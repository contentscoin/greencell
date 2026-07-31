# -*- coding: utf-8 -*-
"""
원본 PDF(그린셀바이옴 소개서)에서 사진·도판을 추출한다.

PowerPoint → PDF 변환 과정에서 큰 사진이 가로 띠 4~5장으로 분할 저장되어 있어,
페이지상의 배치 좌표를 읽어 원본 1장으로 다시 이어 붙인다.
"""
import os
import shutil
from collections import defaultdict

import fitz  # pymupdf
from PIL import Image

PDF = "/projects/sandbox/source/greencell_original_1.pdf"
IMG_DIR = "/projects/sandbox/images"
PAGE_DIR = "/projects/sandbox/pages"

MIN_W, MIN_H = 150, 150      # 아이콘·불릿 등 제외
MIN_BYTES = 12 * 1024
TOL = 1.5                    # 띠 인접 판정 허용 오차(pt)

for d in (IMG_DIR, PAGE_DIR):
    shutil.rmtree(d, ignore_errors=True)
    os.makedirs(d)

doc = fitz.open(PDF)
print(f"원본: {doc.page_count} 페이지\n")


def load_pix(xref):
    pix = fitz.Pixmap(doc, xref)
    if pix.n - pix.alpha >= 4:            # CMYK 등 → RGB
        pix = fitz.Pixmap(fitz.csRGB, pix)
    mode = "RGBA" if pix.alpha else "RGB"
    return Image.frombytes(mode, (pix.width, pix.height), pix.samples)


def save(img, name):
    path = os.path.join(IMG_DIR, name)
    if img.mode == "RGBA":
        img.load()
    img.save(path, "PNG")
    size = os.path.getsize(path)
    if size < MIN_BYTES:
        os.remove(path)
        return None
    return size


saved, merged, skipped = [], 0, 0
seen_xrefs = set()

for pno in range(doc.page_count):
    page = doc[pno]
    # 같은 x 범위 + 세로로 연속된 이미지 = 분할된 한 장
    groups = defaultdict(list)
    for info in page.get_images(full=True):
        xref, w, h = info[0], info[2], info[3]
        for r in page.get_image_rects(xref):
            groups[(round(r.x0, 1), round(r.x1, 1))].append((r.y0, r.y1, xref, w, h))

    for (x0, x1), bands in groups.items():
        bands.sort()
        # 연속 구간으로 쪼개기
        runs, cur = [], [bands[0]]
        for b in bands[1:]:
            same_w = b[3] == cur[-1][3]
            touching = abs(b[0] - cur[-1][1]) <= TOL
            if same_w and touching:
                cur.append(b)
            else:
                runs.append(cur)
                cur = [b]
        runs.append(cur)

        for run in runs:
            xrefs = [b[2] for b in run]
            if all(x in seen_xrefs for x in xrefs):
                continue
            seen_xrefs.update(xrefs)

            if len(run) == 1:
                _, _, xref, w, h = run[0]
                if w < MIN_W or h < MIN_H:
                    skipped += 1
                    continue
                img = load_pix(xref)
                name = f"p{pno+1:02d}_img{xref}_{w}x{h}.png"
            else:
                parts = [load_pix(b[2]) for b in run]
                tw = max(p.width for p in parts)
                th = sum(p.height for p in parts)
                mode = "RGBA" if any(p.mode == "RGBA" for p in parts) else "RGB"
                img = Image.new(mode, (tw, th))
                y = 0
                for p in parts:
                    if p.width != tw:
                        p = p.resize((tw, int(p.height * tw / p.width)), Image.LANCZOS)
                    img.paste(p, (0, y))
                    y += p.height
                name = f"p{pno+1:02d}_merged_{tw}x{th}.png"
                merged += 1

            size = save(img, name)
            if size is None:
                skipped += 1
                continue
            saved.append((pno + 1, name, size, len(run)))

    # 페이지 전체 렌더 (레이아웃 참고용)
    page.get_pixmap(dpi=150).save(os.path.join(PAGE_DIR, f"page{pno+1:02d}.png"))

print(f"이미지 {len(saved)}개 저장  (그중 분할 재조립 {merged}개) / 소형·중복 제외 {skipped}개")
print(f"페이지 렌더 {doc.page_count}장 → pages/\n")
print(f"{'페이지':>5}  {'파일명':46s} {'용량':>8}  비고")
print("-" * 78)
for pno, name, size, nband in saved:
    note = f"{nband}조각 재조립" if nband > 1 else ""
    print(f"{'p'+str(pno):>5}  {name:46s} {size/1024:7.0f}K  {note}")



# ----------------------------------------------------------------------------
# 용도별 분류 + 한글 파일명 정리
# ----------------------------------------------------------------------------
LABELS = {
    # 01 제품 사진 ------------------------------------------------------------
    "p02_img31_934x848.png":      ("01_제품사진", "제품_보틀_대리석_꽃"),
    "p05_merged_1232x864.png":    ("01_제품사진", "제품_보틀2개_숲이슬"),
    "p23_img337_1225x822.png":    ("01_제품사진", "제품_보틀_실험실"),
    "p24_merged_2358x1651.png":   ("01_제품사진", "제품_보틀_로즈박스_문구포함"),
    # 02 브랜드 · 인물 --------------------------------------------------------
    "p03_img43_521x758.png":      ("02_브랜드_인물", "창업자_이상희쉘리"),
    "p02_img32_553x183.png":      ("02_브랜드_인물", "로고_GreenCellBiome_영문"),
    "p01_img12_1626x383.png":     ("02_브랜드_인물", "커버_타이틀_글로우"),
    "p01_img17_231x190.png":      ("02_브랜드_인물", "커버_TM_조각"),
    "p04_img55_511x357.png":      ("02_브랜드_인물", "그래픽_LOVE_하트"),
    "p04_img56_841x561.png":      ("02_브랜드_인물", "그래픽_ION_하트"),
    # 03 성분 도판 ------------------------------------------------------------
    "p12_img208_885x375.png":     ("03_성분도판", "PDRN_작용기전_도식"),
    "p12_img207_241x240.png":     ("03_성분도판", "PDRN_DNA나선"),
    "p10_img175_704x722.png":     ("03_성분도판", "GF5나노좀_피부단면도"),
    "p11_img195_544x448.png":     ("03_성분도판", "성장인자_논문그래프1"),
    "p11_img196_544x448.png":     ("03_성분도판", "성장인자_논문그래프2"),
    "p14_img227_1046x599.png":    ("03_성분도판", "세라마이드_AS_NS_AP_EOP"),
    "p15_img233_1080x598.png":    ("03_성분도판", "세라마이드_NG_NP_EOS"),
    "p17_img259_435x419.png":     ("03_성분도판", "히알루론산_8종_육각형"),
    "p18_img287_416x432.png":     ("03_성분도판", "히알루론산_구조식"),
    "p18_img286_496x198.png":     ("03_성분도판", "히알루론산_피부효능_표"),
    "p18_img288_496x227.png":     ("03_성분도판", "히알루론산_주요특징_표"),
    "p21_img321_480x424.png":     ("03_성분도판", "마이크로바이옴_3종_삼각도"),
    "p22_img327_924x1009.png":    ("03_성분도판", "피부장벽_건강한_유익균"),
    "p22_img328_925x1009.png":    ("03_성분도판", "피부장벽_스트레스_유해균"),
    "p08_img152_814x1047.png":    ("03_성분도판", "사용법_얼굴마사지_일러스트"),
    "p09_merged_1288x1448.png":   ("03_성분도판", "컨셉_그린DNA_손"),
    # 04 배경 · 오버레이 ------------------------------------------------------
    "p01_img11_951x585.png":      ("04_배경", "배경_숲반사_커버"),
    "p06_merged_2573x1448.png":   ("04_배경", "배경_분자_블랙"),
    "p13_merged_2573x1448.png":   ("04_배경", "배경_식물_세라마이드타이틀"),
    "p16_merged_2573x1448.png":   ("04_배경", "배경_물방울_수분"),
    "p19_merged_2573x1448.png":   ("04_배경", "배경_숲_다크"),
    "p24_merged_2573x1448.png":   ("04_배경", "배경_숲_다크2"),
    "p20_img311_1200x794.png":    ("04_배경", "배경_박테리아_블루"),
    "p18_merged_853x1448.png":    ("04_배경", "배경_네이비_그라디언트"),
    "p05_merged_2568x1448.png":   ("04_배경", "배경_크림_단색"),
    "p09_merged_2573x1448.png":   ("04_배경", "배경_딥그린_단색"),
    "p07_img120_699x470.png":     ("04_배경", "배경_카드장식1"),
    "p07_img123_699x470.png":     ("04_배경", "배경_카드장식2"),
    "p20_merged_1572x692.png":    ("04_배경", "배경_크림박스_오버레이"),
    "p25_merged_2478x1652.png":   ("04_배경", "마무리_ThankYou_카드"),
    # 05 텍스트 캡처(재사용 가치 낮음) ----------------------------------------
    "p15_img234_481x262.png":     ("05_텍스트캡처", "세라마이드NG_설명"),
    "p15_img235_463x262.png":     ("05_텍스트캡처", "세라마이드NP_설명"),
    "p15_img236_472x349.png":     ("05_텍스트캡처", "세라마이드EOS_설명"),
    "p15_img237_404x290.png":     ("05_텍스트캡처", "글라이코스핑고리피드_설명"),
}

print("\n용도별 정리")
print("-" * 78)
moved = defaultdict(list)
for old, (cat, label) in LABELS.items():
    src = os.path.join(IMG_DIR, old)
    if not os.path.exists(src):
        print(f"  ! 없음: {old}")
        continue
    page = old.split("_")[0]                       # pNN
    wh = old.rsplit("_", 1)[-1].replace(".png", "")  # WxH
    os.makedirs(os.path.join(IMG_DIR, cat), exist_ok=True)
    new = f"{page}_{label}_{wh}.png"
    os.rename(src, os.path.join(IMG_DIR, cat, new))
    moved[cat].append(new)

for cat in sorted(moved):
    print(f"\n[{cat}]  {len(moved[cat])}개")
    for n in sorted(moved[cat]):
        print(f"    {n}")

left = [f for f in os.listdir(IMG_DIR) if f.endswith(".png")]
if left:
    print(f"\n미분류 {len(left)}개: {left}")
