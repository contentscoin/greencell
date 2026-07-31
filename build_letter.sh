#!/usr/bin/env bash
# 제조사 확인 요청 공문 생성 (OfficeCLI)
set -euo pipefail
export PATH="$HOME/.local/bin:$PATH"
cd "$(dirname "$0")"

DOC="PDRN_제조사_확인요청_공문.docx"
KR="맑은 고딕"

rm -f "$DOC"
officecli create "$DOC" >/dev/null

# 1) 본문 기본 서체(한글) 지정
officecli set "$DOC" /styles/Normal \
  --prop font="$KR" --prop font.ea="$KR" \
  --prop size=10.5 --prop lineSpacing=1.45 --prop spaceAfter=6pt >/dev/null

# 2) 제목 스타일 신규 정의 (빈 템플릿에는 Heading1/2가 없음)
officecli add "$DOC" /styles --type style \
  --prop id=Heading1 --prop name="Heading 1" --prop type=paragraph \
  --prop basedOn=Normal --prop next=Normal --prop qFormat=true \
  --prop uiPriority=9 \
  --prop font="$KR" --prop font.ea="$KR" \
  --prop size=17 --prop bold=true --prop color=#1C8F55 \
  --prop align=center --prop spaceAfter=14pt --prop keepNext=true >/dev/null

# 3) Heading2 스타일 신규 정의
officecli add "$DOC" /styles --type style \
  --prop id=Heading2 --prop name="Heading 2" --prop type=paragraph \
  --prop basedOn=Normal --prop next=Normal --prop qFormat=true \
  --prop uiPriority=9 \
  --prop font="$KR" --prop font.ea="$KR" \
  --prop size=12.5 --prop bold=true --prop color=#14372A \
  --prop spaceBefore=16pt --prop spaceAfter=7pt --prop keepNext=true >/dev/null

# 4) 본문 = 마크다운 확장
officecli add "$DOC" / --type markdown --prop src=letter_src.md >/dev/null

# 5) 마크다운 확장이 남긴 직접 서식(제목 16~18pt) 정리 + 문단 다듬기
find_para() {   # $1 = 스타일명 또는 본문 앞부분
  officecli query "$DOC" paragraph --json |
    python3 -c "
import json,sys
key = sys.argv[1]
for r in json.load(sys.stdin)['data']['results']:
    style = (r.get('style') or '').replace(' ', '')
    if style == key or r.get('text','').startswith(key):
        print(r['path'])
" "$1"
}

# 제목 크기: H1 17pt / H2 13pt
for p in $(find_para Heading1); do
  officecli set "$DOC" "$p" --prop size=17 >/dev/null
done
for p in $(find_para Heading2); do
  officecli set "$DOC" "$p" --prop size=13 >/dev/null
done

# 인사말 문단: 표와 붙지 않도록 위 여백
for p in $(find_para "안녕하십니까"); do
  officecli set "$DOC" "$p" --prop spaceBefore=14pt >/dev/null
done

# '2. 확인 요청 항목'은 표와 함께 2페이지에서 시작 (제목만 남는 것 방지)
for p in $(find_para "2. 확인 요청 항목"); do
  officecli set "$DOC" "$p" --prop pageBreakBefore=true --prop spaceBefore=0pt >/dev/null
done

officecli close "$DOC" >/dev/null
echo "created: $DOC"
officecli view "$DOC" outline
