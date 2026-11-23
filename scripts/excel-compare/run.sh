#!/bin/bash
# 프로그램 화면의 계산값과, 추출한 엑셀을 다시 계산한 값을 항목별로 대조한다.
#
# 준비: 다른 창에서 시험용 서버를 켠다 →  VITE_E2E=1 npx vite --port 4179 --strictPort
# 실행: scripts/excel-compare/run.sh <결과를 둘 폴더> [조건 이름…]      (SEED=숫자 를 앞에 붙이면 다른 견본 숫자로 돈다)
# 필요: node + playwright, LibreOffice(soffice), python3 + openpyxl
#
# 기록부 종류 × 수분 측정법 × 굴뚝 모양·측정점 수 × 정압 입력 방식을 섞은 견본을 화면에 넣어 엑셀로 뽑고,
# LibreOffice로 다시 계산한 뒤 계산 칸마다 화면 값과 비교한다. 맨 끝 줄 "전체 N항목, 불일치 0"이 나와야 한다.
set -e
HERE="$(cd "$(dirname "$0")" && pwd)"; OUT="$(mkdir -p "$1" && cd "$1" && pwd)"; shift
node "$HERE/export-scenarios.cjs" "$OUT" "$@"
(cd "$OUT" && for f in *.xlsm; do timeout 240 soffice -env:UserInstallation=file://$OUT/.lo-profile --headless --convert-to xlsx --outdir lo "$f" >/dev/null 2>&1; done)
python3 "$HERE/compare.py" "$OUT"
