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
# LibreOffice는 엑셀 파일에 저장돼 있는 계산값을 그대로 믿고 읽기도 한다 → 열 때마다 모든 수식을 다시 계산하도록 설정한 전용 프로필을 쓴다
# (이 설정이 없으면 프로그램이 파일에 적어 둔 값을 되읽는 셈이라 비교가 되지 않는다)
mkdir -p "$OUT/.lo-recalc/user"
cat > "$OUT/.lo-recalc/user/registrymodifications.xcu" <<'XCU'
<?xml version="1.0" encoding="UTF-8"?>
<oor:items xmlns:oor="http://openoffice.org/2001/registry" xmlns:xs="http://www.w3.org/2001/XMLSchema" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">
<item oor:path="/org.openoffice.Office.Calc/Formula/Load"><prop oor:name="OOXMLRecalcMode" oor:op="fuse"><value>0</value></prop></item>
<item oor:path="/org.openoffice.Office.Calc/Formula/Load"><prop oor:name="ODFRecalcMode" oor:op="fuse"><value>0</value></prop></item>
</oor:items>
XCU
(cd "$OUT" && for f in *.xlsm; do timeout 240 soffice -env:UserInstallation=file://$OUT/.lo-recalc --headless --convert-to xlsx --outdir lo "$f" >/dev/null 2>&1; done)
python3 "$HERE/compare.py" "$OUT"
# 조건을 따로 고르지 않았으면 통합시료채취기록부도 함께 본다
if [ $# -eq 0 ]; then
  node "$HERE/sampling-log.cjs" "$OUT"
  (cd "$OUT" && timeout 240 soffice -env:UserInstallation=file://$OUT/.lo-recalc --headless --convert-to xlsx --outdir lo-log samplinglog.xlsx >/dev/null 2>&1)
  python3 "$HERE/compare-sampling-log.py" "$OUT"
fi
