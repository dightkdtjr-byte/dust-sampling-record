#!/bin/bash
# lo.sh <폴더> : 폴더 안 xlsx/xlsm을 강제 재계산으로 xlsx(lo/)와 pdf(pdf/)로 바꾼다
# 글자→숫자 바꾸기는 엑셀처럼: 숫자 모양 글자만 숫자로, 빈 글자("")는 0이 아니라 #VALUE! (엑셀에서 나는 오류를 여기서도 보게)
D="$(cd "$1" && pwd)"
mkdir -p "$D/.lo/user"
cat > "$D/.lo/user/registrymodifications.xcu" <<'XCU'
<?xml version="1.0" encoding="UTF-8"?>
<oor:items xmlns:oor="http://openoffice.org/2001/registry" xmlns:xs="http://www.w3.org/2001/XMLSchema" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">
<item oor:path="/org.openoffice.Office.Calc/Formula/Load"><prop oor:name="OOXMLRecalcMode" oor:op="fuse"><value>0</value></prop></item>
<item oor:path="/org.openoffice.Office.Calc/Formula/Load"><prop oor:name="ODFRecalcMode" oor:op="fuse"><value>0</value></prop></item>
<item oor:path="/org.openoffice.Office.Calc/Formula/Syntax"><prop oor:name="StringConversion" oor:op="fuse"><value>3</value></prop></item>
<item oor:path="/org.openoffice.Office.Calc/Formula/Syntax"><prop oor:name="EmptyStringAsZero" oor:op="fuse"><value>false</value></prop></item>
</oor:items>
XCU
cd "$D"
timeout 280 soffice -env:UserInstallation=file://$D/.lo --headless --convert-to xlsx --outdir lo *.xls[xm] >/dev/null 2>&1
if [ "$2" = "pdf" ]; then timeout 280 soffice -env:UserInstallation=file://$D/.lo --headless --convert-to pdf --outdir pdf *.xls[xm] >/dev/null 2>&1; fi
ls lo pdf 2>/dev/null
