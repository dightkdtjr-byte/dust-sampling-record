#!/bin/bash
# lo.sh <폴더> : 폴더 안 xlsx/xlsm을 강제 재계산으로 xlsx(lo/)와 pdf(pdf/)로 바꾼다
D="$(cd "$1" && pwd)"
mkdir -p "$D/.lo/user"
cat > "$D/.lo/user/registrymodifications.xcu" <<'XCU'
<?xml version="1.0" encoding="UTF-8"?>
<oor:items xmlns:oor="http://openoffice.org/2001/registry" xmlns:xs="http://www.w3.org/2001/XMLSchema" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">
<item oor:path="/org.openoffice.Office.Calc/Formula/Load"><prop oor:name="OOXMLRecalcMode" oor:op="fuse"><value>0</value></prop></item>
<item oor:path="/org.openoffice.Office.Calc/Formula/Load"><prop oor:name="ODFRecalcMode" oor:op="fuse"><value>0</value></prop></item>
</oor:items>
XCU
cd "$D"
timeout 280 soffice -env:UserInstallation=file://$D/.lo --headless --convert-to xlsx --outdir lo *.xls[xm] >/dev/null 2>&1
if [ "$2" = "pdf" ]; then timeout 280 soffice -env:UserInstallation=file://$D/.lo --headless --convert-to pdf --outdir pdf *.xls[xm] >/dev/null 2>&1; fi
ls lo pdf 2>/dev/null
