# 시트별 실제 인쇄 쪽수: 시트 하나만 보이게 한 사본을 LibreOffice로 PDF로 찍어 센다
#   python3 -I tools/labform/sheet_pages.py <기록부 폴더> <작업 폴더>  → <작업 폴더>/pages.json {파일: {시트: 쪽수}}
import sys, os, re, glob, json, shutil, zipfile, subprocess
here = os.path.dirname(os.path.abspath(__file__)); sys.path.insert(0, here)
import lo_pool
G, W = sys.argv[1], sys.argv[2]
if os.path.exists(W): shutil.rmtree(W)
os.makedirs(W)
jobs = []
for p in sorted(glob.glob(G + '/*.xls[xm]')):
    z = zipfile.ZipFile(p); wbx = z.read('xl/workbook.xml').decode()
    names = [(m.group(0), m.group(1)) for m in re.finditer(r'<sheet [^>]*name="([^"]+)"[^>]*/>', wbx) if 'state="hidden"' not in m.group(0)]
    for i, (tag, nm) in enumerate(names):
        d = os.path.join(W, f'{len(jobs):03d}'); os.makedirs(d)
        # LibreOffice는 숨긴 시트도 찍으므로, 다른 시트는 내용·그림·인쇄 영역을 비운다 (빈 시트는 쪽이 없음)
        rels = z.read('xl/_rels/workbook.xml.rels').decode()
        empty = set()
        alltags = [(m.group(0), m.group(1)) for m in re.finditer(r'<sheet [^>]*name="([^"]+)"[^>]*/>', wbx)]
        for t2, nm2 in alltags:
            if t2 == tag: continue
            rid = re.search(r'r:id="([^"]+)"', t2).group(1)
            t = re.search(rf'Id="{rid}"[^>]*Target="([^"]+)"', rels) or re.search(rf'Target="([^"]+)"[^>]*Id="{rid}"', rels)
            empty.add('xl/' + t.group(1).lstrip('/').replace('xl/', ''))
        idx = [k for k, m in enumerate(re.finditer(r'<sheet [^>]*/>', wbx)) if m.group(0) == tag][0]
        w2 = re.sub(r'<definedName name="_xlnm\.Print_(Area|Titles)" localSheetId="(\d+)">[^<]*</definedName>',
                    lambda m: m.group(0) if int(m.group(2)) == idx else '', wbx)
        dst = os.path.join(d, os.path.basename(p))
        with zipfile.ZipFile(p) as zi, zipfile.ZipFile(dst, 'w', zipfile.ZIP_DEFLATED) as zo:
            for it in zi.infolist():
                data = zi.read(it.filename)
                if it.filename == 'xl/workbook.xml': data = w2.encode('utf8')
                elif it.filename in empty:
                    x = data.decode('utf8')
                    x = re.sub(r'<sheetData>.*?</sheetData>', '<sheetData/>', x, flags=re.S)
                    x = re.sub(r'<(drawing|legacyDrawing)\b[^>]*/>', '', x)
                    x = re.sub(r'<mergeCells.*?</mergeCells>', '', x, flags=re.S)
                    x = re.sub(r'<conditionalFormatting.*?</conditionalFormatting>', '', x, flags=re.S)
                    x = re.sub(r'<headerFooter\b[^>]*?(?:/>|>.*?</headerFooter>)', '', x, flags=re.S)
                    data = x.encode('utf8')
                elif any(it.filename == e.replace('worksheets/', 'worksheets/_rels/') + '.rels' for e in empty):
                    x = data.decode('utf8')      # 메모(comments)·그림 연결도 끊음 (빈 시트에 메모만 한 쪽 찍히지 않게)
                    x = re.sub(r'<Relationship\b[^>]*(comments|vmlDrawing|drawing)[^>]*/>', '', x)
                    data = x.encode('utf8')
                zo.writestr(it, data)
        jobs.append((d, os.path.basename(p), nm.replace('&amp;', '&')))
lo_pool.run([d for d, _, _ in jobs], pdf=True)
out = {}
for d, f, nm in jobs:
    pdf = glob.glob(os.path.join(d, 'pdf', '*.pdf'))
    n = int(re.search(r'Pages:\s+(\d+)', subprocess.run(['pdfinfo', pdf[0]], capture_output=True, text=True).stdout).group(1)) if pdf else -1
    out.setdefault(f, {})[nm] = n
json.dump(out, open(os.path.join(W, 'pages.json'), 'w'), ensure_ascii=False, indent=1)
print('시트', len(jobs))
