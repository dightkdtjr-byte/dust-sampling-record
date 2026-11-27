import sys, glob, os, re, zipfile, openpyxl, warnings
warnings.filterwarnings('ignore')
for f in sorted(glob.glob(sys.argv[1] + '/*.xlsx')):
    wb = openpyxl.load_workbook(f, data_only=True)
    bad = []
    for ws in wb.worksheets:
        for row in ws.iter_rows():
            for c in row:
                if isinstance(c.value, str) and c.value.strip() in ('#REF!', '#DIV/0!', '#VALUE!', '#N/A', '#NAME?', '#NUM!', '#NULL!', '#SPILL!', '#CALC!'):
                    bad.append(f'{ws.title}!{c.coordinate}={c.value}')
    print(os.path.basename(f), '오류 칸', len(bad), bad[:12])

# 원본 칸이 없는 공유 수식(엑셀에서 수식이 비어 보임): 다시 계산하기 전 파일(lo/의 위 폴더)에서
src = os.path.dirname(os.path.abspath(sys.argv[1]))
for f in sorted(glob.glob(src + '/*.xls[xm]')):
    z = zipfile.ZipFile(f); bad = 0
    for n in z.namelist():
        if n.startswith('xl/worksheets/sheet'):
            x = z.read(n).decode('utf8', 'ignore')
            masters = {m.group(1) for m in re.finditer(r'<f t="shared"(?=[^>]*\bref=")[^>]*\bsi="(\d+)"', x)}
            bad += sum(1 for s_ in re.findall(r'<f t="shared"[^>]*si="(\d+)"[^>]*/>', x) if s_ not in masters)
            bad += len(re.findall(r'<f>\s*</f>|<f/>', x))
    if bad:
        print(os.path.basename(f), '빈 수식 칸', bad)
