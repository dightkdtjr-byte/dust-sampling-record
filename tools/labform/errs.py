import sys, glob, os, openpyxl, warnings
warnings.filterwarnings('ignore')
for f in sorted(glob.glob(sys.argv[1] + '/*.xlsx')):
    wb = openpyxl.load_workbook(f, data_only=True)
    bad = []
    for ws in wb.worksheets:
        for row in ws.iter_rows():
            for c in row:
                if isinstance(c.value, str) and c.value.startswith('#'):
                    bad.append(f'{ws.title}!{c.coordinate}={c.value}')
    print(os.path.basename(f), '오류 칸', len(bad), bad[:12])
