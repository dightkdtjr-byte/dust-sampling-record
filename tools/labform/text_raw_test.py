# 기기 값에 "-"·음수·글자가 들어와도 최종 결과가 나오는지 (중금속·비소): 사본에 값을 넣고 LibreOffice로 다시 계산
#   python3 -I tools/labform/text_raw_test.py <기록부 폴더> <작업 폴더>
import sys, os, glob, shutil, subprocess, warnings
warnings.filterwarnings('ignore')
import openpyxl
here = os.path.dirname(os.path.abspath(__file__))
G, W = sys.argv[1], sys.argv[2]
shutil.rmtree(W, ignore_errors=True); os.makedirs(W)
METALS = ['구리화합물', '납화합물', '아연화합물', '크로뮴화합물', '니켈화합물', '카드뮴화합물', '베릴륨화합물']
CASES = {   # 이름: (파일, {raw 칸: 값})
    '중금속_시료모두대시': ('중금속', {'I19': '-', 'I20': '-', 'I21': '-'}),
    '중금속_바탕모두대시': ('중금속', {'K16': '-', 'K17': '-', 'K18': '-'}),
    '중금속_하나대시': ('중금속', {'J19': '-'}),
    '중금속_글자': ('중금속', {'M19': '<0.001', 'M20': 'N/A', 'M21': '0.004'}),
    '중금속_음수': ('중금속', {'L19': -0.01, 'L20': -0.02, 'L21': -0.03}),
    '중금속_검출_글자숫자': ('중금속', {'I19': 5, 'I20': '5', 'I21': 5}),
    '비소_시료모두대시': ('비소화합물', {'H19': '-', 'H20': '-', 'H21': '-'}),
    '비소_음수': ('비소화합물', {'H22': -0.5, 'H23': -0.5, 'H24': -0.5}),
}
for n, (item, ch) in CASES.items():
    p = glob.glob(os.path.join(G, f'시험기록부_{item}.xls[xm]'))[0]
    wb = openpyxl.load_workbook(p, keep_vba=p.endswith('xlsm'))
    raw = 'Raw data' if item == '중금속' else 'rawdata(As)'
    for a, v in ch.items():
        wb[raw][a] = v
    wb.save(os.path.join(W, n + os.path.splitext(p)[1]))
subprocess.run(['bash', os.path.join(here, 'lo.sh'), W], capture_output=True)
bad = 0; tot = 0
for n, (item, _) in CASES.items():
    v = openpyxl.load_workbook(glob.glob(os.path.join(W, 'lo', n + '.*'))[0], data_only=True)
    sheets = METALS if item == '중금속' else ['시험기록부(As)']
    for sh in sheets:
        ws = v[sh]; r = 17 if item == '중금속' else 18
        res = ws.cell(r, 4).value
        tot += 1
        ok = res not in (None, '') and not str(res).startswith('#')
        if '검출' in n and sh == '구리화합물':       # 크게 넣은 구리는 숫자로 나와야 함 (글자 '5'도 숫자로 읽음)
            ok = ok and str(res) != '불검출'
        calc = v['계산과정'] if item == '중금속' else v['계산과정(As)']
        texts = [c.coordinate for row in calc.iter_rows() for c in row if c.value == '0']
        if texts:
            ok = False
        bad += not ok
        print(('  ' if ok else '✖ ') + f'{n} {sh}: 결과 {res!r}' + (f' 글자 "0" 칸 {texts[:4]}' if texts else ''))
print(f'기기 값 글자·음수 시험: {tot - bad}/{tot} 결과 나옴')
