# 원본의 저장값과, 정리한 양식에 같은 입력을 넣고 다시 계산한 값을 수식 칸마다 비교한다
# 쓰는 법: python3 -I tools/labform/compare.py <원본 폴더> <다시 계산한 xlsx 폴더(lo)>
# 다르게 나와도 되는 칸: 일부러 고친 오류 칸, 그리고 LibreOffice가 빈 이름을 찾은 것으로 치는 벤조(a)피렌 계산과정 102~108행
import sys, openpyxl, warnings
warnings.filterwarnings('ignore')
ORIG, LO = sys.argv[1], sys.argv[2]
PAIRS = [('dust.xlsx', 'labtest-dust.xlsx'), ('gas.xlsx', 'labtest-gas.xlsx'), ('hcn.xlsx', 'labtest-hcn.xlsx'), ('bap.xlsx', 'labtest-bap.xlsx'), ('vocs.xlsm', 'labtest-vocs.xlsx')]
def same(a, b):
    if a in (None, '') and b in (None, ''): return True
    if isinstance(a, (int, float)) and isinstance(b, (int, float)): return abs(a - b) <= 1e-9 * max(1, abs(a), abs(b))
    return str(a) == str(b)
for o, t in PAIRS:
    wo = openpyxl.load_workbook(f'{ORIG}/{o}', data_only=True); wof = openpyxl.load_workbook(f'{ORIG}/{o}')
    wt = openpyxl.load_workbook(f'{LO}/{t}', data_only=True)
    total = 0; diff = []
    for ws in wof.worksheets:
        for row in ws.iter_rows():
            for c in row:
                if isinstance(c.value, str) and c.value.startswith('='):
                    total += 1
                    a = wo[ws.title][c.coordinate].value; b = wt[ws.title][c.coordinate].value
                    if not same(a, b): diff.append(f'{ws.title}!{c.coordinate}: 원본 {a!r} → 정리본 {b!r}')
    print(o, '수식 칸', total, '다른 칸', len(diff))
    for d in diff[:60]: print('   ', d)
