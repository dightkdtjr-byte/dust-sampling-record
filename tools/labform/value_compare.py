# 재검증: 같은 원자료로 원본(LibreOffice로 다시 계산)과 최종의 계산값을 칸 단위로 견줌. python3 -I value_compare.py <원본 계산 폴더> <기록부 폴더> [줄 수]
import sys, os, re, glob, collections, warnings
warnings.filterwarnings('ignore'); import openpyxl
OC, G = sys.argv[1], sys.argv[2]
KEY = {'불소화합물':'플루오린화합물','브롬':'브로민화합물','시안화수소':'사이안화수소','포름알데하이드':'폼알데하이드'}
key = lambda p: (lambda k: KEY.get(k, k))(re.search(r'시험기록부_([^_]+?)(?:_|\.xls)', os.path.basename(p)).group(1))
of = {key(p): p for p in glob.glob(OC + '/*.xls[xm]')}
tot = collections.Counter(); diffs = collections.defaultdict(list)
num = lambda v: isinstance(v, (int, float)) and not isinstance(v, bool)
for k, p in sorted(of.items()):
    wf = openpyxl.load_workbook(p)                                   # 원본 수식
    wo = openpyxl.load_workbook(os.path.join(OC, 'lo', os.path.splitext(os.path.basename(p))[0] + '.xlsx'), data_only=True)
    g = glob.glob(f'{G}/lo/시험기록부_{k}.xlsx')[0]; wg = openpyxl.load_workbook(g, data_only=True)
    sg = [s for s in wg.sheetnames if s != '기준']
    for a, b in zip(wf.sheetnames, sg):
        for row in wf[a].iter_rows():
            for c in row:
                if not (isinstance(c.value, str) and c.value.startswith('=')): continue
                x, y = wo[a][c.coordinate].value, wg[b][c.coordinate].value
                tot['수식 칸'] += 1
                if num(x) and num(y):
                    if abs(x - y) <= 1e-9 * max(1, abs(x)): tot['값 같음'] += 1
                    else: tot['값 다름'] += 1; diffs[k].append((b, c.coordinate, x, y))
                elif x == y or (x in (None, '') and y in (None, '')): tot['같음(글자·빈칸)'] += 1
                else: tot['글자·종류 다름'] += 1; diffs[k].append((b, c.coordinate, x, y))
print(dict(tot))
for k, l in diffs.items():
    print(f'== {k}: {len(l)}칸')
    for d in l[:int(sys.argv[3]) if len(sys.argv) > 3 else 6]: print('    ', d[0][:14], d[1], repr(d[2])[:38], '→', repr(d[3])[:38])
