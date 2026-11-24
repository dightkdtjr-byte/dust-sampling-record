# 통합시료채취기록부: 넣은 값에서 따로 계산한 값과, 추출한 엑셀을 다시 계산한 값을 줄마다 대조한다.
#  - 채취유량 K = (가스미터 종료값 − 시작값) × Yd  (입자 샘플러·가스 샘플러만 Yd, 음압상자·그 밖은 그대로)
#  - 게이지압 N = 압력(mmH2O) ÷ 13.6,  표준 채취량 V = K × 273/(273+온도) × (채취한 날의 대기압 + 게이지압)/760
#  - 시트 전체에 오류 칸(#VALUE! 등)이 없어야 한다
import sys, json, warnings
warnings.filterwarnings('ignore')
import openpyxl
d = sys.argv[1]
src = json.load(open(f'{d}/samplinglog.input.json')); log, kinds = src['log'], src['kinds']
wv = openpyxl.load_workbook(f'{d}/lo-log/samplinglog.xlsx', data_only=True); ws = wv['시료채취 작성']
wf = openpyxl.load_workbook(f'{d}/samplinglog.xlsx'); wsf = wf['시료채취 작성']
n = bad = 0
def cmp(name, x, a, tol=1e-6):
    global n, bad
    n += 1
    ok = isinstance(x, (int, float)) and abs(x - a) <= tol * max(1, abs(a))
    if not ok: bad += 1; print(f'  불일치 {name}: 엑셀 {x!r} / 계산 {a!r}')
for i, day in enumerate(log['days']): cmp(f'{i+1}일차 대기압', ws[f'I{6+i}'].value, float(day['pressure']))
for row, kind in kinds.items():
    r = int(row); rd = log['rows'][row]
    yd = float(log['particleSamplerYd']) if kind == 'particle' else float(log['gasSamplerYd']) if kind.startswith('gas') else 1
    k = round((float(rd['meterEnd']) - float(rd['meterStart'])) * yd * 100 + 1e-9) / 100
    cmp(f'{r}행 채취유량', ws[f'K{r}'].value, k, 1e-9)
    if r >= 14:
        gauge = float(rd['pressureH2O']) / 13.6 if rd['pressureH2O'] else 0
        pa = float(log['days'][rd['dayIndex']]['pressure'])  # 그 줄을 채취한 날의 대기압
        nval = ws[f'N{r}'].value
        if rd['pressureH2O']: cmp(f'{r}행 게이지압', nval, gauge)
        cmp(f'{r}행 표준 채취량', ws[f'V{r}'].value, k * 273 / (273 + float(rd['temp'])) * (pa + gauge) / 760)
errs = [f'{c.coordinate}={c.value}' for rw in ws.iter_rows() for c in rw if isinstance(c.value, str) and c.value.startswith('#') and c.value.rstrip('!?0/AN').startswith('#') and c.value[1:2].isupper()]
n += 1
if errs: bad += 1; print('  오류 칸:', errs[:10])
print(f'통합시료채취기록부: {n}항목 중 불일치 {bad}')
sys.exit(1 if bad else 0)
