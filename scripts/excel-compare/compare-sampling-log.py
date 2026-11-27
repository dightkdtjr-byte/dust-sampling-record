# 통합시료채취기록부: 넣은 값에서 따로 계산한 값과, 추출한 엑셀을 다시 계산한 값을 줄마다 대조한다.
#  - 채취유량 K = (가스미터 종료값 − 시작값) × Yd  (입자 샘플러·가스 샘플러만 Yd, 음압상자·그 밖은 그대로). 숫자가 아니라 수식이어야 한다
#  - 게이지압 N = 압력(mmH2O) ÷ 13.6,  표준 채취량 V = K × 273/(273+온도) × (채취한 날의 대기압 + 게이지압)/760  (온도를 안 적은 줄은 빈칸)
#  - 분석기한 T = 채취일에서 시험검사기간(평일 수)만큼 뒤. 기간이 글자(현장측정)면 그 글자, 비면 빈칸
#  - 일차별 시트: 그날 채취한 줄만 보이고, 시험검사기간·분석기한이 입력 시트와 같으며, 가장 이른/늦은 기한이 맞아야 한다
#  - 모든 시트에 오류 칸(#VALUE! 등)이 없어야 한다
import sys, json, warnings, datetime
warnings.filterwarnings('ignore')
import openpyxl
d = sys.argv[1]
src = json.load(open(f'{d}/samplinglog.input.json')); log, kinds = src['log'], src['kinds']
wv = openpyxl.load_workbook(f'{d}/lo-log/samplinglog.xlsx', data_only=True); ws = wv['시료채취 작성']
wf = openpyxl.load_workbook(f'{d}/samplinglog.xlsx'); wsf = wf['시료채취 작성']
n = bad = 0
blank = lambda x: x is None or x == ''
def cmp(name, x, a, tol=1e-6):
    global n, bad
    n += 1
    ok = isinstance(x, (int, float)) and abs(x - a) <= tol * max(1, abs(a))
    if not ok: bad += 1; print(f'  불일치 {name}: 엑셀 {x!r} / 계산 {a!r}')
def same(name, x, a):
    global n, bad
    n += 1
    if blank(x) and blank(a): return
    if x != a: bad += 1; print(f'  불일치 {name}: 엑셀 {x!r} / 계산 {a!r}')
def workday(start, days):
    cur = start
    while days > 0:
        cur += datetime.timedelta(days=1)
        if cur.weekday() < 5: days -= 1
    return cur
for i, day in enumerate(log['days']): cmp(f'{i+1}일차 대기압', ws[f'I{6+i}'].value, float(day['pressure']))
due = {}  # 줄 → (채취일, 시험검사기간, 분석기한)
for row, kind in kinds.items():
    r = int(row); rd = log['rows'][row]
    date = datetime.datetime.strptime(log['days'][rd['dayIndex']]['date'], '%Y-%m-%d')
    period = ws[f'S{r}'].value
    want = workday(date, int(period)) if isinstance(period, (int, float)) else period
    if r not in (14, 15, 16): same(f'{r}행 분석기한', ws[f'T{r}'].value, want)
    due[r] = (date, period, want)
    if kind == 'thc': continue
    yd = float(log['particleSamplerYd']) if kind == 'particle' else float(log['gasSamplerYd']) if kind.startswith('gas') else 1
    k = round((float(rd['meterEnd']) - float(rd['meterStart'])) * yd * 100 + 1e-9) / 100
    cmp(f'{r}행 채취유량', ws[f'K{r}'].value, k, 1e-9)
    same(f'{r}행 채취유량이 수식인지', str(wsf[f'K{r}'].value).startswith('='), True)
    if r >= 14:
        gauge = float(rd['pressureH2O']) / 13.6 if rd['pressureH2O'] else 0
        pa = float(log['days'][rd['dayIndex']]['pressure'])  # 그 줄을 채취한 날의 대기압
        nval = ws[f'N{r}'].value
        if rd['pressureH2O']: cmp(f'{r}행 게이지압', nval, gauge)
        if rd['temp'] == '': same(f'{r}행 표준 채취량(온도 없음 → 빈칸)', ws[f'V{r}'].value, '')
        else: cmp(f'{r}행 표준 채취량', ws[f'V{r}'].value, k * 273 / (273 + float(rd['temp'])) * (pa + gauge) / 760)
# 비소(입자상 32행·가스상 33행)는 한 물질: 의뢰 칸은 병합해 O 하나, 의뢰 개수에 1로 잡히고, 채취유량이 같다
if '32' in kinds and '33' in kinds:
    same('비소 가스상 의뢰 칸(병합이라 빈칸)', ws['C33'].value, '')
    cmp('비소 입자상·가스상 채취유량이 같은지', ws['K33'].value, ws['K32'].value, 1e-9)
    same('비소 가스상 sampler No.', blank(ws['E33'].value), blank(ws['E32'].value))
# 의뢰 개수: 줄 수가 아니라 화면이 세는 종 수와 같아야 한다 (알데히드는 고른 종 수만큼, 3회차·암모니아 2줄은 한 번만, 비소 입자·가스상은 1)
same('의뢰 개수(C63) = 화면의 의뢰 종 수', ws['C63'].value, src['stats']['requestedCount'])
same('의뢰 개수가 수식인지', str(wsf['C63'].value).startswith('='), True)
# 일차별 시트
for sheet in wv.worksheets[1:]:
    day = sheet['J7'].value; dues = []
    for r in range(12, 58):
        m = r + 1
        if m in due and due[m][0] == day:
            same(f'{sheet.title} {r}행 시험검사기간', sheet[f'M{r}'].value, due[m][1])
            same(f'{sheet.title} {r}행 분석기한', sheet[f'N{r}'].value, due[m][2])
            if isinstance(due[m][2], datetime.datetime): dues.append(due[m][2])
            same(f'{sheet.title} {r}행 시료번호', blank(sheet[f'B{r}'].value), False)
        else:
            same(f'{sheet.title} {r}행(그날 채취 아님) 분석기한', sheet[f'N{r}'].value, '')
    if dues:
        same(f'{sheet.title} 가장 이른 기한', sheet['N58'].value, min(dues)); same(f'{sheet.title} 가장 늦은 기한', sheet['N59'].value, max(dues))
errs = [f'{s.title}!{c.coordinate}={c.value}' for s in wv.worksheets for rw in s.iter_rows() for c in rw if isinstance(c.value, str) and c.value[:1] == '#' and c.value[1:4] in ('VAL', 'DIV', 'REF', 'NAM', 'N/A', 'NUM', 'NUL')]
n += 1
if errs: bad += 1; print('  오류 칸:', errs[:10])
print(f'통합시료채취기록부: {n}항목 중 불일치 {bad}')
sys.exit(1 if bad else 0)
