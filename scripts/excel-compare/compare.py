# 사용: python3 fullcmp.py <dir> [-v]   → 엑셀(LibreOffice 재계산본)의 계산 칸과 화면 계산값을 항목별로 대조
from decimal import Decimal, ROUND_HALF_UP
import openpyxl, json, warnings, glob, sys, re, zipfile
warnings.filterwarnings('ignore')
d = sys.argv[1]; verbose = '-v' in sys.argv
TOL = 1e-6
tot = bad = skipped = 0; groups = {}
def cmp(tag, grp, name, x, a, rnd=None):
    global tot, bad
    tot += 1; g = groups.setdefault(grp, [0, 0]); g[0] += 1
    try:
        xf = float(x); af = float(a)
        if rnd is not None:
            q = Decimal(1).scaleb(-rnd); af = float(Decimal(f'{af:.12g}').quantize(q, ROUND_HALF_UP)); xf = float(Decimal(f'{xf:.12g}').quantize(q, ROUND_HALF_UP))
        ok = abs(xf - af) <= abs(af) * TOL + 1e-9
    except Exception:
        ok = (x in (None, '', '-')) and (a in (None, '', '-'))
    if not ok: bad += 1; g[1] += 1
    if not ok or verbose: print(('  ' if ok else '✗ ') + f'{tag} [{grp}] {name}: 엑셀 {x} / 프로그램 {a}')
def names(path):
    x = zipfile.ZipFile(path).read('xl/workbook.xml').decode()
    return {m.group(1): m.group(2) for m in re.finditer(r'<definedName name="([^"]+)"(?: [^>]*)?>([^<]*)', x) if 'localSheetId' not in m.group(0)}
for f in sorted(glob.glob(d + '/lo/*.xlsx')):
    tag = f.split('/')[-1][:-5]
    wv = openpyxl.load_workbook(f, data_only=True); wf = openpyxl.load_workbook(f'{d}/{tag}.xlsm', keep_vba=True)
    D = json.load(open(f'{d}/{tag}.derived.json')); F = D['full']; nm = names(f'{d}/{tag}.xlsm')
    flow = tag.startswith('flow')
    def nv(name):
        ref = nm[name].replace('$', ''); sh, c = ref.split('!'); return wv[sh.strip("'")][c].value
    def find(sheet, pat):
        for row in wf[sheet].iter_rows():
            for c in row:
                if isinstance(c.value, str) and re.search(pat, c.value): return c
        return None
    # 보이는 칸의 오류값
    for s in wv:
        rs = wf[s.title]
        if rs.sheet_state != 'visible': continue
        for row in s.iter_rows():
            for c in row:
                if isinstance(c.value, str) and re.fullmatch(r'#(DIV/0!|VALUE!|REF!|NAME\?|N/A|NUM!|NULL!)|Err:\d+', c.value) and not rs.row_dimensions[c.row].hidden:
                    cmp(tag, '오류칸', f'{s.title}!{c.coordinate}', c.value, None)
    nz = wv['노즐산정']; g = D['gas']; P = F['pre']
    # ── 가스 조성·분자량·밀도
    for name, cell, a in [('O2 평균', 'D26', g['o2']), ('CO2 평균', 'D27', g['co2']), ('CO vol%', 'D28', g['co'] / 10000), ('N2', 'D29', g['n2']), ('건조가스 분자량', 'H8', g['Md']),
                          ('습가스 분자량', 'I8', g['Ms']), ('밀도 r0', 'H26', g['r0']), ('사전 수분량', 'C15', g['Xw'])]:
        cmp(tag, '가스·수분', name, nz[cell].value, a)
    # ── 예비조사(측정점) 평균과 유속
    npt = len(json.load(open(f'{d}/{tag}.input.json'))['data']['points'])
    if npt <= 5:
        for name, cell, a in [('평균 온도', 'G15', P['avgTs']), ('평균 전압', 'H15', P['avgTp']), ('평균 정압', 'I15', P['avgSp']), ('평균 동압', 'J15', P['avgDp'])]:
            cmp(tag, '예비조사', name, nz[cell].value, a)
    cmp(tag, '예비조사', '밀도 r', nz['K15'].value, P['density']); cmp(tag, '예비조사', '밀도 r(위 칸)', nz['K8'].value, P['density'])
    cmp(tag, '예비조사', '유속', nz['G39'].value, P['velocity'])
    # ── 노즐 표
    if not flow:
        plan = nz['A39'].value; r = 39; byNum = {n['num']: n for n in F['nozzles']}
        while nz[f'B{r}'].value not in (None, ''):
            n = byNum.get(float(nz[f'B{r}'].value))
            if n:
                cmp(tag, '노즐표', f'{r}행 직경', nz[f'C{r}'].value, n['d']); cmp(tag, '노즐표', f'{r}행 K', nz[f'D{r}'].value, n['k']); cmp(tag, '노즐표', f'{r}행 ΔH', nz[f'F{r}'].value, n['dH'])
                cmp(tag, '노즐표', f'{r}행 채취량 L', nz[f'H{r}'].value, n['qm'] * plan); cmp(tag, '노즐표', f'{r}행 채취량 SL', nz[f'I{r}'].value, n['qsl'] * plan)
                cmp(tag, '노즐표', f'{r}행 채취량 Sm3', nz[f'J{r}'].value, n['qsl'] * plan / 1000)
            else: cmp(tag, '노즐표', f'{r}행 노즐번호 {nz[f"B{r}"].value}', 'x', None)
            r += 1
    if not flow:
        for r in range(51, 59):
            n = byNum.get(float(nz[f'N{r}'].value)) if isinstance(nz[f'N{r}'].value, (int, float)) else None
            if n: cmp(tag, '노즐 순위표', f'{r}행 예상시간(분)', nz[f'Q{r}'].value, n['time'])
        used = byNum.get(float(nz['K39'].value)) if isinstance(nz['K39'].value, (int, float)) else None
        if used: cmp(tag, '노즐 순위표', '사용 노즐 유량(SL/min)', nz['O46'].value, used['qsl'])
    # ── 측정위치
    ps = [s for s in wv.sheetnames if '시료측정위치' in s][0]; pv = wv[ps]; dref = nm['D'].replace('$', '').split('!')[1]; dr = int(re.sub(r'\D', '', dref))
    pts = D['points']
    cmp(tag, '측정위치', '내경(깊이)', pv[f'C{dr}'].value, D['diameter']); cmp(tag, '측정위치', '단면적', pv[f'E{dr}'].value, D['area'])
    if pts:
        horiz = json.load(open(f'{d}/{tag}.input.json'))['data']['stackOrientation'] == 'horizontal'
        cmp(tag, '측정위치', '측정점 수', pv[f'F{dr}'].value, (pts['reduced'] or pts['total']) if pts['rect'] else (1 if pts['isCenterOnly'] else pts['perRadius'] * (4 if horiz else 1)))
        for i, w in enumerate(pts['wall']):
            cmp(tag, '측정위치', f'{i+1}번 벽에서 거리', pv[f'C{dr+3+i}'].value, w, 2)
        for i, w in enumerate(pts['near']):
            cmp(tag, '측정위치', f'{i+1}번 삽입 길이', pv[f'D{dr+3+i}'].value, w, 2)
        # 원형은 반대편 측정공에서의 삽입 길이, 사각은 세로 방향(측정공 줄) 위치
        for i, w in enumerate(pts['far']):
            cmp(tag, '측정위치', f"{i+1}번 {'세로 위치' if pts['rect'] else '반대편'}", pv[f'E{dr+3+i}'].value, w, 2)
    # ── 수분량
    mv = wv['수분량']; M = F['moisture']
    if M['method'] == 'auto': cmp(tag, '수분량', '자동측정 평균', mv['B15'].value, M['pre'])
    else:
        cmp(tag, '수분량', '사전 흡습 수분(g)', mv['E24'].value, M['preWater']); cmp(tag, '수분량', '사전 수분량(무게법)', mv['H28'].value, M['pre'])
    if not flow:
        tc = find('수분량', r'^=SUM\(E33:'); tr = tc.row
        cmp(tag, '수분량', '사후 흡습 수분(g)', mv[f'E{tr}'].value, M['water']); cmp(tag, '수분량', '사후 수분량', mv[f'H{tr+6}'].value, M['post'])
        cmp(tag, '수분량', '가스채취량(보정, L)', mv[f'C{tr+6}'].value, M['postDeltaVmL'] * 0.991)
    # ── 기록지
    R = F['result']
    if not flow:
        rv = wv['기록지']; ac = find('기록지', r'^=MAX\(B17:B'); ar = ac.row; A = F['meterAvg']; Pa = rv['J7'].value
        for name, col, a in [('채취시간', 'B', A['minutes']), ('진공게이지압 평균', 'C', A['vacuum']), ('배출가스 온도 평균', 'D', A['ts']), ('동압 평균', 'E', A['dp']), ('오리피스압 평균', 'F', A['orifice']),
                             ('가스미터 채취량(m3)', 'G', A['volumeM3']), ('가스미터 온도 평균', 'H', A['tm']), ('임핀저 출구온도 평균', 'K', A['impingerTemp'])]:
            cmp(tag, '기록지 평균', name, rv[f'{col}{ar}'].value, a)
        cmp(tag, '기록지 평균', '본 측정 정압(절대)', Pa + rv['J12'].value, A['spAbs']); cmp(tag, '기록지 평균', '표준 채취량(Sm3)', rv[f'G{ar+4}'].value, A['vmStdM3'])
        cmp(tag, '결과', '등속흡입계수(기록지)', rv['C13'].value, R['isokinetic'], 2); cmp(tag, '결과', '수분량(기록지)', rv['J9'].value, R['xwReport'])
        rows = D['rows']
        for i in range(1, ar - 17):
            x = 17 + i
            if i >= len(rows): break
            cmp(tag, '기록지 줄별', f'{x}행 오리피스압', rv[f'F{x}'].value, F['rowOrifice'][i])
            cmp(tag, '기록지 줄별', f'{x}행 유속', rv[f'M{x}'].value, rows[i]['velocity']); cmp(tag, '기록지 줄별', f'{x}행 수분량(mL)', rv[f'O{x}'].value, rows[i]['moisture'])
            cmp(tag, '기록지 줄별', f'{x}행 등속흡입계수', rv[f'L{x}'].value, rows[i]['isokinetic'], 2)
    # ── 계산식
    cmp(tag, '결과', '수분량 Xw', nv('Xw'), R['xwReport']); cmp(tag, '결과', '배출가스 밀도', nv('rr'), R['density']); cmp(tag, '결과', '유속', nv('Ve'), R['velocity'])
    cmp(tag, '결과', '건조 유량 Qa', nv('Qa'), R['qa'], 2)
    of = R['o2Factor']
    cmp(tag, '결과', '표준산소 환산 유량 Q', nv('Q'), (round(R['qa'], 2) / of) if (R['qa'] is not None and of) else None)
    if not flow:
        cmp(tag, '결과', '등속흡입계수(계산식)', nv('I'), R['isokinetic'], 2)
        if tag.startswith('dust') or tag.startswith('metal'):
            cmp(tag, '결과', '먼지농도 Ca', nv('Ca'), R['ca']); cmp(tag, '결과', '표준산소 환산 농도', nv('DC'), R['dc'])
print('\n묶음별:'); [print(f'  {k}: {v[0]}항목 중 불일치 {v[1]}') for k, v in groups.items()]
print(f'전체 {tot}항목, 불일치 {bad}')
