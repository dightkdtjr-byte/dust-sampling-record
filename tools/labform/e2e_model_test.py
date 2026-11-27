# 검증 방식 16: 처음부터 끝까지 — 엑셀 수식을 보지 않고 공정시험기준 식으로 따로 만든 계산기(아래 model_*)와 견준다.
# 원자료 칸(기기 농도·무게·시료채취량·온도·압력·산소)에 무작위 값을 넣어 LibreOffice로 다시 계산한 결과 칸 글자 = 계산기 결과
# (표준상태 환산 Vs = V × 273/(273+t) × (Pa+Pm)/760, 농도식은 각 ES 8.0 결과보고, 맺음은 총칙: 계산 자릿수 버림 → 최종 자릿수 4사5입)
#   python3 -I tools/labform/e2e_model_test.py <기록부 폴더> <master.json> <작업 폴더> [경우 수=6]
import sys, os, re, json, random, shutil, glob, warnings
from decimal import Decimal, ROUND_HALF_UP, ROUND_DOWN
warnings.filterwarnings('ignore'); import openpyxl
here = os.path.dirname(os.path.abspath(__file__)); sys.path.insert(0, here)
from tplkit import Book
import lo_pool
G, MASTER, W = sys.argv[1:4]; NC = int(sys.argv[4]) if len(sys.argv) > 4 else 6
random.seed(16)
rows = json.load(open(MASTER, encoding='utf8'))
CR = {}
for r in rows:
    own = next(x for x in rows if x['file'] == r['file'] and not x.get('linked'))
    src = own if r.get('linked') else r
    try: mdl = float(src['mdl'])
    except (TypeError, ValueError): mdl = None
    k_ = (r['file'], r['sheet']); CR[k_] = (int(r['calc']), int(r['final']), mdl, (CR[k_][3] if k_ in CR else []) + [x.strip() for x in str(r['result']).split(',')])
rhu = lambda x, d: Decimal(repr(x)).quantize(Decimal(1).scaleb(-d), rounding=ROUND_HALF_UP)
def fin(x, calc, final, mdl):
    v = rhu(x, calc + 5).quantize(Decimal(1).scaleb(-calc), rounding=ROUND_DOWN).quantize(Decimal(1).scaleb(-final), rounding=ROUND_HALF_UP)
    if v < (Decimal(repr(mdl)) if mdl is not None else 0): return '불검출'
    s = f'{v:.{final}f}'
    return s[1:] if s.startswith('-') and Decimal(s) == 0 else s
U = random.uniform
Vs = lambda V, t, Pa, Pm: V * 273 / (273 + t) * (Pa + Pm) / 760
mean = lambda l: sum(l) / len(l)
def o2():   # (표준산소, 실측산소) 또는 (None, None)
    return (random.choice([4, 6, 12, 15]), round(U(5, 18), 2)) if random.random() < 0.5 else (None, None)
corr = lambda c, s, m: c if s is None else c * (21 - s) / (21 - m)
gas = lambda: (round(U(10, 40), 2), round(U(5, 35), 1), random.randint(740, 775), round(U(5, 40), 1))   # V, t, Pa, 게이지압(mmH2O)
FILES = {os.path.basename(p): p for p in glob.glob(G + '/*.xls[xm]')}
WB = {f: openpyxl.load_workbook(p) for f, p in FILES.items()}
def raw_ref(f, sh, cell):   # =IF('Raw data'!I16="","",…) 꼴에서 원자료 칸
    m = re.search(r"(?:'([^']+)'|([A-Za-z가-힣()]+))!\$?([A-Z]+)\$?(\d+)", WB[f][sh][cell].value)
    return (m.group(1) or m.group(2)), m.group(3) + m.group(4)

def case():
    S, E = [], []      # 넣을 값 (파일, 시트, 칸, 값), 기대 (파일, 시트, 농도)
    put = lambda f, sh, kv: S.extend((f, sh, k, v) for k, v in kv.items())
    # 먼지 (ES 01301.1): md/Vs, md = 채취 후−전 − 현장바탕
    f = '시험기록부_먼지.xlsx'; V, t, Pa, dH = round(U(0.8, 1.5), 4), round(U(10, 30), 1), random.randint(745, 772), round(U(20, 60), 2)
    a0, d0 = round(U(9, 9.3), 4), round(U(9, 9.3), 4); gain, fb = round(U(0.002, 0.03), 4), round(U(0, 0.0008), 4); s, m = o2()
    put(f, '계산과정', {**{f'A{r}': a0 for r in (11, 12, 13)}, **{f'B{r}': round(a0 + gain, 4) for r in (11, 12, 13)}, **{f'D{r}': d0 for r in (11, 12, 13)},
                    **{f'E{r}': round(d0 + fb, 4) for r in (11, 12, 13)}, 'D3': V, 'F3': t, 'G3': Pa, 'H3': dH, 'L3': s, 'M3': m})
    E.append((f, '시험기록부', corr(((round(a0 + gain, 4) - a0) - (round(d0 + fb, 4) - d0)) * 1000 / Vs(V, t, Pa, dH / 13.6), s, m)))
    # 사이안화수소 (ES 01312.1 식 7): (a−b)×10/Vs × 22.4/26.017, a = 농도 × 분취량 25
    f = '시험기록부_사이안화수소.xlsx'; V, t, Pa, g = gas(); fbv = [round(U(0, 0.02), 4) for _ in range(3)]; sv = [round(U(0.03, 0.9), 4) for _ in range(3)]
    put(f, '계산과정', {**{f'B{3 + i}': fbv[i] for i in range(3)}, **{f'B{6 + i}': sv[i] for i in range(3)}, 'M3': V, 'L3': t, 'I3': Pa, 'J3': g})
    E.append((f, '시험기록부', (mean(sv) - mean(fbv)) * 25 * 10 / Vs(V, t, Pa, g / 13.6) * 22.4 / 26.017))
    # 황화수소 (ES 01310.1): 0.698 × (a−b) × 1000/Vs, a = 농도(mg/L) × 0.02 L × (200/20)
    f = '시험기록부_황화수소.xlsx'; V, t, Pa, g = gas(); fbv = [round(U(0, 0.01), 4) for _ in range(3)]; sv = [round(U(0.02, 0.8), 4) for _ in range(3)]; s, m = o2()
    put(f, '계산과정', {**{f'B{3 + i}': fbv[i] for i in range(3)}, **{f'B{6 + i}': sv[i] for i in range(3)}, 'K3': V, 'J3': t, 'G3': Pa, 'H3': g, 'M3': s, 'N3': m})
    E.append((f, '시험기록부', corr(0.698 * (mean(sv) - mean(fbv)) * 0.02 * 10 * 1000 / Vs(V, t, Pa, g / 13.6), s, m)))
    # 암모니아 (ES 01303.1): (a−b) × (250/10)/Vs, a = 농도(μL/mL) × 분취량 10. 붕산(4~9행)·과산화수소(10~15행), 전체 = 큰 값
    f = '시험기록부_암모니아.xlsx'; Pa = random.randint(745, 775); s, m = o2(); res = {}
    for name, r0 in (('시험기록부(BA)', 4), ('시험기록부 (H2O2)', 10)):
        V, t, _, g = gas(); fbv = [round(U(0, 0.005), 4) for _ in range(3)]; sv = [round(U(0.01, 0.4), 4) for _ in range(3)]
        put(f, '계산과정', {**{f'B{r0 + i}': fbv[i] for i in range(3)}, **{f'B{r0 + 3 + i}': sv[i] for i in range(3)}, f'M{r0}': V, f'L{r0}': t, f'J{r0}': g, f'O{r0}': s, f'P{r0}': m})
        res[name] = corr((mean(sv) - mean(fbv)) * 10 * 25 / Vs(V, t, Pa, g / 13.6), s, m); E.append((f, name, res[name]))
    put(f, '계산과정', {'I4': Pa}); E.append((f, '시험기록부 (전체)', max(res.values())))
    # 페놀 (ES 01503.2): 농도 × 10 × (200/10)/Vs, 현장바탕을 뺌
    f = '시험기록부_페놀화합물.xlsx'; V, t, Pa, g = gas(); fbv = [round(U(0, 0.01), 4) for _ in range(3)]; sv = [round(U(0.05, 0.9), 4) for _ in range(3)]
    put(f, '계산과정', {**{f'B{3 + i}': fbv[i] for i in range(3)}, **{f'B{6 + i}': sv[i] for i in range(3)}, 'J3': V, 'I3': t, 'F3': Pa, 'G3': g})
    E.append((f, '시험기록부', (mean(sv) - mean(fbv)) * 10 * 20 / Vs(V, t, Pa, g / 13.6)))
    # 알데하이드 (ES 01501.1): ((Aa−Ab) × V/Vs) × (22.4/Mw) × 분자량비, 시료 3개 평균
    for f, Mw, B in (('시험기록부_아세트알데하이드.xlsx', 44.05, 0.1965), ('시험기록부_폼알데하이드.xlsx', 30.03, 0.1429)):
        fbv = [round(U(0, 0.02), 4) for _ in range(3)]; Pa = random.randint(745, 775); cs = []; kv = {f'B{8 + i}': fbv[i] for i in range(3)}
        for i in range(3):
            sv = [round(U(0.1, 3), 4) for _ in range(3)]; V, t, pm = round(U(8, 14), 2), round(U(15, 32), 1), round(U(0.5, 2), 2)
            kv.update({f'B{11 + 3 * i + j}': sv[j] for j in range(3)}); kv.update({f'G{3 + i}': V, f'H{3 + i}': t, f'J{3 + i}': pm})
            if not str(WB[f]['계산과정'][f'I{3 + i}'].value).startswith('='): kv[f'I{3 + i}'] = Pa
            cs.append((mean(sv) - mean(fbv)) * 5 / Vs(V, t, Pa, pm) * 22.4 / Mw * B)
        put(f, '계산과정', kv); E.append((f, '시험기록부', mean(cs)))
    # 브로민·플루오린·염화수소 (ES 01502.3 식 5 등): (a−b) × 100/Vs × 22.4/M
    for f, M, has_o2 in (('시험기록부_브로민화합물.xlsx', 79.904, False), ('시험기록부_플루오린화합물.xlsx', 18.998, True), ('시험기록부_염화수소.xlsx', 35.453, True)):
        V, t, Pa, pm = round(U(15, 45), 2), round(U(5, 30), 1), random.randint(745, 775), round(U(0.5, 2), 2)
        fbv = [round(U(0, 0.01), 4) for _ in range(3)]; sv = [round(U(0.05, 2), 4) for _ in range(3)]; s, m = o2() if has_o2 else (None, None)
        put(f, 'rawdata', {**{f'H{14 + i}': fbv[i] for i in range(3)}, **{f'H{17 + i}': sv[i] for i in range(3)}})
        put(f, '계산과정', {'H3': V, 'I3': t, 'K3': Pa, 'J3': pm, **({'L3': s, 'M3': m} if has_o2 else {})})
        E.append((f, '시험기록부', corr((mean(sv) - mean(fbv)) * 100 / Vs(V, t, Pa, pm) * 22.4 / M, s, m)))
    # 수은 (ES 01408.1): (흡수액 2단 합 − 현장바탕 2단 합), 각 단 = Hg(ng)/분석량 × 희석배수 × 용액 부피/10^6 / Vs(Sm3)
    f = '시험기록부_수은화합물.xlsx'; ws = WB[f]['계산과정']; V, t, Pa, dH = round(U(0.8, 1.4), 4), round(U(10, 30), 1), random.randint(745, 772), round(U(5, 30), 2); s, m = o2()
    kv = {'I3': V, 'J3': t, 'K3': Pa, 'L3': dH, 'Q15': s, 'Q16': m}; tot = 0
    for r0, sign, lo, hi in ((3, 1, 0.2, 3), (6, 1, 0.2, 3), (9, -1, 0, 0.05), (12, -1, 0, 0.05)):
        v = [round(U(lo, hi), 4) for _ in range(3)]; kv.update({f'C{r0 + i}': v[i] for i in range(3)})
        tot += sign * mean([v[i] / ws[f'D{r0 + i}'].value * ws[f'E{r0 + i}'].value for i in range(3)]) * ws[f'H{r0}'].value / 1e6
    put(f, '계산과정', kv); E.append((f, '시험기록부', corr(tot / Vs(V, t, Pa, dH / 13.6), s, m)))
    # 비소 (ES 01401.3 식 3): (m1+m2)/Vs(L) × 22.41/74.92, m = (시료 − 바탕) 농도 × 용액 부피
    f = '시험기록부_비소화합물.xlsx'; sh = '계산과정(As)'; ws = WB[f][sh]; V, t, Pa, dH = round(U(0.8, 1.4), 4), round(U(10, 30), 1), random.randint(745, 772), round(U(5, 40), 2); s, m = o2(); mm = 0
    for col, vol in (('E', ws['E16'].value), ('K', ws['F16'].value)):
        sv = [round(U(0.05, 0.6), 4) for _ in range(3)]; bv = [round(U(0, 0.01), 4) for _ in range(3)]
        for i in range(3):
            rs, rc = raw_ref(f, sh, f'{col}{5 + i}'); S.append((f, rs, rc, sv[i])); rs, rc = raw_ref(f, sh, f'{col}{9 + i}'); S.append((f, rs, rc, bv[i]))
        mm += (mean(sv) - mean(bv)) * vol
    put(f, sh, {'E17': V, 'E18': t, 'E19': dH, 'E20': Pa, 'E22': m}); put(f, '시험기록부(As)', {'K18': s})
    E.append((f, '시험기록부(As)', corr(mm / (Vs(V, t, Pa, dH / 13.6) * 1000), s, m) * 22.41 / 74.92))
    # 중금속 7종 (ES 01405.2 등): (시료 − 바탕) 농도(μg/mL) × 100 mL / Vs(L)
    f = '시험기록부_중금속.xlsm'; ws = WB[f]['계산과정']; V, t, Pa, dH = round(U(0.8, 1.4), 4), round(U(10, 30), 1), random.randint(745, 772), round(U(5, 60), 2); m_ = round(U(5, 18), 2)
    put(f, '계산과정', {'C18': V, 'C19': t, 'C20': dH, 'C21': Pa, 'C23': m_})
    for col, name in zip('CDEFGHI', ('구리화합물', '납화합물', '아연화합물', '크로뮴화합물', '니켈화합물', '카드뮴화합물', '베릴륨화합물')):
        sv = [round(U(0.005, 0.4), 4) for _ in range(3)]; bv = [round(U(0, 0.002), 4) for _ in range(3)]; s = random.choice([None, 6, 12])
        for i in range(3):
            rs, rc = raw_ref(f, '계산과정', f'{col}{11 + i}'); S.append((f, rs, rc, sv[i])); rs, rc = raw_ref(f, '계산과정', f'{col}{6 + i}'); S.append((f, rs, rc, bv[i]))
        put(f, name, {'K17': s}); E.append((f, name, corr((mean(sv) - mean(bv)) * 100 / (Vs(V, t, Pa, dH / 13.6) * 1000), s, m_)))
    # VOCs 12종 (ES 01511.1 등): 기기 농도(ppb) × 200/주입량 − 바탕, 시료 3개 평균, ppm
    f = '시험기록부_VOCs.xlsm'
    for (ff, sh), _ in [(k, v) for k, v in CR.items() if k[0] == f]:
        bl = [round(U(0, 2), 3) for _ in range(2)]; kv = {'O17': bl[0], 'O18': bl[1]}; cs = []
        for i in range(3):
            v = [round(U(5, 400), 3) for _ in range(3)]; inj = random.choice([200, 100, 50, 20]); kv.update({f'O{19 + 3 * i + j}': v[j] for j in range(3)}); kv[f'M{53 + 3 * i}'] = inj
            cs.append(mean(v) * 200 / inj - mean(bl))
        put(f, sh, kv); E.append((f, sh, mean(cs) / 1000))
    # 이황화탄소 (ppm)·이황화메틸 (ppb): 시료 평균 − 현장바탕 평균
    for f, sh, lo, hi in (('시험기록부_CS2.xlsx', '시험기록부-이황화탄소', 0.1, 3), ('시험기록부_이황화메틸.xlsx', '시험기록부-이황화메틸', 0.02, 9)):
        bv = [round(U(0, lo / 5), 4) for _ in range(3)]; sv = [round(U(lo, hi), 4) for _ in range(3)]
        put(f, sh, {**{f'O{16 + i}': bv[i] for i in range(3)}, **{f'O{19 + i}': sv[i] for i in range(3)}}); E.append((f, sh, mean(sv) - mean(bv)))
    # 벤조(a)피렌 (ES 01525.1a 식 9·10·13·14, 원문 PDF의 식을 보고 확인): 농도 = (As/Ais) × (Qis/RRF) / Vs, RRF = 표준용액들의 (At/Ais) × (Qis/Qt) 평균, 시료 평균 − 현장바탕 평균
    f = '시험기록부_PAHs.xlsx'; c = WB[f]['계산과정']; t_sh, i_sh = WB[f]['Benzo(a)pyrene'], WB[f]['d12-Benzo(a)pyrene']
    def name_of(cell):
        v = c[cell].value
        while isinstance(v, str) and re.fullmatch(r'=\$?[A-Z]+\$?\d+', v): v = c[v[1:].replace('$', '')].value
        return v
    row_of = {t_sh[f'B{r}'].value: r for r in range(16, 46) if t_sh[f'B{r}'].value}
    V, t, Pa, dH = round(U(2.5, 4), 4), round(U(10, 30), 1), random.randint(745, 772), round(U(10, 40), 1); put(f, '계산과정', {'A112': V, 'B112': t, 'C112': Pa, 'D112': dH})
    vs = Vs(V, t, Pa, dH / 13.6); rrfs = []; k0 = U(0.8, 1.3)
    for r in range(43, 51):
        nm = name_of(f'A{r}')
        if nm not in row_of: continue
        rr = row_of[nm]; qt, qis = c[f'D{rr - 9}'].value, c[f'F{rr - 9}'].value; ais = round(U(4e5, 6e5)); at = round(ais * qt / qis * k0 * U(0.97, 1.03))
        S.append((f, 'Benzo(a)pyrene', f'D{rr}', at)); S.append((f, 'd12-Benzo(a)pyrene', f'D{rr}', ais)); rrfs.append(at / ais * qis / qt)
    rrf = mean(rrfs); grp = []
    for r0, lo, hi in ((96, 0, 0.002), (99, 0.01, 0.2)):     # 현장바탕 3, 시료 3
        cs = []
        for r in range(r0, r0 + 3):
            rr = row_of[name_of(f'A{r}')]; qis = c[f'F{rr - 9}'].value; ais = round(U(4e5, 6e5)); at = round(ais * U(lo, hi))
            S.append((f, 'Benzo(a)pyrene', f'D{rr}', at)); S.append((f, 'd12-Benzo(a)pyrene', f'D{rr}', ais)); cs.append(at / ais * qis / rrf / vs)
        grp.append(mean(cs))
    E.append((f, '시험기록부', grp[1] - grp[0]))
    # 총탄화수소 (ES 01507.1): 측정 평균 × K(프로페인 3)
    f = '시험기록부_총탄화수소.xlsx'; v = [round(U(0, 40), 1) for _ in range(120)]
    put(f, 'Raw data', {f'C{14 + i}': v[i] for i in range(120)}); E.append((f, '시험기록부', mean(v) * 3))
    # SOx·NOx·CO (자동측정): 3회 평균, 표준산소 보정
    f = '시험기록부_SOx, CO, NOx.xlsx'; sh = '시험기록부-SOx, CO, NOx'; exp3 = []
    for k in range(3):
        v = [round(U(1, 90), 2) for _ in range(3)]; s, m = o2()
        put(f, '계산과정', {**{f'D{5 + 3 * k + i}': v[i] for i in range(3)}, **{f'F{5 + 3 * k + i}': s for i in range(3)}, **{f'G{5 + 3 * k + i}': m for i in range(3)}})
        exp3.append(mean([corr(x, s, m) for x in v]))
    E.append((f, sh, exp3))
    return S, E

if os.path.exists(W): shutil.rmtree(W)
ALL = []
for i in range(NC):
    S, E = case(); d = os.path.join(W, f'c{i}'); os.makedirs(d); by = {}
    for f, sh, c, v in S: by.setdefault(f, []).append((sh, c, v))
    for f, lst in by.items():
        b = Book(FILES[f])
        for sh, c, v in lst: b.setv(sh, c, v)
        b.finish(os.path.join(d, f))
    ALL.append((d, E))
lo_pool.run([d for d, _ in ALL])
ok = bad = 0; fails = []
clean = lambda v: re.sub(r'\([^()]*\)\s*$', '', str(v if v is not None else '').replace(' (초과)', '')).strip()
for d, E in ALL:
    cache = {}
    for f, sh, x in E:
        p = os.path.join(d, 'lo', os.path.splitext(f)[0] + '.xlsx'); wb = cache.setdefault(p, openpyxl.load_workbook(p, data_only=True))
        calc, final, mdl, cells = CR[(f, sh)]
        xs = x if isinstance(x, list) else [x] * len(cells)
        for cell, xv in zip(cells, xs):
            want = fin(xv, calc, final, mdl); got = clean(wb[sh][cell].value)
            if got == want: ok += 1
            else: bad += 1; fails.append(f'{os.path.basename(d)} {f[6:]} {sh}!{cell}: 엑셀 {got!r} ≠ 계산기 {want!r} (값 {xv:.8g})')
for x in fails[:40]: print('   ✖', x)
print(f'처음부터 끝까지(원문 식 계산기 ↔ 엑셀): {ok + bad}건 중 같음 {ok}, 다름 {bad}')
