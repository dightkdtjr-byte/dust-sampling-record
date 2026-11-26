# 검증 방식 9: 극단값 시험 — 결과 칸마다 경계·특이 입력을 넣고 LibreOffice로 다시 계산해 파이썬 규칙과 견준다.
#   정확히 5(올림), …45(버림), 두 맺음 방식이 갈리는 값, 이진 소수 함정(1.005·2.675꼴), 0, 아주 작은 값, 큰 값, 음수(바탕 보정),
#   앞뒤 빈칸 글자, 지수 표기 글자, "값(12)" 글자(표준산소 꼬리), 숫자 아닌 글자, 빈칸
#   python3 -I tools/labform/edge_test.py <완성본 폴더> <master.json> <작업 폴더>
# 파이썬 쪽 규칙 (엑셀 수식을 보지 않고 README 규칙만으로 구현)
#   v1 = 버림(x, 계산 자릿수), 판정값 = 4사5입(v1, 판정 자릿수), 표시값 = 4사5입(v1, 최종 자릿수)  (맺음은 한 번)
#   판정 자릿수 = max(최종, 방법검출한계 소수 자릿수)
#   판정값 < 방법검출한계 → "불검출"; 아니면 표시값을 최종 자릿수로 쓴 글자 + (표시값 > 허가배출기준 괄호 앞 숫자면 " (초과)")
import sys, os, re, json, random, subprocess, warnings
from decimal import Decimal, ROUND_HALF_UP
import openpyxl
warnings.filterwarnings('ignore')

golden, master, work = sys.argv[1], sys.argv[2], sys.argv[3]
rows = json.load(open(master, encoding='utf8'))


def rhu(x, d):
    return Decimal(str(x)).quantize(Decimal(1).scaleb(-d), rounding=ROUND_HALF_UP)


def trunc(x, d):
    """d자리까지 계산(아래 버림). 이진 오차를 지우려고 d+5자리에서 먼저 4사5입"""
    from decimal import ROUND_DOWN
    return rhu(x, d + 5).quantize(Decimal(1).scaleb(-d), rounding=ROUND_DOWN)


def decs(v):
    t = Decimal(str(v)).normalize().as_tuple()
    return max(0, -t.exponent)


def expected(x, calc, final, mdl, permit):
    jd = final      # 판정은 최종 자릿수로 반올림한 값 (표시값과 판정이 같도록)
    v1 = trunc(x, calc)
    vj, vf = rhu(v1, jd), rhu(v1, final)
    lim = Decimal(str(mdl)) if mdl is not None else Decimal(0)      # 엑셀 N("-") = 0: 방법검출한계가 없으면 0 미만만 불검출
    if vj < lim:
        return '불검출'
    s = f'{vf:.{final}f}'
    L = None
    m = re.match(r'\s*([\d.]+)', permit or '')
    if m:
        L = Decimal(m.group(1))
    if L is not None and vf > L:
        s += ' (초과)'
    return s




def mdl_of(r):
    v = r['mdl']
    m = re.fullmatch(r"='기준'!\$?G\$?(\d+)", str(v))
    if m:
        src = next((x for x in rows if x['file'] == r['file'] and x['row'] == int(m.group(1))), None)
        v = src['mdl'] if src else None
    try:
        return float(v)
    except (TypeError, ValueError):
        return None


def edge_values(calc, final, mdl):
    """(넣을 값, 파이썬 규칙에 넣을 숫자 또는 None(빈칸 기대), 꼬리 글자)"""
    u, uc = Decimal(1).scaleb(-final), Decimal(1).scaleb(-calc)
    m = Decimal(str(mdl)) if mdl else Decimal(1)
    q = lambda d: float(d)
    vals = [
        ('정확히 5', q(m + Decimal('0.5') * u)),
        ('…45 버림', q(m + Decimal('0.45') * u)),
        ('두 방식 갈림', q(m - Decimal('0.505') * u)),
        ('이진 함정 1', q(1 + Decimal('0.5') * u)),
        ('이진 함정 2', q(Decimal('2.67') + Decimal('0.5') * u) if final == 2 else q(Decimal(2) + Decimal('0.6') * u + Decimal('0.5') * u)),
        ('0', 0.0),
        ('아주 작은 값', 1e-9),
        ('큰 값', 12345.678949),
        ('음수', q(-(m * 0 + Decimal('0.4') * u + Decimal('0.05') * u))),
    ]
    out = [(lab, x, x, '') for lab, x in vals]
    out += [('앞뒤 빈칸', f'  {q(m + u)}  ', q(m + u), ''),
            ('지수 표기', f'{q(m * 3):.3E}', float(f'{q(m * 3):.3E}'), ''),
            ('값(12)', f'{q(m * 2 + Decimal("0.5") * uc)}(12)', q(m * 2 + Decimal('0.5') * uc), '(12)'),
            ('숫자 아님', 'N.D.', None, ''),
            ('빈칸', None, None, '')]
    return out


cases = []
for r in rows:
    mdl = mdl_of(r)
    calc, final = int(r['calc']), int(r['final'])
    wb = openpyxl.load_workbook(os.path.join(golden, r['file']), keep_vba=r['file'].endswith('m'))
    for res in r['result'].split(', '):
        fm = str(wb[r['sheet']][res].value)
        m = re.search(r"ROUND\((?:ROUND|TRUNC\(ROUND)\(IFERROR\(VALUE\(LEFT\((\$?[A-Z]{1,3}\$?\d+)&\"\"", fm)
        m2 = re.search(r"ROUND\((?:ROUND|TRUNC\(ROUND)\(IFERROR\(VALUE\(LEFT\(IF\(((?:'[^']+'|[^!(),\s]+)!)?(\$?[A-Z]{1,3}\$?\d+)=\"\"", fm)
        mp = re.search(r'TRIM\(SUBSTITUTE\((\$?[A-Z]{1,3}\$?\d+)&""', fm)
        fallback = False
        if m:
            rsh, rc = r['sheet'], m.group(1).replace('$', '')
        elif m2:
            fallback = True     # IF(보정값="", 보정 전 값, 보정값): 보정값 칸을 비우면 보정 전 값으로 계산하는 양식
            rsh, rc = (m2.group(1) or '').rstrip('!').strip("'") or r['sheet'], m2.group(2).replace('$', '')
        else:
            continue
        for i, (lab, put, x, tail) in enumerate(edge_values(calc, final, mdl)):
            cases.append(dict(file=r['file'], sheet=r['sheet'], result=res, rsh=rsh, rc=rc, pc=mp.group(1).replace('$', '') if mp else None,
                              i=i, label=lab, put=put, x=x, tail=tail, fallback=fallback, calc=calc, final=final, mdl=mdl, item=r['item']))

os.makedirs(work, exist_ok=True)
N = max(c['i'] for c in cases) + 1
LO_DIRS = []
for i in range(N):
    d = os.path.join(work, f't{i}'); os.makedirs(d, exist_ok=True)
    byfile = {}
    for c in cases:
        if c['i'] == i:
            byfile.setdefault(c['file'], []).append(c)
    for f, lst in byfile.items():
        wb = openpyxl.load_workbook(os.path.join(golden, f), keep_vba=f.endswith('m'))
        used = set()
        for c in lst:
            if (c['rsh'], c['rc']) in used:
                continue
            used.add((c['rsh'], c['rc']))
            wb[c['rsh']][c['rc']].value = c['put']
            if c['pc']:
                wb[c['sheet']][c['pc']].value = '해당없음'      # 초과 판정은 빼고 맺음·불검출만
            c['t'] = i
        wb.save(os.path.join(d, os.path.splitext(f)[0] + ('.xlsm' if f.endswith('m') else '.xlsx')))
    LO_DIRS.append(d)

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import lo_pool
lo_pool.run(LO_DIRS)
ok = bad = 0; fails = []; cache = {}; by = {}
for c in cases:
    if 't' not in c:
        continue
    p = os.path.join(work, f"t{c['t']}", 'lo', os.path.splitext(c['file'])[0] + '.xlsx')
    if p not in cache:
        cache[p] = openpyxl.load_workbook(p, data_only=True)
    got = cache[p][c['sheet']][c['result']].value
    got_s = '' if got is None else str(got)
    if c['put'] is None and c['fallback']:
        by.setdefault(c['label'] + '(보정 전 값 사용)', [0, 0])[0] += 1     # 빈칸이면 보정 전 값으로 계산 — 양식 설계대로
        ok += 1
        continue
    if c['x'] is None:
        exp = ''
    else:
        exp = expected(c['x'], c['calc'], c['final'], c['mdl'], '')
        if re.fullmatch(r'-0\.?0*', exp):
            exp = exp[1:]           # 엑셀은 -0을 0으로 표시
        if exp != '불검출':
            exp += c['tail']
    good = got_s.strip() == exp
    ok += good; bad += (not good)
    by.setdefault(c['label'], [0, 0])[0 if good else 1] += 1
    if not good:
        fails.append(f"{c['item']} {c['sheet']}!{c['result']} [{c['label']}] 넣은 값 {c['put']!r} → 엑셀 {got_s!r}, 규칙 {exp!r}")
print(f'극단값 시험: {ok + bad}건 중 같음 {ok}, 다름 {bad}')
for k, (g, b) in by.items():
    print(f'   {k:10} 같음 {g} 다름 {b}')
for f in fails[:40]:
    print('   ✖', f)
