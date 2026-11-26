# 검증 방식 3: 무작위 대량 비교 — 판정 규칙을 파이썬으로 따로 구현하고, 무작위 계산값·허가배출기준을 넣어
# LibreOffice로 다시 계산한 기록부 결과와 한 칸씩 견준다.
#   python3 -I tools/labform/random_test.py <완성본 폴더> <master.json> <작업 폴더> [회수=8] [씨앗=7]
# 파이썬 쪽 규칙 (엑셀 수식을 보지 않고 README 규칙만으로 구현)
#   v1 = 버림(x, 계산 자릿수), 판정값 = 4사5입(v1, 판정 자릿수), 표시값 = 4사5입(v1, 최종 자릿수)  (맺음은 한 번)
#   판정 자릿수 = max(최종, 방법검출한계 소수 자릿수)
#   판정값 < 방법검출한계 → "불검출"; 아니면 표시값을 최종 자릿수로 쓴 글자 + (표시값 > 허가배출기준 괄호 앞 숫자면 " (초과)")
import sys, os, re, json, random, subprocess, warnings
from decimal import Decimal, ROUND_HALF_UP
import openpyxl
warnings.filterwarnings('ignore')

golden, master, work = sys.argv[1], sys.argv[2], sys.argv[3]
N = int(sys.argv[4]) if len(sys.argv) > 4 else 8
random.seed(int(sys.argv[5]) if len(sys.argv) > 5 else 7)
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


cases = []      # (file, sheet, result, raw cell sheet, raw cell, permit cell, x, permit, calc, final, mdl)
def mdl_of(r):
    # 다른 시트 값을 쓰는 줄(암모니아 BA·전체: ='기준'!G2)은 그 줄의 값을 따라간다
    v = r['mdl']
    m = re.fullmatch(r"='기준'!\$?G\$?(\d+)", str(v))
    if m:
        src = next((x for x in rows if x['file'] == r['file'] and x['row'] == int(m.group(1))), None)
        v = src['mdl'] if src else None
    try:
        return float(v)
    except (TypeError, ValueError):
        return None


for r in rows:
    mdl = mdl_of(r)
    calc, final = int(r['calc']), int(r['final'])
    wb = openpyxl.load_workbook(os.path.join(golden, r['file']), keep_vba=r['file'].endswith('m'))
    for res in r['result'].split(', '):
        fm = str(wb[r['sheet']][res].value)
        m = re.search(r"ROUND\((?:ROUND|TRUNC\(ROUND)\(IFERROR\(VALUE\(LEFT\((\$?[A-Z]{1,3}\$?\d+)&\"\"", fm)
        m2 = re.search(r"ROUND\((?:ROUND|TRUNC\(ROUND)\(IFERROR\(VALUE\(LEFT\(IF\(((?:'[^']+'|[^!(),\s]+)!)?(\$?[A-Z]{1,3}\$?\d+)=\"\"", fm)
        mp = re.search(r'TRIM\(SUBSTITUTE\((\$?[A-Z]{1,3}\$?\d+)&""', fm)
        if m:
            rsh, rc = r['sheet'], m.group(1).replace('$', '')
        elif m2:
            rsh, rc = (m2.group(1) or '').rstrip('!').strip("'") or r['sheet'], m2.group(2).replace('$', '')
        else:
            continue
        base = mdl if mdl else 1.0
        for i in range(N):
            # 검출한계 근처·반올림 경계(…5)·허가배출기준 근처·큰 값 고루
            kind = i % 4
            u = 10 ** -max(final, decs(mdl) if mdl else final)
            if kind == 0:
                x = base + random.randint(-3, 3) * u + random.choice([0, 0.5, -0.5, 0.49, 0.51]) * u
            elif kind == 1:
                x = round(random.uniform(0, base * 3), calc + 2)
            elif kind == 2:
                x = round(random.uniform(base, base * 50), calc + 1) + 0.5 * 10 ** -calc   # 계산 자릿수 반올림 경계
            else:
                x = round(random.uniform(0, base * 200), calc + 3)
            x = max(0.0, round(x, calc + 4))
            Lv = round(random.uniform(base, base * 60), final)
            permit = random.choice([f'{Lv:g}', f'{Lv:g}({random.choice([4, 6, 12, 13])})', f'{x:.{final}f}', '해당없음'])
            cases.append(dict(file=r['file'], sheet=r['sheet'], result=res, rsh=rsh, rc=rc, pc=mp.group(1).replace('$', '') if mp else None,
                              x=x, permit=permit, calc=calc, final=final, mdl=mdl, item=r['item']))

# 회차별 파일: i번째 시도끼리 한 파일에 (같은 파일 안 여러 결과 칸은 서로 다른 계산값 칸)
os.makedirs(work, exist_ok=True)
groups = {}
for c in cases:
    groups.setdefault((c['file'], c['result'], c['sheet']), []).append(c)
LO_DIRS = []
for i in range(N):
    d = os.path.join(work, f't{i}'); os.makedirs(d, exist_ok=True)
    byfile = {}
    for k, lst in groups.items():
        if i < len(lst):
            byfile.setdefault(k[0], []).append(lst[i])
    for f, lst in byfile.items():
        wb = openpyxl.load_workbook(os.path.join(golden, f), keep_vba=f.endswith('m'))
        used = set()
        for c in lst:
            if (c['rsh'], c['rc']) in used:
                continue        # 같은 계산값 칸을 쓰는 결과 칸(불소 D17·M19 등)은 한 시도에 하나만 — 덮어쓰기 방지
            used.add((c['rsh'], c['rc']))
            wb[c['rsh']][c['rc']].value = c['x']
            if c['pc']:
                # 허가배출기준에 괄호(표준산소)를 넣으면 산소 보정 값으로 바뀌므로, 초과 비교는 보정 없는 경우만 → 괄호 있는 경우는 표준산소 연동만 확인
                wb[c['sheet']][c['pc']].value = c['permit'] if '(' not in c['permit'] else c['permit'].split('(')[0]
            c['t'] = i
        wb.save(os.path.join(d, os.path.splitext(f)[0] + ('.xlsm' if f.endswith('m') else '.xlsx')))
    LO_DIRS.append(d)

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import lo_pool
lo_pool.run(LO_DIRS)   # 여러 폴더를 동시에 다시 계산
ok = bad = 0; fails = []
cache = {}
for c in cases:
    if 't' not in c:
        continue
    p = os.path.join(work, f"t{c['t']}", 'lo', os.path.splitext(c['file'])[0] + '.xlsx')
    if p not in cache:
        cache[p] = openpyxl.load_workbook(p, data_only=True)
    got = cache[p][c['sheet']][c['result']].value
    permit = c['permit'].split('(')[0] if '(' in c['permit'] else c['permit']
    exp = expected(c['x'], c['calc'], c['final'], c['mdl'], permit if c['pc'] else '')
    got_s = str(got if got is not None else '')
    if got_s.replace('(12)', '') .strip() == exp or got_s == exp:
        ok += 1
    else:
        bad += 1; fails.append(f"{c['item']} {c['result']} x={c['x']} 허가={permit} 계산{c['calc']}·최종{c['final']} MDL={c['mdl']}: 엑셀 {got_s!r} ≠ 파이썬 {exp!r}")
print(f'무작위 대량 비교: {ok + bad}건 중 같음 {ok}, 다름 {bad}')
for f in fails[:40]:
    print('   ✖', f)
