# 검증 방식 10: 프로그램 결과 계산 ↔ 파이썬 규칙 — 무작위 3000건(음수·지수 표기·정확히 5·큰 값)을 프로그램의 evaluate()와
# 파이썬 Decimal로 따로 구현한 규칙(계산 자릿수까지 버림 → 최종 자릿수 한 번 4사5입, 방법검출한계 미만 불검출, 없으면 0 미만)에 넣어 견준다
#   python3 -I tools/labform/prop_test.py <시험기록부.html> <작업 폴더> [건수=3000] [씨앗=11]
import sys, os, json, random, subprocess
from decimal import Decimal, ROUND_HALF_UP, ROUND_DOWN

html, work = sys.argv[1], sys.argv[2]
N = int(sys.argv[3]) if len(sys.argv) > 3 else 3000
random.seed(int(sys.argv[4]) if len(sys.argv) > 4 else 11)
os.makedirs(work, exist_ok=True)
rhu = lambda x, d: Decimal(str(x)).quantize(Decimal(1).scaleb(-d), rounding=ROUND_HALF_UP)
trunc = lambda x, d: rhu(x, d + 5).quantize(Decimal(1).scaleb(-d), rounding=ROUND_DOWN)
cases = []
for i in range(N):
    final = random.randint(0, 6); calc = final + 1
    mdl = random.choice([None, round(random.uniform(0.001, 5), random.randint(1, 4))])
    k = random.random()
    if k < 0.25: x = round(random.uniform(-2, 50), calc + random.randint(0, 3))
    elif k < 0.5: x = (random.randint(0, 999) + 0.5) / 10 ** final + random.choice([0, 0, -0.05, 0.05, -0.005]) / 10 ** final
    elif k < 0.6: x = float(f'{random.uniform(1, 9):.3f}e-{random.randint(3, 9)}')
    elif k < 0.7: x = -round(random.uniform(0, 3), calc + 2)
    else: x = round(random.uniform(0, 1e5), random.randint(0, 8))
    raw = random.choice([repr(x), f' {x!r} ', f'{x:.6E}' if k < 0.6 else repr(x)])
    cases.append(dict(calc=calc, final=final, mdl=mdl, raw=raw))
cf, of = os.path.join(work, 'cases.json'), os.path.join(work, 'js.json')
json.dump(cases, open(cf, 'w'))
here = os.path.dirname(os.path.abspath(__file__))
subprocess.run(['node', os.path.join(here, 'verify_js', 'prop.cjs'), os.path.abspath(html), cf, of], check=True, timeout=600)
js = json.load(open(of))
bad = []
for c, g in zip(cases, js):
    x = float(c['raw'].strip()); vf = rhu(trunc(x, c['calc']), c['final'])
    lim = Decimal(str(c['mdl'])) if c['mdl'] is not None else Decimal(0)
    e = '불검출' if vf < lim else f"{vf:.{c['final']}f}"
    if e.startswith('-') and Decimal(e) == 0:
        e = e[1:]
    if g != e:
        bad.append((c, g, e))
print(f'프로그램 계산 ↔ 파이썬 규칙: {len(cases)}건 중 같음 {len(cases) - len(bad)}, 다름 {len(bad)}')
for b in bad[:20]:
    print('   ✖', b)
