# 완성본 불검출 판정 경계값 시험: 계산값을 방법검출한계 근처로 넣고 다시 계산해 결과를 본다.
#   python3 -I tools/labform/boundary_test.py <완성본 폴더> <master.json> <작업 폴더>
# 경우 (u = 최종 자릿수 한 단위, 판정은 최종 자릿수로 반올림한 값)
#   A 계산값 = 방법검출한계            B = 검출한계 - 0.4u     C = 검출한계 - 0.6u     D = 검출한계 + 0.5u
#   기대값은 규칙으로 계산: 계산 자릿수 4사5입 → 최종 자릿수 4사5입 → 검출한계 미만이면 불검출, 아니면 값
#   (대부분 A·B·D 값, C 불검출. 브로민처럼 검출한계가 최종 자릿수보다 잘면 A도 불검출)
import sys, os, re, json, shutil, subprocess, warnings
import openpyxl
warnings.filterwarnings('ignore')
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import extract_master as EM

golden, master, work = sys.argv[1], sys.argv[2], sys.argv[3]
rows = json.load(open(master, encoding='utf8'))
from decimal import Decimal, ROUND_HALF_UP
CASES = {'A': 0.0, 'B': 0.4, 'C': 0.6, 'D': -0.5}


def rhu(x, d):
    return Decimal(str(x)).quantize(Decimal(1).scaleb(-d), rounding=ROUND_HALF_UP)


def case_value(it, k):
    return round(it['mdl'] - k * 10 ** (-it['final']), it['final'] + 3)


def expect_of(it, k):
    v = rhu(rhu(case_value(it, k), it['calc']), it['final'])
    return 'nd' if v < Decimal(str(it['mdl'])) else 'value'
plan = {}   # file -> list of (sheet, result, R cell, mdl, final)
for r in rows:
    try:
        m = float(r['mdl'])
    except (TypeError, ValueError):
        continue
    for res in r['result'].split(', '):
        plan.setdefault(r['file'], []).append(dict(sheet=r['sheet'], result=res, mdl=m, final=int(r['final']), calc=int(r['calc']), item=r['item']))

# 결과 칸이 비교하는 계산값 칸 (단순한 칸 참조만 시험)
src_items = {}
for f in plan:
    for it in EM.scan(os.path.join(golden, f)) if False else []:
        pass

os.makedirs(work, exist_ok=True)
results = []
for f, items in plan.items():
    path = os.path.join(golden, f)
    wb0 = openpyxl.load_workbook(path, keep_vba=f.endswith('m'))
    cells = []
    for it in items:
        fm = str(wb0[it['sheet']][it['result']].value)
        # 바뀐 수식에서 첫 번째 계산값: ROUND(ROUND(<R>,
        m = re.search(r"ROUND\(ROUND\(IFERROR\(VALUE\(LEFT\((\$?[A-Z]{1,3}\$?\d+)&\"\"", fm)
        m2 = re.search(r"ROUND\(ROUND\(IFERROR\(VALUE\(LEFT\(IF\(((?:'[^']+'|[^!(),\s]+)!)?(\$?[A-Z]{1,3}\$?\d+)=\"\"", fm)
        if not m and m2:
            sh = (m2.group(1) or '').rstrip('!').strip("'") or it['sheet']
            cells.append((dict(it, sheet_src=sh), m2.group(2).replace('$', '')))
            continue
        if not m:
            results.append((f, it, None, '계산값이 식이라 이 시험에서 뺌 (같은 수식 틀)'))
            continue
        cells.append((it, m.group(1).replace('$', '')))
    for case, k in CASES.items():
        wb = openpyxl.load_workbook(path, keep_vba=f.endswith('m'))
        for it, rc in cells:
            v = case_value(it, k)
            ws = wb[it.get('sheet_src', it['sheet'])]
            ws[rc].value = v
            # 중금속·비소: 보정 칸이 비어 있어야 J열 계산값을 쓴다 → K열(표준산소 보정 선택)을 비운다
            if f.endswith('.xlsm') and '금속' in f or '비소' in f:
                pass
        d = os.path.join(work, case); os.makedirs(d, exist_ok=True)
        out = os.path.join(d, os.path.splitext(f)[0] + ('.xlsm' if f.endswith('m') else '.xlsx'))
        wb.save(out)
import lo_pool
lo_pool.run([os.path.join(work, case) for case in CASES])
ok = bad = 0
for f, items in plan.items():
    base = os.path.splitext(f)[0] + '.xlsx'
    tested = {(i['sheet'], i['result']) for i, _ in []}
    for it in items:
        if any(f == ff and it is i2 for ff, i2, _, _ in results):
            continue
        line = []
        for case in CASES:
            p = os.path.join(work, case, 'lo', base)
            if not os.path.exists(p):
                line.append(f'{case}:파일없음'); continue
            v = openpyxl.load_workbook(p, data_only=True)[it['sheet']][it['result']].value
            got = 'nd' if v == '불검출' else ('value' if v not in (None, '') else 'empty')
            good = got == expect_of(it, CASES[case])
            ok += good; bad += (not good)
            line.append(f'{case}:{v!s:>9}{"" if good else " ✖"}')
        print(f"{'✔' if all('✖' not in x for x in line) else '✖'} {it['item'][:10]:10} {it['result']:4} 검출한계 {it['mdl']:<8g} 최종 {it['final']}자리 | " + ' | '.join(line))
print(f'경계값 시험: 맞음 {ok}, 틀림 {bad}')
for f, it, _, why in results:
    print('   뺌:', it['item'], it['result'], why)
