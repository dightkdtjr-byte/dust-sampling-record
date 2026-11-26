# 완성본 허가배출기준 초과·산소 보정 연동 시험
#   python3 -I tools/labform/excess_test.py <완성본 폴더> <master.json> <작업 폴더>
# 경우 (허가배출기준 칸에 넣는 값 / 계산값, u = 최종 자릿수 한 단위)
#   X1 "10" / 10        → 값, 초과 아님 (같으면 초과 아님)
#   X2 "10" / 10 + u    → "(초과)"
#   X3 "10" / 10 - u    → 초과 아님
#   O2 "10(12)" / 그대로 → 연결한 표준산소 칸이 모두 12, 결과 칸 오류 없음
# 그리고 결과 칸마다 "초과" 빨간 조건부 서식이 있는지 XML에서 본다
import sys, os, re, json, subprocess, zipfile, warnings
import openpyxl
warnings.filterwarnings('ignore')
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from apply_master import permit_cell
import extract_master as EM

golden, master, work = sys.argv[1], sys.argv[2], sys.argv[3]
rows = json.load(open(master, encoding='utf8'))
os.makedirs(work, exist_ok=True)
CASES = {'X1': ('10', 0), 'X2': ('10', 1), 'X3': ('10', -1), 'O2': ('10(12)', None)}
plan = {}
for r in rows:
    for res in r['result'].split(', '):
        plan.setdefault(r['file'], []).append(dict(sheet=r['sheet'], result=res, final=int(r['final']), item=r['item']))

o2_cells = {}   # file -> [(sheet, cell)]  : 표준산소로 연결한 칸 (수식에 '괄호 안 숫자' 식이 들어간 칸)
cf_ok = cf_bad = 0
for f, items in plan.items():
    z = zipfile.ZipFile(os.path.join(golden, f))
    wb0 = openpyxl.load_workbook(os.path.join(golden, f), keep_vba=f.endswith('m'))
    for ws in wb0.worksheets:
        for row in ws.iter_rows():
            for c in row:
                if isinstance(c.value, str) and c.value.startswith('=') and 'VALUE(MID(TRIM(SUBSTITUTE(' in c.value:
                    o2_cells.setdefault(f, []).append((ws.title, c.coordinate))
    # 조건부 서식
    wbx = z.read('xl/workbook.xml').decode(); rels = z.read('xl/_rels/workbook.xml.rels').decode()
    for it in items:
        rid = re.search(rf'<sheet name="{re.escape(it["sheet"])}"[^>]*r:id="([^"]+)"', wbx.replace('&amp;', '&')).group(1)
        tgt = re.search(rf'Id="{rid}"[^>]*Target="([^"]+)"|Target="([^"]+)"[^>]*Id="{rid}"', rels)
        part = 'xl/' + (tgt.group(1) or tgt.group(2)).replace('/xl/', '')
        x = z.read(part).decode()
        has = re.search(rf'<conditionalFormatting sqref="{it["result"]}"><cfRule type="containsText"[^>]*text="초과"', x)
        it['cf'] = bool(has)
        cf_ok += bool(has); cf_bad += (not has)

LO_DIRS = []
for case, (permit, k) in CASES.items():
    d = os.path.join(work, case); os.makedirs(d, exist_ok=True)
    for f, items in plan.items():
        path = os.path.join(golden, f)
        wb = openpyxl.load_workbook(path, keep_vba=f.endswith('m'))
        for it in items:
            ws = wb[it['sheet']]
            fm = str(ws[it['result']].value)
            P = None
            m = re.search(r'TRIM\(SUBSTITUTE\((\$?[A-Z]{1,3}\$?\d+)&""', fm)
            if m:
                P = m.group(1)
            it['permit'] = P
            if not P:
                continue
            ws[P].value = permit
            if k is None:
                continue
            mr = re.search(r"ROUND\((?:ROUND|TRUNC\(ROUND)\(IFERROR\(VALUE\(LEFT\((\$?[A-Z]{1,3}\$?\d+)&\"\"", fm)
            mr2 = re.search(r"ROUND\((?:ROUND|TRUNC\(ROUND)\(IFERROR\(VALUE\(LEFT\(IF\(((?:'[^']+'|[^!(),\s]+)!)?(\$?[A-Z]{1,3}\$?\d+)=\"\"", fm)
            u = 10 ** (-it['final'])
            v = round(10 + k * u, it['final'] + 3)
            if mr:
                ws[mr.group(1).replace('$', '')].value = v; it['set'] = True
            elif mr2:
                sh = (mr2.group(1) or '').rstrip('!').strip("'") or it['sheet']
                wb[sh][mr2.group(2).replace('$', '')].value = v; it['set'] = True
            else:
                it['set'] = False
        wb.save(os.path.join(d, os.path.splitext(f)[0] + ('.xlsm' if f.endswith('m') else '.xlsx')))
    LO_DIRS.append(d)

import lo_pool
lo_pool.run(LO_DIRS)   # 여러 폴더를 동시에 다시 계산
ok = bad = 0
expect = {'X1': False, 'X2': True, 'X3': False}
for f, items in plan.items():
    base = os.path.splitext(f)[0] + '.xlsx'
    for it in items:
        line = []
        if not it.get('permit'):
            print(f"-  {it['item'][:10]:10} {it['result']:4} 허가배출기준 칸 없음 (초과 판정 안 함)"); continue
        for case in ('X1', 'X2', 'X3'):
            if not it.get('set', True):
                line.append(f'{case}:계산값이 식'); continue
            v = openpyxl.load_workbook(os.path.join(work, case, 'lo', base), data_only=True)[it['sheet']][it['result']].value
            got = isinstance(v, str) and '초과' in v
            good = got == expect[case]
            ok += good; bad += (not good)
            line.append(f'{case}:{str(v):>12}{"" if good else " ✖"}')
        v = openpyxl.load_workbook(os.path.join(work, 'O2', 'lo', base), data_only=True)[it['sheet']][it['result']].value
        err = isinstance(v, str) and v.startswith('#')
        ok += (not err); bad += err
        line.append(f'O2:{str(v)[:12]:>12}{" ✖오류" if err else ""}')
        line.append('서식O' if it['cf'] else '서식X ✖')
        print(f"{'✔' if '✖' not in ''.join(line) else '✖'} {it['item'][:10]:10} {it['result']:4} 허가 {it['permit']:4} | " + ' | '.join(line))
# 표준산소 칸 = 12
for f, cells in o2_cells.items():
    base = os.path.splitext(f)[0] + '.xlsx'
    wbv = openpyxl.load_workbook(os.path.join(work, 'O2', 'lo', base), data_only=True)
    vals = [(s, c, wbv[s][c].value) for s, c in cells]
    good = all(v in (12, 12.0) for _, _, v in vals)
    ok += good; bad += (not good)
    print(f"{'✔' if good else '✖'} 표준산소 연동 {f[6:20]:14} " + ', '.join(f'{s[:8]}!{c}={v}' for s, c, v in vals[:6]) + (' …' if len(vals) > 6 else ''))
print(f'초과·산소 보정 시험: 맞음 {ok}, 틀림 {bad} / 조건부 서식 있음 {cf_ok}, 없음 {cf_bad}')
