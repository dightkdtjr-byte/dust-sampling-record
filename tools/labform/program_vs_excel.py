# 검증 방식 7-2: 프로그램 판정 ↔ 내려받은 완성본(엑셀) 판정
#   python3 -I tools/labform/program_vs_excel.py <작업 폴더(golden.zip, program_cases.json)>
import sys, os, re, json, zipfile, subprocess, warnings
import openpyxl
warnings.filterwarnings('ignore')
W = sys.argv[1]
cases = json.load(open(os.path.join(W, 'program_cases.json'), encoding='utf8'))
src = os.path.join(W, 'golden'); os.makedirs(src, exist_ok=True)
z = zipfile.ZipFile(os.path.join(W, 'golden.zip'))
for n in z.namelist():
    if n.lower().endswith(('.xlsx', '.xlsm')):
        open(os.path.join(src, os.path.basename(n)), 'wb').write(z.read(n))
# 시도 번호별로 파일 하나씩 (같은 계산값 칸을 쓰는 결과 칸은 한 시도에 하나만)
for c in cases:
    c['res0'] = c['result'].split(', ')[0]
by = {}
for c in cases:
    by.setdefault((c['file'], c['row']), []).append(c)
N = max(len(v) for v in by.values())
LO = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'lo.sh')
LO_DIRS = []
for i in range(N):
    d = os.path.join(W, f't{i}'); os.makedirs(d, exist_ok=True)
    files = {}
    for k, lst in by.items():
        if i < len(lst):
            files.setdefault(k[0], []).append(lst[i])
    for f, lst in files.items():
        wb = openpyxl.load_workbook(os.path.join(src, f), keep_vba=f.endswith('m'))
        used = set()
        for c in lst:
            fm = str(wb[c['sheet']][c['res0']].value)
            m = re.search(r"ROUND\(ROUND\(IFERROR\(VALUE\(LEFT\((\$?[A-Z]{1,3}\$?\d+)&\"\"", fm)
            m2 = re.search(r"ROUND\(ROUND\(IFERROR\(VALUE\(LEFT\(IF\(((?:'[^']+'|[^!(),\s]+)!)?(\$?[A-Z]{1,3}\$?\d+)=\"\"", fm)
            mp = re.search(r'TRIM\(SUBSTITUTE\((\$?[A-Z]{1,3}\$?\d+)&""', fm)
            if m:
                tgt = (c['sheet'], m.group(1).replace('$', ''))
            elif m2:
                tgt = ((m2.group(1) or '').rstrip('!').strip("'") or c['sheet'], m2.group(2).replace('$', ''))
            else:
                c['skip'] = '계산값이 식'; continue
            if tgt in used:
                c['skip'] = '같은 계산값 칸'; continue
            used.add(tgt)
            wb[tgt[0]][tgt[1]].value = c['x']
            if mp:
                wb[c['sheet']][mp.group(1).replace('$', '')].value = c['permit']
            else:
                c['nopermit'] = True
            c['t'] = i
        wb.save(os.path.join(d, os.path.splitext(f)[0] + ('.xlsm' if f.endswith('m') else '.xlsx')))
    LO_DIRS.append(d)
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import lo_pool
lo_pool.run(LO_DIRS)   # 여러 폴더를 동시에 다시 계산
ok = bad = skip = 0; fails = []; cache = {}
for c in cases:
    if 't' not in c:
        skip += 1; continue
    p = os.path.join(W, f"t{c['t']}", 'lo', os.path.splitext(c['file'])[0] + '.xlsx')
    cache[p] = cache.get(p) or openpyxl.load_workbook(p, data_only=True)
    got = str(cache[p][c['sheet']][c['res0']].value or '')
    prog = c['program']
    if c.get('nopermit'):
        prog = prog.replace(' (초과)', '')
    if got.replace('(12)', '') == prog or got == prog:
        ok += 1
    else:
        bad += 1; fails.append(f"{c['file'][6:16]} {c['sheet'][:10]} {c['res0']} x={c['x']} 허가={c['permit']}: 엑셀 {got!r} ≠ 프로그램 {prog!r}")
print(f'프로그램 ↔ 엑셀: {ok + bad}건 중 같음 {ok}, 다름 {bad} (뺌 {skip})')
for f in fails[:30]:
    print('   ✖', f)
