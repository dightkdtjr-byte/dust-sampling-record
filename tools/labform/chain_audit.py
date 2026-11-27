# 검증 방식 15: 결과 칸이 읽는 값의 경로(앞 칸 → 그 앞 칸 …)를 끝까지 따라가 경로 위의 반올림을 찾는다.
# 맺음은 결과 칸에서 한 번: 경로 위에 ROUND(x, n)이 계산 자릿수+2 이하로 있으면(TRUNC(ROUND(x, n+5), n) 꼴은 뺌) 반올림이 두 번 되는 것
#   python3 -I tools/labform/chain_audit.py <기록부 폴더> <master.json>
import sys, os, re, json, glob, collections, warnings
warnings.filterwarnings('ignore'); import openpyxl
from openpyxl.utils import range_boundaries, get_column_letter as CL
G, M = sys.argv[1], sys.argv[2]
rows = json.load(open(M, encoding='utf8'))
ref = re.compile(r"(?:(?:'([^']+)'|([^\s!(),=<>&\"'+*/:^-]+))!)?\$?([A-Z]{1,3})\$?(\d+)(?::\$?([A-Z]{1,3})\$?(\d+))?")
def precedents(f, sheet, names):
    out = []
    for m in ref.finditer(re.sub(r'"[^"]*"', '', f)):
        sh = (m.group(1) or m.group(2) or sheet).strip()
        if sh not in names or sh == '기준': continue
        if m.group(5):
            c1, r1, c2, r2 = range_boundaries(f'{m.group(3)}{m.group(4)}:{m.group(5)}{m.group(6)}')
            if (c2 - c1 + 1) * (r2 - r1 + 1) > 400: continue
            out += [(sh, f'{CL(c)}{r}') for c in range(c1, c2 + 1) for r in range(r1, r2 + 1)]
        else: out.append((sh, m.group(3) + m.group(4)))
    return out
def rounds(f):
    """수식 안의 ROUND(…, n) 가운데 TRUNC(ROUND(…)) 보호 꼴이 아닌 것의 n"""
    out = []
    for m in re.finditer(r'ROUND\(', f):
        if f[max(0, m.start() - 6):m.start()] == 'TRUNC(' or f[max(0, m.start() - 1):m.start()] in ('T',): continue
        d, j = 1, m.end()
        while j < len(f) and d:
            d += f[j] == '('; d -= f[j] == ')'; j += 1
        arg = f[m.end():j - 1]; k = arg.rfind(',')
        n = arg[k + 1:].strip()
        out.append(int(n) if re.fullmatch(r'-?\d+', n) else n)
    return out
bad = []; n_cells = 0; wbs = {}
for r in rows:
    p = os.path.join(G, r['file']); wb = wbs.setdefault(p, openpyxl.load_workbook(p)); names = set(wb.sheetnames)
    calc = int(r['calc'])
    for res in str(r['result']).split(','):
        start = (r['sheet'], res.strip()); seen = {start}; q = collections.deque()
        f0 = wb[start[0]][start[1]].value
        q.extend(precedents(str(f0), start[0], names))
        while q:
            sh, c = q.popleft()
            if (sh, c) in seen: continue
            seen.add((sh, c)); v = wb[sh][c].value
            if not (isinstance(v, str) and v.startswith('=')): continue
            n_cells += 1
            for n in rounds(v):
                if not isinstance(n, int) or n <= calc + 2:
                    bad.append((r['file'][6:], r['item'], f'{sh}!{c}', n, calc, v[:90]))
            q.extend(precedents(v, sh, names))
uniq = {}
for b in bad: uniq.setdefault((b[0], b[2]), b)
for b in list(uniq.values())[:25]: print(f'   ✖ {b[0]} {b[1]}: {b[2]} ROUND {b[3]}자리 (계산 {b[4]}자리) {b[5]}')
print(f'결과 경로 반올림: 경로 위 수식 {n_cells}칸, 중간 반올림 {len(uniq)}칸')
