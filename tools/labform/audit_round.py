# 검증 방식 5: 수식 정적 감사 + 다른 계산 엔진(fast-formula-parser)용 시험 자료 만들기
#   python3 -I tools/labform/audit_round.py <골든 폴더> <master.json> <결과 폴더>
# 정적 감사
#   A 결과 수식이 자기 물질의 기준 줄(같은 행 번호)만 가리키는가 ('기준'!$E/$F/$K$행)
#   B 계산값에 고정 자릿수 ROUND(칸,숫자)가 남아 있지 않은가
#   C 결과 수식에 "불검출"이 있고, 허가배출기준이 있으면 "(초과)"와 기준값 비교가 있는가
#   D 방법검출한계·정량한계·시험방법 칸이 같은 행의 '기준'!$G·$H·$D를 가리키는가 (다른 시트에서 가져오는 칸 제외)
#   E 통합 문서의 어떤 수식에도 "#REF!"가 없는가
#   F 기준 시트 값이 기준표(master.json)와 같은가
#   G 결과 칸마다 "초과" 조건부 서식이 있는가
# 다른 엔진 시험 자료: 결과 칸마다 계산값 6개 × 허가배출기준 → cases.json (평가는 engine_eval.js)
import sys, os, re, json, zipfile, random, warnings
from decimal import Decimal, ROUND_HALF_UP
import openpyxl
warnings.filterwarnings('ignore')
golden, master, out = sys.argv[1:4]
os.makedirs(out, exist_ok=True)
rows = json.load(open(master, encoding='utf8'))
random.seed(11)
problems = []; checked = 0


def bad(msg):
    problems.append(msg)


def rhu(x, d):
    return Decimal(str(x)).quantize(Decimal(1).scaleb(-d), rounding=ROUND_HALF_UP)


def decs(v):
    return max(0, -Decimal(str(v)).normalize().as_tuple().exponent)


def expected(x, calc, final, mdl, L):
    jd = max(final, decs(mdl)) if mdl is not None else final
    v1 = rhu(x, calc); vj, vf = rhu(v1, jd), rhu(v1, final)
    if mdl is not None and vj < Decimal(str(mdl)):
        return '불검출'
    s = f'{vf:.{final}f}'
    if L is not None and vf > Decimal(str(L)):
        s += ' (초과)'
    return s


cases = []
for f in sorted({r['file'] for r in rows}):
    path = os.path.join(golden, f)
    wb = openpyxl.load_workbook(path, keep_vba=f.endswith('m'))
    lo = os.path.join(golden, 'lo', os.path.splitext(f)[0] + '.xlsx')
    wv = openpyxl.load_workbook(lo, data_only=True) if os.path.exists(lo) else None
    gs = wb['기준']
    # E: #REF!
    for ws in wb.worksheets:
        for row in ws.iter_rows():
            for c in row:
                if isinstance(c.value, str) and c.value.startswith('=') and '#REF!' in c.value:
                    bad(f'E {f} {ws.title}!{c.coordinate} 수식에 #REF!')
    z = zipfile.ZipFile(path)
    wbx = z.read('xl/workbook.xml').decode().replace('&amp;', '&'); rels = z.read('xl/_rels/workbook.xml.rels').decode()
    def part_of(sh):
        rid = re.search(rf'<sheet name="{re.escape(sh)}"[^>]*r:id="([^"]+)"', wbx).group(1)
        t = re.search(rf'Id="{rid}"[^>]*Target="([^"]+)"|Target="([^"]+)"[^>]*Id="{rid}"', rels)
        return z.read('xl/' + (t.group(1) or t.group(2)).replace('/xl/', '')).decode()
    for r in [x for x in rows if x['file'] == f]:
        n = r['row']; ws = wb[r['sheet']]
        # F: 기준 시트 값
        for col, k in (('A', 'item'), ('E', 'calc'), ('F', 'final'), ('I', 'unit')):
            v = gs[f'{col}{n}'].value
            if str(v) != str(r[k]):
                bad(f'F {f} 기준!{col}{n}={v!r} ≠ 기준표 {k}={r[k]!r}')
        if not r.get('linked'):
            for col, k in (('G', 'mdl'), ('H', 'loq')):
                v = gs[f'{col}{n}'].value
                if str(v if v is not None else '') != str(r[k] if r[k] is not None else ''):
                    bad(f'F {f} 기준!{col}{n}={v!r} ≠ 기준표 {k}={r[k]!r}')
        # D: 검출한계·정량한계·시험방법 칸
        if not r.get('linked'):
            if str(ws[r['mdl_cell']].value) != f"='기준'!$G${n}":
                bad(f"D {f} {r['sheet']}!{r['mdl_cell']}={ws[r['mdl_cell']].value!r} (기준 G{n}을 가리켜야 함)")
            if r.get('loq_cell') and str(ws[r['loq_cell']].value) != f"='기준'!$H${n}":
                bad(f"D {f} {r['sheet']}!{r['loq_cell']}={ws[r['loq_cell']].value!r} (기준 H{n})")
        if r.get('method_cell') and r['method_row'] == n and str(ws[r['method_cell']].value) != f"='기준'!$D${n}":
            bad(f"D {f} {r['sheet']}!{r['method_cell']}={ws[r['method_cell']].value!r} (기준 D{n})")
        sx = part_of(r['sheet'])
        for res in r['result'].split(', '):
            checked += 1
            fm = str(ws[res].value)
            # A
            rows_ref = set(re.findall(r"'기준'!\$[A-Z]\$(\d+)", fm))
            if rows_ref != {str(n)}:
                bad(f'A {f} {r["sheet"]}!{res} 기준 행 {sorted(rows_ref)} (자기 행 {n}이어야 함)')
            # B
            if re.search(r'ROUND\(\s*\$?[A-Z]{1,3}\$?\d+\s*,\s*\d+\s*\)', fm):
                bad(f'B {f} {r["sheet"]}!{res} 고정 자릿수 ROUND가 남음')
            # C
            mp = re.search(r'TRIM\(SUBSTITUTE\((\$?[A-Z]{1,3}\$?\d+)&""', fm)
            if '"불검출"' not in fm:
                bad(f'C {f} {r["sheet"]}!{res} "불검출" 없음')
            if mp and '" (초과)"' not in fm:
                bad(f'C {f} {r["sheet"]}!{res} 허가배출기준이 있는데 "(초과)" 없음')
            # G
            if not re.search(rf'<conditionalFormatting sqref="{res}"><cfRule type="containsText"[^>]*text="초과"', sx):
                bad(f'G {f} {r["sheet"]}!{res} 초과 조건부 서식 없음')
            # 다른 엔진 시험 자료
            m = re.search(r"ROUND\(ROUND\(IFERROR\(VALUE\(LEFT\((\$?[A-Z]{1,3}\$?\d+)&\"\"", fm)
            m2 = re.search(r"ROUND\(ROUND\(IFERROR\(VALUE\(LEFT\(IF\(((?:'[^']+'|[^!(),\s]+)!)?(\$?[A-Z]{1,3}\$?\d+)=\"\"", fm)
            if not (m or m2):
                continue
            rsh, rc = (r['sheet'], m.group(1).replace('$', '')) if m else ((m2.group(1) or '').rstrip('!').strip("'") or r['sheet'], m2.group(2).replace('$', ''))
            try:
                mdl = float(r['mdl'])
            except (TypeError, ValueError):
                mdl = None
            base = mdl or 1.0; calc, final = int(r['calc']), int(r['final'])
            u = 10 ** -max(final, decs(mdl) if mdl else final)
            for x in [base, base - 0.4 * u, base - 0.6 * u, base + random.randint(1, 30) * u + 0.5 * 10 ** -calc,
                      round(random.uniform(0, base * 40), calc + 2), round(random.uniform(0, base * 400), calc + 3)]:
                x = max(0.0, round(x, calc + 4))
                L = round(random.uniform(base, base * 60), final)
                Lp = random.choice([L, round(x, final), None])
                cases.append(dict(file=f, sheet=r['sheet'], result=res, formula=fm, rsh=rsh, rc=rc, x=x,
                                  pc=mp.group(1).replace('$', '') if mp else None, permit='해당없음' if Lp is None else f'{Lp:g}',
                                  expect=expected(x, calc, final, mdl, Lp if mp else None), item=r['item']))
    # 셀 지도: 수식(골든) + 값(LibreOffice 계산본)
    cells = {}
    for ws in wb.worksheets:
        for row in ws.iter_rows():
            for c in row:
                v = c.value
                if v is None:
                    continue
                cv = wv[ws.title][c.coordinate].value if wv and ws.title in wv.sheetnames else None
                isf = isinstance(v, str) and v.startswith('=')
                # 수식 결과가 비어 있으면 엑셀처럼 빈 글자("") — None으로 넘기면 E23="" 같은 비교가 엑셀과 달라진다
                val = cv if cv is not None else ('' if isf else v)
                if hasattr(val, 'isoformat'):
                    val = 0
                cells[f'{ws.title}!{c.coordinate}'] = dict(f=v if isf else None, v=val)
    json.dump(cells, open(os.path.join(out, f'cells_{os.path.splitext(f)[0]}.json'), 'w', encoding='utf8'), ensure_ascii=False, default=str)
json.dump(cases, open(os.path.join(out, 'cases.json'), 'w', encoding='utf8'), ensure_ascii=False)
print(f'정적 감사: 결과 칸 {checked}개, 문제 {len(problems)}건')
for p in problems[:60]:
    print('   ✖', p)
print(f'다른 엔진 시험 자료: {len(cases)}건')
