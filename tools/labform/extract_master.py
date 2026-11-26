# 시험기록부에서 물질별 기준(결과 칸·방법검출한계·정량한계·시험방법·단위·자릿수)을 찾아 기준표(JSON)로 뽑는다.
#   python3 -I tools/labform/extract_master.py <기록부 폴더> <기준표.json>
# 결과 칸 = "불검출"을 내는 수식이 있는 칸. 수식이 비교하는 칸 = 방법검출한계 칸.
import sys, os, re, glob, json, warnings
import openpyxl
from openpyxl.utils import get_column_letter as CL, column_index_from_string as CI
warnings.filterwarnings('ignore')
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

norm = lambda s: re.sub(r'\s', '', str(s or ''))
REF = r"(?:'[^']+'!|[^\s(),'\"!=<>&*/+-]+!)?\$?[A-Z]{1,3}\$?\d+"


def label_right(ws, words, near=None):
    """라벨(방법검출한계·정량한계 등) 오른쪽 값 칸. near=(r,c)가 있으면 그 근처 것"""
    M = {(g.min_row, g.min_col): g for g in ws.merged_cells.ranges}
    best = None
    for row in ws.iter_rows(max_row=min(ws.max_row, 80), max_col=min(ws.max_column, 14)):
        for c in row:
            if norm(c.value) in words:
                g = M.get((c.row, c.column))
                vc = (g.max_col if g else c.column) + 1
                d = abs(c.row - near[0]) + abs(vc - near[1]) / 3 if near else c.row
                if best is None or d < best[0]:
                    best = (d, c.row, vc)
    return (best[1], best[2]) if best else None


def _start_of(text, pos):
    """pos 바로 앞 식의 시작: 괄호 깊이를 세어 거꾸로, 같은 깊이의 ( , 를 만날 때까지"""
    depth = 0; i = pos - 1
    while i >= 0:
        ch = text[i]
        if ch == ')':
            depth += 1
        elif ch == '(':
            if depth == 0:
                return i + 1
            depth -= 1
        elif ch == ',' and depth == 0:
            return i + 1
        i -= 1
    return 0


def nd_compares(f):
    """불검출 판정 비교 [(값 식, 연산자, 한계 칸)]
       꼴 1: 값<한계,"불검출"     꼴 2: IF(값>한계, …, "불검출")"""
    out = []
    for m in re.finditer(rf'(<=|<)\s*({REF})\s*,\s*"불검출"', f):
        st = _start_of(f, m.start())
        out.append((f[st:m.start()].strip(), m.group(1), m.group(2).replace('$', '')))
    for m in re.finditer(rf'(>=|>)\s*({REF})\s*,', f):
        # 이 IF의 마지막 인자가 "불검출"인지
        st = _start_of(f, m.start())
        if not f[:st].endswith('IF('):
            continue
        depth = 0; j = m.end(); args = 1
        while j < len(f):
            ch = f[j]
            if ch == '(':
                depth += 1
            elif ch == ')':
                if depth == 0:
                    break
                depth -= 1
            elif ch == ',' and depth == 0:
                args += 1; last = j
            j += 1
        if args == 2 and f[last + 1:j].strip() == '"불검출"':
            out.append((f[st:m.start()].strip(), m.group(1), m.group(2).replace('$', '')))
    return out


def lab_basic(ws):
    """시험 기초자료(실험실이 측정한 방법검출한계·정량한계). 비소처럼 줄이 여럿(입자상·가스상)이면 줄마다.
    돌려주는 값: lab_rows=[{label, mdl, loq}], lab_mdl·lab_loq = 첫 줄"""
    for row in ws.iter_rows(max_row=min(ws.max_row, 60), max_col=4):
        for c in row:
            if not norm(c.value).startswith('시험기초자료'):
                continue
            head = c.row; cols = {}
            for rr in range(c.row, c.row + 2):
                for cc in range(3, 15):
                    h = norm(ws.cell(rr, cc).value)
                    if ('검출한계' in h or '정량한계' in h or '정량범위' in h) and '방법바탕' not in h:
                        cols['mdl' if '검출' in h else 'loq'] = cc; head = rr
            out = []
            for rr in range(head + 1, head + 6):
                if rr > head + 1 and re.fullmatch(r'\d{1,2}', str(ws.cell(rr, 1).value or '').strip()):
                    break
                vals = {k: ws.cell(rr, cc).value for k, cc in cols.items()}
                if all(v is None for v in vals.values()):
                    break
                if any(isinstance(v, str) and not v.startswith('=') and not re.fullmatch(r'[-\d.]+', v.strip()) for v in vals.values()):
                    break      # 다른 표(측정결과 등)
                lim = min(cols.values()) if cols else 15
                lab = next((str(ws.cell(rr, cc).value).strip() for cc in range(2, lim) if isinstance(ws.cell(rr, cc).value, str)
                            and not norm(ws.cell(rr, cc).value).startswith('시험기초자료') and not str(ws.cell(rr, cc).value).startswith('=')), '')
                out.append(dict(label=lab, **vals))
            res = dict(lab_rows=out)
            if out:
                res.update(lab_mdl=out[0].get('mdl'), lab_loq=out[0].get('loq'))
            return res
    return {}


def item_of(ws):
    for r in range(1, 12):
        if norm(ws.cell(r, 2).value) == '시험항목':
            return str(ws.cell(r, 4).value or '').strip(), (r, 4)
    return '', None


def digits_of(formula, cell):
    ds = [int(x) for x in re.findall(r'ROUND\([^;]*?,\s*(\d)\)', formula)]
    if ds:
        return max(ds)
    fmt = cell.number_format or ''
    m = re.search(r'0\.(0+)', fmt)
    return len(m.group(1)) if m else None


LAB = [{}]


def scan(path):
    wb = openpyxl.load_workbook(path)
    out = []
    for ws in wb.worksheets:
        if ws.sheet_state != 'visible' or '시험기록부' not in norm(ws['A1'].value):
            continue
        item, item_cell = item_of(ws)
        LAB[0] = lab_basic(ws)
        method = label_right(ws, {'시험방법'})
        for row in ws.iter_rows(max_row=min(ws.max_row, 80)):
            for c in row:
                v = c.value
                if not (isinstance(v, str) and v.startswith('=') and '불검출' in v):
                    continue
                found = nd_compares(v)
                if not found:
                    continue
                lhs, op, mdl = found[0]
                if '!' in mdl or not re.fullmatch(r'\$?[A-Z]{1,3}\$?\d+', mdl):
                    continue
                mr, mc = int(re.sub(r'\D', '', mdl)), CI(re.sub(r'[\d$]', '', mdl))
                lab = ws.cell(mr, mc - 1).value or ws.cell(mr, mc - 2).value
                loq = label_right(ws, {'정량한계'}, near=(mr, mc))
                lim = label_right(ws, {'허가배출기준'}, near=(mr, mc))
                unit = ws.cell(mr, mc + 1).value
                if not isinstance(unit, str) or len(unit) > 12:
                    unit = ws.cell(mr, mc + 2).value if isinstance(ws.cell(mr, mc + 2).value, str) else ''
                sub = item
                # 한 시트에 여러 물질(SOx·NOx·CO): 결과 칸 왼쪽 첫 열의 물질명
                name_here = str(ws.cell(c.row, 1).value or '').split('\n')[0].strip()
                if name_here and not re.fullmatch(r'\d{1,2}', name_here) and not name_here.startswith('='):
                    sub = name_here
                excess = '초과' in v
                out.append(dict(
                    file=os.path.basename(path), sheet=ws.title, item=sub, result=c.coordinate, formula=v,
                    raw=lhs, op=op, compares=[list(x) for x in found], mdl_cell=mdl.replace('$', ''), mdl_label=norm(lab),
                    mdl=ws.cell(mr, mc).value,
                    loq_cell=f'{CL(loq[1])}{loq[0]}' if loq else None, loq=ws.cell(*loq).value if loq else None,
                    limit_cell=f'{CL(lim[1])}{lim[0]}' if lim else None,
                    unit=(unit or '').strip(), method_cell=f'{CL(method[1])}{method[0]}' if method else None,
                    method=str(ws.cell(*method).value or '').strip() if method else '',
                    final_digits=digits_of(v, c), excess=excess, **LAB[0]))
        # 불검출 수식이 없는 양식(먼지 등): "측정결과" 칸에 같은 기준을 건다
        if not any(o['sheet'] == ws.title for o in out):
            res = label_right(ws, {'측정결과'})
            mdl = label_right(ws, {'방법검출한계'}, near=res) if res else None
            if res and mdl:
                c = ws.cell(*res)
                v = c.value
                m = re.fullmatch(rf'=\s*({REF})', str(v or ''))
                if m:
                    R = m.group(1); M = f'{CL(mdl[1])}{mdl[0]}'
                    loq = label_right(ws, {'정량한계'}, near=mdl)
                    unit = ws.cell(res[0], res[1] + 1).value if isinstance(ws.cell(res[0], res[1] + 2).value, str) else ws.cell(mdl[0], mdl[1] + 1).value
                    fmt = re.search(r'0\.(0+)', c.number_format or '')
                    out.append(dict(
                        file=os.path.basename(path), sheet=ws.title, item=item, result=c.coordinate,
                        formula=f'=IF(AND(ISNUMBER({M}),{R}<{M}),"불검출",{R})', raw=R, op='<', compares=[[R, '<', M]],
                        mdl_cell=M, mdl_label='방법검출한계', mdl=ws.cell(*mdl).value,
                        loq_cell=f'{CL(loq[1])}{loq[0]}' if loq else None, loq=ws.cell(*loq).value if loq else None,
                        limit_cell=None, unit=str(ws.cell(mdl[0], mdl[1] + 1).value or '').strip(),
                        method_cell=f'{CL(method[1])}{method[0]}' if method else None,
                        method=str(ws.cell(*method).value or '').strip() if method else '',
                        final_digits=len(fmt.group(1)) if fmt else 0, excess=False, synthetic=True, **LAB[0]))
    return out


if __name__ == '__main__':
    src, dst = sys.argv[1], sys.argv[2]
    rows = []
    for p in sorted(glob.glob(os.path.join(src, '**', '*.xls*'), recursive=True)):
        rows += scan(p)
    json.dump(rows, open(dst, 'w', encoding='utf8'), ensure_ascii=False, indent=1)
    for r in rows:
        print(f"{r['file'][:18]:18} {r['sheet'][:14]:14} {r['item'][:14]:14} 결과 {r['result']:5} 원값 {r['raw'][:28]:28} {r['op']:2} 검출한계 {r['mdl_cell']:4}={str(r['mdl'])[:8]:8}({r['mdl_label'][:6]}) 정량 {r['loq_cell']}={str(r['loq'])[:7]} 단위 {r['unit'][:7]} 자릿수 {r['final_digits']} 초과 {r['excess']}")
    print(len(rows), '항목')
