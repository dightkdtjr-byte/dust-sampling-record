# 시험기록부에 "기준" 시트를 넣고, 방법검출한계·정량한계·시험방법·자릿수를 그 시트 한 곳에서 관리하게 한다.
#   python3 -I tools/labform/apply_master.py <기록부 폴더> <결과(골든 카피 원본) 폴더> <기준표.json>
#
# 기준 시트 (숨김, 기준 관리 프로그램이 이 칸들을 고친다)
#   1행 머리. 2행부터 한 줄 = 한 물질(결과 칸 묶음)
#   A 물질  B 시트  C 결과 칸  D 공정시험기준  E 계산 자릿수  F 최종 자릿수  G 방법검출한계  H 정량한계  I 단위  J 비고  K 판정 자릿수(수식)
# 기록부
#   · 방법검출한계·정량한계·시험방법 칸 = '기준'!G·H·D (다른 시트에서 가져오던 칸은 그대로 둔다: 암모니아 BA·전체)
#   · 결과 수식: 계산값 R을
#       비교 자리 → V = ROUND(ROUND(R, 계산 자릿수), 판정 자릿수)  (판정 자릿수 = MAX(최종 자릿수, 검출한계 소수 자릿수))
#       표시 자리 → TEXT(V, 최종 자릿수)  (끝자리 0까지 보이게)
#     "R > 한계면 값, 아니면 불검출" 꼴은 ">="로: 최종값이 방법검출한계와 같으면 불검출이 아니라 값
#     계산값이 비어 있으면 빈칸, 계산 중 오류도 빈칸
import sys, os, re, glob, json, warnings
from xml.sax.saxutils import escape
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
warnings.filterwarnings('ignore')
from tplkit import Book
from openpyxl.utils import get_column_letter as CL, column_index_from_string as CI

SHEET = '기준'
HEAD = ['물질', '시트', '결과 칸', '공정시험기준', '계산 자릿수', '최종 자릿수', '방법검출한계', '정량한계', '단위', '비고', '판정 자릿수']
COL = {k: CL(i + 1) for i, k in enumerate(['item', 'sheet', 'result', 'method', 'calc', 'final', 'mdl', 'loq', 'unit', 'note', 'judge'])}


def ref(row, key):
    return f"'{SHEET}'!${COL[key]}${row}"


def sheet_q(name):
    return f"'{name}'" if re.search(r"[^A-Za-z0-9_가-힣]", name) else name


# ── 기준 시트 만들기 ──
def cell_xml(r, c, v, s=None):
    st = f' s="{s}"' if s is not None else ''
    if v is None or v == '':
        return f'<c r="{c}{r}"{st}/>'
    if isinstance(v, str) and v.startswith('='):
        return f'<c r="{c}{r}"{st}><f>{escape(v[1:])}</f></c>'
    if isinstance(v, (int, float)) and not isinstance(v, bool):
        return f'<c r="{c}{r}"{st}><v>{v}</v></c>'
    return f'<c r="{c}{r}"{st} t="inlineStr"><is><t xml:space="preserve">{escape(str(v))}</t></is></c>'


def judge_formula(i):
    # 판정 자릿수 = MAX(최종 자릿수, 방법검출한계의 소수 자릿수)  — 검출한계가 숫자가 아니면 최종 자릿수
    g = f'{COL["mdl"]}{i}'; f = f'{COL["final"]}{i}'
    return (f'=IF(ISNUMBER({g}),MAX({f},IFERROR(LEN(TEXT({g},"0.##########"))-FIND(".",TEXT({g},"0.##########")),0)),{f})')


def add_sheet(b, rows):
    body = ['<row r="1">' + ''.join(cell_xml(1, CL(i + 1), h) for i, h in enumerate(HEAD)) + '</row>']
    for i, r in enumerate(rows, start=2):
        r['judge'] = judge_formula(i)
        vals = [r['item'], r['sheet'], r['result'], r['method'], r['calc'], r['final'], r['mdl'], r['loq'], r['unit'], r.get('note', ''), r['judge']]
        body.append(f'<row r="{i}">' + ''.join(cell_xml(i, CL(j + 1), v) for j, v in enumerate(vals)) + '</row>')
    widths = [14, 18, 12, 60, 10, 10, 12, 12, 10, 30, 10]
    cols = ''.join(f'<col min="{i + 1}" max="{i + 1}" width="{w}" customWidth="1"/>' for i, w in enumerate(widths))
    xml = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
           '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
           'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
           f'<cols>{cols}</cols><sheetData>{"".join(body)}</sheetData></worksheet>')
    n = 1
    while f'xl/worksheets/sheet{n}.xml' in b.parts:
        n += 1
    part = f'xl/worksheets/sheet{n}.xml'
    b.parts[part] = xml.encode('utf8')
    rels = b.text('xl/_rels/workbook.xml.rels')
    ids = [int(x) for x in re.findall(r'Id="rId(\d+)"', rels)]
    rid = f'rId{max(ids) + 1}'
    rels = rels.replace('</Relationships>', f'<Relationship Id="{rid}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet{n}.xml"/></Relationships>')
    b.put('xl/_rels/workbook.xml.rels', rels)
    wbx = b.text('xl/workbook.xml')
    sids = [int(x) for x in re.findall(r'<sheet\b[^>]*sheetId="(\d+)"', wbx)]
    wbx = wbx.replace('</sheets>', f'<sheet name="{SHEET}" sheetId="{max(sids) + 1}" state="hidden" r:id="{rid}"/></sheets>')
    b.put('xl/workbook.xml', wbx)
    ct = b.text('[Content_Types].xml')
    ct = ct.replace('</Types>', f'<Override PartName="/{part}" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/></Types>')
    b.put('[Content_Types].xml', ct)
    b.sheets[SHEET] = part


# ── 수식 바꾸기 ──
def set_formula(b, sheet, ref_, formula):
    b.setf(sheet, ref_, formula)
    # 새 수식 칸도 날짜·숫자 칸처럼 "셀에 맞춤" (줄바꿈 칸은 그대로) — 기록부 시트만
    if sheet in b.wb.sheetnames and '시험기록부' in str(b.wb[sheet]['A1'].value or '').replace(' ', ''):
        return _shrink(b, sheet, ref_)


def _shrink(b, sheet, ref_):
    from form_style import Styles, _set_styles, _cell_style
    s = b.sx(sheet)
    st = _cell_style(s, ref_, None, {}, CI(re.sub(r'\d', '', ref_)))
    S = Styles(b)
    x = S.xf(st)
    if 'wrapText="1"' in x or 'shrinkToFit="1"' in x:
        return
    _set_styles(b, sheet, {(int(re.sub(r'\D', '', ref_)), CI(re.sub(r'\d', '', ref_))): S.with_shrink(st)})


def tok(R):
    """계산값 R을 정확히 찾는 정규식 (J21이 J210·$J21 안에서 잡히지 않게)"""
    if re.fullmatch(r'\$?[A-Z]{1,3}\$?\d+', R):
        return rf'(?<![A-Za-z0-9_$!]){re.escape(R)}(?!\d)'
    return re.escape(R)


def num_part(R):
    return f'IFERROR(VALUE(LEFT({R}&"",FIND("(",{R}&"(")-1)),{R})'


def o2_suffix(R):
    return f'IF(ISNUMBER(FIND("(",{R}&"")),MID({R}&"",FIND("(",{R}&""),20),"")'


def rewrite(f, compares, row, permit=None):
    """f: '=' 뺀 결과 수식. compares: [(R, op, 한계 칸)]"""
    calc, final, judge = ref(row, 'calc'), ref(row, 'final'), ref(row, 'judge')
    Rs = []
    for R, op, mdl in compares:
        if R in Rs:
            continue
        Rs.append(R)
        # 계산값이 "0.1(12)"처럼 숫자 뒤에 표준산소를 붙인 글자이면: 괄호 앞 숫자로 판정·표시하고 "(12)"는 뒤에 그대로
        Rn = num_part(R)
        V = f'ROUND(ROUND({Rn},{calc}),{judge})'                                # 판정용 (검출한계와 비교)
        Vf = f'ROUND(ROUND({Rn},{calc}),{final})'                               # 표시용
        T = f'TEXT({Vf},IF({final}=0,"0","0."&REPT("0",{final})))&{o2_suffix(R)}'
        t = tok(R)
        f = re.sub(rf'ROUND\(\s*{t}\s*,\s*\d+\s*\)', '\x00OUT\x00', f)               # 원래 반올림해서 내보내던 자리 → 표시
        f = re.sub(rf'{t}(?=\s*(<=|>=|<>|<|>|=(?!")))', '\x00V\x00', f)                 # 비교 왼쪽 (R="" 빈칸 확인은 그대로)
        f = re.sub(rf'(?<=<=|>=|<>)\s*{t}|(?<=[<>])\s*{t}', '\x00V\x00', f)            # 비교 오른쪽
        f = re.sub(rf'(?<![A-Za-z0-9_$!]){t}(?![\d(])' if not R.startswith('IF(') else t, '\x00OUT\x00', f)   # 나머지 = 내보내는 자리
        f = f.replace('\x00V\x00', V).replace('\x00OUT\x00', T)
        if op == '>':
            # "V > 한계면 값, 아니면 불검출" → V >= 한계 (같으면 값)
            f = re.sub(rf'{re.escape(V)}\s*>\s*(?!=)(\$?{re.escape(mdl)})(?!\d)', rf'{V}>=\1', f, count=1)
    blank = 'AND(' + ','.join(f'{R}&""=""' for R in Rs) + ')'          # 계산값이 모두 비어 있으면 빈칸
    if permit:
        # 허가배출기준 초과: 표시값(보정 전/후 중 고른 값)이 기준값보다 크면 " (초과)" — 원래 초과 표시는 지운다
        import o2_excess as OX
        f = OX.strip_excess(f)
        L, _ = OX.permit_exprs(permit)
        Rsel = OX.selected_value(ORIG[0], Rs)
        Vsel = f'ROUND(ROUND({num_part(Rsel)},{calc}),{final})'
        f = f'{f}&IF(AND(ISNUMBER({L}),{Vsel}>{L})," (초과)","")'
    return f'IFERROR(IF({blank},"",{f}),"")'


ORIG = ['']


import o2_excess as OX
DXF = [None]
PERMITS = {}


def permit_cell(ws, it):
    """결과 칸의 허가배출기준 칸: 같은 줄 라벨 오른쪽, 또는 표 머리(허가배출기준(ppm))면 그 열의 결과 줄"""
    rr, rc = int(re.sub(r'\D', '', it['result'])), CI(re.sub(r'\d', '', it['result']))
    mr = int(re.sub(r'\D', '', it['mdl_cell']))
    M = {}
    for g in ws.merged_cells.ranges:
        for r_ in range(g.min_row, g.max_row + 1):
            for c_ in range(g.min_col, g.max_col + 1):
                M[(r_, c_)] = (g.min_row, g.min_col, g.max_col)
    best = None
    for row in ws.iter_rows(max_row=min(ws.max_row, 80), max_col=min(ws.max_column, 14)):
        for c in row:
            t = OX.norm(c.value)
            if not any(t.startswith(w) for w in OX.PERMIT_WORDS) or len(t) > 16:
                continue
            if t in OX.PERMIT_WORDS:                       # 줄 라벨: 오른쪽 칸
                g = M.get((c.row, c.column), (c.row, c.column, c.column))
                cand = (c.row, g[2] + 1); d = abs(c.row - mr)
            else:                                           # 표 머리 "허가배출기준(ppm)": 그 열의 결과 줄 (병합이면 맨 위)
                g = M.get((rr, c.column), (rr, c.column, c.column))
                cand = (g[0], g[1]); d = abs(c.row - rr) if c.row < rr else 999
            if best is None or d < best[0]:
                best = (d, cand)
    if not best:
        return None
    return f'{CL(best[1][1])}{best[1][0]}'


def link_o2(b, permits, log):
    """표준산소 입력 칸 = 허가배출기준 괄호 안 숫자 (기준값 오른쪽에 따로 적는 칸이 있으면 그 칸)"""
    if not permits:
        return
    wb = b.wb
    formulas = {(ws.title, c.coordinate): c.value for ws in wb.worksheets for row in ws.iter_rows() for c in row
                if isinstance(c.value, str) and c.value.startswith('=')}
    refd = set()
    for (sh, _), f in formulas.items():
        for m in re.finditer(r"(?:(?:'((?:[^']|'')+)'|([^\s!'(),=<>&*/+\-\"]+))!)?\$?([A-Z]{1,3})\$?(\d+)(?![\d(])", f):
            refd.add(((m.group(1) or m.group(2) or sh).replace("''", "'"), f'{m.group(3)}{m.group(4)}'))
    first_sheet = next(iter(permits))
    first_P = next(iter(permits[first_sheet].values()))
    # 기준값 바로 오른쪽 칸(단위 앞)이 비어 있고 수식이 참조하면 = 표준산소 칸 (중금속 K17, 비소 K18)
    side = {}
    for sh, d in permits.items():
        for res, P in d.items():
            pr, pc = int(re.sub(r'\D', '', P)), CI(re.sub(r'\d', '', P))
            nxt = f'{CL(pc + 1)}{pr}'
            v = wb[sh][nxt].value; unit = wb[sh].cell(pr, pc + 2).value
            if (sh, nxt) in refd and (v is None or isinstance(v, (int, float))) and isinstance(unit, str):
                side[sh] = nxt
                L, O2 = OX.permit_exprs(P)
                set_formula(b, sh, nxt, O2)
                log.append(f'{sh}!{nxt}: 표준산소 = {P}의 괄호 안 숫자')
    def o2_for(sh, P):
        sc = side.get(sh)
        return OX.permit_exprs(P, O2cell=sc)[1] if sc else OX.permit_exprs(P)[1]
    done = set(side.items())
    # 거꾸로 찾기: 계산과정 칸 ← 기록부 표의 어느 줄에서 쓰이는가 (SOx: 계산과정 F6 ← 기록부 E24 → SO2 묶음)
    back = {}
    for sh in permits:
        for row in wb[sh].iter_rows():
            for c in row:
                if isinstance(c.value, str) and c.value.startswith('='):
                    t_ = OX.follow(wb, sh, c.coordinate, c.value)
                    if t_ and t_[0] != sh:
                        back.setdefault(t_, (sh, c.row))
    def permit_by_row(sh, rr):
        keys = sorted(permits[sh], key=lambda k: int(re.sub(r'\D', '', k)))
        k = max([k for k in keys if int(re.sub(r'\D', '', k)) <= rr] or [keys[0]], key=lambda k: int(re.sub(r'\D', '', k)))
        return permits[sh][k]
    # 표준산소 머리 칸 → 입력 칸
    order = sorted(wb.worksheets, key=lambda w: w.title not in permits)     # 기록부 시트 먼저 (줄마다 다른 허가배출기준)
    for ws in order:
        if ws.sheet_state != 'visible':
            continue
        for row in ws.iter_rows(max_row=min(ws.max_row, 200), max_col=min(ws.max_column, 30)):
            for c in row:
                t = OX.norm(c.value)
                if not t or not any(t.startswith(w) for w in OX.O2_WORDS) or any(k in t for k in OX.O2_SKIP) or len(t) > 16:
                    continue
                # 머리 아래: 숫자·빈칸·"-"·단순 참조 수식이 이어지는 동안만 (글자가 나오면 다른 표) / 줄 머리: 오른쪽 3칸
                down = []
                for rr in range(c.row + 1, min(c.row + 40, ws.max_row) + 1):
                    v_ = ws.cell(rr, c.column).value
                    if isinstance(v_, str) and not v_.startswith('=') and v_.strip() not in ('-', 'X', 'x'):
                        break
                    down.append((rr, c.column))
                side_c = []
                for cc in range(c.column + 1, c.column + 4):
                    v_ = ws.cell(c.row, cc).value
                    if isinstance(v_, str) and not v_.startswith('=') and v_.strip() not in ('%', '-'):
                        break
                    side_c.append((c.row, cc))
                cands = down + side_c
                for rr, cc in cands:
                    ref_ = f'{CL(cc)}{rr}'
                    if (ws.title, ref_) not in refd:
                        continue
                    tgt = OX.follow(wb, ws.title, ref_, ws.cell(rr, cc).value)
                    if not tgt or tgt in done:
                        continue
                    tv = wb[tgt[0]][tgt[1]].value
                    if isinstance(tv, str) and tv.startswith('=') and tgt != (ws.title, ref_):
                        continue
                    if isinstance(tv, str) and tv.startswith('='):
                        continue                                       # 계산식 칸은 바꾸지 않는다
                    # 허가배출기준: 기록부 표의 한 줄에서 쓰이는 칸이면 그 줄이 속한 물질의 것, 아니면 첫 기록부의 것
                    if ws.title in permits:
                        sh_p, P = ws.title, permit_by_row(ws.title, rr)
                    elif tgt in back:
                        sh_p, P = back[tgt][0], permit_by_row(*back[tgt])
                    else:
                        P, sh_p = first_P, first_sheet
                    expr = o2_for(sh_p, f'{OX.q(sh_p)}!{P}' if tgt[0] != sh_p else P)
                    set_formula(b, tgt[0], tgt[1], expr)
                    done.add(tgt)
                    log.append(f'{tgt[0]}!{tgt[1]}: 표준산소 = {sh_p}!{P}의 괄호 안 숫자 (원래 {tv!r})')


def apply(path, items, dst):
    PERMITS.clear(); DXF[0] = None
    b = Book(path)
    # 한 줄 = (시트, 방법검출한계 칸) 묶음
    groups = []
    for it in items:
        g = next((g for g in groups if g['sheet'] == it['sheet'] and g['mdl_cell'] == it['mdl_cell']), None)
        if g:
            g['results'].append(it)
        else:
            groups.append(dict(sheet=it['sheet'], mdl_cell=it['mdl_cell'], loq_cell=it['loq_cell'] if it['loq_cell'] != it['mdl_cell'] else None,
                               method_cell=it['method_cell'], item=it['item'], results=[it], m=it))
    rows = []
    method_owner = {}
    for i, g in enumerate(groups, start=2):
        it = g['m']
        note = []
        mdl, loq = it['mdl'], it['loq']
        linked = None
        if isinstance(mdl, str) and mdl.startswith('='):
            src = re.match(r"='?([^'!]+)'?!\$?([A-Z]+)\$?(\d+)", mdl)
            linked = next((r for r in rows if src and r['sheet'] == src.group(1) and r['mdl_cell'] == f'{src.group(2)}{src.group(3)}'), None)
        if linked:
            mdl = f"='{SHEET}'!{COL['mdl']}{linked['row']}"
            loq = f"='{SHEET}'!{COL['loq']}{linked['row']}"
            note.append(f'방법검출한계·정량한계는 「{linked["sheet"]}」와 같음')
        if it['loq_cell'] == it['mdl_cell']:
            note.append('이 양식은 정량한계 칸으로 불검출을 판정함')
        final = it['final_digits'] if it['final_digits'] is not None else 2
        mkey = (g['sheet'], g['method_cell'])
        if g['method_cell'] in method_owner.get(g['sheet'], {}):
            note.append(f'시험방법 칸을 {method_owner[g["sheet"]][g["method_cell"]]}행과 같이 씀')
        else:
            method_owner.setdefault(g['sheet'], {})[g['method_cell']] = i
        rows.append(dict(row=i, item=it['item'], sheet=g['sheet'], result=', '.join(r['result'] for r in g['results']),
                         method=it['method'], calc=final + 1, final=final, mdl=mdl, loq=loq if g['loq_cell'] else '',
                         unit=it['unit'] or ('ppm' if 'SOx' in it['sheet'] else ''), note=' · '.join(note),
                         mdl_cell=g['mdl_cell'], loq_cell=g['loq_cell'], method_cell=g['method_cell'], results=g['results'],
                         method_row=method_owner[g['sheet']][g['method_cell']], linked=bool(linked)))
    # 원문 대조: 원문에 결과 자릿수·단위가 있으면 원문 값으로, 시험방법 버전이 원문과 다르면 원문 버전으로
    if REFS:
        import check_master as CM
        for r in rows:
            code, rr, vnote = CM.ref_for(r, REFS)
            if not rr:
                continue
            fixed = []
            if rr['digits'] and rr['digits'][0]['calc'] and rr['digits'][0]['final'] is not None:
                c, f = rr['digits'][0]['calc'], rr['digits'][0]['final']
                if (r['calc'], r['final']) != (c, f):
                    fixed.append(f'자릿수 계산 {r["calc"]}·최종 {r["final"]} → {c}·{f} (ES {code})'); r['calc'], r['final'] = c, f
            ru = rr['digits'][0]['unit'] if rr['digits'] else ''
            nu = lambda u: str(u).lower().replace('³', '3').replace(' ', '')
            if ru and r['unit'] and nu(ru) != nu(r['unit']):
                fixed.append(f'단위 {r["unit"]} → {ru} (ES {code}, 기록부 안의 같은 오타도)')
                old_u = r['unit']; r['unit'] = ru
                # 기록부 글자 칸의 같은 단위 오타도 고친다 (공유 문자열·칸 안 글자)
                for part in [p_ for p_ in b.parts if p_ == 'xl/sharedStrings.xml' or p_.startswith('xl/worksheets/sheet')]:
                    x = b.text(part)
                    y = re.sub(rf'(?<![A-Za-z]){re.escape(old_u)}(?![A-Za-z])', ru, x)
                    if y != x:
                        b.put(part, y)
            if vnote:
                old = re.search(r'ES\s*' + re.escape(code.split('.')[0]) + r'\.\s?\w+', r['method'])
                if old:
                    fixed.append(f'시험방법 {old.group(0)} → ES {code} (받은 원문 버전)')
                    r['method'] = r['method'].replace(old.group(0), f'ES {code}')
            if fixed:
                r['note'] = ' · '.join(filter(None, [r['note'], '원문 대조로 고침: ' + '; '.join(fixed)]))
    add_sheet(b, rows)
    log = []
    for r in rows:
        sh = r['sheet']
        if not r['linked']:
            set_formula(b, sh, r['mdl_cell'], ref(r['row'], 'mdl'))
            if r['loq_cell']:
                set_formula(b, sh, r['loq_cell'], ref(r['row'], 'loq'))
        if r['method_cell'] and r['method_row'] == r['row']:
            set_formula(b, sh, r['method_cell'], ref(r['row'], 'method'))
        for it in r['results']:
            ORIG[0] = it['formula'][1:]
            P = permit_cell(b.wb[sh], it)
            f = rewrite(it['formula'][1:], it['compares'], r['row'], permit=P)
            set_formula(b, sh, it['result'], f)
            if P:
                OX.add_red_rule(b, sh, it['result'], DXF[0] if DXF[0] is not None else DXF.__setitem__(0, OX.red_dxf(b)) or DXF[0])
                PERMITS.setdefault(sh, {})[it['result']] = P
            log.append(f"{sh}!{it['result']}: {r['item']} (기준 {r['row']}행){' 허가배출기준 ' + P if P else ''}")
    link_o2(b, PERMITS, log)
    b.finish(dst)
    return rows, log


REFS = {}

if __name__ == '__main__':
    src, dst, master = sys.argv[1], sys.argv[2], sys.argv[3]
    rp = os.path.join(os.path.dirname(os.path.abspath(master)), 'ref_rules.json')
    if os.path.exists(rp):
        REFS.update(json.load(open(rp, encoding='utf8')))
    os.makedirs(dst, exist_ok=True)
    import extract_master as EM
    allrows = []
    for p in sorted(glob.glob(os.path.join(src, '*.xls*'))):
        items = EM.scan(p)
        if not items:
            print('--', os.path.basename(p), '불검출 수식 없음 (기준 시트만 없이 그대로 복사)')
            import shutil; shutil.copy(p, dst)
            continue
        rows, log = apply(p, items, os.path.join(dst, os.path.basename(p)))
        print('ok', os.path.basename(p), f'기준 {len(rows)}줄')
        for l in log:
            if '표준산소' in l or '허가배출기준' in l:
                print('    ', l)
        for r in rows:
            allrows.append({k: v for k, v in r.items() if k != 'results'} | {'file': os.path.basename(p)})
    json.dump(allrows, open(master, 'w', encoding='utf8'), ensure_ascii=False, indent=1, default=str)
    print(len(allrows), '기준 줄')
