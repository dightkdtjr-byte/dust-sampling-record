# 시험기록부에 "기준" 시트를 넣고, 방법검출한계·정량한계·시험방법·자릿수를 그 시트 한 곳에서 관리하게 한다.
#   python3 -I tools/labform/apply_master.py <기록부 폴더> <결과(완성본 원본) 폴더> <기준표.json>
#
# 기준 시트 (숨김, 기준 관리 프로그램이 이 칸들을 고친다)
#   1행 머리. 2행부터 한 줄 = 한 물질(결과 칸 묶음)
#   A 물질  B 시트  C 결과 칸  D 공정시험기준  E 계산 자릿수  F 최종 자릿수  G 방법검출한계  H 정량한계  I 단위  J 비고  K 판정 자릿수(수식)
# 기록부
#   · 방법검출한계·정량한계·시험방법 칸 = '기준'!G·H·D (다른 시트에서 가져오던 칸은 그대로 둔다: 암모니아 BA·전체)
#   · 결과 수식: 계산값 R을
#       비교 자리 → V = ROUND(TRUNC(ROUND(R, 계산+5), 계산 자릿수), 판정 자릿수)  (판정 자릿수 = 최종 자릿수)
#       계산 자릿수까지 계산(그 아래는 버림) → 최종 자릿수로 한 번 4사5입 (총칙 1.6·2.14.2 → KS Q 5002 4.2.2.2, 맺음은 한 번에)
#       ROUND(R, 계산+5): 이진 소수 오차(0.025 → 0.02499999…) 때문에 버림이 한 자리 내려가지 않게
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
HEAD = ['물질', '시트', '결과 칸', '공정시험기준', '계산 자릿수', '최종 자릿수', '방법검출한계', '정량한계', '단위', '비고', '판정 자릿수',
        '시험 기초자료 방법검출한계 유효숫자', '시험 기초자료 정량한계 유효숫자']
COL = {k: CL(i + 1) for i, k in enumerate(['item', 'sheet', 'result', 'method', 'calc', 'final', 'mdl', 'loq', 'unit', 'note', 'judge', 'sig_mdl', 'sig_loq'])}


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
    # 판정 자릿수 = 최종 자릿수: 최종값(성적서에 적는 값)으로 방법검출한계와 견준다 — 표시값과 판정이 늘 일치
    # (브로민화합물처럼 검출한계 0.04를 최종 1자리로 나타낼 수 없으면 0.0 < 0.04 → 불검출, 0.05부터 0.1로 검출)
    return f'={COL["final"]}{i}'


def sig_count(v):
    """원문 제시값의 유효숫자 개수 (0.016 → 2, 0.008 → 1, 1.2 → 2, 10 → 2), 많아야 2. 숫자가 아니면 None"""
    try:
        x = float(v)
    except (TypeError, ValueError):
        return None
    if x == 0:
        return None
    t = repr(abs(x)).lower()
    if 'e' in t:
        t = t.split('e')[0]
    if '.' in t and t.endswith('.0'):
        t = t[:-2]
    d = t.replace('.', '').lstrip('0')
    return min(2, len(d))       # 정수 10 → 2개. 정도보증 4.3.4.4 「2개 이하」로 많아야 2


def sig_defaults(rows):
    """시험 기초자료 유효숫자 기본값 = 원문 제시값(공정시험기준 방법검출한계·정량한계)의 유효숫자. 제시값이 없으면 빈칸(입력값 그대로 표시).
    다른 시트 값을 가져오는 줄(암모니아 BA·전체)은 같은 파일의 원래 줄 값"""
    own = next((r for r in rows if not r.get('linked')), None)
    for r in rows:
        src = r if not r.get('linked') or own is None else own
        r.setdefault('sig_mdl', sig_count(src.get('mdl')) or '')
        r.setdefault('sig_loq', sig_count(src.get('loq')) or '')


def add_sheet(b, rows):
    sig_defaults(rows)
    body = ['<row r="1">' + ''.join(cell_xml(1, CL(i + 1), h) for i, h in enumerate(HEAD)) + '</row>']
    for i, r in enumerate(rows, start=2):
        r['judge'] = judge_formula(i)
        vals = [r['item'], r['sheet'], r['result'], r['method'], r['calc'], r['final'], r['mdl'], r['loq'], r['unit'], r.get('note', ''), r['judge'], r['sig_mdl'], r['sig_loq']]
        body.append(f'<row r="{i}">' + ''.join(cell_xml(i, CL(j + 1), v) for j, v in enumerate(vals)) + '</row>')
    widths = [14, 18, 12, 60, 10, 10, 12, 12, 10, 30, 10, 14, 14]
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
def set_formula(b, sheet, ref_, formula, shrink=True):
    b.setf(sheet, ref_, formula)
    if shrink == 'auto' and sheet in b.wb.sheetnames:
        # 글자 결과 칸(시험방법): 옆 빈칸으로 이어져 보일 수 있으면 셀에 맞춤을 켜지 않는다 (content_rules.fit_cells와 같은 조건)
        ws = b.wb[sheet]; c = ws[ref_]
        merged = any(ref_ in g for g in ws.merged_cells.ranges)
        nxt = ws.cell(c.row, c.column + 1)
        spill = (not merged and c.alignment.horizontal in (None, 'general', 'left') and nxt.value is None
                 and not any(nxt.coordinate in g for g in ws.merged_cells.ranges))
        if spill:
            return
    elif not shrink:
        return
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


def rounded(x, calc, digits):
    """계산 자릿수까지 계산(아래는 버림)한 뒤 digits 자리로 한 번 4사5입"""
    return f'ROUND(TRUNC(ROUND({x},{calc}+5),{calc}),{digits})'


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
        V = rounded(Rn, calc, judge)                                          # 판정용 (검출한계와 비교)
        Vf = rounded(Rn, calc, final)                                         # 표시용
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
        # 불검출 판정은 기준 시트의 방법검출한계로: N()으로 감싸 "-"(판정 안 함)이면 0 → 불검출이 나오지 않게
        # (Excel은 숫자 < 글자를 참으로 보아, 감싸지 않으면 "-"일 때 늘 불검출이 된다)
        G = f'N({ref(row, "mdl")})'
        f = re.sub(rf'({re.escape(V)}\s*(?:<=|>=|<|>))\s*{re.escape(mdl)}(?!\d)', lambda m: m.group(1) + G, f)
    if len(Rs) == 1:
        # 불검출 판정을 가장 바깥에서 먼저: 원본 수식이 판정을 허가배출기준 비교 안쪽에 두면(PAHs) 기준이 비거나 0일 때 판정을 건너뛴다
        V1 = rounded(num_part(Rs[0]), calc, judge)
        f = f'IF({V1}<N({ref(row, "mdl")}),"불검출",{f})'
    blank = 'AND(' + ','.join(f'{R}&""=""' for R in Rs) + ')'          # 계산값이 모두 비어 있으면 빈칸
    if permit:
        # 허가배출기준 초과: 표시값(보정 전/후 중 고른 값)이 기준값보다 크면 " (초과)" — 원래 초과 표시는 지운다
        import o2_excess as OX
        f = OX.strip_excess(f)
        L, _ = OX.permit_exprs(permit)
        Rsel = OX.selected_value(ORIG[0], Rs)
        Vsel = rounded(num_part(Rsel), calc, final)
        # 불검출이면 "(초과)"를 붙이지 않는다 (허가배출기준이 검출한계보다 작게 적힌 경우 등)
        f = f'IF(({f})="불검출","불검출",({f})&IF(AND(ISNUMBER({L}),{Vsel}>{L})," (초과)",""))'
    need = 'OR(' + ','.join(f'{R}&""="실측산소 입력"' for R in Rs) + ')'   # 표준산소가 있는데 실측산소가 비면 (허가배출기준 비교보다 바깥)
    return f'IFERROR(IF({blank},"",IF({need},"실측산소 입력",{f})),"")'


ORIG = ['']


import o2_excess as OX
DXF = [None]
PERMITS = {}


def in_basic(ws, ref_):
    """칸이 '시험 기초자료' 표 안(실험실 측정값)이면 True — 기준 시트로 다시 연결하지 않는다 (SOx의 검출한계 5 ppm 등)"""
    r = int(re.sub(r'\D', '', ref_))
    start = None
    for row in ws.iter_rows(max_row=min(ws.max_row, 60), max_col=4):
        for c in row:
            if OX.norm(c.value).startswith('시험기초자료'):
                start = c.row
    if not start or r <= start:
        return False
    for rr in range(start + 1, r + 1):
        if re.fullmatch(r'\d{1,2}', str(ws.cell(rr, 1).value or '').strip()):
            return False
    return True


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


def mirror_red(b, rows, log):
    """결과 칸을 가져와 보여 주는 칸(요약 시트의 INDIRECT(…'!D17') 등)에도 빨간 "초과" 서식, 원래 있던 "초과" 서식은 같은 빨간 서식으로"""
    if DXF[0] is None:
        DXF[0] = OX.red_dxf(b)
    res = {(r['sheet'], x) for r in rows for x in r['result'].split(', ')}
    cols = {x for _, x in res}
    for ws in b.wb.worksheets:
        if ws.sheet_state != 'visible':
            continue
        for row in ws.iter_rows():
            for c in row:
                v = c.value
                if not (isinstance(v, str) and v.startswith('=')) or (ws.title, c.coordinate) in res:
                    continue
                hit = any(re.search(rf"!\$?{re.escape(re.sub(r'(\d+)', r'$?\1', x))}(?![\d])", v, re.I) or f"!{x.lower()}\")" in v.lower() for x in cols)
                if hit and ('INDIRECT' in v.upper() or any(f"{OX.q(sh)}!{x}" in v or f"{sh}!{x}" in v for sh, x in res)):
                    OX.add_red_rule(b, ws.title, c.coordinate, DXF[0])
                    log.append(f'{ws.title}!{c.coordinate}: 결과를 가져와 보여 주는 칸에도 빨간 초과 서식')
    for ws in b.wb.worksheets:
        part = b.sheets.get(ws.title)
        if not part or part not in b.parts:
            continue
        x = b.text(part)
        y = re.sub(r'(<cfRule\b[^>]*?)dxfId="\d+"([^>]*text="초과")', lambda m: f'{m.group(1)}dxfId="{DXF[0]}"{m.group(2)}', x)
        if y != x:
            b.put(part, y)
            log.append(f'{ws.title}: 원래 있던 "초과" 서식을 같은 빨간 서식으로')


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


# 물질별로 사용자가 정한 시험방법 (ES 번호 바꾸기): 트라이클로로에틸렌은 ES 01511.1d로 통일 (0.03/0.1과 같은 원문)
METHOD_CHOICE = {'트라이클로로에틸렌': ('01514.1d', '01511.1d')}
# 원문에 방법검출한계 제시값이 없어 빼는 물질 (불검출 판정 안 함, 사용자 결정)
NO_MDL = {'황산화물', '질소산화물', '일산화탄소', '총탄화수소'}
TITLES = {}
FIX_LOG = []
SIG_LOG = []


def own_lab_rows(it):
    """시험 기초자료 줄 가운데 이 물질 줄만 (SOx·NOx·CO처럼 한 표에 물질별 줄이 있을 때). 이름이 맞는 줄이 없으면 모두"""
    rows = it.get('lab_rows') or []
    nm = re.sub(r'\s', '', it.get('item', ''))
    mine = [r for r in rows if r.get('label') and (re.sub(r'\s', '', r['label']) in nm or nm in re.sub(r'\s', '', r['label']))]
    return mine or rows


def apply(path, items, dst):
    PERMITS.clear(); DXF[0] = None
    b = Book(path)
    import names as NM
    pairs = NM.pairs_for(path)
    nrep = NM.rename_book(b, pairs)        # 물질 이름을 공정시험기준 표기로 (글자·수식·시트 이름·이름 정의까지)
    for it in items:
        for k in ('item', 'sheet', 'method', 'formula', 'mdl', 'loq'):
            if isinstance(it.get(k), str):
                it[k] = NM.rename_text(it[k], pairs)
        it['compares'] = [[NM.rename_text(x, pairs) if isinstance(x, str) else x for x in c] for c in it['compares']]
    # 원본 계산 수식 오류 고치기 (form_fixes.py, 근거는 거기에) + 결과가 읽는 계산값 칸의 앞선 반올림을 버림으로
    import form_fixes as FF
    FIX_LOG[:] = FF.apply(b, path)
    SIG_LOG[:] = []
    cell_re = re.compile(r"(?:(?:'([^']+)'|([^\s!(),=<>&\"'+*/-]+))!)?\$?([A-Z]{1,3})\$?(\d+)(?![\d(])")
    R_CELLS = []        # 결과가 읽는 계산값 칸 (계산 자릿수가 정해진 뒤 앞선 반올림을 버림으로 바꾼다)
    for it in items:
        for c in it['compares']:
            R = c[0] if isinstance(c[0], str) else ''
            for m in cell_re.finditer(R):
                sh = m.group(1) or m.group(2) or it['sheet']
                if sh in b.sheets:
                    R_CELLS.append((sh, f'{m.group(3)}{m.group(4)}', it['item'], it['sheet']))
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
                         method_row=method_owner[g['sheet']][g['method_cell']], linked=bool(linked),
                         lab_mdl=it.get('lab_mdl'), lab_loq=it.get('lab_loq'), lab_rows=own_lab_rows(it)))
    for r in rows:
        if r['item'] in NO_MDL:
            r['note'] = ' · '.join(filter(None, [r['note'], f'원문에 방법검출한계 제시값 없음, 불검출 판정 안 함 (기존 값 {r["mdl"]})']))
            r['mdl'] = '-'
            if r.get('loq') in ('', None, 0, '0'):
                r['loq'] = '-'
            r['mdl_in_basic'] = r.get('mdl_in_basic', False)
    # 사용자가 정한 시험방법 (트라이클로로에틸렌 → ES 01511.1d)
    for r in rows:
        ch = METHOD_CHOICE.get(r['item'])
        if ch and ch[0] in r['method']:
            r['method'] = re.sub(r'ES\s*' + re.escape(ch[0]) + r'[^)\n]*', f'ES {ch[1]}', r['method'])
            r['note'] = ' · '.join(filter(None, [r['note'], f'시험방법 ES {ch[1]}로 통일 (공정시험기준 칸 값과 같은 원문)']))
    # ES 번호 띄어쓰기 오타: "ES 01409. 2" → "ES 01409.2"
    for r in rows:
        fixed_m = re.sub(r'(ES\s*\d{5}\.)\s+(\w)', r'\1\2', r['method'])
        if fixed_m != r['method']:
            r['note'] = ' · '.join(filter(None, [r['note'], f'시험방법 ES 번호 띄어쓰기 오타 고침']))
            r['method'] = fixed_m
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
    if TITLES:
        import names as NM2
        for r in rows:
            nm = NM2.fix_method(r['method'], TITLES)
            if nm != r['method']:
                r['method'] = nm
    add_sheet(b, rows)
    log = []
    for r in rows:
        sh = r['sheet']
        if not r['linked'] and not in_basic(b.wb[sh], r['mdl_cell']):
            set_formula(b, sh, r['mdl_cell'], ref(r['row'], 'mdl'))
            if r['loq_cell']:
                set_formula(b, sh, r['loq_cell'], ref(r['row'], 'loq'))
        if r['method_cell'] and r['method_row'] == r['row']:
            set_formula(b, sh, r['method_cell'], ref(r['row'], 'method'), shrink='auto')
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
    mirror_red(b, rows, log)
    # 새로 수식을 넣은 칸(표준산소 등)의 정렬을 정렬 기준대로 다시 맞춘다
    from layout_rules import align
    for sh in {r['sheet'] for r in rows}:
        pa = b.wb[sh].print_area
        pa = pa[0] if isinstance(pa, (list, tuple)) else pa
        if pa:
            log += align(b, sh, pa.split('!')[-1].replace('$', ''))
    FIX_LOG.extend(FF.post_apply(b, path))     # 표준산소 연결·산소 보정 규칙 등 마지막 손질 (link_o2 뒤)
    SIG_LOG.extend(FF.sig_figs(b, path, rows, SHEET, COL))   # 시험 기초자료 유효숫자 표시 ('기준' L·M열)
    done = set()
    for sh, ref_, item, sheet in R_CELLS + [(sh, ref_, it_, s_) for (sh, ref_, it_, s_) in FF.o2_cells(b, path, R_CELLS)]:
        calc = next((r['calc'] for r in rows if r['item'] == item and r['sheet'] == sheet), None)
        if (sh, ref_) not in done and FF.strip_pre_rounding(b, sh, ref_, calc):
            FIX_LOG.append(f'{sh}!{ref_}: 결과가 읽는 계산값의 반올림을 계산 자릿수({calc}) 버림으로 (맺음은 결과 칸에서 한 번)')
        done.add((sh, ref_))
    b.finish(dst)
    return rows, log


REFS = {}

if __name__ == '__main__':
    src, dst, master = sys.argv[1], sys.argv[2], sys.argv[3]
    import names as NM
    TITLES.update(NM.official_titles(os.path.join(os.path.dirname(os.path.abspath(__file__)), 'work', 'ref_zip')))
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
        import names as NM
        base = os.path.basename(p)
        for a_, b_ in NM.FILE_MAP.items():
            base = base.replace(f'_{a_}.', f'_{b_}.')
        rows, log = apply(p, items, os.path.join(dst, base))
        print('ok', os.path.basename(p), f'기준 {len(rows)}줄')
        for l in FIX_LOG:
            print('     원본 수식 고침:', l)
        for l in SIG_LOG:
            print('     유효숫자:', l)
        for l in log:
            if '표준산소' in l or '허가배출기준' in l:
                print('    ', l)
        for r in rows:
            allrows.append({k: v for k, v in r.items() if k != 'results'} | {'file': base})
    json.dump(allrows, open(master, 'w', encoding='utf8'), ensure_ascii=False, indent=1, default=str)
    print(len(allrows), '기준 줄')
