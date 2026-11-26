# 허가배출기준 초과 표시·조건부 서식·산소 보정 연동
#
# 허가배출기준 칸 P 읽기 (사람이 적는 칸 그대로):  "30", "8(12)", "1.6\n(12)", "16 (12)", "해당없음"
#   기준값 L  = 괄호 앞 숫자   (숫자가 아니면 빈칸 → 초과 판정 안 함)
#   표준산소 O2 = 괄호 안 숫자 (없으면 빈칸 → 산소 보정 안 함). 기준값 바로 오른쪽 칸에 숫자로 따로 적는 양식(중금속 K17)은 그 칸
# 결과 수식 끝에: 표시값(산소 보정을 골랐으면 보정값)이 L보다 크면 " (초과)"  — 같으면 초과 아님
#   원래 수식의 초과 표시("( 초과 )", " 초과", ">=" 비교 등)는 지우고 이 규칙 하나로
# 조건부 서식: 결과 칸에 "초과"가 들어 있으면 빨간 굵은 글씨 + 연한 빨강 바탕
# 산소 보정 연동: "표준산소"·"기준 산소농도" 머리 칸 아래(또는 옆)의 입력 칸 중 수식이 실제로 참조하는 칸을
#   = 허가배출기준의 표준산소 로 바꾼다. 그 칸이 다른 칸을 가리키는 수식이면 끝까지 따라가 원래 입력 칸을 바꾼다.
import re
from xml.sax.saxutils import escape
from openpyxl.utils import get_column_letter as CL, column_index_from_string as CI

PERMIT_WORDS = ('허가배출기준', '배출허용기준', '배출허가기준')   # 양식마다 이름이 다르다 (라벨 글자는 바꾸지 않음)
O2_WORDS = ('표준산소', '표준산소농도', '기준산소농도', '기준산소')
O2_SKIP = ('적용', '최종', '환산', 'Os:', '실측')
norm = lambda s: re.sub(r'\s', '', str(s or ''))


def q(sheet):
    return "'" + sheet.replace("'", "''") + "'"


def permit_exprs(P, O2cell=None):
    """P: 허가배출기준 칸 참조(시트 포함 가능). 기준값·표준산소 식"""
    PS = f'TRIM(SUBSTITUTE({P}&"",CHAR(10)," "))'
    L = f'IFERROR(VALUE(TRIM(LEFT({PS},FIND("(",{PS}&"(")-1))),"")'
    O2 = f'IFERROR(VALUE(MID({PS},FIND("(",{PS})+1,FIND(")",{PS})-FIND("(",{PS})-1)),"")'
    if O2cell:
        O2 = f'IF(ISNUMBER({O2cell}),{O2cell},{O2})'
    return L, O2


def strip_excess(f):
    """원래 수식의 초과 표시 문자열을 지운다 (비교는 남아도 두 갈래가 같아진다)"""
    return re.sub(r'&\s*"\s*\(?\s*초\s*과\s*\)?\s*"', '', f)


def top_if(f):
    """f가 IF(a,b,c) 하나로 되어 있으면 (a,b,c)"""
    f = f.strip()
    if f.startswith('(') and f.endswith(')'):
        f = f[1:-1]
    if not f.startswith('IF('):
        return None
    depth = 0; args = []; cur = ''; inq = False
    for ch in f[3:]:
        if ch == '"':
            inq = not inq
        if not inq:
            if ch == '(':
                depth += 1
            elif ch == ')':
                if depth == 0:
                    args.append(cur); break
                depth -= 1
            elif ch == ',' and depth == 0:
                args.append(cur); cur = ''; continue
        cur += ch
    return args if len(args) == 3 else None


def selected_value(orig, Rs):
    """표시되는 계산값: 계산값이 둘(보정 전/후)이면 원래 수식의 바깥 IF 조건으로 고른다"""
    if len(Rs) == 1:
        return Rs[0]
    a = top_if(orig)
    if a and Rs[0] in a[1] and Rs[1] in a[2]:
        return f'IF({a[0]},{Rs[0]},{Rs[1]})'
    return Rs[0]


# ── 조건부 서식 ──
def red_dxf(b):
    st = b.text('xl/styles.xml')
    dxf = '<dxf><font><b/><color rgb="FFC00000"/></font><fill><patternFill><bgColor rgb="FFFFE5E5"/></patternFill></fill></dxf>'
    m = re.search(r'<dxfs count="(\d+)"\s*/>', st)
    if m:
        st = st.replace(m.group(0), f'<dxfs count="1">{dxf}</dxfs>'); idx = 0
    else:
        m = re.search(r'<dxfs count="(\d+)">(.*?)</dxfs>', st, re.S)
        if m:
            items = re.findall(r'<dxf>.*?</dxf>', m.group(2), re.S)
            if dxf in items:
                return items.index(dxf)
            st = st.replace(m.group(0), f'<dxfs count="{len(items) + 1}">{m.group(2)}{dxf}</dxfs>'); idx = len(items)
        else:
            st = re.sub(r'(</cellStyles>|<tableStyles)', lambda k: (f'{k.group(0)}<dxfs count="1">{dxf}</dxfs>' if k.group(0) == '</cellStyles>' else f'<dxfs count="1">{dxf}</dxfs>{k.group(0)}'), st, count=1); idx = 0
    b.put('xl/styles.xml', st)
    return idx


def add_red_rule(b, sheet, ref_, dxf):
    s = b.sx(sheet)
    if f'sqref="{ref_}"' in s and '초과' in s:
        return
    pri = max([int(x) for x in re.findall(r'priority="(\d+)"', s)] + [0]) + 1
    cf = (f'<conditionalFormatting sqref="{ref_}"><cfRule type="containsText" dxfId="{dxf}" priority="{pri}" operator="containsText" text="초과">'
          f'<formula>NOT(ISERROR(SEARCH("초과",{ref_})))</formula></cfRule></conditionalFormatting>')
    m = re.search(r'<(dataValidations|hyperlinks|printOptions|pageMargins|pageSetup|headerFooter|rowBreaks|drawing|legacyDrawing|tableParts|extLst)\b', s)
    s = s[:m.start()] + cf + s[m.start():] if m else s.replace('</worksheet>', cf + '</worksheet>')
    b.setsx(sheet, s)


# ── 산소 보정 연동 ──
def find_permit(ws, near):
    best = None
    for row in ws.iter_rows(max_row=min(ws.max_row, 80), max_col=min(ws.max_column, 14)):
        for c in row:
            if norm(c.value) in PERMIT_WORDS:
                M = {(g.min_row, g.min_col): g for g in ws.merged_cells.ranges}
                g = M.get((c.row, c.column))
                vc = (g.max_col if g else c.column) + 1
                d = abs(c.row - near[0]) if near else c.row
                if best is None or d < best[0]:
                    best = (d, c.row, vc)
    return (best[1], best[2]) if best else None


def o2_targets(wb, form_sheet, formulas):
    """표준산소 입력 칸 [(시트, 칸)] — 수식이 참조하는 칸, 가리키는 수식은 원래 칸까지 따라간다"""
    refd = set()
    for (sh, ref_), f in formulas.items():
        for m in re.finditer(r"(?:'([^']+)'|([A-Za-z0-9_가-힣()\- ]+))?!?\$?([A-Z]{1,3})\$?(\d+)", f):
            pass
    # 참조되는 칸 모으기 (시트 이름 포함)
    for (sh, ref_), f in formulas.items():
        for m in re.finditer(r"(?:(?:'((?:[^']|'')+)'|([^\s!'(),=<>&*/+\-\"]+))!)?\$?([A-Z]{1,3})\$?(\d+)(?![\d(])", f):
            tsh = (m.group(1) or m.group(2) or sh).replace("''", "'")
            refd.add((tsh, f'{m.group(3)}{m.group(4)}'))
    out = []
    for ws in wb.worksheets:
        if ws.sheet_state != 'visible':
            continue
        for row in ws.iter_rows(max_row=min(ws.max_row, 200), max_col=min(ws.max_column, 30)):
            for c in row:
                t = norm(c.value)
                if not t or not any(t.startswith(w) or t == w for w in O2_WORDS) or any(k in t for k in O2_SKIP) or len(t) > 16:
                    continue
                cands = []
                # 아래로 (표 머리)
                for rr in range(c.row + 1, min(c.row + 40, ws.max_row) + 1):
                    cands.append((rr, c.column))
                # 옆으로 (줄 머리: "표준산소 | % | 값")
                for cc in range(c.column + 1, c.column + 4):
                    cands.append((c.row, cc))
                for (rr, cc) in cands:
                    ref_ = f'{CL(cc)}{rr}'
                    if (ws.title, ref_) not in refd:
                        continue
                    v = ws.cell(rr, cc).value
                    out.append(follow(wb, ws.title, ref_, v))
    seen = []
    for t in out:
        if t and t not in seen:
            seen.append(t)
    return seen


def follow(wb, sheet, ref_, v, depth=0):
    """=X, =IF(X="","",X), =IF(X="X","0",X) 꼴이면 X로 따라간다. 계산식이면 None (바꾸지 않음)"""
    if not (isinstance(v, str) and v.startswith('=')):
        return (sheet, ref_)
    if depth > 4:
        return None
    m = re.fullmatch(r"=\s*(?:IF\(\s*)?(?:(?:'((?:[^']|'')+)'|([^\s!'(),=]+))!)?\$?([A-Z]{1,3})\$?(\d+)(?:\s*=\s*\"[^\"]*\"\s*,\s*\"[^\"]*\"\s*,\s*(?:(?:'(?:[^']|'')+'|[^\s!'(),=]+)!)?\$?[A-Z]{1,3}\$?\d+\s*\))?\s*", v)
    if not m:
        return None
    tsh = (m.group(1) or m.group(2) or sheet).replace("''", "'")
    tref = f'{m.group(3)}{m.group(4)}'
    if tsh not in wb.sheetnames:
        return None
    return follow(wb, tsh, tref, wb[tsh][tref].value, depth + 1)
