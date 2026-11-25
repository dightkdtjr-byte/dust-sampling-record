# 시험기록부의 정렬·수식·그림 위치를 모든 양식에서 같은 기준으로 맞춘다.
#
#  정렬 (세로는 모두 가운데)
#   · 번호 칸, 항목명 칸, 회색 머리 칸, 표 안의 값          → 가로 가운데
#   · 기본 정보의 내용 칸(시험항목 ~ 검정곡선 작성), 기호 설명("C : …"), 긴 설명 글 → 왼쪽 (세 글자 이하 단위 칸은 가운데)
#   · 계산식 칸에 직접 쓴 수식(분수 모양 등)은 배치를 그대로 둔다
#   · 한 줄 전체를 차지하는 항목 머리                       → 왼쪽
#   · 도형(그래프 등) 위에 얹은 글자 칸은 도형에 맞춰 놓은 것이라 그대로 둔다
#  수식 (계산식 칸의 수식 글상자)
#   · 글자 크기를 EQ_PT로 (수식 안의 크기 비율은 유지), 위치를 맞추려고 넣은 앞쪽 빈칸은 지운다
#   · 기호 설명 위쪽 줄들의 세로 가운데. 가로는 "* ○○ 농도" 설명 글 바로 뒤, 없으면 칸 왼쪽에서 EQ_INDENT
#   · 글상자가 글자를 자르지 않게 (줄바꿈·잘림 끔). 수식이 수식 줄보다 크면 그 줄 높이를 늘린다
#  그림 (차트)
#   · 차트가 놓인 빈 영역(가로선으로 막힌 빈 줄 묶음, 표 너비 전체)에 상하좌우 PIC_PAD 여백으로 채운다
#   · 너비나 높이가 5px 미만이라 보이지 않는 그림은 지운다
import re
from openpyxl.utils import get_column_letter as CL, range_boundaries

EQ_PT = 20            # 수식 글자 크기(pt)
EQ_INDENT = 16        # 설명 글이 없을 때 칸 왼쪽에서 띄우는 거리(px)
EQ_GAP = 10           # "* ○○ 농도" 설명 글과 수식 사이(px)
PIC_PAD = 8           # 차트 둘레 여백(px)
EQ_PAD = 4            # 수식 위아래 최소 여백(px)
EMU_PX, EMU_PT = 9525, 12700

INFO_LABELS = {'시험항목', '의뢰기관', '측정대상', '시료인수일', '분석일자', '시험방법', '분석기기', '분석기기조작조건', '검정곡선작성'}
LEGEND = re.compile(r'^\s*\S{1,12}\s*:\s*\S')


def _norm(s):
    return re.sub(r'\s', '', str(s or ''))


def _sections(ws, c1, r1, r2):
    """번호가 붙은 항목: [(시작줄, 끝줄, 항목명)]"""
    # 번호 칸에 번호가 있는 줄, 또는 번호 칸에 글자(번호 없는 표 머리, 예: "시료명")가 있는 줄이 항목의 시작
    num = lambda r: re.fullmatch(r'\d{1,2}', str(ws.cell(r, c1).value or '').strip())
    starts = [r for r in range(r1, r2 + 1) if num(r) or (isinstance(ws.cell(r, c1).value, str) and ws.cell(r, c1).value.strip()
                                                           and not ws.cell(r, c1).value.startswith('='))]
    out = []
    for i, r in enumerate(starts):
        end = (starts[i + 1] - 1) if i + 1 < len(starts) else r2
        out.append((r, end, _norm(ws.cell(r, c1 + 1).value) if num(r) else '(표)'))
    return out


def _gray(cell):
    f = cell.fill
    if f.fill_type != 'solid':
        return False
    c = f.fgColor
    if c.type == 'theme':
        return (c.tint or 0) < -0.01 or c.theme not in (0, 1)
    if c.type == 'rgb' and isinstance(c.rgb, str):
        return c.rgb[-6:].upper() not in ('FFFFFF',)
    return False


def _drawing_part(b, sheet):
    rels = b.sheets[sheet].replace('worksheets/', 'worksheets/_rels/') + '.rels'
    if rels not in b.parts:
        return None
    m = re.search(r'<Relationship\b[^>]*Type="[^"]*/drawing"[^>]*>', b.text(rels))
    if not m:
        return None
    t = re.search(r'Target="([^"]+)"', m.group(0)).group(1)
    return 'xl/drawings/' + t.split('/')[-1]


class Grid:
    """열 너비(px)·줄 높이(px)로 칸 위치를 EMU로 바꾼다."""

    def __init__(self, ws):
        from std import px
        self.ws = ws
        dw = ws.sheet_format.defaultColWidth or 8.43
        dh = ws.sheet_format.defaultRowHeight or 15
        self.wpx = {}
        for c in range(1, 80):
            w = None
            for k, d in ws.column_dimensions.items():
                if d.min and d.max and d.min <= c <= d.max and d.width:
                    w = d.width
            self.wpx[c] = px(w if w else dw)
        self.hpx = lambda r: (ws.row_dimensions[r].height if ws.row_dimensions[r].height is not None else dh) * 96 / 72

    def x(self, c):           # c열 왼쪽 끝(px)
        return sum(self.wpx[k] for k in range(1, c))

    def y(self, r):
        return sum(self.hpx(k) for k in range(1, r))

    def at_x(self, xpx):      # px → (0부터 센 열, 열 안 EMU)
        c = 1
        while xpx >= self.wpx[c]:
            xpx -= self.wpx[c]; c += 1
        return c - 1, int(xpx * EMU_PX)

    def at_y(self, ypx):
        r = 1
        while ypx >= self.hpx(r):
            ypx -= self.hpx(r); r += 1
        return r - 1, int(ypx * EMU_PX)


def _anchors(x):
    return list(re.finditer(r'<xdr:(twoCellAnchor|oneCellAnchor|absoluteAnchor)\b.*?</xdr:\1>', x, re.S))


def _pos(body, tag):
    m = re.search(rf'<xdr:{tag}><xdr:col>(\d+)</xdr:col><xdr:colOff>(-?\d+)</xdr:colOff><xdr:row>(\d+)</xdr:row><xdr:rowOff>(-?\d+)</xdr:rowOff></xdr:{tag}>', body)
    return tuple(int(v) for v in m.groups()) if m else None


def _box(G, body):
    """글상자·그림의 위치(px): (x0, y0, x1, y1)"""
    f = _pos(body, 'from')
    x0 = G.x(f[0] + 1) + f[1] / EMU_PX; y0 = G.y(f[2] + 1) + f[3] / EMU_PX
    t = _pos(body, 'to')
    if t:
        return x0, y0, G.x(t[0] + 1) + t[1] / EMU_PX, G.y(t[2] + 1) + t[3] / EMU_PX
    e = re.search(r'<xdr:ext cx="(\d+)" cy="(\d+)"', body) or re.search(r'<a:ext cx="(\d+)" cy="(\d+)"', body)
    return x0, y0, x0 + int(e.group(1)) / EMU_PX, y0 + int(e.group(2)) / EMU_PX


def _set_from(body, G, xpx, ypx):
    c, co = G.at_x(xpx); r, ro = G.at_y(ypx)
    return re.sub(r'<xdr:from>.*?</xdr:from>', f'<xdr:from><xdr:col>{c}</xdr:col><xdr:colOff>{co}</xdr:colOff><xdr:row>{r}</xdr:row><xdr:rowOff>{ro}</xdr:rowOff></xdr:from>', body, count=1, flags=re.S)


def _set_to(body, G, xpx, ypx):
    c, co = G.at_x(xpx); r, ro = G.at_y(ypx)
    return re.sub(r'<xdr:to>.*?</xdr:to>', f'<xdr:to><xdr:col>{c}</xdr:col><xdr:colOff>{co}</xdr:colOff><xdr:row>{r}</xdr:row><xdr:rowOff>{ro}</xdr:rowOff></xdr:to>', body, count=1, flags=re.S)


def shape_rows(b, sheet, ws):
    """수식이 아닌 도형(그래프 등)이 덮는 줄"""
    part = _drawing_part(b, sheet)
    if not part:
        return set()
    G = Grid(ws); rows = set()
    for a in _anchors(b.text(part)):
        body = a.group(0)
        if re.search(r'<xdr:sp\b', body) and 'm:oMath' not in body and 'a14:m' not in body:
            x0, y0, x1, y1 = _box(G, body)
            r0, _ = G.at_y(y0); r1, _ = G.at_y(max(y0, y1 - 1))
            rows |= set(range(r0 + 1, r1 + 2))
    return rows


# ───────── 정렬 ─────────
def align(b, sheet, area, ws=None):
    from std import load
    from form_style import Styles, _set_styles, _cell_style
    ws = ws or load(b)[sheet]
    c1, r1, c2, r2 = range_boundaries(area)
    M = {}
    for g in ws.merged_cells.ranges:
        M[(g.min_row, g.min_col)] = (g.min_col, g.min_row, g.max_col, g.max_row)
    secs = _sections(ws, c1, r1, r2)
    kind = {}
    for a, e, name in secs:
        for r in range(a, e + 1):
            kind[r] = 'info' if name in INFO_LABELS else 'calc' if name == '계산식' else 'table'
    srows = shape_rows(b, sheet, ws)
    skip = {r for a, e, _ in secs if any(a <= k <= e for k in srows) for r in range(a, e + 1)}
    S = Styles(b); s_xml = b.sx(sheet)
    new, notes = {}, []
    for r in range(r1 + 2, r2 + 1):
        for c in range(c1, c2 + 1):
            x = ws.cell(r, c)
            if type(x).__name__ == 'MergedCell' or x.value is None or (isinstance(x.value, str) and not x.value.strip()):
                continue
            a, top, z, bot = M.get((r, c), (c, r, c, r))
            v = str(x.value)
            numbered = re.fullmatch(r'\d{1,2}', str(ws.cell(r, c1).value or '').strip())
            if c == c1:
                h = 'center'
            elif c == c1 + 1 and numbered:
                h = 'left' if (z == c2 and top == bot) else 'center'
            elif r in skip and c > c1 + 1:
                continue
            elif _gray(x):
                h = 'center'
            elif kind.get(r) == 'info':
                first = M.get((r, c1 + 1), (c1 + 1, r, c1 + 1, r))[2] + 1   # 항목명 바로 오른쪽 = 내용 첫 칸
                h = 'center' if c != first and len(v.strip()) <= 3 and not v.startswith('=') else 'left'   # "시" 같은 단위 칸은 가운데
            elif kind.get(r) == 'calc':
                # 계산식 칸: 기호 설명·설명 글만 왼쪽. 칸에 직접 쓴 수식(분수 모양 등)은 배치를 그대로 둔다
                if LEGEND.match(v) or v.strip().startswith('*') or '\n' in v.strip() or len(v) > 25:
                    h = 'left'
                else:
                    continue
            elif not v.startswith('=') and (LEGEND.match(v) or '\n' in v.strip() or len(v) > 25):
                h = 'left'
            else:
                h = 'center'
            cur = (x.alignment.horizontal or 'general', x.alignment.vertical or 'bottom')
            if cur != (h, 'center'):
                st = _cell_style(s_xml, f'{CL(c)}{r}', None, {}, c)
                new[(r, c)] = S.with_align(st, h, 'center')
                notes.append(f'{CL(c)}{r} {cur[0]}/{cur[1]}→{h}/center')
    if new:
        _set_styles(b, sheet, new)
    return [f'정렬 {len(new)}칸 맞춤: ' + ', '.join(notes[:8]) + (' …' if len(notes) > 8 else '')] if new else []


# ───────── 수식·그림 ─────────
def place(b, sheet, area, ws=None):
    from std import load, text_px
    ws = ws or load(b)[sheet]
    part = _drawing_part(b, sheet)
    if not part:
        return []
    c1, r1, c2, r2 = range_boundaries(area)
    G = Grid(ws)
    M = {}
    for g in ws.merged_cells.ranges:
        M[(g.min_row, g.min_col)] = (g.min_col, g.min_row, g.max_col, g.max_row)
    x = b.text(part)
    notes = []
    secs = _sections(ws, c1, r1, r2)
    calc = [(a, e) for a, e, n in secs if n == '계산식']
    from form_style import raw_borders
    B = raw_borders(b.sx(sheet), b.text('xl/styles.xml'))
    out = x
    for am in reversed(_anchors(x)):
        body = am.group(0)
        x0, y0, x1, y1 = _box(G, body)
        is_math = 'm:oMath' in body or 'a14:m' in body
        is_chart = 'c:chart' in body
        if x0 >= Grid(ws).x(c2 + 1) - 1 or y0 >= Grid(ws).y(r2 + 1) - 1:
            continue   # 인쇄 범위 밖의 그림(자료 영역)은 건드리지 않는다
        if (x1 - x0) < 5 or (y1 - y0) < 5:
            out = out[:am.start()] + out[am.end():]
            notes.append(f'보이지 않는 그림 지움 ({"차트" if is_chart else "도형"}, {x1 - x0:.0f}×{y1 - y0:.0f}px)')
            continue
        new = body
        if is_math and calc:
            a, e = calc[0]
            lab = M.get((a, c1 + 1), (c1 + 1, a, c1 + 1, a))
            cl = lab[2] + 1                       # 내용 칸 첫 열
            # 기호 설명이 시작하는 줄(왼쪽 내용 칸) 위까지가 수식 줄
            band_end = e
            for r in range(a, e + 1):
                v = ws.cell(r, cl).value
                if isinstance(v, str) and LEGEND.match(v):
                    band_end = r - 1; break
            band_end = max(band_end, a)
            # 오른쪽에 기호 설명 열이 있으면 그 앞까지
            right = c2
            for c in range(cl + 1, c2 + 1):
                if any(isinstance(ws.cell(r, c).value, str) and LEGEND.match(ws.cell(r, c).value) for r in range(a, e + 1)):
                    right = c - 1; break
            # 글자 크기
            sizes = [int(s) for s in re.findall(r'sz="(\d+)"', new)]
            k = EQ_PT * 100 / max(sizes) if sizes else 1
            if sizes and abs(k - 1) > 0.01:
                new = re.sub(r'sz="(\d+)"', lambda m: f'sz="{round(int(m.group(1)) * k / 100) * 100 or 100}"', new)
                notes.append(f'수식 글자 크기 {max(sizes) / 100:g}→{EQ_PT}pt')
            # 앞쪽 빈칸 지우기
            new2 = re.sub(r'(<m:t>|<a:t>)\s+(</m:t>|</a:t>)', r'\1\2', new)
            if new2 != new:
                notes.append('수식 앞 빈칸 지움'); new = new2
            # 잘리지 않게
            new = re.sub(r'<a:bodyPr\b[^>]*?/?>', lambda m: re.sub(r'\s(vertOverflow|horzOverflow|wrap)="[^"]*"', '', m.group(0)).replace('<a:bodyPr', '<a:bodyPr wrap="none" vertOverflow="overflow" horzOverflow="overflow"'), new, count=1)
            # 수식은 왼쪽 정렬 수식 단락으로 (가운데 정렬 단락·줄 안 수식이 섞여 있으면 같은 자리에 둬도 보이는 자리가 다르다)
            if '<m:oMathPara' in new:
                if '<m:jc ' in new:
                    new = re.sub(r'<m:jc m:val="[^"]*"/>', '<m:jc m:val="left"/>', new)
                elif '<m:oMathParaPr>' in new:
                    new = new.replace('<m:oMathParaPr>', '<m:oMathParaPr><m:jc m:val="left"/>', 1)
                else:
                    new = re.sub(r'(<m:oMathPara\b[^>]*>)', r'\1<m:oMathParaPr><m:jc m:val="left"/></m:oMathParaPr>', new, count=1)
            else:
                new = re.sub(r'(<a14:m>)(<m:oMath\b.*?</m:oMath>)(</a14:m>)',
                             r'\1<m:oMathPara xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math"><m:oMathParaPr><m:jc m:val="left"/></m:oMathParaPr>\2</m:oMathPara>\3',
                             new, count=1, flags=re.S)
            # 높이는 글자 크기 비율대로
            h = (y1 - y0) * k
            # 가로: "* ○○ 농도" 설명 글 뒤, 없으면 들여쓰기에서 시작해 기호 설명 열 앞(없으면 표 오른쪽 끝)까지 꽉 채운다
            left = G.x(cl) + EQ_INDENT
            for r in range(a, band_end + 1):
                v = ws.cell(r, cl).value
                if isinstance(v, str) and v.strip().startswith('*'):
                    f = ws.cell(r, cl).font
                    left = G.x(cl) + 3 + text_px(v.strip(), f.sz or 11, bool(f.b), False) + EQ_GAP
            right_px = G.x(right + 1) - 6
            w = right_px - left
            top_px = G.y(a); bot_px = G.y(band_end + 1)
            if h + 2 * EQ_PAD > bot_px - top_px:
                # 수식이 수식 줄보다 크면 수식 줄의 마지막 줄을 늘린다
                add_pt = (h + 2 * EQ_PAD - (bot_px - top_px)) * 72 / 96
                old = ws.row_dimensions[band_end].height or (ws.sheet_format.defaultRowHeight or 15)
                newh = round(old + add_pt + 0.5)
                sx = b.sx(sheet)
                m = re.search(rf'<row r="{band_end}"([^>]*?)(/?)>', sx)
                at = re.sub(r'\s*(ht|customHeight)="[^"]*"', '', m.group(1))
                b.setsx(sheet, sx[:m.start()] + f'<row r="{band_end}"{at} ht="{newh:g}" customHeight="1"{m.group(2)}>' + sx[m.end():])
                ws.row_dimensions[band_end].height = newh
                notes.append(f'수식이 들어가게 {band_end}행 높이 {old:g}→{newh:g}pt')
                bot_px = G.y(band_end + 1)
            ny = top_px + max(2, (bot_px - top_px - h) / 2)
            new = _set_from(new, G, left, ny)
            new = re.sub(r'<xdr:ext cx="\d+" cy="\d+"/>', f'<xdr:ext cx="{int(w * EMU_PX)}" cy="{int(h * EMU_PX)}"/>', new, count=1)
            new = re.sub(r'(<a:xfrm\b[^>]*>\s*<a:off x=")\d+(" y=")\d+("/>\s*<a:ext cx=")\d+(" cy=")\d+', lambda m: f'{m.group(1)}0{m.group(2)}0{m.group(3)}{int(w * EMU_PX)}{m.group(4)}{int(h * EMU_PX)}', new, count=1)
            notes.append(f'수식 위치: 계산식 {a}~{band_end}행 가운데 높이, 왼쪽 {left - G.x(cl):.0f}px')
        elif is_chart:
            # 차트가 놓인 빈 줄 묶음: 차트 위 끝 줄에서 위아래로, 내용이 없고 너비 전체 가로선으로 막히지 않은 줄
            r0, _ = G.at_y(y0 + 1); r0 += 1
            def empty(r):
                return all(ws.cell(r, c).value is None for c in range(c1, c2 + 1))
            def wall(r):  # r행 아래 너비 전체 가로선
                return all(B(r, c)[3] or B(r + 1, c)[2] for c in range(c1, c2 + 1))
            top = r0
            while top - 1 >= r1 and empty(top - 1) and not wall(top - 1):
                top -= 1
            bot = r0
            while bot + 1 <= r2 and empty(bot + 1) and not wall(bot):
                bot += 1
            if not empty(r0):
                notes.append(f'!! 차트가 내용 있는 {r0}행 위에 있음 — 자리를 그대로 둠')
                continue
            xa, xb = G.x(c1) + PIC_PAD, G.x(c2 + 1) - PIC_PAD
            ya, yb = G.y(top) + PIC_PAD, G.y(bot + 1) - PIC_PAD
            new = _set_from(new, G, xa, ya)
            new = _set_to(new, G, xb - 0.5, yb - 0.5)
            new = re.sub(r'<xdr:twoCellAnchor\b[^>]*>', '<xdr:twoCellAnchor editAs="oneCell">', new, count=1)
            notes.append(f'차트: {top}~{bot}행 빈 영역에 여백 {PIC_PAD}px로 맞춤')
        out = out[:am.start()] + new + out[am.end():]
    b.put(part, out)
    return notes
