# 시험기록부 양식의 서식을 모든 양식에서 같게 맞춘다. 칸 배치(형태)는 양식마다 달라도 서식(양식)은 하나.
#  · 테두리: 선이 있는 자리는 그대로 두고 굵기만 맞춘다
#      바깥 테두리와 쪽이 나뉘는 경계 = 굵은 선(medium), 안쪽 = 가는 선(thin). 이중선·점선 등도 가는 선으로
#      표 아래 내용 없는 줄(지운 문서번호 칸 자리)은 인쇄 범위에서 빼고 선을 지운다
#      빈 칸 사이에 한 칸 길이로만 남은 가로선은 지운다
#  · 제목 글꼴, 서명란 글자, 한 줄 전체를 차지하는 항목 머리의 정렬
import re
from openpyxl.utils import get_column_letter as CL, column_index_from_string as CI, range_boundaries
from tplkit import escape

FRAME, INNER = 'medium', 'thin'
TITLE_FONT = '돋움'
SIGN_TEXT = {'담당': '담 당', '담 당': '담 당', '기술책임자': '기술책임자'}


def raw_borders(sheet_xml, styles_xml):
    """칸마다 실제로 파일에 적힌 테두리 (openpyxl은 병합 칸의 테두리를 바꿔 읽어서 쓰지 않는다).
    돌려주는 값: f(r, c) -> (left, right, top, bottom)"""
    cx = re.search(r'<cellXfs count="\d+">(.*?)</cellXfs>', styles_xml, re.S).group(1)
    xfs = re.findall(r'<xf\b[^>]*?(?:/>|>.*?</xf>)', cx, re.S)
    bd = re.search(r'<borders count="\d+"[^>]*>(.*?)</borders>', styles_xml, re.S).group(1)
    borders = []
    for x in re.findall(r'<border\b[^>]*?(?:/>|>.*?</border>)', bd, re.S):
        sides = []
        for n in ('left', 'right', 'top', 'bottom'):
            m = re.search(rf'<{n}\b([^>]*?)/?>', x)
            st = re.search(r'style="([^"]+)"', m.group(1)) if m else None
            sides.append(st.group(1) if st else None)
        borders.append(tuple(sides))
    def of_style(st):
        m = re.search(r'borderId="(\d+)"', xfs[int(st)])
        return borders[int(m.group(1))] if m else (None,) * 4
    cells, rows, cols = {}, {}, {}
    for rm in re.finditer(r'<row r="(\d+)"([^>]*?)(/>|>(.*?)</row>)', sheet_xml, re.S):
        r = int(rm.group(1))
        if 'customFormat="1"' in rm.group(2):
            st = re.search(r' s="(\d+)"', rm.group(2))
            rows[r] = st.group(1) if st else '0'
        for cm in re.finditer(r'<c r="([A-Z]+)\d+"([^>]*?)(?:/>|>)', rm.group(4) or ''):
            st = re.search(r' s="(\d+)"', cm.group(2))
            cells[(r, CI(cm.group(1)))] = st.group(1) if st else '0'
    cm_ = re.search(r'<cols>(.*?)</cols>', sheet_xml, re.S)
    if cm_:
        for m in re.finditer(r'<col min="(\d+)" max="(\d+)"[^>]*?style="(\d+)"', cm_.group(1)):
            for c in range(int(m.group(1)), int(m.group(2)) + 1):
                cols[c] = m.group(3)
    def f(r, c):
        st = cells.get((r, c)) or rows.get(r) or cols.get(c) or '0'
        return of_style(st)
    return f


def _merges(ws):
    m = {}
    for g in ws.merged_cells.ranges:
        for r in range(g.min_row, g.max_row + 1):
            for c in range(g.min_col, g.max_col + 1):
                m[(r, c)] = (g.min_col, g.min_row, g.max_col, g.max_row)
    return m


def _empty(ws, M, r, c):
    a, top, _, _ = M.get((r, c), (c, r, c, r))
    return ws.cell(top, a).value is None


def trim_area(ws, area, B):
    """표 아래 내용 없이 선만 있는 줄을 인쇄 범위에서 뺀다. 마지막 가로선이 맨 아래가 되게. B: raw_borders"""
    c1, r1, c2, r2 = range_boundaries(area)
    last = max((r for r in range(r1, r2 + 1) if any(ws.cell(r, c).value is not None for c in range(c1, c2 + 1))), default=r1)
    M = _merges(ws)
    def full(r):
        return all(B(r, c)[3] or (r < r2 and B(r + 1, c)[2])
                   or M.get((r, c), (0, 0, 0, 0))[3] > r for c in range(c1, c2 + 1))
    # 내용이 끝난 뒤 처음으로 온 너비 전체의 가로선(병합 칸 안은 선으로 친다)
    ends = [r for r in range(last, r2 + 1) if full(r) and not any(M.get((r, c), (0, 0, 0, 0))[3] > r for c in range(c1, c2 + 1))]
    end = r2
    if ends:
        # 마지막 항목 머리줄 바로 아래 선은 건너뛰고, 그다음 선(내용 칸의 바닥)을 끝으로 본다
        cand = [r for r in ends if r > last] or ends
        end = cand[0] if len(cand) else r2
        if end < r2 and all(ws.cell(r, c).value is None for r in range(end + 1, r2 + 1) for c in range(c1, c2 + 1)):
            return f'{CL(c1)}{r1}:{CL(c2)}{end}', end
    return area, r2


class Styles:
    """styles.xml의 border·font·xf를 복사해 새 스타일 번호를 만든다 (같은 것은 한 번만)."""

    def __init__(self, b):
        self.b = b
        self.memo = {}

    def _st(self):
        return self.b.text('xl/styles.xml')

    def _list(self, tag, inner):
        st = self._st()
        m = re.search(rf'<{tag}s count="(\d+)"[^>]*>(.*?)</{tag}s>', st, re.S)
        items = re.findall(rf'<{inner}\b[^>]*?(?:/>|>.*?</{inner}>)', m.group(2), re.S)
        return st, m, items

    def _add(self, tag, inner, xml):
        st, m, items = self._list(tag, inner)
        if xml in items:
            return items.index(xml)
        new = m.group(0).replace(f'count="{m.group(1)}"', f'count="{len(items) + 1}"', 1).replace(f'</{tag}s>', xml + f'</{tag}s>')
        self.b.put('xl/styles.xml', st.replace(m.group(0), new))
        return len(items)

    def xf(self, s):
        st, m, items = self._list('cellXf', 'xf')
        return items[int(s)]

    def border_xml(self, bid):
        return self._list('border', 'border')[2][int(bid)]

    def with_border(self, s, sides):
        key = ('b', s, sides)
        if key in self.memo:
            return self.memo[key]
        x = self.xf(s)
        bid = re.search(r'borderId="(\d+)"', x)
        old = self.border_xml(bid.group(1) if bid else 0)
        head = re.match(r'<border\b[^>]*?/?>', old).group(0).rstrip('/>').rstrip('/') + '>'
        diag = re.search(r'<diagonal\b.*?(?:/>|</diagonal>)', old, re.S)
        parts = []
        for name, style in zip(('left', 'right', 'top', 'bottom'), sides):
            parts.append(f'<{name} style="{style}"><color indexed="64"/></{name}>' if style else f'<{name}/>')
        xml = head + ''.join(parts) + (diag.group(0) if diag else '<diagonal/>') + '</border>'
        nb = self._add('border', 'border', xml)
        nx = re.sub(r'borderId="\d+"', f'borderId="{nb}"', x, count=1) if bid else x.replace('<xf ', f'<xf borderId="{nb}" ', 1)
        if 'applyBorder' not in nx:
            nx = nx.replace('<xf ', '<xf applyBorder="1" ', 1)
        else:
            nx = re.sub(r'applyBorder="\d"', 'applyBorder="1"', nx)
        r = str(self._add('cellXf', 'xf', nx))
        self.memo[key] = r
        return r

    def with_font_name(self, s, name):
        x = self.xf(s)
        fid = re.search(r'fontId="(\d+)"', x).group(1)
        font = self._list('font', 'font')[2][int(fid)]
        nf = re.sub(r'<name val="[^"]*"/>', f'<name val="{escape(name)}"/>', font)
        nf = re.sub(r'<charset val="\d+"/>', '<charset val="129"/>', nf)
        nid = self._add('font', 'font', nf)
        nx = re.sub(r'fontId="\d+"', f'fontId="{nid}"', x, count=1)
        if 'applyFont' not in nx:
            nx = nx.replace('<xf ', '<xf applyFont="1" ', 1)
        return str(self._add('cellXf', 'xf', nx))

    def with_align(self, s, h, v):
        x = self.xf(s)
        al = f'<alignment horizontal="{h}" vertical="{v}"'
        if '<alignment' in x:
            nx = re.sub(r'<alignment\b([^>]*?)/>', lambda m: al + re.sub(r'\s*(horizontal|vertical)="[^"]*"', '', m.group(1)) + '/>', x, count=1)
        elif x.endswith('/>'):
            nx = re.sub(r'<xf\b([^>]*?)/>', lambda m: f'<xf{m.group(1)}>{al}/></xf>', x, count=1)
        else:
            nx = x.replace('</xf>', f'{al}/></xf>')
        if 'applyAlignment' not in nx:
            nx = nx.replace('<xf ', '<xf applyAlignment="1" ', 1)
        return str(self._add('cellXf', 'xf', nx))

    def with_halign(self, s, h):
        x = self.xf(s)
        if '<alignment' in x:
            nx = re.sub(r'<alignment\b([^>]*?)/>', lambda m: '<alignment' + re.sub(r'\s*horizontal="[^"]*"', '', m.group(1)) + f' horizontal="{h}"/>', x, count=1)
        elif x.endswith('/>'):
            nx = re.sub(r'<xf\b([^>]*?)/>', rf'<xf\1><alignment horizontal="{h}"/></xf>', x, count=1)
        else:
            nx = x.replace('</xf>', f'<alignment horizontal="{h}"/></xf>')
        if 'applyAlignment' not in nx:
            nx = nx.replace('<xf ', '<xf applyAlignment="1" ', 1)
        return str(self._add('cellXf', 'xf', nx))


def _set_styles(b, sheet, new):
    """new: {(r, c): 스타일 번호}. 없는 칸은 만든다."""
    s = b.sx(sheet)
    rows = {}
    for (r, c), st in new.items():
        rows.setdefault(r, {})[c] = st
    def fix_row(m):
        r = int(m.group(1))
        if r not in rows:
            return m.group(0)
        head, body = m.group(2), m.group(4) or ''
        cells = list(re.finditer(r'<c r="([A-Z]+)\d+"([^>]*?)(/>|>.*?</c>)', body, re.S))
        have = {CI(x.group(1)): x for x in cells}
        out = []
        cols = sorted(set(have) | set(rows[r]))
        for c in cols:
            if c in have:
                x = have[c]
                if c in rows[r]:
                    attrs = re.sub(r' s="\d+"', '', x.group(2))
                    out.append(f'<c r="{x.group(1)}{r}" s="{rows[r][c]}"{attrs}{x.group(3)}')
                else:
                    out.append(x.group(0))
            else:
                out.append(f'<c r="{CL(c)}{r}" s="{rows[r][c]}"/>')
        rows.pop(r)
        return f'<row r="{r}"{head}>' + ''.join(out) + '</row>'
    s = re.sub(r'<row r="(\d+)"([^>]*?)(/>|>(.*?)</row>)', lambda m: fix_row(_RowMatch(m)), s, flags=re.S)
    # 줄 자체가 없는 경우
    for r in sorted(rows):
        cells = ''.join(f'<c r="{CL(c)}{r}" s="{st}"/>' for c, st in sorted(rows[r].items()))
        later = [m for m in re.finditer(r'<row r="(\d+)"', s) if int(m.group(1)) > r]
        pos = later[0].start() if later else s.index('</sheetData>')
        s = s[:pos] + f'<row r="{r}">{cells}</row>' + s[pos:]
    b.setsx(sheet, s)


class _RowMatch:
    # <row .../> 와 <row ...>...</row> 를 같은 꼴로 다루려고
    def __init__(self, m):
        self.m = m
    def group(self, i):
        if i == 2:
            return self.m.group(2)
        if i == 4:
            return self.m.group(4) if self.m.group(3) != '/>' else ''
        return self.m.group(i)


def _cell_style(s_xml, ref, row_style, col_styles, c):
    m = re.search(rf'<c r="{ref}"([^>]*?)(?:/>|>)', s_xml)
    if m:
        st = re.search(r' s="(\d+)"', m.group(1))
        return st.group(1) if st else '0'
    if row_style is not None:
        return row_style
    return col_styles.get(c, '0')


def unify(b, sheet, area, breaks, old_r2=None, ws=None):
    """area(인쇄 범위) 안의 테두리 굵기와 제목·머리 서식을 맞춘다. breaks: 쪽 나눔 줄(그 줄 뒤에서 나뉨)."""
    from std import load
    ws = ws or load(b)[sheet]
    B = raw_borders(b.sx(sheet), b.text('xl/styles.xml'))
    c1, r1, c2, r2 = range_boundaries(area)
    M = _merges(ws)
    same = lambda r, c, r_, c_: (r, c) in M and M.get((r_, c_)) == M[(r, c)]
    S = Styles(b)
    notes = []

    # 1) 지금 선이 있는 자리
    H = {}   # H[(r, c)]: r행과 r+1행 사이, c열
    for r in range(r1 - 1, r2 + 1):
        for c in range(c1, c2 + 1):
            if r1 <= r < r2 and same(r, c, r + 1, c):
                continue
            up = B(r, c)[3] if r >= r1 else None
            dn = B(r + 1, c)[2] if r < r2 else None
            H[(r, c)] = bool(up or dn)
    V = {}   # V[(r, c)]: r행, c열과 c+1열 사이
    for r in range(r1, r2 + 1):
        for c in range(c1 - 1, c2 + 1):
            if c1 <= c < c2 and same(r, c, r, c + 1):
                continue
            lf = B(r, c)[1] if c >= c1 else None
            rt = B(r, c + 1)[0] if c < c2 else None
            V[(r, c)] = bool(lf or rt)

    # 2) 빈 칸 사이에 한 칸 길이로만 남은 가로선은 지운다
    for (r, c), on in list(H.items()):
        if on and r1 <= r < r2 and not H.get((r, c - 1)) and not H.get((r, c + 1)) \
                and _empty(ws, M, r, c) and _empty(ws, M, r + 1, c) and c2 > c1:
            H[(r, c)] = False
            notes.append(f'{CL(c)}{r} 아래 혼자 남은 짧은 선 지움')

    # 3) 굵기: 바깥·쪽 경계 = 굵은 선, 안쪽 = 가는 선
    edge_rows = {r1 - 1, r2} | {e for e in breaks if r1 <= e < r2}
    def hs(r, c):
        if r in edge_rows:
            return FRAME
        return INNER if H.get((r, c)) else None
    def vs(r, c):
        if c in (c1 - 1, c2):
            return FRAME
        return INNER if V.get((r, c)) else None

    s_xml = b.sx(sheet)
    row_style = {}
    for m in re.finditer(r'<row r="(\d+)"([^>]*?)/?>', s_xml):
        if 'customFormat="1"' in m.group(2):
            st = re.search(r' s="(\d+)"', m.group(2))
            row_style[int(m.group(1))] = st.group(1) if st else None
    col_styles = {}
    cols = re.search(r'<cols>(.*?)</cols>', s_xml, re.S)
    if cols:
        for m in re.finditer(r'<col min="(\d+)" max="(\d+)"[^>]*?style="(\d+)"', cols.group(1)):
            for c in range(int(m.group(1)), int(m.group(2)) + 1):
                col_styles[c] = m.group(3)

    new = {}
    for r in range(r1, r2 + 1):
        for c in range(c1, c2 + 1):
            a, top, z, bot = M.get((r, c), (c, r, c, r))
            sides = (vs(r, c - 1) if c == a else None, vs(r, c) if c == z else None,
                     hs(r - 1, c) if r == top else None, hs(r, c) if r == bot else None)
            if B(r, c) == sides:
                continue
            st = _cell_style(s_xml, f'{CL(c)}{r}', row_style.get(r), col_styles, c)
            new[(r, c)] = S.with_border(st, sides)
    # 인쇄 범위에서 뺀 아래 줄의 선은 지운다
    for r in range(r2 + 1, (old_r2 or r2) + 1):
        for c in range(c1, c2 + 1):
            if any(B(r, c)):
                st = _cell_style(s_xml, f'{CL(c)}{r}', row_style.get(r), col_styles, c)
                new[(r, c)] = S.with_border(st, (None, None, None, None))
    _set_styles(b, sheet, new)
    notes.append(f'테두리 {len(new)}칸 맞춤 (바깥·쪽 경계 굵은 선, 안쪽 가는 선)')

    # 4) 제목 글꼴, 서명란 글자, 한 줄 전체 항목 머리 정렬
    s_xml = b.sx(sheet)
    fix = {}
    t = ws.cell(r1, c1)
    if isinstance(t.value, str) and '기' in t.value and (t.font.name or '') != TITLE_FONT:
        fix[(r1, c1)] = S.with_font_name(_cell_style(s_xml, f'{CL(c1)}{r1}', None, {}, c1), TITLE_FONT)
        notes.append(f'제목 글꼴 {t.font.name} → {TITLE_FONT}')
    for c in range(c1, c2 + 1):
        v = ws.cell(r1, c).value
        if isinstance(v, str) and v.strip() in SIGN_TEXT and v != SIGN_TEXT[v.strip()]:
            b.setv(sheet, f'{CL(c)}{r1}', SIGN_TEXT[v.strip()])
            notes.append(f'서명란 "{v}" → "{SIGN_TEXT[v.strip()]}"')
    s_xml = b.sx(sheet)
    for r in range(r1 + 2, r2 + 1):
        if not re.fullmatch(r'\d{1,2}', str(ws.cell(r, c1).value or '').strip()):
            continue
        a, top, z, bot = M.get((r, c1 + 1), (c1 + 1, r, c1 + 1, r))
        lab = ws.cell(top, a)
        if z == c2 and top == bot and isinstance(lab.value, str) and lab.alignment.horizontal != 'left':
            fix[(r, a)] = S.with_halign(_cell_style(s_xml, f'{CL(a)}{r}', None, {}, a), 'left')
            notes.append(f'{CL(a)}{r} "{lab.value[:12]}" 한 줄 머리 → 왼쪽 정렬')
    if fix:
        _set_styles(b, sheet, fix)
    return notes
