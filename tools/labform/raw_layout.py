# raw data·계산과정 시트 인쇄 모양 (기록부 시트가 아닌 시트)
#  · 그림만 있는 시트(기기 출력물 사진): 그림을 원래 순서(왼쪽 열 위→아래, 다음 열)대로 한 줄로 세로로 다시 놓고,
#    한 쪽 너비·높이에 들어가게 줄이며, 그림이 쪽 사이에서 잘리지 않게 쪽 나눔. 100% 배율 (작게 몰려 읽을 수 없던 것)
#  · 칸 내용이 있는 시트: 내용 폭이 세로 쪽보다 넓으면 가로, 좁은 여백, 너비 한 쪽 맞춤 (높이는 여러 쪽)
import math
import re

EMU = 12700                       # 1 pt
PAGE_W, PAGE_H = 595, 842         # A4 (pt)
MARG = dict(left=0.3, right=0.3, top=0.4, bottom=0.4, header=0.2, footer=0.2)
ROW_PT = 15.0                     # 그림 시트의 줄 높이 (고정)
GAP_ROWS = 1                      # 그림 사이 빈 줄


def _drawing_part(b, sheet):
    s = b.sx(sheet)
    d = re.search(r'<drawing r:id="([^"]+)"', s)
    if not d:
        return None
    rels = b.text(b.sheets[sheet].replace('worksheets/', 'worksheets/_rels/') + '.rels')
    t = re.search(rf'Id="{d.group(1)}"[^>]*Target="([^"]+)"', rels) or re.search(rf'Target="([^"]+)"[^>]*Id="{d.group(1)}"', rels)
    return 'xl/drawings/' + t.group(1).split('/')[-1]


def _margins_xml():
    return '<pageMargins ' + ' '.join(f'{k}="{v}"' for k, v in MARG.items()) + '/>'


def _set_page(s, landscape, fit, breaks=(), wide=1):
    s = re.sub(r'<pageMargins[^>]*/>', _margins_xml(), s) if '<pageMargins' in s else s
    ps = (f'<pageSetup paperSize="9" fitToWidth="{wide}" fitToHeight="0" orientation="{"landscape" if landscape else "portrait"}"/>' if fit
          else f'<pageSetup paperSize="9" scale="100" orientation="{"landscape" if landscape else "portrait"}"/>')
    s = re.sub(r'<pageSetup[^>]*/>', ps, s) if '<pageSetup' in s else s.replace(_margins_xml(), _margins_xml() + ps)
    fp = '<pageSetUpPr fitToPage="1"/>' if fit else '<pageSetUpPr fitToPage="0"/>'
    if '<pageSetUpPr' in s:
        s = re.sub(r'<pageSetUpPr[^>]*/>', fp, s)
    elif re.search(r'<sheetPr[^>]*/>', s):
        s = re.sub(r'<sheetPr([^>]*)/>', rf'<sheetPr\1>{fp}</sheetPr>', s, count=1)
    elif '<sheetPr' in s:
        s = s.replace('</sheetPr>', fp + '</sheetPr>', 1)
    else:
        s = re.sub(r'(<worksheet[^>]*>)', rf'\1<sheetPr>{fp}</sheetPr>', s, count=1)
    s = re.sub(r'<rowBreaks.*?</rowBreaks>', '', s, flags=re.S)
    s = re.sub(r'<colBreaks.*?</colBreaks>', '', s, flags=re.S)
    if breaks:
        rb = f'<rowBreaks count="{len(breaks)}" manualBreakCount="{len(breaks)}">' + ''.join(f'<brk id="{r}" max="16383" man="1"/>' for r in breaks) + '</rowBreaks>'
        m = re.search(r'<(colBreaks|customProperties|cellWatches|ignoredErrors|smartTags|drawing|legacyDrawing|picture|oleObjects|controls|webPublishItems|tableParts|extLst)\b', s)
        s = s[:m.start()] + rb + s[m.start():] if m else s.replace('</worksheet>', rb + '</worksheet>')
    return s


PLACEHOLDER = re.compile(r'^\(?기기\s*RAW\s*DATA\s*붙이기\)?$', re.I)


def _set_row(xml, r, h):
    m = re.search(rf'<row r="{r}"([^>]*?)(/?)>', xml)
    if m:
        attrs = re.sub(r'\s(ht|customHeight)="[^"]*"', '', m.group(1))
        return xml[:m.start()] + f'<row r="{r}"{attrs} ht="{h}" customHeight="1"{m.group(2)}>' + xml[m.end():]
    rows_ = [(int(mm.group(1)), mm.start()) for mm in re.finditer(r'<row r="(\d+)"', xml)]
    pos = next((p for rr, p in rows_ if rr > r), None)
    new = f'<row r="{r}" ht="{h}" customHeight="1"/>'
    if pos is None:
        return xml.replace('</sheetData>', new + '</sheetData>').replace('<sheetData/>', f'<sheetData>{new}</sheetData>')
    return xml[:pos] + new + xml[pos:]


def picture_sheet(b, sheet, first_row=0):
    """그림만 있는 시트를 한 줄 세로 배치로. 바꿨으면 설명 글, 아니면 None"""
    dp = _drawing_part(b, sheet)
    if not dp or dp not in b.parts:
        return None
    dx = b.text(dp)
    anchors = list(re.finditer(r'<xdr:(twoCellAnchor|oneCellAnchor|absoluteAnchor)\b[^>]*>(.*?)</xdr:\1>', dx, re.S))
    pics = [a for a in anchors if '<xdr:pic>' in a.group(2)]
    if not pics or len(pics) != len(anchors):
        return None
    items = []
    for a in pics:
        body = a.group(2)
        fr = re.search(r'<xdr:from><xdr:col>(\d+)</xdr:col>.*?<xdr:row>(\d+)</xdr:row>', body, re.S)
        ext = re.search(r'<a:ext cx="(\d+)" cy="(\d+)"', body)
        pic = re.search(r'<xdr:pic>.*?</xdr:pic>', body, re.S).group(0)
        items.append((int(fr.group(1)), int(fr.group(2)), int(ext.group(1)) / EMU, int(ext.group(2)) / EMU, pic))
    items.sort(key=lambda t: (t[0], t[1]))
    usable_w = PAGE_W - (MARG['left'] + MARG['right']) * 72 - 4
    usable_h = PAGE_H - (MARG['top'] + MARG['bottom']) * 72 - 8
    rows_per_page = int(usable_h // ROW_PT)
    row = first_row; page_start = 0; breaks = []; out = []
    for _, _, w, h, pic in items:
        k = min(1.0, usable_w / w, (rows_per_page - 1) * ROW_PT / h)
        w2, h2 = w * k, h * k
        need = int(-(-h2 // ROW_PT))
        if row > page_start and row + need > page_start + rows_per_page:      # 이 쪽에 안 들어가면 다음 쪽 처음으로
            row = page_start + rows_per_page; breaks.append(row); page_start = row
        cx, cy = int(w2 * EMU), int(h2 * EMU)
        pic2 = re.sub(r'<a:ext cx="\d+" cy="\d+"/>', f'<a:ext cx="{cx}" cy="{cy}"/>', pic, count=1)
        out.append(f'<xdr:oneCellAnchor><xdr:from><xdr:col>0</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>{row}</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from>'
                   f'<xdr:ext cx="{cx}" cy="{cy}"/>{pic2}<xdr:clientData/></xdr:oneCellAnchor>')
        row += need + GAP_ROWS
    last = row
    head = dx[:anchors[0].start()]; tail = dx[anchors[-1].end():]
    b.put(dp, head + ''.join(out) + tail)
    # 줄 높이 고정 (칸 내용이 없는 시트), 인쇄 영역 = 그림이 있는 범위, 100%, 쪽 나눔
    s = b.sx(sheet)
    s = re.sub(r'<sheetFormatPr[^>]*/>', f'<sheetFormatPr defaultRowHeight="{ROW_PT}" customHeight="1"/>', s) if '<sheetFormatPr' in s \
        else s.replace('<sheetData', f'<sheetFormatPr defaultRowHeight="{ROW_PT}" customHeight="1"/><sheetData', 1)
    s = re.sub(r'<row r="(\d+)"([^>]*?)\sht="[^"]*"', r'<row r="\1"\2', s)
    s = re.sub(r'(<row r="\d+"[^>]*?)\scustomHeight="1"', r'\1', s)
    for r_ in range(1, last + 2):          # 줄 높이를 줄마다 박아 둠 (기본 줄 높이는 프로그램·글꼴마다 달리 잡힘)
        s = _set_row(s, r_, ROW_PT)
    # 열 12개의 너비 합 = 쪽 너비 (엑셀 기준). 글꼴이 넓은 프로그램에서는 너비 맞춤으로 조금 줄어 그림이 오른쪽으로 삐져나가지 않음
    cols = 12
    cw = round((usable_w + 4) / 0.75 / 8 / cols, 3)
    newcols = f'<cols><col min="1" max="{cols}" width="{cw}" customWidth="1"/></cols>'
    s = re.sub(r'<cols>.*?</cols>', newcols, s, flags=re.S) if '<cols>' in s else s.replace('<sheetData', newcols + '<sheetData', 1)
    s = _set_page(s, landscape=False, fit=True, breaks=breaks)
    b.setsx(sheet, s)
    b.print_area(sheet, f'A1:{_col(cols)}{last}')
    return f'{sheet}: 기기 출력물 그림 {len(items)}장을 한 줄로 세로 배치, 한 쪽에 맞춰 {len(breaks) + 1}쪽 (그림이 쪽 사이에서 잘리지 않게)'


def _col(n):
    s = ''
    while n:
        n, r = divmod(n - 1, 26); s = chr(65 + r) + s
    return s


def content_sheet(b, sheet, ws):
    """칸 내용이 있는 시트: 폭에 맞춰 방향·여백"""
    from openpyxl.utils import get_column_letter as CL
    cells = [(c.row, c.column) for row in ws.iter_rows() for c in row if c.value not in (None, '')]
    if not cells:
        return None
    pa = ws.print_area
    pa = pa[0] if isinstance(pa, (list, tuple)) else pa
    if pa:
        from openpyxl.utils import range_boundaries
        c1, r1, c2, r2 = range_boundaries(pa.split('!')[-1].replace('$', ''))
    else:
        c1 = min(c for _, c in cells); c2 = max(c for _, c in cells)
    wpt = sum((ws.column_dimensions[CL(c)].width or 8.43) * 5.6 for c in range(c1, c2 + 1))
    port_w = PAGE_W - (MARG['left'] + MARG['right']) * 72
    land = wpt > port_w * 1.08
    usable = (PAGE_H if land else PAGE_W) - (MARG['left'] + MARG['right']) * 72
    pct = min(100, round(usable / wpt * 100))
    wide = 2 if pct < 60 else 1          # 너무 작아지면(60 % 미만) 가로 두 쪽으로 나눠 글자 크기를 살림
    s = _set_page(b.sx(sheet), landscape=land, fit=True, wide=wide)
    b.setsx(sheet, s)
    return f'{sheet}: {"가로" if land else "세로"}, 너비 {wide}쪽 (배율 약 {min(100, pct * wide)}%)'


def is_raw_paste(ws):
    """수식이 하나도 없는 시트 = 기기 출력물을 그대로 붙여 넣는 곳 (계산과정에는 반드시 수식이 있다). 이런 시트는 손대지 않는다"""
    return not any(isinstance(c.value, str) and c.value.startswith('=') for row in ws.iter_rows() for c in row)


def _px(w):
    """엑셀 열 너비(글자) → 화면 점 (맑은 고딕 11: 한 글자 8점, 저장 너비에 여백 포함)"""
    return int((w if w else 9) * 8)


def col_widths(x):
    """시트 XML의 <col> (범위 묶음 포함) → {열: (너비, 숨김)}; 없는 열은 기본 너비"""
    d = re.search(r'<sheetFormatPr[^>]*defaultColWidth="([\d.]+)"', x)
    base = re.search(r'<sheetFormatPr[^>]*baseColWidth="([\d.]+)"', x)
    dflt = float(d.group(1)) if d else (float(base.group(1)) + 0.625 if base else 9.0)
    out = {}
    for m in re.finditer(r'<col\b([^>]*)/>', x):
        a = m.group(1)
        lo, hi = int(re.search(r'min="(\d+)"', a).group(1)), int(re.search(r'max="(\d+)"', a).group(1))
        w = re.search(r'width="([\d.]+)"', a)
        for c in range(lo, min(hi, 16384) + 1):
            out[c] = (float(w.group(1)) if w else dflt, 'hidden="1"' in a)
    return out, dflt


def _auto_h(ws, r, c1, c2):
    """높이를 정하지 않은 줄: 글자 줄 수로 (줄바꿈 칸은 열 너비로 나눠 셈)"""
    from openpyxl.utils import get_column_letter as CL
    best = 15.0
    merged = {}
    for g in ws.merged_cells.ranges:
        if g.min_row == r == g.max_row:
            merged[g.min_col] = g.max_col
    for c in range(c1, c2 + 1):
        cell = ws.cell(r, c); v = cell.value
        if v in (None, '') or type(cell).__name__ == 'MergedCell':
            continue
        size = (cell.font.sz or 11) if cell.font else 11
        txt = str(v) if not str(v).startswith('=') else ''
        lines = txt.count('\n') + 1
        if cell.alignment is not None and cell.alignment.wrap_text and txt:
            wpx = sum(_px(ws.column_dimensions[CL(k)].width) for k in range(c, merged.get(c, c) + 1)) - 6
            per = max(1, int(wpx / (size * 1.33 * 0.9)))       # 한글 한 자 ≈ 글자 크기
            lines = sum(max(1, -(-len(part) // per)) for part in txt.split('\n'))
        best = max(best, lines * size * 1.32 + 4)
    return round(best, 2)


def _freeze_anchors(dx, colpx, rowpt):
    """두 칸에 걸친 그림·차트(twoCellAnchor)를 한 칸 기준 + 고정 크기(oneCellAnchor)로: 지금 열 너비·줄 높이로 크기를 재서 박는다"""
    def size(m):
        body = m.group(2)
        fr = re.search(r'<xdr:from><xdr:col>(\d+)</xdr:col><xdr:colOff>(-?\d+)</xdr:colOff><xdr:row>(\d+)</xdr:row><xdr:rowOff>(-?\d+)</xdr:rowOff></xdr:from>', body)
        to = re.search(r'<xdr:to><xdr:col>(\d+)</xdr:col><xdr:colOff>(-?\d+)</xdr:colOff><xdr:row>(\d+)</xdr:row><xdr:rowOff>(-?\d+)</xdr:rowOff></xdr:to>', body)
        if not (fr and to):
            return m.group(0)
        c0, co0, r0, ro0 = map(int, fr.groups()); c1, co1, r1, ro1 = map(int, to.groups())
        cx = sum(colpx.get(c + 1, 72) * 9525 for c in range(c0, c1)) - co0 + co1
        cy = int(sum(rowpt.get(r + 1, 15) * 12700 for r in range(r0, r1))) - ro0 + ro1
        if cx <= 0 or cy <= 0:
            return m.group(0)
        body = body.replace(to.group(0), '').replace(fr.group(0), fr.group(0) + f'<xdr:ext cx="{cx}" cy="{cy}"/>', 1)
        return f'<xdr:oneCellAnchor>{body}</xdr:oneCellAnchor>'
    return re.sub(r'<xdr:twoCellAnchor\b([^>]*)>(.*?)</xdr:twoCellAnchor>', size, dx, flags=re.S)


def fit_sheet(b, sheet, ws):
    """계산과정·검정곡선 같은 시트를 쪽에 꽉 차게: 표·차트 묶음이 갈라지지 않게 쪽을 나누고, 열 너비·줄 높이를 늘려 쪽을 채운다.
    배율·쪽 나눔은 숫자로 고정 (엑셀·LibreOffice가 따로 쪽을 나누지 않게)"""
    from openpyxl.utils import get_column_letter as CL, range_boundaries
    pa = ws.print_area
    pa = pa[0] if isinstance(pa, (list, tuple)) else pa
    cells = [(c.row, c.column) for row in ws.iter_rows() for c in row if c.value not in (None, '')]
    if not cells:
        return None
    if pa:
        c1, r1, c2, r2 = range_boundaries(pa.split('!')[-1].replace('$', ''))
    else:
        r1, c1 = 1, 1; r2 = max(r for r, _ in cells); c2 = max(c for _, c in cells)
    # 그림·차트가 걸친 줄·열까지
    dp = _drawing_part(b, sheet); anc = []
    if dp and dp in b.parts:
        for m in re.finditer(r'<xdr:twoCellAnchor\b[^>]*>(.*?)</xdr:twoCellAnchor>', b.text(dp), re.S):
            fr = re.search(r'<xdr:from><xdr:col>(\d+)</xdr:col>.*?<xdr:row>(\d+)</xdr:row>', m.group(1), re.S)
            to = re.search(r'<xdr:to><xdr:col>(\d+)</xdr:col>.*?<xdr:row>(\d+)</xdr:row>', m.group(1), re.S)
            if fr and to:
                anc.append((int(fr.group(2)) + 1, int(to.group(2)) + 1, int(fr.group(1)) + 1, int(to.group(1)) + 1))
    if not pa and anc:
        r2 = max([r2] + [a[1] for a in anc]); c2 = max([c2] + [a[3] for a in anc])
    # 병합 칸이 인쇄 영역 경계에 걸리면 영역을 넓힌다 (반쪽만 찍히지 않게)
    grow = True
    while grow:
        grow = False
        for g in ws.merged_cells.ranges:
            if g.max_row < r1 or g.min_row > r2 or g.max_col < c1 or g.min_col > c2:
                continue
            if g.max_col > c2 or g.max_row > r2 or g.min_col < c1 or g.min_row < r1:
                c1, r1, c2, r2 = min(c1, g.min_col), min(r1, g.min_row), max(c2, g.max_col), max(r2, g.max_row); grow = True
    # 줄 높이 (정하지 않은 줄은 글자 줄 수로 정해 고정)
    H = {}
    for r in range(r1, r2 + 1):
        d = ws.row_dimensions[r]
        H[r] = 0 if d.hidden else (d.height if d.height else _auto_h(ws, r, c1, c2))
    CW, dflt_w = col_widths(b.sx(sheet))
    cw = {c: CW.get(c, (dflt_w, False)) for c in range(c1, c2 + 1)}
    Wpx = {c: (0 if cw[c][1] else _px(cw[c][0])) for c in range(c1, c2 + 1)}
    W = sum(Wpx.values()) * 0.75
    # 묶음: 빈 줄로 나뉜 표 (값·테두리가 있는 줄), 차트가 걸친 줄은 그 묶음에
    filled = set()
    for row in ws.iter_rows(min_row=r1, max_row=r2, min_col=c1, max_col=c2):
        for c in row:
            if c.value not in (None, '') or (c.border and any(getattr(c.border, sd).style for sd in ('top', 'bottom', 'left', 'right'))):
                filled.add(c.row)
    vals = set(filled)                         # 값·테두리가 있는 줄
    for a in anc:
        filled.update(range(a[0], a[1] + 1))
    cut_bad = {x for a in anc for x in range(a[0] + 1, a[1] + 1)}     # 이 줄 앞에서 나누면 그림·차트가 잘림
    merged_spans = [(g.min_row, g.max_row) for g in ws.merged_cells.ranges if g.max_row > g.min_row]
    blocks = []; r = r1
    while r <= r2:
        if r in filled:
            a = r
            while r + 1 <= r2 and r + 1 in filled:
                r += 1
            blocks.append([a, r])
        r += 1
    # 묶음 사이 빈 줄은 앞 묶음에 붙여 둠 (쪽 나눔은 묶음 시작 앞에서만)
    for i in range(len(blocks) - 1):
        blocks[i][1] = blocks[i + 1][0] - 1
    blocks[0][0] = r1; blocks[-1][1] = r2
    bh = [sum(H[x] for x in range(a, e + 1)) for a, e in blocks]
    def plan(land, s):
        uw = ((PAGE_H if land else PAGE_W) - (MARG['left'] + MARG['right']) * 72) * 0.995
        uh = ((PAGE_W if land else PAGE_H) - (MARG['top'] + MARG['bottom']) * 72) * 0.97
        if W * s > uw + 0.1:
            return None
        pages = []; cur = []; acc = 0
        for i, h in enumerate(bh):
            if h * s > uh:                        # 한 쪽보다 긴 묶음: 줄 단위로 (머리 줄 반복)
                if cur:
                    pages.append(cur); cur = []; acc = 0
                a, e = blocks[i]; start = a; acc2 = 0
                x = a
                while x <= e:
                    if acc2 + H[x] * s > uh and x > start:
                        # 나눌 줄: 그림·차트·병합 칸을 자르지 않는 곳, 되도록 빈 줄 바로 다음 (표 사이)
                        # 한 쪽보다 긴 병합 칸(예: 120줄을 덮는 K 값 칸)은 어차피 나뉘므로 따지지 않는다
                        merged_bad = {z for m0, m1 in merged_spans if sum(H[q] for q in range(m0, m1 + 1)) * s <= uh for z in range(m0 + 1, m1 + 1)}
                        ok_ = [y for y in range(start + 1, x + 1) if y not in cut_bad and y not in merged_bad]
                        gap = [y for y in ok_ if y - 1 not in vals]
                        y = (gap or ok_ or [x])[-1]
                        if gap and ok_ and sum(H[z] for z in range(gap[-1], ok_[-1])) * s > uh * 0.35:
                            y = ok_[-1]                 # 빈 줄이 너무 위쪽이면 표 안 줄 경계에서
                        pages.append([(start, y - 1)]); start = y
                        acc2 = (H[r1] * s if start > r1 else 0) + sum(H[z] for z in range(y, x)) * s
                    acc2 += H[x] * s
                    x += 1
                cur = [(start, e)]; acc = acc2
                continue
            if acc + h * s > uh and cur:
                pages.append(cur); cur = []; acc = 0
            cur.append(tuple(blocks[i])); acc += h * s
        if cur:
            pages.append(cur)
        return pages, uw, uh
    best = None
    for land in (False, True):
        uw = ((PAGE_H if land else PAGE_W) - (MARG['left'] + MARG['right']) * 72) * 0.995
        s_w = math.floor(min(1.0, uw / W) * 1000) / 1000      # 너비 맞춤 배율 (엑셀 열 너비 기준; 글꼴이 더 넓은 프로그램에서는 배율만 작아져 높이는 남음)
        for k in range(0, 16):                    # 배율을 15 %까지 줄여 쪽 수가 줄면 그 배율로 (열을 넓혀 너비 맞춤 배율이 그 값이 되게)
            s = round(s_w * (1 - k * 0.01), 3)
            r_ = plan(land, s)
            if r_:
                key = (len(r_[0]), -s)
                if best is None or key < best[0]:
                    best = (key, land, s, r_)
    if best is None:
        return None
    _, land, s, (pages, uw, uh) = best
    # 쪽 채우기: 열 너비를 늘려 너비를 채우고 (1.8배까지), 쪽마다 줄 높이를 늘려 높이를 채운다 (줄마다 찍힌 높이 약 1.5 cm까지, 머리 줄 반복 쪽은 그 줄 몫 뺌).
    # 그래도 남는 한 쪽짜리는 위아래 가운데에 놓는다
    fw = min(1.8, uw / (W * s))
    titles = None
    if dp and dp in b.parts and anc:
        b.put(dp, _freeze_anchors(b.text(dp), {c: Wpx[c] for c in Wpx}, dict(H)))   # 줄·열을 늘려도 그림·차트 크기와 비율은 그대로
    for pg in pages:
        rows_ = [x for a, e in pg for x in range(a, e + 1)]
        ph = sum(H[x] for x in rows_) * s
        if len(pages) > 1 and any(e - a > 30 for a, e in pg) and titles is None:
            titles = r1
        extra = (H[titles] * s if titles and pg[0][0] != r1 else 0)
        # 줄마다 f배 늘리되, 찍힌 높이가 42pt(약 1.5 cm)를 넘는 줄은 거기서 멈춤 (원래 그보다 높은 머리 줄은 그대로).
        # 쪽 높이를 넘지 않는 가장 큰 f (3배까지)를 이분법으로 찾는다
        avail = uh - extra
        grown = lambda f: sum(min(H[x] * s * f, max(H[x] * s, 42.0)) for x in rows_)
        lo_, hi_ = 1.0, 3.0
        if grown(hi_) <= avail:
            lo_ = hi_
        else:
            for _ in range(40):
                mid = (lo_ + hi_) / 2
                lo_, hi_ = (mid, hi_) if grown(mid) <= avail else (lo_, mid)
        fh = lo_
        if fh > 1.01:
            for x in rows_:
                if H[x]:
                    H[x] = round(min(H[x] * fh, max(H[x], 42.0 / s)) - 0.005, 2)
    # 시트 XML: 줄 높이·열 너비
    x = b.sx(sheet)
    for r_, h in H.items():
        if h:
            x = _set_row(x, r_, h)
    if fw > 1.01:
        cols = re.search(r'<cols>(.*?)</cols>', x, re.S)
        styl = {}
        for m in re.finditer(r'<col\b([^>]*)/>', cols.group(1) if cols else ''):
            st = re.search(r'style="(\d+)"', m.group(1))
            if st:
                for c in range(int(re.search(r'min="(\d+)"', m.group(1)).group(1)), int(re.search(r'max="(\d+)"', m.group(1)).group(1)) + 1):
                    styl[c] = st.group(1)
        new = ''.join(f'<col min="{c}" max="{c}" width="{round(cw[c][0] * fw, 2)}"' + (f' style="{styl[c]}"' if c in styl else '') + ' customWidth="1"' + (' hidden="1"' if cw[c][1] else '') + '/>' for c in range(c1, c2 + 1))
        keep = ''
        if cols:
            for m in re.finditer(r'<col min="(\d+)" max="(\d+)"[^>]*/>', cols.group(1)):
                lo, hi = int(m.group(1)), int(m.group(2))
                if hi < c1 or lo > c2:
                    keep += m.group(0)
            x = x.replace(cols.group(0), f'<cols>{new}{keep}</cols>')
        else:
            x = x.replace('<sheetData', f'<cols>{new}</cols><sheetData', 1)
    breaks = [pg[0][0] - 1 for pg in pages[1:]]
    x = _set_page(x, landscape=land, fit=True, breaks=breaks)      # 너비 1쪽 맞춤 + 수동 쪽 나눔
    one = len(pages) == 1 and sum(H[x] for a_, e_ in pages[0] for x in range(a_, e_ + 1)) * s < uh * 0.9
    x = re.sub(r'<printOptions[^>]*/>', '', x)
    x = x.replace('<pageMargins', '<printOptions horizontalCentered="1"' + (' verticalCentered="1"' if one else '') + '/><pageMargins', 1)
    b.setsx(sheet, x)
    b.print_area(sheet, f'{CL(c1)}{r1}:{CL(c2)}{r2}')
    if titles:
        _print_titles(b, sheet, titles)
    return f'{sheet}: {"가로" if land else "세로"} {len(pages)}쪽, 배율 {int(s * 100)}%, 표 묶음 {len(blocks)}개를 쪽 사이에서 가르지 않음, 너비 ×{fw:.2f}' + (f', 머리 줄 {titles}행 쪽마다 반복' if titles else '')


def _print_titles(b, sheet, row):
    wbx = b.text('xl/workbook.xml'); idx = list(b.sheets).index(sheet)
    q = "'" + sheet.replace("'", "''") + "'"
    dn = f'<definedName name="_xlnm.Print_Titles" localSheetId="{idx}">{q}!${row}:${row}</definedName>'
    wbx = re.sub(rf'<definedName name="_xlnm.Print_Titles" localSheetId="{idx}">[^<]*</definedName>', '', wbx)
    wbx = wbx.replace('<definedNames>', '<definedNames>' + dn, 1) if '<definedNames>' in wbx else wbx.replace('</sheets>', f'</sheets><definedNames>{dn}</definedNames>', 1)
    b.put('xl/workbook.xml', wbx)


def tidy(b, others, wb):
    log = []
    for sh in others:
        if sh not in b.sheets or sh not in wb.sheetnames:
            continue
        ws = wb[sh]
        texts = [c for row in ws.iter_rows() for c in row if c.value not in (None, '')]
        has_text = any(not PLACEHOLDER.match(str(c.value).strip()) for c in texts)
        m = None if has_text else picture_sheet(b, sh, first_row=max((c.row for c in texts), default=0) + (1 if texts else 0))
        if m is None and has_text:
            m = fit_sheet(b, sh, ws) or content_sheet(b, sh, ws)      # 계산과정만 들어옴 (기기 자료 시트는 run_all에서 뺌)
        if m:
            log.append(m)
    return log


def trim_form_merges(b, sheet):
    """기록부: 인쇄 영역 오른쪽 경계를 넘는 병합 칸을 경계까지로 줄인다 (원본에서 한 칸 더 병합된 곳; 글자가 반쪽만 찍히지 않게)"""
    from openpyxl.utils import range_boundaries, get_column_letter as CL
    wbx = b.text('xl/workbook.xml'); idx = list(b.sheets).index(sheet)
    m = re.search(rf'<definedName name="_xlnm.Print_Area" localSheetId="{idx}">([^<]*)</definedName>', wbx)
    if not m:
        return []
    c1, r1, c2, r2 = range_boundaries(m.group(1).split('!')[-1].replace('$', ''))
    x = b.sx(sheet); out = []
    def fix(mm):
        a1, b1, a2, b2 = range_boundaries(mm.group(1))
        if a1 <= c2 < a2 and r1 <= b1 <= r2:
            out.append(f'{sheet}!{mm.group(1)} → {CL(a1)}{b1}:{CL(c2)}{b2}')
            if a1 == c2 and b1 == b2:
                return ''
            return f'<mergeCell ref="{CL(a1)}{b1}:{CL(c2)}{b2}"/>'
        return mm.group(0)
    x2 = re.sub(r'<mergeCell ref="([A-Z]+\d+:[A-Z]+\d+)"/>', fix, x)
    if out:
        n = len(re.findall(r'<mergeCell ref=', x2))
        x2 = re.sub(r'<mergeCells count="\d+"', f'<mergeCells count="{n}"', x2)
        b.setsx(sheet, x2)
        # 병합 끝 칸의 오른쪽 테두리(표 바깥 굵은 선)를 새 끝 칸(인쇄 영역 끝 열)으로 옮긴다
        from form_style import Styles, _set_styles, _cell_style
        S = Styles(b); new = {}
        for o in out:
            old = re.search(r'!([A-Z]+)(\d+):([A-Z]+)(\d+) →', o)
            r_top, r_bot, old_last = int(old.group(2)), int(old.group(4)), old.group(3)
            for r in range(r_top, r_bot + 1):
                xs = b.sx(sheet)
                bx_old = S.border_xml((re.search(r'borderId="(\d+)"', S.xf(_cell_style(xs, f'{old_last}{r}', None, {}, 0))) or [0, '0'])[1])
                right = re.search(r'<right style="([^"]+)"', bx_old)
                st = _cell_style(xs, f'{CL(c2)}{r}', None, {}, c2)
                bx = S.border_xml((re.search(r'borderId="(\d+)"', S.xf(st)) or [0, '0'])[1])
                cur = [(re.search(rf'<{n_} style="([^"]+)"', bx) or [0, None])[1] for n_ in ('left', 'right', 'top', 'bottom')]
                if right and cur[1] != right.group(1):
                    new[(r, c2)] = S.with_border(st, (cur[0], right.group(1), cur[2], cur[3]))
        if new:
            _set_styles(b, sheet, new)
    return [f'병합 칸을 인쇄 영역 안으로: {o}' for o in out]
