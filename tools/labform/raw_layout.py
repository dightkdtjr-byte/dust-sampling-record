# raw data·계산과정 시트 인쇄 모양 (기록부 시트가 아닌 시트)
#  · 그림만 있는 시트(기기 출력물 사진): 그림을 원래 순서(왼쪽 열 위→아래, 다음 열)대로 한 줄로 세로로 다시 놓고,
#    한 쪽 너비·높이에 들어가게 줄이며, 그림이 쪽 사이에서 잘리지 않게 쪽 나눔. 100% 배율 (작게 몰려 읽을 수 없던 것)
#  · 칸 내용이 있는 시트: 내용 폭이 세로 쪽보다 넓으면 가로, 좁은 여백, 너비 한 쪽 맞춤 (높이는 여러 쪽)
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
    s = _set_page(s, landscape=False, fit=False, breaks=breaks)
    b.setsx(sheet, s)
    cols = max(1, int(usable_w // 48) + 1)
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
            m = content_sheet(b, sh, ws)
        if m:
            log.append(m)
    return log
