# 시험기록부 양식의 인쇄 모양을 한 기준으로 맞춘다.
#  ① 서명란(담당·기술책임자)의 너비·높이를 모든 양식에서 같게  ② 인쇄 폭·배율·여백을 같게
#  ③ 열을 좁혀 글자가 넘치게 된 칸은 줄바꿈을 켜고 줄 높이를 늘림  ④ 쪽은 항목 경계에서만 나누고 한 쪽에 다 들어가게
#  ⑤ 문서번호는 칸 대신 바닥글에 넣어 모든 쪽에 찍히게
import io, re
import openpyxl
from openpyxl.utils import get_column_letter as CL, column_index_from_string as CI, range_boundaries
from tplkit import CELL, split_ref, escape

MDW = 8                       # 기본 글꼴(돋움 11)의 숫자 한 자 너비(px)
SIGN_W = 17.5                 # 서명 칸 하나의 너비(글자 수)
TOTAL_PX = 1048               # 인쇄되는 열 너비의 합(px): 먼지 양식(A~J)의 값
ROW1, ROW2 = 30.75, 59.25     # 서명란 이름 줄·서명 줄 높이(pt)
# 인쇄 배율은 숫자로 박지 않고 "한 쪽 너비에 맞춤"으로 둔다: 배율을 박으면 프린터·글꼴에 따라 마지막 열이 다음 쪽으로 넘어갈 수 있다.
# 모든 양식의 열 너비 합(TOTAL_PX)과 여백이 같으므로 엑셀이 정하는 배율도 모두 같다 (이 너비·여백에서는 65~66%).
MARGINS = dict(left=0.45, right=0.45, top=0.45, bottom=0.45, header=0.2, footer=0.2)
WORST_SCALE = 67              # 쪽 높이를 따질 때 가정하는 가장 큰 배율(%): 이보다 작게 찍히면 여유가 더 생긴다
PAGE_PT = (11.69 - MARGINS['top'] - MARGINS['bottom']) * 72 / (WORST_SCALE / 100) * 0.99   # 한 쪽에 들어가는 줄 높이 합(pt), 여유 1%
DOCNO = re.compile(r'[A-Z]{2,6}-[A-Z]{1,4}-\d{1,3}-\d{1,3}')
# 문서번호 칸: 칸 전체가 "■ MAMO-QP-14-01" 꼴 (시료명 "0629-TCW-#A1-DMDS-FB-1-1" 같은 것을 문서번호로 잘못 읽지 않게)
DOCNO_CELL = re.compile(r'\s*■?\s*[A-Z]{2,6}-[A-Z]{1,4}-\d{1,3}-\d{1,3}\s*')
is_docno = lambda v: isinstance(v, str) and DOCNO_CELL.fullmatch(v) is not None

px = lambda w: int(w * MDW + 5)
width_of = lambda p: round((p - 5) / MDW, 4)


def load(b):
    buf = io.BytesIO()
    import zipfile
    with zipfile.ZipFile(buf, 'w') as z:
        for n, d in b.parts.items():
            z.writestr(n, d)
    buf.seek(0)
    return openpyxl.load_workbook(buf, keep_vba=False)


def em(ch, dotum):
    o = ord(ch)
    if ch == ' ':
        return 0.32
    if 0xAC00 <= o <= 0xD7A3 or 0x3130 <= o <= 0x318F or 0x4E00 <= o <= 0x9FFF or 0xFF00 <= o <= 0xFFEF or 0x3000 <= o <= 0x33FF or 0x2460 <= o <= 0x27BF or 0x2100 <= o <= 0x22FF or 0x25A0 <= o <= 0x25FF:
        return 1.0 if dotum else 0.95
    if ch.isdigit():
        return 0.56
    if ch.isupper():
        return 0.68
    if ch.islower():
        return 0.55
    if o > 0x2000:
        return 0.9
    return 0.4


def text_px(text, size, bold, dotum):
    return sum(em(ch, dotum) for ch in text) * size * 96 / 72 * (1.06 if bold else 1)


def wrap_lines(text, inner, size, bold, dotum):
    # 낱말 단위로 줄을 바꾸고, 한 낱말이 칸보다 길면 글자 단위로 자른다
    lines = 0
    for para in str(text).split('\n'):
        cur = 0.0; n = 1
        for word in re.split(r'(\s+)', para):
            w = text_px(word, size, bold, dotum)
            if cur + w <= inner:
                cur += w; continue
            if w > inner and not word.isspace():
                for ch in word:
                    cw = text_px(ch, size, bold, dotum)
                    if cur + cw > inner:
                        n += 1; cur = 0.0
                    cur += cw
            elif not word.isspace():
                n += 1; cur = w
        lines += n
    return lines


class Layout:
    """한 시트의 인쇄 범위·열 너비·줄 높이·병합을 읽고 고친다."""

    def __init__(self, b, sheet, area):
        self.b, self.sheet = b, sheet
        c1, r1, c2, r2 = range_boundaries(area)
        self.c1, self.c2, self.r1, self.r2 = c1, c2, r1, r2
        self.ws = load(b)[sheet]
        ws = self.ws
        dflt_w = ws.sheet_format.defaultColWidth or 8.43
        widths, _ = b.col_widths(sheet)
        self.w = {c: widths.get(c, dflt_w) for c in range(c1, c2 + 1)}
        self.w0 = dict(self.w)
        dflt_h = ws.sheet_format.defaultRowHeight or 13.5
        self.h = {r: (ws.row_dimensions[r].height if ws.row_dimensions[r].height is not None else dflt_h) for r in range(r1, r2 + 1)}
        self.h0 = dict(self.h)
        self.merge = {}
        self.merges = []
        for m in ws.merged_cells.ranges:
            self.merges.append((m.min_col, m.min_row, m.max_col, m.max_row))
            for r in range(m.min_row, m.max_row + 1):
                for c in range(m.min_col, m.max_col + 1):
                    self.merge[(r, c)] = (m.min_col, m.min_row, m.max_col, m.max_row)
        self.notes = []

    def box(self, r, c):
        return self.merge.get((r, c), (c, r, c, r))

    # ── ① 서명란 ──
    def sign_groups(self):
        groups = []
        for c in range(self.c1, self.c2 + 1):
            v = self.ws.cell(self.r1, c).value
            if isinstance(v, str) and re.sub(r'\s', '', v) in ('담당', '기술책임자', '담당자', '품질책임자', '검토', '승인', '작성'):
                a, _, z, _ = self.box(self.r1, c)
                groups.append((a, z))
        return groups

    def set_widths(self, like=None):
        groups = self.sign_groups()
        if len(groups) < 2:
            # 서명란이 없는 쪽(물질별 시트 등): 같은 열 짜임의 첫 장과 같은 너비를 쓰고, 없으면 전체를 같은 비율로 맞춘다
            if like is not None and set(like.w) == set(self.w):
                self.w = dict(like.w)
                self.notes.append('서명란 없음 → 첫 장과 같은 열 너비')
            else:
                tot = sum(self.w0.values()); acc = 0; cols = sorted(self.w0)
                for i, c in enumerate(cols):
                    p = round(TOTAL_PX * self.w0[c] / tot) if i < len(cols) - 1 else TOTAL_PX - acc
                    acc += p; self.w[c] = width_of(p)
                self.notes.append('서명란 없음 → 열 너비를 같은 비율로 맞춤')
            return
        sign_cols = {c for a, z in groups for c in range(a, z + 1)}
        newpx = {}
        for a, z in groups:
            tot = sum(self.w0[c] for c in range(a, z + 1))
            target = px(SIGN_W) if a == z else None
            if a == z:
                newpx[a] = target
            else:
                # 여러 열을 합친 서명 칸: 합이 한 칸짜리 서명 칸과 같은 px이 되게 나눈다
                want = px(SIGN_W)
                parts = [max(12, round(want * self.w0[c] / tot)) for c in range(a, z + 1)]
                parts[-1] += want - sum(parts)
                for c, p in zip(range(a, z + 1), parts):
                    newpx[c] = p
        rest = [c for c in range(self.c1, self.c2 + 1) if c not in sign_cols]
        rest_target = TOTAL_PX - sum(newpx.values())
        rest_tot = sum(self.w0[c] for c in rest)
        acc = 0
        for i, c in enumerate(rest):
            p = round(rest_target * self.w0[c] / rest_tot)
            if i == len(rest) - 1:
                p = rest_target - acc
            newpx[c] = p; acc += p
        for c, p in newpx.items():
            self.w[c] = width_of(p)
        assert sum(px(self.w[c]) for c in self.w) == TOTAL_PX, (self.sheet, sum(px(self.w[c]) for c in self.w))
        self.h[self.r1] = ROW1; self.h[self.r1 + 1] = ROW2
        self.notes.append(f'서명란 열 {[(CL(a), CL(z)) for a, z in groups]} → 칸마다 {SIGN_W}자 너비, 나머지 열 ×{rest_target / sum(px(self.w0[c]) for c in rest):.3f}')

    # ── ③ 글자 맞춤 ──
    def fit_text(self):
        wrap_cells = []
        for r in range(self.r1, self.r2 + 1):
            for c in range(self.c1, self.c2 + 1):
                cell = self.ws.cell(r, c)
                v = cell.value
                if type(cell).__name__ == 'MergedCell' or not isinstance(v, str) or v.startswith('=') or not v.strip():
                    continue
                a, top, z, bot = self.box(r, c)
                z = min(z, self.c2)
                size = cell.font.sz or 11; bold = bool(cell.font.b); dotum = (cell.font.name or '') in ('돋움', '굴림', '바탕', '돋움체', '굴림체')
                inner0 = sum(px(self.w0[k]) for k in range(a, z + 1)) - 6
                inner1 = sum(px(self.w[k]) for k in range(a, z + 1)) - 6
                wrap = bool(cell.alignment.wrap_text)
                spill = (a == z and top == bot and c < self.c2 and cell.alignment.horizontal in (None, 'general', 'left')
                         and self.ws.cell(r, c + 1).value is None and (r, c + 1) not in self.merge)
                if inner1 >= inner0 and (spill or not self._overflows(v, inner1, size, bold, dotum, wrap, top, bot)):
                    continue
                lh = size * 1.45
                avail = sum(self.h[k] for k in range(top, bot + 1))
                if wrap:
                    l0 = wrap_lines(v, inner0, size, bold, dotum); l1 = wrap_lines(v, inner1, size, bold, dotum)
                    if l1 > l0 or l1 * lh + 3 > avail:
                        add = max(0.0, l1 * lh + 3 - avail)
                        if add > 0:
                            self.h[bot] += add
                            self.notes.append(f'{CL(c)}{r} "{v[:14].strip()}…" 줄 수 {l0}→{l1}, {bot}행 +{add:.1f}pt')
                else:
                    longest = max(text_px(line, size, bold, dotum) for line in v.split('\n'))
                    if longest > inner1 and not spill:
                        l1 = wrap_lines(v, inner1, size, bold, dotum)
                        add = max(0.0, l1 * lh + 3 - avail)
                        wrap_cells.append(f'{CL(c)}{r}')
                        if add > 0:
                            self.h[bot] += add
                        self.notes.append(f'{CL(c)}{r} "{v[:14].strip()}" 칸보다 길어져 줄바꿈 켬' + (f', {bot}행 +{add:.1f}pt' if add > 0 else ''))
        return wrap_cells

    def _overflows(self, v, inner, size, bold, dotum, wrap, top, bot):
        avail = sum(self.h[k] for k in range(top, bot + 1)); lh = size * 1.45
        if wrap:
            return wrap_lines(v, inner, size, bold, dotum) * lh + 3 > avail
        return max(text_px(t, size, bold, dotum) for t in v.split('\n')) > inner

    # ── ④ 쪽 나누기 ──
    def row_empty(self, r):
        return all(self.ws.cell(r, c).value is None for c in range(self.c1, self.c2 + 1))

    def can_break_after(self, r):
        return not any(top <= r < bot for (_, top, _, bot) in self.merges)

    def section_starts(self):
        # A열에 항목 번호가 있는 줄, 또는 A열이 회색 머리 칸인 줄(번호 없는 표 머리, 예: "시료명")이 항목의 시작
        from layout_rules import _gray
        out = []
        for r in range(self.r1, self.r2 + 1):
            cell = self.ws.cell(r, self.c1)
            if type(cell).__name__ == 'MergedCell' or cell.value is None:
                continue
            v = str(cell.value).strip()
            if re.fullmatch(r'\d{1,2}', v) or (isinstance(cell.value, str) and v and not v.startswith('=') and _gray(cell)):
                out.append(r)
        return out

    def paginate(self, hints):
        H = self.h
        total = lambda a, b: sum(H[r] for r in range(a, b + 1))
        starts = set(self.section_starts())
        strong = [r for r in range(self.r1, self.r2) if (r + 1) in starts and self.can_break_after(r)]
        # 원본의 쪽 나눔 자리는 다음 줄이 항목 시작일 때만 쓴다 (아니면 항목 경계에서 다시 정한다)
        ends = sorted(set(h for h in hints if self.r1 <= h < self.r2 and self.can_break_after(h) and (h + 1) in starts) | {self.r2})
        pages = []
        a = self.r1
        for e in ends:
            # 원래 한 쪽이던 구간이 한 쪽에 안 들어가면 항목 경계에서 더 나눈다
            while total(a, e) > PAGE_PT:
                cand = [r for r in strong if a <= r < e and total(a, r) <= PAGE_PT]
                if not cand:
                    cand = [r for r in range(a, e) if self.can_break_after(r) and total(a, r) <= PAGE_PT]
                if not cand:
                    break
                pages.append((a, cand[-1])); a = cand[-1] + 1
            pages.append((a, e)); a = e + 1
        # 이웃한 쪽을 합쳐도 (빈 줄을 조금 줄이면) 한 쪽에 들어가면 합친다
        i = 0
        while i < len(pages) - 1:
            a, e = pages[i][0], pages[i + 1][1]
            if self.squeeze(a, e, dry=True):
                self.squeeze(a, e)
                pages[i:i + 2] = [(a, e)]
            else:
                i += 1
        for a, e in pages:
            if total(a, e) > PAGE_PT and not self.squeeze(a, e):
                self.notes.append(f'!! {a}~{e}행이 한 쪽({PAGE_PT:.0f}pt)을 넘음: {total(a, e):.0f}pt')
        self.pages = pages
        return pages

    def squeeze(self, a, e, dry=False):
        # 넘치는 만큼: 먼저 내용 없는 큰 줄(여백·그림 자리)을 35%까지, 그래도 모자라면 넉넉한 줄 전체를 3%까지 줄인다
        H = self.h
        over = sum(H[r] for r in range(a, e + 1)) - PAGE_PT
        if over <= 0:
            return True
        fill = [r for r in range(a, e + 1) if self.row_empty(r) and H[r] >= 45]
        room = sum(H[r] for r in fill) * 0.35
        loose = [r for r in range(a, e + 1) if r not in fill and H[r] >= 26 and r not in (self.r1, self.r1 + 1)
                 and r not in getattr(self, 'tight', set())]   # 글자가 꽉 찬 줄은 줄이지 않는다
        room2 = sum(H[r] for r in loose) * 0.03
        if over > room + room2:
            return False
        if dry:
            return True
        if fill:
            cut = min(over, room); f = cut / sum(H[r] for r in fill)
            for r in fill:
                H[r] = round(H[r] * (1 - f), 2)
            over -= cut
            self.notes.append(f'{a}~{e}행: 빈 줄 {fill} 높이를 {f * 100:.0f}% 줄임')
        if over > 0.01:
            f = over / sum(H[r] for r in loose)
            for r in loose:
                H[r] = round(H[r] * (1 - f) - 0.005, 2)
            self.notes.append(f'{a}~{e}행: 줄 높이를 고르게 {f * 100:.1f}% 줄임')
        return True

    # ── 고친 값을 XML에 적는다 ──
    def write(self, wrap_cells, docno):
        b, sheet = self.b, self.sheet
        s = b.sx(sheet)
        # 열 너비
        cols = re.search(r'<cols>.*?</cols>', s, re.S)
        attrs = {}; tail = []
        for m in re.finditer(r'<col min="(\d+)" max="(\d+)" width="([\d.]+)"([^>]*)/>', cols.group(0)):
            lo, hi = int(m.group(1)), int(m.group(2))
            extra = re.sub(r'\s*(customWidth|bestFit)="[^"]*"', '', m.group(4))
            for c in range(lo, min(hi, self.c2) + 1):
                attrs[c] = extra
            if hi > self.c2:
                tail.append(f'<col min="{max(lo, self.c2 + 1)}" max="{hi}" width="{m.group(3)}"{m.group(4)}/>')
        items = [f'<col min="{c}" max="{c}" width="{self.w[c]}"{attrs.get(c, "")} customWidth="1"/>' for c in range(self.c1, self.c2 + 1)]
        s = s.replace(cols.group(0), '<cols>' + ''.join(items) + ''.join(tail) + '</cols>')
        # 줄 높이
        for r in range(self.r1, self.r2 + 1):
            if abs(self.h[r] - self.h0[r]) < 0.005:
                continue
            m = re.search(rf'<row r="{r}"([^>]*?)(/?)>', s)
            if not m:
                continue
            a = re.sub(r'\s*(ht|customHeight)="[^"]*"', '', m.group(1))
            s = s[:m.start()] + f'<row r="{r}"{a} ht="{self.h[r]:g}" customHeight="1"{m.group(2)}>' + s[m.end():]
        b.setsx(sheet, s)
        # 줄바꿈 켜기
        cache = {}
        for ref in wrap_cells:
            st = b.style_of(sheet, ref)
            if st not in cache:
                cache[st] = style_with_wrap(b, st)
            s = b.sx(sheet)
            m = CELL(ref).search(s)
            new = re.sub(r' s="\d+"', f' s="{cache[st]}"', m.group(0), count=1)
            b.setsx(sheet, s[:m.start()] + new + s[m.end():])
        # 문서번호 칸은 비우고 바닥글로
        for r in range(self.r1, self.r2 + 1):
            for c in range(self.c1, self.c2 + 1):
                cell = self.ws.cell(r, c)
                if type(cell).__name__ != 'MergedCell' and is_docno(cell.value):
                    b.setv(sheet, f'{CL(c)}{r}', None)
        page_setup(b, sheet, docno, breaks=[e for _, e in self.pages[:-1]], last_col=self.c2)


def style_with_wrap(b, base):
    st = b.text('xl/styles.xml')
    cx = re.search(r'<cellXfs count="(\d+)">(.*?)</cellXfs>', st, re.S)
    xfs = re.findall(r'<xf\b[^>]*?(?:/>|>.*?</xf>)', cx.group(2), re.S)
    x = xfs[int(base)]
    if '<alignment' in x:
        x = re.sub(r'<alignment\b([^>]*?)/>', lambda m: '<alignment' + re.sub(r'\s*(wrapText|shrinkToFit)="[^"]*"', '', m.group(1)) + ' wrapText="1"/>', x, count=1)
    else:
        x = re.sub(r'<xf\b([^>]*?)/>', r'<xf\1><alignment wrapText="1"/></xf>', x, count=1) if x.endswith('/>') else x.replace('</xf>', '<alignment wrapText="1"/></xf>')
    if 'applyAlignment' not in x:
        x = x.replace('<xf ', '<xf applyAlignment="1" ', 1)
    st = st.replace(cx.group(0), f'<cellXfs count="{int(cx.group(1)) + 1}">{cx.group(2)}{x}</cellXfs>')
    b.put('xl/styles.xml', st)
    return str(int(cx.group(1)))


def page_setup(b, sheet, docno, breaks=(), last_col=None, fixed_scale=False, landscape=False):
    s = b.sx(sheet)
    pm = '<pageMargins ' + ' '.join(f'{k}="{v}"' for k, v in MARGINS.items()) + '/>'
    assert '<pageMargins' in s, sheet
    s = re.sub(r'<pageMargins[^>]*/>', pm, s)
    if fixed_scale:
        ps = f'<pageSetup paperSize="9" scale="{fixed_scale}" orientation="portrait"/>'
    else:
        ps = f'<pageSetup paperSize="9" fitToWidth="1" fitToHeight="0" orientation="{"landscape" if landscape else "portrait"}"/>'
    s = re.sub(r'<pageSetup[^>]*/>', '', s)
    s = re.sub(r'<headerFooter\b[^>]*?(?:/>|>.*?</headerFooter>)', '', s, flags=re.S)
    s = re.sub(r'<printOptions[^>]*/>', '', s)
    s = re.sub(r'<rowBreaks.*?</rowBreaks>', '', s, flags=re.S)
    s = re.sub(r'<colBreaks.*?</colBreaks>', '', s, flags=re.S)
    # 문서번호: 인쇄 배율과 상관없이 같은 크기로, 오른쪽 여백선에 맞춰 모든 쪽 아래에
    hf = ''
    if docno:
        hf = f'<headerFooter scaleWithDoc="0" alignWithMargins="1"><oddFooter>&amp;R&amp;"맑은 고딕,Regular"&amp;9{escape(docno)}</oddFooter></headerFooter>'
    rb = ''
    if breaks:
        rb = f'<rowBreaks count="{len(breaks)}" manualBreakCount="{len(breaks)}">' + ''.join(f'<brk id="{r}" max="{(last_col or 16384) - 1}" man="1"/>' for r in breaks) + '</rowBreaks>'
    s = s.replace(pm, '<printOptions horizontalCentered="1"/>' + pm + ps + hf + rb)
    # 배율 고정 / 한 쪽 너비에 맞춤
    fit = '0' if fixed_scale else '1'
    if '<pageSetUpPr' in s:
        s = re.sub(r'<pageSetUpPr[^>]*/>', f'<pageSetUpPr fitToPage="{fit}"/>', s)
    elif re.search(r'<sheetPr[^>]*/>', s):
        s = re.sub(r'<sheetPr([^>]*)/>', rf'<sheetPr\1><pageSetUpPr fitToPage="{fit}"/></sheetPr>', s, count=1)
    elif '<sheetPr' in s:
        s = s.replace('</sheetPr>', f'<pageSetUpPr fitToPage="{fit}"/></sheetPr>', 1)
    else:
        s = re.sub(r'(<worksheet[^>]*>)', rf'\1<sheetPr><pageSetUpPr fitToPage="{fit}"/></sheetPr>', s, count=1)
    b.setsx(sheet, s)
    rels_part = b.sheets[sheet].replace('worksheets/', 'worksheets/_rels/') + '.rels'
    if rels_part in b.parts:
        b.put(rels_part, re.sub(r'<Relationship\b[^>]*printerSettings[^>]*/>', '', b.text(rels_part)))


def find_docno(b, sheet):
    ws = load(b)[sheet]
    for row in ws.iter_rows():
        for c in row:
            if is_docno(c.value):
                return c.value.strip()
    # 이미 정리한 양식: 문서번호가 칸에서 바닥글로 옮겨져 있다. 다시 돌려도 바닥글을 잃지 않게 바닥글에서 찾는다
    m = re.search(r'<oddFooter>(.*?)</oddFooter>', b.sx(sheet), re.S)
    if m:
        t = re.sub(r'&amp;(?:"[^"]*"|\d+|[A-Z])', '', m.group(1))
        k = DOCNO.search(t)
        if k:
            return t[max(0, t.rfind('■', 0, k.start())):k.end()].strip()
    return None


def docno_rows(b, sheet):
    ws = load(b)[sheet]
    return [c.row for row in ws.iter_rows() for c in row if is_docno(c.value)]


def standardize(b, sheet, area, docno, hints=(), like=None):
    from form_style import trim_area, unify, raw_borders
    from content_rules import unify_fonts, examples, fit_cells, bold_rule
    pre = unify_fonts(b)
    main = any(isinstance(c.value, str) and re.sub(r'\s', '', c.value) in ('담당', '기술책임자') for c in load(b)[sheet][range_boundaries(area)[1]])
    pre += examples(b, sheet, area, fill=main)   # 예시 값은 서명란이 있는 첫 장에만, 날짜 형식은 모든 양식 시트에
    area0 = area
    area, old_r2 = trim_area(load(b)[sheet], area, raw_borders(b.sx(sheet), b.text('xl/styles.xml')))
    # 정렬·굵기는 글자 맞춤보다 먼저 (굵게·가운데로 바뀐 글이 나중에 칸을 넘지 않게)
    from layout_rules import align as _align
    pre += _align(b, sheet, area)
    pre += bold_rule(b, sheet, area)
    lay = Layout(b, sheet, area)
    lay.notes += pre
    if area != area0:
        lay.notes.append(f'표 아래 내용 없는 줄 {range_boundaries(area)[3] + 1}~{range_boundaries(area0)[3]}행을 인쇄 범위에서 뺌')
    lay.set_widths(like)
    wraps = fit_cells(lay)
    hints = sorted(set(hints) | set(docno_rows(b, sheet)))
    pages = lay.paginate(hints)
    lay.write(wraps, docno)
    b.print_area(sheet, area)
    lay.notes += unify(b, sheet, area, [e for _, e in pages[:-1]], old_r2=range_boundaries(area0)[3])
    from layout_rules import align, place
    lay.notes += align(b, sheet, area)     # 테두리 단계에서 바뀐 한 줄 머리 등 마무리 (대부분 바뀌는 것 없음)
    lay.notes += place(b, sheet, area)
    # 수식이 들어가게 줄을 늘린 뒤(place) 쪽이 넘치면: 줄 높이를 다시 읽어 쪽 나누기·테두리를 다시 한다
    for _ in range(2):
        ws2 = load(b)[sheet]
        hh = {r: (ws2.row_dimensions[r].height if ws2.row_dimensions[r].height is not None else lay.h0.get(r, 15)) for r in lay.h}
        if all(sum(hh[r] for r in range(a, e + 1)) <= PAGE_PT + 0.5 for a, e in pages):
            break
        lay.notes.append('수식 줄을 늘린 뒤 쪽이 넘쳐 쪽 나누기를 다시 함')
        lay.h = dict(hh); lay.h0 = dict(hh)
        pages = lay.paginate(hints)
        lay.write([], docno)
        b.print_area(sheet, area)
        lay.notes += unify(b, sheet, area, [e for _, e in pages[:-1]], old_r2=range_boundaries(area0)[3])
        lay.notes += place(b, sheet, area)
    return lay, pages
