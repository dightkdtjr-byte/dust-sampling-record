# 시험기록부의 글꼴·굵기·글자 맞춤·예시 값을 모든 양식에서 같은 기준으로 맞춘다.
#
#  글꼴   · 모든 시트(rawdata·계산과정·물질별 포함)의 모든 칸, 칸 안 일부 글자 서식, 메모, 글상자, 차트 글자를 맑은 고딕으로
#            (수식 글꼴 Cambria Math는 그대로)
#  굵기   · 굵게: 제목, 기본 정보 줄(시험항목·의뢰기관·측정대상·시료인수일·분석일자·시험방법·분석기기)에 적는 값 칸(빈 입력 칸 포함)
#          · 그 밖(항목명, 회색 머리 칸, 표, 설명)은 굵지 않게
#  글자   · 칸에 다 들어가지 않는 글은 띄어쓰기 자리에 줄바꿈을 직접 넣는다
#            (Excel은 한글을 글자 단위로도 줄바꿈해서, 자동 줄바꿈에 맡기면 단어 중간이 잘린다)
#          · 좁고 세로로 긴 칸(병합)에 안 들어가는 글은 90° 돌려 세로로 쓴다 ("Sample", "Field Blank" 등)
#          · 그 밖에 낱말 하나가 칸보다 길면 그 칸의 글자 크기를 줄인다 (MIN_PT까지)
#          · 날짜·숫자·수식 칸은 "셀에 맞춤"을 켠다: 값이 칸보다 길어도 ###이 아니라 작게 찍힌다
#  날짜   · 날짜 칸의 표시 형식을 yyyy-mm-dd로
#  예시   · 기본 정보 입력 칸에 모든 양식 같은 예시 값을 넣는다 (이미 수식·글이 있는 칸은 그대로)
import re, datetime
from openpyxl.utils import get_column_letter as CL, range_boundaries
from tplkit import escape

FONT = '맑은 고딕'
MIN_PT = 8
FIT = 0.92                 # 글자 폭 추정이 틀려도 잘리지 않게 칸 안쪽 너비의 92%까지만 쓴다
LINE = 1.35                # 줄 높이 = 글자 크기 × LINE
DATE_FMT = 'yyyy-mm-dd'
BOLD_ROWS = {'시험항목', '의뢰기관', '측정대상', '시료인수일', '분석일자', '시험방법', '분석기기'}
# 예시 값: 사용자가 준 먼지 원본 기록(A16)의 값을 2027년 기준 날짜로 (인수·채취 3월 31일, 분석 다음 날 4월 1일)
EXAMPLE = {
    '의뢰기관': '수도권대기환경청 대기총량과',
    '측정대상': '대일개발(주) #A16',
    '시료채취일': datetime.date(2027, 3, 31),
    '분석일자': datetime.date(2027, 4, 1),
    '시각': 17,
    '시험검사기간': 2,
}
_norm = lambda s: re.sub(r'\s', '', str(s or ''))


# ───────── 글꼴 ─────────
def unify_fonts(b):
    if getattr(b, '_fonts_done', False):
        return []
    b._fonts_done = True
    notes = []
    st = b.text('xl/styles.xml')
    names = set(re.findall(r'<name val="([^"]+)"/>', st))
    def fix_font(m):
        f = m.group(0)
        f = re.sub(r'<name val="[^"]*"/>', f'<name val="{FONT}"/>', f)
        f = re.sub(r'<scheme val="[^"]*"/>', '', f)
        f = re.sub(r'<family val="\d+"/>', '<family val="3"/>', f)
        if '<charset' in f:
            f = re.sub(r'<charset val="\d+"/>', '<charset val="129"/>', f)
        else:
            f = f.replace(f'<name val="{FONT}"/>', f'<name val="{FONT}"/><charset val="129"/>')
        return f
    st = re.sub(r'<font\b[^>]*>.*?</font>', fix_font, st, flags=re.S)
    b.put('xl/styles.xml', st)
    if names - {FONT}:
        notes.append('칸 글꼴 ' + ', '.join(sorted(names - {FONT})) + f' → {FONT}')
    # 칸 안 일부 글자에 따로 지정된 글꼴(서식 있는 텍스트)과 메모의 글꼴
    rich = set()
    for p in [p for p in b.parts if re.match(r'xl/(sharedStrings|comments\d*|worksheets/sheet\d+)\.xml$', p)]:
        x = b.text(p)
        rich |= set(re.findall(r'<rFont val="([^"]+)"/>', x)) - {FONT}
        y = re.sub(r'<rFont val="[^"]*"/>', f'<rFont val="{FONT}"/>', x)
        y = re.sub(r'<scheme val="[^"]*"/>', '', y) if y != x else y
        if y != x:
            b.put(p, y)
    for p in [p for p in b.parts if re.match(r'xl/drawings/vmlDrawing\d+\.vml$', p)]:
        x = b.text(p)
        y = re.sub(r'(face|font-family)(=|:)"?[^";>]*"?', lambda m: f'{m.group(1)}{m.group(2)}"{FONT}"', x)
        if y != x:
            b.put(p, y)
    if rich:
        notes.append('일부 글자·메모 글꼴 ' + ', '.join(sorted(rich)) + f' → {FONT}')
    n = 0
    for p in [p for p in b.parts if re.match(r'xl/(drawings/drawing|charts/chart)\d+\.xml$', p)]:
        x = b.text(p)
        y = re.sub(r'<a:(latin|ea|cs) typeface="(?!Cambria Math)[^"]*"([^>]*)/>', lambda m: f'<a:{m.group(1)} typeface="{FONT}"/>', x)
        if y != x:
            b.put(p, y); n += 1
    if n:
        notes.append(f'글상자·차트 {n}개 글꼴 → {FONT}')
    return notes


# ───────── 예시 값·날짜 형식 ─────────
def _serial(d):
    return (d - datetime.date(1899, 12, 30)).days


def examples(b, sheet, area, fill=True):
    from std import load
    from form_style import Styles, _set_styles, _cell_style
    ws = load(b)[sheet]
    c1, r1, c2, r2 = range_boundaries(area)
    M = {}
    for g in ws.merged_cells.ranges:
        for r in range(g.min_row, g.max_row + 1):
            for c in range(g.min_col, g.max_col + 1):
                M[(r, c)] = (g.min_col, g.min_row, g.max_col, g.max_row)
    notes = []
    S = Styles(b)
    fmt_cache = {}

    def put(r, c, val, why):
        cell = ws.cell(r, c)
        if type(cell).__name__ == 'MergedCell' or cell.value not in (None, ''):
            return
        ref = f'{CL(c)}{r}'
        if isinstance(val, datetime.date):
            b.setv(sheet, ref, _serial(val))
        else:
            b.setv(sheet, ref, val)
        notes.append(f'{ref} {why} = {val}')

    def right_of(r, c):   # 칸(병합 포함) 바로 오른쪽 칸
        z = M.get((r, c), (c, r, c, r))[2]
        return z + 1

    for r in (range(r1 + 2, r2 + 1) if fill else ()):
        for c in range(c1, c2 + 1):
            v = ws.cell(r, c).value
            k = _norm(v)
            if not k or type(ws.cell(r, c)).__name__ == 'MergedCell':
                continue
            if k in ('의뢰기관', '측정대상', '분석일자', '시료채취일', '시험검사기간'):
                t = right_of(r, c)
                if t <= c2:
                    put(r, t, EXAMPLE[k], k)
            if k == '시' and c > c1:
                put(r, c - 1, EXAMPLE['시각'], '시료인수 시각')
    # 날짜 칸 표시 형식
    s_xml = b.sx(sheet)
    new = {}
    for r in range(r1, r2 + 1):
        for c in range(c1, c2 + 1):
            cell = ws.cell(r, c)
            f = cell.number_format or ''
            if type(cell).__name__ == 'MergedCell' or f == DATE_FMT:
                continue
            if re.search(r'(^|[^"])(yy|mm-dd|m/d|d/m)', f) and 'h' not in f.lower():
                st = _cell_style(s_xml, f'{CL(c)}{r}', None, {}, c)
                if st not in fmt_cache:
                    fmt_cache[st] = b.style_with_format(st, DATE_FMT)
                new[(r, c)] = fmt_cache[st]
    if new:
        _set_styles(b, sheet, new)
        notes.append(f'날짜 칸 {len(new)}개 표시 형식 → {DATE_FMT}')
    return notes


# ───────── 굵기 ─────────
def bold_rule(b, sheet, area):
    from std import load
    from form_style import Styles, _set_styles, _cell_style
    import layout_rules as LR
    ws = load(b)[sheet]
    c1, r1, c2, r2 = range_boundaries(area)
    secs = LR._sections(ws, c1, r1, r2)
    name = {}
    for a, e, n in secs:
        for r in range(a, e + 1):
            name[r] = n
    M = {}
    for g in ws.merged_cells.ranges:
        M[(g.min_row, g.min_col)] = (g.min_col, g.min_row, g.max_col, g.max_row)
    S = Styles(b); s_xml = b.sx(sheet)
    new = {}
    for r in range(r1, r2 + 1):
        lab_end = M.get((r, c1 + 1), (c1 + 1, r, c1 + 1, r))[2]
        for c in range(c1, c2 + 1):
            cell = ws.cell(r, c)
            if type(cell).__name__ == 'MergedCell':
                continue
            if r == r1 and c == c1:
                want = True                                  # 제목
            elif r < r1 + 2:
                want = False                                 # 서명란
            elif name.get(r) in BOLD_ROWS and c > lab_end and not LR._gray(cell) \
                    and not (c > lab_end + 1 and isinstance(cell.value, str) and len(_norm(cell.value)) <= 3 and not cell.value.startswith('=')):
                want = True                                  # 기본 정보 값 칸 ("시" 같은 단위 칸 빼고)
            else:
                want = False
            if cell.value is None and not want:
                continue
            if bool(cell.font.b) != want:
                st = _cell_style(s_xml, f'{CL(c)}{r}', None, {}, c)
                new[(r, c)] = S.with_font(st, bold=want)
    if new:
        _set_styles(b, sheet, new)
        return [f'굵기 {len(new)}칸 맞춤 (제목·기본 정보 값만 굵게)']
    return []


# ───────── 글자 맞춤 ─────────
def split_lines(text, inner, size, bold, measure):
    """띄어쓰기 자리에서만 나눠 inner(px)에 들어가게. 낱말 하나가 넘치면 None"""
    out = []
    for para in text.split('\n'):
        words = para.split(' ')
        # 띄어쓰기 없이 긴 영문 이름(d12-Benzo(a)pyrene 등)은 "-" 뒤, "(" 앞에서도 나눌 수 있게
        if any(measure(w, size, bold) > inner for w in words):
            words = [p for w in words for p in (re.findall(r'[^-(]*-|\(?[^-(]+', w) if measure(w, size, bold) > inner else [w])]
            return _join_parts(words, inner, size, bold, measure, out, text)
        cur = ''
        for w in words:
            cand = (cur + ' ' + w) if cur else w
            if measure(cand, size, bold) <= inner:
                cur = cand
            else:
                if measure(w, size, bold) > inner:
                    return None
                if cur:
                    out.append(cur)
                cur = w
        out.append(cur)
    return out


def _join_parts(parts, inner, size, bold, measure, out, text):
    # 쪼갠 조각을 이어 붙이되, 원래 띄어쓰기가 있던 자리만 공백을 넣는다
    rest = text.split('\n')[len(out)] if len(out) < len(text.split('\n')) else ''
    pos = 0; cur = ''; lines = list(out)
    for p in parts:
        i = rest.find(p, pos)
        sep = ' ' if i > pos and rest[pos:i].strip() == '' and pos > 0 else ''
        pos = i + len(p) if i >= 0 else pos
        cand = cur + (sep if cur else '') + p
        if measure(cand, size, bold) <= inner:
            cur = cand
        else:
            if measure(p, size, bold) > inner:
                return None
            if cur:
                lines.append(cur)
            cur = p
    lines.append(cur)
    # 남은 문단은 원래 방식으로
    remain = '\n'.join(text.split('\n')[len(out) + 1:])
    if remain:
        more = split_lines(remain, inner, size, bold, measure)
        if more is None:
            return None
        lines += more
    return lines


def fit_cells(lay):
    """Layout의 열 너비·줄 높이로 글자가 칸에 들어가게 한다. 줄바꿈을 켜야 하는 칸 목록을 돌려준다."""
    from std import px, text_px
    from form_style import Styles, _set_styles, _cell_style
    b, sheet, ws = lay.b, lay.sheet, lay.ws
    S = Styles(b)
    measure = lambda t, sz, bold: text_px(t, sz, bold, False)
    wraps, size_fix, shrink, rotate = [], {}, {}, {}
    s_xml = b.sx(sheet)
    for r in range(lay.r1, lay.r2 + 1):
        for c in range(lay.c1, lay.c2 + 1):
            cell = ws.cell(r, c)
            v = cell.value
            if type(cell).__name__ == 'MergedCell' or v is None:
                continue
            a, top, z, bot = lay.box(r, c)
            z = min(z, lay.c2)
            inner = (sum(px(lay.w[k]) for k in range(a, z + 1)) - 6) * FIT
            avail = sum(lay.h[k] for k in range(top, bot + 1))
            ref = f'{CL(c)}{r}'
            if not isinstance(v, str) or v.startswith('='):
                # 날짜·숫자·수식: 셀에 맞춤 (### 대신 작게)
                # 단, 글자 결과가 옆 빈칸으로 넘쳐 보일 수 있는 왼쪽 정렬 수식 칸(표시 형식이 일반)은 줄이지 않는다
                text_spill = (isinstance(v, str) and (cell.number_format or 'General') == 'General' and a == z and top == bot
                              and c < lay.c2 and cell.alignment.horizontal in (None, 'general', 'left')
                              and ws.cell(r, c + 1).value is None and (r, c + 1) not in lay.merge)
                if text_spill and cell.alignment.shrink_to_fit:
                    lay.notes.append(f'!! {ref} 넘쳐 보여야 할 칸에 셀에 맞춤이 켜져 있음')
                if not text_spill and not cell.alignment.shrink_to_fit and not cell.alignment.wrap_text:
                    shrink[(r, c)] = True
                continue
            if not v.strip():
                continue
            size = cell.font.sz or 11; bold = bool(cell.font.b)
            hpx = avail * 96 / 72
            if cell.alignment.text_rotation in (90, 180, 255):
                one = ' '.join(v.split())
                if measure(one, size, bold) > (hpx - 6) * FIT:
                    lay.notes.append(f'!! {ref} "{v[:12]}" 세로로 써도 칸 높이보다 김')
                continue
            spill = (a == z and top == bot and c < lay.c2 and cell.alignment.horizontal in (None, 'general', 'left')
                     and not cell.alignment.wrap_text and ws.cell(r, c + 1).value is None and (r, c + 1) not in lay.merge)
            lines = v.split('\n')
            if all(measure(t, size, bold) <= inner for t in lines):
                need = len(lines)
            elif spill and len(lines) == 1:
                continue
            else:
                new_size = size
                got = split_lines(v, inner, new_size, bold, measure)
                one = ' '.join(v.split())
                if got is None and top < bot and hpx > inner * 2 and measure(one, size, bold) <= (hpx - 6) * FIT:
                    rotate[(r, c)] = True
                    if one != v:
                        b.setv(sheet, ref, one)
                    lay.notes.append(f'{ref} "{one[:12]}" 좁고 긴 칸이라 90° 돌려 세로로')
                    continue
                while got is None and new_size > MIN_PT:
                    new_size -= 0.5
                    got = split_lines(v, inner, new_size, bold, measure)
                if got is None:
                    lay.notes.append(f'!! {ref} "{v[:12]}" {MIN_PT}pt로도 안 들어감')
                    continue
                if new_size != size:
                    size_fix[(r, c)] = new_size
                    lay.notes.append(f'{ref} "{v[:12].strip()}" 글자 {size:g}→{new_size:g}pt')
                    size = new_size
                newv = '\n'.join(got)
                if newv != v:
                    b.setv(sheet, ref, newv)
                    lay.notes.append(f'{ref} "{v[:12].strip()}" 띄어쓰기 자리에서 {len(got)}줄로')
                need = len(got)
            if need > 1 and not cell.alignment.wrap_text:
                wraps.append(ref)
            want = need * size * LINE + 4
            if want > avail + 0.5:
                lay.h[bot] = round(lay.h[bot] + want - avail, 2)
                lay.notes.append(f'{ref} {need}줄이 들어가게 {bot}행 +{want - avail:.1f}pt')
    s_xml = b.sx(sheet)
    new = {}
    for (r, c), sz in size_fix.items():
        new[(r, c)] = S.with_font(_cell_style(s_xml, f'{CL(c)}{r}', None, {}, c), size=sz)
    for (r, c) in rotate:
        new[(r, c)] = S.with_rotation(_cell_style(s_xml, f'{CL(c)}{r}', None, {}, c), 90)
    for (r, c) in shrink:
        new[(r, c)] = S.with_shrink(new.get((r, c)) or _cell_style(s_xml, f'{CL(c)}{r}', None, {}, c))
    if new:
        _set_styles(b, sheet, new)
    if shrink:
        lay.notes.append(f'날짜·숫자·수식 칸 {len(shrink)}개 "셀에 맞춤" 켬')
    return wraps
