# 정리한 시험기록부 양식이 6가지 기준을 지키는지 차례로 검사한다.
#   python3 -I verify.py <정리본 폴더> [--ref <원본 폴더>] [--pdf <PDF 폴더>]
#   ① 서명란 크기·인쇄 배율이 모든 양식에서 같은가
#   ② 쪽이 항목 경계에서만 나뉘고, 병합 칸·항목이 쪽 사이로 잘리지 않으며, 한 쪽 높이를 넘지 않는가
#   ③ 테두리가 원본과 같은가 (--ref), 표 밖으로 남은 선이 없는가
#   ④ 글자가 칸을 넘거나 잘리지 않는가 (글자 폭은 추정값)
#   ⑤ 정렬·글꼴이 원본과 같은가 (--ref)
#   ⑥ 문서번호가 인쇄되는 모든 시트의 바닥글에 있고, PDF 모든 쪽에 찍히는가 (--pdf)
# 원본과 정리본은 파일 이름 앞부분(확장자·'labtest-' 뺀 것)으로 짝짓는다.
import sys, os, re, glob, subprocess, warnings, zipfile
import openpyxl
from openpyxl.utils import get_column_letter as CL, range_boundaries
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from form_style import raw_borders
from std import SIGN_W, TOTAL_PX, ROW1, ROW2, MARGINS, PAGE_PT, DOCNO, px, text_px, wrap_lines

warnings.filterwarnings('ignore')
SIGN_WORDS = ('담당', '기술책임자', '담당자', '품질책임자', '검토', '승인', '작성')

args = sys.argv[1:]
def opt(name):
    if name in args:
        i = args.index(name); v = args[i + 1]; del args[i:i + 2]; return v
ref_dir, pdf_dir = opt('--ref'), opt('--pdf')
out_dir = args[0]

key = lambda p: re.sub(r'^labtest-', '', os.path.splitext(os.path.basename(p))[0])
files = sorted(glob.glob(os.path.join(out_dir, '*.xls[xm]')))
refs = {key(p): p for p in glob.glob(os.path.join(ref_dir, '*.xls[xm]'))} if ref_dir else {}
problems = {}


def bad(form, crit, msg):
    problems.setdefault(form, {}).setdefault(crit, []).append(msg)


def area_of(ws):
    pa = ws.print_area
    if not pa:
        return 1, 1, ws.max_column, ws.max_row
    pa = pa[0] if isinstance(pa, (list, tuple)) else pa
    return range_boundaries(pa.split('!')[-1].replace('$', ''))


def col_w(ws, c):
    d = ws.column_dimensions.get(CL(c))
    # openpyxl은 min~max로 묶인 열을 첫 열에만 둔다
    for k, v in ws.column_dimensions.items():
        if v.min and v.max and v.min <= c <= v.max and v.width:
            return v.width
    return (d.width if d and d.width else None) or ws.sheet_format.defaultColWidth or 8.43


def row_h(ws, r):
    h = ws.row_dimensions[r].height
    return h if h is not None else (ws.sheet_format.defaultRowHeight or 15)


def merges_of(ws):
    m = {}
    for g in ws.merged_cells.ranges:
        for r in range(g.min_row, g.max_row + 1):
            for c in range(g.min_col, g.max_col + 1):
                m[(r, c)] = (g.min_col, g.min_row, g.max_col, g.max_row)
    return m


def breaks_of(ws):
    return sorted(b.id for b in ws.row_breaks.brk)


def first_sheet_borders(path):
    z = zipfile.ZipFile(path)
    wbx = z.read('xl/workbook.xml').decode('utf8'); rels = z.read('xl/_rels/workbook.xml.rels').decode('utf8')
    rid = re.search(r'<sheet\b[^>]*r:id="([^"]+)"', wbx).group(1)
    tgt = re.search(rf'<Relationship\b[^>]*Id="{rid}"[^>]*Target="([^"]+)"', rels) or re.search(rf'<Relationship\b[^>]*Target="([^"]+)"[^>]*Id="{rid}"', rels)
    part = 'xl/' + re.sub(r'^/?xl/', '', tgt.group(1))
    return raw_borders(z.read(part).decode('utf8'), z.read('xl/styles.xml').decode('utf8'))


class _Side:
    def __init__(self, st): self.style = st


class _Bd:
    def __init__(self, t): self.left, self.right, self.top, self.bottom = (_Side(x) for x in t)


def footer_docno(ws):
    t = (ws.oddFooter.right.text or '') + (ws.oddFooter.center.text or '') + (ws.oddFooter.left.text or '')
    m = DOCNO.search(t)
    return m.group(0) if m else None


summary = []
for f in files:
    form = key(f)
    wb = openpyxl.load_workbook(f)
    ws = wb.worksheets[0]
    c1, r1, c2, r2 = area_of(ws)
    M = merges_of(ws)
    RB = first_sheet_borders(f)
    BD = lambda r, c: _Bd(RB(r, c))

    # ① 서명란·배율
    sign = []
    for c in range(c1, c2 + 1):
        v = ws.cell(r1, c).value
        if isinstance(v, str) and re.sub(r'\s', '', v) in SIGN_WORDS:
            a, _, z, _ = M.get((r1, c), (c, r1, c, r1))
            sign.append(sum(px(col_w(ws, k)) for k in range(a, z + 1)))
    tot = sum(px(col_w(ws, c)) for c in range(c1, c2 + 1))
    ps, pm = ws.page_setup, ws.page_margins
    fit = ws.sheet_properties.pageSetUpPr.fitToPage if ws.sheet_properties.pageSetUpPr else False
    row = dict(form=form, sign=sign, h1=row_h(ws, r1), h2=row_h(ws, r1 + 1), total=tot,
               fit=f'{ps.fitToWidth}x{ps.fitToHeight}' if fit else f'scale {ps.scale}',
               margin=(pm.left, pm.right, pm.top, pm.bottom), paper=ps.paperSize, orient=ps.orientation)
    summary.append(row)
    if len(sign) < 2:
        bad(form, '①', '서명란(담당·기술책임자)을 찾지 못함')
    for w in sign:
        if w != px(SIGN_W):
            bad(form, '①', f'서명 칸 너비 {w}px (기준 {px(SIGN_W)}px)')
    if abs(row['h1'] - ROW1) > .01 or abs(row['h2'] - ROW2) > .01:
        bad(form, '①', f'서명란 줄 높이 {row["h1"]}/{row["h2"]}pt (기준 {ROW1}/{ROW2})')
    if tot != TOTAL_PX:
        bad(form, '①', f'인쇄 열 너비 합 {tot}px (기준 {TOTAL_PX}px) → 배율이 달라짐')
    if not fit or ps.fitToWidth != 1 or (ps.fitToHeight or 0) != 0:
        bad(form, '①', f'배율 설정 {row["fit"]} (기준: 한 쪽 너비에 맞춤 1x0)')
    if (pm.left, pm.right, pm.top, pm.bottom) != (MARGINS['left'], MARGINS['right'], MARGINS['top'], MARGINS['bottom']):
        bad(form, '①', f'여백 {row["margin"]}')

    # ② 쪽 나눔
    starts = {r for r in range(r1, r2 + 1) if re.fullmatch(r'\d{1,2}', str(ws.cell(r, c1).value or '').strip())}
    cuts = [r1 - 1] + [b for b in breaks_of(ws) if r1 <= b < r2] + [r2]
    for b in cuts[1:-1]:
        if (b + 1) not in starts:
            bad(form, '②', f'{b}행 뒤에서 쪽이 나뉨: 다음 줄({b + 1}행)이 항목 시작이 아님')
        for (a, top, z, bot) in set(M.values()):
            if top <= b < bot:
                bad(form, '②', f'병합 칸 {CL(a)}{top}:{CL(z)}{bot} 가 쪽 사이로 잘림')
    for a, b in zip(cuts, cuts[1:]):
        h = sum(row_h(ws, r) for r in range(a + 1, b + 1))
        if h > PAGE_PT + 0.5:
            bad(form, '②', f'{a + 1}~{b}행 높이 {h:.1f}pt > 한 쪽 {PAGE_PT:.1f}pt → 엑셀이 중간에서 자동으로 나눔')

    # ③ 테두리: 표 밖(인쇄 범위 아래·오른쪽)에 남은 선, 내용 없이 선만 있는 맨 아래 줄
    for r in range(r2 + 1, min(ws.max_row, r2 + 20) + 1):
        for c in range(c1, c2 + 1):
            b = BD(r, c)
            if any(getattr(b, s).style for s in ('left', 'right', 'top', 'bottom')):
                bad(form, '③', f'인쇄 범위 밖 {CL(c)}{r}에 테두리가 남음'); break
    # 표 아래쪽 내용 없는 줄에 남은 가로선: 맨 아래 테두리 한 줄 말고는 없어야 한다
    def hline(r):  # r행 아래 경계에 선이 그어진 열
        return {c for c in range(c1, c2 + 1)
                if BD(r, c).bottom.style or (r < r2 and BD(r + 1, c).top.style)}
    last_content = max((r for r in range(r1, r2 + 1) if any(ws.cell(r, c).value is not None for c in range(c1, c2 + 1))), default=r1)
    lines = [(r, hline(r)) for r in range(last_content, r2 + 1)]
    lines = [(r, cols) for r, cols in lines if cols]
    full = [r for r, cols in lines if len(cols) == c2 - c1 + 1]
    part = [(r, cols) for r, cols in lines if len(cols) < c2 - c1 + 1]
    if len(full) > 2:   # 마지막 항목 머리줄 아래 선 + 맨 아래 선
        gap = [r for r in full[1:-1]]
        bad(form, '③', f'표 아래 빈 줄에 가로선이 겹침: {", ".join(f"{r}행 아래" for r in gap)} (맨 아래 선은 {full[-1]}행)')
    for r, cols in part:
        if r > last_content:
            bad(form, '③', f'{r}행 아래에 일부 열({CL(min(cols))}~{CL(max(cols))})에만 가로선이 남음')
    # 인쇄 범위 바깥 오른쪽 열에 선만 있는 줄
    # ④ 글자 넘침·잘림 (값이 있는 글자 칸만; 수식 결과는 빼고 본다)
    for r in range(r1, r2 + 1):
        for c in range(c1, c2 + 1):
            cell = ws.cell(r, c)
            v = cell.value
            if type(cell).__name__ == 'MergedCell' or not isinstance(v, str) or v.startswith('=') or not v.strip():
                continue
            a, top, z, bot = M.get((r, c), (c, r, c, r))
            z = min(z, c2)
            size = cell.font.sz or 11; bold = bool(cell.font.b)
            dotum = (cell.font.name or '') in ('돋움', '굴림', '바탕', '돋움체', '굴림체')
            inner = sum(px(col_w(ws, k)) for k in range(a, z + 1)) - 6
            avail = sum(row_h(ws, k) for k in range(top, bot + 1))
            lh = size * 1.3
            if cell.alignment.wrap_text:
                n = wrap_lines(v, inner, size, bold, dotum)
                if n * lh > avail + 1:
                    bad(form, '④', f'{CL(c)}{r} "{v[:16].strip()}" {n}줄 필요({n * lh:.0f}pt) > 칸 높이 {avail:.0f}pt')
            else:
                lines = v.split('\n')
                w = max(text_px(t, size, bold, dotum) for t in lines)
                spill = (a == z and top == bot and c < c2 and cell.alignment.horizontal in (None, 'general', 'left')
                         and ws.cell(r, c + 1).value is None and (r, c + 1) not in M)
                if w > inner and not spill:
                    bad(form, '④', f'{CL(c)}{r} "{v[:16].strip()}" 글자 폭 {w:.0f}px > 칸 {inner}px')
                if len(lines) > 1 and len(lines) * lh > avail + 1:
                    bad(form, '④', f'{CL(c)}{r} "{v[:16].strip()}" {len(lines)}줄 > 칸 높이 {avail:.0f}pt')

    # ③ 테두리 굵기 규칙: 바깥·쪽 경계 = 굵은 선, 안쪽 = 가는 선만
    brk = set(b for b in breaks_of(ws) if r1 <= b < r2)
    wrong = []
    for r in range(r1, r2 + 1):
        for c in range(c1, c2 + 1):
            a, top, z, bot = M.get((r, c), (c, r, c, r))
            bd = BD(r, c)
            want = {'left': 'medium' if c == c1 else None, 'right': 'medium' if c == c2 else None,
                    'top': 'medium' if (r == r1 or (r - 1) in brk) else None,
                    'bottom': 'medium' if (r == r2 or r in brk) else None}
            for side, w in want.items():
                st = getattr(bd, side).style
                edge = (side == 'left' and c == a) or (side == 'right' and c == z) or (side == 'top' and r == top) or (side == 'bottom' and r == bot)
                if not edge:
                    continue
                if w and st != w:
                    wrong.append(f'{CL(c)}{r} {side} {st}(굵은 선이어야 함)')
                elif not w and st not in (None, 'thin'):
                    wrong.append(f'{CL(c)}{r} {side} {st}(가는 선이어야 함)')
    if wrong:
        bad(form, '③', f'테두리 굵기가 규칙과 다른 곳 {len(wrong)}군데: ' + ', '.join(wrong[:5]))
    # ⑤ 서식 통일: 제목 글꼴, 서명란 글자, 한 줄 머리 정렬
    t = ws.cell(r1, c1)
    row['title'] = (t.font.name, t.font.sz, bool(t.font.b))
    row['signtext'] = tuple(str(ws.cell(r1, c).value) for c in range(c1, c2 + 1) if isinstance(ws.cell(r1, c).value, str) and re.sub(r'\s', '', ws.cell(r1, c).value) in SIGN_WORDS)
    for r in range(r1 + 2, r2 + 1):
        if re.fullmatch(r'\d{1,2}', str(ws.cell(r, c1).value or '').strip()):
            a, top, z, bot = M.get((r, c1 + 1), (c1 + 1, r, c1 + 1, r))
            lab = ws.cell(top, a)
            if z == c2 and top == bot and isinstance(lab.value, str) and lab.alignment.horizontal != 'left':
                bad(form, '⑤', f'{CL(a)}{r} "{lab.value[:12]}" 한 줄 머리가 왼쪽 정렬이 아님')

    # ③·⑤ 원본과 비교
    if form in refs:
        rb = openpyxl.load_workbook(refs[form]); rw = rb.worksheets[0]; RR = first_sheet_borders(refs[form])
        nb = nf = na = 0; ex = {'b': [], 'f': [], 'a': []}
        for r in range(r1, r2 + 1):
            for c in range(c1, c2 + 1):
                x, y = ws.cell(r, c), rw.cell(r, c)
                bx = tuple(getattr(BD(r, c), s).style for s in ('left', 'right', 'top', 'bottom'))
                by = tuple(getattr(_Bd(RR(r, c)), s).style for s in ('left', 'right', 'top', 'bottom'))
                if bx != by:
                    nb += 1; ex['b'].append(f'{CL(c)}{r} {by}→{bx}')
                fx, fy = (x.font.name, x.font.sz, x.font.b, x.font.i), (y.font.name, y.font.sz, y.font.b, y.font.i)
                if fx != fy:
                    nf += 1; ex['f'].append(f'{CL(c)}{r} {fy}→{fx}')
                ax, ay = (x.alignment.horizontal, x.alignment.vertical), (y.alignment.horizontal, y.alignment.vertical)
                if ax != ay:
                    na += 1; ex['a'].append(f'{CL(c)}{r} {ay}→{ax}')
        if nb: bad(form, '③', f'원본과 테두리가 다른 칸 {nb}개: ' + ', '.join(ex['b'][:6]))
        if nf: bad(form, '⑤', f'원본과 글꼴이 다른 칸 {nf}개: ' + ', '.join(ex['f'][:6]))
        if na: bad(form, '⑤', f'원본과 정렬이 다른 칸 {na}개: ' + ', '.join(ex['a'][:6]))

    # ⑥ 문서번호
    for s in wb.worksheets:
        if s.sheet_state != 'visible':
            continue
        if not footer_docno(s):
            bad(form, '⑥', f'"{s.title}" 시트 바닥글에 문서번호 없음')
    if pdf_dir:
        cand = [p for p in glob.glob(os.path.join(pdf_dir, '*.pdf')) if key(p) == form]
        if not cand:
            bad(form, '⑥', 'PDF 없음')
        else:
            n = int(re.search(r'Pages:\s+(\d+)', subprocess.run(['pdfinfo', cand[0]], capture_output=True, text=True).stdout).group(1))
            for i in range(1, n + 1):
                t = subprocess.run(['pdftotext', '-f', str(i), '-l', str(i), cand[0], '-'], capture_output=True, text=True).stdout
                if not DOCNO.search(t):
                    bad(form, '⑥', f'PDF {i}쪽에 문서번호가 안 찍힘')

print('양식        서명칸(px)     서명줄(pt)    열합(px) 배율      여백')
for s in summary:
    print(f'{s["form"]:<10} {str(s["sign"]):<14} {s["h1"]}/{s["h2"]:<8} {s["total"]:<7} {s["fit"]:<8} {s["margin"]}')
same = len({(tuple(s['sign']), s['h1'], s['h2'], s['total'], s['fit'], s['margin']) for s in summary}) == 1
for k, label in (('title', '제목 글꼴'), ('signtext', '서명란 글자')):
    vals = {s[k] for s in summary}
    if len(vals) > 1:
        for s_ in summary:
            bad(s_['form'], '⑤', f'{label}가 양식마다 다름: {s_[k]}')
print('→ 서명란·배율', '모두 같음' if same else '양식마다 다름')
print()
for crit in '①②③④⑤⑥':
    hits = {f: p[crit] for f, p in problems.items() if crit in p}
    print(f'{crit} ' + ('통과' if not hits else f'문제 {sum(len(v) for v in hits.values())}건'))
    for f, msgs in hits.items():
        for m in msgs:
            print(f'    {f}: {m}')
sys.exit(1 if problems else 0)
