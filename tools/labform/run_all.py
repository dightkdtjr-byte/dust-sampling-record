# 시험기록부 원본 기록(xlsx/xlsm)을 모두 같은 기준으로 양식화한다. 양식마다 코드를 따로 쓰지 않는 공통 실행기.
#   python3 -I tools/labform/run_all.py <원본 폴더(하위 폴더 포함)> <결과 폴더>
# 파일마다:
#   · A1에 "시험기록부"가 있는 보이는 시트 = 기록부 양식 → std.standardize (8가지 기준 + 예시 값)
#     본문 범위: 오른쪽 끝은 서명란 끝 열, 없으면 문서번호 칸의 열, 그것도 없으면 첫 기록부와 같은 열. 아래 끝은 인쇄 범위(없으면 마지막 줄)
#   · 나머지 시트(계산과정·rawdata 등) → 글꼴·바닥글(문서번호)·여백만
#   · 시험기한 수식의 DAY() 버그를 고친다 (기간을 날짜로 적고 DAY()로 꺼내면 31일을 넘을 때 틀린다)
#   · 기본 정보 칸(의뢰기관·측정대상·날짜·시각·기간·서명)은 표준 예시로 바꾼다 (content_rules.OVERWRITE)
#   · 원본을 LibreOffice로 다시 계산해 오류(#REF!·#DIV/0! 등)가 나는 수식은 IFERROR(…,"")로 감싼다 (빈 줄의 INDIRECT, 빈 자료의 AVERAGE 등)
#   · 계산·자료 시트도 날짜 yyyy-mm-dd, 숫자·수식 칸 셀에 맞춤
import sys, os, re, glob, warnings
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
warnings.filterwarnings('ignore')
from openpyxl.utils import get_column_letter as CL, range_boundaries
from tplkit import Book
import content_rules as CR
import subprocess, shutil, tempfile, openpyxl
from std import standardize, page_setup, find_docno, load, is_docno

DEFAULT_DOCNO = '■ MAMO-QP-14-01'
SRC, DST = sys.argv[1], sys.argv[2]
os.makedirs(DST, exist_ok=True)
CR.OVERWRITE = True


def item_name(path):
    n = os.path.splitext(os.path.basename(path))[0]
    m = re.search(r'시험기록부_([^_]+)', n)
    return (m.group(1) if m else n).strip()


def is_form(ws):
    return ws.sheet_state == 'visible' and '시험기록부' in str(ws['A1'].value or '').replace(' ', '')


def body_area(ws, first_c2=None):
    pa = ws.print_area
    pa = pa[0] if isinstance(pa, (list, tuple)) else pa
    if pa:
        c1, r1, c2, r2 = range_boundaries(pa.split('!')[-1].replace('$', ''))
    else:
        c1, r1, c2, r2 = 1, 1, ws.max_column, ws.max_row
    M = {}
    for g in ws.merged_cells.ranges:
        M[(g.min_row, g.min_col)] = g.max_col
    sign = [M.get((1, c), c) for c in range(1, 40) if re.sub(r'\s', '', str(ws.cell(1, c).value or '')) in ('담당', '기술책임자')]
    doc_c = [c.column for row in ws.iter_rows() for c in row if is_docno(c.value)]
    if sign:
        c2 = max(sign)
    elif doc_c:
        c2 = max(doc_c)
    elif first_c2:
        c2 = first_c2
    # 아래 끝: 인쇄 범위가 없으면 내용이 있는 마지막 줄
    if not pa:
        r2 = max((c.row for row in ws.iter_rows() for c in row if c.value is not None and c.column <= c2), default=r2)
    return f'A1:{CL(c2)}{r2}', c2


def fix_due(b, sheet):
    s = b.sx(sheet); n = 0
    def fx(m):
        nonlocal n
        f = m.group(0)
        if 'WORKDAY' not in f:
            return f
        g = re.sub(r'DAY\((\$?[A-Z]+\$?\d+)\s*-\s*1\)', r'(\1-1)', f)
        g = re.sub(r'DAY\((\$?[A-Z]+\$?\d+)\)', r'\1', g)
        if g != f:
            n += 1
        return g
    s = re.sub(r'<f\b[^>]*>[^<]*</f>', fx, s)
    b.setsx(sheet, s)
    return n


ERR = ('#REF!', '#DIV/0!', '#VALUE!', '#N/A', '#NAME?', '#NUM!', '#NULL!')


def error_cells(path):
    """원본을 LibreOffice로 강제 재계산해 오류 값이 나는 칸: {시트: [칸]}"""
    tmp = tempfile.mkdtemp(dir=os.path.join(os.path.dirname(DST.rstrip('/')) or '.', ''))
    ext = os.path.splitext(path)[1]
    shutil.copy(path, os.path.join(tmp, 'x' + ext))
    subprocess.run([os.path.join(os.path.dirname(os.path.abspath(__file__)), 'lo.sh'), tmp], capture_output=True, timeout=400)
    out = {}
    calc = os.path.join(tmp, 'lo', 'x.xlsx')
    if os.path.exists(calc):
        wb = openpyxl.load_workbook(calc, data_only=True)
        for ws in wb.worksheets:
            refs = [c.coordinate for row in ws.iter_rows() for c in row if isinstance(c.value, str) and c.value.strip() in ERR]
            if refs:
                out[ws.title] = refs
    shutil.rmtree(tmp, ignore_errors=True)
    return out


def run(path):
    errs = error_cells(path)
    b = Book(path)
    wb = b.wb
    forms, others = [], []
    first_c2 = None
    for ws in wb.worksheets:
        if is_form(ws):
            area, c2 = body_area(ws, first_c2)
            first_c2 = first_c2 or c2
            forms.append((ws.title, area))
        elif ws.sheet_state == 'visible':
            others.append(ws.title)
    if not forms:
        return None, ['기록부 시트 없음']
    log = []
    for sheet, refs in errs.items():
        if sheet in b.sheets:
            b.unshare(sheet)
            k = CR.wrap_errors(b, sheet, refs)
            if k:
                log.append(f'{sheet}: 오류 나는 수식 {k}곳 IFERROR로 감쌈 ({", ".join(refs[:4])}{" …" if len(refs) > 4 else ""})')
    for sheet, _ in forms:
        b.unshare(sheet)
        k = fix_due(b, sheet)
        if k:
            log.append(f'{sheet}: 시험기한 수식 DAY() 버그 고침 {k}곳')
    doc = find_docno(b, forms[0][0]) or DEFAULT_DOCNO
    first = None
    for sheet, area in forms:
        hints = [brk.id for brk in b.wb[sheet].row_breaks.brk]
        lay, pages = standardize(b, sheet, area, doc, hints, like=first)
        first = first or lay
        warn = [n for n in lay.notes if n.startswith('!!')]
        log.append(f'{sheet} {area} → {len(pages)}쪽' + (f' 경고 {warn}' if warn else ''))
    for sheet in others:
        land = (load(b)[sheet].page_setup.orientation == 'landscape')
        page_setup(b, sheet, doc, landscape=land)
        CR.tidy_other(b, sheet)
    ext = os.path.splitext(path)[1]
    dst = os.path.join(DST, f'시험기록부_{item_name(path)}{ext}')
    b.finish(dst)
    return dst, log


files = sorted(p for p in glob.glob(os.path.join(SRC, '**', '*.xls*'), recursive=True) if not os.path.basename(p).startswith('~$'))
ok = 0
for p in files:
    try:
        dst, log = run(p)
        print('ok' if dst else '--', os.path.basename(p))
        for l in log:
            print('    ', l)
        ok += bool(dst)
    except Exception as e:
        import traceback
        print('!! 실패', os.path.basename(p), repr(e))
        traceback.print_exc(limit=3)
print(f'{ok}/{len(files)} 파일 처리')
