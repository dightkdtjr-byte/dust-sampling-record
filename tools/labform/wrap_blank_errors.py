# 빈 양식 오류 막기: 계산과정·raw data의 입력 숫자를 모두 지운 사본을 엑셀처럼(빈 글자 계산은 #VALUE!) 다시 계산해
# 오류가 나는 수식 칸을 IFERROR(원래 수식,"")로 감싼다 — 값을 다 넣으면 계산은 그대로, 입력 전·일부만 넣었을 때 #DIV/0!·#VALUE! 대신 빈칸
#   python3 -I tools/labform/wrap_blank_errors.py <기록부 폴더>   (제자리에서 고침, 두 번 돌려도 같음)
import sys, os, re, shutil, glob, subprocess, tempfile, warnings
warnings.filterwarnings('ignore')
here = os.path.dirname(os.path.abspath(__file__)); sys.path.insert(0, here)
import openpyxl
from tplkit import Book
golden = sys.argv[1]
ERR = ('#REF!', '#DIV/0!', '#VALUE!', '#N/A', '#NUM!', '#NULL!')
norm = lambda v: re.sub(r'\s+', '', str(v or ''))


def blank_copy(p, dst):
    wb = openpyxl.load_workbook(p); b = Book(p)
    for ws in wb.worksheets:
        if ws.title == '기준' or '시험기록부' in norm(ws['A1'].value):
            continue
        for row in ws.iter_rows():
            for c in row:
                if isinstance(c.value, (int, float)) and not isinstance(c.value, bool):
                    b.setv(ws.title, c.coordinate, None)
    b.finish(dst)


total = 0
for rnd in range(3):          # 감싼 뒤 다시 계산해 남은 오류가 없을 때까지 (보통 한 번)
    tmp = tempfile.mkdtemp(dir=os.path.dirname(os.path.abspath(golden)))
    files = sorted(glob.glob(golden + '/*.xls[xm]'))
    for p in files:
        blank_copy(p, os.path.join(tmp, os.path.basename(p)))
    subprocess.run(['bash', os.path.join(here, 'lo.sh'), tmp], stdout=subprocess.DEVNULL, timeout=900)
    n_round = 0
    for p in files:
        lo = os.path.join(tmp, 'lo', os.path.splitext(os.path.basename(p))[0] + '.xlsx')
        if not os.path.exists(lo):
            continue
        wv = openpyxl.load_workbook(lo, data_only=True)
        bad = {}
        for ws in wv.worksheets:
            if ws.title == '기준':
                continue
            for row in ws.iter_rows():
                for c in row:
                    if isinstance(c.value, str) and c.value.strip() in ERR:
                        bad.setdefault(ws.title, []).append(c.coordinate)
        if not bad:
            continue
        b = Book(p); k = 0
        for sh, refs in bad.items():
            if sh not in b.sheets:
                continue
            if 't="shared"' in b.sx(sh):
                b.unshare(sh)
            for ref in refs:
                v = b.wb[sh][ref].value
                if isinstance(v, str) and v.startswith('=') and not v.upper().startswith('=IFERROR('):
                    f = f'IFERROR({v[1:]},"")'
                    b.setf(sh, ref, f); b.wb[sh][ref].value = '=' + f; k += 1
        if k:
            b.finish(p); n_round += k
            print(f'{os.path.basename(p)}: 빈 양식에서 오류 나는 수식 {k}칸 → 빈칸 ({", ".join(f"{s}!{r[0]}" + (f" 외 {len(r) - 1}" if len(r) > 1 else "") for s, r in bad.items())})')
    shutil.rmtree(tmp, ignore_errors=True)
    total += n_round
    if not n_round:
        break
print(f'빈 양식 오류 막기: {total}칸')
