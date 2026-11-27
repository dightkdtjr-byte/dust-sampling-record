# 검증 방식 12: 시험 기초자료 유효숫자 표시 — (가) 조건부 서식 수식(소수 자릿수 = MAX(0,S-1-INT(LOG10(|S개로 맺은 값|))))을
# 파이썬 Decimal 규칙과 무작위 2만 건(경계 0.00995·9.95 포함)으로 견주고, (나) 모든 기록부의 시험 기초자료 칸에 값을 넣어
# LibreOffice로 인쇄(PDF)한 글자가 '기준' 시트 L·M열 개수대로 나오는지 본다
#   python3 -I tools/labform/sig_test.py <기록부 폴더> <작업 폴더>
import sys, os, re, math, random, subprocess, warnings, shutil
from decimal import Decimal, ROUND_HALF_UP
warnings.filterwarnings('ignore')
here = os.path.dirname(os.path.abspath(__file__)); sys.path.insert(0, here)
import openpyxl
from tplkit import Book
import form_fixes as FF

golden, work = sys.argv[1], sys.argv[2]
random.seed(12)


def xround(x, n):          # 엑셀 ROUND (0에서 먼 쪽으로 반올림, 음수 자리 가능)
    q = Decimal(1).scaleb(-n)
    return float(Decimal(repr(x)).quantize(q, rounding=ROUND_HALF_UP)) if n >= 0 else float((Decimal(repr(x)) / q).quantize(Decimal(1), rounding=ROUND_HALF_UP) * q)


def formula_decs(x, S):
    lg = lambda v: math.floor(math.log10(abs(v)) + 1e-9)
    return max(0, S - 1 - lg(xround(x, S - 1 - lg(x))))


def rule_text(x, S):       # 유효숫자 S개: S개로 맺은 값의 자리로 소수 자릿수, 표시는 원래 값을 그 자리로 4사5입
    d = Decimal(repr(x)); e = d.adjusted()
    r = d.quantize(Decimal(1).scaleb(e - S + 1), rounding=ROUND_HALF_UP)
    dec = max(0, S - 1 - r.adjusted())
    return dec, str(d.quantize(Decimal(1).scaleb(-dec), rounding=ROUND_HALF_UP))


bad = 0; n = 0
cases = [0.00995, 0.0995, 9.95, 0.995, 0.000995, 99.5, 0.00994, 9.94, 1, 10, 0.1, 0.001]
cases += [round(random.uniform(1, 9.999), random.randint(1, 6)) * 10 ** random.randint(-7, 2) for _ in range(20000)]
for x in cases:
    for S in (1, 2, 3, 4):
        n += 1
        if formula_decs(x, S) != rule_text(x, S)[0]:
            bad += 1
            if bad < 6: print('   ✖ 수식', x, S, formula_decs(x, S), rule_text(x, S))
print(f'(가) 유효숫자 수식 ↔ 규칙: {n}건 중 같음 {n - bad}, 다름 {bad}')

# (나) 인쇄
if os.path.exists(work): shutil.rmtree(work)
os.makedirs(work)
MANT = ['7.3816', '4.6271', '2.9183', '8.1547', '3.5729', '6.2938', '5.8412', '1.7364', '9.2715', '3.0849']
expect = []; i = 0
for f in sorted(os.listdir(golden)):
    if not f.endswith(('.xlsx', '.xlsm')): continue
    p = os.path.join(golden, f); wb = openpyxl.load_workbook(p)
    if '기준' not in wb.sheetnames: continue
    G = wb['기준']; b = Book(p); touched = False
    for ws in wb.worksheets:
        if ws.sheet_state != 'visible' or '시험기록부' not in FF._norm(ws['A1'].value): continue
        x = b.sx(ws.title)
        for ref, kind, label in FF.lab_cells(ws):
            m = re.search(rf"<conditionalFormatting sqref=\"{ref}\">.*?N\('기준'!\$([LM])\$(\d+)\)", x)
            if not m: continue
            S = G[f'{m.group(1)}{m.group(2)}'].value
            if not S: continue
            v = float(f'{MANT[i % len(MANT)]}e{-(1 + i % 5)}'); i += 1
            b.setv(ws.title, ref, v); touched = True
            expect.append((f, ws.title, ref, v, int(S), rule_text(v, int(S))[1]))
    if touched: b.finish(os.path.join(work, f))
subprocess.run(['bash', os.path.join(here, 'lo.sh'), work, 'pdf'], stdout=subprocess.DEVNULL, timeout=900)
ok = bad = 0
txt = {}
for f, sh, ref, v, S, t in expect:
    pdf = os.path.join(work, 'pdf', os.path.splitext(f)[0] + '.pdf')
    if pdf not in txt:
        txt[pdf] = subprocess.run(['pdftotext', '-layout', pdf, '-'], capture_output=True, text=True).stdout if os.path.exists(pdf) else ''
    if re.search(rf'(?<![\d.]){re.escape(t)}(?![\d])', txt[pdf]):
        ok += 1
    else:
        bad += 1; print(f'   ✖ {f[6:]} {sh} {ref}: {v} 유효숫자 {S}개 → {t} 안 보임')
print(f'(나) 인쇄 글자: {ok + bad}칸 중 맞음 {ok}, 틀림 {bad}')
nb = sum(1 for x in cases for S in (1, 2, 3, 4) if formula_decs(x, S) != rule_text(x, S)[0])
print(f'유효숫자 표시 시험: 수식 {n}건·인쇄 {ok + bad}칸, 틀림 {nb + bad}')
