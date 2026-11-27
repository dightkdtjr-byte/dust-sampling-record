# 검증 방식 13: 엑셀에서 나는 오류 찾기 — LibreOffice를 엑셀처럼(빈 글자 "" 계산은 #VALUE!) 두고
# (가) 그대로, (나) 빈 양식(계산과정·raw data 시트의 입력 숫자를 모두 지움)으로 다시 계산해 기록부 시트에 오류가 있는지 본다
#   python3 -I tools/labform/blank_scan.py <기록부 폴더> <작업 폴더>
import sys, os, re, shutil, glob, subprocess, warnings
warnings.filterwarnings('ignore')
here = os.path.dirname(os.path.abspath(__file__)); sys.path.insert(0, here)
import openpyxl
from tplkit import Book
golden, work = sys.argv[1], sys.argv[2]
if os.path.exists(work): shutil.rmtree(work)
A, B = os.path.join(work, 'as_is'), os.path.join(work, 'blank')
os.makedirs(A); os.makedirs(B)
ERR = ('#REF!', '#DIV/0!', '#VALUE!', '#N/A', '#NAME?', '#NUM!', '#NULL!')
norm = lambda v: re.sub(r'\s+', '', str(v or ''))
for p in sorted(glob.glob(golden + '/*.xls[xm]')):
    shutil.copy(p, A)
    wb = openpyxl.load_workbook(p); b = Book(p); n = 0
    for ws in wb.worksheets:
        if ws.title == '기준' or '시험기록부' in norm(ws['A1'].value):
            continue
        for row in ws.iter_rows():
            for c in row:
                if isinstance(c.value, (int, float)) and not isinstance(c.value, bool):
                    b.setv(ws.title, c.coordinate, None); n += 1
    b.finish(os.path.join(B, os.path.basename(p)))
subprocess.run(['python3', '-I', os.path.join(here, 'lo_pool.py'), A, B], stdout=subprocess.DEVNULL, timeout=1800) if os.path.exists(os.path.join(here, 'lo_pool.py')) and False else None
for d in (A, B):
    subprocess.run(['bash', os.path.join(here, 'lo.sh'), d], stdout=subprocess.DEVNULL, timeout=900)
tot = {}
for d, name in ((A, '그대로'), (B, '빈 양식')):
    k = 0
    for p in sorted(glob.glob(os.path.join(d, 'lo', '*.xlsx'))):
        wb = openpyxl.load_workbook(p, data_only=True)
        for ws in wb.worksheets:
            if name == '빈 양식' and '시험기록부' not in norm(ws['A1'].value):
                continue   # 빈 양식: 기록부 시트만 (계산과정은 입력이 없으면 오류가 날 수 있음)
            for row in ws.iter_rows():
                for c in row:
                    if isinstance(c.value, str) and c.value.strip() in ERR:
                        k += 1
                        if k <= 30: print(f'   ✖ {name} {os.path.basename(p)[6:]} {ws.title}!{c.coordinate} {c.value}')
    tot[name] = k
print(f"엑셀 오류 찾기: 그대로 {tot['그대로']}칸, 빈 양식 기록부 시트 {tot['빈 양식']}칸")
