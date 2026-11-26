# 검증 방식 6: 눈으로 보는 검증 — 인쇄(PDF)에 실제로 찍히는지
#   python3 -I tools/labform/visual_round.py <완성본 폴더> <master.json> <작업 폴더>
#   · 보통 상태와 "초과"를 일부러 만든 상태(허가배출기준 아주 작게, 계산값 = 방법검출한계의 5배)를 PDF로 찍어
#     빨간 픽셀 수를 비교한다 → 늘어나야 조건부 서식이 인쇄에 찍힌 것
#   · PDF 글자에 "(초과)"가 찍히는지, "불검출 (초과)" 같은 잘못된 조합이 없는지
#   · 19개 파일 첫 쪽을 한 장으로 모은 그림 (사람이 보는 용)
import sys, os, re, json, subprocess, glob, warnings
import openpyxl
from PIL import Image
warnings.filterwarnings('ignore')
golden, master, work = sys.argv[1:4]
rows = json.load(open(master, encoding='utf8'))
base_d, over_d = os.path.join(work, 'base'), os.path.join(work, 'over')
os.makedirs(base_d, exist_ok=True); os.makedirs(over_d, exist_ok=True)
LO = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'lo.sh')
for f in sorted({r['file'] for r in rows}):
    src = os.path.join(golden, f)
    import shutil
    shutil.copy(src, os.path.join(base_d, f))          # 보통 상태 = 완성본 그대로
    wb = openpyxl.load_workbook(src, keep_vba=f.endswith('m'))
    for r in [x for x in rows if x['file'] == f]:
        try:
            mdl = float(r['mdl'])
        except (TypeError, ValueError):
            mdl = 0.1
        ws = wb[r['sheet']]
        for res in r['result'].split(', '):
            fm = str(ws[res].value)
            mp = re.search(r'TRIM\(SUBSTITUTE\((\$?[A-Z]{1,3}\$?\d+)&""', fm)
            m = re.search(r"ROUND\((?:ROUND|TRUNC\(ROUND)\(IFERROR\(VALUE\(LEFT\((\$?[A-Z]{1,3}\$?\d+)&\"\"", fm)
            m2 = re.search(r"ROUND\((?:ROUND|TRUNC\(ROUND)\(IFERROR\(VALUE\(LEFT\(IF\(((?:'[^']+'|[^!(),\s]+)!)?(\$?[A-Z]{1,3}\$?\d+)=\"\"", fm)
            if mp:
                ws[mp.group(1).replace('$', '')].value = '0.0000001'
            if m:
                ws[m.group(1).replace('$', '')].value = mdl * 5
            elif m2:
                sh = (m2.group(1) or '').rstrip('!').strip("'") or r['sheet']
                wb[sh][m2.group(2).replace('$', '')].value = mdl * 5
    wb.save(os.path.join(over_d, f))
for d in (base_d, over_d):
    subprocess.run([LO, d, 'pdf'], capture_output=True, timeout=1800)


def red_pixels(pdf, tag):
    out = os.path.join(work, 'png', tag); os.makedirs(out, exist_ok=True)
    subprocess.run(['pdftoppm', '-r', '40', '-png', pdf, os.path.join(out, 'p')], capture_output=True)
    n = 0
    for p in glob.glob(os.path.join(out, 'p*.png')):
        im = Image.open(p).convert('RGB')
        n += sum(1 for (r, g, b) in im.getdata() if r > 170 and g < 90 and b < 90)
    return n


ok = bad = 0
for f in sorted({r['file'] for r in rows}):
    stem = os.path.splitext(f)[0]
    pb, po = os.path.join(base_d, 'pdf', stem + '.pdf'), os.path.join(over_d, 'pdf', stem + '.pdf')
    if not (os.path.exists(pb) and os.path.exists(po)):
        print('✖', f, 'PDF 없음'); bad += 1; continue
    rb, ro = red_pixels(pb, stem + '_b'), red_pixels(po, stem + '_o')
    t = subprocess.run(['pdftotext', po, '-'], capture_output=True, text=True).stdout
    n_over = len(re.findall(r'\(초과\)', t))
    wrong = len(re.findall(r'불검출\s*\(초과\)', t))
    want = sum(len(r['result'].split(', ')) for r in rows if r['file'] == f)
    good = ro > rb and n_over > 0 and wrong == 0
    ok += good; bad += (not good)
    print(f"{'✔' if good else '✖'} {stem[6:]:14} 빨간 픽셀 {rb} → {ro} | PDF의 \"(초과)\" {n_over}개 (결과 칸 {want}) | 잘못된 조합 {wrong}")
print(f'눈으로 보는 검증: 맞음 {ok}, 틀림 {bad}')
# 첫 쪽 모음
ims = []
for f in sorted({r['file'] for r in rows}):
    stem = os.path.splitext(f)[0]
    p = os.path.join(work, 'png', stem + '_o')
    first = sorted(glob.glob(os.path.join(p, 'p*.png')))
    if first:
        ims.append(Image.open(first[0]).convert('RGB'))
if ims:
    w, h = ims[0].size; cols = 5
    sheet = Image.new('RGB', (w * cols, h * ((len(ims) + cols - 1) // cols)), 'white')
    for i, im in enumerate(ims):
        sheet.paste(im.resize((w, h)), ((i % cols) * w, (i // cols) * h))
    sheet.save(os.path.join(work, '첫쪽_초과표시.png'))
