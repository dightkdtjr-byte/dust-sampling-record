# 시험기록부 기준 관리 프로그램(HTML 한 파일)을 만든다.
#   python3 -I tools/labform/build_program.py <골든 카피 폴더> <master_checked.json> <jszip.min.js> <결과 html>
# 프로그램 안에 골든 카피 원본(xlsx/xlsm)·기준표·JSZip이 모두 들어 있어 인터넷 없이 브라우저에서 연다.
# 기준을 고친 뒤 "골든 카피 내려받기"를 누르면 각 기록부의 숨김 '기준' 시트 칸(D·E·F·G·H·I)을 고쳐 zip으로 내려받는다.
import sys, os, json, base64, glob, datetime

VERSION = '1.0.0'
golden, master, jszip, out = sys.argv[1:5]
rows = json.load(open(master, encoding='utf8'))
keep = ['file', 'row', 'item', 'sheet', 'result', 'method', 'calc', 'final', 'mdl', 'loq', 'unit', 'note', 'linked', 'method_row', 'ref', 'lab_mdl', 'lab_loq']
data = [{k: r.get(k) for k in keep} for r in rows]
files = {}
for p in sorted(glob.glob(os.path.join(golden, '*.xls*'))):
    files[os.path.basename(p)] = base64.b64encode(open(p, 'rb').read()).decode()
# 원문 문장 등 RULES (총칙·정도보증)
RULES = {
    'calc': '총칙 ES 01000.b 2.14.2 — 시험성적수치는 마지막 유효숫자의 다음 단위까지 계산하여 KS Q 5002 4.2.2.2 4사5입법으로 기록',
    'nd': '총칙 ES 01000.b 2.14.3 — 방법검출한계 미만의 시험결과 값은 불검출로 표시 (같으면 값)',
    'unit': '총칙 ES 01000.b 2.14.1 — 가스상 ppm(μmol/mol)·ppb(nmol/mol), 입자상 mg/Sm³·μg/Sm³·ng/Sm³',
    'sig': '정도보증 ES 01001.b 4.3.4.4 — 방법검출한계·정량한계는 유효숫자 2개 이하로 표기',
    'ratio': '정도보증 ES 01001.b 4.3.2·4.3.3 — 방법검출한계 = 표준편차 × 3.14, 정량한계 = 표준편차 × 10 (비 ≈ 3.18)',
    'ref_mdl': '각 시험법 6.1 — 측정한 방법검출한계는 시험방법에서 제시한 값 이하이어야 한다 (시험 기초자료). 공정시험기준 칸은 원문 제시값과 같아야 함',
}
html = open(os.path.join(os.path.dirname(os.path.abspath(__file__)), 'program_template.html'), encoding='utf8').read()
html = (html.replace('/*__JSZIP__*/', open(jszip, encoding='utf8').read())
            .replace('"__DATA__"', json.dumps(data, ensure_ascii=False, default=str))
            .replace('"__FILES__"', json.dumps(files))
            .replace('"__RULES__"', json.dumps(RULES, ensure_ascii=False))
            .replace('__VERSION__', VERSION)
            .replace('__BUILT__', datetime.date.today().isoformat()))
open(out, 'w', encoding='utf8').write(html)
print(out, f'{os.path.getsize(out) / 1e6:.1f} MB', len(data), '줄', len(files), '파일')
