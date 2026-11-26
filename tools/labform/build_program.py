# 시험기록부 프로그램(HTML 한 파일)을 만든다.
#   python3 -I tools/labform/build_program.py <기록부 폴더> <master_checked.json> <jszip.min.js> <결과 html> [evidence.json]
# 프로그램 안에 기록부 원본(xlsx/xlsm)·기준표·원문 그림·JSZip이 모두 들어 있어 인터넷 없이 브라우저에서 연다.
# 기준을 고친 뒤 내려받으면 각 기록부의 숨김 '기준' 시트 칸(D·E·F·G·H·I)을 고쳐 준다 (기록부 하나 또는 전체 zip).
import sys, os, json, base64, glob, datetime

VERSION = '2.2.0'
golden, master, jszip, out = sys.argv[1:5]
evidence = sys.argv[5] if len(sys.argv) > 5 else None
rows = json.load(open(master, encoding='utf8'))
keep = ['file', 'row', 'item', 'sheet', 'result', 'method', 'calc', 'final', 'mdl', 'loq', 'unit', 'note', 'linked', 'method_row', 'ref', 'lab_mdl', 'lab_loq', 'lab_rows']
data = [{k: r.get(k) for k in keep} for r in rows]
files = {}
for p in sorted(glob.glob(os.path.join(golden, '*.xls*'))):
    files[os.path.basename(p)] = base64.b64encode(open(p, 'rb').read()).decode()
# 원문 문장 등 RULES (총칙·정도보증)
RULES = {
    'calc': '총칙 ES 01000.b 1.6·2.14.2 (마지막 유효숫자의 다음 단위까지 계산, KS Q 5002 4.2.2.2 4사5입, 맺음은 한 번에)',
    'nd': '총칙 ES 01000.b 2.14.3 (방법검출한계 미만은 불검출)',
    'unit': '총칙 ES 01000.b 2.14.1 (가스상 ppm·ppb, 입자상 mg/Sm³·μg/Sm³·ng/Sm³)',
    'sig': '정도보증 ES 01001.b 4.3.4.4 (유효숫자 2개 이하)',
    'ratio': '정도보증 ES 01001.b 4.3.2·4.3.3 (방법검출한계 = 3.14s, 정량한계 = 10s)',
    'ref_mdl': '각 시험법 6.1 (측정한 방법검출한계는 제시한 값 이하)',
}
js = lambda t: t.replace('</', '<\\/')      # <script> 안 JSON에 "</script>"가 있어도 끊기지 않게
html = open(os.path.join(os.path.dirname(os.path.abspath(__file__)), 'program_template.html'), encoding='utf8').read()
html = (html.replace('/*__JSZIP__*/', open(jszip, encoding='utf8').read())
            .replace('"__DATA__"', js(json.dumps(data, ensure_ascii=False, default=str)))
            .replace('"__FILES__"', json.dumps(files))
            .replace('"__RULES__"', js(json.dumps(RULES, ensure_ascii=False)))
            .replace('"__EV__"', js(open(evidence, encoding='utf8').read()) if evidence else '{"codes":{},"rules":{},"img":{}}')
            .replace('__VERSION__', VERSION)
            .replace('__BUILT__', datetime.date.today().isoformat()))
open(out, 'w', encoding='utf8').write(html)
print(out, f'{os.path.getsize(out) / 1e6:.1f} MB', '기준', len(data), '항목', len(files), '파일')
