# 시험기록부 프로그램(HTML 한 파일)을 만든다.
#   python3 -I tools/labform/build_program.py <기록부 폴더> <master_checked.json> <jszip.min.js> <결과 html> [evidence.json]
# 프로그램 안에 기록부 원본(xlsx/xlsm)·기준표·원문 그림·JSZip이 모두 들어 있어 인터넷 없이 브라우저에서 연다.
# 기준을 고친 뒤 내려받으면 각 기록부의 숨김 '기준' 시트 칸(D·E·F·G·H·I)을 고쳐 준다 (기록부 하나 또는 전체 zip).
import sys, os, json, base64, glob, datetime

VERSION = '2.3.0'
golden, master, jszip, out = sys.argv[1:5]
evidence = sys.argv[5] if len(sys.argv) > 5 else None
rows = json.load(open(master, encoding='utf8'))
keep = ['file', 'row', 'item', 'sheet', 'result', 'method', 'calc', 'final', 'mdl', 'loq', 'unit', 'note', 'linked', 'method_row', 'ref', 'sig_mdl', 'sig_loq']
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from form_fixes import O2_FIX
# 산소 보정을 하는 기록부 (결과 계산에 표준산소·실측산소 칸)
o2 = lambda f: any(k in f for k in O2_FIX)
data = [{k: r.get(k) for k in keep} | {'o2': o2(r['file'])} for r in rows]
files = {}
for p in sorted(glob.glob(os.path.join(golden, '*.xls*'))):
    files[os.path.basename(p)] = base64.b64encode(open(p, 'rb').read()).decode()
# 원문 문장 등 RULES (총칙·정도보증)
RULES = {
    'calc': '총칙 1.6·2.14.2',
    'nd': '총칙 2.14.3',
    'unit': '총칙 2.14.1',
    'sig': '정도보증 4.3.4.4',
    'ratio': '정도보증 4.3.2·4.3.3',
}
js = lambda t: t.replace('</', '<\\/')      # <script> 안 JSON에 "</script>"가 있어도 끊기지 않게
html = open(os.path.join(os.path.dirname(os.path.abspath(__file__)), 'program_template.html'), encoding='utf8').read()
SCENE = os.path.join(os.path.dirname(os.path.abspath(__file__)), '..', 'suite', 'vendor', 'stackpilot_scene.js')
html = (html.replace('/*__JSZIP__*/', open(jszip, encoding='utf8').read())
            .replace('/*__SCENE__*/', open(SCENE, encoding='utf8').read().replace('</script', '<\\/script') if os.path.exists(SCENE) else '')
            .replace('/*__WORDMARK__*/', open(os.path.join(os.path.dirname(SCENE), 'sp_wordmark.js'), encoding='utf8').read().replace('</script', '<\\/script') if os.path.exists(os.path.join(os.path.dirname(SCENE), 'sp_wordmark.js')) else '')
            .replace('"__DATA__"', js(json.dumps(data, ensure_ascii=False, default=str)))
            .replace('"__FILES__"', json.dumps(files))
            .replace('"__RULES__"', js(json.dumps(RULES, ensure_ascii=False)))
            .replace('"__EV__"', js(open(evidence, encoding='utf8').read()) if evidence else '{"codes":{},"rules":{},"img":{}}')
            .replace('__VERSION__', VERSION)
            .replace('__BUILT__', datetime.date.today().isoformat()))
open(out, 'w', encoding='utf8').write(html)
print(out, f'{os.path.getsize(out) / 1e6:.1f} MB', '기준', len(data), '항목', len(files), '파일')
