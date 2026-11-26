# 통합 프로그램(HTML 한 파일)을 만든다: 채팅(작업)마다 따로 만든 프로그램 HTML을 한 파일 안에 넣고
# 처음 화면·위 막대에서 골라 연다. 프로그램끼리 코드를 나누지 않으므로 한쪽을 고쳐도 다른 쪽이 깨지지 않는다.
#   python3 -I tools/suite/build_suite.py <suite.json> <결과 html>
# suite.json: {"title": ..., "kicker": ..., "lead": ..., "apps": [{"id", "title", "kicker", "desc", "color", "file", "version"}]}
#   file이 없거나 비어 있으면 그 카드는 "아직 합치지 않음"으로 나온다.
import sys, os, json, base64, datetime

man_path, out = sys.argv[1], sys.argv[2]
man = json.load(open(man_path, encoding='utf8'))
base = os.path.dirname(os.path.abspath(man_path))
apps, blobs = [], []
for a in man['apps']:
    f = a.get('file') and os.path.join(base, a['file'])
    ready = bool(f and os.path.exists(f))
    apps.append({k: a.get(k) for k in ('id', 'title', 'kicker', 'desc', 'color', 'version')} | {'ready': ready})
    if ready:
        blobs.append(f'<script type="application/octet-stream" id="app-{a["id"]}">{base64.b64encode(open(f, "rb").read()).decode()}</script>')
tpl = open(os.path.join(os.path.dirname(os.path.abspath(__file__)), 'suite_template.html'), encoding='utf8').read()
title = man.get('title', '실험실 프로그램')
html = (tpl.replace('"__MANIFEST__"', json.dumps(apps, ensure_ascii=False))
           .replace('__APPS__', '\n'.join(blobs))
           .replace('__TITLE_SPACED__', ' '.join(title.replace(' ', '')))
           .replace('__TITLE__', title)
           .replace('__KICKER__', man.get('kicker', ''))
           .replace('__LEAD__', man.get('lead', ''))
           .replace('__BUILT__', datetime.date.today().isoformat()))
open(out, 'w', encoding='utf8').write(html)
print(out, f'{os.path.getsize(out) / 1e6:.1f} MB', ', '.join(f'{a["title"]}({"O" if a["ready"] else "없음"})' for a in apps))
