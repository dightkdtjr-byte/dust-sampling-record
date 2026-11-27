# 검증 방식 14: 원본 ↔ 최종 대조 — (가) 원본에서 수식이던 칸이 값·빈칸이 되지 않았는지(따오던 것 유지), 바뀐 수식이 가리키던 칸을
# 그대로 가리키는지(빠진 칸은 의도한 변경 목록과 대조), 없는 시트를 가리키는 수식 (나) 원본의 약속 문자: 표준산소 칸에 "-"·"X"를 적으면 보정 안 함
#   python3 -I tools/labform/link_audit.py <원본 폴더> <기록부 폴더> <작업 폴더>
import sys, os, re, glob, json, collections, shutil, subprocess, warnings
warnings.filterwarnings('ignore'); import openpyxl
here = os.path.dirname(os.path.abspath(__file__)); sys.path.insert(0, here)
from tplkit import Book
SRC, G, W = sys.argv[1:4]
if os.path.exists(W): shutil.rmtree(W)
os.makedirs(W)
KEY = {'불소화합물': '플루오린화합물', '브롬': '브로민화합물', '시안화수소': '사이안화수소', '포름알데하이드': '폼알데하이드'}
def key(p):
    k = re.search(r'시험기록부_([^_]+?)(?:_|\.xls)', os.path.basename(p)).group(1); return KEY.get(k, k)
orig = {key(p): p for p in glob.glob(SRC + '/**/*.xls*', recursive=True)}
gold = {key(p): p for p in glob.glob(G + '/*.xls[xm]')}
ref = re.compile(r"(?:(?:'([^']+)'|([^\s!(),=<>&\"'+*/:^-]+))!)?\$?([A-Z]{1,3})\$?(\d+)(?::\$?([A-Z]{1,3})\$?(\d+))?")
def refs(f):
    f = re.sub(r'"[^"]*"', '', f)
    return {((m.group(1) or m.group(2) or '').strip(), m.group(3) + m.group(4) + (':' + m.group(5) + m.group(6) if m.group(5) else '')) for m in ref.finditer(f)}
# 의도한 변경: 빠져도 되는 칸 (기록부 열쇠, 칸 정규식, 빠진 칸 정규식, 까닭)
OK = [(r'.*', r'.*', r'^[H-K](1[6-9]|2[0-9])$', '방법검출한계 칸 → 기준 시트'),
      (r'VOCs', r'O(53|56|59)', r'^N\d+$', '희석배수는 H열에서 한 번만'),
      (r'VOCs', r'D(32|33|34)', r'^L\d+$', '시료 농도 = 반복 주입 평균(O열, L열 세 칸의 평균)'),
      (r'먼지', r'J18', r'^M3$', '표준산소(L3) 유무로 보정'),
      (r'브로민화합물|플루오린화합물|염화수소', r'F3', r'^B8$', '현장바탕 측정 여부는 B8:B10 개수로'),
      (r'수은화합물', r'R3', r'^[PQ]3$', '표준산소·실측산소 입력 칸(Q15·Q16)을 직접'),
      (r'중금속', r'J22', r'^H25$', '베릴륨이 카드뮴 열을 읽던 것 고침')]
tot = collections.Counter(); bad = []
for k in sorted(orig):
    if k not in gold: bad.append(f'{k}: 최종 파일 없음'); continue
    wo, wg = openpyxl.load_workbook(orig[k]), openpyxl.load_workbook(gold[k])
    so, sg = wo.sheetnames, [s for s in wg.sheetnames if s != '기준']
    if len(so) != len(sg): bad.append(f'{k}: 시트 수 {len(so)} → {len(sg)}'); continue
    rn = {a: b for a, b in zip(so, sg) if a != b}
    old_names = set(so)
    for a, b in zip(so, sg):
        for row in wo[a].iter_rows():
            for c in row:
                v = c.value
                if not (isinstance(v, str) and v.startswith('=')): continue
                tot['원본 수식'] += 1
                w = wg[b][c.coordinate].value
                if not (isinstance(w, str) and w.startswith('=')):
                    bad.append(f'{k} {b}!{c.coordinate}: 수식이 값·빈칸이 됨 ({v[:50]} → {w!r})'); continue
                if w == v: tot['그대로'] += 1; continue
                tot['바뀜'] += 1
                full = lambda sh_, r_: (rn.get(sh_, sh_) if sh_ else b, r_)
                A = {full(*x) for x in refs(v) if not x[0] or x[0] in old_names or x[0] in rn}
                B = {(x[0] or b, x[1]) for x in refs(w)}
                for sh_, r_ in A - B:
                    if any(re.fullmatch(kk, k) and re.fullmatch(cc, c.coordinate) and re.fullmatch(gg, r_) for kk, cc, gg, _ in OK):
                        tot['의도한 변경으로 빠진 칸'] += 1
                    else:
                        bad.append(f'{k} {b}!{c.coordinate}: 가리키던 {sh_}!{r_}가 빠짐 ({v[:60]} → {w[:60]})')
    for ws in wg.worksheets:
        for row in ws.iter_rows():
            for c in row:
                w = c.value
                if isinstance(w, str) and w.startswith('=') and 'INDIRECT' not in w:
                    for sh_, _ in refs(w):
                        if sh_ and sh_ not in wg.sheetnames: bad.append(f'{k} {ws.title}!{c.coordinate}: 없는 시트 {sh_}')
print(f"(가) 원본 수식 {tot['원본 수식']}칸: 그대로 {tot['그대로']}, 바뀜 {tot['바뀜']} (의도한 변경으로 빠진 칸 {tot['의도한 변경으로 빠진 칸']}), 문제 {len(bad)}")
for x in bad[:40]: print('   ✖', x)
# (나) 약속 문자
CASES = [('암모니아', '계산과정', 'O4', '시험기록부(BA)', 'D17'), ('황화수소', '계산과정', 'M3', '시험기록부', 'D17'), ('수은화합물', '계산과정', 'Q15', '시험기록부', 'D21')]
exp = []
for i, (k, sh, cell, fs, res) in enumerate(CASES):
    for j, val in enumerate(('', '-', 'X')):
        d = os.path.join(W, f's{i}{j}'); os.makedirs(d)
        b = Book(gold[k]); b.setv(sh, cell, val if val else None); b.finish(os.path.join(d, os.path.basename(gold[k])))
        exp.append((k, val, d, fs, res))
import lo_pool
lo_pool.run(sorted({e[2] for e in exp}))
got = {}
for k, val, d, fs, res in exp:
    p = glob.glob(os.path.join(d, 'lo', '*.xlsx'))
    got[(k, val)] = str(openpyxl.load_workbook(p[0], data_only=True)[fs][res].value) if p else '파일 없음'
nb = 0
for k, *_ in CASES:
    base = got[(k, '')]
    for val in ('-', 'X'):
        ok = got[(k, val)] == base and base not in ('', 'None') and not base.startswith('#')
        nb += not ok
        print(f"   {'✔' if ok else '✖'} {k} 표준산소 칸 \"{val}\": 결과 {got[(k, val)]!r} (빈칸일 때 {base!r})")
print(f'원본 ↔ 최종 대조: 수식 {tot["원본 수식"]}칸 문제 {len(bad)}, 약속 문자 {len(CASES) * 2}건 중 틀림 {nb}')
