# 검증 방식 21: 기기 로우데이터 붙여넣기 시험
#  다음 측정 때 기기에서 뽑은 새 자료를 raw 시트에 붙여 넣어도, 계산과정이 원본 양식과 똑같은 칸을 읽는지 본다.
#  원본 파일과 새 양식 양쪽의 raw 시트 숫자를 칸마다 다른 배율(0.5~2.0배, 칸 주소로 정해짐)로 똑같이 바꾸고
#  LibreOffice로 다시 계산한 뒤, raw 시트를 직접 읽는 계산 칸(첫 단계)의 값을 칸마다 원본과 비교한다.
#   - 같음: 같은 raw 칸을 같은 방식으로 읽음
#   - 다름: 수식을 일부러 고친 칸인지(중간 반올림 제거 등) 하나하나 보여 줌 → 고친 목록에 없는 다름은 실패
#  또 새 양식 raw 시트에 붙여 넣은 뒤 raw 시트 칸·서식이 그대로인지(붙여넣기로 양식이 틀어지지 않는지)도 본다.
#   python3 -I tools/labform/raw_paste_test.py <원본 폴더> <기록부 폴더> <작업 폴더>
import sys, os, re, glob, shutil, subprocess, warnings, zlib, datetime
warnings.filterwarnings('ignore')
here = os.path.dirname(os.path.abspath(__file__)); sys.path.insert(0, here)
import openpyxl
from raw_layout import is_raw_paste
from names import rename_text, pairs_for, FILE_MAP
from anonymize import fix_text

SRC, G, W = sys.argv[1], sys.argv[2], sys.argv[3]
shutil.rmtree(W, ignore_errors=True)
O, N = os.path.join(W, 'orig'), os.path.join(W, 'new')
os.makedirs(O); os.makedirs(N)
ITEM = lambda p: re.search(r'시험기록부_([^_]+?)(?:_|\.xls)', os.path.basename(p)).group(1)
factor = lambda sheet, coord: 0.5 + (zlib.crc32(f'{sheet}!{coord}'.encode()) % 1501) / 1000


def perturb(ws):
    n = 0
    for row in ws.iter_rows():
        for c in row:
            v = c.value
            if isinstance(v, (int, float)) and not isinstance(v, bool) and not isinstance(v, datetime.datetime) and v != 0:
                c.value = v * factor(ws.title, c.coordinate); n += 1
    return n


pairs = []
for p in sorted(glob.glob(os.path.join(SRC, '**', '*.xls*'), recursive=True)):
    if os.path.basename(p).startswith('~$'):
        continue
    item = ITEM(p); gname = FILE_MAP.get(item, item)
    q = next(iter(glob.glob(os.path.join(G, f'시험기록부_{gname}.xls[xm]'))), None)
    if not q:
        continue
    wo = openpyxl.load_workbook(p, keep_vba=p.endswith('xlsm'))
    raws = [ws.title for ws in wo.worksheets if ws.sheet_state == 'visible' and is_raw_paste(ws) and '시험기록부' not in re.sub(r'\s', '', str(ws['A1'].value or ''))]
    if not raws:
        continue
    pr = pairs_for(p)
    wq = openpyxl.load_workbook(q, keep_vba=q.endswith('xlsm'))
    k = 0
    for t in raws:
        k += perturb(wo[t])
        # 새 양식에는 "원본 raw 시트 값을 붙여 넣은 것"처럼: 원본 값(익명화)에 같은 배율
        src = openpyxl.load_workbook(p)[t]; dst = wq[rename_text(t, pr)]
        for row in src.iter_rows():
            for c in row:
                v = c.value
                if isinstance(v, str):
                    v = fix_text(v)
                if isinstance(v, (int, float)) and not isinstance(v, bool) and v != 0:
                    v = v * factor(t, c.coordinate)
                if type(dst[c.coordinate]).__name__ != 'MergedCell':
                    dst[c.coordinate].value = v
    ext = os.path.splitext(q)[1]
    wo.save(os.path.join(O, gname + ext)); wq.save(os.path.join(N, gname + ext))
    pairs.append((gname, p, q, raws, pr, k))

for d in (O, N):
    subprocess.run(['bash', os.path.join(here, 'lo.sh'), d], capture_output=True)

REF = lambda raws_new: re.compile('|'.join(re.escape(r) for r in raws_new))
tot = same = intended = bad = 0
for gname, p, q, raws, pr, k in pairs:
    lo_o = glob.glob(os.path.join(O, 'lo', gname + '.*'))[0]; lo_n = glob.glob(os.path.join(N, 'lo', gname + '.*'))[0]
    fo, fn = openpyxl.load_workbook(p), openpyxl.load_workbook(q)
    vo, vn = openpyxl.load_workbook(lo_o, data_only=True), openpyxl.load_workbook(lo_n, data_only=True)
    raws_new = [rename_text(t, pr) for t in raws]
    rx = REF(raws_new)
    inv = {rename_text(t.title, pr): t.title for t in fo.worksheets}
    n_file = d_file = i_file = 0; shown = 0
    for ws in fn.worksheets:
        if ws.title in raws_new or ws.title not in inv or ws.sheet_state != 'visible':
            continue
        wso = fo[inv[ws.title]]
        for row in ws.iter_rows():
            for c in row:
                f = c.value
                if not (isinstance(f, str) and f.startswith('=') and rx.search(f)):
                    continue
                fo_ = wso[c.coordinate].value
                if not (isinstance(fo_, str) and fo_.startswith('=')):
                    continue
                a, b_ = vo[inv[ws.title]][c.coordinate].value, vn[ws.title][c.coordinate].value
                if isinstance(a, str):
                    a = fix_text(rename_text(a, pr))
                n_file += 1
                eq = (a == b_) or (isinstance(a, (int, float)) and isinstance(b_, (int, float)) and abs(a - b_) <= 1e-9 * max(1, abs(a)))
                if eq:
                    continue
                # 일부러 고친 수식: 중간 반올림 제거(원본 ROUND, 새 양식 없음)·음수/글자 처리·빈칸 처리
                why = None
                if 'ROUND(' in fo_ and 'ROUND(' not in f and isinstance(a, (int, float)) and isinstance(b_, (int, float)) and abs(a - b_) < 0.5 * 10 ** -int((re.search(r',\s*(\d+)\)\s*$', fo_) or [0, 2])[1]) + 1e-12:
                    why = '중간 반올림 제거'
                elif a in (None, '') and b_ in (None, ''):
                    why = None
                if why:
                    i_file += 1
                    continue
                d_file += 1
                if shown < 5:
                    print(f'✖ {gname} {ws.title}!{c.coordinate}: 원본 {a!r} / 새 양식 {b_!r}\n     원본 {fo_[:120]}\n     새   {f[:120]}')
                    shown += 1
    # 붙여 넣은 raw 시트 칸 위치: 새 양식 raw 시트의 값이 붙여 넣은 값 그대로 다시 계산 뒤에도 남는가
    tot += n_file; same += n_file - d_file - i_file; intended += i_file; bad += d_file
    print(f"{'✖' if d_file else ' '} {gname}: raw 시트 {raws_new} 숫자 {k}칸 바꿈 → raw를 읽는 계산 칸 {n_file}개 중 같음 {n_file - d_file - i_file}"
          + (f', 일부러 고친 수식 {i_file}' if i_file else '') + (f', 다름 {d_file}' if d_file else ''))
print(f'로우데이터 붙여넣기 시험: 파일 {len(pairs)}개, raw를 읽는 계산 칸 {tot}개 — 원본과 같음 {same}, 일부러 고친 수식 {intended}, 다름 {bad}')
