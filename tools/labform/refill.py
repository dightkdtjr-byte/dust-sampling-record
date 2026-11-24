# 정리한 양식에 원본의 예시 입력값을 도로 넣는다 (계산이 원본과 같은지 보려고). 수식 칸은 건드리지 않는다
# 쓰는 법: python3 -I tools/labform/refill.py <원본 폴더> <정리본 폴더> <결과 폴더>
import sys, openpyxl, warnings
warnings.filterwarnings('ignore')
ORIG, TPL, OUT = sys.argv[1], sys.argv[2], sys.argv[3]
PAIRS = [('dust.xlsx', 'labtest-dust.xlsx'), ('gas.xlsx', 'labtest-gas.xlsx'), ('hcn.xlsx', 'labtest-hcn.xlsx'), ('bap.xlsx', 'labtest-bap.xlsx'), ('vocs.xlsm', 'labtest-vocs.xlsm')]
for o, t in PAIRS:
    wo = openpyxl.load_workbook(f'{ORIG}/{o}'); wt = openpyxl.load_workbook(f'{TPL}/{t}')
    n = 0
    for ws in wo.worksheets:
        wt_s = wt[ws.title]
        for row in ws.iter_rows():
            for c in row:
                v = c.value
                if v is None or (isinstance(v, str) and v.startswith('=')): continue
                tc = wt_s[c.coordinate]
                if type(tc).__name__ == 'MergedCell': continue
                if tc.value is None:
                    tc.value = v; n += 1
    # 시료채취일(입력 칸으로 바꾼 곳)은 원본의 계산된 날짜를 넣는다
    wv = openpyxl.load_workbook(f'{ORIG}/{o}', data_only=True)
    main = wo.worksheets[0].title
    for ref in ('I6', 'J6'):
        if wt[main][ref].value is None and wv[main][ref].value is not None and type(wt[main][ref]).__name__ != 'MergedCell':
            wt[main][ref].value = wv[main][ref].value
    wt.save(f'{OUT}/' + t.rsplit('.', 1)[0] + '.xlsx'); print(t, '채운 칸', n)
