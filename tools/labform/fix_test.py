# 검증 방식 11: 원본 계산 수식 고친 것 확인 — 결과 칸이 아니라 앞 단계 입력 칸에 값을 넣고 LibreOffice로 다시 계산해
# 고친 계산이 맞는 결과를 내는지 본다 (form_fixes.py의 항목마다 실패하던 예를 그대로)
#   python3 -I tools/labform/fix_test.py <완성본 폴더> <작업 폴더>
import sys, os, shutil, warnings
import openpyxl
warnings.filterwarnings('ignore')
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import lo_pool

golden, work = sys.argv[1], sys.argv[2]
approx = lambda want, tol=1e-6: (lambda v: isinstance(v, (int, float)) and abs(v - want) <= tol, f'≈{want}')
eq = lambda want: (lambda v: str(v if v is not None else '').strip() == want, repr(want))
nonempty = (lambda v: v not in (None, ''), '값 있음')

V = ['Benzene', 'Dichloromethane', 'Styrene', 'Chloroform', '1,3-Butadiene']
CASES = [
    ('PAHs 시료 − 바탕', '시험기록부_PAHs.xlsx',
     {'시험기록부': {'I21': 15, 'I22': 15, 'I23': 15, 'I24': 1500, 'I25': 1500, 'I26': 1500, 'I17': '해당없음'}},
     [('시험기록부', 'J21', approx(1485)), ('시험기록부', 'D17', eq('1485.0'))]),
    ('PAHs 기준 빈칸이어도 불검출 판정', '시험기록부_PAHs.xlsx',
     {'시험기록부': {'I21': 0, 'I22': 0, 'I23': 0, 'I24': 1, 'I25': 1, 'I26': 1, 'I17': None}},
     [('시험기록부', 'D17', eq('불검출'))]),
    ('베릴륨이 카드뮴을 읽지 않음', '시험기록부_중금속.xlsm',
     {'계산과정': {'H14': 0.05, 'H9': 0, 'I14': 0.0000001, 'I9': 0}, '베릴륨화합물': {'J17': '해당없음'}, '카드뮴화합물': {'J17': '해당없음'}},
     [('베릴륨화합물', 'D17', eq('불검출')), ('카드뮴화합물', 'D17', (lambda v: v not in ('불검출', None, ''), '카드뮴은 검출'))]),
    ('금속마다 자기 표준산소 (아연 4(12), 납 0.8(6))', '시험기록부_중금속.xlsm',
     {'아연화합물': {'J17': '4(12)'}, '납화합물': {'J17': '0.8(6)'}, '구리화합물': {'J17': '4'}},
     [('계산과정', 'E22', approx(12)), ('계산과정', 'D22', approx(6)), ('계산과정', 'C22', eq(''))]),
    ('금속 요약에 납·카드뮴·베릴륨', '시험기록부_중금속.xlsm', {},
     [('시험기록부-금속화합물', 'A16', eq('납화합물')), ('시험기록부-금속화합물', 'I16', nonempty),
      ('시험기록부-금속화합물', 'I17', nonempty), ('시험기록부-금속화합물', 'I18', nonempty)]),
    ('이황화메틸 ppb 그대로', '시험기록부_이황화메틸.xlsx',
     {'시험기록부-이황화메틸': {'H32': 0, 'H33': 0, 'H34': 0, 'H35': 5, 'H36': 5, 'H37': 5, 'I27': '해당없음'}},
     [('시험기록부-이황화메틸', 'I32', approx(5)), ('시험기록부-이황화메틸', 'D27', eq('5.00'))]),
    ('브로민 표준상태 시료량 /760', '시험기록부_브로민화합물.xlsx',
     {'계산과정': {'H3': 20, 'I3': 0, 'J3': 0, 'K3': 740}},
     [('계산과정', 'N3', approx(20 * 740 / 760, 1e-6))]),
    ('플루오린 표준상태 시료량 /760', '시험기록부_플루오린화합물.xlsx',
     {'계산과정': {'H3': 20, 'I3': 0, 'J3': 0, 'K3': 740}},
     [('계산과정', 'N3', approx(20 * 740 / 760, 1e-6))]),
    ('염화수소 표준상태 시료량 /760', '시험기록부_염화수소.xlsx',
     {'계산과정': {'H3': 20, 'I3': 0, 'J3': 0, 'K3': 740}},
     [('계산과정', 'N3', approx(19.47, 0.006))]),
    # 산소 보정 규칙 (사용자 결정): 표준산소(괄호)가 있을 때만 실측산소로 보정, 없으면 보정 안 한 값, 있는데 실측산소가 비면 "실측산소 입력"
    ('황화수소 괄호 없음 → 보정 안 한 값', '시험기록부_황화수소.xlsx',
     {'계산과정': {'E6': 0.003, 'E3': 0, 'L3': 1, 'N3': 15}, '시험기록부': {'J17': '1.6'}},
     [('시험기록부', 'D17', eq('2.1 (초과)'))]),
    ('황화수소 괄호 있고 실측산소 빈칸', '시험기록부_황화수소.xlsx',
     {'계산과정': {'E6': 0.003, 'E3': 0, 'L3': 1, 'N3': None}, '시험기록부': {'J17': '1.6(12)'}},
     [('시험기록부', 'D17', eq('실측산소 입력'))]),
    ('황화수소 괄호 있고 실측산소 15 %', '시험기록부_황화수소.xlsx',
     {'계산과정': {'E6': 0.003, 'E3': 0, 'L3': 1, 'N3': 15}, '시험기록부': {'J17': '1.6(12)'}},
     [('시험기록부', 'D17', eq('3.1 (초과)'))]),
    ('수은 괄호 없음 → 보정 안 한 값', '시험기록부_수은화합물.xlsx',
     {'계산과정': {'N3': 0.05, 'N6': 0, 'N9': 0, 'N12': 0}, '시험기록부': {'J21': '0.04'}},
     [('시험기록부', 'D21', eq('0.05000 (초과)'))]),
    ('수은 괄호 있고 실측산소 빈칸', '시험기록부_수은화합물.xlsx',
     {'계산과정': {'N3': 0.05, 'N6': 0, 'N9': 0, 'N12': 0, 'Q16': None}, '시험기록부': {'J21': '0.04(12)'}},
     [('시험기록부', 'D21', eq('실측산소 입력'))]),
    ('암모니아 괄호 없음 → 보정 안 한 값 (전에는 빈칸)', '시험기록부_암모니아.xlsx',
     {'계산과정': {'F7': 1, 'F4': 0, 'N4': 25, 'F13': 1, 'F10': 0, 'N10': 25}, '시험기록부 (H2O2)': {'K17': '16'}},
     [('시험기록부 (H2O2)', 'D17', eq('1.0'))]),
    ('먼지 괄호 없음 + 실측산소 적음 → 보정 안 한 값 (전에는 빈칸)', '시험기록부_먼지.xlsx',
     {'계산과정': {'M3': 10}, '시험기록부': {'I14': '해당없음'}},
     [('시험기록부', 'D14', (lambda v: v not in (None, '', '실측산소 입력'), '보정 안 한 값'))]),
    ('아연 괄호 있고 실측산소 빈칸', '시험기록부_중금속.xlsm',
     {'아연화합물': {'J17': '4(12)'}, '계산과정': {'C23': None}},
     [('아연화합물', 'D17', eq('실측산소 입력'))]),
    ('아세트알데하이드 현장바탕을 뺌', '시험기록부_아세트알데하이드.xlsx',
     {'계산과정': {'B8': 1, 'B9': 1, 'B10': 1}},
     [('계산과정', 'D3', approx(1)), ('계산과정', 'D5', approx(1))]),
    ('총탄화수소 평균을 반올림하지 않음', '시험기록부_총탄화수소.xlsx',
     {'계산과정': {'C2': 4.9849, **{f'C{r}': None for r in range(3, 122)}}},
     [('계산과정', 'C122', approx(4.9849, 1e-9))]),
    ('VOCs 주입량 100 mL (희석배수 2 한 번)', '시험기록부_VOCs.xlsm',
     {s: {'M53': 100, 'M56': 100, 'M59': 100, 'L51': 0, 'L52': 0, **{f'L{r}': 30 for r in range(53, 62)}, 'I27': '해당없음'} for s in V},
     [(s, 'I32', approx(0.06, 1e-9)) for s in V] + [(s, 'D27', eq('불검출' if s == 'Dichloromethane' else '0.06')) for s in V]),   # 다이클로로메테인 방법검출한계 0.17
    ('VOCs 요약에 염화바이닐', '시험기록부_VOCs.xlsm', {},
     [('시험기록부-VOCs', 'A37', eq('VInyl chloride')), ('시험기록부-VOCs', 'I37', nonempty)]),
    ('VOCs 앞선 반올림 없음 (0.0249 → 0.024 → 0.02 < 0.03 불검출)', '시험기록부_VOCs.xlsm',
     {'Benzene': {'H32': 24.9, 'H33': 24.9, 'H34': 24.9, 'I27': '해당없음'}},
     [('Benzene', 'D27', eq('불검출'))]),
]

os.makedirs(work, exist_ok=True)
dirs = []
for i, (name, f, sets, _) in enumerate(CASES):
    d = os.path.join(work, f'c{i:02d}'); os.makedirs(d, exist_ok=True)
    wb = openpyxl.load_workbook(os.path.join(golden, f), keep_vba=f.endswith('m'))
    for sh, cells in sets.items():
        for ref, v in cells.items():
            wb[sh][ref].value = v
    wb.save(os.path.join(d, f))
    dirs.append(d)
lo_pool.run(dirs)
ok = bad = 0
for i, (name, f, _, checks) in enumerate(CASES):
    p = os.path.join(work, f'c{i:02d}', 'lo', os.path.splitext(f)[0] + '.xlsx')
    wb = openpyxl.load_workbook(p, data_only=True) if os.path.exists(p) else None
    res = []
    for sh, ref, (test, want) in checks:
        v = wb[sh][ref].value if wb else '(파일 없음)'
        good = bool(wb) and test(v)
        ok += good; bad += (not good)
        res.append(f'{sh}!{ref}={v!r}' + ('' if good else f' ✖ (기대 {want})'))
    print(('✔ ' if all('✖' not in r for r in res) else '✖ ') + name + ': ' + ', '.join(res))
print(f'원본 수식 고친 것 확인: 맞음 {ok}, 틀림 {bad}')
