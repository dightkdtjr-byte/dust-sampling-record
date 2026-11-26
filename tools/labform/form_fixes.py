# 원본 기록부의 계산 수식 오류 고치기 (결과 수식 앞 단계). 독립 수식 감사(LibreOffice에 값을 넣어 재현)로 찾고,
# 원본 파일·공정시험기준 원문으로 하나씩 확인한 것만 고친다. 각 항목에 근거를 적는다.
import re

VOCS = ['Benzene', '1,3-Butadiene', 'Dichloromethane', 'Trichloroethylene', 'Tetrachloroethylene', '1,2-Dichloroethane',
        'Chloroform', 'Acrylonitrile', 'Styrene', 'Ethylbenzene', 'Carbon tetrachloride', 'VInyl chloride']

# 파일 이름 조각 → [(시트, 칸, 새 값(=으로 시작하면 수식), 까닭)]
FIXES = {
    'PAHs': [
        ('시험기록부', 'J21', '=AVERAGE(I24:I26)-AVERAGE(I21:I23)',
         '시료 − 현장바탕으로 (원본은 21~23행 현장바탕(FB) − 24~26행 시료로 거꾸로 빼서 시료가 늘 음수·불검출)'),
    ],
    '중금속': [
        ('베릴륨화합물', 'J22', '=IF((E22-G22)*100/I22<0,0,(E22-G22)*100/I22)',
         '베릴륨 농도를 자기 시트 값으로 (원본은 계산과정 H25 = 카드뮴 열을 읽음. 다른 6종과 같은 식)'),
        ('베릴륨화합물', 'D20', None, '머리글 "Cd" → "Be"'),
        ('시험기록부-금속화합물', 'A16', '납화합물', '요약에 빠진 납 (수식은 있고 물질 이름만 비어 있었음)'),
        ('시험기록부-금속화합물', 'A17', '카드뮴화합물', '요약에 빠진 카드뮴'),
        ('시험기록부-금속화합물', 'A18', '베릴륨화합물', '요약에 빠진 베릴륨'),
    ],
    '이황화메틸': [
        (None, 'I32', None, '최종 농도를 ppb 그대로 (원본은 이황화탄소 양식의 ppb→ppm 환산 /1000이 남아 결과가 1000배 작음. 머리글·기준·방법검출한계는 ppb)'),
        (None, 'J27', 'ppb', '허가배출기준 단위를 결과 단위(ppb)와 같게 (원본 ppm이면 초과 비교가 1000배 어긋남)'),
    ],
    '브롬': [('계산과정', 'N3', None, '표준상태 시료채취량 분모를 760 mmHg로 (ES 01502.3 식 4: (Pa+Pm)/760. 원본은 대기압으로 나눔)')],
    '브로민': [('계산과정', 'N3', None, '표준상태 시료채취량 분모를 760 mmHg로 (ES 01502.3 식 4)')],
    '불소': [('계산과정', 'N3', None, '표준상태 시료채취량 분모를 760 mmHg로 (ES 01311.2 식: (Pa+Pm)/760)')],
    '플루오린': [('계산과정', 'N3', None, '표준상태 시료채취량 분모를 760 mmHg로 (ES 01311.2)')],
    '염화수소': [('계산과정', 'N3', None, '표준상태 시료채취량 분모를 760 mmHg로 (ES 01305.1e 식: (Pa+Pm)/760)')],
    'VOCs': [('시험기록부-VOCs', 'A37', 'VInyl chloride', '요약에 빠진 염화바이닐 (수식은 있고 물질 이름만 비어 있었음)')],
}


def _vocs_fixes():
    """VOCs 12종을 1,3-Butadiene 식으로 통일: 희석배수 200/주입량을 한 번만(H열), 주입 반복은 평균, 음수는 숫자 0"""
    out = []
    for sh in VOCS:
        for k, (r, lr) in enumerate([(32, 53), (33, 56), (34, 59)]):
            out += [(sh, f'N{lr}', f'=200/M{lr}', '희석배수 = 200 / 주입량 (Benzene은 원본이 거꾸로 주입량/200)'),
                    (sh, f'O{lr}', f'=IFERROR(AVERAGE(L{lr}:L{lr + 2}),0)', '반복 주입 평균, 희석배수는 여기서 곱하지 않음 (원본 4종은 여기와 H열에서 두 번 곱함)'),
                    (sh, f'D{r}', f'=O{lr}', '시료 농도 = 반복 주입 평균 (원본 6종은 첫 주입값만)'),
                    (sh, f'H{r}', f'=ROUND(((D{r}*F{r})-G$32),3)', '음수도 숫자로 평균에 넣음 (Dichloromethane은 원본이 글자 "0.000"으로 바꿔 평균에서 빠짐)')]
    return out


def apply(b, path):
    """b: tplkit.Book (물질 이름을 바꾼 뒤). 고친 내용 목록"""
    import os
    base = os.path.basename(path)
    todo = [x for k, v in FIXES.items() if k in base for x in v]
    if 'VOCs' in base:
        todo += _vocs_fixes()
    first = next((ws.title for ws in b.wb.worksheets if '시험기록부' in str(ws['A1'].value or '').replace(' ', '')), None)
    log = []
    for sheet, ref, val, why in todo:
        sheet = sheet or first
        if sheet not in b.sheets:
            continue
        cur = b.wb[sheet][ref].value
        new = val
        if ref == 'N3' and isinstance(cur, str):
            new = re.sub(r'\(\(K3\+J3\)/K3\)', '((K3+J3)/760)', cur)
        elif ref == 'I32' and '/1000' in str(cur) and '이황화메틸' in base:
            new = str(cur).replace('/1000', '')
        elif ref == 'D20' and isinstance(cur, str):
            new = cur.replace('Cd', 'Be')
        if new is None or new == cur:
            continue
        if isinstance(new, str) and new.startswith('='):
            b.setf(sheet, ref, new[1:])
        else:
            b.setv(sheet, ref, new)
        b.wb[sheet][ref].value = new
        log.append(f'{sheet}!{ref}: {why}')
    return log


METALS = [('C', '구리화합물'), ('D', '납화합물'), ('E', '아연화합물'), ('F', '크로뮴화합물'), ('G', '니켈화합물'), ('H', '카드뮴화합물'), ('I', '베릴륨화합물')]


def post_apply(b, path):
    """결과 수식·표준산소 연결을 마친 뒤 손질. 중금속: 계산과정 22행(표준산소)을 금속마다 자기 시트의 표준산소(K17)로
    (계산과정 열 순서 = Raw data Cu·Pb·Zn·Cr·Ni·Cd·Be). 전에는 구리 시트 기준만 C·D열에 들어가 나머지 금속은 보정이 안 되거나 구리 기준으로 됨"""
    import os
    log = []
    if '중금속' in os.path.basename(path) and '계산과정' in b.sheets:
        for col, sh in METALS:
            if sh not in b.sheets:
                continue
            f = f"=IF('{sh}'!K17=\"\",\"\",'{sh}'!K17)"
            if b.wb['계산과정'][f'{col}22'].value != f:
                b.setf('계산과정', f'{col}22', f[1:]); b.wb['계산과정'][f'{col}22'].value = f
                log.append(f'계산과정!{col}22: 표준산소 = {sh} 기준의 괄호 안 숫자 (전에는 구리 기준만 C·D열에)')
    return log


def _trunc_rounds(f):
    """수식 글의 ROUND(식, n)(n은 숫자)을 TRUNC(ROUND(식, n+5), n)으로: n자리까지 계산(아래 버림). 괄호 짝을 센다"""
    out, i = [], 0
    pat = re.compile(r'(?<![A-Z_.])ROUND\(')
    while True:
        m = pat.search(f, i)
        if not m:
            out.append(f[i:]); break
        j, depth, comma = m.end(), 1, None
        while j < len(f) and depth:
            ch = f[j]
            if ch == '"':
                j = f.index('"', j + 1)
            elif ch == '(':
                depth += 1
            elif ch == ')':
                depth -= 1
            elif ch == ',' and depth == 1:
                comma = j
            j += 1
        n = f[comma + 1:j - 1].strip() if comma else ''
        if re.fullmatch(r'-?\d+', n) and not f[max(0, m.start() - 6):m.start()].endswith('TRUNC('):
            out.append(f[i:m.start()]); out.append(f'TRUNC(ROUND({_trunc_rounds(f[m.end():comma])},{int(n) + 5}),{n})'); i = j
        else:
            out.append(f[i:m.end()]); i = m.end()
    return ''.join(out)


def strip_pre_rounding(b, sheet, ref):
    """결과 수식이 읽는 계산값 칸 R의 ROUND(식, n)을 버림(TRUNC)으로: 결과 수식이 계산 자릿수에서 버리고 최종 자릿수에서
    한 번 맺으므로, 앞에서 먼저 반올림하면 맺음이 두 번이 된다(0.0249 → 0.025 → 0.03). 표시 자릿수는 그대로. 바뀌면 True"""
    v = b.wb[sheet][ref].value
    if not (isinstance(v, str) and v.startswith('=')):
        return False
    nv = '=' + _trunc_rounds(v[1:])
    if nv == v:
        return False
    b.setf(sheet, ref, nv[1:])
    b.wb[sheet][ref].value = nv
    return True
