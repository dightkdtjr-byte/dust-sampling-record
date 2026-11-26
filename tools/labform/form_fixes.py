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



NEED_O2 = '실측산소 입력'
# 표준산소(허가배출기준 괄호 안)가 있을 때만 실측산소로 보정. 표준산소가 없으면 보정하지 않은 값, 있는데 실측산소가 비면 "실측산소 입력"
# (사용자 결정. 원본은 실측산소가 비면 0 %로 보고 보정하거나, 괄호가 없으면 결과가 빈칸이었다)
# 파일 이름 조각 → [(시트, 칸, 새 수식)]
O2_FIX = {
    '황화수소': [('계산과정', 'P3', '=IF(M3="","",IF(N3="","{N}",O6*((21-M3)/(21-N3))))'),
                 ('시험기록부', 'L21', '=IF(계산과정!M3="","",IF(ISNUMBER(계산과정!P3),ROUND(TRUNC(ROUND(계산과정!P3,\'기준\'!$E$2+5),\'기준\'!$E$2),\'기준\'!$F$2)&"("&계산과정!M3&")",계산과정!P3))')],   # 보정값 표시 칸: 맺음 규칙 같게, 글자면 그대로
    '수은': [('계산과정', 'R3', '=IF(Q15="",O3,IF(Q16="","{N}",IF(N(O3)>0,O3*((21-Q15)/(21-Q16)),0)))')],
    '암모니아': [('계산과정', 'R7', '=IF(O4="","",IF(P4="","{N}",Q7*(21-O4)/(21-P4)))'),
                 ('계산과정', 'R13', '=IF(O10="","",IF(P10="","{N}",Q13*(21-O10)/(21-P10)))')],
    '플루오린': [('계산과정', 'P3', '=IF(L3="","",IF(M3="","{N}",ROUND((((E3-F3)*100)/N3)*(22.4/18.998)*((21-L3)/(21-M3)),2)&"("&L3&")"))')],
    '불소': [('계산과정', 'P3', '=IF(L3="","",IF(M3="","{N}",ROUND((((E3-F3)*100)/N3)*(22.4/18.998)*((21-L3)/(21-M3)),2)&"("&L3&")"))')],
    '비소': [('계산과정(As)', 'F25', '=IF(E21="","",IF(E22="","{N}",((H13+N13)/$E$23)*((21-$E$21)/(21-$E$22))))')],
    '먼지': [('계산과정', 'N3', '=IF(L$3="","",IF(M3="","{N}",K3*((21-L3)/(21-M3))))'),
             ('시험기록부', 'J18', '=IF(계산과정!L3="",ROUND(계산과정!K3,2),계산과정!N3)')],
    '염화수소': [('계산과정', 'P3', '=IF(L3="","",IF(M3="","{N}",ROUND((((E3-F3)*100)/N3)*(22.4/35.453)*((21-L3)/(21-M3)),2)&"("&L3&")"))')],
    'SOx': [('계산과정', 'I5', '=IF(F5="",D5,IF(G5="","{N}",D5*((21-F5)/(21-G5))))')],
    '중금속': [('계산과정', f'{c}27', f'=IF({c}$22="",{c}25,IF($C$23="","{{N}}",{c}$25*((21-{c}$22)/(21-$C$23))))') for c in 'CDEFGHI'],
}
OTHER_FIX = {
    '아세트알데하이드': [('계산과정', f'D{r}', '=C8', '현장바탕 평균(C8)을 뺌 (원본은 0으로 박혀 있어 빼지 않음. 폼알데하이드와 같게, 사용자 결정)') for r in (3, 4, 5)],
    '총탄화수소': [('계산과정', 'C122', '=AVERAGE(C2:C121)', '평균을 반올림하지 않고 다음 계산에 (맺음은 결과 칸에서 한 번, 사용자 결정)')],
}

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
    base = os.path.basename(path)
    for key, lst in O2_FIX.items():
        if key not in base:
            continue
        for sh, ref, f in lst:
            if sh not in b.sheets:
                continue
            f = f.replace('{N}', NEED_O2)
            if b.wb[sh][ref].value != f:
                b.setf(sh, ref, f[1:]); b.wb[sh][ref].value = f
                log.append(f'{sh}!{ref}: 표준산소가 있을 때만 실측산소로 보정, 없으면 보정 안 한 값, 실측산소가 비면 "{NEED_O2}"')
    for key, lst in OTHER_FIX.items():
        if key not in base:
            continue
        for sh, ref, f, why in lst:
            if sh in b.sheets and b.wb[sh][ref].value != f:
                b.setf(sh, ref, f[1:]); b.wb[sh][ref].value = f
                log.append(f'{sh}!{ref}: {why}')
    return log


def o2_cells(b, path, r_cells):
    """결과가 읽는 계산값 칸이 다시 가리키는 칸(한 단계 더): 예) 염화수소 결과 J21 → 계산과정!P3(보정값).
    거기의 ROUND(…, n)도 계산 자릿수 버림으로 바꾼다 (계산 자릿수 이상에서 버리면 결과는 같고, 그 아래에서 반올림하던 것만 고쳐짐)"""
    cell_re = re.compile(r"(?<![:A-Za-z$])(?:(?:'([^']+)'|([^\s!(),=<>&\"'+*/:-]+))!)?\$?([A-Z]{1,3})\$?(\d+)(?![\d(:])")   # 범위(H32:H34)는 빼고
    out = []
    for sh, ref, item, sheet in r_cells:
        v = b.wb[sh][ref].value
        if not (isinstance(v, str) and v.startswith('=')):
            continue
        for m in cell_re.finditer(re.sub(r'"[^"]*"', '', v)):
            s2 = m.group(1) or m.group(2) or sh
            if s2 in b.sheets:
                out.append((s2, f'{m.group(3)}{m.group(4)}', item, sheet))
    return out


def _trunc_rounds(f, calc=None):
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
            k = max(int(n), calc) if calc is not None else int(n)
            out.append(f[i:m.start()]); out.append(f'TRUNC(ROUND({_trunc_rounds(f[m.end():comma], calc)},{k + 5}),{k})'); i = j
        else:
            out.append(f[i:m.end()]); i = m.end()
    return ''.join(out)


def strip_pre_rounding(b, sheet, ref, calc=None):
    """결과 수식이 읽는 계산값 칸 R의 ROUND(식, n)을 버림(TRUNC)으로: 결과 수식이 계산 자릿수에서 버리고 최종 자릿수에서
    한 번 맺으므로, 앞에서 먼저 반올림하면 맺음이 두 번이 된다(0.0249 → 0.025 → 0.03). 표시 자릿수는 그대로. 바뀌면 True"""
    v = b.wb[sheet][ref].value
    if not (isinstance(v, str) and v.startswith('=')):
        return False
    nv = '=' + _trunc_rounds(v[1:], calc)
    if nv == v:
        return False
    b.setf(sheet, ref, nv[1:])
    b.wb[sheet][ref].value = nv
    return True
