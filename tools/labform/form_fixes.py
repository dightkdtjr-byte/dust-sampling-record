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
                    (sh, f'O{lr}', f'=IFERROR(AVERAGE(L{lr}:L{lr + 2}),"")', '반복 주입 평균, 희석배수는 여기서 곱하지 않음 (원본 4종은 여기와 H열에서 두 번 곱함). 주입값이 모두 비면 빈칸(분석 안 한 시료), 0을 적으면 0'),
                    (sh, f'D{r}', f'=O{lr}', '시료 농도 = 반복 주입 평균 (원본 6종은 첫 주입값만)'),
                    (sh, f'H{r}', f'=IF(D{r}="","",(D{r}*F{r})-G$32)', '분석 안 한 시료는 빈칸으로 평균에서 뺌(0을 적은 시료는 0으로 들어감), 음수도 숫자로 (Dichloromethane은 원본이 글자 "0.000"으로 바꿔 평균에서 빠짐)')]
        out.append((sh, 'O51', '=IFERROR(AVERAGE(L51:L52),0)', '바탕시료 평균을 미리 반올림하지 않음 (맺음은 결과 칸에서 한 번)'))
        out += [(sh, f'P{lr}', f'=IF(O{lr}="","",O{lr}-O$51)', '분석 안 한 시료는 빈칸 (빈칸 − 바탕은 엑셀에서 #VALUE!)') for lr in (53, 56, 59)]
        out += [(sh, f'B{r}', f'=IFERROR(LEFT(K{lr},FIND(".",K{lr})-1),K{lr}&"")', '시료명: 파일 이름에 점이 없거나 비면 그대로 (원본은 #VALUE!)') for r, lr in [(32, 53), (33, 56), (34, 59)]]
        out.append((sh, 'I32', '=IFERROR(IF((AVERAGE(H32:H34)/1000)<0,0,TRUNC(ROUND(AVERAGE(H32:H34)/1000,8),3)),"")', '시료 평균은 분석한 시료만 (원본은 빈 시료가 0으로 평균에 들어가 결과가 작아짐), 시료가 모두 비면 결과 빈칸'))
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
            m_ = re.fullmatch(r'=ROUND\((.*),\s*\d\)', new)      # 염화수소: 시료량을 미리 반올림하지 않음
            if m_:
                new = '=' + m_.group(1)
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
    '황화수소': [('계산과정', 'P3', '=IF(OR(M3="",M3="-",M3="X"),"",IF(N3="","{N}",O6*((21-M3)/(21-N3))))'),
                 ('시험기록부', 'L21', '=IF(OR(계산과정!M3="",계산과정!M3="-",계산과정!M3="X"),"",IF(ISNUMBER(계산과정!P3),ROUND(TRUNC(ROUND(계산과정!P3,\'기준\'!$E$2+5),\'기준\'!$E$2),\'기준\'!$F$2)&"("&계산과정!M3&")",계산과정!P3))')],   # 보정값 표시 칸: 맺음 규칙 같게, 글자면 그대로
    '수은': [('계산과정', 'R3', '=IF(O3="","",IF(OR(Q15="",Q15="-",Q15="X"),O3,IF(Q16="","{N}",IF(N(O3)>0,O3*((21-Q15)/(21-Q16)),0))))')],
    '암모니아': [('계산과정', 'R7', '=IF(OR(O4="",O4="-",O4="X"),"",IF(P4="","{N}",Q7*(21-O4)/(21-P4)))'),
                 ('계산과정', 'R13', '=IF(OR(O10="",O10="-",O10="X"),"",IF(P10="","{N}",Q13*(21-O10)/(21-P10)))')],
    '플루오린': [('계산과정', 'P3', '=IF(OR(L3="",L3="-",L3="X"),"",IF(M3="","{N}",ROUND((((E3-F3)*100)/N3)*(22.4/18.998)*((21-L3)/(21-M3)),2)&"("&L3&")"))')],
    '불소': [('계산과정', 'P3', '=IF(OR(L3="",L3="-",L3="X"),"",IF(M3="","{N}",ROUND((((E3-F3)*100)/N3)*(22.4/18.998)*((21-L3)/(21-M3)),2)&"("&L3&")"))')],
    '비소': [('계산과정(As)', 'F25', '=IF(OR(E21="",E21="-",E21="X"),"",IF(E22="","{N}",((H13+N13)/$E$23)*((21-$E$21)/(21-$E$22))))')],
    '먼지': [('계산과정', 'N3', '=IF(OR(L$3="",L$3="-",L$3="X"),"",IF(M3="","{N}",K3*((21-L3)/(21-M3))))'),
             ('시험기록부', 'J18', '=IF(OR(계산과정!L3="",계산과정!L3="-",계산과정!L3="X"),ROUND(계산과정!K3,2),계산과정!N3)')],
    '염화수소': [('계산과정', 'P3', '=IF(OR(L3="",L3="-",L3="X"),"",IF(M3="","{N}",ROUND((((E3-F3)*100)/N3)*(22.4/35.453)*((21-L3)/(21-M3)),2)&"("&L3&")"))')],
    'SOx': [('계산과정', f'I{r}', f'=IF(D{r}="","",IF(OR(F{r}="",F{r}="-",F{r}="X"),D{r},IF(G{r}="","{{N}}",D{r}*((21-F{r})/(21-G{r})))))') for r in range(5, 14)],
    '중금속': [('계산과정', f'{c}27', f'=IF(OR({c}$22="",{c}$22="-",{c}$22="X"),{c}25,IF($C$23="","{{N}}",{c}$25*((21-{c}$22)/(21-$C$23))))') for c in 'CDEFGHI'],
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
    # 암모니아 (전체): 원본은 보정값이 없는지를 MAX(K21:L32)="" 로 봤는데 MAX는 빈칸만 있어도 0이라 늘 거짓 → 표준산소가 없으면
    # 결과가 0(불검출)이 됨. 보정값(숫자)이 하나도 없으면 보정 전 값의 큰 값, 실측산소가 비었으면 그 글자를 그대로.
    # 결과 수식은 이미 바뀐 뒤라 시트 XML에서 바꾼다 (수식 안의 > 는 &gt;)
    if '암모니아' in base:
        old = 'IF(MAX(K21:L32)="",MAX(J21:J32),MAX(K21:L32))'
        new = 'IF(COUNTIF(K21:L32,"실측산소 입력")&gt;0,"실측산소 입력",IF(COUNT(K21:L32)=0,MAX(J21:J32),MAX(K21:L32)))'
        for sh in list(b.sheets):
            x = b.sx(sh)
            if old in x:
                b.setsx(sh, x.replace(old, new))
                log.append(f'{sh}: 표준산소가 없으면 보정 전 값의 큰 값 (원본은 MAX(…)="" 비교가 늘 거짓이라 0 → 불검출)')
    log += blank_rows(b, path)
    # 글자가 될 수 있는 칸(빈칸 "" 또는 "실측산소 입력")을 ROUND로 감싼 표시 칸: 숫자일 때만 맺고 아니면 그대로 (엑셀은 ROUND("")가 #VALUE!)
    ref_re = re.compile(r"^=ROUND\(((?:'[^']+'|[^!(),'=]+)!)?\$?([A-Z]{1,3})\$?(\d+),(\d)\)$")
    for ws in b.wb.worksheets:
        if ws.title == '기준':
            continue
        for row in ws.iter_rows():
            for c in row:
                m = ref_re.match(str(c.value or ''))
                if not m:
                    continue
                sh2 = (m.group(1) or '').rstrip('!').strip("'") or ws.title
                if sh2 not in b.sheets:
                    continue
                tv = str(b.wb[sh2][f'{m.group(2)}{m.group(3)}'].value or '')
                if tv.startswith('=') and ('""' in tv or '실측산소' in tv):
                    x = (m.group(1) or '') + f'{m.group(2)}{m.group(3)}'
                    n_ = int(m.group(4))     # 표시는 그 자릿수까지 버림: 결과 칸이 이 칸을 다시 읽어도(암모니아 전체) 반올림이 두 번 되지 않게
                    if _set(b, ws.title, c.coordinate, f'=IF(ISNUMBER({x}),TRUNC(ROUND({x},{n_ + 5}),{n_}),{x}&"")'):
                        log.append(f'{ws.title}!{c.coordinate}: 보정값이 빈칸·「실측산소 입력」이면 그대로 표시')
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


def _unshare(b, sh):
    """공유 수식(한 칸 수식을 여러 칸이 같이 씀)을 칸마다 풀어 둔다. 원본 칸을 바꾸면 나머지 칸 수식이 비는 것을 막음"""
    if 't="shared"' in b.sx(sh):
        b.unshare(sh)


def _wrap_blank(b, sh, ref, cond):
    """칸 수식을 IF(cond="","",원래 수식)으로: cond(입력 칸)가 비면 빈칸 → 평균(AVERAGE)에서 빠진다"""
    v = b.wb[sh][ref].value
    if not (isinstance(v, str) and v.startswith('=')) or v.startswith(f'=IF({cond}="",""'):
        return False
    f = f'=IF({cond}="","",{v[1:]})'
    _unshare(b, sh)
    b.setf(sh, ref, f[1:]); b.wb[sh][ref].value = f
    return True


def _set(b, sh, ref, f):
    if b.wb[sh][ref].value == f:
        return False
    _unshare(b, sh)
    b.setf(sh, ref, f[1:]); b.wb[sh][ref].value = f
    return True


def blank_rows(b, path):
    """분석하지 않은 값(입력 칸이 빔)은 평균에서 뺀다. 0이나 n.a.(검출 안 됨)는 0으로 평균에 넣는다 (사용자 결정).
    엑셀은 빈 칸을 가리키는 수식(=B11)과 빈 칸의 곱(B×C)을 0으로 계산하므로, 입력 칸이 비었는지를 직접 보고 빈칸("")을 돌려준다"""
    import os
    base = os.path.basename(path); log = []
    def note(ok, where, why):
        if ok:
            log.append(f'{where}: {why}')
    W = lambda sh: sh in b.sheets
    if '암모니아' in base and W('계산과정'):
        for r in range(4, 16):
            note(_wrap_blank(b, '계산과정', f'E{r}', f'B{r}'), f'계산과정!E{r}', '측정값 빈칸이면 평균에서 뺌')
    if '황화수소' in base and W('계산과정'):
        for r in range(3, 9):
            note(_wrap_blank(b, '계산과정', f'D{r}', f'B{r}'), f'계산과정!D{r}', '측정값 빈칸이면 평균에서 뺌')
    if ('사이안화수소' in base or '시안화수소' in base) and W('계산과정'):
        for r in range(3, 9):
            note(_wrap_blank(b, '계산과정', f'E{r}', f'B{r}'), f'계산과정!E{r}', '측정값 빈칸이면 평균에서 뺌')
    if any(k in base for k in ('브롬', '브로민', '불소', '플루오린', '염화수소')) and W('계산과정'):
        for r in range(8, 14):
            raw = f'rawdata!H{r + 6}'
            note(_set(b, '계산과정', f'B{r}', f'=IF({raw}="","",IF({raw}="n.a.",0,{raw}))'), f'계산과정!B{r}', '기기값 빈칸이면 빈칸(평균에서 뺌), n.a.는 0')
        note(_set(b, '계산과정', 'E3', '=IF(COUNT(C3:C5)=0,"",(IF(ISNUMBER(C3),C3*D3,0)+IF(ISNUMBER(C4),C4*D4,0)+IF(ISNUMBER(C5),C5*D5,0))/COUNT(C3:C5))'),
             '계산과정!E3', '시료 평균은 측정한 반복만')
        note(_set(b, '계산과정', 'F3', '=IF(COUNT(B8:B10)=0,0,C8)'), '계산과정!F3', '현장바탕을 측정하지 않았으면 0')
        first = next((ws.title for ws in b.wb.worksheets if '시험기록부' in str(ws['A1'].value or '').replace(' ', '')), None)
        if first:   # 브로민·플루오린도 염화수소처럼 (빈 반복 × 시료량은 엑셀에서 #VALUE!)
            note(_set(b, first, 'F21', '=IF(COUNT(D21:D23)=0,"",(IF(ISNUMBER(D21),D21*E21,0)+IF(ISNUMBER(D22),D22*E22,0)+IF(ISNUMBER(D23),D23*E23,0))/COUNT(D21:D23))'),
                 f'{first}!F21', '시료 평균은 측정한 반복만')
    if any(k in base for k in ('아세트알데하이드', '포름알데하이드', '폼알데하이드')) and W('계산과정'):
        for r0 in (8, 11, 14, 17):     # 반복 평균: 모두 비면 빈칸 (AVERAGE가 #DIV/0!)
            note(_set(b, '계산과정', f'C{r0}', f'=IF(COUNT(B{r0}:B{r0 + 2})=0,"",AVERAGE(B{r0}:B{r0 + 2}))'), '계산과정!C8·C11·C14·C17', '반복을 모두 분석하지 않았으면 빈칸')
        for r, b0 in ((3, 11), (4, 14), (5, 17)):
            note(_set(b, '계산과정', f'D{r}', '=IF(COUNT(B8:B10)=0,0,C8)'), f'계산과정!D{r}', '현장바탕 평균 (측정하지 않았으면 0)')
            v = b.wb['계산과정'][f'N{r}'].value
            if isinstance(v, str) and v.startswith('=') and 'COUNT(' not in v:
                note(_set(b, '계산과정', f'N{r}', f'=IF(COUNT(B{b0}:B{b0 + 2})=0,"",{v[1:]})'), f'계산과정!N{r}', '시료를 분석하지 않았으면 빈칸(평균에서 뺌)')
    if '비소' in base and W('계산과정(As)'):
        sh = '계산과정(As)'
        for r in (5, 6, 7, 9, 10, 11):
            for c in ('E', 'K'):
                v = b.wb[sh][f'{c}{r}'].value
                m = re.search(r"('rawdata\(As\)'!\$?[A-Z]+\$?\d+)", str(v))
                if m:
                    note(_wrap_blank(b, sh, f'{c}{r}', m.group(1)), f'{sh}!{c}{r}', '기기값 빈칸이면 평균에서 뺌')
            note(_wrap_blank(b, sh, f'H{r}', f'E{r}'), f'{sh}!H{r}', '빈칸이면 평균에서 뺌')
            note(_wrap_blank(b, sh, f'N{r}', f'K{r}'), f'{sh}!N{r}', '빈칸이면 평균에서 뺌')
    if '중금속' in base and W('계산과정'):
        for col in 'CDEFGHI':
            for r in (6, 7, 8, 11, 12, 13):
                v = b.wb['계산과정'][f'{col}{r}'].value
                m = re.search(r"('Raw data'!\$?[A-Z]+\$?\d+)", str(v))
                if m:
                    note(_wrap_blank(b, '계산과정', f'{col}{r}', m.group(1)), f'계산과정!{col}{r}', '기기값 빈칸이면 평균에서 뺌')
    if '총탄화수소' in base and W('계산과정'):
        for r in range(2, 122):
            v = b.wb['계산과정'][f'C{r}'].value
            m = re.search(r"('Raw data'!\$?[A-Z]+\$?\d+)", str(v))
            if m:
                raw_ = m.group(1)
                note(_set(b, '계산과정', f'C{r}', f'=IF({raw_}="","",IF({raw_}<0,0,{raw_}))'), '계산과정!C2:C121', '측정이 없는 시각은 평균에서 뺌, 기기값을 미리 반올림하지 않음')
            # 시각: 기기 파일의 " 2026/11/05 13:08:05"에서 시·분·초를 숫자로 꺼냄 (TIMEVALUE는 엑셀 지역 설정·앞 공백에 따라 #VALUE!),
            # 엑셀이 날짜·시각으로 바꿔 넣은 값이면 시각 부분, 비면 빈칸
            rb = f"'Raw data'!B{r + 12}"
            t = f'RIGHT(TRIM({rb}),8)'
            note(_set(b, '계산과정', f'B{r}', f'=IFERROR(IF(ISNUMBER({rb}),MOD({rb},1),IF(TRIM({rb})="","",TIME(VALUE(MID({t},1,2)),VALUE(MID({t},4,2)),VALUE(MID({t},7,2))))),"")'),
                 '계산과정!B2:B121', '시각을 시·분·초 숫자로 (TIMEVALUE는 엑셀 설정에 따라 #VALUE!), 비면 빈칸')
            ev = b.wb['계산과정'][f'E{r}'].value
            if isinstance(ev, str) and ev.startswith('=') and not ev.startswith(f'=IF(C{r}="",""'):
                note(_wrap_blank(b, '계산과정', f'E{r}', f'C{r}'), '계산과정!E2:E121', '측정이 없는 시각은 빈칸 (빈칸 × K는 엑셀에서 #VALUE!)')
    if 'CS2' in base or '이황화메틸' in base:
        # 최종 결과값(H32:H37)이 원본부터 숫자 0으로 박혀 있어 시료 농도가 결과에 들어가지 않음 → 시료 농도 × 희석배수 (VOCs와 같게).
        # 이황화탄소: 시료 농도 ppm(표준가스 5.01 ppm, 검정곡선 0.25~2.5) → 결과값 ppb는 ×1000, 이황화메틸: ppb 그대로
        sh = next((ws.title for ws in b.wb.worksheets if '시험기록부' in str(ws['A1'].value or '').replace(' ', '')), None)
        cs2 = 'CS2' in base; col = 'D' if cs2 else 'E'; k = '*1000' if cs2 else ''; d = '/1000' if cs2 else ''
        if sh:
            for r in range(32, 38):
                v = b.wb[sh][f'{col}{r}'].value
                if isinstance(v, str):     # 시트 이름을 바꾸기 전 이름(자기 시트)을 가리키는 글자가 남아 있으면 뗌
                    v = re.sub(r"'([^']+)'!", lambda m: m.group(0) if m.group(1) in b.sheets else '', v)
                if isinstance(v, str) and v.startswith('=') and not v.startswith(f'=IF(O{r - 16}=""'):
                    note(_set(b, sh, f'{col}{r}', f'=IF(O{r - 16}="","",{v[1:]})'), f'{sh}!{col}32:{col}37', '기기 농도가 빈 줄(분석 안 함)은 빈칸으로 평균에서 뺌')
                note(_set(b, sh, f'H{r}', f'=IF({col}{r}="","",{col}{r}*G{r}{k})'), f'{sh}!H32:H37', '최종 결과값 = 시료 농도 × 희석배수 (원본은 숫자 0으로 박혀 있어 시료 농도가 결과에 들어가지 않음)')
            S_, B_ = f'AVERAGE(H35:H37){d}', f'IF(COUNT(H32:H34)=0,0,AVERAGE(H32:H34){d})'
            note(_set(b, sh, 'I32', f'=IF(COUNT(H35:H37)=0,"",IF({S_}-{B_}<0,0,TRUNC(ROUND({S_}-{B_},8),3)))'), f'{sh}!I32', '시료 평균 − 바탕 평균 (바탕을 분석하지 않았으면 0), 시료가 모두 비면 빈칸')
    if 'PAHs' in base and W('계산과정'):
        # 검증시료(Re_STD) 상대감응계수: 원본은 검증시료의 벤조(a)피렌 면적을 다른 주입(표준용액)의 내부표준 면적·농도로 나눔
        # → 같은 주입의 값으로 (ES 01525.1a 식 9: RRF = An/Ais × Cis/Cn, 모두 그 표준용액의 값)
        for r in range(55, 62):
            v = b.wb['계산과정'][f'B{r}'].value
            if isinstance(v, str) and v.startswith('=') and f'MATCH(A{r},' in v:
                new = re.sub(r'MATCH\(A\d+,', f'MATCH(A{r},', v)
                note(_set(b, '계산과정', f'B{r}', new), f'계산과정!B{r}', '검증시료 RRF를 같은 주입의 내부표준 면적·농도로 (식 9)')
    if 'PAHs' in base and W('시험기록부'):
        for r in range(21, 27):
            v = b.wb['시험기록부'][f'I{r}'].value
            if isinstance(v, str) and v.startswith('=') and 'ISTEXT' not in v:
                note(_set(b, '시험기록부', f'I{r}', f'=IF(ISTEXT(B{r}),{v[1:]},"")'), f'시험기록부!I{r}', '시료명이 없는 줄은 평균에서 뺌 (면적 n.a.는 0)')
    # 같은 메시지는 한 번만
    out = []
    for l in log:
        if l not in out:
            out.append(l)
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


def _guard_truncs(f):
    """TRUNC(식, n)(n은 숫자, 식이 ROUND로 시작하지 않음) → TRUNC(ROUND(식, n+5), n): 이진 소수 오차(0.025 → 0.02499…)로
    버림이 한 단위 내려가지 않게"""
    out, i = [], 0
    pat = re.compile(r'(?<![A-Z_.])TRUNC\(')
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
        arg = f[m.end():comma] if comma else ''
        if re.fullmatch(r'-?\d+', n) and not arg.lstrip().startswith('ROUND('):
            out.append(f[i:m.start()]); out.append(f'TRUNC(ROUND({_guard_truncs(arg)},{int(n) + 5}),{n})'); i = j
        else:
            out.append(f[i:m.end()]); i = m.end()
    return ''.join(out)


def strip_pre_rounding(b, sheet, ref, calc=None):
    """결과 수식이 읽는 계산값 칸 R의 ROUND(식, n)을 버림(TRUNC)으로: 결과 수식이 계산 자릿수에서 버리고 최종 자릿수에서
    한 번 맺으므로, 앞에서 먼저 반올림하면 맺음이 두 번이 된다(0.0249 → 0.025 → 0.03). 표시 자릿수는 그대로. 바뀌면 True"""
    v = b.wb[sheet][ref].value
    if not (isinstance(v, str) and v.startswith('=')):
        return False
    nv = '=' + _guard_truncs(_trunc_rounds(v[1:], calc))
    if nv == v:
        return False
    b.setf(sheet, ref, nv[1:])
    b.wb[sheet][ref].value = nv
    return True


# ── 시험 기초자료 방법검출한계·정량한계: 유효숫자 표시 (정도보증 ES 01001.b 4.3.4.4 「유효 숫자는 2 개 이하로 표기한다」) ──
# 유효숫자 개수는 물질마다·방법검출한계와 정량한계마다 따로: 숨김 '기준' 시트 L열(방법검출한계)·M열(정량한계).
# 기본값은 원문 제시값의 유효숫자 (구리 0.016 → 2, 0.05 → 1). 빈칸이면 입력한 값 그대로 표시.


def _norm(v):
    return re.sub(r'\s+', '', str(v or ''))


def lab_cells(ws):
    """시험 기초자료 표의 방법검출한계·정량한계 입력 칸 [(칸, 'mdl'|'loq', 줄 이름)] (수식 칸은 뺌)"""
    out = []
    for row in ws.iter_rows(max_row=min(ws.max_row, 60), max_col=4):
        for c in row:
            if not _norm(c.value).startswith('시험기초자료'):
                continue
            head = c.row; cols = {}
            for rr in range(c.row, c.row + 2):
                for cc in range(3, 15):
                    h = _norm(ws.cell(rr, cc).value)
                    if ('검출한계' in h or '정량한계' in h or '정량범위' in h) and '방법바탕' not in h:
                        cols[cc] = 'mdl' if '검출' in h else 'loq'; head = rr
            for rr in range(head + 1, head + 6):
                if re.fullmatch(r'\d{1,2}', str(ws.cell(rr, 1).value or '').strip()):
                    break                      # 다음 번호 칸(측정결과 등)
                cells = [(ws.cell(rr, cc), k) for cc, k in cols.items()]
                if all(x.value is None for x, _ in cells):
                    break
                lim = min(cols) if cols else 15
                label = next((str(ws.cell(rr, cc).value).strip() for cc in range(2, lim) if isinstance(ws.cell(rr, cc).value, str)
                              and not _norm(ws.cell(rr, cc).value).startswith('시험기초자료')), '')
                out += [(x.coordinate, k, label) for x, k in cells if not (isinstance(x.value, str) and x.value.startswith('='))
                        and (x.value is None or isinstance(x.value, (int, float)))]
            return out
    return out


def _numfmt_dxfs(b, decs):
    """소수 자릿수마다 표시 형식만 바꾸는 조건부 서식 스타일(dxf) 번호"""
    st = b.text('xl/styles.xml')
    ids = [int(x) for x in re.findall(r'numFmtId="(\d+)"', st)]
    base = max(ids + [163]) + 1
    code = lambda d: '0' if d == 0 else '0.' + '0' * d
    m = re.search(r'<dxfs count="\d+"\s*/>', st)
    if m:
        st = st.replace(m.group(0), '<dxfs count="0"></dxfs>')
    m = re.search(r'<dxfs count="(\d+)">(.*?)</dxfs>', st, re.S)
    if not m:
        st = re.sub(r'(</cellStyles>)', r'\1<dxfs count="0"></dxfs>', st, count=1)
        m = re.search(r'<dxfs count="(\d+)">(.*?)</dxfs>', st, re.S)
    items = re.findall(r'<dxf>.*?</dxf>|<dxf/>', m.group(2), re.S)
    idx = {}
    for d in decs:
        old = next((i for i, it in enumerate(items) if it.startswith('<dxf><numFmt') and it.endswith(f'formatCode="{code(d)}"/></dxf>')), None)
        if old is None:
            items.append(f'<dxf><numFmt numFmtId="{base + d}" formatCode="{code(d)}"/></dxf>'); old = len(items) - 1
        idx[d] = old
    st = st.replace(m.group(0), f'<dxfs count="{len(items)}">{"".join(items)}</dxfs>')
    b.put('xl/styles.xml', st)
    return idx


DECS = range(0, 10)


def sig_rule(ref, S):
    """칸 ref를 유효숫자 S개(S = '기준' 칸 참조)로 나타낼 때 소수 자릿수 = MAX(0, S-1-INT(LOG10(|S개로 맺은 값|)))
    맺은 값으로 자릿수를 정해 0.00996 → 0.010처럼 자리가 올라가도 맞게. LOG10의 이진 오차는 1E-9로 막음"""
    lg = lambda x: f'INT(LOG10(ABS({x}))+1E-9)'
    r = f'ROUND({ref},{S}-1-{lg(ref)})'
    return f'MAX(0,{S}-1-{lg(r)})'


def sig_figs(b, path, rows, SHEET='기준', COL=None):
    """시험 기초자료 칸: 값을 길게 적어도 '기준' 시트 L·M열의 유효숫자 개수로 보이게 (소수 자릿수마다 조건부 서식).
    저장값은 그대로 (결과 수식은 이 칸을 읽지 않음). 유효숫자 칸이 비면 그 칸의 원래 표시 형식"""
    log = []
    D = None
    col = {'mdl': (COL or {}).get('sig_mdl', 'L'), 'loq': (COL or {}).get('sig_loq', 'M')}
    for ws in b.wb.worksheets:
        if ws.sheet_state != 'visible' or '시험기록부' not in _norm(ws['A1'].value):
            continue
        mine = [r for r in rows if r['sheet'] == ws.title]
        if not mine:
            continue
        cells = lab_cells(ws)
        if not cells:
            continue
        D = D or _numfmt_dxfs(b, DECS)
        s = b.sx(ws.title)
        pri = max([int(x) for x in re.findall(r'priority="(\d+)"', s)] + [0])
        cf = ''; done = []
        for ref, kind, label in cells:
            r = next((r for r in mine if label and _norm(r['item']) == _norm(label)), mine[0])   # SOx·NOx·CO: 줄 이름 = 물질
            S = f"N('{SHEET}'!${col[kind]}${r['row']})"
            if f'<conditionalFormatting sqref="{ref}"><cfRule type="expression" dxfId="{D[0]}"' in s:
                continue
            rules = ''
            for d in DECS:
                pri += 1
                rules += (f'<cfRule type="expression" dxfId="{D[d]}" priority="{pri}"><formula>'
                          f'AND(ISNUMBER({ref}),{ref}&lt;&gt;0,{S}&gt;=1,{sig_rule(ref, S)}={d})</formula></cfRule>')
            cf += f'<conditionalFormatting sqref="{ref}">{rules}</conditionalFormatting>'
            done.append(f'{ref}({"검출" if kind == "mdl" else "정량"} ' + (f'{r["sig_" + kind]}개' if r['sig_' + kind] else '빈칸: 입력값 그대로') + ')')
        if not cf:
            continue
        m = re.search(r'<(dataValidations|hyperlinks|printOptions|pageMargins|pageSetup|headerFooter|rowBreaks|colBreaks|customProperties|drawing|legacyDrawing|tableParts|extLst)\b', s)
        s = s[:m.start()] + cf + s[m.start():] if m else s.replace('</worksheet>', cf + '</worksheet>')
        b.setsx(ws.title, s)
        log.append(f'{ws.title}: 시험 기초자료 ' + ', '.join(done))
    return log

