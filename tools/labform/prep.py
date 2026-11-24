# 시험기록부 양식 정리: 예시 값(업체명·측정값·그림)을 비우고, 수식 오류를 바로잡고, 인쇄 기준(std.py)을 적용한다
# 쓰는 법: python3 -I tools/labform/prep.py <원본 폴더> <결과 폴더>
#   원본 폴더에는 dust.xlsx, gas.xlsx, hcn.xlsx, bap.xlsx, vocs.xlsm 이름으로 회사 양식을 둔다 (회사 양식은 저장소에 올리지 않는다)
import os, sys, re
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from tplkit import Book, expand
from std import standardize, page_setup, find_docno

SRC, DST = sys.argv[1], sys.argv[2]
PERIOD_FMT = '0"일"'


def due(date, hour, period):
    # 시험기한: 13시 전에 인수하면 하루를 덜 센다. DAY()로 날수를 꺼내던 것은 뺐다 (31일을 넘으면 틀리고, 다른 프로그램에서는 하루 어긋난다)
    return f'IF(OR({date}="",{period}=""),"",WORKDAY({date},IF(N({hour})<13,{period}-1,{period})))'


def finish_print(b, forms, others=()):
    # forms: [(시트, 인쇄 범위)] 기록부 본문 → 기준에 맞춘다.  others: [(시트, 가로 여부)] 계산·자료 시트
    doc = find_docno(b, forms[0][0])
    first = None
    for sheet, area in forms:
        hints = [brk.id for brk in b.wb[sheet].row_breaks.brk]
        lay, pages = standardize(b, sheet, area, doc, hints, like=first)
        first = first or lay
        if sheet in [f[0] for f in forms[:2]]:
            print('  ', sheet, '쪽', pages, '| 쪽별 높이', [round(sum(lay.h[r] for r in range(a, e + 1))) for a, e in pages])
            for n in lay.notes: print('     ·', n)
        elif any(n.startswith('!!') for n in lay.notes):
            print('  ', sheet, [n for n in lay.notes if n.startswith('!!')])
    for sheet, land in others:
        page_setup(b, sheet, doc, landscape=land)
    return doc


# ───────── 먼지 ─────────
def dust():
    b = Book(f'{SRC}/dust.xlsx')
    M, C = '시험기록부', '계산과정'
    for s in (M, C): b.unshare(s)
    b.setv(M, 'I2', '(인)'); b.setv(M, 'J2', '(인)')
    b.clear(M, 'D4,D5,E6,I6,D7,I14')
    b.setf(M, 'D6', 'IF(I6="","",I6)')
    b.setv(M, 'H7', None, style=b.style_with_format(b.style_of(M, 'H7'), PERIOD_FMT))
    b.setf(M, 'I7', due('D6', 'E6', 'H7'))
    b.setf(M, 'D14', 'IF(J18="","",J18)')
    b.setf(M, 'A18', f'IF({C}!A9="","",{C}!A9)')
    for ref, src in [('D18', 'C3'), ('F18', 'D3'), ('G18', 'F3'), ('H18', 'G3'), ('I18', 'H3')]:
        b.setf(M, ref, f'IF({C}!{src}="","",{C}!{src})')
    # 표준산소를 적지 않았는데 실측산소만 적으면 #VALUE!가 나던 것: 보정값(N3)이 있을 때만 보정값을 쓴다
    b.setf(M, 'J18', f'IF({C}!K3="","",IF({C}!N3="",ROUND({C}!K3,2),ROUND({C}!N3,2)))')
    b.clear(C, 'A3,D3,F3,G3,H3,L3,M3,B6,E6,B7,E7,B8,C8,E8,F8,A11:B13,D11:E13')
    # 먼지량: 0보다 작으면 글자 "0"이 아니라 숫자 0
    b.setf(C, 'C3', 'IF(C11="","",MAX(0,C11-N(F11)))')
    b.setf(C, 'I3', 'IF(H3="","",H3/13.6)')
    b.setf(C, 'J3', 'IF(OR(D3="",F3="",G3=""),"",D3*((G3+N(I3))/760)*(273/(273+F3)))')
    b.setf(C, 'K3', 'IF(OR(C3="",J3=""),"",C3/J3)')
    b.setf(C, 'N3', 'IF(OR(L3="",M3="",K3=""),"",K3*((21-L3)/(21-M3)))')
    b.setf(C, 'A9', 'IF(A3="","",A3&"-1")')
    b.setf(C, 'D9', 'IF(A3="","",A3&"-FB-1")')
    for col in 'ABDE':
        b.setf(C, f'{col}14', f'IF(COUNT({col}11:{col}13)=0,"",AVERAGE({col}11:{col}13))')
    b.setf(C, 'C11', 'IF(OR(A14="",B14=""),"",(B14-A14)*1000)')
    b.setf(C, 'F11', 'IF(OR(D14="",E14=""),"",(E14-D14)*1000)')
    for s in (M, C): b.drop_cached(s)
    finish_print(b, [(M, 'A1:J31')], [(C, True)])
    b.finish(f'{DST}/labtest-dust.xlsx')


# ───────── 가스 (황산화물·질소산화물·일산화탄소) ─────────
def gas():
    b = Book(f'{SRC}/gas.xlsx')
    M, C = '시험기록부-SOx, CO, NOx', '계산과정'
    for s in (M, C): b.unshare(s)
    b.setv(M, 'H2', '(인)'); b.setv(M, 'J2', '(인)')
    b.clear(M, 'D4,D5,I6,J23,J26,J29')
    b.setf(M, 'D7', 'IF(I6="","",I6)')
    q = C
    for i in range(9):
        r, c = 23 + i, 5 + i
        b.setf(M, f'C{r}', f'IF({q}!C{c}="","",{q}!C{c})')
        b.setf(M, f'D{r}', f'IF({q}!D{c}="","",{q}!D{c})')
        b.setf(M, f'E{r}', f'IF({q}!F{c}="","",{q}!F{c})')
        b.setf(M, f'F{r}', f'IF({q}!G{c}="","",{q}!G{c})')
        b.setf(M, f'G{r}', f'IF({q}!I{c}="","",{q}!I{c})')
    for r in (23, 26, 29):
        b.setf(M, f'H{r}', f'IF(COUNT(G{r}:G{r + 2})=0,"",TRUNC(AVERAGE(G{r}:G{r + 2}),2))')
        b.setf(M, f'I{r}', f'IF(H{r}="","",ROUND(H{r},1))')
    b.clear(C, 'C5:C7,D5:D13,F5:F13,G5:G7,H5')
    for r in range(5, 14):
        b.setf(C, f'I{r}', f'IF(D{r}="","",IF(OR(F{r}="",G{r}=""),D{r},D{r}*((21-F{r})/(21-G{r}))))')
        if r >= 8:
            b.setf(C, f'C{r}', f'IF(C{r - 3}="","",C{r - 3})')
            b.setf(C, f'G{r}', f'IF(G{r - 3}="","",G{r - 3})')
        if r >= 6:
            b.setf(C, f'H{r}', 'IF($H$5="","",$H$5)')
    for s in (M, C): b.drop_cached(s)
    n = b.drop_pictures()
    finish_print(b, [(M, 'A1:J36')], [(C, True), ('rawdata', True)])
    b.finish(f'{DST}/labtest-gas.xlsx')


# ───────── 사이안화수소 ─────────
def hcn():
    b = Book(f'{SRC}/hcn.xlsx')
    M, C = '시험기록부', '계산과정'
    for s in (M, C): b.unshare(s)
    b.setv(M, 'I2', '(인)'); b.setv(M, 'K2', '(인)')
    b.clear(M, 'D4,D5,F6,J6,D7,J17,G28:G31,I28:I31')
    b.setf(M, 'D6', 'IF(J6="","",J6)')
    b.setv(M, 'I7', None, style=b.style_with_format(b.style_of(M, 'I7'), PERIOD_FMT))
    b.setf(M, 'J7', due('D6', 'F6', 'I7'))
    b.setf(M, 'D17', 'IF(J21="","",IF(J21<=J17,IF(J21<J18,"불검출",ROUND(J21,2)),ROUND(J21,2)&" 초과"))')
    for i in range(3):
        r = 21 + i
        b.setf(M, f'A{r}', f'IF({C}!A{6 + i}="","",{C}!A{6 + i})')
        b.setf(M, f'D{r}', f'IF({C}!E{6 + i}="","",{C}!E{6 + i})')
        # 희석배수는 계산과정에 적은 값을 따라간다 (양식에는 1이 따로 적혀 있었다)
        b.setf(M, f'E{r}', f'IF({C}!C{6 + i}="","",{C}!C{6 + i})')
    b.setf(M, 'F21', f'IF({C}!F6="","",{C}!F6)')
    b.setf(M, 'G21', f'IF({C}!F3="","",{C}!F3)')
    b.setf(M, 'H21', f'IF({C}!N3="","",{C}!N3)')
    b.setf(M, 'J21', f'IF({C}!P3="","",TRUNC({C}!P3,3))')
    b.clear(C, 'A3:D8,G3,H3,I3,J3,L3,M3')
    for r in range(3, 9):
        b.setf(C, f'E{r}', f'IF(OR(B{r}="",C{r}="",D{r}=""),"",B{r}*C{r}*D{r})')
    b.setf(C, 'F3', 'IF(COUNT(E3:E5)=0,"",AVERAGE(E3:E5))')
    b.setf(C, 'F6', 'IF(COUNT(E6:E8)=0,"",AVERAGE(E6:E8))')
    b.setf(C, 'K3', 'IF(J3="","",J3/13.6)')
    b.setf(C, 'N3', 'IF(OR(M3="",L3="",I3=""),"",M3*(273/(273+L3))*((I3+N(K3))/760))')
    b.setf(C, 'O3', 'IF(OR(F6="",N3=""),"",((F6-N(F3))*10/N3)*(22.4/26.017))')
    b.setf(C, 'P3', 'IF(O3="","",O3)')
    for s in (M, C): b.drop_cached(s)
    n = b.drop_pictures()
    finish_print(b, [(M, 'A1:L47')], [(C, True), ('rawdata', False)])
    b.finish(f'{DST}/labtest-hcn.xlsx')


# ───────── 벤조(a)피렌 ─────────
BAP_SHEETS = ['Benzo(a)pyrene', 'd12-Benzo(a)pyrene', 'd12-Benzo(e)pyrene', 'd14-Terphenyl', 'd10-Fluoranthene', 'd10-Pyrene']


def bap():
    b = Book(f'{SRC}/bap.xlsx')
    M, C = '시험기록부', '계산과정'
    for s in (M, C): b.unshare(s)
    b.setv(M, 'I2', '(인)'); b.setv(M, 'J2', '(인)')
    b.clear(M, 'D4,D5,E6,I6,D7,I17')
    b.setf(M, 'D6', 'IF(I6="","",I6)')
    b.setv(M, 'H7', None, style=b.style_with_format(b.style_of(M, 'H7'), PERIOD_FMT))
    b.setf(M, 'I7', due('D6', 'E6', 'H7'))
    b.setf(M, 'D17', 'IF(J21="","",IF(J21<=I17,IF(J21<I18,"불검출",ROUND(J21,1)),ROUND(J21,1)&" 초과"))')
    b.setf(M, 'J21', 'IF(COUNT(D21:D23)=0,"",TRUNC(AVERAGE(I21:I23)-N(I24),2))')
    for i in range(4):
        r = 21 + i
        for col, src in [('B', 'A'), ('D', 'B'), ('E', 'C'), ('F', 'D')]:
            b.setf(M, f'{col}{r}', f'IF({C}!{src}{94 + i}="","",{C}!{src}{94 + i})')
        b.setf(M, f'G{r}', f'IF({C}!A{94 + i}="","",{C}!E{94 + i})')
        b.setf(M, f'H{r}', f'IF({C}!A{94 + i}="","",{C}!F{94 + i})')
        b.setf(M, f'I{r}', f'IF({C}!A{94 + i}="","",IFERROR((D{r}/E{r})*(F{r}/G{r})*(1/H{r}),0))')
    for r in range(29, 37):
        b.setf(M, f'F{r}', f'IF({C}!M{r + 14}="","",{C}!M{r + 14})')
        b.setf(M, f'H{r}', f'IF({C}!N{r + 14}="","",{C}!N{r + 14})')
    b.setf(M, 'J50', 'IF(COUNT(D51:F51)=0,"",IF(AND(D51>50,D51<150,E51>50,E51<150,F51>50,F51<150),"적합","부적합"))')
    b.setf(M, 'J51', 'IF(G51="","",IF(AND(G51>-30,G51<30),"적합","부적합"))')
    for ref, src in [('D51', 'C86'), ('E51', 'E86'), ('F51', 'F86'), ('G51', 'B52')]:
        b.setf(M, ref, f'IF({C}!{src}="","",{C}!{src})')

    # 기기 시트 이름을 글자로 이어 붙여 찾던 수식(INDIRECT)을 시트를 바로 가리키는 수식으로
    def direct(m):
        name = b.wb[C][f'{m.group(1)}2'].value
        return f"'{name}'!{m.group(2)}"
    n = b.sub_formulas(C, r'''INDIRECT\("'"&([A-Z])\$2&"'!([A-Z]+\d+)"\)''', direct)
    # 양식 8행(B8)은 조건이 B16을 보고 있었다 → 제 줄을 보게
    b.setf(C, 'B8', "IF('Benzo(a)pyrene'!B17=\"\",\"\",'Benzo(a)pyrene'!B17)")
    b.clear(C, 'D7:D36,F7:F36,H7:H36,J7:J36,L7:L36,N7:N36,A71:A74,A112:D112,F112,G112')
    look = lambda col, key: f'INDEX(${col}$7:${col}$36,MATCH({key},$B$7:$B$36,0))'
    # 검정곡선 표: 13~18행이라는 자리가 아니라 STD 이름으로 찾는다 (주입 순서가 바뀌어도 맞게)
    for r in range(43, 49):
        b.setf(C, f'L{r}', f'IF(ISNUMBER(MATCH(A{r},$B$7:$B$36,0)),A{r},"")')
        b.setf(C, f'M{r}', f'IFERROR({look("D", f"A{r}")},"")')
        b.setf(C, f'N{r}', f'IFERROR({look("C", f"A{r}")}/{look("G", f"A{r}")}*100,"")')
    b.setf(C, 'B51', 'IF(COUNT(B43:B50)=0,"",AVERAGE(B43:B50))')
    for col in 'CEF':
        b.setf(C, f'{col}51', f'IF(COUNT({col}43:{col}50)=0,"",AVERAGE({col}43:{col}50))')
    b.setf(C, 'B52', 'IF(COUNT(B43:B50)<2,"",(STDEV(B43:B50)/AVERAGE(B43:B50))*100)')
    b.setf(C, 'B59', 'IF(COUNT(B57:B58)<2,"",(STDEV(B57:B58)/AVERAGE(B57:B58)*100))')
    # 방법바탕시료 비교의 표준용액 평균 면적도 STD 이름으로 모은다
    b.setf(C, 'E64', 'IFERROR(AVERAGEIF($B$7:$B$36,"STD-*",$E$7:$E$36),"")')
    b.setf(C, 'E65', 'IFERROR(AVERAGEIF($B$7:$B$36,"STD-*",$K$7:$K$36),"")')
    for r in (64, 65):
        b.setf(C, f'B{r}', f'IFERROR({look("E", f"$A${r}")},"")')
        b.setf(C, f'C{r}', f'IFERROR({look("K", f"$A${r}")},"")')
    b.setf(C, 'B66', 'IFERROR(AVERAGE(B64:B65)/E64*100,"")')
    b.setf(C, 'C66', 'IFERROR(AVERAGE(C64:C65)/E65*100,"")')
    for col in 'BCEF':
        b.setf(C, f'{col}86', f'IF(COUNT({col}71:{col}74)=0,"",AVERAGE({col}71:{col}74))')
        # 검증시료 평균: 벤조(a)피렌 칸만 78~81행을 보고 있었다 → 다른 칸과 같이 77~78행
        b.setf(C, f'{col}87', f'IF(COUNT({col}77:{col}78)=0,"",AVERAGE({col}77:{col}78))')
    for r in range(94, 109):
        if r <= 97:
            b.setf(C, f'A{r}', f'IF(A{r - 23}="","",A{r - 23})')
        elif r <= 101:
            b.setf(C, f'A{r}', f'IF(ISNUMBER(MATCH(A{r - 23},$B$7:$B$36,0)),A{r - 23},"")')
        b.setf(C, f'E{r}', f'IF(A{r}="","",$B$51)')
        b.setf(C, f'F{r}', f'IF(A{r}="","",$E$112)')
        b.setf(C, f'G{r}', f'IF(A{r}="","",IFERROR((B{r}/C{r})*(D{r}/E{r})*(1/F{r}),0))')
    b.setf(C, 'E112', 'IF(OR(A112="",B112="",C112=""),"",A112*(273/(273+B112))*((C112+(N(D112)/13.6))/760))')
    for s in (M, C): b.drop_cached(s)
    # 기기 시트: 주입 목록(16행부터)과 머리의 예시 정보를 비운다
    for name in BAP_SHEETS:
        ws = b.wb[name]
        b.clear(name, f'A16:G{max(ws.max_row, 45)}')
        b.clear(name, 'C4,C5,C6,C7,G4,G5,G6,G7')
    finish_print(b, [(M, 'A1:J54')], [(C, False)] + [(name, False) for name in BAP_SHEETS])
    b.finish(f'{DST}/labtest-bap.xlsx')


# ───────── 휘발성유기화합물 ─────────
def vocs():
    b = Book(f'{SRC}/vocs.xlsm')
    M = '시험기록부-VOCs'
    subs = [s for s in b.sheets if s != M]
    b.unshare(M)
    b.setv(M, 'I2', '(인)'); b.setv(M, 'J2', '(인)')
    b.clear(M, 'D4,D5,E6,I6,D7')
    b.setf(M, 'D6', 'IF(I6="","",I6)')
    # 물질 목록: 이름(시트 이름)을 적은 줄만 그 시트의 값을 가져온다. 이름이 빈 줄은 #REF! 대신 빈칸
    for r in range(26, 41):
        at = lambda ref: f'''INDIRECT("'"&$A{r}&"'!{ref}")'''
        b.setf(M, f'D{r}', f'IF($A{r}="","",{at("I32")})')
        b.setf(M, f'F{r}', f'IF($A{r}="","",{at("W18")})')
        b.setf(M, f'G{r}', f'IF($A{r}="","",IF({at("I27")}="","",{at("I27")}))')
        b.setf(M, f'I{r}', f'IF($A{r}="","",{at("D27")})')
    b.setv(M, 'H7', None, style=b.style_with_format(b.style_of(M, 'H7'), PERIOD_FMT))
    b.setf(M, 'I7', due('D6', 'E6', 'H7'))
    b.setsx(M, re.sub(r'<autoFilter\b[^>]*?(?:/>|>.*?</autoFilter>)', '', b.sx(M), flags=re.S))
    b.drop_cached(M)
    q = f"'{M}'"
    period_style = None
    for s in subs:
        b.unshare(s)
        b.setf(s, 'D4', f'IF({q}!D4="","",{q}!D4)')
        b.setf(s, 'D5', f'IF({q}!D5="","",{q}!D5)')
        b.setf(s, 'D6', f'IF({q}!D6="","",{q}!D6)')
        b.setf(s, 'E6', f'IF({q}!E6="","",{q}!E6)')
        b.setf(s, 'I6', f'IF({q}!I6="","",{q}!I6)')
        b.setf(s, 'D7', f'IF({q}!D7="","",{q}!D7)')
        if period_style is None:
            period_style = {}
        base = b.style_of(s, 'H7')
        if base not in period_style:
            period_style[base] = b.style_with_format(base, PERIOD_FMT)
        b.setf(s, 'H7', f'IF({q}!H7="","",{q}!H7)', style=period_style[base])
        b.setf(s, 'I7', due('D6', 'E6', 'H7'))
        b.clear(s, 'K4:Q13,K17:O27,I27')
        # 시료 표의 주입량·희석배수는 오른쪽 표(W·X열)에 적은 값을 따라간다
        for r, src in [(32, 16), (33, 17), (34, 18), (35, 21), (36, 24)]:
            b.setf(s, f'F{r}', f'IF(W{src}="","",W{src})')
            b.setf(s, f'G{r}', f'IF(X{src}="","",X{src})')
        b.setf(s, 'B32', 'IF(U16="","",U16)'); b.setf(s, 'B33', 'IF(U17="","",U17)')
        b.setf(s, 'D32', 'IF(U16="","",V16)'); b.setf(s, 'D33', 'IF(U17="","",V17)')
        b.setf(s, 'V16', 'IF(O17="",0,O17)'); b.setf(s, 'V17', 'IF(O18="",0,O18)')
        for r, src in [(34, 18), (35, 21), (36, 24)]:
            b.setf(s, f'D{r}', f'IF(U{src}="","",Y{src})')
            # 벤젠 시트 34행에는 수식 대신 숫자 6이 적혀 있었다 → 다른 줄과 같은 수식으로
            b.setf(s, f'H{r}', f'IF(D{r}="","",IFERROR(D{r}*G{r},0))')
        b.setf(s, 'H32', 'IF(COUNT(D32:E33)=0,"",IFERROR(AVERAGE(D32:E33),0))')
        b.setf(s, 'I32', 'IF(COUNT(H34:H36)=0,"",IFERROR(IF((AVERAGE(H34:H36)/1000)-(N(H32)/1000)<0,0,TRUNC((AVERAGE(H34:H36)/1000)-(N(H32)/1000),4)),0))')
        b.setf(s, 'D27', 'IF(I32="","",IF(I32<I28,"불검출",IF(AND(I27<>"",I32>=I27),ROUND(I32,2)&" ( 초과 ) ",ROUND(I32,2))))')
        b.drop_cached(s)
    n = b.drop_pictures()
    finish_print(b, [(M, 'A1:J41')] + [(s, 'A1:J64') for s in subs])
    # 자동 필터용 숨은 이름은 예시 범위라 지운다
    wbx = b.text('xl/workbook.xml')
    wbx = re.sub(r'<definedName name="_xlnm\._FilterDatabase"[^>]*>[^<]*</definedName>', '', wbx)
    b.put('xl/workbook.xml', wbx)
    b.finish(f'{DST}/labtest-vocs.xlsm')


for fn in (dust, gas, hcn, bap, vocs):
    fn()
    print('ok', fn.__name__)
