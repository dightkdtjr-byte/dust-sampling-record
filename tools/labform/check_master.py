# 기준표(master.json)를 대기오염공정시험기준(총칙·정도보증·시료채취방법·분석법)과 대조한다.
#   python3 -I tools/labform/check_master.py <master.json> [골든 카피 폴더]
# 근거
#   총칙 ES 01000.b 1.6·2.14.2  마지막 유효숫자의 다음 단위까지 계산 → KS Q 5002 4.2.2.2 4사5입법으로 기록  ⇒ 계산 자릿수 = 최종 자릿수 + 1
#   총칙 ES 01000.b 2.14.3      방법검출한계 "미만"의 결과는 불검출  ⇒ 같으면 값
#   총칙 ES 01000.b 2.14.1      단위: 가스상 ppm·ppb, 입자상 mg/Sm³·μg/Sm³·ng/Sm³
#   정도보증 ES 01001.b 4.3.2~4.3.4.4  방법검출한계 = 3.14 s, 정량한계 = 10 s, 유효숫자 2개 이하
#   입자상 시료채취 ES 01112.d 8.1.1  먼지 농도: 소수점 둘째 자리까지 계산, 첫째 자리까지 표기
#   비분산적외선 ES 01204.a 8.2.2     측정값 소수점 둘째 자리까지 계산, 결과 첫째 자리까지
import sys, json, re, os, glob, math, zipfile
from decimal import Decimal

RULES = {   # 항목별 자릿수 규정 (물질 이름에 들어 있는 낱말 → (계산, 최종, 근거))
    '먼지': (2, 1, 'ES 01112.d 8.1.1 먼지 농도: 소수 둘째 자리 계산, 첫째 자리 표기'),
}
NDIR = (2, 1, 'ES 01204.a 8.2.2 비분산적외선: 소수 둘째 자리 계산, 첫째 자리 표시')
UNITS = {'ppm', 'ppb', 'mg/sm3', 'μg/sm3', 'ug/sm3', 'ng/sm3', 'mg/sm³', 'μg/sm³', 'ng/sm³'}


def sig(x):
    """유효숫자 개수"""
    d = Decimal(str(x)).normalize()
    t = d.as_tuple()
    digits = len(t.digits)
    if t.exponent > 0:   # 30 → 3E+1
        return len(str(int(d)).rstrip('0')) or 1
    return digits


def decimals(x):
    d = Decimal(str(x)).normalize()
    return max(0, -d.as_tuple().exponent)


def num(v):
    try:
        return float(v)
    except (TypeError, ValueError):
        return None


def check(r):
    res = []   # (판정, 내용, 근거)  판정: 맞음/확인/틀림
    calc, final = r.get('calc'), r.get('final')
    item = r.get('item', ''); method = str(r.get('method', ''))
    # 1 계산 = 최종 + 1
    if isinstance(calc, int) and isinstance(final, int):
        if calc == final + 1:
            res.append(('맞음', f'계산 {calc}자리 = 최종 {final}자리 + 1', '총칙 2.14.2'))
        else:
            res.append(('틀림', f'계산 {calc}자리 ≠ 최종 {final}자리 + 1', '총칙 2.14.2 (마지막 유효숫자의 다음 단위까지 계산)'))
    # 2 항목별 규정
    rule = next((v for k, v in RULES.items() if k in item), None)
    if rule is None and ('비분산' in method or 'NDIR' in method.upper()):
        rule = NDIR
    if rule:
        ok = (calc, final) == rule[:2]
        res.append(('맞음' if ok else '틀림', f'계산 {calc}·최종 {final}자리' + ('' if ok else f' → 규정은 계산 {rule[0]}·최종 {rule[1]}자리'), rule[2]))
    # 3 유효숫자 2개 이하
    mdl, loq = num(r.get('mdl')), num(r.get('loq'))
    for name, v in (('방법검출한계', mdl), ('정량한계', loq)):
        if v is None:
            if str(r.get('mdl' if name == '방법검출한계' else 'loq', '')).startswith('='):
                continue
            if name == '방법검출한계':
                res.append(('확인', f'방법검출한계가 숫자가 아님({r.get("mdl")!r}) → 불검출 판정을 하지 않음', '총칙 2.14.3'))
            continue
        s = sig(v)
        res.append(('맞음' if s <= 2 else '틀림', f'{name} {v:g}: 유효숫자 {s}개', '정도보증 4.3.4.4 (유효숫자 2개 이하)'))
    # 4 정량한계/방법검출한계 ≈ 10/3.14
    if mdl and loq:
        k = loq / mdl
        ok = 2.5 <= k <= 4.0
        res.append(('맞음' if ok else '확인', f'정량한계/방법검출한계 = {k:.2f} (이론값 10/3.14 = 3.18)', '정도보증 4.3.2·4.3.3 (MDL = 3.14s, LOQ = 10s)'))
    # 5 최종 자릿수로 방법검출한계를 나타낼 수 있는가
    if mdl is not None and isinstance(final, int):
        d = decimals(mdl)
        res.append(('맞음' if d <= final else '틀림', f'방법검출한계 {mdl:g}는 소수 {d}자리, 최종 {final}자리' + ('' if d <= final else ' → 최종값으로 검출한계를 나타낼 수 없어 판정이 어긋남'), '총칙 2.14.2·2.14.3'))
    # 6 단위
    u = str(r.get('unit', '')).strip()
    un = u.lower().replace(' ', '').replace('³', '3')
    if u:
        res.append(('맞음' if un in {x.replace('³', '3') for x in UNITS} else '틀림', f'단위 {u}' + ('' if un in {x.replace('³', '3') for x in UNITS} else ' → 총칙 단위가 아님 (오타?)'), '총칙 2.14.1'))
    else:
        res.append(('확인', '단위가 비어 있음', '총칙 2.14.1'))
    return res


SYMBOL = {'크로뮴': 'Cr', '구리': 'Cu', '납': 'Pb', '아연': 'Zn', '크롬': 'Cr', '크로뮴': 'Cr', '니켈': 'Ni', '카드뮴': 'Cd', '베릴륨': 'Be', '비소': 'As', '수은': 'Hg'}
ALIAS = {'클로로포름': '클로로폼', '염화비닐': '염화바이닐', '1,3-부타디엔': '1.3-뷰타다이엔', '디클로로메탄': '다이클로로메테인',
         '트리클로로에틸렌': '트라이클로로에틸렌', '불소화합물': '플루오린화합물', '브롬화합물': '브로민화합물', '시안화수소': '사이안화수소',
         '포름알데히드': '폼알데하이드', '아세트알데히드': '아세트알데하이드'}
GAS = {'황산화물': '01307', '질소산화물': '01308', '일산화탄소': '01304'}


def ref_for(r, refs):
    """기준표 한 줄에 해당하는 원문 (코드, 원문 규칙, 버전 메모)"""
    codes = [c.replace(' ', '') for c in re.findall(r'ES\s*(\d{5}(?:\.\s?\w+)?)', str(r.get('method', '')))]
    item = r.get('item', '')
    if item in GAS:
        codes = [c for c in codes if c.startswith(GAS[item])] or codes
    for c in codes:
        if c in refs:
            return c, refs[c], ''
        base = c.split('.')[0]
        ver = c.split('.')[1] if '.' in c else ''
        cand = sorted(k for k in refs if k.split('.')[0] == base and (not ver or k.split('.')[1][:1] == ver[:1]))
        if cand:
            return cand[-1], refs[cand[-1]], f'기록부 {c} ↔ 받은 원문 {cand[-1]} (버전 다름)'
    return None, None, ''


def ref_values(r, ref):
    """원문에서 이 물질의 제시 방법검출한계(pick)와 정량범위 하한"""
    item = r.get('item', '')
    key = ALIAS.get(item, item)
    sym = next((v for k, v in SYMBOL.items() if k in item), None)
    m = [x for x in ref['mdl'] if x.get('value') is not None]
    pick = None
    if sym:
        pick = next((x for x in m if x['name'] == sym), None)
    if pick is None and len({x['value'] for x in m}) > 1:
        pick = next((x for x in m if key[:3] in x['name'] or item[:3] in x['name']), None)
    if pick is None and m:
        pick = m[0]
    # 정량범위는 본문(1.2 적용범위 등)에 있을 때만 쓴다 — 부록 요약표에만 있는 값(총탄화수소)은 쓰지 않는다
    rg = [x for x in ref['range'] if (not sym or x.get('name') in (None, sym)) and x.get('where', '본문') == '본문']
    low = rg[0]['low'] if rg and rg[0].get('low') else None
    if ref.get('range') and all((x.get('high') or 0) >= 1000 and (x.get('low') or 0) == 0 for x in ref['range']):
        low = None      # 자동측정법 "(0 ∼ 1 000) ppm"은 측정범위라 정량범위 하한으로 보지 않는다
    return pick, low


def check_ref(r, refs):
    """원문과 대조: 자릿수, 제시 방법검출한계, 정량범위, 단위, 버전"""
    res = []
    code, ref, vnote = ref_for(r, refs)
    if not ref:
        return [('확인', '시험방법에서 ES 번호를 찾지 못해 원문 대조를 못 함', '-')]
    src = f'ES {code}' + (f' (→ ES {ref["from"]})' if ref.get('from') else '')
    if vnote:
        res.append(('틀림', vnote + ' → 시험방법 칸의 버전 확인 필요', src))
    if ref['digits']:
        d = ref['digits'][0]
        ok = (r.get('calc'), r.get('final')) == (d['calc'], d['final'])
        res.append(('맞음' if ok else '틀림', f'원문: {d["unit"]} 계산 {d["calc"]}·표기 {d["final"]}자리 / 기준표: 계산 {r.get("calc")}·최종 {r.get("final")}자리', f'{src} 「{d["text"][:70]}…」'))
    else:
        res.append(('확인', '원문에서 결과 자릿수 문장을 찾지 못함', src))
    pick, rg_low = ref_values(r, ref)
    mdl, loq = num(r.get('mdl')), num(r.get('loq'))
    linked = str(r.get('mdl', '')).startswith('=')
    # 공정시험기준 칸 = 원문 제시값과 같아야 한다 (불검출 판정에 쓰는 칸)
    if pick and not linked:
        if mdl is not None and abs(mdl - pick['value']) < 1e-12:
            res.append(('맞음', f'공정시험기준 방법검출한계 {mdl:g} = 원문 {pick["value"]:g} {pick["unit"]}', src))
        else:
            res.append(('틀림', f'공정시험기준 방법검출한계 {r.get("mdl")} ≠ 원문 {pick["value"]:g} {pick["unit"]}', src))
    elif not pick and not linked:
        res.append(('확인', f'원문에 방법검출한계 제시값 없음 (공정시험기준 칸 {r.get("mdl")})', src))
    if rg_low is not None and not linked:
        if loq is not None and abs(loq - rg_low) < 1e-12:
            res.append(('맞음', f'공정시험기준 정량한계 {loq:g} = 원문 정량범위 하한 {rg_low:g}', src))
        else:
            res.append(('틀림', f'공정시험기준 정량한계 {r.get("loq")} ≠ 원문 정량범위 하한 {rg_low:g}', src))
    # 시험 기초자료(실험실 측정값) ≤ 원문 제시값
    for k, lab, refv, nm in (('lab_mdl', num(r.get('lab_mdl')), pick['value'] if pick else None, '방법검출한계'),
                             ('lab_loq', num(r.get('lab_loq')), rg_low, '정량한계')):
        if refv is None or k not in r:
            continue
        if lab is None or lab == 0:
            res.append(('확인', f'시험 기초자료 {nm}가 비어 있음 ({r.get(k)!r})', src))
        elif lab <= refv + 1e-12:
            res.append(('맞음', f'시험 기초자료 {nm} {lab:.4g} ≤ 원문 {refv:g}', f'{src} 「측정한 방법검출한계는 시험방법에서 제시한 값 이하」'))
        else:
            res.append(('틀림', f'시험 기초자료 {nm} {lab:.4g} > 원문 {refv:g}', f'{src} 「측정한 방법검출한계는 시험방법에서 제시한 값 이하」'))
    u = str(r.get('unit', '')).lower().replace('³', '3').replace(' ', '')
    ru = (ref['digits'][0]['unit'] if ref['digits'] and ref['digits'][0]['unit'] else (pick or {}).get('unit', '')).lower().replace('³', '3')
    if u and ru:
        res.append(('맞음' if u == ru else '틀림', f'단위 기준표 {r.get("unit")} / 원문 {ru}', src))
    return res


def check_formulas(golden):
    """골든 카피 결과 수식이 '미만이면 불검출'인지 (같으면 값)"""
    out = {}
    for p in glob.glob(os.path.join(golden, '*.xls*')):
        z = zipfile.ZipFile(p)
        for n in z.namelist():
            if not re.match(r'xl/worksheets/sheet\d+\.xml$', n):
                continue
            x = z.read(n).decode('utf8')
            for m in re.finditer(r'<c r="([A-Z]+\d+)"[^>]*><f[^>]*>([^<]*불검출[^<]*)</f>', x):
                f = m.group(2).replace('&amp;', '&').replace('&quot;', '"').replace('&lt;', '<').replace('&gt;', '>')
                bad = re.search(r'ROUND\(ROUND\([^<>]*\)\)\s*>\s*(?!=)\$?[A-Z]+\d+\s*,\s*(?:IF\(|TEXT\()', f) and '"불검출"' in f.split(',')[-1]
                lt = re.findall(r'\)\)\s*(<=|<|>=|>)', f)
                out[(os.path.basename(p), n, m.group(1))] = (lt, f)
    return out


if __name__ == '__main__':
    rows = json.load(open(sys.argv[1], encoding='utf8'))
    refs = json.load(open(os.path.join(os.path.dirname(sys.argv[1]), 'ref_rules.json'), encoding='utf8')) if os.path.exists(os.path.join(os.path.dirname(sys.argv[1]), 'ref_rules.json')) else {}
    total = {'맞음': 0, '확인': 0, '틀림': 0}
    for r in rows:
        res = check(r) + (check_ref(r, refs) if refs else [])
        if refs:
            code, rf, vnote = ref_for(r, refs)
            if rf:
                d = rf['digits'][0] if rf['digits'] else {}
                sym = next((v for k, v in SYMBOL.items() if k in r.get('item', '')), None)
                mm = [x for x in rf['mdl'] if x.get('value') is not None]
                pk = next((x for x in mm if sym and x['name'] == sym), None) or \
                     (next((x for x in mm if ALIAS.get(r['item'], r['item'])[:3] in x['name'] or r['item'][:3] in x['name']), None) if len({x['value'] for x in mm}) > 1 else None) or (mm[0] if mm else None)
                rg = [x for x in rf['range'] if (not sym or x.get('name') in (None, sym))]
                r['ref'] = dict(code=code, via=rf.get('from'), calc=d.get('calc'), final=d.get('final'), unit=d.get('unit', ''),
                                text=d.get('text', ''), mdl=(pk or {}).get('value'), mdl_unit=(pk or {}).get('unit', ''),
                                range_low=rg[0]['low'] if rg else None, version_note=vnote)
        r['checks'] = res
        bad = [x for x in res if x[0] != '맞음']
        for x in res:
            total[x[0]] += 1
        flag = '✔' if not bad else ('✖' if any(x[0] == '틀림' for x in res) else '△')
        print(f"{flag} {r['file'][6:22]:16} {r['item'][:12]:12} 계산 {r['calc']} 최종 {r['final']} MDL {str(r['mdl'])[:8]:8} LOQ {str(r['loq'])[:8]:8} {r['unit']}")
        for x in bad:
            print(f"      {x[0]}: {x[1]}  [{x[2]}]")
    print('판정 합계', total)
    if len(sys.argv) > 2:
        fs = check_formulas(sys.argv[2])
        wrong = [(k, v) for k, v in fs.items() if any(op in ('<=',) for op in v[0])]
        print(f'골든 카피 불검출 수식 {len(fs)}개 중 "같으면 불검출"(<=)로 판정하는 수식 {len(wrong)}개')
    json.dump(rows, open(sys.argv[1].replace('.json', '_checked.json'), 'w', encoding='utf8'), ensure_ascii=False, indent=1, default=str)
