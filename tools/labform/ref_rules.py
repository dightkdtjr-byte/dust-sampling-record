# 대기오염공정시험기준 원문(텍스트)에서 물질별 결과 자릿수·단위·제시 방법검출한계·정량범위를 뽑는다.
#   python3 -I tools/labform/ref_rules.py <원문 텍스트 폴더> <결과.json>
# 원문 문장 예
#   "측정결과는 ppm 단위로 소수점 셋째 자리까지 계산하고, 결과 표시는 소수점 둘째 자리로 표기한다."
#   요약표 "방법검출한계: 0.03 ppm", "정량범위: 0.10 ppm 이상"
import sys, os, re, glob, json

NUMW = {'첫째': 1, '둘째': 2, '셋째': 3, '넷째': 4, '다섯째': 5, '여섯째': 6, '일곱째': 7}
UNITN = r'(?:ppm|ppb|mg/Sm3|μg/Sm3|ug/Sm3|ng/Sm3|mg/m3|μg/m3|ng/m3|mg/Sm³|μg/Sm³|ng/Sm³|%)'
UNIT = r'(ppm|ppb|mg/Sm3|μg/Sm3|ug/Sm3|ng/Sm3|mg/m3|μg/m3|ng/m3|mg/Sm³|μg/Sm³|ng/Sm³|%)'


def clean(t):
    t = re.sub(r'-\s*\d+\s*-\s*\n\s*\d{4}\s+ES\s+\d{5}\S*', ' ', t)     # 쪽 바닥글
    t = re.sub(r'\s+', ' ', t)
    t = t.replace('Sm 3', 'Sm3').replace('m 3', 'm3').replace('㎍', 'μg')
    t = re.sub(r'(\d) (\d{3})(?!\d)', r'\1\2', t)                       # 2 500 → 2500
    # 줄바꿈으로 끊긴 낱말: "자리까 지", "다섯 째", "첫 째자리", "방법검출한 계"
    t = re.sub(r'자\s*리\s*까\s*지', '자리까지', t)
    t = re.sub(r'(첫|둘|셋|넷|다섯|여섯|일곱)\s+째', r'\1째', t)
    t = re.sub(r'째\s*자\s*리', '째 자리', t)
    t = re.sub(r'방법\s*검출한\s*계', '방법검출한계', t)
    for _ in range(3):
        t = re.sub(r'(\.\d{3}) (\d{1,3})(?!\d)', r'\1\2', t)                 # 0.000 5 → 0.0005 (소수점 아래 세 자리 띄어쓰기)
    return t


def parse(path):
    raw = open(path, encoding='utf8', errors='ignore').read()
    t = clean(raw)
    out = dict(file=os.path.basename(path), digits=[], mdl=[], range=[])
    for m in re.finditer(r'측정\s*결과는\s*(\S+?)\s*단위\S*\s*(?:의\s*)?소수점\s*(\S+?)\s*자\s*리까지[^.。]{0,40}?계산하고[^.。]{0,40}?소수점\s*(\S+?)\s*자\s*리', t):
        c, f = NUMW.get(m.group(2)), NUMW.get(m.group(3))
        out['digits'].append(dict(unit=m.group(1), calc=c, final=f, text=m.group(0)[:160]))
    # 먼지처럼 "다음 식에 의해 소수점 둘째 자리까지 계산하고 소수점 첫째 자리까지 표기"
    for m in re.finditer(r'소수점\s*(\S+?)\s*자\s*리까지\s*계산하고\s*(?:결과\s*표시는\s*)?소수점\s*(\S+?)\s*자\s*리', t):
        c, f = NUMW.get(m.group(1)), NUMW.get(m.group(2))
        if not any(d['calc'] == c and d['final'] == f for d in out['digits']):
            out['digits'].append(dict(unit='', calc=c, final=f, text=m.group(0)[:160]))
    # 요약표의 방법검출한계 (물질 이름이 앞에 붙기도 함)
    for m in re.finditer(r'방법검출한계\s*:\s*(.{0,160}?)(?=정밀도|정량한계\s*:|정확도|$)', t):
        seg = m.group(1)
        for k in re.finditer(rf'(?:([가-힣A-Za-z0-9,.\-() ]{{2,40}}?)\s*:\s*)?([\d.]+)\s*{UNIT}', seg):
            out['mdl'].append(dict(name=(k.group(1) or '').strip(' ,'), value=float(k.group(2)), unit=k.group(3)))
        if '해당 없음' in seg:
            out['mdl'].append(dict(name='', value=None, unit='', note='해당 없음'))
    # 본문 "방법검출한계는 0.03 ppm"
    for m in re.finditer(rf'([가-힣A-Za-z0-9,.\-() ]{{0,40}}?)방법검출한\s*계는\s*([\d.]+)\s*{UNIT}', t):
        v = float(m.group(2))
        if not any(abs((x['value'] or -1) - v) < 1e-12 for x in out['mdl']):
            out['mdl'].append(dict(name=m.group(1).strip()[-30:], value=v, unit=m.group(3), where='본문'))
    for m in re.finditer(rf'정량범위\s*[:은는]\s*(.{{0,120}}?)(?=간섭물질|방법검출|시료채취|이며|이고|$)', t):
        seg = m.group(1)
        k = re.search(rf'\(?\s*([\d.]+)\s*(?:{UNITN})?\s*[∼～~-]\s*([\d.]+)\s*\)?\s*{UNIT}', seg)
        if k:
            out['range'].append(dict(low=float(k.group(1)), high=float(k.group(2)), unit=k.group(3), text=seg[:100]))
            continue
        k = re.search(rf'([\d.]+)\s*{UNIT}\s*이상', seg)
        if k:
            out['range'].append(dict(low=float(k.group(1)), high=None, unit=k.group(2), text=seg[:100]))
    # 금속 공통 원문의 표: "Cd 226.50/214.44/228.80 0.005 이상 0.002"
    for k in re.finditer(r'\b(Cd|Pb|Cr|Cu|Ni|Zn|Be|As|Hg|Fe|Mn)\s+[\d./]+\s+(\d+\.\d+)\s*이상\s+(\d+\.\d+?)(?=\d\.\d|[^\d.]|$)', t):
        out['mdl'].append(dict(name=k.group(1), value=float(k.group(3)), unit='mg/Sm3', where='표 1'))
        out['range'].append(dict(name=k.group(1), low=float(k.group(2)), high=None, unit='mg/Sm3', text=k.group(0)))
    # "ES 01400.2 … 을 따른다" : 따르는 원문
    m = re.search(r'(ES\s*\d{5}(?:\.\d)?)\s*[^.]{0,60}?을\s*따른다', t)
    out['follows'] = m.group(1).replace(' ', '') if m else None
    return out


if __name__ == '__main__':
    src, dst = sys.argv[1], sys.argv[2]
    res = {}
    for p in sorted(glob.glob(os.path.join(src, 'ES *.txt'))):
        m = re.match(r'ES (\d{5}(?:\.\w+)?)', os.path.basename(p))
        if m:
            res[m.group(1)] = parse(p)
    # 따르는 원문의 자릿수·검출한계를 가져온다 (ES 01405.2d → ES 01400.2d)
    for code, r in res.items():
        f = r.get('follows')
        if f and not r['digits']:
            base = f.replace('ES', '')
            src_code = next((c for c in res if c.startswith(base)), None)
            if src_code:
                r['digits'] = res[src_code]['digits']; r['from'] = src_code
                if not r['mdl']:
                    r['mdl'] = res[src_code]['mdl']; r['range'] = r['range'] or res[src_code]['range']
    json.dump(res, open(dst, 'w', encoding='utf8'), ensure_ascii=False, indent=1)
    for code, r in res.items():
        if r['digits'] or r['mdl']:
            d = '; '.join(f"{x['unit']} 계산{x['calc']}·표기{x['final']}" for x in r['digits'][:2])
            md = '; '.join(f"{x['name'][:14]} {x['value']}{x['unit']}" for x in r['mdl'][:5])
            rg = '; '.join(f"{x['low']}~{x['high']} {x['unit']}" for x in r['range'][:2])
            print(f'{code:9} 자릿수[{d}] 검출한계[{md}] 정량범위[{rg}]')
