# 원문 근거 캡처: 공정시험기준 PDF에서 판정 근거가 되는 문장(결과 자릿수·방법검출한계·정량범위·총칙 규정)을 찾아
# 그 부분을 잘라 형광펜 표시한 그림(JPEG, base64)으로 만든다. 프로그램의 "근거 보기"에서 원문 그대로 보여 준다.
#   python3 -I tools/labform/evidence.py <공정시험기준 PDF 폴더> <master_checked.json> <결과 evidence.json>
import sys, os, re, json, glob, io, base64
import pdfplumber
from PIL import Image, ImageDraw

REF_DIR, MASTER, OUT = sys.argv[1:4]
DPI = 120
S = DPI / 72
rows = json.load(open(MASTER, encoding='utf8'))
SYMBOL = {'구리': 'Cu', '납': 'Pb', '아연': 'Zn', '크로뮴': 'Cr', '니켈': 'Ni', '카드뮴': 'Cd', '베릴륨': 'Be'}
ns = lambda t: re.sub(r'\s+', '', t or '')


def pdf_of(code):
    c = glob.glob(os.path.join(REF_DIR, f'ES {code} *.pdf')) or glob.glob(os.path.join(REF_DIR, f'ES {code}*.pdf'))
    return c[0] if c else None


_cache = {}


def lines_of(path):
    """쪽마다 줄 목록: (쪽 번호, 줄 글자(공백 뺌), 단어들)"""
    if path in _cache:
        return _cache[path]
    out = []
    with pdfplumber.open(path) as pdf:
        for pi, page in enumerate(pdf.pages):
            words = page.extract_words(keep_blank_chars=False, use_text_flow=True)
            rows_ = []
            for w in sorted(words, key=lambda w: (round(w['top']), w['x0'])):
                if rows_ and abs(rows_[-1][0]['top'] - w['top']) < 3.5:
                    rows_[-1].append(w)
                else:
                    rows_.append([w])
            for ws in rows_:
                ws.sort(key=lambda w: w['x0'])
                out.append(dict(page=pi, text=ns(''.join(w['text'] for w in ws)), words=ws,
                                top=min(w['top'] for w in ws), bottom=max(w['bottom'] for w in ws),
                                x0=min(w['x0'] for w in ws), x1=max(w['x1'] for w in ws), pw=float(page.width)))
    _cache[path] = out
    return out


def find(path, test, after=None, skip_toc=True):
    """test(줄, 줄 목록, i)가 참인 첫 줄 번호. 목차(첫 쪽의 '…' 줄)는 건너뛴다"""
    L = lines_of(path)
    for i, ln in enumerate(L):
        if after is not None and i <= after:
            continue
        if skip_toc and re.search(r'[·…]{3,}|\.{5,}', ln['text']):
            continue
        if test(ln, L, i):
            return i
    return None


SEC = re.compile(r'^[1-9]\d*(\.\d+){1,3}[가-힣“"(]')               # 절 번호로 시작하는 줄 (문장 시작)
SEC_ONLY = re.compile(r'^[1-9]\d*(\.\d+){1,3}[가-힣\s]{0,12}$')   # 절 제목 줄
HDR = re.compile(r'^(ES\d{5}\.\w+(\d{4})?|-\d+-|\d{4}|ES\d{5}\.\w+)$')
is_hdr = lambda ln: HDR.match(ln['text']) is not None


def _render(path, pg):
    import pypdfium2 as pdfium
    doc = pdfium.PdfDocument(path)
    img = doc[pg].render(scale=S).to_pil().convert('RGB')
    doc.close()
    return img


def shot(path, i0, i1, mark=None, pad=2):
    """i0~i1 줄을 형광펜으로, 앞뒤 pad줄(머리·바닥 줄 제외)을 문맥으로 잘라 JPEG base64.
    문장이 쪽을 넘어가면 쪽마다 잘라 위아래로 잇는다"""
    L = lines_of(path)
    mark = mark if mark is not None else list(range(i0, i1 + 1))
    body = [k for k in range(len(L)) if not is_hdr(L[k])]
    pos = {k: n for n, k in enumerate(body)}
    p0 = pos.get(i0, 0); p1 = pos.get(i1, p0)
    keep = body[max(0, p0 - pad): p1 + pad + 1]
    pages = sorted({L[k]['page'] for k in range(i0, i1 + 1)})
    keep = [k for k in keep if L[k]['page'] in pages]
    parts = []
    for pg in pages:
        ks = [k for k in keep if L[k]['page'] == pg]
        top = min(L[k]['top'] for k in ks) - 8
        bot = max(L[k]['bottom'] for k in ks) + 8
        pl = [l for l in L if l['page'] == pg and not is_hdr(l)] or [L[ks[0]]]
        x0 = max(0, min(l['x0'] for l in pl) - 10)
        x1 = min(L[ks[0]]['pw'], max(l['x1'] for l in pl) + 10)
        img = _render(path, pg)
        ov = Image.new('RGBA', img.size, (0, 0, 0, 0))
        d = ImageDraw.Draw(ov)
        for k in mark:
            if L[k]['page'] == pg and not is_hdr(L[k]):
                d.rectangle([(L[k]['x0'] - 2) * S, (L[k]['top'] - 2) * S, (L[k]['x1'] + 2) * S, (L[k]['bottom'] + 2) * S], fill=(255, 214, 0, 95))
        img = Image.alpha_composite(img.convert('RGBA'), ov).convert('RGB')
        parts.append(img.crop((int(x0 * S), int(max(0, top) * S), int(x1 * S), int(bot * S))))
    W = max(p.width for p in parts)
    gap = 14
    out = Image.new('RGB', (W, sum(p.height for p in parts) + gap * (len(parts) - 1)), 'white')
    y = 0
    for n, p in enumerate(parts):
        out.paste(p, (0, y)); y += p.height
        if n < len(parts) - 1:   # 쪽 나뉨 표시: 점선
            ImageDraw.Draw(out).line([(0, y + gap // 2), (W, y + gap // 2)], fill=(170, 176, 180), width=1)
            for xx in range(0, W, 12):
                ImageDraw.Draw(out).line([(xx + 6, y + gap // 2), (xx + 12, y + gap // 2)], fill='white', width=1)
            y += gap
    img = out
    if img.width > 900:
        img = img.resize((900, int(img.height * 900 / img.width)), Image.LANCZOS)
    buf = io.BytesIO()
    img.save(buf, 'JPEG', quality=72, optimize=True)
    text = ' / '.join(L[k]['text'] for k in mark if not is_hdr(L[k]))
    return dict(page=pages[0] + 1 if len(pages) == 1 else f'{pages[0] + 1}~{pages[-1] + 1}', img=base64.b64encode(buf.getvalue()).decode(), text=text)


def sentence(path, i, j=None):
    """i(~j)줄이 들어 있는 문장 전체의 줄 범위. 머리·바닥 줄(ES 번호·쪽 번호)은 건너뛰어 쪽을 넘는 문장도 잇는다"""
    L = lines_of(path)
    j = i if j is None else j
    body = [k for k in range(len(L)) if not is_hdr(L[k])]
    pos = {k: n for n, k in enumerate(body)}
    def near(a, b):   # body 위치 a, b(=a+1)가 같은 문단으로 이어지는가
        A, B = L[body[a]], L[body[b]]
        if A['page'] == B['page']:
            return 0 <= B['top'] - A['bottom'] < 14
        return B['page'] == A['page'] + 1     # 쪽 끝 줄 → 다음 쪽 첫 줄
    a = pos.get(i, 0)
    while a > 0 and a > pos.get(i, 0) - 3 and near(a - 1, a) and not re.search(r'(다\.|\.)$', L[body[a - 1]]['text']) \
            and not SEC_ONLY.match(L[body[a - 1]]['text']) and not SEC.match(L[body[a]]['text']):
        a -= 1
    b = pos.get(j, a)
    while b < len(body) - 1 and b < pos.get(j, a) + 3 and not re.search(r'다\.$', L[body[b]]['text']) and near(b, b + 1):
        b += 1
    return body[a], body[b]


def numtxt(v):
    s = f'{v:.6f}'.rstrip('0').rstrip('.')
    return s


def has_num(text, v):
    """줄에 숫자 v가 따로 있는가: 1.00 = 1.0, 0.045 ≠ 0.04, 절 번호(1.2.2)·"1배"의 1 같은 다른 수와 구별"""
    for m in re.finditer(r'(?<![\d.])(\d+(?:\.\d+)?)(?![\d]|\.\d)', text):
        if abs(float(m.group(1)) - v) < 1e-12 and (('.' in m.group(1)) == ('.' in numtxt(v)) or v >= 1 and not text[m.end():m.end() + 1] in '배회개번차'):
            return True
    return False


ALIAS_NAME = {'1,2-다이클로로에테인': '다이클로로에테인', '1,3-뷰타다이엔': '뷰타다이엔', '1.3-뷰타다이엔': '뷰타다이엔'}


def best(path, cands, item):
    """후보 줄 가운데: 물질 이름이 같은 문장에 있는 것 > 본문 서술형(요약표 '항목:' 꼴이 아닌 것) > 앞쪽"""
    if not cands:
        return None
    L = lines_of(path)
    nm = ns(ALIAS_NAME.get(item, item))[:4]
    def score(i):
        a, b = sentence(path, i)
        t = ''.join(L[k]['text'] for k in range(a, b + 1))
        named = nm in t
        # 문장에 "○○의 방법검출한계"가 있으면 ○○(앞 30자 안)에 이 물질이 있어야 제 값
        wins = [t[max(0, m.start() - 30):m.start()] for m in re.finditer(r'의방법검출한계', t)]
        own = not wins or any(nm in w for w in wins)
        summary = re.search(r'(방법검출한계|정량범위):', L[i]['text']) is not None
        return (not own, summary, not named, i)
    return min(cands, key=score)


def all_lines(path, test):
    L = lines_of(path)
    return [i for i, ln in enumerate(L) if not re.search(r'[·…]{3,}|\.{5,}', ln['text']) and not is_hdr(ln) and test(ln, L, i)]


def span(path, i, label):
    """요약표('방법검출한계:' 꼴)면 머리말 줄~값 줄만, 본문이면 문장 전체"""
    L = lines_of(path)
    for k in range(i, max(-1, i - 3), -1):
        if label + ':' in L[k]['text'] and L[k]['page'] == L[i]['page']:
            return k, i
    return sentence(path, i)


def captures(code, r):
    path = pdf_of(code)
    if not path:
        return []
    ref = r.get('ref') or {}
    item = r.get('item', '')
    out = []
    two = lambda L, i: L[i]['text'] + (L[i + 1]['text'] if i + 1 < len(L) else '')
    prev = lambda L, i: (L[i - 1]['text'] if i > 0 else '') + L[i]['text']
    # 결과 자릿수
    i = find(path, lambda l, L, i: re.search(r'자리까지(유효자리수를)?계산', two(L, i)) is not None and ('자리' in l['text'] or '소수' in l['text']))
    if i is not None:
        a, b = sentence(path, i)
        out.append(dict(kind='digits', title='결과 자릿수', **shot(path, a, b)))
    # 방법검출한계
    mv = ref.get('mdl')
    if mv is not None:
        sym = next((v for k, v in SYMBOL.items() if k in item), None)
        if sym:
            i = find(path, lambda l, L, i: l['text'].startswith(sym) and has_num(l['text'], mv))
            hdr = None if i is None else next((k for k in range(i - 1, max(-1, i - 25), -1) if '검출한계' in lines_of(path)[k]['text'] and lines_of(path)[k]['page'] == lines_of(path)[i]['page']), None)
            if i is not None:
                out.append(dict(kind='mdl', title='방법검출한계', **shot(path, hdr if hdr is not None else i, i, mark=[k for k in (hdr, i) if k is not None], pad=0)))
        else:
            c = all_lines(path, lambda l, L, i: has_num(l['text'], mv) and '방법검출한계' in prev(L, i))
            i = best(path, c, item)
            if i is not None:
                a, b = span(path, i, '방법검출한계')
                out.append(dict(kind='mdl', title='방법검출한계', **shot(path, a, b)))
    # 정량범위: 본문 서술("정량범위는 … 이상")을 요약표("정량범위: …")보다 먼저
    lo = ref.get('range_low')
    if lo is not None:
        c = all_lines(path, lambda l, L, i: has_num(l['text'], lo) and '정량범위' in prev(L, i)
                      and re.search(r'이상|∼|～|~', two(L, i)) is not None)
        i = best(path, c, item)
        if i is not None:
            a, b = span(path, i, '정량범위')
            out.append(dict(kind='range', title='정량범위', **shot(path, a, b)))
    # 측정한 방법검출한계는 제시한 값 이하 (시험 기초자료): 그 문장만
    L = lines_of(path)
    i = find(path, lambda l, L, i: '측정한' in two(L, i) and '제시한' in two(L, i) and ('제시한' in l['text'] or '측정한' in l['text']))
    if i is not None:
        j = i
        while j < len(L) - 1 and j < i + 2 and not re.search(r'이하이어야한다\.', ''.join(L[k]['text'] for k in range(i, j + 1))):
            j += 1
        if L[i]['text'].find('측정한') < 0 and i > 0:
            pass
        out.append(dict(kind='lab', title='시험 기초자료 기준', **shot(path, i, j)))
    return out


def rule_shots():
    """총칙·정도보증의 공통 근거"""
    out = {}
    g = pdf_of('01000.b')
    if g:
        L = lines_of(g)
        for key, title, test in (
                ('calc', '총칙 2.14.2 계산과 반올림', lambda l, L, i: '다음단위까지계산' in l['text'] or ('유효숫자' in l['text'] and '다음' in l['text'] and '계산' in l['text'])),
                ('nd', '총칙 2.14.3 불검출', lambda l, L, i: '방법검출한계' in l['text'] and '미만' in l['text']),
                ('lt', '총칙 2.13.6 미만', lambda l, L, i: '미만' in l['text'] and '기산점' in l['text']),
                ('unit', '총칙 2.14.1 단위', lambda l, L, i: 'μmol/mol' in l['text'] or 'µmol/mol' in l['text'] or 'mol/mol' in l['text'])):
            i = find(g, test)
            if i is not None:
                a, b = sentence(g, i)
                out[key] = dict(title=title, code='01000.b', **shot(g, a, b))
    q = pdf_of('01001.b')
    if q:
        L = lines_of(q)
        for key, title, test in (
                ('sig', '정도보증 4.3.4.4 유효숫자', lambda l, L, i: '유효숫자' in l['text'] and ('2개' in l['text'] or '두' in l['text'])),
                ('ratio', '정도보증 4.3.2 방법검출한계 계산', lambda l, L, i: '3.14' in l['text'])):
            i = find(q, test)
            if i is not None:
                a, b = sentence(q, i)
                out[key] = dict(title=title, code='01001.b', **shot(q, a, b))
    return out


ev = {'codes': {}, 'rules': rule_shots()}
for r in rows:
    ref = r.get('ref') or {}
    code = ref.get('code')
    if not code:
        continue
    k = f'{code}|{r.get("item")}'
    if k in ev['codes']:
        continue
    caps = captures(code, r)
    if ref.get('via'):
        have = {c['kind'] for c in caps}
        caps += [dict(c, title=c['title'] + f' (ES {ref["via"]})', code=ref['via']) for c in captures(ref['via'], r) if c['kind'] not in have]
    for c in caps:
        c.setdefault('code', code)
    ev['codes'][k] = caps
    print(f'{code:10} {r.get("item", "")[:10]:10} ' + ', '.join(f'{c["title"]} {c["page"]}쪽' for c in caps))
import hashlib
pool = {}
def pooled(c):
    h = hashlib.sha1(c['img'].encode()).hexdigest()[:12]
    pool[h] = c['img']
    return dict({k: v for k, v in c.items() if k != 'img'}, img=h)
ev['codes'] = {k: [pooled(c) for c in v] for k, v in ev['codes'].items()}
ev['rules'] = {k: pooled(v) for k, v in ev['rules'].items()}
ev['img'] = pool
json.dump(ev, open(OUT, 'w', encoding='utf8'), ensure_ascii=False)
n = len(ev['img'])
print(f'근거 그림 {n}개, {os.path.getsize(OUT) / 1e6:.1f} MB; 공통 규정: {", ".join(v["title"] for v in ev["rules"].values())}')
