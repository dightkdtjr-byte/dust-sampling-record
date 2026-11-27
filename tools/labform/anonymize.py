# 예시 사업장명: 원본 기록에 있던 실제 사업장명과 시료명 약어를 지어낸 이름으로 바꾼다 (모든 시트의 글자·수식·차트·메모)
# 원본마다 다른 사업장이 들어 있어 이름·약어를 짝지어 바꾼다. 굴뚝 번호(#A16 등)는 그대로.
import re

# (실제 이름 정규식, 지어낸 이름) — 긴 것부터
NAMES = [
    (r'\(주\)\s*대일개발|대일개발\s*\(주\)|대일개발\s*㈜|대일개발', '가람산업(주)'),
    (r'신대한정유산업\s*\(주\)|신대한정유산업\s*㈜|신대한정유산업', '누리정유(주)'),
    (r'악조노벨분체도료\s*\(주\)|악조노벨분체도료\s*㈜|악조노벨분체도료', '다온도료(주)'),
    (r'지에스파워\s*㈜\s*부천열병합발전처|지에스파워\s*\(주\)\s*부천열병합발전처|지에스파워\s*㈜|지에스파워\s*\(주\)|지에스파워', '라온발전(주) 열병합발전처'),
    (r'화성그린에너지\s*[·ㆍ]?\s*밸류\s*\(주\)|화성그린에너지\s*[·ㆍ]?\s*밸류\s*㈜|화성그린에너지\s*[·ㆍ]?\s*밸류', '마루에너지(주)'),
    (r'서울특별시\s*탄천물재생센터|탄천물재생센터', '바다물재생센터'),
    (r'케이지스틸\s*㈜\s*인천공장|케이지스틸\s*\(주\)\s*인천공장|케이지스틸\s*㈜|케이지스틸\s*\(주\)|케이지스틸', '사랑제강(주) 공장'),
]
# 시료명·파일 이름 속 사업장 약어 (앞뒤가 - _ ( ) 공백 또는 칸 처음·끝일 때만)
CODES = {'DIG': 'GRS', 'SDH': 'NRJ', 'AZN': 'DOD', 'GSP': 'RON', 'HGE': 'MRE', 'TCW': 'BDW', 'GS': 'KD', 'SCR': 'HNR'}
_code_re = re.compile(r'(?<![A-Za-z0-9])(' + '|'.join(sorted(CODES, key=len, reverse=True)) + r')(?![A-Za-z0-9])')
_name_res = [(re.compile(p), r) for p, r in NAMES]
_codes_ok = re.compile(r'(^|[-_(\s])(' + '|'.join(CODES) + r')($|[-_)\s])')


def fix_text(s):
    for rx, rep in _name_res:
        s = rx.sub(rep, s)
    # 약어는 시료명처럼 보이는 글자(숫자·-·_가 섞임)나 약어만 있는 칸에서만
    if _codes_ok.search(s) and (re.search(r'\d', s) or s.strip() in CODES):
        s = _code_re.sub(lambda m: CODES[m.group(1)], s)
    return s


def fix_xml(x):
    """XML 한 부분: 글자(<t>), 수식(<f>), 수식이 저장해 둔 결과(<v>), 차트(<c:v>·<a:t>), 속성값의 이름을 바꾸고
    원본 파일 경로(x15ac:absPath: 사용자 폴더·사업장 폴더 이름)는 지운다. (새 XML, 바꾼 곳 수)"""
    n = 0
    y = re.sub(r'<mc:AlternateContent[^>]*>\s*<mc:Choice Requires="x15">\s*<x15ac:absPath [^>]*/>\s*</mc:Choice>\s*</mc:AlternateContent>', '', x)
    y = re.sub(r'<x15ac:absPath [^>]*/>', '', y)
    if y != x:
        n += 1
    def sub(m):
        nonlocal n
        new = fix_text(m.group(2))
        if new != m.group(2):
            n += 1
        return m.group(1) + new + m.group(3)
    y = re.sub(r'(<(?:t|f|v|c:v|a:t)(?:\s[^>]*)?>)([^<]*)(</(?:t|f|v|c:v|a:t)>)', sub, y)
    y = re.sub(r'(=")([^"]*)(")', sub, y)     # 속성값 (정의된 이름·메모 작성자 등)
    return y, n


def anonymize(b):
    """b.parts의 모든 XML(시트·공유 글자·차트·메모·통합 문서·문서 속성)에서 실제 이름을 바꾼다. 바꾼 곳 수를 돌려준다"""
    n = 0
    for part in list(b.parts):
        if not part.endswith(('.xml', '.vml', '.rels')) or part.startswith('xl/styles'):
            continue
        raw = b.parts[part]
        enc = 'utf-16' if raw[:2] in (b'\xff\xfe', b'\xfe\xff') else 'utf8'
        try:
            x = raw.decode(enc)
        except UnicodeDecodeError:
            continue
        y, k = fix_xml(x)
        if y != x:
            b.parts[part] = y.encode(enc); n += k
    return n


def leftovers(text):
    """남은 실제 이름 (검증용)"""
    out = [m.group(0) for rx, _ in _name_res for m in rx.finditer(text) if m.group(0) not in {r for _, r in NAMES}]
    return out


if __name__ == '__main__':   # 마지막 단계로 폴더의 기록부를 제자리에서 (앞 단계가 남긴 저장 결과까지)
    import sys, os, glob, warnings
    warnings.filterwarnings('ignore')
    sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
    from tplkit import Book
    tot = 0
    for p in sorted(glob.glob(os.path.join(sys.argv[1], '*.xls[xm]'))):
        b = Book(p); k = anonymize(b)
        if k:
            b.finish(p); tot += k
    left = 0
    import zipfile
    for p in sorted(glob.glob(os.path.join(sys.argv[1], '*.xls[xm]'))):
        z = zipfile.ZipFile(p)
        for nme in z.namelist():
            raw = z.read(nme) if nme.endswith(('.xml', '.vml', '.rels')) else b''
            t = raw.decode('utf-16' if raw[:2] in (b'\xff\xfe', b'\xfe\xff') else 'utf8', 'ignore')
            l = leftovers(t) + [m.group(1) for m in _code_re.finditer(t) if _codes_ok.search(t[max(0, m.start() - 1):m.end() + 1])]
            if l:
                left += len(l); print('   ✖', os.path.basename(p), nme, l[:3])
    print(f'예시 사업장명: {tot}곳 바꿈, 남은 실제 이름 {left}곳')
