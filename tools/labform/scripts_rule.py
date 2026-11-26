# 위첨자·아래첨자 통일: 글자 칸의 "Sm3"·"m3"(단위)는 3을 위첨자로, 화학식(NH3·H2S·H2O2·SO2·O2 …)의 숫자와
# RRFavg의 avg는 아래첨자로 서식 있는 글자(rich text)로 바꾼다. 이미 위첨자·아래첨자인 부분은 그대로 둔다.
#   · 글자 값은 바뀌지 않는다(보이는 모양만): 이 칸을 참조하는 수식·조건부 서식에 영향 없음
#   · 덧붙이는 글자 조각에는 그 칸의 글꼴(맑은 고딕·크기·굵기·색)을 그대로 적는다
#   · 공유 문자열 하나를 글꼴이 다른 칸들이 같이 쓰면 글꼴별로 복사해 칸마다 연결한다
import re

CHEM = ['H2SO4', 'HNO3', 'H2O2', 'H2S', 'NH3', 'CS2', 'SO2', 'NO2', 'CO2', 'CH4', 'O2', 'N2']
PAT = re.compile(r'(?P<sup>(?<=Sm)3(?!\d)|(?<=/m)3(?!\d)|(?<=\(m)3(?=\)))'
                 r'|(?<![A-Za-z0-9#_\-])(?P<chem>' + '|'.join(CHEM) + r')(?![A-Za-z0-9])'
                 r'|(?<=RRF)(?P<avg>avg)')


def pieces(text):
    """글자를 [(조각, None|'superscript'|'subscript')]로. 바꿀 것이 없으면 None"""
    out, pos = [], 0
    for m in PAT.finditer(text):
        if m.group('sup'):
            out += [(text[pos:m.start()], None), ('3', 'superscript')]
        elif m.group('avg'):
            out += [(text[pos:m.start()], None), ('avg', 'subscript')]
        else:
            out.append((text[pos:m.start()], None))
            for t in re.split(r'(\d+)', m.group('chem')):
                if t:
                    out.append((t, 'subscript' if t.isdigit() else None))
        pos = m.end()
    if not out:
        return None
    out.append((text[pos:], None))
    merged = []
    for t, va in out:
        if not t:
            continue
        if merged and merged[-1][1] == va:
            merged[-1] = (merged[-1][0] + t, va)
        else:
            merged.append((t, va))
    return merged


def _fonts(b):
    st = b.text('xl/styles.xml')
    fonts = re.findall(r'<font\b[^>]*?(?:/>|>.*?</font>)', re.search(r'<fonts\b[^>]*>(.*?)</fonts>', st, re.S).group(1), re.S)
    xfs = re.findall(r'<xf\b[^>]*?(?:/>|>.*?</xf>)', re.search(r'<cellXfs\b[^>]*>(.*?)</cellXfs>', st, re.S).group(1), re.S)
    font_of = [int((re.search(r'fontId="(\d+)"', x) or [0, 0])[1]) for x in xfs]
    return fonts, font_of


def rpr_of(font_xml, va=None):
    """<font> → <rPr> (name → rFont). 순서는 상관없다 (CT_RPrElt는 choice)"""
    inner = re.sub(r'^<font\b[^>]*>|</font>$', '', font_xml.strip()) if not font_xml.strip().endswith('/>') else ''
    inner = inner.replace('<name ', '<rFont ')
    inner = re.sub(r'<vertAlign\b[^>]*/>', '', inner)
    if va:
        inner += f'<vertAlign val="{va}"/>'
    return f'<rPr>{inner}</rPr>'


def _runs_of_si(si):
    """<si> → [(rPr 또는 None, 글자)]"""
    if '<r>' not in si:
        m = re.search(r'<t(?:\s[^>]*)?>(.*?)</t>|<t\s*/>', si, re.S)
        return [(None, m.group(1) if m and m.group(1) is not None else '')]
    out = []
    for r in re.findall(r'<r>(.*?)</r>', si, re.S):
        p = re.search(r'<rPr>.*?</rPr>', r, re.S)
        t = re.search(r'<t(?:\s[^>]*)?>(.*?)</t>', r, re.S)
        out.append((p.group(0) if p else None, t.group(1) if t else ''))
    return out


def convert_si(si, font_xml):
    """바꿀 것이 있으면 새 <si>, 없으면 None. 끝의 발음 정보(rPh·phoneticPr)는 그대로"""
    runs = _runs_of_si(si)
    changed = False
    new = []
    for rpr, text in runs:
        if rpr and 'vertAlign' in rpr:
            new.append((rpr, text)); continue
        ps = pieces(text)
        if ps is None:
            new.append((rpr, text)); continue
        changed = True
        for t, va in ps:
            if rpr:
                p = rpr.replace('</rPr>', f'<vertAlign val="{va}"/></rPr>') if va else rpr
            else:
                p = rpr_of(font_xml, va)
            new.append((p, t))
    if not changed:
        return None
    tail = ''.join(re.findall(r'<rPh\b.*?</rPh>|<phoneticPr\b[^>]*/>', si, re.S))
    body = ''.join(f'<r>{p or ""}<t xml:space="preserve">{t}</t></r>' for p, t in new)
    return f'<si>{body}{tail}</si>'


def unify_scripts(b, sheets):
    """sheets 시트의 글자 칸 위첨자·아래첨자를 통일. 바꾼 칸 수"""
    if 'xl/sharedStrings.xml' not in b.parts:
        return 0
    fonts, font_of = _fonts(b)
    sst = b.text('xl/sharedStrings.xml')
    head = re.search(r'<sst\b[^>]*>', sst).group(0)
    sis = re.findall(r'<si>.*?</si>|<si/>', sst, re.S)
    # 공유 문자열 → 쓰는 칸들의 글꼴
    use = {}
    for sh in sheets:
        x = b.sx(sh)
        for m in re.finditer(r'<c r="([A-Z]+\d+)"([^>]*?)>\s*<v>(\d+)</v>', x):
            if 't="s"' not in m.group(2):
                continue
            st = re.search(r' s="(\d+)"', m.group(2))
            fid = font_of[int(st.group(1))] if st and int(st.group(1)) < len(font_of) else 0
            use.setdefault(int(m.group(3)), {}).setdefault(fid, []).append((sh, m.group(1)))
    remap = {}     # (시트, 칸) → 새 문자열 번호
    n = 0
    for idx, byfont in sorted(use.items()):
        if idx >= len(sis):
            continue
        first = True
        for fid, cells in byfont.items():
            new = convert_si(sis[idx], fonts[fid] if fid < len(fonts) else fonts[0])
            if new is None:
                break
            if first:
                sis[idx] = new; first = False
            else:
                sis.append(new)
                for c in cells:
                    remap[c] = len(sis) - 1
            n += len(cells)
    # 칸 안의 글자(inlineStr)
    for sh in sheets:
        x = b.sx(sh)
        def fix(m):
            nonlocal n
            st = re.search(r' s="(\d+)"', m.group(1))
            fid = font_of[int(st.group(1))] if st and int(st.group(1)) < len(font_of) else 0
            new = convert_si('<si>' + m.group(2) + '</si>', fonts[fid] if fid < len(fonts) else fonts[0])
            if new is None:
                return m.group(0)
            n += 1
            return f'<c{m.group(1)}><is>{new[4:-5]}</is></c>'
        y = re.sub(r'<c((?=[^>]*t="inlineStr")[^>]*)>\s*<is>(.*?)</is>\s*</c>', fix, x, flags=re.S)
        if y != x:
            b.setsx(sh, y)
    if not n:
        return 0
    for sh in sheets:
        cells = {ref: i for (s2, ref), i in remap.items() if s2 == sh}
        if not cells:
            continue
        x = b.sx(sh)
        x = re.sub(r'(<c r="([A-Z]+\d+)"[^>]*?>\s*<v>)(\d+)(</v>)', lambda m: m.group(1) + str(cells[m.group(2)]) + m.group(4) if m.group(2) in cells else m.group(0), x)
        b.setsx(sh, x)
    head2 = re.sub(r'uniqueCount="\d+"', f'uniqueCount="{len(sis)}"', head)
    if 'count="' in head2:
        cnt = int(re.search(r'\bcount="(\d+)"', head2).group(1)) if re.search(r'\bcount="(\d+)"', head2) else 0
        head2 = re.sub(r'\bcount="\d+"', f'count="{max(cnt, len(sis))}"', head2)
    b.put('xl/sharedStrings.xml', re.sub(r'<sst\b[^>]*>.*</sst>', lambda m: head2 + ''.join(sis) + '</sst>', sst, flags=re.S))
    return n
