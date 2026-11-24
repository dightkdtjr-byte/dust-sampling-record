# 시험기록부 양식 원본을 다듬는 도구: 원본 xlsx의 XML을 직접 고친다 (그림·도형·차트를 잃지 않으려고)
import zipfile, re, warnings, io
from xml.sax.saxutils import escape
import openpyxl
from openpyxl.utils import column_index_from_string as CI, get_column_letter as CL

warnings.filterwarnings('ignore')
CELL = lambda ref: re.compile(rf'<c r="{ref}"((?: [^>]*?)?)(?:/>|>.*?</c>)', re.S)


def split_ref(ref):
    m = re.match(r'([A-Z]+)(\d+)$', ref)
    return m.group(1), int(m.group(2))


class Book:
    def __init__(self, src):
        self.src = src
        z = zipfile.ZipFile(src)
        self.parts = {n: z.read(n) for n in z.namelist()}
        self.wb = openpyxl.load_workbook(src)
        wbx = self.text('xl/workbook.xml')
        rels = self.text('xl/_rels/workbook.xml.rels')
        relmap = {}
        for m in re.finditer(r'<Relationship\b[^>]*>', rels):
            rid = re.search(r'Id="([^"]+)"', m.group(0)).group(1)
            tgt = re.search(r'Target="([^"]+)"', m.group(0)).group(1)
            relmap[rid] = 'xl/' + re.sub(r'^/?xl/', '', tgt)
        self.sheets = {}
        for m in re.finditer(r'<sheet\b[^>]*>', wbx):
            name = re.search(r'name="([^"]+)"', m.group(0)).group(1).replace('&amp;', '&')
            rid = re.search(r'r:id="([^"]+)"', m.group(0)).group(1)
            self.sheets[name] = relmap[rid]
        self.log = []

    def text(self, part):
        return self.parts[part].decode('utf8')

    def put(self, part, text):
        self.parts[part] = text.encode('utf8')

    def sx(self, sheet):
        return self.text(self.sheets[sheet])

    def setsx(self, sheet, xml):
        self.put(self.sheets[sheet], xml)

    # ── 칸 고치기 ──
    def _replace_cell(self, sheet, ref, body, style_from=None, t=None, style=None):
        s = self.sx(sheet)
        m = CELL(ref).search(s)
        tattr = f' t="{t}"' if t else ''
        if m:
            st = re.search(r' s="(\d+)"', m.group(1))
            st = style if style is not None else (st.group(1) if st else None)
            new = f'<c r="{ref}"' + (f' s="{st}"' if st else '') + tattr + (f'>{body}</c>' if body else '/>')
            s = s[:m.start()] + new + s[m.end():]
        else:
            col, row = split_ref(ref)
            st = style
            if st is None and style_from:
                sm = CELL(style_from).search(s)
                sm = re.search(r' s="(\d+)"', sm.group(1)) if sm else None
                st = sm.group(1) if sm else None
            new = f'<c r="{ref}"' + (f' s="{st}"' if st else '') + tattr + (f'>{body}</c>' if body else '/>')
            rm = re.search(rf'<row r="{row}"[^>]*?(/>|>(.*?)</row>)', s, re.S)
            if not rm:
                # 줄이 없으면 만든다 (줄 번호 순서에 맞춰)
                rows = list(re.finditer(r'<row r="(\d+)"[^>]*?(?:/>|>.*?</row>)', s, re.S))
                pos = None
                for r in rows:
                    if int(r.group(1)) > row:
                        pos = r.start(); break
                if pos is None:
                    pos = s.index('</sheetData>')
                s = s[:pos] + f'<row r="{row}">{new}</row>' + s[pos:]
            elif rm.group(1) == '/>':
                head = rm.group(0)[:-2]
                s = s[:rm.start()] + head + '>' + new + '</row>' + s[rm.end():]
            else:
                pos = rm.start(2)
                for cm in re.finditer(r'<c r="([A-Z]+)\d+"(?:[^>]*?)(?:/>|>.*?</c>)', rm.group(2), re.S):
                    if CI(cm.group(1)) < CI(col):
                        pos = rm.start(2) + cm.end()
                s = s[:pos] + new + s[pos:]
        self.setsx(sheet, s)

    def setf(self, sheet, ref, formula, **kw):
        self._replace_cell(sheet, ref, f'<f>{escape(formula)}</f>', **kw)
        self.log.append((sheet, ref, '=' + formula))

    def setv(self, sheet, ref, value, **kw):
        if value is None:
            self._replace_cell(sheet, ref, '', **kw)
        elif isinstance(value, (int, float)):
            self._replace_cell(sheet, ref, f'<v>{value}</v>', **kw)
        else:
            self._replace_cell(sheet, ref, f'<is><t xml:space="preserve">{escape(str(value))}</t></is>', t='inlineStr', **kw)

    def clear(self, sheet, refs):
        for ref in expand(refs):
            if CELL(ref).search(self.sx(sheet)):
                self.setv(sheet, ref, None)

    def style_of(self, sheet, ref):
        m = CELL(ref).search(self.sx(sheet))
        m = re.search(r' s="(\d+)"', m.group(1)) if m else None
        return m.group(1) if m else None

    # 공유 수식을 칸마다의 수식으로 푼다 (openpyxl이 풀어 읽은 수식을 그대로 적는다)
    def unshare(self, sheet):
        ws = self.wb[sheet]
        s = self.sx(sheet)

        def fix(m):
            ref = m.group(1); v = ws[ref].value
            assert isinstance(v, str) and v.startswith('='), (sheet, ref, v)
            st = re.search(r' s="(\d+)"', m.group(2))
            return f'<c r="{ref}"' + (f' s="{st.group(1)}"' if st else '') + f'><f>{escape(v[1:])}</f></c>'

        s, n = re.subn(r'<c r="([A-Z]+\d+)"((?: [^>]*?)?)><f t="shared"[^>]*?(?:/>|>[^<]*</f>)(?:<v>[^<]*</v>|<v/>)?</c>', fix, s)
        assert 't="shared"' not in s, sheet
        self.setsx(sheet, s)
        return n

    # 모든 수식 칸의 저장값(<v>)을 지운다 (예시 값이 남지 않게). 엑셀은 열 때 다시 계산한다
    def drop_cached(self, sheet):
        s = self.sx(sheet)
        s = re.sub(r'(<c r="[A-Z]+\d+"(?: [^>]*?)?)(?: t="(?:str|e|b|n)")?((?: [^>]*?)?>)(<f>.*?</f>)(?:<v>[^<]*</v>|<v/>)', lambda m: re.sub(r' t="(?:str|e|b|n)"', '', m.group(1) + m.group(2)) + m.group(3), s, flags=re.S)
        self.setsx(sheet, s)

    def sub_formulas(self, sheet, pattern, repl):
        s = self.sx(sheet)
        count = [0]

        def fix(m):
            f = unescape(m.group(2))
            g = re.sub(pattern, repl, f)
            if g != f:
                count[0] += 1
            return m.group(1) + escape(g) + m.group(3)

        s = re.sub(r'(<f>)(.*?)(</f>)', fix, s, flags=re.S)
        self.setsx(sheet, s)
        return count[0]

    # ── 그림(붙여 넣은 사진·화면 갈무리) 지우기: 도형·차트는 둔다 ──
    def drop_pictures(self):
        removed = 0
        for part in [p for p in self.parts if re.match(r'xl/drawings/drawing\d+\.xml$', p)]:
            x = self.text(part)
            rels_part = part.replace('drawings/', 'drawings/_rels/') + '.rels'
            rels = self.text(rels_part) if rels_part in self.parts else ''

            def fix(m):
                nonlocal removed, rels
                if '<xdr:pic>' not in m.group(0):
                    return m.group(0)
                for rid in re.findall(r'r:embed="([^"]+)"', m.group(0)):
                    rm = re.search(rf'<Relationship\b[^>]*Id="{rid}"[^>]*/>', rels)
                    if rm:
                        tgt = re.search(r'Target="([^"]+)"', rm.group(0)).group(1)
                        media = 'xl/' + tgt.replace('../', '')
                        rels = rels.replace(rm.group(0), '')
                        if not any(media.split('/')[-1] in self.text(p) for p in self.parts if p.endswith('.rels') and p != rels_part) and media.split('/')[-1] not in rels:
                            self.parts.pop(media, None)
                removed += 1
                return ''

            x = re.sub(r'<xdr:(twoCellAnchor|oneCellAnchor|absoluteAnchor)\b[^>]*>.*?</xdr:\1>', fix, x, flags=re.S)
            self.put(part, x)
            if rels_part in self.parts:
                self.put(rels_part, rels)
        return removed

    # ── 인쇄 설정 ──
    def page_setup(self, sheet, landscape=False, fit_height=0, margins=None, center=True, print_area=None):
        s = self.sx(sheet)
        m = margins or dict(left=0.39, right=0.39, top=0.59, bottom=0.39, header=0.2, footer=0.2)
        pm = '<pageMargins ' + ' '.join(f'{k}="{v}"' for k, v in m.items()) + '/>'
        s = re.sub(r'<pageMargins[^>]*/>', pm, s) if '<pageMargins' in s else s
        ps = f'<pageSetup paperSize="9" fitToWidth="1" fitToHeight="{fit_height}" orientation="{"landscape" if landscape else "portrait"}"/>'
        if '<pageSetup' in s:
            s = re.sub(r'<pageSetup[^>]*/>', ps, s)
        else:
            s = s.replace(pm, pm + ps)
        # fitToPage
        if '<pageSetUpPr' in s:
            s = re.sub(r'<pageSetUpPr[^>]*/>', '<pageSetUpPr fitToPage="1"/>', s)
        elif '<sheetPr' in s:
            if re.search(r'<sheetPr[^>]*/>', s):
                s = re.sub(r'<sheetPr([^>]*)/>', r'<sheetPr\1><pageSetUpPr fitToPage="1"/></sheetPr>', s, count=1)
            else:
                s = s.replace('</sheetPr>', '<pageSetUpPr fitToPage="1"/></sheetPr>', 1)
        else:
            s = re.sub(r'(<worksheet[^>]*>)', r'\1<sheetPr><pageSetUpPr fitToPage="1"/></sheetPr>', s, count=1)
        po = '<printOptions horizontalCentered="1"/>' if center else ''
        s = re.sub(r'<printOptions[^>]*/>', '', s)
        if po:
            s = s.replace(pm, po + pm)
        s = re.sub(r'<rowBreaks.*?</rowBreaks>', '', s, flags=re.S)
        s = re.sub(r'<colBreaks.*?</colBreaks>', '', s, flags=re.S)
        self.setsx(sheet, s)
        # 프린터 설정 파일 연결은 지운다 (특정 프린터 값이 배율을 덮어쓰지 않게)
        rels_part = self.sheets[sheet].replace('worksheets/', 'worksheets/_rels/') + '.rels'
        if rels_part in self.parts:
            rels = self.text(rels_part)
            rels = re.sub(r'<Relationship\b[^>]*printerSettings[^>]*/>', '', rels)
            self.put(rels_part, rels)
        if print_area:
            self.print_area(sheet, print_area)

    def print_area(self, sheet, area):
        wbx = self.text('xl/workbook.xml')
        idx = list(self.sheets).index(sheet)
        quoted = "'" + sheet.replace("'", "''") + "'"
        a, b = area.split(':')
        ref = f'{quoted}!${split_ref(a)[0]}${split_ref(a)[1]}:${split_ref(b)[0]}${split_ref(b)[1]}'
        dn = f'<definedName name="_xlnm.Print_Area" localSheetId="{idx}">{escape(ref)}</definedName>'
        wbx = re.sub(rf'<definedName name="_xlnm.Print_Area" localSheetId="{idx}">[^<]*</definedName>', '', wbx)
        if '<definedNames>' in wbx:
            wbx = wbx.replace('<definedNames>', '<definedNames>' + dn, 1)
        elif '<definedNames/>' in wbx:
            wbx = wbx.replace('<definedNames/>', f'<definedNames>{dn}</definedNames>', 1)
        else:
            wbx = wbx.replace('</sheets>', f'</sheets><definedNames>{dn}</definedNames>', 1)
        self.put('xl/workbook.xml', wbx)

    def col_widths(self, sheet):
        s = self.sx(sheet)
        out = {}
        dflt = re.search(r'defaultColWidth="([\d.]+)"', s)
        for a, b, w in re.findall(r'<col min="(\d+)" max="(\d+)" width="([\d.]+)"', s):
            for c in range(int(a), min(int(b), 60) + 1):
                out[c] = float(w)
        return out, float(dflt.group(1)) if dflt else 8.43

    def scale_cols(self, sheet, first, last, factor):
        s = self.sx(sheet)
        widths, dflt = self.col_widths(sheet)
        cols = re.search(r'<cols>.*?</cols>', s, re.S)
        items = []
        attrs = {}
        for m in re.finditer(r'<col min="(\d+)" max="(\d+)" width="([\d.]+)"([^>]*)/>', cols.group(0)):
            for c in range(int(m.group(1)), min(int(m.group(2)), 16384) + 1):
                if c > 80:
                    break
                attrs[c] = m.group(4)
        tail = [m.group(0) for m in re.finditer(r'<col min="(\d+)" max="(\d+)"[^>]*/>', cols.group(0)) if int(m.group(2)) > 80]
        for c in sorted(set(list(attrs) + list(range(first, last + 1)))):
            w = widths.get(c, dflt)
            if first <= c <= last:
                w = round(w * factor, 4)
            a = attrs.get(c, '')
            if 'customWidth' not in a:
                a += ' customWidth="1"'
            items.append(f'<col min="{c}" max="{c}" width="{w}"{a}/>')
        fixed_tail = []
        for t in tail:
            mn = int(re.search(r'min="(\d+)"', t).group(1))
            fixed_tail.append(re.sub(r'min="\d+"', f'min="{max(mn, 81)}"', t))
        s = s.replace(cols.group(0), '<cols>' + ''.join(items) + ''.join(fixed_tail) + '</cols>')
        self.setsx(sheet, s)

    # 숫자 표시 형식을 가진 새 스타일을 만든다 (base 스타일을 복사)
    def style_with_format(self, base_style, code):
        st = self.text('xl/styles.xml')
        nf = re.search(r'<numFmts count="(\d+)">(.*?)</numFmts>', st, re.S)
        ids = [int(x) for x in re.findall(r'numFmtId="(\d+)"', nf.group(2))] if nf else []
        nid = max(ids + [199]) + 1
        item = f'<numFmt numFmtId="{nid}" formatCode="{escape(code, {chr(34): "&quot;"})}"/>'
        if nf:
            st = st.replace(nf.group(0), f'<numFmts count="{int(nf.group(1)) + 1}">{nf.group(2)}{item}</numFmts>')
        else:
            st = re.sub(r'(<styleSheet[^>]*>)', rf'\1<numFmts count="1">{item}</numFmts>', st, count=1)
        cx = re.search(r'<cellXfs count="(\d+)">(.*?)</cellXfs>', st, re.S)
        xfs = re.findall(r'<xf\b[^>]*?(?:/>|>.*?</xf>)', cx.group(2), re.S)
        base = xfs[int(base_style)]
        new = re.sub(r'numFmtId="\d+"', f'numFmtId="{nid}"', base, count=1)
        if 'applyNumberFormat' not in new:
            new = new.replace('<xf ', '<xf applyNumberFormat="1" ', 1)
        st = st.replace(cx.group(0), f'<cellXfs count="{int(cx.group(1)) + 1}">{cx.group(2)}{new}</cellXfs>')
        self.put('xl/styles.xml', st)
        return str(int(cx.group(1)))

    def finish(self, dst):
        # 계산 순서표·프린터 설정은 지우고, 열 때 다시 계산하게 한다
        self.parts.pop('xl/calcChain.xml', None)
        ct = self.text('[Content_Types].xml')
        ct = re.sub(r'<Override[^>]*calcChain[^>]*/>', '', ct)
        for p in [p for p in self.parts if p.startswith('xl/printerSettings/')]:
            self.parts.pop(p)
        rels = self.text('xl/_rels/workbook.xml.rels')
        rels = re.sub(r'<Relationship\b[^>]*calcChain[^>]*/>', '', rels)
        self.put('xl/_rels/workbook.xml.rels', rels)
        # 지운 그림 파일의 확장자 선언은 남아도 무방
        self.put('[Content_Types].xml', ct)
        wbx = self.text('xl/workbook.xml')
        if re.search(r'<calcPr\b', wbx):
            wbx = re.sub(r'<calcPr\b[^>]*/>', '<calcPr calcId="191029" fullCalcOnLoad="1"/>', wbx)
        else:
            wbx = wbx.replace('</workbook>', '<calcPr calcId="191029" fullCalcOnLoad="1"/></workbook>')
        self.put('xl/workbook.xml', wbx)
        # 문서 정보에서 작성자 이름을 지운다
        if 'docProps/core.xml' in self.parts:
            core = self.text('docProps/core.xml')
            core = re.sub(r'<dc:creator>[^<]*</dc:creator>', '<dc:creator></dc:creator>', core)
            core = re.sub(r'<cp:lastModifiedBy>[^<]*</cp:lastModifiedBy>', '<cp:lastModifiedBy></cp:lastModifiedBy>', core)
            self.put('docProps/core.xml', core)
        # 어디에서도 쓰지 않는 그림 파일은 뺀다
        used = ' '.join(self.text(p) for p in self.parts if p.endswith('.rels'))
        for p in [p for p in self.parts if p.startswith('xl/media/')]:
            if p.split('/')[-1] not in used:
                self.parts.pop(p)
        with zipfile.ZipFile(dst, 'w', zipfile.ZIP_DEFLATED) as z:
            for name, data in self.parts.items():
                z.writestr(name, data)


def unescape(s):
    return s.replace('&lt;', '<').replace('&gt;', '>').replace('&quot;', '"').replace('&apos;', "'").replace('&amp;', '&')


def expand(refs):
    if isinstance(refs, str):
        refs = [refs]
    out = []
    for r in refs:
        for part in r.split(','):
            part = part.strip()
            if ':' in part:
                a, b = part.split(':')
                (c1, r1), (c2, r2) = split_ref(a), split_ref(b)
                for rr in range(r1, r2 + 1):
                    for cc in range(CI(c1), CI(c2) + 1):
                        out.append(f'{CL(cc)}{rr}')
            else:
                out.append(part)
    return out
