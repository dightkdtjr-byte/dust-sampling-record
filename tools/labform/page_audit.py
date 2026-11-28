# 인쇄 쪽 검사: 모든 시트(기록부·계산과정·raw data)에서
#  (가) 표 묶음이 두 쪽으로 갈라지는 곳 — 기록부는 번호 매긴 항목(1, 2, 3 …)이 한 묶음, 그 밖의 시트는 빈 줄로 나뉜 표·차트·그림이 한 묶음
#  (나) 쪽 채움 — 쪽마다 인쇄 내용이 쓸 수 있는 높이·너비의 몇 %를 채우는지
#  쪽 나눔은 시트 설정(배율·너비 맞춤·수동 쪽 나눔)과 줄 높이·열 너비로 계산하고, LibreOffice로 찍은 PDF 쪽수와 견줘 계산이 맞는지 확인한다
#   python3 -I tools/labform/page_audit.py <기록부 폴더> [PDF 폴더]
import sys, os, re, glob, zipfile, subprocess, warnings
warnings.filterwarnings('ignore')
import openpyxl
from openpyxl.utils import range_boundaries, get_column_letter as CL

MDW = 8
px = lambda w: int(w * MDW + 5)
norm = lambda v: re.sub(r'\s+', '', str(v or ''))
PAPER = {9: (8.27, 11.69)}


def sheet_xml(z, wbx, rels, name):
    for m in re.finditer(r'<sheet [^>]*name="([^"]+)"[^>]*r:id="([^"]+)"', wbx):
        if m.group(1).replace('&amp;', '&') == name:
            t = re.search(rf'Id="{m.group(2)}"[^>]*Target="([^"]+)"', rels) or re.search(rf'Target="([^"]+)"[^>]*Id="{m.group(2)}"', rels)
            p = 'xl/' + t.group(1).lstrip('/').replace('xl/', '')
            return p, z.read(p).decode('utf8')
    return None, ''


def anchors(z, sp, x):
    """그림·차트가 걸친 (첫 줄, 끝 줄) 1부터"""
    d = re.search(r'<drawing r:id="([^"]+)"', x)
    if not d:
        return []
    rp = sp.replace('worksheets/', 'worksheets/_rels/') + '.rels'
    rels = z.read(rp).decode()
    t = re.search(rf'Id="{d.group(1)}"[^>]*Target="([^"]+)"', rels) or re.search(rf'Target="([^"]+)"[^>]*Id="{d.group(1)}"', rels)
    dx = z.read('xl/drawings/' + t.group(1).split('/')[-1]).decode()
    out = []
    for a in re.finditer(r'<xdr:(twoCellAnchor|oneCellAnchor)\b[^>]*>(.*?)</xdr:\1>', dx, re.S):
        fr = re.search(r'<xdr:from><xdr:col>(\d+)</xdr:col>.*?<xdr:row>(\d+)</xdr:row>', a.group(2), re.S)
        to = re.search(r'<xdr:to><xdr:col>(\d+)</xdr:col>.*?<xdr:row>(\d+)</xdr:row>', a.group(2), re.S)
        ext = re.search(r'<xdr:ext cx="(\d+)" cy="(\d+)"', a.group(2))
        r1, c1 = int(fr.group(2)) + 1, int(fr.group(1)) + 1
        if to:
            r2, c2 = int(to.group(2)) + 1, int(to.group(1)) + 1
        else:
            r2, c2 = r1, c1      # oneCellAnchor: 높이는 따로 (ext)
        out.append((r1, r2, c1, c2, int(ext.group(2)) / 12700 if ext and not to else None))
    return out


def analyse(path):
    wb = openpyxl.load_workbook(path)
    z = zipfile.ZipFile(path); wbx = z.read('xl/workbook.xml').decode(); rels = z.read('xl/_rels/workbook.xml.rels').decode()
    res = []
    for ws in wb.worksheets:
        if ws.sheet_state != 'visible':
            continue
        from raw_layout import is_raw_paste
        if is_raw_paste(ws):          # 기기 자료 붙이는 시트: 원본 그대로 두므로 검사하지 않음
            continue
        sp, x = sheet_xml(z, wbx, rels, ws.title)
        form = '시험기록부' in norm(ws['A1'].value)
        pa = ws.print_area
        pa = pa[0] if isinstance(pa, (list, tuple)) else pa
        if pa:
            c1, r1, c2, r2 = range_boundaries(pa.split('!')[-1].replace('$', ''))
        else:
            cells = [(c.row, c.column) for row in ws.iter_rows() for c in row if c.value not in (None, '')]
            anc = anchors(z, sp, x)
            if not cells and not anc:
                continue
            rows_ = [r for r, _ in cells] + [a[0] for a in anc] + [a[1] for a in anc]
            cols_ = [c for _, c in cells] + [a[2] for a in anc] + [a[3] for a in anc]
            r1, r2, c1, c2 = 1, max(rows_), 1, max(cols_)
        # 줄 높이·열 너비
        dflt = re.search(r'<sheetFormatPr[^>]*defaultRowHeight="([\d.]+)"', x)
        dflt = float(dflt.group(1)) if dflt else 15.0
        hh = {}
        for r in range(r1, r2 + 1):
            d = ws.row_dimensions.get(r) if hasattr(ws.row_dimensions, 'get') else ws.row_dimensions[r]
            hh[r] = 0 if (d is not None and d.hidden) else (d.height if d is not None and d.height else dflt)
        sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
        from raw_layout import col_widths, _px
        CW, dw = col_widths(x)
        wpx = sum(0 if CW.get(c, (dw, False))[1] else _px(CW.get(c, (dw, False))[0]) for c in range(c1, c2 + 1))
        W = wpx * 0.75
        # 쪽 설정
        ps = re.search(r'<pageSetup\b[^>]*/>', x); ps = ps.group(0) if ps else ''
        land = 'orientation="landscape"' in ps
        mg = re.search(r'<pageMargins left="([\d.]+)" right="([\d.]+)" top="([\d.]+)" bottom="([\d.]+)"', x)
        L, R, T, B = (float(v) for v in mg.groups()) if mg else (0.7, 0.7, 0.75, 0.75)
        pw, ph = PAPER[9]
        if land:
            pw, ph = ph, pw
        uw, uh = (pw - L - R) * 72, (ph - T - B) * 72
        fit = re.search(r'<pageSetUpPr[^>]*fitToPage="1"', x) is not None
        fw = int((re.search(r'fitToWidth="(\d+)"', ps) or [0, 1])[1]); fh = int((re.search(r'fitToHeight="(\d+)"', ps) or [0, 1])[1])
        sc = re.search(r'scale="(\d+)"', ps)
        H = sum(hh.values())
        if fit:
            s = min(1.0, (uw * max(fw, 1)) / W if fw else 1.0, (uh * fh) / H if fh else 1.0)
        else:
            s = int(sc.group(1)) / 100 if sc else 1.0
        brk = sorted(int(b) for b in re.findall(r'<brk id="(\d+)"[^>]*man="1"', (re.search(r'<rowBreaks.*?</rowBreaks>', x, re.S) or [''])[0])) if '<rowBreaks' in x else []
        # 쪽 나누기: 수동 나눔 뒤, 넘치면 자동 (넘치는 줄 앞에서)
        pages = []; start = r1; acc = 0.0
        for r in range(r1, r2 + 1):
            if r - 1 in brk and r > start:
                pages.append((start, r - 1)); start = r; acc = 0.0
            if acc + hh[r] * s > uh + 0.5 and r > start:
                pages.append((start, r - 1)); start = r; acc = 0.0
            acc += hh[r] * s
        pages.append((start, r2))
        # 묶음
        anc = anchors(z, sp, x)
        units = []
        if form:
            # 묶음 = 번호 항목 안의 표: 항목 시작 줄, 또는 회색 머리 줄(글자 칸 3개 이상, 60 % 이상이 회색)로 새 표가 시작
            sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
            from layout_rules import _gray as gray      # 양식 정리와 같은 회색 판정 (테마 색 포함)
            def head(r):
                cs = [ws.cell(r, c) for c in range(c1, c2 + 1)]
                filled = [c for c in cs if type(c).__name__ != 'MergedCell' and c.value not in (None, '')]
                g = [c for c in filled if gray(c)]
                return len(g) >= 3 and len(g) >= 0.6 * len(filled)
            nums = [r for r in range(r1, r2 + 1) if re.fullmatch(r'\d{1,2}', str(ws.cell(r, c1).value or '').strip())]
            starts = sorted(set(nums) | {r for r in range(r1, r2 + 1) if r not in nums and any(n < r for n in nums) and head(r) and not (r > r1 and head(r - 1))})
            for i, a in enumerate(starts):
                sec = max(n for n in nums if n <= a)
                units.append((a, (starts[i + 1] - 1) if i + 1 < len(starts) else r2, f'{ws.cell(sec, c1).value}번 {norm(ws.cell(sec, c1 + 1).value)[:8]}' + ('' if a == sec else f' {a}행 표')))
        else:
            filled = set()
            for row in ws.iter_rows(min_row=r1, max_row=r2, min_col=c1, max_col=c2):
                for c in row:
                    if c.value not in (None, '') or (c.border and any(getattr(c.border, sd).style for sd in ('top', 'bottom', 'left', 'right'))):
                        filled.add(c.row)
            for a in anc:
                a2 = a[1]
                if a[4]:
                    rest = a[4]; a2 = a[0]
                    while rest > 0 and a2 <= r2:
                        rest -= hh.get(a2, dflt); a2 += 1
                filled.update(range(a[0], a2 + 1))
            r = r1
            while r <= r2:
                if r in filled:
                    a = r
                    while r + 1 <= r2 and r + 1 in filled:
                        r += 1
                    units.append((a, r, f'{a}~{r}행'))
                r += 1
        splits = []
        for a, e, nm in units:
            pg = [i for i, (p1, p2) in enumerate(pages) if not (p2 < a or p1 > e)]
            if len(pg) > 1:
                need = sum(hh[r] for r in range(a, e + 1)) * s
                splits.append((nm, len(pg), need > uh))       # 한 쪽보다 길면 어쩔 수 없음
        fill = [round(sum(hh[r] for r in range(p1, p2 + 1)) * s / uh * 100) for p1, p2 in pages]
        cut = [str(g) for g in ws.merged_cells.ranges if not (g.max_row < r1 or g.min_row > r2 or g.max_col < c1 or g.min_col > c2)
               and (g.max_col > c2 or g.max_row > r2 or g.min_col < c1 or g.min_row < r1)]
        if not form:     # 기록부는 병합이 영역 안에 있게 따로 맞춤; 그 밖의 시트도 같은 기준
            pass
        res.append(dict(sheet=ws.title, cut=cut, form=form, pages=len(pages), scale=round(s * 100), wfill=round(min(1, W * s / (uw * max(fw, 1))) * 100),
                        hfill=fill, splits=splits, land=land))
    return res


if __name__ == '__main__':
    # python3 -I page_audit.py <기록부 폴더> [pages.json]  — pages.json(sheet_pages.py 결과)이 있으면 시트마다 실제 PDF 쪽수와 견줌
    import json
    sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
    G = sys.argv[1]; J = json.load(open(sys.argv[2])) if len(sys.argv) > 2 else None
    tot_split = 0; mism = []; nsheet = 0
    for p in sorted(glob.glob(G + '/*.xls[xm]')):
        res = analyse(p)
        real_pages = (J or {}).get(os.path.basename(p), {})
        for r in res:
            nsheet += 1
            if J is not None and real_pages.get(r['sheet']) != r['pages']:
                mism.append(f"{os.path.basename(p)} {r['sheet']}: 계산 {r['pages']}쪽, PDF {real_pages.get(r['sheet'])}쪽")
            real = [s for s in r['splits'] if not s[2]] + [('병합 칸 잘림 ' + c, 0, False) for c in r['cut']]
            tot_split += len(real)
            flag = '✖' if real else ' '
            print(f" {flag} {os.path.basename(p)[6:16]:10} {r['sheet'][:16]:16} {'기록부' if r['form'] else '      '} {'가로' if r['land'] else '세로'} {r['pages']}쪽 배율 {r['scale']}% 너비 {r['wfill']}% 높이 {r['hfill']}"
                  + (f"  갈라짐 {[s[0] for s in real]}" if real else '') + (f"  (한 쪽보다 긴 묶음 {[s[0] for s in r['splits'] if s[2]]})" if any(s[2] for s in r['splits']) else ''))
    for m in mism:
        print('   ! 쪽수 다름', m)
    print(f'인쇄 쪽 검사 (기기 자료 시트 뺀 {nsheet}개 시트): 갈라진 표 묶음·잘린 병합 칸 {tot_split}곳' + (f', PDF 쪽수와 계산이 다른 시트 {len(mism)}개' if J is not None else ''))
