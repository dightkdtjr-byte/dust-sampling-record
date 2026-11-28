# 기기 자료 붙이는 시트(수식이 하나도 없는 시트)가 원본 그대로인지 검사
#  칸 위치·값(익명화한 글자만 예외)·칸 서식(글꼴·숫자 형식·맞춤·테두리·채우기)·열 너비·줄 높이·병합·그림 위치·인쇄 설정을 원본과 견준다.
#  기기에서 내보낸 자료를 그대로 붙여 넣는 곳이라, 하나라도 달라지면 붙여 넣은 뒤 계산과정 수식이 엉뚱한 칸을 읽을 수 있다.
#   python3 -I tools/labform/raw_untouched.py <원본 폴더> <결과 폴더>
import sys, os, re, glob, zipfile, warnings
warnings.filterwarnings('ignore')
here = os.path.dirname(os.path.abspath(__file__)); sys.path.insert(0, here)
from anonymize import fix_text
from raw_layout import is_raw_paste
import openpyxl

ITEM = lambda p: re.search(r'시험기록부_([^_]+?)(?:_|\.xls)', os.path.basename(p)).group(1)


def sheet_xml(z, name):
    wbx = z.read('xl/workbook.xml').decode(); rels = z.read('xl/_rels/workbook.xml.rels').decode()
    for m in re.finditer(r'<sheet [^>]*name="([^"]+)"[^>]*r:id="([^"]+)"', wbx):
        if m.group(1).replace('&amp;', '&') == name:
            t = re.search(rf'Id="{m.group(2)}"[^>]*Target="([^"]+)"', rels) or re.search(rf'Target="([^"]+)"[^>]*Id="{m.group(2)}"', rels)
            p = 'xl/' + t.group(1).lstrip('/').replace('xl/', '')
            x = z.read(p).decode()
            d = re.search(r'<drawing r:id="([^"]+)"', x); dx = ''
            if d:
                rp = p.replace('worksheets/', 'worksheets/_rels/') + '.rels'
                r = z.read(rp).decode()
                t2 = re.search(rf'Id="{d.group(1)}"[^>]*Target="([^"]+)"', r) or re.search(rf'Target="([^"]+)"[^>]*Id="{d.group(1)}"', r)
                dx = z.read('xl/drawings/' + t2.group(1).split('/')[-1]).decode()
            return x, dx
    return '', ''


def fmt(c):
    """칸의 실제 서식 (서식 번호가 아니라 내용: 글꼴·숫자 형식·맞춤·테두리·채우기·보호)"""
    if type(c).__name__ == 'MergedCell':      # 병합으로 가려진 칸: 보이지 않음 (openpyxl은 통합 문서 기본 서식을 돌려줌)
        return None
    f, a, bd, fl = c.font, c.alignment, c.border, c.fill
    return (c.number_format, f.name, f.sz, f.b, f.i, f.u, f.vertAlign, str(f.color.rgb if f.color is not None else None),
            a.horizontal, a.vertical, a.wrap_text, a.shrink_to_fit, a.indent, a.text_rotation,
            tuple(getattr(bd, sd).style for sd in ('left', 'right', 'top', 'bottom')), fl.fill_type, str(fl.fgColor.rgb), c.protection.locked)


def shape(x, dx):
    """비교할 모양: 열 너비, 줄 높이·숨김, 병합, 인쇄 설정, 그림 자리"""
    cols = re.findall(r'<col\b[^>]*/>', x)
    rows = [(m.group(1), (re.search(r'\sht="([^"]+)"', m.group(2)) or [0, ''])[1], 'hidden="1"' in m.group(2)) for m in re.finditer(r'<row r="(\d+)"([^>]*)>', x)]
    merges = sorted(re.findall(r'<mergeCell ref="([^"]+)"', x))
    page = [re.sub(r'\s+r:id="[^"]*"', '', t) for t in re.findall(r'<(?:pageSetup|pageMargins|printOptions|pageSetUpPr|sheetFormatPr|rowBreaks|colBreaks)\b[^>]*>', x)]
    anchors = re.findall(r'<xdr:from>.*?</xdr:from>(?:<xdr:to>.*?</xdr:to>)?|<xdr:ext cx="\d+" cy="\d+"/>', dx, re.S)
    return dict(열너비=cols, 줄=rows, 병합=merges, 인쇄설정=page, 그림=anchors)


def main(src, out):
    origs = {ITEM(p): p for p in glob.glob(os.path.join(src, '**', '*.xls*'), recursive=True) if not os.path.basename(p).startswith('~$')}
    bad = 0; n = 0
    for q in sorted(glob.glob(os.path.join(out, '*.xls[xm]'))):
        p = origs.get(ITEM(q))
        if not p:
            continue
        wo, wq = openpyxl.load_workbook(p), openpyxl.load_workbook(q)
        zo, zq = zipfile.ZipFile(p), zipfile.ZipFile(q)
        for ws in wo.worksheets:
            if ws.sheet_state != 'visible' or not is_raw_paste(ws) or '시험기록부' in re.sub(r'\s', '', str(ws['A1'].value or '')):
                continue
            n += 1
            if ws.title not in wq.sheetnames:
                print(f'✖ {os.path.basename(q)} {ws.title}: 시트가 없어짐'); bad += 1; continue
            wsq = wq[ws.title]; diff = []
            co = {c.coordinate: (fix_text(c.value) if isinstance(c.value, str) else c.value, fmt(c)) for row in ws.iter_rows() for c in row if c.value is not None or c.has_style}
            cq = {c.coordinate: (c.value, fmt(c)) for row in wsq.iter_rows() for c in row if c.value is not None or c.has_style}
            dv = [k for k in set(co) | set(cq) if (co.get(k) or (None,))[0] != (cq.get(k) or (None,))[0]]
            if dv:
                diff.append(f'값 다른 칸 {len(dv)}개 {sorted(dv)[:5]}')
            ds = [k for k in set(co) & set(cq) if co[k][1] != cq[k][1]]
            if ds:
                diff.append(f'서식 다른 칸 {len(ds)}개 {sorted(ds)[:5]}')
            so, sq = shape(*sheet_xml(zo, ws.title)), shape(*sheet_xml(zq, ws.title))
            for k in so:
                if so[k] != sq[k]:
                    diff.append(f'{k} 다름')
            if diff:
                bad += 1
                print(f'✖ {os.path.basename(q)} {ws.title}: ' + '; '.join(diff))
            else:
                print(f'  {os.path.basename(q)} {ws.title}: 원본 그대로')
    print(f'기기 자료 시트 원본 유지: {n - bad}/{n}')


if __name__ == '__main__':
    main(sys.argv[1], sys.argv[2])
