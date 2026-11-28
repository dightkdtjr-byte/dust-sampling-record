# 검증 방식 20: 계산식 그림 — 계산식 칸이 있는 기록부 시트마다
#  ① 식 그림이 있고 계산식 칸(항목 줄 범위·내용 열) 안에 들어가는가  ② 기호 설명 그림이 있고 칸 안인가
#  ③ 그림끼리 겹치지 않는가  ④ 원래 엑셀 수식 개체(OMML)가 남아 있지 않은가  ⑤ 그림이 많이 줄지 않았는가(원래 크기의 90 % 이상)
#  ⑥ 그림 파일이 eq/ 의 원본과 같은가 (픽셀 크기)
#   python3 -I tools/labform/eq_audit.py <기록부 폴더>
import sys, os, re, glob, zipfile, json, warnings
warnings.filterwarnings('ignore')
here = os.path.dirname(os.path.abspath(__file__)); sys.path.insert(0, here)
import openpyxl
from openpyxl.utils import get_column_letter as CL
from equations import EQ_DIR, MDW


def main(G):
    sizes = json.load(open(os.path.join(EQ_DIR, 'sizes.json'))); sizes.update(json.load(open(os.path.join(EQ_DIR, 'legend', 'sizes.json'))))
    bad = n = 0
    for p in sorted(glob.glob(G + '/*.xls[xm]')):
        wb = openpyxl.load_workbook(p); z = zipfile.ZipFile(p)
        wbx = z.read('xl/workbook.xml').decode(); rels = z.read('xl/_rels/workbook.xml.rels').decode()
        for ws in wb.worksheets:
            if '시험기록부' not in re.sub(r'\s', '', str(ws['A1'].value or '')):
                continue
            nums = [r for r in range(1, ws.max_row + 1) if re.fullmatch(r'\d{1,2}', str(ws.cell(r, 1).value or '').strip())]
            sec = [r for r in nums if '계산식' in str(ws.cell(r, 2).value or '')]
            if not sec:
                continue
            n += 1
            a = sec[0]; e = min([q for q in nums if q > a] + [ws.max_row + 1]) - 1
            pa = ws.print_area; pa = pa[0] if isinstance(pa, (list, tuple)) else pa
            c2 = openpyxl.utils.range_boundaries(pa.split('!')[-1].replace('$', ''))[2]
            X = lambda c: sum(int((ws.column_dimensions[CL(k)].width or 9) * MDW) for k in range(1, c))
            Y = lambda r: sum((ws.row_dimensions[k].height or 15) for k in range(1, r)) * 96 / 72
            rid = re.search(r'<sheet [^>]*name="%s"[^>]*r:id="([^"]+)"' % re.escape(ws.title.replace('&', '&amp;')), wbx).group(1)
            sp = 'xl/' + re.search(r'Id="%s"[^>]*Target="([^"]+)"' % rid, rels).group(1).replace('/xl/', '')
            x = z.read(sp).decode(); d = re.search(r'<drawing r:id="([^"]+)"', x)
            probs = []
            pics = []
            if d:
                srel = z.read(sp.replace('worksheets/', 'worksheets/_rels/') + '.rels').decode()
                dp = 'xl/drawings/' + re.search(r'Id="%s"[^>]*Target="([^"]+)"' % d.group(1), srel).group(1).split('/')[-1]
                dx = z.read(dp).decode(); drel = z.read(dp.replace('drawings/', 'drawings/_rels/') + '.rels').decode()
                for m in re.finditer(r'<xdr:(twoCellAnchor|oneCellAnchor)\b[^>]*>(.*?)</xdr:\1>', dx, re.S):
                    body = m.group(2)
                    fr = re.search(r'<xdr:from><xdr:col>(\d+)</xdr:col><xdr:colOff>(-?\d+)</xdr:colOff><xdr:row>(\d+)</xdr:row><xdr:rowOff>(-?\d+)</xdr:rowOff>', body)
                    row = int(fr.group(3)) + 1
                    if 'm:oMath' in body and a <= row <= e:
                        probs.append('④ 엑셀 수식 개체가 남음')
                    nm = re.search(r'<xdr:cNvPr id="\d+" name="([^"]+)"', body)
                    if nm and (nm.group(1).startswith('계산식 ') or nm.group(1).startswith('기호 설명 ')):
                        ext = re.search(r'<xdr:ext cx="(\d+)" cy="(\d+)"', body)
                        x0 = X(int(fr.group(1)) + 1) + int(fr.group(2)) / 9525; y0 = Y(row) + int(fr.group(4)) / 9525
                        w, h = int(ext.group(1)) / 9525, int(ext.group(2)) / 9525
                        key = nm.group(1).split(' ')[-1]
                        emb = re.search(r'r:embed="([^"]+)"', body).group(1)
                        tgt = re.search(r'Id="%s"[^>]*Target="../media/([^"]+)"' % emb, drel).group(1)
                        from PIL import Image
                        import io
                        im = Image.open(io.BytesIO(z.read('xl/media/' + tgt)))
                        sw, sh_ = sizes[key]
                        if abs(im.width - sw * 4) > 4 or abs(im.height - sh_ * 4) > 4:
                            probs.append(f'⑥ {key} 그림 파일 크기 다름')
                        if w / sw < 0.9:
                            probs.append(f'⑤ {key} {w / sw:.0%}로 줄어듦')
                        pics.append((nm.group(1), x0, y0, x0 + w, y0 + h))
            eqs = [p_ for p_ in pics if p_[0].startswith('계산식')]; lgs = [p_ for p_ in pics if p_[0].startswith('기호')]
            if not eqs:
                probs.append('① 식 그림 없음')
            if not lgs:
                probs.append('② 기호 설명 그림 없음')
            sx0, sy0, sx1, sy1 = X(4), Y(a), X(c2 + 1), Y(e + 1)
            for nm, x0, y0, x1, y1 in pics:
                if x0 < sx0 - 1 or y0 < sy0 - 1 or x1 > sx1 + 1 or y1 > sy1 + 1:
                    probs.append(f'①② {nm} 계산식 칸({a}~{e}행) 밖으로 나감')
            for i in range(len(pics)):
                for j in range(i + 1, len(pics)):
                    p1, p2 = pics[i], pics[j]
                    if p1[1] < p2[3] - 1 and p2[1] < p1[3] - 1 and p1[2] < p2[4] - 1 and p2[2] < p1[4] - 1:
                        probs.append(f'③ {p1[0]}·{p2[0]} 겹침')
            if probs:
                bad += 1
                print(f'✖ {os.path.basename(p)} {ws.title}: ' + '; '.join(probs))
    print(f'계산식 그림 검사: {n - bad}/{n} 시트 통과')


if __name__ == '__main__':
    main(sys.argv[1])
