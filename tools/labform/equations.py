# 기록부 "계산식" 칸의 식을 한 모양으로: KaTeX로 조판한 그림(eq/*.png)을 원래 식 자리에 넣는다
#  원래 식은 기록부마다 엑셀 수식 개체(글꼴·크기 제각각, 비소는 상자에 안 들어가 잘림), 칸·테두리로 만든 분수, 글자로 섞여 있었다.
#  식 내용은 원본 그대로 옮겼다 (원본 수식 개체 글자·칸 글자). 변수 설명 칸은 그대로 둔다.
#   python3 -I tools/labform/equations.py render   → eq/*.png 다시 만들기 (node·playwright·katex 필요)
import os, re, json, sys, hashlib
from openpyxl.utils import get_column_letter as CL

HERE = os.path.dirname(os.path.abspath(__file__))
EQ_DIR = os.path.join(HERE, 'eq')
T = lambda s: r'\text{' + s + '}'

# 식 (KaTeX). 한글은 \text{}
TEX = {
    'cs2': r'C = \mathrm{Average}\left[\,(C_n \times f_n) - (C_b \times f_b)\,\right]',
    'vocs': r'C = \dfrac{\mathrm{Average}\left[\,(C_n \times f_n) - (C_b \times f_b)\,\right]}{1000}',
    'sox': r'C = C_a \times \dfrac{21 - O_s}{21 - O_a}',
    'dmds': r'x_i = \dfrac{y_i - b}{a}',
    'pahs': r'C = \dfrac{A_s}{A_{\mathrm{is}}} \times \dfrac{I_{\mathrm{is}}}{\mathit{RRF}_{\mathrm{avg}}} \times \dfrac{1}{V_{\mathrm{m(std)}}}',
    'dust': T('먼지 농도 (mg/Sm') + r'^3' + T(')') + r'\ = \dfrac{m_d}{V_m \times \dfrac{P_a + \Delta H / 13.6}{760} \times \dfrac{273}{273 + \theta_m}}',
    'br': T('브로민화합물 농도 (ppm)') + r'\ = \dfrac{(a - b) \times 100}{V_s} \times \dfrac{22.4}{79.904}',
    'hcl': T('염화수소 농도 (ppm)') + r'\ = \dfrac{(a - b) \times 100}{V_s} \times \dfrac{22.4}{35.453}',
    'hf': T('플루오린화합물 농도 (ppm)') + r'\ = \dfrac{(a - b) \times V}{V_s} \times \dfrac{22.4}{18.998}',
    'as': r'C = \dfrac{m_1 + m_2}{V_s} \times \dfrac{22.41}{74.92}',
    'hcn': r'C = \dfrac{(a - b) \times 10 \times 22.4}{V_s \times 26.017}',
    'nh3': r'C = \dfrac{(a - b) \times 25}{V_s}',
    'h2s': r'C = \dfrac{(a - b) \times 0.698 \times V\!/20 \times 1000}{V_s}',
    'phenol': T('페놀화합물 농도 (ppm)') + r'\ = \dfrac{(a - b) \times 20 \times f}{V_s}',
    'ald': T('알데하이드 농도 (ppm)') + r'\ = \dfrac{(A_a - A_b) \times V}{V_m \times \dfrac{P_a}{760} \times \dfrac{273}{273 + T_a}} \times \dfrac{22.4}{\mathit{MW}_a} \times B',
    'hg': T('수은화합물 농도 (mg/Sm') + r'^3' + T(')') + r'\ = \dfrac{m_{\mathrm{(HCl)Hg}} + m_{\mathrm{(sol)Hg}}}{V_s} \times 10^{-3}',
    'hg_hcl': r'm_{\mathrm{(HCl)Hg}} = \left[\,\{C_{\mathrm{(HCl)Hg}} \times V_{\mathrm{f(HCl)}}\} - \{C_{\mathrm{(HCl\,bk)Hg}} \times V_{\mathrm{f(HCl\,bk)}}\}\,\right]',
    'hg_sol': r'm_{\mathrm{(sol)Hg}} = \left[\,\{C_{\mathrm{(sol)Hg}} \times V_{\mathrm{f(sol)}}\} - \{C_{\mathrm{(sol\,bk)Hg}} \times V_{\mathrm{f(sol\,bk)}}\}\,\right]',
    'thc': r'C_c\ \text{(ppm)} = K \times C_{\text{측정}}',
    'metal': r'C = \dfrac{(a - b) \times V}{V_s}',
}

# 기록부마다: (파일 이름에 든 말, 시트 이름 규칙) → [(식, 첫 줄, 끝 줄, 첫 열, 끝 열, 지울 칸)]
#   끝 열이 None이면 인쇄 영역 끝 열. 영역 안 수식 개체는 지우고, 지울 칸의 글자와 영역 안쪽 테두리(칸 분수의 가로선)를 지운다
SPECS = [
    ('_CS2', r'시험기록부', [('cs2', 21, 24, 4, 8, [])]),
    ('_이황화메틸', r'시험기록부', [('dmds', 21, 24, 4, 5, [])]),
    ('_VOCs', r'.*', [('vocs', 21, 24, 4, 8, [])]),
    ('_SOx', r'시험기록부', [('sox', 13, 16, 4, 8, [])]),
    ('_PAHs', r'시험기록부', [('pahs', 12, 13, 4, None, ['D12'])]),
    ('_먼지', r'시험기록부', [('dust', 11, 12, 4, None, ['D11'])]),
    ('_브로민', r'시험기록부', [('br', 12, 13, 4, None, ['D12'])]),
    ('_브롬', r'시험기록부', [('br', 12, 13, 4, None, ['D12'])]),
    ('_염화수소', r'시험기록부', [('hcl', 12, 13, 4, None, ['D12'])]),
    ('_플루오린', r'시험기록부', [('hf', 12, 13, 4, None, ['D12'])]),
    ('_불소', r'시험기록부', [('hf', 12, 13, 4, None, ['D12'])]),
    ('_비소', r'시험기록부', [('as', 12, 13, 4, None, [])]),
    ('_사이안화수소', r'시험기록부', [('hcn', 12, 13, 4, None, ['D12', 'E12', 'F12', 'F13'])]),
    ('_시안화수소', r'시험기록부', [('hcn', 12, 13, 4, None, ['D12', 'E12', 'F12', 'F13'])]),
    ('_암모니아', r'시험기록부', [('nh3', 12, 13, 4, None, ['D12', 'E12', 'F12', 'F13'])]),
    ('_황화수소', r'시험기록부', [('h2s', 12, 13, 4, None, ['D12', 'E12', 'F12', 'F13'])]),
    ('_페놀', r'시험기록부', [('phenol', 12, 13, 4, None, ['D12', 'F12', 'F13'])]),
    ('_아세트알데하이드', r'시험기록부', [('ald', 12, 13, 4, None, ['D12'])]),
    ('_폼알데하이드', r'시험기록부', [('ald', 12, 13, 4, None, ['D12'])]),
    ('_포름알데하이드', r'시험기록부', [('ald', 12, 13, 4, None, ['D12'])]),
    ('_수은', r'시험기록부', [('hg', 12, 13, 4, None, ['D12', 'F12', 'H12', 'I12', 'F13']),
                           ('hg_hcl', 15, 15, 4, None, ['D15']), ('hg_sol', 17, 17, 4, None, ['D17'])]),
    ('_총탄화수소', r'시험기록부', [('thc', 11, 11, 4, None, ['D11', 'E11'])]),
    ('_중금속', r'^(구리|납|아연|크롬|크로뮴|니켈|카드뮴|베릴륨)화합물$', [('metal', 12, 13, 4, None, [], 14)]),   # 원래 식은 14행 설명 왼쪽(D14)에 작게, 12·13행은 빈 줄
]
# 같은 항목 안에서 줄 높이 옮기기 (항목 전체 높이는 그대로라 쪽 나눔은 같음): 중금속은 식 줄(12·13행)이 20pt라 식이 테두리에 붙음
ROW_MOVE = {'_중금속': {12: 30.0, 13: 30.0, 14: None}}     # None = 나머지
PAD_X, PAD_Y = 14, 4          # 영역 왼쪽·위아래 여백 (화면 점)
MDW = 8                        # 열 너비 한 글자 = 8점 (맑은 고딕 11)


def _shrink(folder):
    """검은 글자·흰 바탕이라 회색 16단계로 줄여도 같아 보인다 (파일 크기 약 1/5)"""
    from PIL import Image
    import glob as _g
    for f in _g.glob(os.path.join(folder, '*.png')):
        Image.open(f).convert('L').quantize(16).save(f, optimize=True)


def render():
    """eq/*.png와 eq/sizes.json 만들기"""
    import subprocess, tempfile
    kdist = os.environ.get('KATEX_DIST') or next((p for p in [os.path.join(os.environ.get('SCRATCH', ''), 'katex/node_modules/katex/dist')] if os.path.isdir(p)), None)
    os.makedirs(EQ_DIR, exist_ok=True)
    with tempfile.NamedTemporaryFile('w', suffix='.json', delete=False, encoding='utf8') as f:
        json.dump([{'id': k, 'tex': v} for k, v in TEX.items()], f, ensure_ascii=False)
    subprocess.run(['node', os.path.join(HERE, 'eq_render.cjs'), f.name, EQ_DIR, kdist], check=True)
    _shrink(EQ_DIR)
    print('식 그림', len(TEX), '개 →', EQ_DIR)


def _spec_for(path, sheet):
    base = os.path.basename(path)
    for key, pat, eqs in SPECS:
        if key in base and re.search(pat, sheet):
            return eqs
    return None


def _ensure_drawing(b, sheet):
    """시트의 그림 파트 (없으면 만든다) → (그림 파트 경로, 그림 rels 경로)"""
    sp = b.sheets[sheet]
    s = b.sx(sheet)
    srels = sp.replace('worksheets/', 'worksheets/_rels/') + '.rels'
    d = re.search(r'<drawing r:id="([^"]+)"', s)
    if d:
        rels = b.text(srels)
        t = re.search(rf'Id="{d.group(1)}"[^>]*Target="([^"]+)"', rels) or re.search(rf'Target="([^"]+)"[^>]*Id="{d.group(1)}"', rels)
        dp = 'xl/drawings/' + t.group(1).split('/')[-1]
    else:
        n = 1
        while f'xl/drawings/drawing{n}.xml' in b.parts:
            n += 1
        dp = f'xl/drawings/drawing{n}.xml'
        b.put(dp, '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"></xdr:wsDr>')
        rels = b.text(srels) if srels in b.parts else '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"></Relationships>'
        rid = 'rIdEq1'
        while f'Id="{rid}"' in rels:
            rid += 'x'
        rels = rels.replace('</Relationships>', f'<Relationship Id="{rid}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing" Target="../drawings/drawing{n}.xml"/></Relationships>')
        b.put(srels, rels)
        if 'xmlns:r=' not in s[:500]:
            s = s.replace('<worksheet ', '<worksheet xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" ', 1)
        m = re.search(r'<(legacyDrawing|legacyDrawingHF|picture|oleObjects|controls|webPublishItems|tableParts|extLst)\b', s)
        tag = f'<drawing r:id="{rid}"/>'
        s = s[:m.start()] + tag + s[m.start():] if m else s.replace('</worksheet>', tag + '</worksheet>')
        b.setsx(sheet, s)
        ct = b.text('[Content_Types].xml')
        if f'/{dp}"' not in ct:
            ct = ct.replace('</Types>', f'<Override PartName="/{dp}" ContentType="application/vnd.openxmlformats-officedocument.drawing+xml"/></Types>')
            b.put('[Content_Types].xml', ct)
    drels = dp.replace('drawings/', 'drawings/_rels/') + '.rels'
    if drels not in b.parts:
        b.put(drels, '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"></Relationships>')
    return dp, drels


def _interior_borders(b, sheet, r1, r2, c1, c2, S):
    """영역 안쪽 테두리를 지운다 (칸으로 만든 분수의 가로선 등). 영역 바깥 테두리는 그대로"""
    from form_style import _set_styles, _cell_style
    x = b.sx(sheet); new = {}
    for r in range(r1, r2 + 1):
        for c in range(c1, c2 + 1):
            st = _cell_style(x, f'{CL(c)}{r}', None, {}, c)
            bx = S.border_xml((re.search(r'borderId="(\d+)"', S.xf(st)) or [0, '0'])[1])
            sides = []
            for name, keep in (('left', c == c1), ('right', c == c2), ('top', r == r1), ('bottom', r == r2)):
                m = re.search(rf'<{name} style="([^"]+)"', bx)
                sides.append(m.group(1) if (m and keep) else None)
            cur = [(re.search(rf'<{n} style="([^"]+)"', bx) or [0, None])[1] for n in ('left', 'right', 'top', 'bottom')]
            if cur != sides:
                new[(r, c)] = S.with_border(st, tuple(sides))
    if new:
        _set_styles(b, sheet, new)
    return len(new)


def _place(b, ws, dx, drels, key, png, size, r1, r2, c1, c2, name, valign='middle'):
    """그림 하나를 영역(r1~r2, c1~c2)에 넣는다: 넘치면 줄이고(비율 그대로), 왼쪽 여백, 위아래 가운데(또는 위)"""
    wpx = sum(int((ws.column_dimensions[CL(c)].width or 9) * MDW) for c in range(c1, c2 + 1))
    hpx = sum((ws.row_dimensions[r].height or 15) for r in range(r1, r2 + 1)) * 96 / 72
    w, h = size
    k = min(1.0, (wpx - 2 * PAD_X) / w, (hpx - 2 * PAD_Y) / h)
    w2, h2 = w * k, h * k
    yoff = (hpx - h2) / 2 if valign == 'middle' else min(PAD_Y + 2, (hpx - h2) / 2)
    row, acc = r1, 0.0
    while row < r2 and acc + (ws.row_dimensions[row].height or 15) * 96 / 72 <= yoff:
        acc += (ws.row_dimensions[row].height or 15) * 96 / 72; row += 1
    rowoff = int((yoff - acc) * 9525)
    fname = ('eq_' + key if not key.startswith('lg_') else key) + '.png'
    media = 'xl/media/' + fname
    if media not in b.parts:
        b.parts[media] = open(png, 'rb').read()
    rels = b.text(drels)
    rid = next((m.group(1) for m in re.finditer(r'<Relationship Id="([^"]+)"[^>]*Target="../media/' + re.escape(fname) + '"', rels)), None)
    if not rid:
        n = 1
        while f'Id="rIdEq{n}"' in rels:
            n += 1
        rid = f'rIdEq{n}'
        b.put(drels, rels.replace('</Relationships>', f'<Relationship Id="{rid}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/{fname}"/></Relationships>'))
    ids = [int(i) for i in re.findall(r'<xdr:cNvPr id="(\d+)"', dx)] or [1]
    pid = max(ids) + 1
    cx, cy = int(w2 * 9525), int(h2 * 9525)
    anchor = (f'<xdr:oneCellAnchor><xdr:from><xdr:col>{c1 - 1}</xdr:col><xdr:colOff>{PAD_X * 9525}</xdr:colOff><xdr:row>{row - 1}</xdr:row><xdr:rowOff>{rowoff}</xdr:rowOff></xdr:from>'
              f'<xdr:ext cx="{cx}" cy="{cy}"/><xdr:pic><xdr:nvPicPr><xdr:cNvPr id="{pid}" name="{name} {key}" descr="{name}"/><xdr:cNvPicPr><a:picLocks noChangeAspect="1"/></xdr:cNvPicPr></xdr:nvPicPr>'
              f'<xdr:blipFill><a:blip xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:embed="{rid}"/><a:stretch><a:fillRect/></a:stretch></xdr:blipFill>'
              f'<xdr:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="{cx}" cy="{cy}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></xdr:spPr></xdr:pic><xdr:clientData/></xdr:oneCellAnchor>')
    return dx.replace('</xdr:wsDr>', anchor + '</xdr:wsDr>'), k


# ── 기호 설명 ──
# 원본 설명 칸 (파일 이름에 든 말 → [(칸들, 첫 열, 끝 열)]; 끝 열 None = 인쇄 영역 끝)
LEGENDS = [
    ('_CS2', [(['I21', 'I22', 'I23', 'I24'], 9, 10)]),
    ('_이황화메틸', [(['F21', 'F22', 'F23', 'F24'], 6, 10)]),
    ('_VOCs', [(['I21', 'I22', 'I23', 'I24'], 9, 10)]),
    ('_SOx', [(['I13', 'I14', 'I15', 'I16'], 9, 10)]),
    ('_PAHs', [(['D14'], 4, None)]), ('_먼지', [(['D13'], 4, None)]),
    ('_브로민', [(['D14'], 4, None)]), ('_브롬', [(['D14'], 4, None)]), ('_염화수소', [(['D14'], 4, None)]),
    ('_플루오린', [(['D14'], 4, None)]), ('_불소', [(['D14'], 4, None)]), ('_비소', [(['E14'], 4, None)]),
    ('_사이안화수소', [(['D14'], 4, None)]), ('_시안화수소', [(['D14'], 4, None)]), ('_암모니아', [(['D14'], 4, None)]),
    ('_황화수소', [(['D14'], 4, None)]), ('_페놀', [(['D14'], 4, None)]),
    ('_아세트알데하이드', [(['D14'], 4, None)]), ('_폼알데하이드', [(['D14'], 4, None)]), ('_포름알데하이드', [(['D14'], 4, None)]),
    ('_수은', [(['D14'], 4, None), (['D16'], 4, None), (['D18'], 4, None)]),
    ('_총탄화수소', [(['D12'], 4, None)]), ('_중금속', [(['E14'], 4, None)]),
]
SYM = {
    'C': 'C', 'Ca': 'C_a', 'Cb': 'C_b', 'Cc': 'C_c', 'Cn': 'C_n', 'Cis': r'C_{\mathrm{is}}', 'C측정': r'C_{\text{측정}}',
    'Iis': r'I_{\mathrm{is}}', 'As': 'A_s', 'Ais': r'A_{\mathrm{is}}', 'Aa': 'A_a', 'Ab': 'A_b', 'B': 'B', 'K': 'K',
    'MWa': r'\mathit{MW}_a', 'Oa': 'O_a', 'Os': 'O_s', 'Pa': 'P_a', 'RRFavg': r'\mathit{RRF}_{\mathrm{avg}}', 'Ta': 'T_a',
    'V': 'V', 'Vm': 'V_m', 'Vm(std)': r'V_{\mathrm{m(std)}}', 'Vs': 'V_s', 'Xi': 'x_i', 'yi': 'y_i', 'a': 'a', 'b': 'b', 'f': 'f',
    'm1': 'm_1', 'm2': 'm_2', 'md': 'm_d', 'ΔH': r'\Delta H', '⊿H': r'\Delta H', 'Θm': r'\theta_m', 'θm': r'\theta_m',
    'C(HCl)Hg': r'C_{\mathrm{(HCl)Hg}}', 'C(HCl bk)Hg': r'C_{\mathrm{(HCl\,bk)Hg}}', 'C(sol)Hg': r'C_{\mathrm{(sol)Hg}}', 'C(sol bk)Hg': r'C_{\mathrm{(sol\,bk)Hg}}',
    'Vf(HCl)': r'V_{\mathrm{f(HCl)}}', 'Vf(HCl bk)': r'V_{\mathrm{f(HCl\,bk)}}', 'Vf(sol)': r'V_{\mathrm{f(sol)}}', 'Vf(sol bk)': r'V_{\mathrm{f(sol\,bk)}}',
    'm(HCl)Hg': r'm_{\mathrm{(HCl)Hg}}', 'm(sol)Hg': r'm_{\mathrm{(sol)Hg}}',
}


def _legend_spec(path, sheet):
    if _spec_for(path, sheet) is None:
        return []
    base = os.path.basename(path)
    return next((v for k, v in LEGENDS if k in base), [])


def tidy_desc(d):
    """설명 글 손질: 단위 표기 통일, 끝 쉼표·겹친 띄어쓰기 정리 (내용은 그대로)"""
    d = re.sub(r'\s+', ' ', d).strip().rstrip(',').strip()
    d = d.replace('ug/mL', 'μg/mL').replace('mmH2O', 'mmH₂O').replace('(Cl-)', '(Cl⁻)').replace('mmHg)', 'mmHg)')
    d = re.sub(r'\b(S?m)3\b', lambda m: m.group(1) + '³', d)
    d = re.sub(r'\(\s+', '(', d); d = re.sub(r'\s+\)', ')', d)
    d = re.sub(r'([가-힣])\(', r'\1 (', d)       # 괄호 앞 띄어쓰기
    d = d.replace('(sL)', '(L)').replace('1 atm)', '1 atm)')
    return d


def parse_legend(text):
    """'기호 : 설명' 줄들 → [(기호 TeX, 설명)]. 기호 없는 줄은 앞 설명에 이어 붙이고, 한 줄에 둘 있으면 나눈다"""
    lines = []
    for ln in text.split('\n'):
        lines += re.split(r'(?<=\)),?\s+(?=[A-Za-z][\w()]{0,8}\s:\s)', ln)
    out = []
    for ln in lines:
        m = re.match(r'\s*([^:=]+?)\s*[:=]\s*(.*)$', ln)
        sym = m.group(1).strip() if m else None
        if m and (sym in SYM or re.fullmatch(r'\d+(\.\d+)?', sym)):
            out.append([SYM.get(sym, sym), m.group(2)])
        elif out and ln.strip():
            out[-1][1] += ' ' + ln.strip()
        elif ln.strip():
            raise ValueError(f'기호를 모르는 설명 줄: {ln!r}')
    return [[s_, tidy_desc(d)] for s_, d in out]


def render_legends(collect):
    """LEGEND_COLLECT로 모은 설명들 → eq/legend/*.png"""
    import subprocess, tempfile
    kdist = os.environ.get('KATEX_DIST') or os.path.join(os.environ.get('SCRATCH', ''), 'katex/node_modules/katex/dist')
    items = {}
    for ln in open(collect, encoding='utf8'):
        d = json.loads(ln); items[d['id']] = d
    out = os.path.join(EQ_DIR, 'legend'); os.makedirs(out, exist_ok=True)
    with tempfile.NamedTemporaryFile('w', suffix='.json', delete=False, encoding='utf8') as f:
        json.dump(list(items.values()), f, ensure_ascii=False)
    subprocess.run(['node', os.path.join(HERE, 'eq_render.cjs'), f.name, out, kdist], check=True)
    _shrink(out)
    print('기호 설명 그림', len(items), '개 →', out)


def apply(b, path, sheets, print_cols, wb):
    """기록부 시트들의 계산식을 그림으로. print_cols: {시트: 인쇄 영역 끝 열}"""
    from form_style import Styles
    sizes = json.load(open(os.path.join(EQ_DIR, 'sizes.json')))
    log = []
    for sheet in sheets:
        eqs = _spec_for(path, sheet)
        if not eqs:
            continue
        ws = wb[sheet]
        S = Styles(b)
        mv = next((v for k_, v in ROW_MOVE.items() if k_ in os.path.basename(path)), None)
        if mv:
            from raw_layout import _set_row
            x = b.sx(sheet); tot0 = sum((ws.row_dimensions[r].height or 15) for r in mv)
            for r, h in mv.items():
                nh = h if h is not None else round(tot0 - sum(v for v in mv.values() if v is not None), 2)
                x = _set_row(x, r, nh); ws.row_dimensions[r].height = nh
            assert abs(sum(ws.row_dimensions[r].height for r in mv) - tot0) < 0.01, (sheet, tot0)
            b.setsx(sheet, x)
        dp, drels = _ensure_drawing(b, sheet)
        dx = b.text(dp)
        for key, r1, r2, c1, c2, clears, *more in eqs:
            drop_r2 = more[0] if more else r2
            c2 = c2 or print_cols[sheet]
            # 1) 영역 안 수식 개체(엑셀 수식) 지우기
            def drop(m):
                fr = re.search(r'<xdr:from><xdr:col>(\d+)</xdr:col>.*?<xdr:row>(\d+)</xdr:row>', m.group(2), re.S)
                row, col = int(fr.group(2)) + 1, int(fr.group(1)) + 1
                return '' if ('m:oMath' in m.group(2) and r1 <= row <= drop_r2 and c1 <= col <= c2) else m.group(0)
            dx2 = re.sub(r'<xdr:(twoCellAnchor|oneCellAnchor)\b[^>]*>(.*?)</xdr:\1>', drop, dx, flags=re.S)
            n_obj = (len(re.findall('m:oMath', dx)) - len(re.findall('m:oMath', dx2))) // 2 or (dx2 != dx)
            dx = dx2
            # 2) 식 칸 글자·안쪽 테두리 지우기
            for ref in clears:
                if ws[ref].value is not None:
                    b.clear(sheet, [ref]); ws[ref].value = None
            nb = _interior_borders(b, sheet, r1, r2, c1, c2, S) if clears else 0
            # 3) 그림 넣기: 영역에 맞춰 (작으면 그대로, 크면 줄임) 왼쪽 여백·위아래 가운데
            dx, k = _place(b, ws, dx, drels, key, os.path.join(EQ_DIR, key + '.png'), sizes[key], r1, r2, c1, c2, '계산식')
            log.append(f'{sheet}: 계산식 {key} 그림 ({CL(c1)}{r1}:{CL(c2)}{r2}, 배율 {k:.2f}' + (f', 수식 개체 지움' if n_obj else '') + (f', 칸 {",".join(clears)} 글자·안쪽 테두리 {nb}칸 지움' if clears else '') + ')'
                       + (' !! 많이 줄어듦' if k < 0.75 else ''))
        # 4) 기호 설명: 칸 글자를 기호(식과 같은 조판)·콜론 맞춘 표 그림으로
        for cells, c1, c2 in _legend_spec(path, sheet):
            c2 = c2 or print_cols[sheet]
            texts = [str(ws[ref].value or '') for ref in cells]
            if not any(t.strip() for t in texts):
                continue
            rows = set()
            for ref in cells:
                g = next((g for g in ws.merged_cells.ranges if ref in g), None)
                rows |= set(range(g.min_row, g.max_row + 1)) if g else {ws[ref].row}
            r1, r2 = min(rows), max(rows)
            entries = parse_legend('\n'.join(texts))
            wpx = sum(int((ws.column_dimensions[CL(c)].width or 9) * MDW) for c in range(c1, c2 + 1))
            maxw = int(wpx - 2 * PAD_X)
            lid = 'lg_' + hashlib.md5(json.dumps([entries, maxw], ensure_ascii=False).encode()).hexdigest()[:10]
            png = os.path.join(EQ_DIR, 'legend', lid + '.png')
            lsizes = json.load(open(os.path.join(EQ_DIR, 'legend', 'sizes.json'))) if os.path.exists(os.path.join(EQ_DIR, 'legend', 'sizes.json')) else {}
            if lid not in lsizes or not os.path.exists(png):
                if os.environ.get('LEGEND_COLLECT'):
                    with open(os.environ['LEGEND_COLLECT'], 'a', encoding='utf8') as f:
                        f.write(json.dumps({'id': lid, 'legend': entries, 'maxw': maxw}, ensure_ascii=False) + '\n')
                log.append(f'{sheet}: !! 기호 설명 그림 없음 ({lid}) — equations.py legends로 만들기')
                continue
            for ref in cells:
                b.clear(sheet, [ref]); ws[ref].value = None
            nb = _interior_borders(b, sheet, r1, r2, c1, c2, S)
            dx, k = _place(b, ws, dx, drels, lid, png, lsizes[lid], r1, r2, c1, c2, '기호 설명', valign='middle' if c1 > 5 else 'top')   # 식 오른쪽 설명은 가운데
            log.append(f'{sheet}: 기호 설명 {len(entries)}개 그림 ({CL(c1)}{r1}:{CL(c2)}{r2}, 배율 {k:.2f}, 안쪽 테두리 {nb}칸 지움)' + (' !! 많이 줄어듦' if k < 0.8 else ''))
        b.put(dp, dx)
    ct = b.text('[Content_Types].xml')
    if 'Extension="png"' not in ct:
        b.put('[Content_Types].xml', ct.replace('<Default ', '<Default Extension="png" ContentType="image/png"/><Default ', 1))
    return log


if __name__ == '__main__' and sys.argv[1:] == ['render']:
    render()
if __name__ == '__main__' and sys.argv[1:2] == ['legends']:
    render_legends(sys.argv[2])
