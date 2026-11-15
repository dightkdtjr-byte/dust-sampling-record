// 계산식 시트의 수식 그림(캡처)을 엑셀 수식 개체(Office Math)로 바꾼다.
// 엑셀에서는 실제 수식으로 보이고(선명하게 인쇄, 수식 도구로 고칠 수 있음), 수식 개체를 모르는 프로그램
// (휴대폰 미리보기·한셀 등)은 mc:Fallback에 남겨 둔 원래 그림을 보여 준다.
const A14 = 'http://schemas.microsoft.com/office/drawing/2010/main';
const MC = 'http://schemas.openxmlformats.org/markup-compatibility/2006';
const M = 'http://schemas.openxmlformats.org/officeDocument/2006/math';
const esc = (text) => String(text).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');

// ── 수식 조각 (OMML) ──
const rPr = (sz, italic) => `<a:rPr lang="en-US" sz="${sz}"${italic ? ' i="1"' : ' i="0"'}><a:latin typeface="Cambria Math" panose="02040503050406030204" pitchFamily="18" charset="0"/><a:ea typeface="Cambria Math" panose="02040503050406030204" pitchFamily="18" charset="0"/></a:rPr>`;
const ctrl = (sz) => `<m:ctrlPr>${rPr(sz, true)}</m:ctrlPr>`;
// 글자: 변수는 기울임, 숫자·기호·단위는 바로 세움
const v = (text) => ({ kind: 'run', text, italic: true });
const n = (text) => ({ kind: 'run', text, italic: false });
const sub = (base, index) => ({ kind: 'sub', base, index });
const sup = (base, index) => ({ kind: 'sup', base, index });
const frac = (num, den) => ({ kind: 'frac', num, den });
const sqrt = (body) => ({ kind: 'rad', body });
const group = (body, open = '(', close = ')') => ({ kind: 'delim', body, open, close });
const render = (node, sz) => {
  if (Array.isArray(node)) return node.map((x) => render(x, sz)).join('');
  switch (node.kind) {
    case 'run': return `<m:r>${node.italic ? '' : '<m:rPr><m:sty m:val="p"/></m:rPr>'}${rPr(sz, node.italic)}<m:t>${esc(node.text)}</m:t></m:r>`;
    case 'sub': return `<m:sSub><m:sSubPr>${ctrl(sz)}</m:sSubPr><m:e>${render(node.base, sz)}</m:e><m:sub>${render(node.index, sz)}</m:sub></m:sSub>`;
    case 'sup': return `<m:sSup><m:sSupPr>${ctrl(sz)}</m:sSupPr><m:e>${render(node.base, sz)}</m:e><m:sup>${render(node.index, sz)}</m:sup></m:sSup>`;
    case 'frac': return `<m:f><m:fPr>${ctrl(sz)}</m:fPr><m:num>${render(node.num, sz)}</m:num><m:den>${render(node.den, sz)}</m:den></m:f>`;
    case 'rad': return `<m:rad><m:radPr><m:degHide m:val="1"/>${ctrl(sz)}</m:radPr><m:deg/><m:e>${render(node.body, sz)}</m:e></m:rad>`;
    case 'delim': return `<m:d><m:dPr><m:begChr m:val="${esc(node.open)}"/><m:endChr m:val="${esc(node.close)}"/>${ctrl(sz)}</m:dPr><m:e>${render(node.body, sz)}</m:e></m:d>`;
    default: return '';
  }
};

// ── 계산식 시트의 수식들 (그림 이름 → 수식). 양식의 그림과 같은 식이다 ──
const X = n(' × ');
const tempRatio = (theta) => frac(n('273'), [n('273 + '), sub(v('θ'), v(theta))]);
const o2Ratio = frac([n('21 − '), sub(v('O'), v('s'))], [n('21 − '), sub(v('O'), v('a'))]);
const moistTerm = [frac(n('22.4'), n('18')), sub(v('m'), v('a'))];
export const CALC_SHEET_EQUATIONS = {
  // 수분량 (%)
  _x83189624: [sub(v('X'), v('w')), n(' = '), frac(moistTerm, [sub(v('V'), v('m')), X, tempRatio('m'), X, frac([sub(v('P'), v('a')), n(' + '), sub(v('P'), v('m'))], n('760')), n(' + '), ...moistTerm]), n(' × 100')],
  // 배출가스 밀도
  _x83201912: [v('r'), n(' = '), sub(v('r'), n('0')), X, tempRatio('s'), X, frac([sub(v('P'), v('a')), n(' + '), sub(v('P'), v('s'))], n('760'))],
  // 배출가스 유속
  _x134303384: [v('V'), n(' = '), v('C'), sqrt(frac([n('2'), v('gh')], v('r')))],
  // 등속흡입계수
  _x133882184: [v('I'), n('(%) = '), frac(
    [sub(v('T'), v('s')), group([n('0.00346 '), sub(v('V'), v('ic')), n(' + '), frac(sub(sup(v('V'), n('′')), v('m')), sub(v('T'), v('m'))), group([sub(v('P'), v('bar')), n(' + '), frac([n('Δ'), v('H')], n('13.6'))])], '[', ']')],
    [sub(v('P'), v('s')), n(' · '), v('t'), n(' · '), v('V'), n(' · '), sub(v('A'), v('n'))],
  ), n(' × '), group([n('1.667 × '), sup(n('10'), n('4'))])],
  // 배출가스 유량
  _x83433656: [sub(v('Q'), v('a')), n(' = '), v('v'), X, v('A'), X, tempRatio('s'), X, frac([sub(v('P'), v('a')), n(' + '), sub(v('P'), v('s'))], n('760')), X, group([n('1 − '), frac(sub(v('X'), v('w')), n('100'))]), n(' × 3,600')],
  // 표준산소 환산 유량
  _x86843880: [v('Q'), n(' = '), sub(v('Q'), v('a')), n(' ÷ '), o2Ratio],
  // 먼지농도
  '그림 5': [sub(v('C'), v('a')), n(' = '), frac(sub(v('m'), v('d')), [sub(sup(v('V'), n('′')), v('m')), X, tempRatio('m'), X, frac([sub(v('P'), v('a')), n(' + '), frac([n('Δ'), v('H')], n('13.6'))], n('760'))])],
  // 표준산소 환산 농도
  _x86845240: [v('C'), n(' = '), sub(v('C'), v('a')), X, o2Ratio],
};

// 그림 하나(xdr:pic)를 수식 글상자로 바꾼 XML. 원래 그림은 Fallback으로 남긴다
const equationShape = (picXml, nodes, sz) => {
  const id = (/<xdr:cNvPr\b[^>]*\bid="(\d+)"/.exec(picXml) || [])[1] || '1';
  const name = (/<xdr:cNvPr\b[^>]*\bname="([^"]*)"/.exec(picXml) || [])[1] || 'Equation';
  const xfrm = (/<a:xfrm\b[^>]*>[\s\S]*?<\/a:xfrm>/.exec(picXml) || [])[0] || '<a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/></a:xfrm>';
  const math = `<m:oMathPara xmlns:m="${M}"><m:oMathParaPr><m:jc m:val="left"/></m:oMathParaPr><m:oMath xmlns:m="${M}">${render(nodes, sz)}</m:oMath></m:oMathPara>`;
  const shape = `<xdr:sp macro="" textlink=""><xdr:nvSpPr><xdr:cNvPr id="${id}" name="${name}"/><xdr:cNvSpPr txBox="1"/></xdr:nvSpPr>`
    + `<xdr:spPr>${xfrm}<a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:noFill/><a:ln><a:noFill/></a:ln></xdr:spPr>`
    + `<xdr:txBody><a:bodyPr vertOverflow="overflow" horzOverflow="overflow" wrap="none" lIns="0" tIns="0" rIns="0" bIns="0" rtlCol="0" anchor="ctr"><a:noAutofit/></a:bodyPr><a:lstStyle/>`
    + `<a:p><a:pPr algn="l"/><a14:m>${math}</a14:m></a:p></xdr:txBody></xdr:sp>`;
  return `<mc:AlternateContent xmlns:mc="${MC}"><mc:Choice xmlns:a14="${A14}" Requires="a14">${shape}</mc:Choice><mc:Fallback>${picXml}</mc:Fallback></mc:AlternateContent>`;
};

// 시트에 딸린 그림 파일에서, 이름이 equations에 있는 그림을 수식 개체로 바꾼다. 바꾼 개수를 돌려준다
// sizes: 그림 이름별 글자 크기 (분수가 겹친 큰 수식은 아래 숫자 줄을 덮지 않게 작게)
export async function replacePicturesWithEquations(zip, sheetPath, equations = CALC_SHEET_EQUATIONS, { fontPt = 18, sizes = { _x83189624: 14, '그림 5': 15 }, startCol = 2 } = {}) {
  const dir = sheetPath.replace(/[^/]+$/, '');
  const relsFile = zip.file(`${dir}_rels/${sheetPath.split('/').pop()}.rels`);
  if (!relsFile) return 0;
  const rels = await relsFile.async('string');
  let count = 0;
  for (const m of rels.matchAll(/<Relationship\b([^>]*)\/?>/g)) {
    if (!/Type="[^"]*\/drawing"/.test(m[1])) continue;
    const target = (/Target="([^"]*)"/.exec(m[1]) || [])[1] || '';
    const stack = [];
    for (const part of (dir + target).split('/')) { if (part === '..') stack.pop(); else if (part && part !== '.') stack.push(part); }
    const file = zip.file(stack.join('/'));
    if (!file) continue;
    const drawing = await file.async('string');
    const next = drawing.replace(/<xdr:(twoCellAnchor|oneCellAnchor)\b([^>]*)>([\s\S]*?)<\/xdr:\1>/g, (whole, kind, attrs, body) => {
      if (body.includes('<mc:AlternateContent')) return whole;
      const pic = /<xdr:pic\b[\s\S]*?<\/xdr:pic>/.exec(body);
      if (!pic || /<xdr:grpSp\b/.test(body)) return whole;
      const name = (/<xdr:cNvPr\b[^>]*\bname="([^"]*)"/.exec(pic[0]) || [])[1];
      if (!name || !Object.prototype.hasOwnProperty.call(equations, name)) return whole;
      count += 1;
      // 수식 위치 정리: 모든 수식을 같은 열(C열)에서 시작하게 한다
      const fromCol = Number((/<xdr:from>\s*<xdr:col>(\d+)<\/xdr:col>/.exec(body) || [])[1]);
      if (Number.isFinite(fromCol)) {
        const delta = startCol - fromCol;
        body = body.replace(/(<xdr:from>\s*)<xdr:col>\d+<\/xdr:col>(\s*)<xdr:colOff>-?\d+<\/xdr:colOff>/, (w, a, gap) => `${a}<xdr:col>${startCol}</xdr:col>${gap}<xdr:colOff>0</xdr:colOff>`)
          .replace(/(<xdr:to>\s*)<xdr:col>(\d+)<\/xdr:col>/, (w, a, c) => `${a}<xdr:col>${Math.max(startCol + 1, Number(c) + delta)}</xdr:col>`);
      }
      return `<xdr:${kind}${attrs}>${body.replace(pic[0], () => equationShape(pic[0], equations[name], Math.round((sizes[name] || fontPt) * 100)))}</xdr:${kind}>`;
    });
    if (next !== drawing) zip.file(stack.join('/'), next);
  }
  return count;
}

// ── 기호 설명(범례): 그림 대신 글상자의 실제 글자로 ──
// 수식 개체가 아니라 보통 글자(아래첨자·기울임)로 적어 엑셀·휴대폰·한셀 어디서나 똑같이 보이고 선명하게 인쇄된다.
// 한 줄 = [기호 조각들, 설명]. 기호 조각: 'P' (기울임) 또는 ['P', 's'] (기호 + 아래첨자)
const S = (base, index) => [base, index];
const LEGENDS = {
  _x83189624: {
    left: [[[S('X', 'w')], '배출가스 중의 수증기의 부피 백분율 (%)'], [[S('m', 'a')], '흡습 수분의 질량 (mₐ₂ − mₐ₁) (g)'], [[S('V', 'm')], '흡입한 건조 가스량 (건식가스미터에서 읽은 값) (L)']],
    right: [[[S('θ', 'm')], '가스미터에서의 흡입 가스온도 (℃)'], [[S('P', 'a')], '대기압 (mmHg)'], [[S('P', 'm')], '가스미터에서의 가스의 게이지압 (mmHg)']],
  },
  _x83201912: {
    left: [[['r'], '굴뚝 내의 배출가스 밀도 (kg/m³)'], [[S('r', '0')], '온도 0 ℃ 기압 760 mmHg로 환산한 습한 배출가스 밀도 (kg/Sm³)'], [[S('θ', 's')], '각 측정점에서 배출가스 온도의 평균치 (℃)']],
    right: [[[S('P', 'a')], '대기압 (mmHg)'], [[S('P', 's')], '각 측정점에서 배출가스 정압의 평균치 (mmHg)']],
  },
  _x134303384: {
    left: [[['V'], '유속 (m/s)'], [['C'], '피토관 계수']],
    right: [[['h'], '피토관에 의한 동압 측정치 (mmH₂O)'], [['g'], '중력가속도 (9.81 m/s²)'], [['r'], '굴뚝 내의 배출가스 밀도 (kg/m³)']],
  },
  _x133882184: {
    left: [[['I'], '등속흡입계수 (%)'], [[S('P', 's')], '배출가스 압력 (mmHg: 760 + P′s)'], [[S('P′', 's')], '배출가스 정압'], [[S('T', 's')], '배출가스 평균 절대온도 (K: 273 + θs (℃))'], [[S('V', 'ic')], '임핀저와 실리카겔에 채취된 물의 총량 (mL)'], [[S('V′', 'm')], '건식가스미터에서 읽은 가스시료채취량 (m³)']],
    right: [[[S('T', 'm')], '건식가스미터의 평균 절대온도 (K: 273 + θm (℃))'], [[S('P', 'bar')], '측정공 위치의 대기압 (mmHg)'], [['ΔH'], '오리피스 압차 (mmH₂O)'], [['t'], '총시료채취시간 (분)'], [['V'], '배출가스 유속 (m/s)'], [[S('A', 'n')], '노즐의 단면적 (cm²)']],
  },
  _x83433656: {
    left: [[[S('Q', 'a')], '배출가스 유량 (Sm³/hr)'], [['v'], '배출가스 유속 (m/s)'], [['A'], '굴뚝 내부 면적 (m²)']],
    right: [[[S('P', 'a')], '대기압 (mmHg)'], [[S('P', 's')], '각 측정점에서 배출가스 정압의 평균치 (mmHg)'], [[S('θ', 's')], '각 측정점에서 배출가스 온도의 평균치 (℃)'], [[S('X', 'w')], '습윤 배출가스 중의 수증기 부피 백분율 (%)']],
  },
  _x86843880: {
    left: [[['Q'], '배출가스 유량 (Sm³/hr)'], [[S('Q', 'a')], '실측 배출가스 유량 (Sm³/hr)']],
    right: [[[S('O', 's')], '표준산소농도 (%)'], [[S('O', 'a')], '실측산소농도 (%)']],
  },
  '그림 5': {
    left: [[[S('C', 'a')], '먼지농도 (mg/Sm³)'], [[S('m', 'd')], '포집된 먼지량 (mg)'], [[S('V′', 'm')], '건식가스미터에서 읽은 가스시료채취량 (m³)']],
    right: [[[S('θ', 'm')], '건식가스미터의 평균온도 (℃)'], [[S('P', 'a')], '측정공 위치의 대기압 (mmHg)'], [['ΔH'], '오리피스 압력차 (mmH₂O)']],
  },
  _x86845240: {
    left: [[['C'], '오염물질 농도 (mg/Sm³ 또는 ppm)'], [[S('C', 'a')], '실측 오염물질 농도 (mg/Sm³ 또는 ppm)']],
    right: [[[S('O', 's')], '표준산소농도 (%)'], [[S('O', 'a')], '실측산소농도 (%)']],
  },
};
const legendParagraph = (line, sz) => {
  const sym = `<a:latin typeface="Cambria Math"/><a:ea typeface="Cambria Math"/>`;
  const kor = `<a:latin typeface="맑은 고딕"/><a:ea typeface="맑은 고딕"/>`;
  const runs = [];
  for (const part of line[0]) {
    const [base, index] = Array.isArray(part) ? part : [part, ''];
    runs.push(`<a:r><a:rPr lang="en-US" altLang="ko-KR" sz="${sz}" i="1">${sym}</a:rPr><a:t>${esc(base)}</a:t></a:r>`);
    if (index) runs.push(`<a:r><a:rPr lang="en-US" altLang="ko-KR" sz="${sz}" i="1" baseline="-25000">${sym}</a:rPr><a:t>${esc(index)}</a:t></a:r>`);
  }
  runs.push(`<a:r><a:rPr lang="ko-KR" altLang="en-US" sz="${sz}">${kor}</a:rPr><a:t> = ${esc(line[1])}</a:t></a:r>`);
  // 줄 간격을 행 높이(16.5pt)에 맞춰 고정: 한 줄이 한 행을 차지해 아래 묶음으로 넘치지 않는다
  return `<a:p><a:pPr><a:lnSpc><a:spcPts val="1650"/></a:lnSpc></a:pPr>${runs.join('')}</a:p>`;
};
const legendAnchor = (id, name, col, row, lines, fontPt) => {
  const sz = Math.round(fontPt * 100);
  const cy = Math.round(lines.length * 16.5 * 12700);
  return `<xdr:oneCellAnchor><xdr:from><xdr:col>${col}</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>${row}</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from><xdr:ext cx="5400000" cy="${cy}"/>`
    + `<xdr:sp macro="" textlink=""><xdr:nvSpPr><xdr:cNvPr id="${id}" name="${name}"/><xdr:cNvSpPr txBox="1"/></xdr:nvSpPr>`
    + `<xdr:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="5400000" cy="${cy}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:noFill/><a:ln><a:noFill/></a:ln></xdr:spPr>`
    + `<xdr:txBody><a:bodyPr vertOverflow="overflow" horzOverflow="overflow" wrap="none" lIns="0" tIns="0" rIns="0" bIns="0" rtlCol="0" anchor="t"><a:noAutofit/></a:bodyPr><a:lstStyle/>${lines.map((line) => legendParagraph(line, sz)).join('')}</xdr:txBody></xdr:sp><xdr:clientData/></xdr:oneCellAnchor>`;
};
// 수식 그림(이름이 LEGENDS에 있는 것) 아래의 설명 그림들을 지우고, 같은 자리에 글상자를 넣는다. 바꾼 묶음 수를 돌려준다
export async function replaceLegendPictures(zip, sheetPath, { fontPt = 11, leftCol = 1, rightCol = 9 } = {}) {
  const dir = sheetPath.replace(/[^/]+$/, '');
  const relsFile = zip.file(`${dir}_rels/${sheetPath.split('/').pop()}.rels`);
  if (!relsFile) return 0;
  const rels = await relsFile.async('string');
  let count = 0;
  for (const m of rels.matchAll(/<Relationship\b([^>]*)\/?>/g)) {
    if (!/Type="[^"]*\/drawing"/.test(m[1])) continue;
    const target = (/Target="([^"]*)"/.exec(m[1]) || [])[1] || '';
    const stack = [];
    for (const part of (dir + target).split('/')) { if (part === '..') stack.pop(); else if (part && part !== '.') stack.push(part); }
    const path = stack.join('/');
    const file = zip.file(path);
    if (!file) continue;
    const drawing = await file.async('string');
    const anchors = [...drawing.matchAll(/<xdr:(twoCellAnchor|oneCellAnchor)\b[\s\S]*?<\/xdr:\1>/g)].map((a) => {
      const row = Number((/<xdr:from>[\s\S]*?<xdr:row>(\d+)<\/xdr:row>/.exec(a[0]) || [])[1]);
      const name = (/<xdr:cNvPr\b[^>]*\bname="([^"]*)"/.exec(a[0]) || [])[1] || '';
      return { xml: a[0], row, name, formula: Object.prototype.hasOwnProperty.call(LEGENDS, name) };
    });
    const formulas = anchors.filter((a) => a.formula).sort((a, b) => a.row - b.row);
    if (formulas.length === 0) continue;
    // 설명 그림 → 바로 위의 수식 묶음
    const sections = new Map(); // 수식 anchor → 설명 그림들
    for (const a of anchors) {
      if (a.formula || !Number.isFinite(a.row)) continue;
      const owner = [...formulas].reverse().find((f) => f.row <= a.row);
      if (!owner) continue;
      if (!sections.has(owner)) sections.set(owner, []);
      sections.get(owner).push(a);
    }
    let nextId = Math.max(0, ...[...drawing.matchAll(/<xdr:cNvPr\b[^>]*\bid="(\d+)"/g)].map((x) => Number(x[1]))) + 1;
    let out = drawing;
    const added = [];
    for (const [owner, legends] of sections) {
      const def = LEGENDS[owner.name];
      const row = Math.min(...legends.map((a) => a.row));
      for (const a of legends) out = out.replace(a.xml, () => '');
      added.push(legendAnchor(nextId++, `설명 ${owner.name} L`, leftCol, row, def.left, fontPt));
      added.push(legendAnchor(nextId++, `설명 ${owner.name} R`, rightCol, row, def.right, fontPt));
      count += 1;
    }
    if (added.length) {
      out = out.replace(/<\/xdr:wsDr>\s*$/, () => `${added.join('')}</xdr:wsDr>`);
      zip.file(path, out);
    }
  }
  return count;
}
