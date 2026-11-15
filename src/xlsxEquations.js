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
  const math = `<m:oMathPara xmlns:m="${M}"><m:oMathParaPr><m:jc m:val="centerGroup"/></m:oMathParaPr><m:oMath xmlns:m="${M}">${render(nodes, sz)}</m:oMath></m:oMathPara>`;
  const shape = `<xdr:sp macro="" textlink=""><xdr:nvSpPr><xdr:cNvPr id="${id}" name="${name}"/><xdr:cNvSpPr txBox="1"/></xdr:nvSpPr>`
    + `<xdr:spPr>${xfrm}<a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:noFill/><a:ln><a:noFill/></a:ln></xdr:spPr>`
    + `<xdr:txBody><a:bodyPr vertOverflow="overflow" horzOverflow="overflow" wrap="none" lIns="0" tIns="0" rIns="0" bIns="0" rtlCol="0" anchor="ctr"><a:noAutofit/></a:bodyPr><a:lstStyle/>`
    + `<a:p><a:pPr algn="ctr"/><a14:m>${math}</a14:m></a:p></xdr:txBody></xdr:sp>`;
  return `<mc:AlternateContent xmlns:mc="${MC}"><mc:Choice xmlns:a14="${A14}" Requires="a14">${shape}</mc:Choice><mc:Fallback>${picXml}</mc:Fallback></mc:AlternateContent>`;
};

// 시트에 딸린 그림 파일에서, 이름이 equations에 있는 그림을 수식 개체로 바꾼다. 바꾼 개수를 돌려준다
export async function replacePicturesWithEquations(zip, sheetPath, equations = CALC_SHEET_EQUATIONS, { fontPt = 20 } = {}) {
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
      return `<xdr:${kind}${attrs}>${body.replace(pic[0], () => equationShape(pic[0], equations[name], Math.round(fontPt * 100)))}</xdr:${kind}>`;
    });
    if (next !== drawing) zip.file(stack.join('/'), next);
  }
  return count;
}
