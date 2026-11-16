// 엑셀 서식 정리: 내보내기 마지막에 모든 시트에 적용한다.
//  1) 숫자 표시 형식의 꼬리 공백("0.00_ ")·음수 빨간 괄호를 없앤다 → 가운데 정렬한 숫자가 칸마다 어긋나지 않게
//  2) 보이는 시트는 한 장 너비(onePageSheets면 한 장)에 맞춰 인쇄 (오른쪽이 잘려 다음 장으로 넘어가지 않게)
//  3) 표 테두리: 바깥선은 굵게(medium), 안쪽선은 가늘게(thin). 이미 그어진 선의 굵기만 맞추고 새로 긋지는 않는다
//  4) 시트별 규칙: 같은 표의 행 높이 맞추기, 같은 열의 숫자 자릿수 맞추기
const attr = (attrs, name) => {
  const m = new RegExp(`\\b${name}="([^"]*)"`).exec(attrs);
  return m ? m[1] : null;
};
const decodeXml = (text) => String(text).replace(/&quot;/g, '"').replace(/&lt;/g, '<').replace(/&gt;/g, '>').replace(/&amp;/g, '&');
const encodeAttr = (text) => String(text).replace(/&/g, '&amp;').replace(/"/g, '&quot;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
const colToNum = (letters) => [...letters].reduce((n, ch) => n * 26 + ch.charCodeAt(0) - 64, 0);
const numToColName = (n) => {
  let out = '';
  while (n > 0) { const m = (n - 1) % 26; out = String.fromCharCode(65 + m) + out; n = Math.floor((n - 1) / 26); }
  return out;
};
const parseRef = (ref) => {
  const m = /^\$?([A-Z]{1,3})\$?(\d+)$/.exec(String(ref).trim());
  return m ? { col: colToNum(m[1]), row: Number(m[2]) } : null;
};
const parseRange = (range) => {
  const [a, b] = String(range).split(':');
  const from = parseRef(a); const to = parseRef(b || a);
  return from && to ? { from, to } : null;
};

// "0.00_ " / "0.0_);[Red]\(0.0\)" / "#,##0.00_ " → "0.00" 처럼 숫자 부분만 남긴다
export const cleanNumberFormat = (code) => {
  const text = String(code);
  const core = text.split(';')[0].replace(/_[ )]/g, '').replace(/\\ /g, '').trim();
  return /^(?:#,##)?0(?:\.0+)?$/.test(core) ? core : text;
};

class StyleBook {
  constructor(xml) {
    this.xml = xml;
    this.numFmts = new Map(); // id → code
    for (const m of xml.matchAll(/<numFmt\b([^>]*)\/>/g)) this.numFmts.set(Number(attr(m[1], 'numFmtId')), decodeXml(attr(m[1], 'formatCode') || ''));
    const xfsMatch = /<cellXfs\b[^>]*>([\s\S]*?)<\/cellXfs>/.exec(xml);
    this.xfs = xfsMatch ? [...xfsMatch[1].matchAll(/<xf\b[^>]*?(?:\/>|>[\s\S]*?<\/xf>)/g)].map((m) => m[0]) : [];
    this.added = new Map();
    const bordersMatch = /<borders\b[^>]*>([\s\S]*?)<\/borders>/.exec(xml);
    this.borders = bordersMatch ? [...bordersMatch[1].matchAll(/<border\b[^>]*?(?:\/>|>[\s\S]*?<\/border>)/g)].map((m) => m[0]) : [];
    this.bordersChanged = false;
    const fillsMatch = /<fills\b[^>]*>([\s\S]*?)<\/fills>/.exec(xml);
    this.fills = fillsMatch ? [...fillsMatch[1].matchAll(/<fill\b[^>]*?(?:\/>|>[\s\S]*?<\/fill>)/g)].map((m) => m[0]) : [];
    const fontsMatch = /<fonts\b[^>]*>([\s\S]*?)<\/fonts>/.exec(xml);
    this.fonts = fontsMatch ? [...fontsMatch[1].matchAll(/<font\b[^>]*?(?:\/>|>[\s\S]*?<\/font>)/g)].map((m) => m[0]) : [];
    this.fontsChanged = false;
  }

  numFmtIdOf(s) {
    const xf = this.xfs[s] || this.xfs[0] || '';
    return Number(attr(xf, 'numFmtId') || 0);
  }

  // 표시 형식 번호의 종류: 'general' | 'number' | 'date' | 'text' | 'other'
  formatKind(id) {
    if (id === 0) return 'general';
    if (id === 49) return 'text';
    if ((id >= 14 && id <= 22) || (id >= 27 && id <= 36) || (id >= 45 && id <= 47) || (id >= 50 && id <= 58)) return 'date';
    if (id >= 1 && id <= 13) return 'number';
    const code = this.numFmts.get(id);
    if (code === undefined) return 'other';
    if (code === '@') return 'text';
    if (/[ymdhs]/i.test(code.replace(/"[^"]*"/g, '').replace(/\[[^\]]*\]/g, ''))) return 'date';
    if (/[0#]/.test(code)) return 'number';
    return 'other';
  }

  fontOf(s) {
    const xf = this.xfs[s] || this.xfs[0] || '';
    const font = this.fonts[Number(attr(xf, 'fontId') || 0)] || '';
    const name = /<name\b[^>]*val="([^"]*)"/.exec(font);
    const sz = /<sz\b[^>]*val="([\d.]+)"/.exec(font);
    return { name: name ? decodeXml(name[1]) : '', sz: sz ? Number(sz[1]) : 11 };
  }

  isWrapped(s) {
    const xf = this.xfs[s] || this.xfs[0] || '';
    return /wrapText="(?:1|true)"/.test(xf);
  }

  alignmentOf(s) {
    const xf = this.xfs[s] || this.xfs[0] || '';
    const a = /<alignment\b([^>]*)\/>/.exec(xf);
    return { horizontal: (a && attr(a[1], 'horizontal')) || 'general', vertical: (a && attr(a[1], 'vertical')) || 'bottom' };
  }

  // 원래 스타일을 복사해 글꼴 이름·크기만 바꾼 스타일 번호
  withFontChange(s, change) {
    const key = `${s}|f|${change.name || ''}|${change.sz || ''}`;
    if (this.added.has(key)) return this.added.get(key);
    const xf = this.xfs[s] || this.xfs[0];
    let font = this.fonts[Number(attr(xf, 'fontId') || 0)] || '<font/>';
    if (change.name) font = /<name\b/.test(font) ? font.replace(/<name\b[^>]*\/>/, () => `<name val="${encodeAttr(change.name)}"/>`) : font.replace(/<\/font>$/, () => `<name val="${encodeAttr(change.name)}"/></font>`);
    if (change.sz) font = /<sz\b/.test(font) ? font.replace(/<sz\b[^>]*\/>/, `<sz val="${change.sz}"/>`) : font.replace(/^<font\b([^>]*)>/, `<font$1><sz val="${change.sz}"/>`);
    // 한글 글꼴은 charset 129 (없으면 그대로)
    this.fonts.push(font);
    this.fontsChanged = true;
    const fontId = this.fonts.length - 1;
    const next = xf.replace(/\s*fontId="\d+"/, '').replace(/\s*applyFont="\d+"/, '').replace(/^<xf\b/, `<xf fontId="${fontId}" applyFont="1"`);
    this.xfs.push(next);
    const index = this.xfs.length - 1;
    this.added.set(key, index);
    return index;
  }

  // 원래 스타일을 복사해 가로·세로 정렬을 바꾼 스타일 번호 (줄바꿈 등 다른 정렬 속성은 그대로)
  withAlignment(s, change) {
    const key = `${s}|a|${change.horizontal || ''}|${change.vertical || ''}`;
    if (this.added.has(key)) return this.added.get(key);
    const base = this.xfs[s] || this.xfs[0];
    const set = (a) => {
      let out = a;
      if (change.horizontal) out = `${out.replace(/\s*horizontal="[^"]*"/, '')} horizontal="${change.horizontal}"`;
      if (change.vertical) out = `${out.replace(/\s*vertical="[^"]*"/, '')} vertical="${change.vertical}"`;
      return out;
    };
    let xf;
    if (/<alignment\b[^>]*\/>/.test(base)) xf = base.replace(/<alignment\b([^>]*)\/>/, (w, a) => `<alignment${set(a)}/>`);
    else if (/\/>$/.test(base)) xf = base.replace(/\s*\/>$/, `><alignment${set('')}/></xf>`);
    else xf = base.replace(/<\/xf>$/, `<alignment${set('')}/></xf>`);
    xf = xf.replace(/\s*applyAlignment="\d+"/, '').replace(/^<xf\b/, '<xf applyAlignment="1"');
    this.xfs.push(xf);
    const index = this.xfs.length - 1;
    this.added.set(key, index);
    return index;
  }

  // 원래 스타일을 복사해 표시 형식 번호만 바꾼 스타일 번호
  withNumFmtId(s, id) {
    const key = `${s}|n|${id}`;
    if (this.added.has(key)) return this.added.get(key);
    const base = this.xfs[s] || this.xfs[0];
    const xf = base.replace(/\s*numFmtId="\d+"/, '').replace(/\s*applyNumberFormat="\d+"/, '').replace(/^<xf\b/, `<xf numFmtId="${id}" applyNumberFormat="1"`);
    this.xfs.push(xf);
    const index = this.xfs.length - 1;
    this.added.set(key, index);
    return index;
  }

  // 칸에 단색 채우기가 있는지 (머리줄 판단용)
  isFilled(s) {
    const xf = this.xfs[s] || this.xfs[0] || '';
    const fill = this.fills[Number(attr(xf, 'fillId') || 0)] || '';
    return /patternType="solid"/.test(fill);
  }

  // 글씨가 굵은지 (머리줄 판단용)
  isBold(s) {
    const xf = this.xfs[s] || this.xfs[0] || '';
    const font = this.fonts[Number(attr(xf, 'fontId') || 0)] || '';
    return /<b(?:\s*\/>|\s+val="(?:1|true)"\s*\/>)/.test(font);
  }

  // 원래 스타일을 복사해 세로 정렬만 바꾼 스타일 번호 (소제목 위에 여백을 줄 때 글자를 아래로 붙인다)
  withVertical(s, vertical) {
    const key = `${s}|v|${vertical}`;
    if (this.added.has(key)) return this.added.get(key);
    const base = this.xfs[s] || this.xfs[0];
    let xf;
    if (/<alignment\b[^>]*\/>/.test(base)) {
      xf = base.replace(/<alignment\b([^>]*)\/>/, (w, a) => `<alignment${a.replace(/\s*vertical="[^"]*"/, '')} vertical="${vertical}"/>`);
    } else if (/\/>$/.test(base)) {
      xf = base.replace(/\s*\/>$/, `><alignment vertical="${vertical}"/></xf>`);
    } else {
      xf = base.replace(/<\/xf>$/, `<alignment vertical="${vertical}"/></xf>`);
    }
    xf = xf.replace(/\s*applyAlignment="\d+"/, '').replace(/^<xf\b/, '<xf applyAlignment="1"');
    this.xfs.push(xf);
    const index = this.xfs.length - 1;
    this.added.set(key, index);
    return index;
  }

  // 칸 스타일(s)의 네 변 선 종류 { left, right, top, bottom } (선 없음은 null)
  sidesOf(s) {
    const xf = this.xfs[s] || this.xfs[0] || '';
    const border = this.borders[Number(attr(xf, 'borderId') || 0)] || '';
    const side = (name) => {
      const m = new RegExp(`<${name}\\b([^>]*?)(?:/>|>)`).exec(border);
      return m ? attr(m[1], 'style') : null;
    };
    return { left: side('left'), right: side('right'), top: side('top'), bottom: side('bottom') };
  }

  // 원래 스타일을 복사해 일부 변의 선 굵기만 바꾼 스타일 번호. changes = { left: 'thin', ... }
  // 같은 서식에서 채우기 색만 바꾼 서식 번호 (rgb: 'EDEDED')
  withFill(s, rgb) {
    const key = `${s}|fill|${rgb}`;
    if (this.added.has(key)) return this.added.get(key);
    const xf = this.xfs[s] || this.xfs[0];
    this.fills.push(`<fill><patternFill patternType="solid"><fgColor rgb="FF${rgb}"/><bgColor indexed="64"/></patternFill></fill>`);
    this.fillsChanged = true;
    const fillId = this.fills.length - 1;
    const next = xf.replace(/\s*fillId="\d+"/, '').replace(/\s*applyFill="\d+"/, '').replace(/^<xf\b/, `<xf fillId="${fillId}" applyFill="1"`);
    this.xfs.push(next);
    const index = this.xfs.length - 1;
    this.added.set(key, index);
    return index;
  }
  // 양식의 채우기 색을 기록부 색 테마로 바꾼다: 노랑(머리칸) → header, 연두(입력칸) → input
  recolorFills({ header, input } = {}) {
    const swap = (fill, rgb) => fill.replace(/<fgColor\b[^>]*\/>/, `<fgColor rgb="FF${rgb}"/>`);
    this.fills = this.fills.map((fill) => {
      if (header && /<fgColor\b[^>]*(?:rgb="FFFFFF99"|indexed="43"|rgb="FFFFFF00"|indexed="13"|indexed="11"|rgb="FF00FF00"|rgb="FF00FFFF"|indexed="15")/.test(fill)) { this.fillsChanged = true; return swap(fill, header); }
      if (input && /<fgColor\b[^>]*(?:rgb="FFCCFFCC"|indexed="42"|rgb="FFE2EFDA")/.test(fill)) { this.fillsChanged = true; return swap(fill, input); }
      return fill;
    });
  }
  // 같은 서식에 '셀에 맞춤'(shrinkToFit)을 켠 서식 번호: 칸보다 긴 글자·숫자를 넘치거나 ####로 보이지 않게 줄여 보여 준다
  withShrink(s, force = false) {
    const key = `${s}|shrink${force ? '!' : ''}`;
    if (this.added.has(key)) return this.added.get(key);
    const base = this.xfs[s] || this.xfs[0];
    if (!force && (/\bwrapText="(?:1|true)"/.test(base) || /\bshrinkToFit="(?:1|true)"/.test(base))) return s;
    // force: 줄 바꿈을 끄고 한 줄로 줄여 보여 준다
    const base2 = force ? base.replace(/\s*\bwrapText="[^"]*"/, '').replace(/\s*\bshrinkToFit="[^"]*"/, '') : base;
    let xf;
    if (/<alignment\b[^>]*\/>/.test(base2)) xf = base2.replace(/<alignment\b([^>]*)\/>/, (w, a) => `<alignment${a} shrinkToFit="1"/>`);
    else if (/\/>$/.test(base2)) xf = base2.replace(/\s*\/>$/, '><alignment shrinkToFit="1"/></xf>');
    else xf = base2.replace(/<\/xf>$/, '<alignment shrinkToFit="1"/></xf>');
    xf = xf.replace(/\s*applyAlignment="\d+"/, '').replace(/^<xf\b/, '<xf applyAlignment="1"');
    this.xfs.push(xf);
    const index = this.xfs.length - 1;
    this.added.set(key, index);
    return index;
  }
  // 같은 서식에서 채우기만 없앤 서식 번호
  withNoFill(s) {
    const key = `${s}|nofill`;
    if (this.added.has(key)) return this.added.get(key);
    const xf = this.xfs[s] || this.xfs[0];
    if (!Number(attr(xf, 'fillId') || 0)) return s;
    this.xfs.push(xf.replace(/\s*fillId="\d+"/, '').replace(/^<xf\b/, '<xf fillId="0"'));
    const index = this.xfs.length - 1;
    this.added.set(key, index);
    return index;
  }
  // 같은 서식에서 지정한 변의 선만 지운 서식 번호
  withoutBorderSides(s, sides) {
    const names = ['left', 'right', 'top', 'bottom'].filter((k) => sides.includes(k));
    if (names.length === 0) return s;
    const key = `${s}|nb|${names.join(',')}`;
    if (this.added.has(key)) return this.added.get(key);
    const xf = this.xfs[s] || this.xfs[0];
    let border = this.borders[Number(attr(xf, 'borderId') || 0)] || '<border/>';
    for (const name of names) border = border.replace(new RegExp(`<${name}\\b[^>]*?(?:/>|>[\\s\\S]*?</${name}>)`), () => `<${name}/>`);
    this.borders.push(border);
    this.bordersChanged = true;
    const borderId = this.borders.length - 1;
    const next = xf.replace(/\s*borderId="\d+"/, '').replace(/\s*applyBorder="\d+"/, '').replace(/^<xf\b/, `<xf borderId="${borderId}" applyBorder="1"`);
    this.xfs.push(next);
    const index = this.xfs.length - 1;
    this.added.set(key, index);
    return index;
  }
  withBorderStyles(s, changes) {
    const key = `${s}|b|${['left', 'right', 'top', 'bottom'].map((k) => changes[k] || '').join(',')}`;
    if (this.added.has(key)) return this.added.get(key);
    const xf = this.xfs[s] || this.xfs[0];
    let border = this.borders[Number(attr(xf, 'borderId') || 0)] || '<border/>';
    for (const [name, style] of Object.entries(changes)) {
      border = border.replace(new RegExp(`<${name}\\b([^>]*?)(/?)>`), (w, a, close) => `<${name}${a.replace(/\s*style="[^"]*"/, '')} style="${style}"${close}>`);
    }
    this.borders.push(border);
    this.bordersChanged = true;
    const borderId = this.borders.length - 1;
    const next = xf.replace(/\s*borderId="\d+"/, '').replace(/\s*applyBorder="\d+"/, '').replace(/^<xf\b/, `<xf borderId="${borderId}" applyBorder="1"`);
    this.xfs.push(next);
    const index = this.xfs.length - 1;
    this.added.set(key, index);
    return index;
  }

  // 표시 형식 코드에서 꼬리 공백·빨간 괄호를 지운다 (모든 칸에 함께 적용)
  cleanAll() {
    this.xml = this.xml.replace(/<numFmt\b([^>]*)\/>/g, (whole, attrs) => {
      const code = decodeXml(attr(attrs, 'formatCode') || '');
      const next = cleanNumberFormat(code);
      if (next === code) return whole;
      this.numFmts.set(Number(attr(attrs, 'numFmtId')), next);
      return `<numFmt numFmtId="${attr(attrs, 'numFmtId')}" formatCode="${encodeAttr(next)}"/>`;
    });
  }

  fmtId(code) {
    const builtin = { General: 0, '0': 1, '0.00': 2 };
    if (code in builtin) return builtin[code];
    for (const [id, c] of this.numFmts) if (c === code) return id;
    const id = Math.max(163, ...this.numFmts.keys()) + 1;
    this.numFmts.set(id, code);
    const tag = `<numFmt numFmtId="${id}" formatCode="${encodeAttr(code)}"/>`;
    if (/<numFmts\b[^>]*>/.test(this.xml)) {
      this.xml = this.xml.replace(/<numFmts\b([^>]*)>/, (w, a) => `<numFmts${a.replace(/\s*count="\d+"/, '')} count="${this.numFmts.size}">${tag}`);
    } else {
      this.xml = this.xml.replace(/(<styleSheet\b[^>]*>)/, (w) => `${w}<numFmts count="1">${tag}</numFmts>`);
    }
    return id;
  }

  // 원래 스타일(s)을 복사해 표시 형식만 바꾼 스타일 번호
  withFormat(s, code) {
    const key = `${s}|${code}`;
    if (this.added.has(key)) return this.added.get(key);
    const base = this.xfs[s] || this.xfs[0];
    const id = this.fmtId(code);
    const xf = base
      .replace(/\s*numFmtId="\d+"/, '')
      .replace(/\s*applyNumberFormat="\d+"/, '')
      .replace(/^<xf\b/, `<xf numFmtId="${id}" applyNumberFormat="1"`);
    this.xfs.push(xf);
    const index = this.xfs.length - 1;
    this.added.set(key, index);
    return index;
  }

  // 원래 스타일을 복사해 가로 가운데 정렬만 바꾼 스타일 번호
  centered(s) {
    const key = `${s}|center`;
    if (this.added.has(key)) return this.added.get(key);
    const base = this.xfs[s] || this.xfs[0];
    let xf;
    if (/<alignment\b[^>]*\/>/.test(base)) {
      xf = base.replace(/<alignment\b([^>]*)\/>/, (w, a) => `<alignment${a.replace(/\s*horizontal="[^"]*"/, '')} horizontal="center"/>`);
    } else if (/\/>$/.test(base)) {
      xf = base.replace(/\s*\/>$/, '><alignment horizontal="center"/></xf>');
    } else {
      xf = base.replace(/<\/xf>$/, '<alignment horizontal="center"/></xf>');
    }
    xf = xf.replace(/\s*applyAlignment="\d+"/, '').replace(/^<xf\b/, '<xf applyAlignment="1"');
    this.xfs.push(xf);
    const index = this.xfs.length - 1;
    this.added.set(key, index);
    return index;
  }

  serialize() {
    let out = this.xml;
    if (this.fontsChanged) out = out.replace(/<fonts\b[^>]*>[\s\S]*?<\/fonts>/, () => `<fonts count="${this.fonts.length}">${this.fonts.join('')}</fonts>`);
    if (this.fillsChanged) out = out.replace(/<fills\b[^>]*>[\s\S]*?<\/fills>/, () => `<fills count="${this.fills.length}">${this.fills.join('')}</fills>`);
    if (this.bordersChanged) out = out.replace(/<borders\b[^>]*>[\s\S]*?<\/borders>/, () => `<borders count="${this.borders.length}">${this.borders.join('')}</borders>`);
    if (this.added.size === 0) return out;
    return out.replace(/<cellXfs\b[^>]*>[\s\S]*?<\/cellXfs>/, () => `<cellXfs count="${this.xfs.length}">${this.xfs.join('')}</cellXfs>`);
  }
}

// 한 장 너비에 맞춤: <sheetPr><pageSetUpPr fitToPage="1"/></sheetPr> + <pageSetup fitToWidth="1" fitToHeight="0">
const fitToPageWidth = (xml, { onePage = false } = {}) => {
  let out = ensurePageElements(xml);
  if (/<sheetPr\b[^>]*\/>/.test(out)) out = out.replace(/<sheetPr\b([^>]*)\/>/, '<sheetPr$1><pageSetUpPr fitToPage="1"/></sheetPr>');
  else if (/<sheetPr\b[^>]*>/.test(out)) {
    if (/<pageSetUpPr\b[^>]*\/>/.test(out)) out = out.replace(/<pageSetUpPr\b([^>]*)\/>/, (w, a) => `<pageSetUpPr${a.replace(/\s*fitToPage="[^"]*"/, '')} fitToPage="1"/>`);
    else out = out.replace(/<\/sheetPr>/, '<pageSetUpPr fitToPage="1"/></sheetPr>');
  } else out = out.replace(/(<worksheet\b[^>]*>)/, '$1<sheetPr><pageSetUpPr fitToPage="1"/></sheetPr>');
  const fit = `fitToWidth="1" fitToHeight="${onePage ? 1 : 0}"`;
  // 시트 수준 pageSetup만 (사용자 지정 보기 안의 pageSetup은 그대로)
  const [head, body] = splitAtCustomViews(out);
  let tail = body;
  if (/<pageSetup\b[^>]*\/>/.test(tail)) {
    tail = tail.replace(/<pageSetup\b([^>]*)\/>/, (w, a) => `<pageSetup${a.replace(/\s*fitTo(?:Width|Height)="[^"]*"/g, '')} ${fit}/>`);
  } else if (/<pageMargins\b[^>]*\/>/.test(tail)) {
    tail = tail.replace(/(<pageMargins\b[^>]*\/>)/, (w) => `${w}<pageSetup paperSize="9" orientation="portrait" ${fit}/>`);
  }
  return head + tail;
};

const setRowHeights = (xml, heights) => {
  let out = xml;
  for (const [rowText, ht] of Object.entries(heights)) {
    const row = Number(rowText);
    const re = new RegExp(`<row\\b([^>]*?)\\br="${row}"([^>]*?)(/?)>`);
    if (!re.test(out)) continue;
    out = out.replace(re, (w, a, b, selfClose) => {
      const attrs = `${a}r="${row}"${b}`.replace(/\s*ht="[^"]*"/, '').replace(/\s*customHeight="[^"]*"/, '');
      return `<row${attrs} ht="${ht}" customHeight="1"${selfClose}>`;
    });
  }
  return out;
};

const setCentered = (xml, styles, ranges) => {
  const targets = ranges.map(parseRange).filter(Boolean);
  if (targets.length === 0) return xml;
  return xml.replace(/<c\b([^>]*?)(\/?)>/g, (whole, attrs, selfClose) => {
    const ref = parseRef(attr(attrs, 'r') || '');
    if (!ref || !targets.some((r) => ref.row >= r.from.row && ref.row <= r.to.row && ref.col >= r.from.col && ref.col <= r.to.col)) return whole;
    const next = styles.centered(Number(attr(attrs, 's') || 0));
    return `<c${attrs.replace(/\s*\bs="\d+"/, '')} s="${next}"${selfClose}>`;
  });
};

const setNumberFormats = (xml, styles, formats) => {
  const targets = formats.map(({ range, code }) => ({ r: parseRange(range), code })).filter((x) => x.r);
  if (targets.length === 0) return xml;
  return xml.replace(/<c\b([^>]*?)(\/?)>/g, (whole, attrs, selfClose) => {
    const ref = parseRef(attr(attrs, 'r') || '');
    if (!ref) return whole;
    const hit = targets.find(({ r }) => ref.row >= r.from.row && ref.row <= r.to.row && ref.col >= r.from.col && ref.col <= r.to.col);
    if (!hit) return whole;
    const s = Number(attr(attrs, 's') || 0);
    const next = styles.withFormat(s, hit.code);
    const cleaned = attrs.replace(/\s*\bs="\d+"/, '');
    return `<c${cleaned} s="${next}"${selfClose}>`;
  });
};

// 표 테두리 맞추기
//  - 표 = 세 변 이상이 선으로 둘러싸인 칸이 상하좌우로 이어진 덩어리 (병합 칸은 통째로, 표 안에 갇힌 빈칸 포함).
//    안쪽에 세로선과 가로선이 모두 있어야 표로 본다 (밑줄만 있는 서식 칸은 표가 아니다)
//  - 표 바깥과 맞닿은 변은 medium, 표 안쪽 변은 thin. 강조선(머리줄 아래, 평균 줄 위)과 양식의 이중선 자리도 medium으로 통일
const EMPHASIS_STYLES = new Set(['double']);
const cellKind = (attrs, inner) => {
  if (!inner) return 'empty';
  if (/<f\b/.test(inner)) return 'formula';
  const t = attr(attrs, 't');
  if (t === 's' || t === 'inlineStr' || t === 'str') return /<(?:v|t)\b[^>]*>[^<]/.test(inner) ? 'text' : 'empty';
  return /<v>[^<]/.test(inner) ? 'number' : 'empty';
};

// 시트의 표 분석: 칸 스타일·값 종류·글자, 병합, 선, 표 칸(member), 표 덩어리(tables)
export const analyzeTables = (xml, styles, sharedStrings = []) => {
  const cells = new Map(); // "r,c" → s
  const kinds = new Map(); // "r,c" → 'text' | 'number' | 'formula' | 'empty'
  const texts = new Map(); // "r,c" → 글자 (평균 줄 찾기용)
  const vtypes = new Map(); // "r,c" → 계산값까지 본 값 종류 'number' | 'text' | 'bool' | 'error' | 'empty'
  const raws = new Map(); // "r,c" → <v> 원문 (숫자 자릿수 판단용)
  const formulas = new Map(); // "r,c" → 수식 글
  for (const m of xml.matchAll(/<c\b([^>]*?)(?:\/>|>([\s\S]*?)<\/c>)/g)) {
    const ref = parseRef(attr(m[1], 'r') || '');
    if (!ref) continue;
    cells.set(`${ref.row},${ref.col}`, Number(attr(m[1], 's') || 0));
    kinds.set(`${ref.row},${ref.col}`, cellKind(m[1], m[2]));
    const inner = m[2] || '';
    if (attr(m[1], 't') === 's') { const v = /<v>(\d+)<\/v>/.exec(inner); if (v) texts.set(`${ref.row},${ref.col}`, sharedStrings[Number(v[1])] || ''); }
    else if (attr(m[1], 't') === 'inlineStr') texts.set(`${ref.row},${ref.col}`, [...inner.matchAll(/<t\b[^>]*>([^<]*)<\/t>/g)].map((x) => decodeXml(x[1])).join(''));
    const t = attr(m[1], 't');
    const v = /<v>([^<]*)<\/v>/.exec(inner);
    const f = /<f\b[^>]*>([^<]*)<\/f>/.exec(inner) || (/<f\b[^>]*\/>/.test(inner) ? ['', ''] : null);
    if (f) formulas.set(`${ref.row},${ref.col}`, decodeXml(f[1]));
    let vt = 'empty';
    if (t === 's' || t === 'inlineStr') vt = texts.get(`${ref.row},${ref.col}`) ? 'text' : 'empty';
    else if (t === 'str') vt = v && v[1] !== '' ? 'text' : 'empty';
    else if (t === 'b') vt = 'bool';
    else if (t === 'e') vt = 'error';
    else if (v && v[1] !== '') { vt = 'number'; raws.set(`${ref.row},${ref.col}`, v[1]); }
    if (t === 'str' && v) texts.set(`${ref.row},${ref.col}`, decodeXml(v[1]));
    vtypes.set(`${ref.row},${ref.col}`, vt);
  }
  const mergeOf = new Map(); // "r,c" → 병합 영역 번호
  const merges = [];
  for (const m of xml.matchAll(/<mergeCell\b[^>]*ref="([^"]+)"/g)) {
    const r = parseRange(m[1]);
    if (!r) continue;
    merges.push(r);
    for (let row = r.from.row; row <= r.to.row; row++) for (let col = r.from.col; col <= r.to.col; col++) mergeOf.set(`${row},${col}`, merges.length - 1);
  }
  const K = (row, col) => `${row},${col}`;
  const sameMerge = (a, b) => mergeOf.has(a) && mergeOf.get(a) === mergeOf.get(b);
  const sideCache = new Map();
  const own = (key) => {
    if (!cells.has(key)) return { left: null, right: null, top: null, bottom: null };
    const s = cells.get(key);
    if (!sideCache.has(s)) sideCache.set(s, styles.sidesOf(s));
    return sideCache.get(s);
  };
  // 선이 그어진 경계 (자기 칸이든 옆 칸이든 어느 한쪽에 있으면 그어진 것). 병합 칸 안쪽 경계는 그려지지 않는다
  const hLine = (row, col) => !sameMerge(K(row - 1, col), K(row, col)) && Boolean(own(K(row - 1, col)).bottom || own(K(row, col)).top);
  const vLine = (row, col) => !sameMerge(K(row, col - 1), K(row, col)) && Boolean(own(K(row, col - 1)).right || own(K(row, col)).left);
  const sideCount = (row, col) => [vLine(row, col), vLine(row, col + 1), hLine(row, col), hLine(row + 1, col)].filter(Boolean).length;

  // 표 칸: 세 변 이상이 선으로 둘러싸인 칸 (병합 칸은 영역 둘레로 따진다)
  const candidates = new Set();
  for (const key of cells.keys()) {
    const x = own(key);
    if (!(x.left || x.right || x.top || x.bottom)) continue;
    const [row, col] = key.split(',').map(Number);
    candidates.add(key);
    candidates.add(K(row - 1, col)); candidates.add(K(row + 1, col)); candidates.add(K(row, col - 1)); candidates.add(K(row, col + 1));
  }
  const member = new Set();
  const mergeDone = new Set();
  for (const key of candidates) {
    const [row, col] = key.split(',').map(Number);
    if (row < 1 || col < 1) continue;
    if (mergeOf.has(key)) {
      const id = mergeOf.get(key);
      if (mergeDone.has(id)) continue;
      mergeDone.add(id);
      const r = merges[id];
      let left = false; let right = false; let top = false; let bottom = false;
      for (let rr = r.from.row; rr <= r.to.row; rr++) { left = left || vLine(rr, r.from.col); right = right || vLine(rr, r.to.col + 1); }
      for (let cc = r.from.col; cc <= r.to.col; cc++) { top = top || hLine(r.from.row, cc); bottom = bottom || hLine(r.to.row + 1, cc); }
      if ([left, right, top, bottom].filter(Boolean).length >= 3) {
        for (let rr = r.from.row; rr <= r.to.row; rr++) for (let cc = r.from.col; cc <= r.to.col; cc++) member.add(K(rr, cc));
      }
      continue;
    }
    if (sideCount(row, col) >= 3) member.add(key);
  }

  // 위아래가 선으로 닫히고 양 끝이 표 칸에 막힌 가로 줄 칸(병합 없이 이어 놓은 빈칸), 좌우가 닫히고 위아래 끝이 막힌 세로 칸도 표 칸
  if (member.size) {
    let minR = Infinity; let maxR = -Infinity; let minC = Infinity; let maxC = -Infinity;
    for (const key of member) { const [row, col] = key.split(',').map(Number); minR = Math.min(minR, row); maxR = Math.max(maxR, row); minC = Math.min(minC, col); maxC = Math.max(maxC, col); }
    const strips = [];
    for (let row = minR; row <= maxR; row++) {
      let col = minC;
      while (col <= maxC) {
        if (member.has(K(row, col)) || !(hLine(row, col) && hLine(row + 1, col))) { col += 1; continue; }
        const start = col;
        while (col <= maxC && !member.has(K(row, col)) && hLine(row, col) && hLine(row + 1, col)) col += 1;
        if (member.has(K(row, start - 1)) && member.has(K(row, col))) for (let c = start; c < col; c++) strips.push(K(row, c));
      }
    }
    for (let col = minC; col <= maxC; col++) {
      let row = minR;
      while (row <= maxR) {
        if (member.has(K(row, col)) || !(vLine(row, col) && vLine(row, col + 1))) { row += 1; continue; }
        const start = row;
        while (row <= maxR && !member.has(K(row, col)) && vLine(row, col) && vLine(row, col + 1)) row += 1;
        if (member.has(K(start - 1, col)) && member.has(K(row, col))) for (let r = start; r < row; r++) strips.push(K(r, col));
      }
    }
    strips.forEach((key) => member.add(key));
  }

  const DIRS = { left: [0, -1], right: [0, 1], top: [-1, 0], bottom: [1, 0] };
  const seen = new Set();
  const tables = [];
  for (const start of member) {
    if (seen.has(start)) continue;
    const comp = new Set([start]); const queue = [start]; seen.add(start);
    while (queue.length) {
      const key = queue.pop();
      const [row, col] = key.split(',').map(Number);
      for (const [dr, dc] of Object.values(DIRS)) {
        const next = K(row + dr, col + dc);
        if (member.has(next) && !seen.has(next)) { seen.add(next); comp.add(next); queue.push(next); }
      }
    }
    // 표 안에 갇힌 빈칸(옆 칸의 선만 빌려 쓰는 칸)도 표에 넣는다
    let r0 = Infinity; let r1 = -Infinity; let c0 = Infinity; let c1 = -Infinity;
    for (const key of comp) { const [row, col] = key.split(',').map(Number); r0 = Math.min(r0, row); r1 = Math.max(r1, row); c0 = Math.min(c0, col); c1 = Math.max(c1, col); }
    const open = new Set(); const flood = [];
    for (let row = r0; row <= r1; row++) for (let col = c0; col <= c1; col++) {
      const key = K(row, col);
      if (!comp.has(key) && (row === r0 || row === r1 || col === c0 || col === c1)) { open.add(key); flood.push(key); }
    }
    while (flood.length) {
      const [row, col] = flood.pop().split(',').map(Number);
      for (const [dr, dc] of Object.values(DIRS)) {
        const nr = row + dr; const nc = col + dc; const key = K(nr, nc);
        if (nr < r0 || nr > r1 || nc < c0 || nc > c1 || comp.has(key) || open.has(key)) continue;
        open.add(key); flood.push(key);
      }
    }
    for (let row = r0; row <= r1; row++) for (let col = c0; col <= c1; col++) {
      const key = K(row, col);
      if (!comp.has(key) && !open.has(key)) comp.add(key);
    }
    // 안쪽 선이 세로·가로 모두 있어야 표 (밑줄만 있는 서식은 표가 아니다)
    let innerV = false; let innerH = false;
    for (const key of comp) {
      const [row, col] = key.split(',').map(Number);
      if (comp.has(K(row, col + 1)) && vLine(row, col + 1)) innerV = true;
      if (comp.has(K(row + 1, col)) && hLine(row + 1, col)) innerH = true;
      if (innerV && innerH) break;
    }
    if (!innerV || !innerH) continue;
    tables.push({ comp, r0, r1, c0, c1 });
  }
  const topLeft = (key) => {
    if (!mergeOf.has(key)) return key;
    const r = merges[mergeOf.get(key)];
    return K(r.from.row, r.from.col);
  };
  return { cells, kinds, texts, vtypes, raws, formulas, mergeOf, merges, K, sameMerge, own, hLine, vLine, member, tables, topLeft, DIRS };
};

const SUMMARY_LABEL = /^(평균|합계|계|avg\.?|average|mean|total)$/i;
export const normalizeTableBorders = (xml, styles, sharedStrings = []) => {
  const { cells, kinds, texts, merges, K, sameMerge, own, hLine, vLine, member, tables, topLeft, DIRS } = analyzeTables(xml, styles, sharedStrings);
  void merges;
  const OPP = { left: 'right', right: 'left', top: 'bottom', bottom: 'top' };
  const changes = new Map(); // "r,c" → { side: style }
  const setChange = (key, side, style) => {
    const current = own(key)[side];
    if (!current || current === style) return;
    changes.set(key, { ...(changes.get(key) || {}), [side]: style });
  };
  for (const { comp, r0, r1, c0, c1 } of tables) {
    // 머리줄: 숫자·수식이 없고, 칸마다 채우기나 굵은 글씨가 있는 줄 (병합 칸은 왼쪽 위 칸으로 본다)
    // 머리줄 바로 아래(첫 자료줄 위) 선은 강조선으로 굵게 둔다
    const headerCache = new Map();
    const isHeaderRow = (row) => {
      if (headerCache.has(row)) return headerCache.get(row);
      let any = false; let hasText = false; let ok = true;
      for (let col = c0; col <= c1 && ok; col++) {
        const key = K(row, col);
        if (!member.has(key) || !comp.has(key)) continue;
        any = true;
        const head = topLeft(key);
        const kind = kinds.get(head) || 'empty';
        if (kind === 'number' || kind === 'formula') ok = false;
        else if (!cells.has(head) || !(styles.isFilled(cells.get(head)) || styles.isBold(cells.get(head)))) ok = false;
        if (kind === 'text') hasText = true;
      }
      const result = any && ok && hasText;
      headerCache.set(row, result);
      return result;
    };
    const headerEdges = new Set(); // 가로선 "r,c" = r행 위쪽 선
    for (let row = r0; row < r1; row++) {
      if (!isHeaderRow(row) || isHeaderRow(row + 1)) continue;
      for (let col = c0; col <= c1; col++) {
        const up = K(row, col); const down = K(row + 1, col);
        if (comp.has(up) && comp.has(down) && !sameMerge(up, down) && hLine(row + 1, col)) headerEdges.add(down);
      }
    }
    // 평균·합계 줄 바로 위 선도 강조선으로 굵게
    for (let row = r0 + 1; row <= r1; row++) {
      let summary = false;
      for (let col = c0; col <= c1 && !summary; col++) {
        const key = K(row, col);
        if (comp.has(key) && SUMMARY_LABEL.test(String(texts.get(key) || '').replace(/\s+/g, ''))) summary = true;
      }
      if (!summary) continue;
      for (let col = c0; col <= c1; col++) {
        const up = K(row - 1, col); const down = K(row, col);
        if (comp.has(up) && comp.has(down) && !sameMerge(up, down) && hLine(row, col)) headerEdges.add(down);
      }
    }
    for (const key of comp) {
      const [row, col] = key.split(',').map(Number);
      for (const [side, [dr, dc]] of Object.entries(DIRS)) {
        const next = K(row + dr, col + dc);
        if (sameMerge(key, next)) continue; // 병합 칸 안쪽 변은 그려지지 않는다
        const line = side === 'left' ? vLine(row, col) : side === 'right' ? vLine(row, col + 1) : side === 'top' ? hLine(row, col) : hLine(row + 1, col);
        if (!line) continue;
        const edgeKey = side === 'bottom' ? K(row + 1, col) : side === 'top' ? K(row, col) : null;
        // 양식에서 이중선으로 그어 둔 강조선 자리도 굵은 선(medium)으로
        const nextOwn = own(next)[OPP[side]];
        const emphasis = EMPHASIS_STYLES.has(own(key)[side]) || EMPHASIS_STYLES.has(nextOwn);
        const want = !comp.has(next) || emphasis || (edgeKey && headerEdges.has(edgeKey)) ? 'medium' : 'thin';
        setChange(key, side, want);
        // 표 밖 칸이 같은 선을 따로 가지고 있으면 그 칸도 같은 굵기로 (엑셀은 두 칸 중 굵은 쪽을 그린다)
        if (!comp.has(next) && cells.has(next)) setChange(next, OPP[side], want);
      }
    }
  }
  if (changes.size === 0) return xml;
  return xml.replace(/<c\b([^>]*?)(\/?)>/g, (whole, attrs, close) => {
    const ref = parseRef(attr(attrs, 'r') || '');
    if (!ref) return whole;
    const change = changes.get(K(ref.row, ref.col));
    if (!change) return whole;
    const next = styles.withBorderStyles(Number(attr(attrs, 's') || 0), change);
    return `<c${attrs.replace(/\s*\bs="\d+"/, '')} s="${next}"${close}>`;
  });
};

// 칸 정리: 글꼴 이름 통일, 표 칸 가운데 정렬(가로·세로), 표 글자 크기 통일, 숫자 칸은 숫자 서식·글자 칸은 일반 서식
//  - 가운데로 바꾸지 않는 칸: 양식번호(■…), 표 안쪽 세로선이 없는 줄(캡션·제목·양식번호 줄),
//    '이름 + (인)'·'값 + 단위'처럼 오른쪽·왼쪽으로 붙여 짝지어 둔 칸
//  - 글자 크기: 시트의 표에서 가장 많이 쓴 크기로. 큰 글자는 줄이고, 작은 글자는 칸에 들어갈 때만 키운다 (18pt 이상 제목은 그대로)
//  - 숫자 칸이 '일반' 서식이면 같은 열에서 이어지는 숫자 칸의 서식을 따르고, 없으면 값에 맞춘다 (정수 '0', 소수는 자릿수만큼)
const DATE_FUNCS = /\b(?:WORKDAY|DATE|TODAY|EDATE|EOMONTH|NOW)\s*\(/i;
const charWidth = (ch, sz) => {
  const code = ch.codePointAt(0);
  if (code >= 0x1100) return sz; // 한글·한자 등 전각
  if (ch === ' ') return sz * 0.3;
  return sz * 0.55;
};
const decimalsOf = (raw, isFormula) => {
  const x = Number(raw);
  if (!Number.isFinite(x)) return 0;
  for (let d = 0; d <= 10; d++) {
    if (Math.abs(Math.round(x * 10 ** d) / 10 ** d - x) <= 1e-9 * Math.max(1, Math.abs(x))) {
      if (isFormula && d > 4) return 2;
      return Math.min(d, 4);
    }
  }
  return isFormula ? 2 : 4;
};
export const normalizeCellStyles = (xml, styles, { sharedStrings = [], cols = null, align = true, fontName = '맑은 고딕', sizes = true, formats = true } = {}) => {
  const A = analyzeTables(xml, styles, sharedStrings);
  const { cells, kinds, texts, vtypes, raws, formulas, K, sameMerge, vLine, tables, mergeOf, merges } = A;
  const inCols = (col) => !cols || (col >= cols.from && col <= cols.to);
  // 열 너비(pt) — 글자가 칸에 들어가는지 볼 때 쓴다 (좁게 어림해 넘치지 않게)
  const fmt = /<sheetFormatPr\b([^>]*)\/?>/.exec(xml);
  const baseCol = Number((fmt && attr(fmt[1], 'baseColWidth')) || 8);
  const defaultCol = Number((fmt && attr(fmt[1], 'defaultColWidth')) || baseCol + 0.71);
  const colDefs = [...xml.matchAll(/<col\b([^>]*)\/?>/g)].map((m) => ({ min: Number(attr(m[1], 'min')), max: Number(attr(m[1], 'max')), width: Number(attr(m[1], 'width') || defaultCol) }));
  const colPt = (col) => { const d = colDefs.find((c) => col >= c.min && col <= c.max); return Math.trunc(((256 * (d ? d.width : defaultCol) + Math.trunc(128 / 7)) / 256) * 7) * 0.75; };
  const cellWidthPt = (key) => {
    const [row, col] = key.split(',').map(Number);
    if (mergeOf.has(key)) { const r = merges[mergeOf.get(key)]; let w = 0; for (let c = r.from.col; c <= r.to.col; c++) w += colPt(c); return w; }
    void row;
    return colPt(col);
  };
  const textOf = (key) => (texts.has(key) ? texts.get(key) : raws.get(key) || '');
  const isFooter = (key) => /^\s*■/.test(textOf(key));

  // 표 칸 가운데 정렬 대상
  const alignKeys = new Set();
  const sizeKeys = new Set();
  for (const { comp, r0, r1, c0, c1 } of tables) {
    for (let row = r0; row <= r1; row++) {
      let inner = false;
      for (let col = c0; col < c1 && !inner; col++) {
        if (comp.has(K(row, col)) && comp.has(K(row, col + 1)) && vLine(row, col + 1)) inner = true;
      }
      if (!inner) continue; // 캡션·제목·양식번호 줄
      for (let col = c0; col <= c1; col++) {
        const key = K(row, col);
        if (!comp.has(key) || !cells.has(key) || !inCols(col) || isFooter(key)) continue;
        if (mergeOf.has(key) && A.topLeft(key) !== key) continue;
        alignKeys.add(key);
        if (vtypes.get(key) && vtypes.get(key) !== 'empty') sizeKeys.add(key);
      }
    }
  }
  // '이름 + (인)'·'값 + 단위'처럼 짝지어 붙여 둔 칸은 정렬을 그대로 둔다
  for (const key of [...alignKeys]) {
    const [row, col] = key.split(',').map(Number);
    const right = K(row, col + 1);
    const here = styles.alignmentOf(cells.get(key)).horizontal;
    if (here === 'right' && vtypes.get(key) !== 'empty' && cells.has(right) && vtypes.get(right) && vtypes.get(right) !== 'empty'
      && ['general', 'left'].includes(styles.alignmentOf(cells.get(right)).horizontal) && !sameMerge(key, right)) {
      alignKeys.delete(key); alignKeys.delete(right); sizeKeys.delete(key); sizeKeys.delete(right);
    }
  }
  // 표 글자 크기: 가장 많이 쓴 크기
  const sizeCount = new Map();
  for (const key of sizeKeys) { const sz = styles.fontOf(cells.get(key)).sz; if (sz < 18) sizeCount.set(sz, (sizeCount.get(sz) || 0) + 1); }
  const majority = [...sizeCount].sort((a, b) => b[1] - a[1] || b[0] - a[0])[0]?.[0];

  const changes = new Map(); // key → { font: {name, sz}, align: {h, v}, numFmt }
  const change = (key) => { if (!changes.has(key)) changes.set(key, {}); return changes.get(key); };
  // 1) 글꼴 이름 통일 (값이 있는 칸)
  if (fontName) {
    for (const [key, s] of cells) {
      const col = Number(key.split(',')[1]);
      if (!inCols(col) || !vtypes.get(key) || vtypes.get(key) === 'empty') continue;
      if (styles.fontOf(s).name !== fontName) change(key).font = { ...(change(key).font || {}), name: fontName };
    }
  }
  // 2) 가운데 정렬
  if (align) {
    for (const key of alignKeys) {
      const a = styles.alignmentOf(cells.get(key));
      const want = {};
      if (!['center', 'centerContinuous'].includes(a.horizontal)) want.horizontal = 'center';
      if (a.vertical !== 'center') want.vertical = 'center';
      if (Object.keys(want).length) change(key).align = want;
    }
  }
  // 3) 글자 크기
  if (sizes && align && majority) {
    for (const key of sizeKeys) {
      const s = cells.get(key);
      const sz = styles.fontOf(s).sz;
      if (sz >= 18 || sz === majority) continue;
      if (sz < majority) {
        if (styles.isWrapped(s)) continue;
        const text = String(textOf(key)).trim();
        const width = [...text].reduce((sum, ch) => sum + charWidth(ch, majority), 0);
        if (width > cellWidthPt(key) - 4) continue; // 키우면 칸을 넘친다
      }
      change(key).font = { ...(change(key).font || {}), sz: majority };
    }
  }
  // 4) 숫자 서식
  if (formats) {
    // (숫자 서식은 화면에서도 보이므로 인쇄 범위 밖 칸까지 맞춘다)
    for (const [key, s] of cells) {
      const id = styles.numFmtIdOf(s);
      const kind = styles.formatKind(id);
      // 글자(수식 아님) 칸에 붙은 숫자·날짜 서식은 일반으로
      if (kinds.get(key) === 'text' && !['general', 'text'].includes(kind)) change(key).numFmt = 0;
      // 글자를 가져오는 수식 칸에 붙은 날짜 서식도 일반으로 (숫자 서식은 다른 줄에서 숫자가 나올 수 있어 둔다)
      else if (kinds.get(key) === 'formula' && vtypes.get(key) === 'text' && kind === 'date') change(key).numFmt = 0;
    }
    // 숫자 칸의 '일반' 서식
    const byCol = new Map();
    for (const [key] of cells) {
      if (vtypes.get(key) !== 'number') continue;
      const [row, col] = key.split(',').map(Number);
      if (!byCol.has(col)) byCol.set(col, []);
      byCol.get(col).push(row);
    }
    for (const [col, rowsList] of byCol) {
      rowsList.sort((a, b) => a - b);
      const segments = [];
      for (const row of rowsList) {
        const last = segments[segments.length - 1];
        if (last && row - last[last.length - 1] === 1) last.push(row); else segments.push([row]);
      }
      for (const seg of segments) {
        const keys = seg.map((row) => K(row, col));
        const generals = keys.filter((key) => styles.numFmtIdOf(cells.get(key)) === 0);
        if (!generals.length) continue;
        const explicit = new Map();
        keys.forEach((key) => {
          const id = styles.numFmtIdOf(cells.get(key));
          if (id !== 0 && ['number', 'date'].includes(styles.formatKind(id))) explicit.set(id, (explicit.get(id) || 0) + 1);
        });
        let target = null;
        if (explicit.size) target = { id: [...explicit].sort((a, b) => b[1] - a[1])[0][0] };
        else if (keys.some((key) => DATE_FUNCS.test(formulas.get(key) || ''))) target = { code: 'yyyy-mm-dd' };
        else {
          const decimals = Math.max(...keys.map((key) => decimalsOf(raws.get(key), formulas.has(key))));
          target = { code: decimals === 0 ? '0' : `0.${'0'.repeat(decimals)}` };
        }
        for (const key of generals) change(key).numFmt = target.id !== undefined ? target.id : styles.fmtId(target.code);
      }
    }
  }
  if (changes.size === 0) return xml;
  return xml.replace(/<c\b([^>]*?)(\/?)>/g, (whole, attrs, close) => {
    const ref = parseRef(attr(attrs, 'r') || '');
    if (!ref) return whole;
    const c = changes.get(K(ref.row, ref.col));
    if (!c) return whole;
    let next = Number(attr(attrs, 's') || 0);
    if (c.font) next = styles.withFontChange(next, c.font);
    if (c.align) next = styles.withAlignment(next, c.align);
    if (c.numFmt !== undefined) next = styles.withNumFmtId(next, c.numFmt);
    return `<c${attrs.replace(/\s*\bs="\d+"/, '')} s="${next}"${close}>`;
  });
};

// 간격 맞추기: 소제목(▶·◆) 앞, 표와 표 사이, 마지막 표와 양식번호 칸 사이의 빈 간격을 같게 하고,
// 소제목 높이를 통일하며, 소제목과 그 아래 표 사이 빈 줄은 좁힌다. 줄을 지우지 않고 높이만 바꾼다
// (계산식이 행 번호를 참조하므로). 그림이 걸친 줄은 건드리지 않는다
const ensureRowHeight = (xml, row, ht) => {
  const re = new RegExp(`<row\\b([^>]*?)\\br="${row}"([^>]*?)(/?)>`);
  if (re.test(xml)) {
    return xml.replace(re, (w, a, b, close) => {
      const attrs = `${a}r="${row}"${b}`.replace(/\s*ht="[^"]*"/, '').replace(/\s*customHeight="[^"]*"/, '');
      return `<row${attrs} ht="${ht}" customHeight="1"${close}>`;
    });
  }
  const tag = `<row r="${row}" ht="${ht}" customHeight="1"/>`;
  let placed = false;
  const out = xml.replace(/<row\b[^>]*\br="(\d+)"/g, (w, n) => {
    if (!placed && Number(n) > row) { placed = true; return `${tag}${w}`; }
    return w;
  });
  if (placed) return out;
  if (/<sheetData\s*\/>/.test(out)) return out.replace(/<sheetData\s*\/>/, () => `<sheetData>${tag}</sheetData>`);
  return out.replace('</sheetData>', () => `${tag}</sheetData>`);
};

export const normalizeSectionSpacing = (xml, styles, { sharedStrings = [], drawingRows = new Set(), gap = 18, titleHeight = 21, afterTitle = 4, cols = null } = {}) => {
  const inCols = (col) => !cols || (col >= cols.from && col <= cols.to);
  const rows = new Map(); // row → { value, side, title, styles: [[ref, s]] }
  const valueRowsAnywhere = new Set(); // 인쇄 범위 밖까지 포함해 값이 있는 줄 (이런 줄은 줄이지 않는다)
  const fmtRow = /<sheetFormatPr\b([^>]*)\/?>/.exec(xml);
  const defaultRowHeight = Number((fmtRow && attr(fmtRow[1], 'defaultRowHeight')) || 15);
  const rowHeight = new Map();
  for (const m of xml.matchAll(/<row\b([^>]*?)\/?>/g)) {
    const ht = attr(m[1], 'ht');
    if (ht !== null) rowHeight.set(Number(attr(m[1], 'r')), Number(ht));
  }
  const info = (row) => {
    if (!rows.has(row)) rows.set(row, { value: false, side: false, title: false, cells: [] });
    return rows.get(row);
  };
  for (const m of xml.matchAll(/<c\b([^>]*?)(?:\/>|>([\s\S]*?)<\/c>)/g)) {
    const ref = parseRef(attr(m[1], 'r') || '');
    if (ref && cellKind(m[1], m[2] || '') !== 'empty') valueRowsAnywhere.add(ref.row);
    if (!ref || !inCols(ref.col)) continue;
    const row = info(ref.row);
    const s = Number(attr(m[1], 's') || 0);
    row.cells.push(s);
    const inner = m[2] || '';
    const kind = cellKind(m[1], inner);
    if (kind !== 'empty') row.value = true;
    let text = '';
    if (attr(m[1], 't') === 's') { const v = /<v>(\d+)<\/v>/.exec(inner); text = v ? sharedStrings[Number(v[1])] || '' : ''; }
    else if (attr(m[1], 't') === 'inlineStr') text = [...inner.matchAll(/<t\b[^>]*>([^<]*)<\/t>/g)].map((x) => decodeXml(x[1])).join('');
    else if (kind === 'formula') { const f = /<f\b[^>]*>([^<]*)<\/f>/.exec(inner); text = f && /"▶/.test(decodeXml(f[1])) ? '▶' : ''; }
    if (ref.col <= 3 && /^\s*[▶◆]/.test(text)) row.title = true;
    const x = styles.sidesOf(s);
    if (x.left || x.right) row.side = true;
  }
  // 병합 칸은 영역 전체가 내용 줄
  for (const m of xml.matchAll(/<mergeCell\b[^>]*ref="([^"]+)"/g)) {
    const r = parseRange(m[1]);
    if (!r || r.from.row === r.to.row || !inCols(r.from.col)) continue;
    const head = rows.get(r.from.row);
    if (!head || !(head.value || head.side)) continue;
    for (let row = r.from.row; row <= r.to.row; row++) { const x = info(row); x.value = x.value || head.value; x.side = x.side || head.side; }
  }
  const content = (row) => drawingRows.has(row) || (rows.has(row) && (rows.get(row).value || rows.get(row).side));
  const isTable = (row) => rows.has(row) && rows.get(row).side;
  const isTitle = (row) => rows.has(row) && rows.get(row).title;
  if (drawingRows.unknown) return xml; // 위치를 줄로 알 수 없는 그림이 있으면 줄 높이를 바꾸지 않는다
  const contentRows = [...new Set([...rows.keys(), ...drawingRows])].filter(content).sort((a, b) => a - b);
  if (contentRows.length < 2) return xml;
  const MAX_GAP_ROWS = 3; // 이보다 긴 빈 구간은 일부러 비워 둔 자리로 보고 그대로 둔다
  const heights = new Map(); // row → ht
  const bottomRows = new Set();
  for (let i = 1; i < contentRows.length; i++) {
    const prev = contentRows[i - 1]; const next = contentRows[i];
    const blanks = [];
    for (let row = prev + 1; row < next; row++) if (!drawingRows.has(row)) blanks.push(row);
    if (blanks.length !== next - prev - 1) continue; // 그림이 걸친 구간은 그대로
    if (blanks.length > MAX_GAP_ROWS) continue;
    // 인쇄 범위 밖에 값이 있는 줄(예: 노즐산정 오른쪽 설명)이 끼면 줄이지는 않고, 모자랄 때만 키운다
    const guarded = blanks.some((row) => valueRowsAnywhere.has(row));
    const currentHeight = (row) => (rowHeight.has(row) ? rowHeight.get(row) : defaultRowHeight);
    const setGap = (target) => {
      if (guarded) {
        const total = blanks.reduce((sum, row) => sum + currentHeight(row), 0);
        if (total >= target - 0.01) return; // 이미 같거나 크면 그대로 (줄이면 바깥 내용이 눌린다)
        const extra = (target - total) / blanks.length;
        blanks.forEach((row) => heights.set(row, Math.round((currentHeight(row) + extra) * 100) / 100));
        return;
      }
      blanks.forEach((row) => heights.set(row, Math.round((target / blanks.length) * 100) / 100));
    };
    if (isTitle(next) && !isTitle(prev)) {
      if (blanks.length) setGap(gap);
      else { heights.set(next, titleHeight + gap); bottomRows.add(next); }
    } else if (isTitle(prev) && !isTitle(next)) {
      if (!guarded) blanks.forEach((row) => heights.set(row, afterTitle));
    } else if (isTable(prev) && isTable(next) && blanks.length) {
      setGap(gap);
    }
  }
  for (const row of contentRows) if (isTitle(row) && !heights.has(row)) heights.set(row, titleHeight);
  let out = xml;
  for (const [row, ht] of [...heights].sort((a, b) => a[0] - b[0])) out = ensureRowHeight(out, row, ht);
  if (bottomRows.size) {
    out = out.replace(/<c\b([^>]*?)(\/?)>/g, (whole, attrs, close) => {
      const ref = parseRef(attr(attrs, 'r') || '');
      if (!ref || !bottomRows.has(ref.row)) return whole;
      const next = styles.withVertical(Number(attr(attrs, 's') || 0), 'bottom');
      return `<c${attrs.replace(/\s*\bs="\d+"/, '')} s="${next}"${close}>`;
    });
  }
  return out;
};

// 시트에 붙은 그림·도형이 걸친 줄 (1부터)
// 결과: 줄 번호 Set. 위치를 줄로 알 수 없는 그림(절대 위치)이 있으면 rowsSet.unknown = true
const anchorRows = (text, rowsSet) => {
  // <xdr:from>…<xdr:row>n</xdr:row> 또는 OLE 개체의 <from>…<xdr:row>n</xdr:row> (0부터)
  const from = /<(?:xdr:)?from>[\s\S]*?<xdr:row>(\d+)<\/xdr:row>/.exec(text);
  const to = /<(?:xdr:)?to>[\s\S]*?<xdr:row>(\d+)<\/xdr:row>/.exec(text);
  if (!from) return;
  const start = Number(from[1]) + 1; const end = to ? Number(to[1]) + 1 : start;
  for (let row = start; row <= end; row++) rowsSet.add(row);
};
// 그림·개체가 걸친 가장 아래 행·오른쪽 열 (1부터) — 인쇄 범위를 잴 때 그림이 잘리지 않게
const readDrawingExtent = async (zip, sheetPath, sheetXml = '') => {
  let maxRow = 0; let maxCol = 0;
  const take = (text) => {
    for (const tag of ['from', 'to']) {
      const m = new RegExp(`<(?:xdr:)?${tag}>[\\s\\S]*?<xdr:col>(\\d+)</xdr:col>[\\s\\S]*?<xdr:row>(\\d+)</xdr:row>`).exec(text);
      if (m) { maxCol = Math.max(maxCol, Number(m[1]) + 1); maxRow = Math.max(maxRow, Number(m[2]) + 1); }
    }
  };
  for (const a of sheetXml.matchAll(/<anchor\b[\s\S]*?<\/anchor>/g)) take(a[0]);
  const dir = sheetPath.replace(/[^/]+$/, '');
  const relsFile = zip.file(`${dir}_rels/${sheetPath.split('/').pop()}.rels`);
  if (relsFile) {
    const rels = await relsFile.async('string');
    for (const m of rels.matchAll(/<Relationship\b([^>]*)\/?>/g)) {
      if (!/\/drawing$/.test(attr(m[1], 'Type') || '')) continue;
      const stack = [];
      for (const part of (dir + (attr(m[1], 'Target') || '')).split('/')) { if (part === '..') stack.pop(); else if (part && part !== '.') stack.push(part); }
      const file = zip.file(stack.join('/'));
      if (!file) continue;
      const drawing = await file.async('string');
      for (const a of drawing.matchAll(/<xdr:(twoCellAnchor|oneCellAnchor)\b[\s\S]*?<\/xdr:\1>/g)) take(a[0]);
    }
  }
  return { maxRow, maxCol };
};

const readDrawingRows = async (zip, sheetPath, sheetXml = '') => {
  const rowsSet = new Set();
  // 시트 XML 안의 OLE 개체·컨트롤 위치 (예: 시료측정위치의 원형굴뚝 표 개체)
  for (const a of sheetXml.matchAll(/<anchor\b[\s\S]*?<\/anchor>/g)) anchorRows(a[0], rowsSet);
  const dir = sheetPath.replace(/[^/]+$/, '');
  const relsFile = zip.file(`${dir}_rels/${sheetPath.split('/').pop()}.rels`);
  if (!relsFile) return rowsSet;
  const rels = await relsFile.async('string');
  for (const m of rels.matchAll(/<Relationship\b([^>]*)\/?>/g)) {
    // attr()는 따옴표를 뺀 값을 돌려주므로 값이 '/drawing'으로 끝나는지 본다 (vmlDrawing은 메모·옛 도형이라 제외)
    if (!/\/drawing$/.test(attr(m[1], 'Type') || '')) continue;
    const target = attr(m[1], 'Target') || '';
    const parts = (dir + target).split('/');
    const stack = [];
    for (const part of parts) { if (part === '..') stack.pop(); else if (part && part !== '.') stack.push(part); }
    const file = zip.file(stack.join('/'));
    if (!file) continue;
    const drawing = await file.async('string');
    for (const a of drawing.matchAll(/<xdr:(twoCellAnchor|oneCellAnchor)\b[\s\S]*?<\/xdr:\1>/g)) anchorRows(a[0], rowsSet);
    if (/<xdr:absoluteAnchor\b/.test(drawing)) rowsSet.unknown = true;
  }
  return rowsSet;
};

// ── 글씨 넘침 막기: 양옆에 세로선이 있는 칸(표 칸)에 값이 있으면 '셀에 맞춤'을 켠다 ──
// 줄 바꿈이 켜진 칸은 그대로 둔다. 선 없는 글(제목·설명)은 옆 빈칸으로 넘쳐 보여도 되므로 건드리지 않는다
export const shrinkTableCells = (xml, styles, sharedStrings = []) => {
  const { kinds, merges, mergeOf, K, vLine } = analyzeTables(xml, styles, sharedStrings);
  return xml.replace(/<c\b([^>]*?)(\/?)>/g, (whole, attrs, close) => {
    const ref = parseRef(attr(attrs, 'r') || '');
    if (!ref) return whole;
    const key = K(ref.row, ref.col);
    const kind = kinds.get(key);
    if (!kind || kind === 'empty') return whole;
    const b = mergeOf.has(key) ? merges[mergeOf.get(key)] : { from: ref, to: ref };
    if (!(vLine(b.from.row, b.from.col) && vLine(b.from.row, b.to.col + 1))) return whole;
    const s0 = Number(attr(attrs, 's') || 0);
    const next = styles.withShrink(s0);
    return next === s0 ? whole : `<c${attrs.replace(/\s*\bs="\d+"/, '')} s="${next}"${close}>`;
  });
};

// ── 시트 꾸밈 통일: 모든 시트의 제목·양식번호 줄을 인쇄했을 때 같은 크기로, 쪽 전체를 두른 바깥 테두리는 없앤다 ──
//  - 제목(1행부터 인쇄 폭 전체를 합친 칸): 상자 선을 지우고, 인쇄 크기 기준 같은 글자 크기·같은 높이로 (시트마다 인쇄 배율이 달라 배율로 나눈다)
//  - 양식번호 줄(인쇄 범위 마지막 행의 ■…): 선을 지우고 같은 글자 크기·높이, 오른쪽 정렬
//  - 바깥 테두리: 인쇄 범위 왼쪽·오른쪽 끝 칸의 바깥 선을 지운다. 그 칸이 표 칸(안쪽 세로선과 위·아래 선이 있는 칸)이면 표의 선이므로 둔다
//  - sectionRule: 가로 구분선으로 묶음을 나누는 시트(계산식)는 제목 아래·양식번호 줄 위에도 같은 구분선을 둔다
export const normalizePageFurniture = (xml, styles, {
  sharedStrings = [], range = null, scale = 1, titlePt = 18, titleHeight = 32, footerPt = 9, footerHeight = 14,
  footerPattern = /^\s*■/, sectionRule = false, frame = true, sectionBand = '',
} = {}) => {
  if (!range || !range.from.row || !range.to.row || !(scale > 0)) return xml;
  const { cells, texts, merges, mergeOf, K, sameMerge, own } = analyzeTables(xml, styles, sharedStrings);
  const { from, to } = range;
  const sides = new Map(); // "r,c" → { left, right, top, bottom } (바꿔 가는 값)
  const side = (key) => {
    if (!sides.has(key)) sides.set(key, { ...own(key) });
    return sides.get(key);
  };
  const hLine = (row, col) => !sameMerge(K(row - 1, col), K(row, col)) && Boolean(side(K(row - 1, col)).bottom || side(K(row, col)).top);
  const vLine = (row, col) => !sameMerge(K(row, col - 1), K(row, col)) && Boolean(side(K(row, col - 1)).right || side(K(row, col)).left);
  const blockOf = (row, col) => (mergeOf.has(K(row, col)) ? merges[mergeOf.get(K(row, col))] : { from: { row, col }, to: { row, col } });
  const fullWidth = (b) => b.from.col <= from.col && b.to.col >= to.col;
  // 칸 묶음이 표 칸인지: 맞은편 세로선이 있고 위·아래 선이 있다 (폭 전체를 합친 칸은 위·아래 선이 모두 있어야 한다)
  const boxed = (b, edge) => {
    const opposite = edge === 'left' ? vLine(b.from.row, b.to.col + 1) : vLine(b.from.row, b.from.col);
    const top = hLine(b.from.row, edge === 'left' ? b.from.col : b.to.col);
    const bottom = hLine(b.to.row + 1, edge === 'left' ? b.from.col : b.to.col);
    return opposite && (fullWidth(b) ? top && bottom : top || bottom);
  };
  const gridCell = (row, col) => vLine(row, col) && vLine(row, col + 1); // 양옆에 세로선이 있는 칸
  const noFrame = frame === true || frame === 'none' || frame === 'footer';
  // 인쇄 범위 아래에 이어지는 내용(같은 열 범위)의 마지막 행
  let lastRow = to.row;
  for (const key of cells.keys()) { const [row, col] = key.split(',').map(Number); if (col >= from.col && col <= to.col && row > lastRow) lastRow = row; }
  const fonts = new Map(); // "r,c" → 글자 크기
  const aligns = new Map(); // "r,c" → { horizontal, vertical }
  const heights = new Map(); // 행 → 높이
  const half = (v) => Math.round(v * 2) / 2;

  // 제목: 인쇄 범위 첫 칸에서 시작해 폭 전체를 합친, 글자가 있는 칸
  const titleBlock = (() => {
    const b = blockOf(from.row, from.col);
    if (b.from.row !== from.row || !fullWidth(b) || b.to.row - b.from.row > 5) return null;
    return String(texts.get(K(b.from.row, b.from.col)) || '').trim() ? b : null;
  })();
  // 양식번호 줄: 인쇄 범위 마지막 행
  const footerCol = (() => {
    for (let col = from.col; col <= to.col; col++) if (footerPattern.test(String(texts.get(K(to.row, col)) || ''))) return col;
    return 0;
  })();
  const furnitureRows = new Set();
  if (titleBlock) for (let row = titleBlock.from.row; row <= titleBlock.to.row; row++) furnitureRows.add(row);
  if (footerCol) furnitureRows.add(to.row);

  if (titleBlock) {
    const below = titleBlock.to.row + 1;
    // 제목 바로 아래 줄이 표인지: 인쇄 폭 안쪽에 세로선이 있다 (바깥 테두리의 양 끝 선은 세지 않는다)
    let tableRowBelow = false;
    for (let col = from.col + 1; col <= to.col; col++) if (vLine(below, col)) { tableRowBelow = true; break; }
    const hadBottom = []; // 제목 상자 아랫선이 있던 열
    for (let col = from.col; col <= to.col; col++) hadBottom[col] = hLine(below, col);
    if (noFrame) {
      for (let row = titleBlock.from.row; row <= titleBlock.to.row; row++) for (let col = from.col; col <= to.col; col++) Object.assign(side(K(row, col)), { left: null, right: null, top: null, bottom: null });
      for (let col = from.col; col <= to.col; col++) {
        const key = K(below, col);
        if (!cells.has(key)) continue;
        // 제목 바로 아래가 표 칸이면 표의 윗선으로 남기고, 아니면 제목 상자의 아랫선이므로 지운다
        if (hadBottom[col] && (sectionRule || tableRowBelow)) side(key).top = 'medium';
        else side(key).top = null;
      }
    }
    const n = titleBlock.to.row - titleBlock.from.row + 1;
    for (let row = titleBlock.from.row; row <= titleBlock.to.row; row++) heights.set(row, Math.round((titleHeight / scale / n) * 100) / 100);
    fonts.set(K(titleBlock.from.row, titleBlock.from.col), half(titlePt / scale));
    aligns.set(K(titleBlock.from.row, titleBlock.from.col), { horizontal: 'center', vertical: 'center' });
    const titleStyle = cells.get(K(titleBlock.from.row, titleBlock.from.col));
    for (const m of merges) {
      if (m === titleBlock || m.from.row <= titleBlock.to.row || m.to.row >= to.row || !fullWidth(m)) continue;
      const key = K(m.from.row, m.from.col);
      // 제목과 같은 서식이거나, 굵고 큰 글씨(16pt 이상)인 폭 전체 병합 칸
      const big = cells.has(key) && styles.isBold(cells.get(key)) && styles.fontOf(cells.get(key)).sz >= 16;
      if ((cells.get(key) !== titleStyle && !big) || !String(texts.get(key) || '').trim()) continue;
      const rows = m.to.row - m.from.row + 1;
      for (let row = m.from.row; row <= m.to.row; row++) heights.set(row, Math.round((titleHeight / scale / rows) * 100) / 100);
      fonts.set(key, half(titlePt / scale));
      aligns.set(key, { horizontal: 'center', vertical: 'center' });
    }
  }
  if (footerCol) {
    const above = to.row - 1;
    if (noFrame) {
      const hadTop = [];
      for (let col = from.col; col <= to.col; col++) hadTop[col] = hLine(to.row, col);
      for (let col = from.col; col <= to.col; col++) Object.assign(side(K(to.row, col)), { left: null, right: null, top: null, bottom: null });
      for (let col = from.col; col <= to.col; col++) {
        const key = K(above, col);
        if (!cells.has(key) || furnitureRows.has(above)) continue;
        // 양식번호 줄 바로 위가 표 칸이면 표의 아랫선을 남긴다 (양식번호 줄의 윗선으로만 그어져 있던 경우 포함)
        if (sectionRule ? true : (gridCell(above, col) && hadTop[col])) side(key).bottom = side(key).bottom || 'medium';
      }
    }
    heights.set(to.row, Math.round((footerHeight / scale) * 100) / 100);
    const b = blockOf(to.row, footerCol);
    fonts.set(K(b.from.row, b.from.col), half(footerPt / scale));
    aligns.set(K(b.from.row, b.from.col), { horizontal: 'right', vertical: 'center' });
  }
  // 바깥 테두리: 왼쪽·오른쪽 끝
  if (noFrame) {
    for (const edge of ['left', 'right']) {
      const col = edge === 'left' ? from.col : to.col;
      let row = from.row;
      while (row <= lastRow) {
        const b = blockOf(row, col);
        const last = Math.max(row, b.to.row);
        if (!furnitureRows.has(row)) {
          const has = edge === 'left' ? vLine(row, col) : vLine(row, col + 1);
          if (has && !boxed(b, edge)) {
            for (let r = Math.max(row, b.from.row); r <= last; r++) {
              side(K(r, col))[edge] = null;
              // 인쇄 범위 바로 밖 칸이 같은 선을 가지고 있으면 그것도 지운다
              const outside = K(r, edge === 'left' ? col - 1 : col + 1);
              if (cells.has(outside)) side(outside)[edge === 'left' ? 'right' : 'left'] = null;
            }
          }
        }
        row = last + 1;
      }
    }
  }

  // 전체 테두리(frame: 'box'): 인쇄 범위 둘레, 제목 아래, 양식번호 줄 위를 모두 굵은 선으로 — 모든 시트 같은 모양
  let boxed0 = xml;
  if (frame === 'box' || (frame === 'footer' && footerCol)) {
    const need = new Map(); // 행 → 필요한 열들
    const want = (row, col, name, style = 'medium') => { side(K(row, col))[name] = style; if (!need.has(row)) need.set(row, new Set()); need.get(row).add(col); };
    if (frame === 'footer') {
      // 양식번호 줄: 위쪽에 가는 선 한 줄만
      for (let col = from.col; col <= to.col; col++) want(to.row, col, 'top', 'thin');
    } else {
      for (let row = from.row; row <= to.row; row++) { want(row, from.col, 'left'); want(row, to.col, 'right'); }
      for (let col = from.col; col <= to.col; col++) {
        want(from.row, col, 'top'); want(to.row, col, 'bottom');
        if (titleBlock) want(titleBlock.to.row, col, 'bottom');
        if (footerCol) want(to.row, col, 'top');
      }
    }
    // 테두리를 입힐 칸이 파일에 없으면 빈 칸을 만들어 둔다
    for (const [row, colSet] of need) {
      const missing = [...colSet].filter((col) => !cells.has(K(row, col))).sort((x, y) => x - y);
      if (missing.length === 0) continue;
      const make = (col) => `<c r="${numToColName(col)}${row}"/>`;
      const rowRe = new RegExp(`<row\\b(?=[^>]*\\br="${row}")([^>]*?)(?:/>|>([\\s\\S]*?)</row>)`);
      let m = rowRe.exec(boxed0);
      if (!m) {
        const tag = `<row r="${row}"/>`;
        let placed = false;
        boxed0 = boxed0.replace(/<row\b[^>]*\br="(\d+)"/g, (w, n) => { if (!placed && Number(n) > row) { placed = true; return `${tag}${w}`; } return w; });
        if (!placed) boxed0 = boxed0.replace('</sheetData>', () => `${tag}</sheetData>`);
        m = rowRe.exec(boxed0);
        if (!m) continue;
      }
      const existing = [...(m[2] || '').matchAll(/<c\b[^>]*?\br="([A-Z]+)\d+"[^>]*?(?:\/>|>[\s\S]*?<\/c>)/g)].map((c) => ({ col: colToNum(c[1]), xml: c[0] }));
      const all = [...existing, ...missing.map((col) => ({ col, xml: make(col) }))].sort((x, y) => x.col - y.col);
      boxed0 = boxed0.replace(rowRe, () => `<row${m[1].replace(/\s*\/$/, '')}>${all.map((c) => c.xml).join('')}</row>`);
    }
  }
  // 묶음 제목 띠: 첫 열에 굵은 글씨 제목만 있는 줄을 인쇄 폭 전체에 옅은 색으로 칠한다 (테두리 없이도 묶음이 구분되게)
  const bandRows = new Set();
  if (sectionBand) {
    for (let row = from.row; row <= lastRow; row++) {
      if (furnitureRows.has(row)) continue;
      const key = K(row, from.col);
      const text = String(texts.get(key) || '').trim();
      if (!text || /=$/.test(text) || !cells.has(key) || !styles.isBold(cells.get(key))) continue;
      let alone = true;
      for (let col = from.col + 1; col <= to.col; col++) if (String(texts.get(K(row, col)) || '').trim()) { alone = false; break; }
      if (alone) bandRows.add(row);
    }
    if (sectionRule) {
      for (const row of bandRows) for (let col = from.col; col <= to.col; col++) if (cells.has(K(row, col))) side(K(row, col)).top = 'medium';
    }
  }
  let out = boxed0.replace(/<c\b([^>]*?)(\/?)>/g, (whole, attrs, close) => {
    const ref = parseRef(attr(attrs, 'r') || '');
    if (!ref) return whole;
    const key = K(ref.row, ref.col);
    const original = Number(attr(attrs, 's') || 0);
    let next = original;
    if (bandRows.has(ref.row) && ref.col >= from.col && ref.col <= to.col) next = styles.withFill(next, sectionBand);
    if (sides.has(key)) {
      const before = own(key); const after = sides.get(key);
      const cleared = ['left', 'right', 'top', 'bottom'].filter((k) => before[k] && !after[k]);
      const set = Object.fromEntries(['left', 'right', 'top', 'bottom'].filter((k) => after[k] && after[k] !== before[k]).map((k) => [k, after[k]]));
      if (cleared.length) next = styles.withoutBorderSides(next, cleared);
      if (Object.keys(set).length) next = styles.withBorderStyles(next, set);
    }
    if (fonts.has(key)) next = styles.withFontChange(next, { sz: fonts.get(key) });
    if (aligns.has(key)) next = styles.withAlignment(next, aligns.get(key));
    return next === original ? whole : `<c${attrs.replace(/\s*\bs="\d+"/, '')} s="${next}"${close}>`;
  });
  for (const [row, ht] of heights) out = ensureRowHeight(out, row, ht);
  return out;
};

// ── 인쇄 크기 통일: 한 파일의 모든 시트가 같은 폭으로 인쇄되게, 시트마다 배율을 정한다 ──
// 공통 폭 = 용지 인쇄 가능 폭과, 시트마다 '세로가 한 장에 들어가는 폭' 중 가장 작은 값
const A4 = { width: 595.3, height: 841.9 };
// ECMA-376 18.3.1.13: 픽셀 = Truncate(((256*width + Truncate(128/MDW)) / 256) * MDW)
const colWidthPt = (chars, mdw) => Math.trunc(((256 * chars + Math.trunc(128 / mdw)) / 256) * mdw) * 0.75;
const isTrue = (v) => v === '1' || v === 'true';
const measurePrintArea = (xml, range, mdw, styles, extent = null) => {
  const fmt = /<sheetFormatPr\b([^>]*)\/?>/.exec(xml);
  const defaultRow = Number((fmt && attr(fmt[1], 'defaultRowHeight')) || 15);
  const baseCol = Number((fmt && attr(fmt[1], 'baseColWidth')) || 8);
  const defaultCol = Number((fmt && attr(fmt[1], 'defaultColWidth')) || baseCol + 0.71);
  const colPt = (chars) => colWidthPt(chars, mdw);
  const cols = [];
  for (const m of xml.matchAll(/<col\b([^>]*)\/?>/g)) {
    cols.push({ min: Number(attr(m[1], 'min')), max: Number(attr(m[1], 'max')), width: Number(attr(m[1], 'width') || defaultCol), hidden: isTrue(attr(m[1], 'hidden')) });
  }
  let r = range;
  if (r && (!r.from.row || !r.to.row)) {
    const rowsWithContent = [...xml.matchAll(/<c\b[^>]*\br="[A-Z]+(\d+)"[^>]*>(?:<f|<v|<is)/g)].map((m) => Number(m[1]));
    r = { from: { col: r.from.col, row: 1 }, to: { col: r.to.col, row: Math.max(1, ...rowsWithContent) } };
  }
  if (!r) {
    // 인쇄 범위가 없으면 값이나 테두리가 있는 칸까지 (기록 범위 dimension은 빈 서식 칸까지 넓게 잡혀 있을 수 있다)
    let maxRow = 0; let maxCol = 0;
    for (const m of xml.matchAll(/<c\b([^>]*?)(?:\/>|>([\s\S]*?)<\/c>)/g)) {
      const ref = parseRef(attr(m[1], 'r') || '');
      if (!ref) continue;
      const x = styles.sidesOf(Number(attr(m[1], 's') || 0));
      if (cellKind(m[1], m[2]) === 'empty' && !(x.left || x.right || x.top || x.bottom)) continue;
      maxRow = Math.max(maxRow, ref.row); maxCol = Math.max(maxCol, ref.col);
    }
    for (const m of xml.matchAll(/<mergeCell\b[^>]*ref="([^"]+)"/g)) {
      const mr = parseRange(m[1]);
      if (mr && mr.from.row <= maxRow) { maxRow = Math.max(maxRow, mr.to.row); maxCol = Math.max(maxCol, mr.to.col); }
    }
    if (extent) { maxRow = Math.max(maxRow, extent.maxRow); maxCol = Math.max(maxCol, extent.maxCol); }
    r = maxRow ? { from: { row: 1, col: 1 }, to: { row: maxRow, col: maxCol } } : null;
  }
  if (!r) return null;
  let width = 0;
  for (let col = r.from.col; col <= r.to.col; col++) {
    const def = cols.find((c) => col >= c.min && col <= c.max);
    if (def && def.hidden) continue;
    width += colPt(def ? def.width : defaultCol);
  }
  const rowHeights = new Map();
  for (const m of xml.matchAll(/<row\b([^>]*?)\/?>/g)) {
    const row = Number(attr(m[1], 'r'));
    rowHeights.set(row, isTrue(attr(m[1], 'hidden')) ? 0 : Number(attr(m[1], 'ht') || defaultRow));
  }
  let height = 0;
  for (let row = r.from.row; row <= r.to.row; row++) height += rowHeights.has(row) ? rowHeights.get(row) : defaultRow;
  return { width, height, range: r, fromContent: !range };
};
// pageMargins·pageSetup이 없으면 규격 순서에 맞는 자리에 넣는다
// 시트 수준의 인쇄 설정은 customSheetViews 뒤에 있다. 그 앞부분(사용자 지정 보기의 인쇄 설정)은 건드리지 않게 나눠서 다룬다
const splitAtCustomViews = (xml) => {
  const end = xml.indexOf('</customSheetViews>');
  if (end < 0) return ['', xml];
  const cut = end + '</customSheetViews>'.length;
  return [xml.slice(0, cut), xml.slice(cut)];
};
// 시트 맨 끝의 확장 정보(extLst)만 찾는다 (조건부 서식 안의 extLst는 아님)
const topLevelInsertAt = (xml) => {
  const named = xml.search(/<(?:headerFooter|rowBreaks|colBreaks|customProperties|cellWatches|ignoredErrors|smartTags|drawing|legacyDrawing|legacyDrawingHF|drawingHF|picture|oleObjects|controls|webPublishItems|tableParts)\b/);
  if (named >= 0) return named;
  const tail = /<extLst\b(?:(?!<extLst\b)[\s\S])*<\/extLst>\s*<\/worksheet>\s*$/.exec(xml);
  if (tail) return tail.index;
  return xml.search(/<\/worksheet>/);
};
const ensurePageElements = (xml) => {
  const [head, body] = splitAtCustomViews(xml);
  let out = body;
  if (!/<pageMargins\b/.test(out)) {
    const at = out.search(/<pageSetup\b/) >= 0 ? out.search(/<pageSetup\b/) : topLevelInsertAt(out);
    if (at >= 0) out = `${out.slice(0, at)}<pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/>${out.slice(at)}`;
  }
  if (!/<pageSetup\b/.test(out)) {
    out = out.replace(/(<pageMargins\b[^>]*\/>)/, '$1<pageSetup paperSize="9" orientation="portrait"/>');
  }
  // <pageSetup …></pageSetup> → 스스로 닫는 태그로 (안의 확장 정보는 버린다)
  out = out.replace(/<pageSetup\b([^>]*?)>[\s\S]*?<\/pageSetup>/, '<pageSetup$1/>');
  return head + out;
};

const setPrintScale = (xml, scale, margins) => {
  const [head0, body0] = splitAtCustomViews(ensurePageElements(xml));
  // 시트 속성(sheetPr)은 맨 앞에 있으므로 앞부분에서, 나머지는 뒷부분에서
  const sheetPrFixed = head0 ? head0.replace(/<pageSetUpPr\b([^>]*)\/>/, (w, a) => `<pageSetUpPr${a.replace(/\s*fitToPage="[^"]*"/, '')}/>`) : '';
  let out = body0;
  if (!head0) out = out.replace(/<pageSetUpPr\b([^>]*)\/>/, (w, a) => `<pageSetUpPr${a.replace(/\s*fitToPage="[^"]*"/, '')}/>`);
  // 양식의 수동 페이지 나누기는 지운다 (한 시트 = 한 장)
  out = out.replace(/<rowBreaks\b[^>]*(?:\/>|>[\s\S]*?<\/rowBreaks>)/, '').replace(/<colBreaks\b[^>]*(?:\/>|>[\s\S]*?<\/colBreaks>)/, '');
  if (/<pageSetup\b[^>]*\/>/.test(out)) {
    out = out.replace(/<pageSetup\b([^>]*)\/>/, (w, a) => {
      const orientation = attr(a, 'orientation') || 'portrait';
      return `<pageSetup paperSize="9" scale="${scale}" orientation="${orientation}"${a.replace(/\s*(?:paperSize|scale|orientation|fitToWidth|fitToHeight)="[^"]*"/g, '')}/>`;
    });
  }
  const m = margins;
  const marginTag = `<pageMargins left="${m.left}" right="${m.right}" top="${m.top}" bottom="${m.bottom}" header="${m.header}" footer="${m.footer}"/>`;
  out = /<pageMargins\b[^>]*\/>/.test(out) ? out.replace(/<pageMargins\b[^>]*\/>/, marginTag) : out;
  // 가로 가운데
  if (/<printOptions\b[^>]*\/>/.test(out)) out = out.replace(/<printOptions\b([^>]*)\/>/, (w, a) => `<printOptions${a.replace(/\s*horizontalCentered="[^"]*"/, '')} horizontalCentered="1"/>`);
  else out = out.replace(/<pageMargins\b/, '<printOptions horizontalCentered="1"/><pageMargins');
  return sheetPrFixed + out;
};

export async function tidyWorkbook(zip, { rules = {}, onePageSheets = [], equalWidth = false, furniture = null, palette = null } = {}) {
  const read = async (path) => { const f = zip.file(path); return f ? f.async('string') : null; };
  const workbookXml = await read('xl/workbook.xml');
  const relsXml = await read('xl/_rels/workbook.xml.rels');
  const stylesXml = await read('xl/styles.xml');
  if (!workbookXml || !relsXml || !stylesXml) return;
  const rels = {};
  for (const m of relsXml.matchAll(/<Relationship\b([^>]*)\/?>/g)) rels[attr(m[1], 'Id')] = attr(m[1], 'Target');
  const styles = new StyleBook(stylesXml);
  styles.cleanAll();
  if (palette) styles.recolorFills(palette);
  const sharedXml = await read('xl/sharedStrings.xml');
  const sharedStrings = sharedXml
    ? [...sharedXml.matchAll(/<si>([\s\S]*?)<\/si>/g)].map((m) => [...m[1].matchAll(/<t\b[^>]*>([^<]*)<\/t>/g)].map((t) => decodeXml(t[1])).join(''))
    : [];
  // 시트별 인쇄 범위의 열 (간격은 인쇄되는 부분만 본다)
  const printCols = new Map();
  const printRanges = new Map();
  const multiAreaSheets = new Set();
  for (const m of workbookXml.matchAll(/<definedName\b([^>]*)>([^<]*)<\/definedName>/g)) {
    if (attr(m[1], 'name') !== '_xlnm.Print_Area' || attr(m[1], 'localSheetId') === null) continue;
    const text = decodeXml(m[2]);
    const id = Number(attr(m[1], 'localSheetId'));
    // 여러 구역(쉼표로 이은 범위)은 장을 나눠 인쇄하므로 폭 맞추기에서 뺀다
    if (text.replace(/'(?:[^']|'')*'/g, '').includes(',')) { multiAreaSheets.add(id); continue; }
    const range = /!\$?([A-Z]{1,3})\$?(\d*):\$?([A-Z]{1,3})\$?(\d*)$/.exec(text);
    if (!range) { multiAreaSheets.add(id); continue; } // 행 전체 등 읽지 못한 범위: 그대로 둔다
    printCols.set(id, { from: colToNum(range[1]), to: colToNum(range[3]) });
    printRanges.set(id, { from: { col: colToNum(range[1]), row: range[2] ? Number(range[2]) : null }, to: { col: colToNum(range[3]), row: range[4] ? Number(range[4]) : null } });
  }
  const sheetsForWidth = []; // 인쇄 크기 통일에 쓸 시트들
  let sheetIndex = -1;
  const printAreaUpdates = new Map(); // 시트 순서 → { name, range }
  const hiddenSheets = new Set(); // 숨긴 시트 순서 (인쇄 범위를 지워 LibreOffice·한셀이 인쇄하지 않게)
  for (const m of workbookXml.matchAll(/<sheet\b([^>]*)\/>/g)) {
    sheetIndex += 1;
    const name = decodeXml(attr(m[1], 'name') || '');
    const state = attr(m[1], 'state');
    if (state === 'hidden' || state === 'veryHidden') { hiddenSheets.add(sheetIndex); continue; } // 숨긴 시트는 그대로
    const target = rels[attr(m[1], 'r:id')];
    if (!target) continue;
    const path = target.startsWith('/') ? target.slice(1) : `xl/${target}`;
    let xml = await read(path);
    if (!xml) continue;
    xml = fitToPageWidth(xml, { onePage: onePageSheets === 'all' || onePageSheets.includes(name) });
    // 시트 탭 색은 없앤다 (유량 양식의 두 시트만 검은 탭이었다)
    xml = xml.replace(/<tabColor\b[^>]*\/>/, '');
    const rule = typeof rules === 'function' ? rules(name) : rules[name];
    if (rule?.printArea) printAreaUpdates.set(sheetIndex, { name, range: rule.printArea });
    if (rule?.rowHeights) xml = setRowHeights(xml, rule.rowHeights);
    // 쓰지 않는 줄 숨기기: [첫 행, 끝 행]
    if (rule?.hideRows) for (let row = rule.hideRows[0]; row <= rule.hideRows[1]; row++) xml = ensureRowHidden(xml, row);
    // 병합 칸이 걸친 줄 전체를 같은 높이로 (양식마다 길이가 다른 표: 예) 노즐표의 채취시간 칸 A39:A56/A53/A57)
    if (rule?.rowHeightsByMerge) {
      for (const { cell, ht } of rule.rowHeightsByMerge) {
        const at = parseRef(cell);
        const merge = at && [...xml.matchAll(/<mergeCell\b[^>]*ref="([^"]+)"/g)].map((m) => parseRange(m[1])).find((r) => r && at.row >= r.from.row && at.row <= r.to.row && at.col >= r.from.col && at.col <= r.to.col);
        if (merge) xml = setRowHeights(xml, rowsAt(merge.from.row, merge.to.row, ht));
      }
    }
    if (rule?.numberFormats) xml = setNumberFormats(xml, styles, rule.numberFormats);
    if (rule?.center) xml = setCentered(xml, styles, rule.center);
    // 상자 테두리: 범위 안 모든 칸의 네 변에 가는 선
    if (rule?.boxes) {
      const targets = rule.boxes.map(parseRange).filter(Boolean);
      xml = xml.replace(/<c\b([^>]*?)(\/?)>/g, (whole, attrs, selfClose) => {
        const ref = parseRef(attr(attrs, 'r') || '');
        if (!ref || !targets.some((r) => ref.row >= r.from.row && ref.row <= r.to.row && ref.col >= r.from.col && ref.col <= r.to.col)) return whole;
        const s0 = Number(attr(attrs, 's') || 0);
        const next = styles.withBorderStyles(s0, { left: 'thin', right: 'thin', top: 'thin', bottom: 'thin' });
        return next === s0 ? whole : `<c${attrs.replace(/\s*\bs="\d+"/, '')} s="${next}"${selfClose}>`;
      });
    }
    // 채우기 색 바꾸기: [{ range, fill: 'DDEBF7' | null(없앰) }]
    if (rule?.fills) {
      const targets = rule.fills.map((f) => ({ r: parseRange(f.range), fill: f.fill, noBorder: f.noBorder })).filter((f) => f.r);
      xml = xml.replace(/<c\b([^>]*?)(\/?)>/g, (whole, attrs, selfClose) => {
        const ref = parseRef(attr(attrs, 'r') || '');
        const hit = ref && targets.find((f) => ref.row >= f.r.from.row && ref.row <= f.r.to.row && ref.col >= f.r.from.col && ref.col <= f.r.to.col);
        if (!hit) return whole;
        const s0 = Number(attr(attrs, 's') || 0);
        let next = hit.fill ? styles.withFill(s0, hit.fill) : styles.withNoFill(s0);
        if (hit.noBorder) next = styles.withoutBorderSides(next, ['left', 'right', 'top', 'bottom']);
        return next === s0 ? whole : `<c${attrs.replace(/\s*\bs="\d+"/, '')} s="${next}"${selfClose}>`;
      });
    }
    // 빈 줄 숨기기(hideBlankRows: [행, …]): 칸에 값이 하나도 없는 줄만 숨긴다 (세로를 줄여 더 크게 인쇄되게)
    if (rule?.hideBlankRows) {
      const { kinds } = analyzeTables(xml, styles, sharedStrings);
      const used = new Set();
      for (const [key, kind] of kinds) if (kind && kind !== 'empty') used.add(Number(key.split(',')[0]));
      for (const row of rule.hideBlankRows) if (!used.has(row)) xml = ensureRowHidden(xml, row);
    }
    // 넓은 열 줄이기(scaleCols: { factor, minChars, floor }): minChars 이상인 열만 factor배로(floor보다 좁아지지는 않게).
    // 숫자 칸과, 오른쪽 칸에 값이 있어 넘쳐 쓸 수 없는 글자 칸은 '셀에 맞춤'으로 해서 잘리거나 ####이 되지 않게 한다
    if (rule?.scaleCols) {
      const { factor = 0.7, minChars = 7, floor = 6 } = rule.scaleCols;
      const range = printCols.get(sheetIndex);
      const fmtPr = /<sheetFormatPr\b([^>]*)\/?>/.exec(xml);
      const defaultCol = Number((fmtPr && attr(fmtPr[1], 'defaultColWidth')) || 8.88671875);
      const body = /<cols>([\s\S]*?)<\/cols>/.exec(xml);
      const els = body ? [...body[1].matchAll(/<col\b([^>]*?)\/>/g)].map((m) => ({ attrs: m[1], min: Number(attr(m[1], 'min')), max: Number(attr(m[1], 'max')) })) : [];
      if (range) {
        const keep = (attrs) => attrs.replace(/\s*\b(?:min|max|width|customWidth)="[^"]*"/g, '');
        const out = [];
        for (const el of els) {
          if (el.max < range.from || el.min > range.to) { out.push({ min: el.min, tag: `<col${el.attrs}/>` }); continue; }
          const w = attr(el.attrs, 'width');
          if (el.min < range.from) out.push({ min: el.min, tag: `<col min="${el.min}" max="${range.from - 1}"${w !== null ? ` width="${w}" customWidth="1"` : ''}${keep(el.attrs)}/>` });
          if (el.max > range.to) out.push({ min: range.to + 1, tag: `<col min="${range.to + 1}" max="${el.max}"${w !== null ? ` width="${w}" customWidth="1"` : ''}${keep(el.attrs)}/>` });
        }
        for (let c = range.from; c <= range.to; c++) {
          const el = els.find((x) => c >= x.min && c <= x.max);
          const w = el && attr(el.attrs, 'width') !== null ? Number(attr(el.attrs, 'width')) : defaultCol;
          const next = w >= minChars ? Math.max(floor, Math.round(w * factor * 100) / 100) : w;
          out.push({ min: c, tag: `<col min="${c}" max="${c}" width="${next}" customWidth="1"${el ? keep(el.attrs) : ''}/>` });
        }
        out.sort((a, b) => a.min - b.min);
        const colsXml = `<cols>${out.map((x) => x.tag).join('')}</cols>`;
        xml = body ? xml.replace(/<cols>[\s\S]*?<\/cols>/, () => colsXml) : xml.replace(/<sheetData\b/, (w) => `${colsXml}${w}`);
        const { kinds, K } = analyzeTables(xml, styles, sharedStrings);
        xml = xml.replace(/<c\b([^>]*?)(\/?)>/g, (whole, attrs, selfClose) => {
          const ref = parseRef(attr(attrs, 'r') || '');
          if (!ref) return whole;
          const kind = kinds.get(K(ref.row, ref.col));
          if (!kind || kind === 'empty') return whole;
          const right = kinds.get(K(ref.row, ref.col + 1));
          if (kind === 'text' && (!right || right === 'empty')) return whole;
          const s0 = Number(attr(attrs, 's') || 0);
          const next = styles.withShrink(s0);
          return next === s0 ? whole : `<c${attrs.replace(/\s*\bs="\d+"/, '')} s="${next}"${selfClose}>`;
        });
      }
    }
    // 열 너비 고르기: 같은 머리말 아래 칸들(전·후·차, 입구·출구 …)은 같은 너비로, 거의 같은 이웃 열도 같은 너비로
    if (rule && rule.equalCols !== false && rule.align !== false) xml = equalizeColumns(xml, styles, sharedStrings, printCols.get(sheetIndex) || null);
    if (rule?.borders !== false) xml = normalizeTableBorders(xml, styles, sharedStrings);
    if (rule?.cellStyles !== false) xml = normalizeCellStyles(xml, styles, { sharedStrings, cols: printCols.get(sheetIndex) || null, align: rule?.align !== false });
    const landscape = /<pageSetup\b[^>]*orientation="landscape"/.test(xml);
    if (equalWidth && !multiAreaSheets.has(sheetIndex) && !landscape) sheetsForWidth.push({ path, index: sheetIndex, name });
    if (rule?.shrink !== false) xml = shrinkTableCells(xml, styles, sharedStrings);
    // 글자가 정해진 모양(shrinkText 정규식)인 칸도 '셀에 맞춤' (예: 계산식의 'Q =', 'I(%) =' 같은 좁은 칸의 라벨이 잘리지 않게)
    if (rule?.shrinkText) {
      const { texts, K } = analyzeTables(xml, styles, sharedStrings);
      xml = xml.replace(/<c\b([^>]*?)(\/?)>/g, (whole, attrs, selfClose) => {
        const ref = parseRef(attr(attrs, 'r') || '');
        const text = String(texts.get(K(ref.row, ref.col)) || '');
        if (!ref || !rule.shrinkText.test(text)) return whole;
        if (rule.shrinkRows && (ref.row < rule.shrinkRows[0] || ref.row > rule.shrinkRows[1])) return whole;
        const s0 = Number(attr(attrs, 's') || 0);
        // 줄 바꿈이 켜진 칸도 한 줄짜리 글이면 줄 바꿈을 끄고 줄인다 (끝의 ':'만 아래 줄로 떨어지지 않게)
        const next = styles.withShrink(s0, true);
        return next === s0 ? whole : `<c${attrs.replace(/\s*\bs="\d+"/, '')} s="${next}"${selfClose}>`;
      });
    }
    if (rule?.spacing !== false) xml = normalizeSectionSpacing(xml, styles, { sharedStrings, drawingRows: await readDrawingRows(zip, path, xml), cols: printCols.get(sheetIndex) || null });
    zip.file(path, xml);
  }
  // 인쇄 크기 통일 (여백 같게, 가로 가운데, 같은 폭)
  if (equalWidth && sheetsForWidth.length) {
    const margins = { left: 0.4, right: 0.4, top: 0.5, bottom: 0.5, header: 0.3, footer: 0.3 };
    const printableW = A4.width - (margins.left + margins.right) * 72;
    const printableH = A4.height - (margins.top + margins.bottom) * 72;
    const baseFont = /<font\b[^>]*>[\s\S]*?<sz val="([\d.]+)"/.exec(stylesXml);
    // 기본 글꼴의 숫자 폭(px). 11pt 기준 7~9px로 환경마다 다르므로 넓은 쪽(9px)으로 잡아 넘치지 않게 한다
    const mdw = 9 * (baseFont ? Number(baseFont[1]) / 11 : 1);
    const measured = [];
    for (const item of sheetsForWidth) {
      const xml = await read(item.path);
      const range = printAreaUpdates.has(item.index) ? parseRange(printAreaUpdates.get(item.index).range) : printRanges.get(item.index) || null;
      const size = xml ? measurePrintArea(xml, range, mdw, styles, range ? null : await readDrawingExtent(zip, item.path, xml)) : null;
      // 인쇄 범위가 없는 시트에서 인쇄할 끝 열을 정해 둔 경우(maxPrintCol): 그 오른쪽의 보조 칸은 인쇄 범위에 넣지 않는다
      const itemRule = typeof rules === 'function' ? rules(item.name) : rules[item.name];
      if (size && size.fromContent && itemRule?.maxPrintCol && size.range.to.col > itemRule.maxPrintCol) {
        const clipped = { from: size.range.from, to: { row: size.range.to.row, col: itemRule.maxPrintCol } };
        const again = measurePrintArea(xml, clipped, mdw, styles);
        if (again && again.width > 0) { measured.push({ ...item, ...again, range: clipped, fromContent: true }); continue; }
      }
      if (size && size.width > 0 && size.height > 0) measured.push({ ...item, ...size });
    }
    // 가로는 글꼴 폭 어림 오차로 3%, 세로는 머리글·바닥글 공간으로 5% 여유
    const maxW = printableW * 0.97;
    const maxH = printableH * 0.95;
    const bound = Math.min(maxW, ...measured.map((x) => (maxH / x.height) * x.width));
    // 인쇄 범위가 없는 시트는 잰 범위(값·테두리가 있는 칸)를 인쇄 범위로 지정한다 (엑셀이 빈 서식 칸까지 잡지 않게)
    for (const x of measured) {
      if (!x.fromContent || printAreaUpdates.has(x.index)) continue;
      printAreaUpdates.set(x.index, { name: x.name, range: `A1:${numToColName(x.range.to.col)}${x.range.to.row}` });
    }
    const fixedWidth = typeof equalWidth === 'object' && equalWidth.width ? equalWidth.width : null;
    // 배율은 정수(%)만 되므로, 배율이 작은 시트(계산식 등)는 반올림 오차로 폭이 어긋난다.
    // 공통 폭을 상한에서 조금씩 낮춰 가며, 모든 시트가 들어가면서 장끼리 폭 차이가 가장 작은 값을 고른다
    const scalesFor = (common, round) => measured.map((x) => Math.max(10, Math.min(400, round((common / x.width) * 100))));
    let best = null;
    for (let k = 0; k <= 120; k++) {
      const common = bound * (1 - k * 0.001);
      const scales = scalesFor(common, Math.round);
      const widths = measured.map((x, i) => (x.width * scales[i]) / 100);
      const fits = measured.every((x, i) => widths[i] <= maxW + 0.01 && (x.height * scales[i]) / 100 <= maxH + 0.01);
      if (!fits) continue;
      const spread = Math.max(...widths) / Math.min(...widths);
      // 폭 차이가 0.3% 넘게 줄어들 때만 더 작은 공통 폭을 택한다 (괜히 작아지지 않게)
      if (!best || spread < best.spread - 0.003) best = { spread, scales };
    }
    let scales = best ? best.scales : scalesFor(bound, Math.floor);
    // 기록부 종류가 달라도 같은 폭이 되게 정해 둔 공통 폭이 있으면 그 폭으로 (이 파일 시트가 모두 들어갈 때만)
    if (fixedWidth && fixedWidth <= bound) {
      const fixed = scalesFor(fixedWidth, Math.round);
      if (measured.every((x, i) => (x.width * fixed[i]) / 100 <= maxW + 0.01 && (x.height * fixed[i]) / 100 <= maxH + 0.01)) scales = fixed;
    }
    for (let i = 0; i < measured.length; i++) {
      let xml = await read(measured[i].path);
      // 제목·양식번호 줄·바깥 테두리 통일: 인쇄 배율이 정해진 뒤에 한다 (인쇄 크기 기준으로 글자·높이를 맞추므로)
      const rule = typeof rules === 'function' ? rules(measured[i].name) : rules[measured[i].name];
      if (furniture && rule?.furniture !== false) {
        const next = normalizePageFurniture(xml, styles, { sharedStrings, range: measured[i].range, scale: scales[i] / 100, ...furniture, ...(rule?.furniture || {}) });
        // 제목·양식번호 줄 높이가 바뀌어도 한 장에 들어갈 때만 적용한다
        const size = measurePrintArea(next, measured[i].range, mdw, styles);
        if (size && (size.height * scales[i]) / 100 <= maxH + 0.01) xml = next;
      }
      // 쪽 채우기(fillRows: [첫 행, 끝 행]): 그 행들의 높이를 늘려 양식번호 줄이 A4 맨 아래에 오게 한다
      if (rule?.fillRows) {
        const size = measurePrintArea(xml, measured[i].range, mdw, styles);
        const [first, last] = rule.fillRows;
        const extra = size ? maxH / (scales[i] / 100) - size.height : 0;
        if (extra > 1 && last >= first) {
          const fmt = /<sheetFormatPr\b([^>]*)\/?>/.exec(xml);
          const defaultRow = Number((fmt && attr(fmt[1], 'defaultRowHeight')) || 15);
          const heights = {};
          for (let row = first; row <= last; row++) {
            const m = new RegExp(`<row\\b(?=[^>]*\\br="${row}")([^>]*?)/?>`).exec(xml);
            const ht = m && attr(m[1], 'ht') !== null ? Number(attr(m[1], 'ht')) : defaultRow;
            heights[row] = Math.round((ht + extra / (last - first + 1)) * 100) / 100;
          }
          for (const [row, ht] of Object.entries(heights)) xml = ensureRowHeight(xml, Number(row), ht);
        }
      }
      // 양식번호 줄을 A4 맨 아래로(furniture.footerAtBottom)
      // 1) 양식번호 줄 위쪽에 '줄 전체가 빈' 줄이 있으면 그 줄을 늘린다 (인쇄 범위 밖 오른쪽에 표가 있는 줄은 빈 줄이 아니다)
      // 2) 없으면 한 줄만 길어지지 않게, 제목·양식번호 줄을 뺀 모든 줄을 같은 비율로 조금씩 늘린다
      else if (furniture && furniture.footerAtBottom && !furniture.pageFooter && rule?.furniture !== false && measured[i].range) {
        const size = measurePrintArea(xml, measured[i].range, mdw, styles);
        let extra = size ? maxH / (scales[i] / 100) - size.height : 0;
        const { from, to } = measured[i].range;
        const hasFooter = new RegExp(`<c\\b[^>]*\\br="[A-Z]+${to.row}"[^>]*>`).test(xml);
        if (extra > 1 && hasFooter && to.row > from.row) {
          const { cells, kinds, merges } = analyzeTables(xml, styles, sharedStrings);
          const fmt = /<sheetFormatPr\b([^>]*)\/?>/.exec(xml);
          const defaultRow = Number((fmt && attr(fmt[1], 'defaultRowHeight')) || 15);
          const rowInfo = (row) => {
            const m = new RegExp(`<row\\b(?=[^>]*\\br="${row}")([^>]*?)/?>`).exec(xml);
            return { hidden: !!(m && isTrue(attr(m[1], 'hidden'))), ht: m && attr(m[1], 'ht') !== null ? Number(attr(m[1], 'ht')) : defaultRow };
          };
          // 줄 전체(모든 열)에 값도 테두리도 없는 줄
          const used = new Set();
          for (const [key, kind] of kinds) {
            const row = Number(key.split(',')[0]);
            if (used.has(row)) continue;
            const sides = styles.sidesOf(cells.get(key) || 0) || {};
            if ((kind && kind !== 'empty') || Object.values(sides).some(Boolean)) used.add(row);
          }
          for (const m of merges) for (let row = m.from.row; row <= m.to.row; row++) if (m.to.row > m.from.row) used.add(row);
          const MAX_ROW = 400; // 엑셀의 행 높이 한계(409.5pt)보다 작게
          const blanks = [];
          // 양식번호 줄에서 위로 올라가며, 보이는 내용 줄을 만나기 전까지의 빈 줄을 모은다.
          // 숨긴 줄 가운데 빈 줄(묶음 사이 띄움 줄)은 다시 보이게 해서 쓴다 (유량 기록부처럼 아래쪽 묶음을 통째로 숨긴 시트)
          let passedContent = false; // 보이는 내용 줄을 지났는가 (그 위의 숨긴 줄은 쓰지 않는다)
          for (let row = to.row - 1; row > from.row; row--) {
            const info = rowInfo(row);
            if (used.has(row)) { if (!info.hidden) passedContent = true; }
            else if (info.hidden) { if (!passedContent) blanks.push({ row, ht: 0, hidden: true }); }
            else if (row >= to.row - 6) blanks.push({ row, ht: info.ht, hidden: false });
            if (passedContent && row < to.row - 6) break;
          }
          // 보이는 빈 줄이 있으면 그 줄만 쓴다
          if (blanks.some((x) => !x.hidden)) blanks.splice(0, blanks.length, ...blanks.filter((x) => !x.hidden));
          if (blanks.length) {
            for (const { row, ht, hidden } of blanks) {
              if (extra <= 0.5) break;
              const add = Math.min(extra, MAX_ROW - ht);
              if (add <= 0) continue;
              xml = ensureRowHeight(xml, row, Math.round((ht + add) * 100) / 100);
              if (hidden) xml = xml.replace(new RegExp(`(<row\\b(?=[^>]*\\br="${row}")[^>]*?)\\s*hidden="[^"]*"`), (w, a2) => a2);
              extra -= add;
            }
          } else {
            const title = merges.find((m) => m.from.row === from.row);
            const first = (title ? title.to.row : from.row) + 1;
            const body = [];
            for (let row = first; row < to.row; row++) { const info = rowInfo(row); if (!info.hidden) body.push({ row, ht: info.ht }); }
            const total = body.reduce((sum, x) => sum + x.ht, 0);
            // 너무 많이 늘려야 하면(30% 넘게) 줄들이 부풀어 보이므로 그대로 둔다
            if (total > 0 && extra / total <= 0.3) {
              const factor = 1 + extra / total;
              for (const { row, ht } of body) xml = ensureRowHeight(xml, row, Math.floor(ht * factor * 100) / 100);
            }
          }
        }
      }
      // 양식번호를 인쇄 바닥글로(furniture.pageFooter): 양식번호 줄(인쇄 범위의 마지막 줄)의 글자를 바닥글 오른쪽에 넣고 그 줄은 숨긴다.
      // 줄을 억지로 늘리지 않아도 인쇄하면 항상 종이 맨 아래에 찍힌다. 인쇄 배율과 상관없이 같은 크기(scaleWithDoc 0)
      let footerText = '';
      if (furniture && furniture.pageFooter && rule?.furniture !== false && measured[i].range) {
        const { to } = measured[i].range;
        const { texts, kinds: footKinds } = analyzeTables(xml, styles, sharedStrings);
        // 양식번호 줄 바로 윗줄이 값이 하나도 없는 빈 줄인가 (표의 마지막 줄이 아니다)
        let aboveBlank = true;
        for (const [key, kind] of footKinds) if (Number(key.split(',')[0]) === to.row - 1 && kind && kind !== 'empty') { aboveBlank = false; break; }
        for (const [key, text] of texts) if (Number(key.split(',')[0]) === to.row && String(text || '').trim()) { footerText = String(text).trim(); break; }
        if (footerText && (!furniture.footerPattern || furniture.footerPattern.test(footerText))) {
          xml = ensureRowHidden(xml, to.row);
          // 양식번호 줄 위에 그었던 마감선은 지운다: 그 줄이 숨겨지면 아래 묶음의 윗선과 붙어 굵은 선이 두 줄로 보인다
          xml = xml.replace(/<c\b([^>]*?)(\/?)>/g, (whole, attrs, selfClose) => {
            const ref = parseRef(attr(attrs, 'r') || '');
            // 묶음 구분선을 쓰는 시트(계산식)에서만: 다른 시트는 그 줄이 표의 마지막 줄일 수 있다
            if (!ref || ref.row !== to.row - 1 || !(aboveBlank || (rule?.furniture && rule.furniture.sectionRule))) return whole;
            const s0 = Number(attr(attrs, 's') || 0);
            const next = styles.withoutBorderSides(s0, ['bottom']);
            return next === s0 ? whole : `<c${attrs.replace(/\s*\bs="\d+"/, '')} s="${next}"${selfClose}>`;
          });
        }
        else footerText = '';
      }
      let outXml = setPrintScale(xml, scales[i], margins);
      if (footerText) {
        const esc = footerText.replace(/&/g, '&&').replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
        const tag = `<headerFooter scaleWithDoc="0" alignWithMargins="1"><oddFooter>&amp;R&amp;"맑은 고딕,보통"&amp;${furniture.footerPt || 9}${esc}</oddFooter></headerFooter>`;
        if (/<headerFooter\b[^>]*(?:\/>|>[\s\S]*?<\/headerFooter>)/.test(outXml)) outXml = outXml.replace(/<headerFooter\b[^>]*(?:\/>|>[\s\S]*?<\/headerFooter>)/, () => tag);
        else outXml = outXml.replace(/(<pageSetup\b[^>]*\/>)/, (w) => `${w}${tag}`);
      }
      zip.file(measured[i].path, outXml);
    }
  }
  zip.file('xl/styles.xml', styles.serialize());
  // 인쇄 범위 바꾸기 (프로그램이 표를 넓혀 그린 시트), 숨긴 시트의 인쇄 범위 지우기
  if (printAreaUpdates.size || hiddenSheets.size) {
    let wb = workbookXml;
    for (const index of hiddenSheets) {
      wb = wb.replace(new RegExp(`<definedName\\b(?=[^>]*name="_xlnm\\.Print_Area")(?=[^>]*localSheetId="${index}")[^>]*>[^<]*</definedName>`), '');
    }
    wb = wb.replace(/<definedNames>\s*<\/definedNames>/, '');
    for (const [index, { name, range }] of printAreaUpdates) {
      const [a, b] = range.split(':');
      const abs = (ref) => ref.replace(/^([A-Z]+)(\d+)$/, '$$$1$$$2');
      const text = `'${name.replace(/'/g, "''")}'!${abs(a)}:${abs(b || a)}`.replace(/&/g, '&amp;').replace(/</g, '&lt;');
      const re = new RegExp(`<definedName\\b(?=[^>]*name="_xlnm\\.Print_Area")(?=[^>]*localSheetId="${index}")[^>]*>[^<]*</definedName>`);
      const tag = `<definedName name="_xlnm.Print_Area" localSheetId="${index}">${text}</definedName>`;
      // 바꿀 글(인쇄 범위)에 $A$1처럼 '$'가 들어 있으므로 문자열이 아니라 함수로 넘긴다
      // (문자열로 넘기면 '$1'·'$18'이 '찾은 부분' 기호로 해석돼 $A$1:$H$46이 $A:$H$46으로 깨졌다)
      if (re.test(wb)) wb = wb.replace(re, () => tag);
      else if (/<definedNames\b[^>]*>/.test(wb)) wb = wb.replace(/<definedNames\b([^>]*)>/, (w, a) => `<definedNames${a}>${tag}`);
      else wb = wb.replace(/(<\/sheets>)/, (w) => `${w}<definedNames>${tag}</definedNames>`);
    }
    zip.file('xl/workbook.xml', wb);
  }
}

// ── 열 너비 고르기 ──
// 1) 가로로 병합된 머리말 바로 아래에 칸마다 작은 머리말이 있으면(가스미터지시값 → 전·후·차) 그 열들을 같은 너비로
// 2) 이웃한 열들의 너비가 거의 같으면(가장 넓은 열 ÷ 가장 좁은 열 ≤ tolerance) 같은 너비로
// 전체 너비(합)는 그대로이다. cols: 인쇄 범위의 [첫 열, 끝 열]
export function equalizeColumns(xml, styles, sharedStrings, cols = null, { tolerance = 1.22 } = {}) {
  const fmt = /<sheetFormatPr\b([^>]*)\/?>/.exec(xml);
  const baseCol = Number((fmt && attr(fmt[1], 'baseColWidth')) || 8);
  const defaultCol = Number((fmt && attr(fmt[1], 'defaultColWidth')) || baseCol + 0.71);
  const body = /<cols>([\s\S]*?)<\/cols>/.exec(xml);
  const els = body ? [...body[1].matchAll(/<col\b([^>]*?)\/>/g)].map((m) => ({ attrs: m[1], min: Number(attr(m[1], 'min')), max: Number(attr(m[1], 'max')) })) : [];
  const { cells, texts, merges, mergeOf, K } = analyzeTables(xml, styles, sharedStrings);
  const first = cols ? cols.from : 1; let last = cols ? cols.to : 0;
  if (!cols) for (const key of cells.keys()) last = Math.max(last, Number(key.split(',')[1]));
  if (!(last > first)) return xml;
  const elOf = (col) => els.find((x) => col >= x.min && col <= x.max);
  if ([...Array(last - first + 1)].some((_, i) => { const el = elOf(first + i); return el && isTrue(attr(el.attrs, 'hidden')); })) return xml;
  const width = new Map();
  for (let col = first; col <= last; col++) { const el = elOf(col); width.set(col, el && attr(el.attrs, 'width') !== null ? Number(attr(el.attrs, 'width')) : defaultCol); }
  const before = new Map(width);
  const fixed = new Set(); // 이미 묶음으로 맞춘 열
  const setMean = (from, to) => {
    let sum = 0; for (let col = from; col <= to; col++) sum += width.get(col);
    const mean = Math.round((sum / (to - from + 1)) * 100) / 100;
    for (let col = from; col <= to; col++) { width.set(col, mean); fixed.add(col); }
  };
  // 1) 병합 머리말 아래의 형제 칸
  const groups = [];
  for (const m of merges) {
    if (m.to.col <= m.from.col || m.from.col < first || m.to.col > last) continue;
    if (!String(texts.get(K(m.from.row, m.from.col)) || '').trim()) continue;
    const below = m.to.row + 1;
    let ok = true;
    for (let col = m.from.col; col <= m.to.col && ok; col++) {
      const key = K(below, col);
      const own = mergeOf.has(key) ? merges[mergeOf.get(key)] : null;
      if (!String(texts.get(key) || '').trim() || (own && own.to.col > own.from.col)) ok = false;
    }
    if (ok) groups.push([m.from.col, m.to.col]);
  }
  // 넓은 묶음부터 (겹치면 넓은 쪽이 이긴다)
  groups.sort((a, b) => (b[1] - b[0]) - (a[1] - a[0]));
  for (const [from, to] of groups) {
    let free = true; for (let col = from; col <= to; col++) if (fixed.has(col)) free = false;
    if (free) setMean(from, to);
  }
  // 2) 거의 같은 이웃 열 (묶음으로 맞춘 열은 묶음째 한 단위로 본다)
  let col = first;
  while (col <= last) {
    let end = col; let lo = width.get(col); let hi = lo;
    while (end + 1 <= last) {
      const w = width.get(end + 1);
      if (Math.max(hi, w) / Math.min(lo, w) > tolerance) break;
      lo = Math.min(lo, w); hi = Math.max(hi, w); end += 1;
    }
    if (end > col && hi - lo > 0.005) {
      let sum = 0; for (let c = col; c <= end; c++) sum += width.get(c);
      const mean = Math.round((sum / (end - col + 1)) * 100) / 100;
      for (let c = col; c <= end; c++) width.set(c, mean);
    }
    col = end + 1;
  }
  let changed = false;
  for (const [c, w] of width) if (Math.abs(w - before.get(c)) > 0.005) changed = true;
  if (!changed) return xml;
  // <cols>를 열 하나씩으로 다시 쓴다 (범위 밖 열은 그대로)
  const out = [];
  const maxCol = Math.max(last, ...els.map((x) => Math.min(x.max, 16384)));
  const keep = (attrs) => attrs.replace(/\s*\b(?:min|max|width|customWidth)="[^"]*"/g, '');
  for (const el of els) {
    if (el.max < first || el.min > last) { out.push({ min: el.min, tag: `<col${el.attrs}/>` }); continue; }
    if (el.min < first) out.push({ min: el.min, tag: `<col min="${el.min}" max="${first - 1}"${attr(el.attrs, 'width') !== null ? ` width="${attr(el.attrs, 'width')}" customWidth="1"` : ''}${keep(el.attrs)}/>` });
    if (el.max > last) out.push({ min: last + 1, tag: `<col min="${last + 1}" max="${el.max}"${attr(el.attrs, 'width') !== null ? ` width="${attr(el.attrs, 'width')}" customWidth="1"` : ''}${keep(el.attrs)}/>` });
  }
  void maxCol;
  for (let c = first; c <= last; c++) {
    const el = elOf(c);
    out.push({ min: c, tag: `<col min="${c}" max="${c}" width="${width.get(c)}" customWidth="1"${el ? keep(el.attrs) : ''}/>` });
  }
  out.sort((a, b) => a.min - b.min);
  const colsXml = `<cols>${out.map((x) => x.tag).join('')}</cols>`;
  return body ? xml.replace(/<cols>[\s\S]*?<\/cols>/, () => colsXml) : xml.replace(/<sheetData\b/, (w) => `${colsXml}${w}`);
}

// ── 표 상자 테두리: 범위(예: 'A57:K58') 안 모든 칸에 가는 선, 범위 바깥 둘레는 outer 굵기 ──
// boxes: [{ path: 시트 파일 경로, range }]. 한 파일의 서식(styles.xml)을 한 번만 읽고 쓴다
export async function boxSheetRanges(zip, boxes, { outer = 'medium' } = {}) {
  const stylesFile = zip.file('xl/styles.xml');
  if (!stylesFile) return 0;
  const styles = new StyleBook(await stylesFile.async('string'));
  let count = 0;
  for (const { path, range } of boxes) {
    const file = zip.file(path); const r = parseRange(range || '');
    if (!file || !r) continue;
    const xml = await file.async('string');
    zip.file(path, xml.replace(/<c\b([^>]*?)(\/?)>/g, (whole, attrs, selfClose) => {
      const ref = parseRef(attr(attrs, 'r') || '');
      if (!ref || ref.row < r.from.row || ref.row > r.to.row || ref.col < r.from.col || ref.col > r.to.col) return whole;
      const s0 = Number(attr(attrs, 's') || 0);
      const next = styles.withBorderStyles(s0, {
        left: ref.col === r.from.col ? outer : 'thin', right: ref.col === r.to.col ? outer : 'thin',
        top: ref.row === r.from.row ? outer : 'thin', bottom: ref.row === r.to.row ? outer : 'thin',
      });
      count += 1;
      return next === s0 ? whole : `<c${attrs.replace(/\s*\bs="\d+"/, '')} s="${next}"${selfClose}>`;
    }));
  }
  zip.file('xl/styles.xml', styles.serialize());
  return count;
}

// 행 범위 [시작, 끝]을 같은 높이로
export const rowsAt = (from, to, ht) => Object.fromEntries(Array.from({ length: to - from + 1 }, (_, i) => [from + i, ht]));

// ── 시트에 딸린 그림(drawing) 파일 경로들 ──
const resolveRelTarget = (dir, target) => {
  const stack = [];
  for (const part of (dir + target).split('/')) { if (part === '..') stack.pop(); else if (part && part !== '.') stack.push(part); }
  return stack.join('/');
};
const readSheetDrawingPaths = async (zip, sheetPath) => {
  const dir = sheetPath.replace(/[^/]+$/, '');
  const relsFile = zip.file(`${dir}_rels/${sheetPath.split('/').pop()}.rels`);
  if (!relsFile) return [];
  const rels = await relsFile.async('string');
  const out = [];
  for (const m of rels.matchAll(/<Relationship\b([^>]*)\/?>/g)) {
    if (/\/drawing$/.test(attr(m[1], 'Type') || '')) out.push(resolveRelTarget(dir, attr(m[1], 'Target') || ''));
  }
  return out;
};
const readSharedStringList = async (zip) => {
  const file = zip.file('xl/sharedStrings.xml');
  if (!file) return [];
  const xml = await file.async('string');
  return [...xml.matchAll(/<si>([\s\S]*?)<\/si>/g)].map((m) => [...m[1].matchAll(/<t\b[^>]*>([^<]*)<\/t>/g)].map((t) => decodeXml(t[1])).join(''));
};
// 칸 하나의 글자 (공유 문자열·인라인 문자열·수식 결과 글자)
const cellTextOf = (xml, ref, sharedStrings) => {
  const m = new RegExp(`<c\\b(?=[^>]*\\br="${ref}")([^>]*?)(?:/>|>([\\s\\S]*?)</c>)`).exec(xml);
  if (!m || !m[2]) return '';
  const t = attr(m[1], 't');
  if (t === 's') { const v = /<v>([^<]*)<\/v>/.exec(m[2]); return v ? String(sharedStrings[Number(v[1])] ?? '') : ''; }
  if (t === 'inlineStr') return [...m[2].matchAll(/<t\b[^>]*>([^<]*)<\/t>/g)].map((x) => decodeXml(x[1])).join('');
  if (t === 'str') { const v = /<v>([^<]*)<\/v>/.exec(m[2]); return v ? decodeXml(v[1]) : ''; }
  return '';
};
const ensureRowHidden = (xml, row) => {
  const re = new RegExp(`<row\\b([^>]*?)\\br="${row}"([^>]*?)(/?)>`);
  if (re.test(xml)) {
    return xml.replace(re, (w, a, b, close) => `<row${`${a}r="${row}"${b}`.replace(/\s*\bhidden="[^"]*"/, '')} hidden="1"${close}>`);
  }
  const tag = `<row r="${row}" hidden="1"/>`;
  let placed = false;
  const out = xml.replace(/<row\b[^>]*\br="(\d+)"/g, (w, n) => {
    if (!placed && Number(n) > row) { placed = true; return `${tag}${w}`; }
    return w;
  });
  if (placed) return out;
  if (/<sheetData\s*\/>/.test(out)) return out.replace(/<sheetData\s*\/>/, () => `<sheetData>${tag}</sheetData>`);
  return out.replace('</sheetData>', () => `${tag}</sheetData>`);
};

// ── 해당 없는 계산 묶음 숨기기 (예: 표준산소농도가 없을 때의 '표준산소 환산 배출가스 유량') ──
// blocks: [[첫 행, 끝 행], ...]. titleIncludes가 있으면 첫 행에 그 글자가 있는 묶음만 숨긴다 (양식이 바뀌어 엉뚱한 줄을 숨기지 않게)
// 숨긴 행 안에만 걸친 그림·도형은 지운다: '셀과 함께 이동만' 하는 그림은 행을 숨겨도 그대로 보여 다른 칸 위에 떠 버리므로
export async function hideRowBlocks(zip, sheetPath, blocks, { titleIncludes = '', titleCols = ['A', 'B', 'C'] } = {}) {
  const file = zip.file(sheetPath);
  if (!file || !Array.isArray(blocks) || blocks.length === 0) return [];
  let xml = await file.async('string');
  const sharedStrings = titleIncludes ? await readSharedStringList(zip) : [];
  const chosen = blocks.filter(([from, to]) => Number.isFinite(from) && Number.isFinite(to) && to >= from
    && (!titleIncludes || titleCols.some((col) => cellTextOf(xml, `${col}${from}`, sharedStrings).includes(titleIncludes))));
  if (chosen.length === 0) return [];
  const isHidden = (row) => chosen.some(([from, to]) => row >= from && row <= to);
  for (const [from, to] of chosen) for (let row = from; row <= to; row++) xml = ensureRowHidden(xml, row);
  zip.file(sheetPath, xml);
  for (const path of await readSheetDrawingPaths(zip, sheetPath)) {
    const drawingFile = zip.file(path);
    if (!drawingFile) continue;
    const drawing = await drawingFile.async('string');
    const next = drawing.replace(/<xdr:(twoCellAnchor|oneCellAnchor)\b[\s\S]*?<\/xdr:\1>/g, (anchor) => {
      const from = /<xdr:from>[\s\S]*?<xdr:row>(\d+)<\/xdr:row>/.exec(anchor);
      if (!from) return anchor;
      const to = /<xdr:to>[\s\S]*?<xdr:row>(\d+)<\/xdr:row>\s*<xdr:rowOff>(\d+)<\/xdr:rowOff>/.exec(anchor);
      const start = Number(from[1]) + 1;
      // 끝 칸의 위쪽 경계(rowOff 0)에서 끝나는 그림은 그 행을 차지하지 않는다
      const end = to ? Math.max(start, Number(to[1]) + (Number(to[2]) > 0 ? 1 : 0)) : start;
      for (let row = start; row <= end; row++) if (!isHidden(row)) return anchor;
      return '';
    });
    if (next !== drawing) zip.file(path, next);
  }
  return chosen;
}

// 칸 하나의 글자 읽기 (양식 확인용: 예상한 제목·꼬리말이 그 자리에 있을 때만 고친다)
export async function readSheetCellText(zip, sheetPath, ref) {
  const file = zip.file(sheetPath);
  if (!file) return '';
  return cellTextOf(await file.async('string'), ref, await readSharedStringList(zip));
}

// ── 한 행의 칸(값·서식·가로 병합)을 다른 행으로 옮기고 원래 행은 숨긴다 ──
// 예: 계산식의 꼬리말(■MAMO-QP-13-03)을 표준산소 환산 유량 묶음 아래로 옮겨, 그 묶음까지 한 장에 인쇄
// 수식이 있는 행이나, 옮길 자리에 걸친 병합이 있으면 건드리지 않는다.
// 옮길 자리(toRow)에 걸친 그림도 확인한다: 그 행 안에서 시작하는 그림은 다음 행 맨 위에서 시작하도록 내리고, 그만큼 다음 행을
// 높여 그 아래 칸과의 간격은 그대로 둔다 (그대로 두면 엑셀에서 꼬리말 줄 위에 그림 윗부분이 겹쳐 인쇄된다).
// 위에서 내려와 toRow에 걸치거나 toRow 안에서 끝나는 그림이 있으면 옮기지 않는다
export async function moveRowCells(zip, sheetPath, fromRow, toRow) {
  const file = zip.file(sheetPath);
  if (!file || fromRow === toRow) return false;
  let xml = await file.async('string');
  const rowRe = (r) => new RegExp(`<row\\b(?=[^>]*\\br="${r}")[^>]*?(?:/>|>[\\s\\S]*?</row>)`);
  const src = rowRe(fromRow).exec(xml);
  if (!src) return false;
  const fmt = /<sheetFormatPr\b([^>]*)\/?>/.exec(xml);
  const defaultHt = Number((fmt && attr(fmt[1], 'defaultRowHeight')) || 15);
  const heightOf = (text, r) => { const m = rowRe(r).exec(text); const ht = m ? attr(/^<row\b([^>]*?)\/?>/.exec(m[0])[1], 'ht') : null; return ht !== null ? Number(ht) : defaultHt; };
  const oldTargetHt = heightOf(xml, toRow);
  const anchorPos = (anchor, tag) => {
    const m = new RegExp(`<xdr:${tag}>[\\s\\S]*?<xdr:row>(\\d+)</xdr:row>\\s*<xdr:rowOff>(\\d+)</xdr:rowOff>`).exec(anchor);
    return m ? { row: Number(m[1]) + 1, off: Number(m[2]) } : null;
  };
  const drawings = [];
  for (const path of await readSheetDrawingPaths(zip, sheetPath)) {
    const drawingFile = zip.file(path);
    if (drawingFile) drawings.push({ path, xml: await drawingFile.async('string') });
  }
  const ANCHOR_RE = /<xdr:(twoCellAnchor|oneCellAnchor)\b[\s\S]*?<\/xdr:\1>/g;
  for (const { xml: drawing } of drawings) {
    for (const [anchor] of drawing.matchAll(ANCHOR_RE)) {
      const from = anchorPos(anchor, 'from');
      if (!from) continue;
      const to = anchorPos(anchor, 'to');
      const end = to ? (to.off > 0 ? to.row : to.row - 1) : from.row;
      if (from.row < toRow && end >= toRow) return false;
      if (from.row === toRow && end <= toRow) return false;
    }
  }
  const open = /^<row\b([^>]*?)\/?>/.exec(src[0]);
  const cells = src[0].endsWith('/>') && !src[0].includes('</row>') ? '' : src[0].replace(/^<row\b[^>]*>/, '').replace(/<\/row>$/, '');
  if (/<f\b/.test(cells)) return false;
  const merges = [...xml.matchAll(/<mergeCell\b[^>]*\bref="([A-Z]+)(\d+):([A-Z]+)(\d+)"[^>]*\/>/g)];
  if (merges.some((m) => Number(m[2]) <= toRow && Number(m[4]) >= toRow)) return false;
  const moved = cells.replace(/(<c\b[^>]*?\br=")([A-Z]+)\d+(")/g, (w, a, col, b) => `${a}${col}${toRow}${b}`);
  const attrs = open[1].replace(/\s*\br="\d+"/, '').replace(/\s*\bhidden="[^"]*"/, '');
  const newRow = `<row r="${toRow}"${attrs}>${moved}</row>`;
  if (rowRe(toRow).test(xml)) {
    xml = xml.replace(rowRe(toRow), () => newRow);
  } else {
    let placed = false;
    xml = xml.replace(/<row\b[^>]*\br="(\d+)"/g, (w, n) => {
      if (!placed && Number(n) > toRow) { placed = true; return `${newRow}${w}`; }
      return w;
    });
    if (!placed) xml = xml.replace('</sheetData>', () => `${newRow}</sheetData>`);
  }
  const added = merges.filter((m) => Number(m[2]) === fromRow && Number(m[4]) === fromRow).map((m) => `<mergeCell ref="${m[1]}${toRow}:${m[3]}${toRow}"/>`);
  if (added.length) {
    xml = xml.replace(/<mergeCells\b([^>]*)>([\s\S]*?)<\/mergeCells>/, (w, a, body) => {
      const nextBody = body + added.join('');
      return `<mergeCells${a.replace(/\s*\bcount="\d+"/, '')} count="${(nextBody.match(/<mergeCell\b/g) || []).length}">${nextBody}</mergeCells>`;
    });
  }
  xml = ensureRowHidden(xml, fromRow);
  // toRow 안에서 시작하던 그림을 다음 행 맨 위로 내린다
  const newTargetHt = heightOf(xml, toRow);
  let grow = 0;
  for (const drawing of drawings) {
    const next = drawing.xml.replace(ANCHOR_RE, (anchor) => {
      const from = anchorPos(anchor, 'from');
      if (!from || from.row !== toRow) return anchor;
      grow = Math.max(grow, oldTargetHt - from.off / 12700);
      const shiftEmu = Math.round(newTargetHt * 12700 - from.off);
      return anchor
        .replace(/(<xdr:from>[\s\S]*?)<xdr:row>\d+<\/xdr:row>(\s*)<xdr:rowOff>\d+<\/xdr:rowOff>/, (w, head, gap) => `${head}<xdr:row>${toRow}</xdr:row>${gap}<xdr:rowOff>0</xdr:rowOff>`)
        .replace(/(<xdr:spPr>[\s\S]*?<a:off\b[^>]*\by=")(\d+)(")/, (w, a, y, b) => `${a}${Number(y) + shiftEmu}${b}`);
    });
    if (next !== drawing.xml) zip.file(drawing.path, next);
  }
  void grow;
  zip.file(sheetPath, xml);
  return true;
}

// ── 그림을 칸(box, 예: 'E5:G10') 안에 맞춘다: 가로세로 비율(aspect = 가로 ÷ 세로)을 지키면서 가장 크게, 가운데에 ──
// 서식 정리가 끝난 뒤의 실제 열 너비·행 높이로 재고, 크기를 절대값(oneCellAnchor + ext)으로 적어
// 열 너비가 달라져도 그림이 찌그러지거나 칸 밖으로 나가지 않게 한다. 칸 안에서 시작하는 그림만 바꾼다.
export async function fitPictureInBox(zip, sheetPath, { box, aspect = 1, pad = 3, mdw = 9, fill = false } = {}) {
  const file = zip.file(sheetPath);
  const range = parseRange(box || '');
  if (!file || !range || !(aspect > 0)) return 0;
  const xml = await file.async('string');
  const relsFile = zip.file(sheetPath.replace(/([^/]+)$/, '_rels/$1.rels'));
  const rels = relsFile ? await relsFile.async('string') : '';
  const target = (/Type="[^"]*\/relationships\/drawing"[^>]*Target="([^"]+)"/.exec(rels) || /Target="([^"]+)"[^>]*Type="[^"]*\/relationships\/drawing"/.exec(rels) || [])[1];
  const drawingPath = target ? `xl/${target.replace(/^(\.\.\/)+/, '')}` : '';
  const drawingFile = drawingPath ? zip.file(drawingPath) : null;
  if (!drawingFile) return 0;
  const fmt = /<sheetFormatPr\b([^>]*)\/?>/.exec(xml);
  const defaultRow = Number((fmt && attr(fmt[1], 'defaultRowHeight')) || 15);
  const defaultCol = Number((fmt && attr(fmt[1], 'defaultColWidth')) || 8.88671875);
  const colsBody = /<cols>([\s\S]*?)<\/cols>/.exec(xml);
  const colEls = colsBody ? [...colsBody[1].matchAll(/<col\b([^>]*?)\/>/g)].map((m) => ({ attrs: m[1], min: Number(attr(m[1], 'min')), max: Number(attr(m[1], 'max')) })) : [];
  const EMU_PX = 9525; const EMU_PT = 12700;
  const colEmu = (col) => {
    const c = colEls.find((x) => col >= x.min && col <= x.max);
    const w = c && attr(c.attrs, 'width') !== null ? Number(attr(c.attrs, 'width')) : defaultCol;
    return Math.trunc(((256 * w + Math.trunc(128 / mdw)) / 256) * mdw) * EMU_PX;
  };
  const rowEmu = (row) => {
    const m = new RegExp(`<row\\b(?=[^>]*\\br="${row}")([^>]*?)/?>`).exec(xml);
    return Math.round((m && attr(m[1], 'ht') !== null ? Number(attr(m[1], 'ht')) : defaultRow) * EMU_PT);
  };
  let boxW = 0; for (let c = range.from.col; c <= range.to.col; c++) boxW += colEmu(c);
  let boxH = 0; for (let r = range.from.row; r <= range.to.row; r++) boxH += rowEmu(r);
  const availW = boxW - pad * 2 * EMU_PT; const availH = boxH - pad * 2 * EMU_PT;
  if (availW <= 0 || availH <= 0) return 0;
  let h = availH; let w = Math.round(h * aspect);
  if (w > availW) { w = availW; h = Math.round(w / aspect); }
  let x = Math.round((boxW - w) / 2); let y = Math.round((boxH - h) / 2);
  let col0 = range.from.col - 1; while (col0 < range.to.col - 1 && x >= colEmu(col0 + 1)) { x -= colEmu(col0 + 1); col0 += 1; }
  let row0 = range.from.row - 1; while (row0 < range.to.row - 1 && y >= rowEmu(row0 + 1)) { y -= rowEmu(row0 + 1); row0 += 1; }
  const drawing = await drawingFile.async('string');
  let count = 0;
  const next = drawing.replace(/<xdr:(twoCellAnchor|oneCellAnchor)\b[^>]*>([\s\S]*?)<\/xdr:\1>/g, (whole, kind, inner) => {
    const from = /<xdr:from>\s*<xdr:col>(\d+)<\/xdr:col>[\s\S]*?<xdr:row>(\d+)<\/xdr:row>[\s\S]*?<\/xdr:from>/.exec(inner);
    if (!from || !/<xdr:pic\b/.test(inner)) return whole;
    const c = Number(from[1]) + 1; const r = Number(from[2]) + 1;
    if (c < range.from.col || c > range.to.col || r < range.from.row || r > range.to.row) return whole;
    count += 1;
    const body = inner.replace(/<xdr:from>[\s\S]*?<\/xdr:from>/, '').replace(/<xdr:to>[\s\S]*?<\/xdr:to>/, '').replace(/<xdr:ext\b[^>]*\/>/, '')
      .replace(/<a:xfrm>[\s\S]*?<\/a:xfrm>/, () => `<a:xfrm><a:off x="0" y="0"/><a:ext cx="${w}" cy="${h}"/></a:xfrm>`);
    // fill: 칸 전체를 채운다 (칸 모서리에 붙여 두 칸 기준으로 고정 → 열 너비가 달라도 칸에 꼭 맞는다)
    if (fill) {
      const p = pad * EMU_PT;
      const fw = boxW - p * 2; const fh = boxH - p * 2;
      const filled = inner.replace(/<xdr:from>[\s\S]*?<\/xdr:from>/, '').replace(/<xdr:to>[\s\S]*?<\/xdr:to>/, '').replace(/<xdr:ext\b[^>]*\/>/, '')
        .replace(/<a:xfrm>[\s\S]*?<\/a:xfrm>/, () => `<a:xfrm><a:off x="0" y="0"/><a:ext cx="${fw}" cy="${fh}"/></a:xfrm>`);
      return `<xdr:twoCellAnchor editAs="twoCell"><xdr:from><xdr:col>${range.from.col - 1}</xdr:col><xdr:colOff>${p}</xdr:colOff><xdr:row>${range.from.row - 1}</xdr:row><xdr:rowOff>${p}</xdr:rowOff></xdr:from>`
        + `<xdr:to><xdr:col>${range.to.col - 1}</xdr:col><xdr:colOff>${colEmu(range.to.col) - p}</xdr:colOff><xdr:row>${range.to.row - 1}</xdr:row><xdr:rowOff>${rowEmu(range.to.row) - p}</xdr:rowOff></xdr:to>${filled}</xdr:twoCellAnchor>`;
    }
    return `<xdr:oneCellAnchor><xdr:from><xdr:col>${col0}</xdr:col><xdr:colOff>${x}</xdr:colOff><xdr:row>${row0}</xdr:row><xdr:rowOff>${y}</xdr:rowOff></xdr:from><xdr:ext cx="${w}" cy="${h}"/>${body}</xdr:oneCellAnchor>`;
  });
  if (count) zip.file(drawingPath, next);
  return fill ? { count, aspect: (boxW - pad * 2 * EMU_PT) / (boxH - pad * 2 * EMU_PT) } : count;
}

// ── 연결 그림(카메라)의 원본 칸을 정해진 비율(aspect = 가로 ÷ 세로)로: 열을 넓히고 칸 안의 그림을 가로 가운데로 옮긴다 ──
// 원본 칸과 그림 틀의 비율이 같아야 연결 그림이 찌그러지지 않고 틀을 꽉 채운다. 그림 크기는 그대로 둔다.
export async function setPictureSourceAspect(zip, sheetPath, { aspect, row = 1, mdw = 9 } = {}) {
  const file = zip.file(sheetPath);
  if (!file || !(aspect > 0)) return 0;
  let xml = await file.async('string');
  const rowTag = new RegExp(`<row\\b(?=[^>]*\\br="${row}")([^>]*?)/?>`).exec(xml);
  const ht = rowTag && attr(rowTag[1], 'ht') !== null ? Number(attr(rowTag[1], 'ht')) : 0;
  if (!(ht > 0)) return 0;
  const px = Math.round(((ht / 0.75) * aspect));
  const width = Math.round((px / mdw) * 10000) / 10000;
  const colPx = Math.trunc(((256 * width + Math.trunc(128 / mdw)) / 256) * mdw);
  xml = xml.replace(/<cols>([\s\S]*?)<\/cols>/, (w, body) => `<cols>${body.replace(/\bwidth="[^"]*"/g, `width="${width}"`)}</cols>`);
  zip.file(sheetPath, xml);
  const relsFile = zip.file(sheetPath.replace(/([^/]+)$/, '_rels/$1.rels'));
  const rels = relsFile ? await relsFile.async('string') : '';
  const target = (/Type="[^"]*\/relationships\/drawing"[^>]*Target="([^"]+)"/.exec(rels) || /Target="([^"]+)"[^>]*Type="[^"]*\/relationships\/drawing"/.exec(rels) || [])[1];
  const drawingPath = target ? `xl/${target.replace(/^(\.\.\/)+/, '')}` : '';
  const drawingFile = drawingPath ? zip.file(drawingPath) : null;
  if (!drawingFile) return 0;
  const colEmu = colPx * 9525;
  let count = 0;
  const drawing = await drawingFile.async('string');
  const next = drawing.replace(/<xdr:twoCellAnchor\b([^>]*)>([\s\S]*?)<\/xdr:twoCellAnchor>/g, (whole, attrs, inner) => {
    const from = /<xdr:from>\s*<xdr:col>(\d+)<\/xdr:col>\s*<xdr:colOff>(-?\d+)<\/xdr:colOff>/.exec(inner);
    const ext = /<a:ext cx="(\d+)" cy="(\d+)"/.exec(inner);
    if (!from || !ext || !/<xdr:pic\b/.test(inner)) return whole;
    const w = Number(ext[1]);
    const left = Math.max(0, Math.round((colEmu - w) / 2));
    count += 1;
    const body = inner
      .replace(/(<xdr:from>\s*<xdr:col>\d+<\/xdr:col>\s*<xdr:colOff>)-?\d+/, (m, a) => `${a}${left}`)
      .replace(/<xdr:to>\s*<xdr:col>\d+<\/xdr:col>\s*<xdr:colOff>-?\d+/, () => `<xdr:to><xdr:col>${from[1]}</xdr:col><xdr:colOff>${left + w}`);
    return `<xdr:twoCellAnchor${attrs}>${body}</xdr:twoCellAnchor>`;
  });
  if (count) zip.file(drawingPath, next);
  return count;
}

// ── 그림을 인쇄 폭(firstCol~lastCol) 가로 가운데로 (표지 로고) ──
// 크기는 그대로 두고 위치만 옮긴다. 열 폭은 양식을 만든 엑셀의 기준(맑은 고딕 11pt, 숫자 폭 9px, 기본 열 80px)으로 잰다:
// 양식에 저장된 로고 위치(a:off)가 이 기준으로 'B열 + 여백'과 정확히 맞는다
export async function centerSheetPictures(zip, sheetPath, { firstCol = 1, lastCol = 8, mdw = 9, defaultWidth = 8.88671875 } = {}) {
  const file = zip.file(sheetPath);
  if (!file || lastCol < firstCol) return 0;
  const xml = await file.async('string');
  const pxOf = (w) => Math.trunc(((256 * w + Math.trunc(128 / mdw)) / 256) * mdw);
  const colsBody = /<cols>([\s\S]*?)<\/cols>/.exec(xml);
  const colEls = colsBody ? [...colsBody[1].matchAll(/<col\b([^>]*?)\/>/g)].map((m) => ({ attrs: m[1], min: Number(attr(m[1], 'min')), max: Number(attr(m[1], 'max')) })) : [];
  const colPx = (col) => {
    const c = colEls.find((x) => col >= x.min && col <= x.max);
    return pxOf(c && attr(c.attrs, 'width') !== null ? Number(attr(c.attrs, 'width')) : defaultWidth);
  };
  const EMU = 9525;
  const leftOf = (col0) => { let x = 0; for (let c = 1; c <= col0; c++) x += colPx(c) * EMU; return x; }; // col0: 0부터 센 열 번호
  const toAnchor = (x) => {
    let col0 = 0; let rest = Math.max(0, x);
    while (col0 < 16383 && rest >= colPx(col0 + 1) * EMU) { rest -= colPx(col0 + 1) * EMU; col0 += 1; }
    return { col0, off: Math.round(rest) };
  };
  const spanLeft = leftOf(firstCol - 1);
  const spanRight = leftOf(lastCol);
  let count = 0;
  for (const path of await readSheetDrawingPaths(zip, sheetPath)) {
    const drawingFile = zip.file(path);
    if (!drawingFile) continue;
    const drawing = await drawingFile.async('string');
    const next = drawing.replace(/<xdr:(twoCellAnchor|oneCellAnchor)\b([^>]*)>([\s\S]*?)<\/xdr:\1>/g, (whole, kind, attrs, body) => {
      const read = (tag) => {
        const m = new RegExp(`<xdr:${tag}>[\\s\\S]*?<xdr:col>(\\d+)</xdr:col>\\s*<xdr:colOff>(-?\\d+)</xdr:colOff>`).exec(body);
        return m ? leftOf(Number(m[1])) + Number(m[2]) : null;
      };
      const from = read('from');
      if (from === null) return whole;
      const to = kind === 'twoCellAnchor' ? read('to') : null;
      const ext = /<xdr:spPr>[\s\S]*?<a:ext\b[^>]*\bcx="(\d+)"/.exec(body);
      const width = ext ? Number(ext[1]) : (to !== null ? to - from : 0);
      if (!(width > 0) || width >= spanRight - spanLeft) return whole;
      const left = Math.round(spanLeft + (spanRight - spanLeft - width) / 2);
      const put = (text, tag, x) => text.replace(new RegExp(`(<xdr:${tag}>[\\s\\S]*?)<xdr:col>\\d+</xdr:col>(\\s*)<xdr:colOff>-?\\d+</xdr:colOff>`), (w, head, gap) => {
        const a = toAnchor(x);
        return `${head}<xdr:col>${a.col0}</xdr:col>${gap}<xdr:colOff>${a.off}</xdr:colOff>`;
      });
      let nextBody = put(body, 'from', left);
      if (to !== null) nextBody = put(nextBody, 'to', left + width);
      nextBody = nextBody.replace(/(<xdr:spPr>[\s\S]*?<a:off\b[^>]*\bx=")(\d+)(")/, (w, a, x, b) => `${a}${left}${b}`);
      count += 1;
      return `<xdr:${kind}${attrs}>${nextBody}</xdr:${kind}>`;
    });
    if (next !== drawing) zip.file(path, next);
  }
  return count;
}
