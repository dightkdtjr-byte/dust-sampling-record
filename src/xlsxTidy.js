// 엑셀 서식 정리: 내보내기 마지막에 모든 시트에 적용한다.
//  1) 숫자 표시 형식의 꼬리 공백("0.00_ ")·음수 빨간 괄호를 없앤다 → 가운데 정렬한 숫자가 칸마다 어긋나지 않게
//  2) 보이는 시트는 한 장 너비(onePageSheets면 한 장)에 맞춰 인쇄 (오른쪽이 잘려 다음 장으로 넘어가지 않게)
//  3) 시트별 규칙: 같은 표의 행 높이 맞추기, 같은 열의 숫자 자릿수 맞추기
const attr = (attrs, name) => {
  const m = new RegExp(`\\b${name}="([^"]*)"`).exec(attrs);
  return m ? m[1] : null;
};
const decodeXml = (text) => String(text).replace(/&quot;/g, '"').replace(/&lt;/g, '<').replace(/&gt;/g, '>').replace(/&amp;/g, '&');
const encodeAttr = (text) => String(text).replace(/&/g, '&amp;').replace(/"/g, '&quot;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
const colToNum = (letters) => [...letters].reduce((n, ch) => n * 26 + ch.charCodeAt(0) - 64, 0);
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
      this.xml = this.xml.replace(/(<styleSheet\b[^>]*>)/, `$1<numFmts count="1">${tag}</numFmts>`);
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
    if (this.added.size === 0) return this.xml;
    return this.xml.replace(/<cellXfs\b[^>]*>[\s\S]*?<\/cellXfs>/, `<cellXfs count="${this.xfs.length}">${this.xfs.join('')}</cellXfs>`);
  }
}

// 한 장 너비에 맞춤: <sheetPr><pageSetUpPr fitToPage="1"/></sheetPr> + <pageSetup fitToWidth="1" fitToHeight="0">
const fitToPageWidth = (xml, { onePage = false } = {}) => {
  let out = xml;
  if (/<sheetPr\b[^>]*\/>/.test(out)) out = out.replace(/<sheetPr\b([^>]*)\/>/, '<sheetPr$1><pageSetUpPr fitToPage="1"/></sheetPr>');
  else if (/<sheetPr\b[^>]*>/.test(out)) {
    if (/<pageSetUpPr\b[^>]*\/>/.test(out)) out = out.replace(/<pageSetUpPr\b([^>]*)\/>/, (w, a) => `<pageSetUpPr${a.replace(/\s*fitToPage="[^"]*"/, '')} fitToPage="1"/>`);
    else out = out.replace(/<\/sheetPr>/, '<pageSetUpPr fitToPage="1"/></sheetPr>');
  } else out = out.replace(/(<worksheet\b[^>]*>)/, '$1<sheetPr><pageSetUpPr fitToPage="1"/></sheetPr>');
  const fit = `fitToWidth="1" fitToHeight="${onePage ? 1 : 0}"`;
  if (/<pageSetup\b[^>]*\/>/.test(out)) {
    out = out.replace(/<pageSetup\b([^>]*)\/>/, (w, a) => `<pageSetup${a.replace(/\s*fitTo(?:Width|Height)="[^"]*"/g, '')} ${fit}/>`);
  } else if (/<pageMargins\b[^>]*\/>/.test(out)) {
    out = out.replace(/(<pageMargins\b[^>]*\/>)/, `$1<pageSetup paperSize="9" orientation="portrait" ${fit}/>`);
  }
  return out;
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

export async function tidyWorkbook(zip, { rules = {}, onePageSheets = [] } = {}) {
  const read = async (path) => { const f = zip.file(path); return f ? f.async('string') : null; };
  const workbookXml = await read('xl/workbook.xml');
  const relsXml = await read('xl/_rels/workbook.xml.rels');
  const stylesXml = await read('xl/styles.xml');
  if (!workbookXml || !relsXml || !stylesXml) return;
  const rels = {};
  for (const m of relsXml.matchAll(/<Relationship\b([^>]*)\/?>/g)) rels[attr(m[1], 'Id')] = attr(m[1], 'Target');
  const styles = new StyleBook(stylesXml);
  styles.cleanAll();
  for (const m of workbookXml.matchAll(/<sheet\b([^>]*)\/>/g)) {
    const name = decodeXml(attr(m[1], 'name') || '');
    if (attr(m[1], 'state')) continue; // 숨긴 시트는 그대로
    const target = rels[attr(m[1], 'r:id')];
    if (!target) continue;
    const path = target.startsWith('/') ? target.slice(1) : `xl/${target}`;
    let xml = await read(path);
    if (!xml) continue;
    xml = fitToPageWidth(xml, { onePage: onePageSheets === 'all' || onePageSheets.includes(name) });
    const rule = typeof rules === 'function' ? rules(name) : rules[name];
    if (rule?.rowHeights) xml = setRowHeights(xml, rule.rowHeights);
    if (rule?.numberFormats) xml = setNumberFormats(xml, styles, rule.numberFormats);
    if (rule?.center) xml = setCentered(xml, styles, rule.center);
    zip.file(path, xml);
  }
  zip.file('xl/styles.xml', styles.serialize());
}

// 행 범위 [시작, 끝]을 같은 높이로
export const rowsAt = (from, to, ht) => Object.fromEntries(Array.from({ length: to - from + 1 }, (_, i) => [from + i, ht]));
