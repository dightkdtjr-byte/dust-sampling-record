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
    let out = this.xml;
    if (this.bordersChanged) out = out.replace(/<borders\b[^>]*>[\s\S]*?<\/borders>/, `<borders count="${this.borders.length}">${this.borders.join('')}</borders>`);
    if (this.added.size === 0) return out;
    return out.replace(/<cellXfs\b[^>]*>[\s\S]*?<\/cellXfs>/, `<cellXfs count="${this.xfs.length}">${this.xfs.join('')}</cellXfs>`);
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

// 표 테두리 맞추기
//  - 표 = 세 변 이상이 선으로 둘러싸인 칸이 상하좌우로 이어진 덩어리 (병합 칸은 통째로, 표 안에 갇힌 빈칸 포함).
//    안쪽에 세로선과 가로선이 모두 있어야 표로 본다 (밑줄만 있는 서식 칸은 표가 아니다)
//  - 표 바깥과 맞닿은 변은 medium, 표 안쪽 변은 thin. 이중선(double)은 일부러 쓴 것이라 그대로 둔다
const RANK_KEEP = new Set(['double']);
export const normalizeTableBorders = (xml, styles) => {
  const cells = new Map(); // "r,c" → s
  for (const m of xml.matchAll(/<c\b([^>]*?)(?:\/>|>)/g)) {
    const ref = parseRef(attr(m[1], 'r') || '');
    if (ref) cells.set(`${ref.row},${ref.col}`, Number(attr(m[1], 's') || 0));
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

  const DIRS = { left: [0, -1], right: [0, 1], top: [-1, 0], bottom: [1, 0] };
  const OPP = { left: 'right', right: 'left', top: 'bottom', bottom: 'top' };
  const seen = new Set();
  const changes = new Map(); // "r,c" → { side: style }
  const setChange = (key, side, style) => {
    const current = own(key)[side];
    if (!current || current === style || RANK_KEEP.has(current)) return;
    changes.set(key, { ...(changes.get(key) || {}), [side]: style });
  };
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
    for (const key of comp) {
      const [row, col] = key.split(',').map(Number);
      for (const [side, [dr, dc]] of Object.entries(DIRS)) {
        const next = K(row + dr, col + dc);
        if (sameMerge(key, next)) continue; // 병합 칸 안쪽 변은 그려지지 않는다
        const line = side === 'left' ? vLine(row, col) : side === 'right' ? vLine(row, col + 1) : side === 'top' ? hLine(row, col) : hLine(row + 1, col);
        if (!line) continue;
        const want = comp.has(next) ? 'thin' : 'medium';
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
    if (rule?.borders !== false) xml = normalizeTableBorders(xml, styles);
    zip.file(path, xml);
  }
  zip.file('xl/styles.xml', styles.serialize());
}

// 행 범위 [시작, 끝]을 같은 높이로
export const rowsAt = (from, to, ht) => Object.fromEntries(Array.from({ length: to - from + 1 }, (_, i) => [from + i, ht]));
