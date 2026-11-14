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
    const fillsMatch = /<fills\b[^>]*>([\s\S]*?)<\/fills>/.exec(xml);
    this.fills = fillsMatch ? [...fillsMatch[1].matchAll(/<fill\b[^>]*?(?:\/>|>[\s\S]*?<\/fill>)/g)].map((m) => m[0]) : [];
    const fontsMatch = /<fonts\b[^>]*>([\s\S]*?)<\/fonts>/.exec(xml);
    this.fonts = fontsMatch ? [...fontsMatch[1].matchAll(/<font\b[^>]*?(?:\/>|>[\s\S]*?<\/font>)/g)].map((m) => m[0]) : [];
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
const cellKind = (attrs, inner) => {
  if (!inner) return 'empty';
  if (/<f\b/.test(inner)) return 'formula';
  const t = attr(attrs, 't');
  if (t === 's' || t === 'inlineStr' || t === 'str') return /<(?:v|t)\b[^>]*>[^<]/.test(inner) ? 'text' : 'empty';
  return /<v>[^<]/.test(inner) ? 'number' : 'empty';
};

const SUMMARY_LABEL = /^(평균|합계|계|avg\.?|average|mean|total)$/i;
export const normalizeTableBorders = (xml, styles, sharedStrings = []) => {
  const cells = new Map(); // "r,c" → s
  const kinds = new Map(); // "r,c" → 'text' | 'number' | 'formula' | 'empty'
  const texts = new Map(); // "r,c" → 글자 (평균 줄 찾기용)
  for (const m of xml.matchAll(/<c\b([^>]*?)(?:\/>|>([\s\S]*?)<\/c>)/g)) {
    const ref = parseRef(attr(m[1], 'r') || '');
    if (!ref) continue;
    cells.set(`${ref.row},${ref.col}`, Number(attr(m[1], 's') || 0));
    kinds.set(`${ref.row},${ref.col}`, cellKind(m[1], m[2]));
    const inner = m[2] || '';
    if (attr(m[1], 't') === 's') { const v = /<v>(\d+)<\/v>/.exec(inner); if (v) texts.set(`${ref.row},${ref.col}`, sharedStrings[Number(v[1])] || ''); }
    else if (attr(m[1], 't') === 'inlineStr') texts.set(`${ref.row},${ref.col}`, [...inner.matchAll(/<t\b[^>]*>([^<]*)<\/t>/g)].map((x) => decodeXml(x[1])).join(''));
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
    // 머리줄: 숫자·수식이 없고, 칸마다 채우기나 굵은 글씨가 있는 줄 (병합 칸은 왼쪽 위 칸으로 본다)
    // 머리줄 바로 아래(첫 자료줄 위) 선은 강조선으로 굵게 둔다
    const topLeft = (key) => {
      if (!mergeOf.has(key)) return key;
      const r = merges[mergeOf.get(key)];
      return K(r.from.row, r.from.col);
    };
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
        const want = !comp.has(next) || (edgeKey && headerEdges.has(edgeKey)) ? 'medium' : 'thin';
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
  if (/<sheetData\s*\/>/.test(out)) return out.replace(/<sheetData\s*\/>/, `<sheetData>${tag}</sheetData>`);
  return out.replace('</sheetData>', `${tag}</sheetData>`);
};

export const normalizeSectionSpacing = (xml, styles, { sharedStrings = [], drawingRows = new Set(), gap = 18, titleHeight = 21, afterTitle = 4, cols = null } = {}) => {
  const inCols = (col) => !cols || (col >= cols.from && col <= cols.to);
  const rows = new Map(); // row → { value, side, title, styles: [[ref, s]] }
  const info = (row) => {
    if (!rows.has(row)) rows.set(row, { value: false, side: false, title: false, cells: [] });
    return rows.get(row);
  };
  for (const m of xml.matchAll(/<c\b([^>]*?)(?:\/>|>([\s\S]*?)<\/c>)/g)) {
    const ref = parseRef(attr(m[1], 'r') || '');
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
  const contentRows = [...new Set([...rows.keys(), ...drawingRows])].filter(content).sort((a, b) => a - b);
  if (contentRows.length < 2) return xml;
  const heights = new Map(); // row → ht
  const bottomRows = new Set();
  for (let i = 1; i < contentRows.length; i++) {
    const prev = contentRows[i - 1]; const next = contentRows[i];
    const blanks = [];
    for (let row = prev + 1; row < next; row++) if (!drawingRows.has(row)) blanks.push(row);
    if (blanks.length !== next - prev - 1) continue; // 그림이 걸친 구간은 그대로
    if (isTitle(next) && !isTitle(prev)) {
      if (blanks.length) blanks.forEach((row) => heights.set(row, Math.round((gap / blanks.length) * 100) / 100));
      else { heights.set(next, titleHeight + gap); bottomRows.add(next); }
    } else if (isTitle(prev) && !isTitle(next)) {
      blanks.forEach((row) => heights.set(row, afterTitle));
    } else if (isTable(prev) && isTable(next) && blanks.length) {
      blanks.forEach((row) => heights.set(row, Math.round((gap / blanks.length) * 100) / 100));
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
const readDrawingRows = async (zip, sheetPath) => {
  const rowsSet = new Set();
  const dir = sheetPath.replace(/[^/]+$/, '');
  const relsFile = zip.file(`${dir}_rels/${sheetPath.split('/').pop()}.rels`);
  if (!relsFile) return rowsSet;
  const rels = await relsFile.async('string');
  for (const m of rels.matchAll(/<Relationship\b([^>]*)\/?>/g)) {
    if (!/\/drawing"/.test(attr(m[1], 'Type') || '')) continue;
    const target = attr(m[1], 'Target') || '';
    const parts = (dir + target).split('/');
    const stack = [];
    for (const part of parts) { if (part === '..') stack.pop(); else if (part && part !== '.') stack.push(part); }
    const file = zip.file(stack.join('/'));
    if (!file) continue;
    const drawing = await file.async('string');
    for (const a of drawing.matchAll(/<xdr:(twoCellAnchor|oneCellAnchor)\b[\s\S]*?<\/xdr:\1>/g)) {
      const from = /<xdr:from>[\s\S]*?<xdr:row>(\d+)<\/xdr:row>/.exec(a[0]);
      const to = /<xdr:to>[\s\S]*?<xdr:row>(\d+)<\/xdr:row>/.exec(a[0]);
      if (!from) continue;
      const start = Number(from[1]) + 1; const end = to ? Number(to[1]) + 1 : start;
      for (let row = start; row <= end; row++) rowsSet.add(row);
    }
  }
  return rowsSet;
};

// ── 인쇄 크기 통일: 한 파일의 모든 시트가 같은 폭으로 인쇄되게, 시트마다 배율을 정한다 ──
// 공통 폭 = 용지 인쇄 가능 폭과, 시트마다 '세로가 한 장에 들어가는 폭' 중 가장 작은 값
const A4 = { width: 595.3, height: 841.9 };
const measurePrintArea = (xml, range, mdw, styles) => {
  const fmt = /<sheetFormatPr\b([^>]*)\/?>/.exec(xml);
  const defaultRow = Number((fmt && attr(fmt[1], 'defaultRowHeight')) || 15);
  const baseCol = Number((fmt && attr(fmt[1], 'baseColWidth')) || 8);
  const defaultCol = Number((fmt && attr(fmt[1], 'defaultColWidth')) || baseCol + 0.71);
  const colPt = (chars) => Math.trunc(chars * mdw + 5) * 0.75;
  const cols = [];
  for (const m of xml.matchAll(/<col\b([^>]*)\/?>/g)) {
    cols.push({ min: Number(attr(m[1], 'min')), max: Number(attr(m[1], 'max')), width: Number(attr(m[1], 'width') || defaultCol), hidden: attr(m[1], 'hidden') === '1' });
  }
  let r = range;
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
    rowHeights.set(row, attr(m[1], 'hidden') === '1' ? 0 : Number(attr(m[1], 'ht') || defaultRow));
  }
  let height = 0;
  for (let row = r.from.row; row <= r.to.row; row++) height += rowHeights.has(row) ? rowHeights.get(row) : defaultRow;
  return { width, height };
};
const setPrintScale = (xml, scale, margins) => {
  let out = xml;
  // '한 장에 맞춤'을 끄고 배율로 인쇄
  out = out.replace(/<pageSetUpPr\b([^>]*)\/>/, (w, a) => `<pageSetUpPr${a.replace(/\s*fitToPage="[^"]*"/, '')}/>`);
  // 양식의 수동 페이지 나누기는 지운다 (한 시트 = 한 장)
  out = out.replace(/<rowBreaks\b[^>]*(?:\/>|>[\s\S]*?<\/rowBreaks>)/, '').replace(/<colBreaks\b[^>]*(?:\/>|>[\s\S]*?<\/colBreaks>)/, '');
  const page = `paperSize="9" scale="${scale}" orientation="portrait"`;
  if (/<pageSetup\b[^>]*\/>/.test(out)) {
    out = out.replace(/<pageSetup\b([^>]*)\/>/, (w, a) => `<pageSetup ${page}${a.replace(/\s*(?:paperSize|scale|orientation|fitToWidth|fitToHeight)="[^"]*"/g, '')}/>`);
  }
  const m = margins;
  const marginTag = `<pageMargins left="${m.left}" right="${m.right}" top="${m.top}" bottom="${m.bottom}" header="${m.header}" footer="${m.footer}"/>`;
  out = /<pageMargins\b[^>]*\/>/.test(out) ? out.replace(/<pageMargins\b[^>]*\/>/, marginTag) : out;
  // 가로 가운데
  if (/<printOptions\b[^>]*\/>/.test(out)) out = out.replace(/<printOptions\b([^>]*)\/>/, (w, a) => `<printOptions${a.replace(/\s*horizontalCentered="[^"]*"/, '')} horizontalCentered="1"/>`);
  else out = out.replace(/<pageMargins\b/, '<printOptions horizontalCentered="1"/><pageMargins');
  return out;
};

export async function tidyWorkbook(zip, { rules = {}, onePageSheets = [], equalWidth = false } = {}) {
  const read = async (path) => { const f = zip.file(path); return f ? f.async('string') : null; };
  const workbookXml = await read('xl/workbook.xml');
  const relsXml = await read('xl/_rels/workbook.xml.rels');
  const stylesXml = await read('xl/styles.xml');
  if (!workbookXml || !relsXml || !stylesXml) return;
  const rels = {};
  for (const m of relsXml.matchAll(/<Relationship\b([^>]*)\/?>/g)) rels[attr(m[1], 'Id')] = attr(m[1], 'Target');
  const styles = new StyleBook(stylesXml);
  styles.cleanAll();
  const sharedXml = await read('xl/sharedStrings.xml');
  const sharedStrings = sharedXml
    ? [...sharedXml.matchAll(/<si>([\s\S]*?)<\/si>/g)].map((m) => [...m[1].matchAll(/<t\b[^>]*>([^<]*)<\/t>/g)].map((t) => decodeXml(t[1])).join(''))
    : [];
  // 시트별 인쇄 범위의 열 (간격은 인쇄되는 부분만 본다)
  const printCols = new Map();
  const printRanges = new Map();
  for (const m of workbookXml.matchAll(/<definedName\b([^>]*)>([^<]*)<\/definedName>/g)) {
    if (attr(m[1], 'name') !== '_xlnm.Print_Area' || attr(m[1], 'localSheetId') === null) continue;
    const range = /!\$?([A-Z]{1,3})\$?(\d+):\$?([A-Z]{1,3})\$?(\d+)/.exec(decodeXml(m[2]));
    if (!range) continue;
    printCols.set(Number(attr(m[1], 'localSheetId')), { from: colToNum(range[1]), to: colToNum(range[3]) });
    printRanges.set(Number(attr(m[1], 'localSheetId')), { from: { col: colToNum(range[1]), row: Number(range[2]) }, to: { col: colToNum(range[3]), row: Number(range[4]) } });
  }
  const sheetsForWidth = []; // 인쇄 크기 통일에 쓸 시트들
  let sheetIndex = -1;
  const printAreaUpdates = new Map(); // 시트 순서 → { name, range }
  for (const m of workbookXml.matchAll(/<sheet\b([^>]*)\/>/g)) {
    sheetIndex += 1;
    const name = decodeXml(attr(m[1], 'name') || '');
    if (attr(m[1], 'state')) continue; // 숨긴 시트는 그대로
    const target = rels[attr(m[1], 'r:id')];
    if (!target) continue;
    const path = target.startsWith('/') ? target.slice(1) : `xl/${target}`;
    let xml = await read(path);
    if (!xml) continue;
    xml = fitToPageWidth(xml, { onePage: onePageSheets === 'all' || onePageSheets.includes(name) });
    const rule = typeof rules === 'function' ? rules(name) : rules[name];
    if (rule?.printArea) printAreaUpdates.set(sheetIndex, { name, range: rule.printArea });
    if (rule?.rowHeights) xml = setRowHeights(xml, rule.rowHeights);
    if (rule?.numberFormats) xml = setNumberFormats(xml, styles, rule.numberFormats);
    if (rule?.center) xml = setCentered(xml, styles, rule.center);
    if (rule?.borders !== false) xml = normalizeTableBorders(xml, styles, sharedStrings);
    if (equalWidth) sheetsForWidth.push({ path, index: sheetIndex, name });
    if (rule?.spacing !== false) xml = normalizeSectionSpacing(xml, styles, { sharedStrings, drawingRows: await readDrawingRows(zip, path), cols: printCols.get(sheetIndex) || null });
    zip.file(path, xml);
  }
  zip.file('xl/styles.xml', styles.serialize());
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
      const size = xml ? measurePrintArea(xml, range, mdw, styles) : null;
      if (size && size.width > 0 && size.height > 0) measured.push({ ...item, ...size });
    }
    // 가로는 글꼴 폭 어림 오차로 3%, 세로는 머리글·바닥글 공간으로 5% 여유
    const common = Math.min(printableW * 0.97, ...measured.map((x) => ((printableH * 0.95) / x.height) * x.width));
    for (const item of measured) {
      const scale = Math.max(10, Math.min(400, Math.floor((common / item.width) * 100)));
      const xml = await read(item.path);
      zip.file(item.path, setPrintScale(xml, scale, margins));
    }
  }
  // 인쇄 범위 바꾸기 (프로그램이 표를 넓혀 그린 시트)
  if (printAreaUpdates.size) {
    let wb = workbookXml;
    for (const [index, { name, range }] of printAreaUpdates) {
      const [a, b] = range.split(':');
      const abs = (ref) => ref.replace(/^([A-Z]+)(\d+)$/, '$$$1$$$2');
      const text = `'${name.replace(/'/g, "''")}'!${abs(a)}:${abs(b || a)}`.replace(/&/g, '&amp;').replace(/</g, '&lt;');
      const re = new RegExp(`<definedName\\b(?=[^>]*name="_xlnm\\.Print_Area")(?=[^>]*localSheetId="${index}")[^>]*>[^<]*</definedName>`);
      const tag = `<definedName name="_xlnm.Print_Area" localSheetId="${index}">${text}</definedName>`;
      if (re.test(wb)) wb = wb.replace(re, tag);
      else if (/<definedNames\b[^>]*>/.test(wb)) wb = wb.replace(/<definedNames\b([^>]*)>/, `<definedNames$1>${tag}`);
      else wb = wb.replace(/(<\/sheets>)/, `$1<definedNames>${tag}</definedNames>`);
    }
    zip.file('xl/workbook.xml', wb);
  }
}

// 행 범위 [시작, 끝]을 같은 높이로
export const rowsAt = (from, to, ht) => Object.fromEntries(Array.from({ length: to - from + 1 }, (_, i) => [from + i, ht]));
