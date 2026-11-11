// 공개 양식 엑셀 만들기: 회사 양식 파일 없이, 프로그램이 처음부터 직접 그리는 통합 문서(xlsx).
// 배포용 빌드의 엑셀 추출에 쓴다. 값은 화면에서 계산한 결과를 그대로 적는다 (화면 = 엑셀).
import JSZip from 'jszip';

const NS_MAIN = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';
const NS_REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';

const escapeXml = (value) => String(value)
  // XML에 쓸 수 없는 제어문자는 뺀다
  // eslint-disable-next-line no-control-regex
  .replace(/[\u0000-\u0008\u000B\u000C\u000E-\u001F￾￿]/g, '')
  .slice(0, 32767)
  .replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');

export const columnName = (index) => {
  let n = index + 1;
  let name = '';
  while (n > 0) { const rem = (n - 1) % 26; name = String.fromCharCode(65 + rem) + name; n = Math.floor((n - 1) / 26); }
  return name;
};

// ── 서식 ──
// 이름으로 고르는 셀 서식. 순서가 곧 styles.xml의 cellXfs 번호다.
const STYLE_DEFS = [
  { key: 'plain', font: 0, fill: 0, border: 0 },
  { key: 'title', font: 1, fill: 0, border: 0, h: 'center' },
  { key: 'subtitle', font: 2, fill: 0, border: 0, h: 'center' },
  { key: 'section', font: 3, fill: 2, border: 1, h: 'left' },
  { key: 'label', font: 4, fill: 3, border: 1, h: 'center', wrap: true },
  { key: 'text', font: 0, fill: 0, border: 1, h: 'center', wrap: true },
  { key: 'textLeft', font: 0, fill: 0, border: 1, h: 'left', wrap: true },
  { key: 'num', font: 0, fill: 0, border: 1, h: 'center' },
  { key: 'num1', font: 0, fill: 0, border: 1, h: 'center', numFmt: 164 },
  { key: 'num2', font: 0, fill: 0, border: 1, h: 'center', numFmt: 2 },
  { key: 'num3', font: 0, fill: 0, border: 1, h: 'center', numFmt: 165 },
  { key: 'num4', font: 0, fill: 0, border: 1, h: 'center', numFmt: 166 },
  { key: 'head', font: 4, fill: 3, border: 1, h: 'center', wrap: true },
  { key: 'result', font: 5, fill: 0, border: 1, h: 'center' },
  { key: 'result2', font: 5, fill: 0, border: 1, h: 'center', numFmt: 2 },
  { key: 'note', font: 2, fill: 0, border: 0, h: 'left', wrap: true },
  { key: 'formula', font: 6, fill: 0, border: 1, h: 'left', wrap: true },
  { key: 'calc', font: 0, fill: 4, border: 1, h: 'center' },
  { key: 'calc2', font: 0, fill: 4, border: 1, h: 'center', numFmt: 2 },
  // 공정시험기준 기록지 양식용: 이름표(테두리 없음), 밑줄 칸, 무늬 없는 표 머리, 그림 칸
  { key: 'formTitle', font: 8, fill: 0, border: 0, h: 'left' },
  { key: 'formLabel', font: 7, fill: 0, border: 0, h: 'left' },
  { key: 'formValue', font: 7, fill: 0, border: 2, h: 'center' },
  { key: 'formHead', font: 9, fill: 0, border: 1, h: 'center', wrap: true },
  { key: 'formBox', font: 9, fill: 0, border: 1, h: 'center', wrap: true },
  { key: 'formCaption', font: 9, fill: 0, border: 0, h: 'center', wrap: true },
  { key: 'formCell', font: 7, fill: 0, border: 1, h: 'center' },
  { key: 'formCell4', font: 7, fill: 0, border: 1, h: 'center', numFmt: 166 },
  { key: 'formNote', font: 9, fill: 0, border: 0, h: 'left', wrap: true },
];
const STYLE_INDEX = Object.fromEntries(STYLE_DEFS.map((def, index) => [def.key, index]));

const FONT_NAME = '맑은 고딕';
const stylesXml = () => {
  const font = (size, { bold = false, color = null } = {}) => (
    `<font>${bold ? '<b/>' : ''}<sz val="${size}"/>${color ? `<color rgb="${color}"/>` : ''}<name val="${FONT_NAME}"/><family val="2"/></font>`
  );
  const formFont = (size, bold = false) => `<font>${bold ? '<b/>' : ''}<sz val="${size}"/><name val="바탕"/><family val="1"/></font>`;
  const fonts = [
    font(10), font(16, { bold: true }), font(9, { color: 'FF5F6368' }), font(11, { bold: true }),
    font(9, { bold: true }), font(11, { bold: true }), font(9),
    // 공정시험기준 기록지 양식: 원문과 같은 바탕(명조) 글꼴
    formFont(10), formFont(11, true), formFont(9),
  ];
  const solid = (rgb) => `<fill><patternFill patternType="solid"><fgColor rgb="${rgb}"/><bgColor indexed="64"/></patternFill></fill>`;
  const fills = [
    '<fill><patternFill patternType="none"/></fill>', '<fill><patternFill patternType="gray125"/></fill>',
    solid('FFDDE5E0'), solid('FFF1F4F2'), solid('FFF7F7F2'),
  ];
  const side = (name) => `<${name} style="thin"><color rgb="FF8A8F8C"/></${name}>`;
  const borders = [
    '<border><left/><right/><top/><bottom/><diagonal/></border>',
    `<border>${side('left')}${side('right')}${side('top')}${side('bottom')}<diagonal/></border>`,
    `<border><left/><right/><top/>${side('bottom')}<diagonal/></border>`,
  ];
  const xfs = STYLE_DEFS.map((def) => (
    `<xf numFmtId="${def.numFmt || 0}" fontId="${def.font}" fillId="${def.fill}" borderId="${def.border}" xfId="0"`
    + `${def.numFmt ? ' applyNumberFormat="1"' : ''} applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1">`
    + `<alignment horizontal="${def.h || 'general'}" vertical="center"${def.wrap ? ' wrapText="1"' : ''}/></xf>`
  ));
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="${NS_MAIN}">
<numFmts count="3"><numFmt numFmtId="164" formatCode="0.0"/><numFmt numFmtId="165" formatCode="0.000"/><numFmt numFmtId="166" formatCode="0.0000"/></numFmts>
<fonts count="${fonts.length}">${fonts.join('')}</fonts>
<fills count="${fills.length}">${fills.join('')}</fills>
<borders count="${borders.length}">${borders.join('')}</borders>
<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>
<cellXfs count="${xfs.length}">${xfs.join('')}</cellXfs>
<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>
</styleSheet>`;
};

// ── 시트 그리기 도구 ──
// 한 칸씩 "어느 줄, 어느 열부터 몇 칸"으로 적어 나가면 병합·테두리를 알아서 채운다.
export const createSheet = (name, { columns, landscape = false } = {}) => {
  const rows = new Map(); // 줄 번호(1부터) → { height, cells: Map(열 번호 → { v, s }) }
  const merges = [];
  const rowOf = (r) => { if (!rows.has(r)) rows.set(r, { height: null, cells: new Map() }); return rows.get(r); };
  const sheet = {
    name, columns, landscape,
    // (r, c)에서 가로 span칸·세로 rowSpan줄을 차지하는 칸. value가 숫자면 숫자, { f }면 수식, 그 밖엔 글자
    put(r, c, value, style = 'text', span = 1, rowSpan = 1) {
      for (let dr = 0; dr < rowSpan; dr++) {
        for (let dc = 0; dc < span; dc++) rowOf(r + dr).cells.set(c + dc, { v: dr === 0 && dc === 0 ? value : null, s: style });
      }
      if (span > 1 || rowSpan > 1) merges.push(`${columnName(c)}${r}:${columnName(c + span - 1)}${r + rowSpan - 1}`);
      return sheet;
    },
    height(r, points) { rowOf(r).height = points; return sheet; },
    // 그림 한 장: png(Uint8Array)를 (fromCol, fromRow)~(toCol, toRow) 칸 범위에 맞춰 넣는다 (0부터 세는 칸 번호, to는 끝 다음 칸)
    setImage(png, fromCol, fromRow, toCol, toRow) { sheet.image = { png, fromCol, fromRow, toCol, toRow }; return sheet; },
    toXml() {
      const cols = (columns || []).map((width, i) => `<col min="${i + 1}" max="${i + 1}" width="${width}" customWidth="1"/>`).join('');
      const rowXml = [...rows.keys()].sort((a, b) => a - b).map((r) => {
        const row = rows.get(r);
        const cells = [...row.cells.keys()].sort((a, b) => a - b).map((c) => {
          const { v, s } = row.cells.get(c);
          const ref = `${columnName(c)}${r}`;
          const sAttr = ` s="${STYLE_INDEX[s] ?? 0}"`;
          if (v === null || v === undefined || v === '') return `<c r="${ref}"${sAttr}/>`;
          if (typeof v === 'number' && Number.isFinite(v)) return `<c r="${ref}"${sAttr}><v>${v}</v></c>`;
          if (typeof v === 'object' && typeof v.f === 'string') return `<c r="${ref}"${sAttr}><f>${escapeXml(v.f)}</f></c>`;
          return `<c r="${ref}"${sAttr} t="inlineStr"><is><t xml:space="preserve">${escapeXml(v)}</t></is></c>`;
        }).join('');
        return `<row r="${r}"${row.height ? ` ht="${row.height}" customHeight="1"` : ''}>${cells}</row>`;
      }).join('');
      const mergeXml = merges.length ? `<mergeCells count="${merges.length}">${merges.map((ref) => `<mergeCell ref="${ref}"/>`).join('')}</mergeCells>` : '';
      return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="${NS_MAIN}" xmlns:r="${NS_REL}">
<sheetPr><pageSetUpPr fitToPage="1"/></sheetPr>
<sheetViews><sheetView workbookViewId="0" showGridLines="0"/></sheetViews>
<sheetFormatPr defaultRowHeight="18"/>
${cols ? `<cols>${cols}</cols>` : ''}
<sheetData>${rowXml}</sheetData>
${mergeXml}
<printOptions horizontalCentered="1"/>
<pageMargins left="0.4" right="0.4" top="0.5" bottom="0.5" header="0.2" footer="0.2"/>
<pageSetup paperSize="9" orientation="${landscape ? 'landscape' : 'portrait'}" fitToWidth="1" fitToHeight="0"/>
<headerFooter><oddFooter>&amp;C&amp;P / &amp;N</oddFooter></headerFooter>
${sheet.image ? '<drawing r:id="rId1"/>' : ''}
</worksheet>`;
    },
  };
  return sheet;
};

// 시트 이름에 쓸 수 없는 글자를 빼고 31자로 자른다
const safeSheetName = (name) => String(name || 'Sheet').replace(/[\\/?*[\]:]/g, ' ').slice(0, 31) || 'Sheet';

export const buildWorkbook = async (sheets) => {
  const zip = new JSZip();
  const names = sheets.map((sheet) => safeSheetName(sheet.name));
  zip.file('[Content_Types].xml', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
<Default Extension="xml" ContentType="application/xml"/>
<Default Extension="png" ContentType="image/png"/>
${sheets.map((sheet, i) => (sheet.image ? `<Override PartName="/xl/drawings/drawing${i + 1}.xml" ContentType="application/vnd.openxmlformats-officedocument.drawing+xml"/>` : '')).join('')}
<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>
${sheets.map((_, i) => `<Override PartName="/xl/worksheets/sheet${i + 1}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>`).join('\n')}
</Types>`);
  zip.file('_rels/.rels', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>`);
  zip.file('xl/workbook.xml', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="${NS_MAIN}" xmlns:r="${NS_REL}">
<sheets>${names.map((name, i) => `<sheet name="${escapeXml(name)}" sheetId="${i + 1}" r:id="rId${i + 1}"/>`).join('')}</sheets>
</workbook>`);
  zip.file('xl/_rels/workbook.xml.rels', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
${sheets.map((_, i) => `<Relationship Id="rId${i + 1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet${i + 1}.xml"/>`).join('\n')}
<Relationship Id="rId${sheets.length + 1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
</Relationships>`);
  zip.file('xl/styles.xml', stylesXml());
  sheets.forEach((sheet, i) => {
    zip.file(`xl/worksheets/sheet${i + 1}.xml`, sheet.toXml());
    if (!sheet.image) return;
    const n = i + 1;
    const { png, fromCol, fromRow, toCol, toRow } = sheet.image;
    const pad = 38100; // 칸 테두리에 붙지 않게 4px쯤 안쪽으로
    zip.file(`xl/media/image${n}.png`, png);
    zip.file(`xl/worksheets/_rels/sheet${n}.xml.rels`, `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing" Target="../drawings/drawing${n}.xml"/></Relationships>`);
    zip.file(`xl/drawings/_rels/drawing${n}.xml.rels`, `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/image${n}.png"/></Relationships>`);
    zip.file(`xl/drawings/drawing${n}.xml`, `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="${NS_REL}">
<xdr:twoCellAnchor editAs="twoCell">
<xdr:from><xdr:col>${fromCol}</xdr:col><xdr:colOff>${pad}</xdr:colOff><xdr:row>${fromRow}</xdr:row><xdr:rowOff>${pad}</xdr:rowOff></xdr:from>
<xdr:to><xdr:col>${toCol - 1}</xdr:col><xdr:colOff>${600000}</xdr:colOff><xdr:row>${toRow - 1}</xdr:row><xdr:rowOff>${228600}</xdr:rowOff></xdr:to>
<xdr:pic><xdr:nvPicPr><xdr:cNvPr id="2" name="측정점 배열"/><xdr:cNvPicPr><a:picLocks noChangeAspect="1"/></xdr:cNvPicPr></xdr:nvPicPr>
<xdr:blipFill><a:blip r:embed="rId1"/><a:stretch><a:fillRect/></a:stretch></xdr:blipFill>
<xdr:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></xdr:spPr></xdr:pic>
<xdr:clientData/></xdr:twoCellAnchor></xdr:wsDr>`);
  });
  return zip.generateAsync({ type: 'arraybuffer', compression: 'DEFLATE' });
};

// ── 굴뚝단면 및 측정점 배열 그림 ──
// 원형 굴뚝의 반경 구분 수: 화면의 측정점 계산과 같은 기준 (단면적 0.25㎡ 이하는 중심 1점)
export const getStackPointCoeffs = (diameter) => {
  const D = Number(diameter);
  if (!Number.isFinite(D) || D <= 0) return null;
  if (Math.PI * D * D / 4 <= 0.25) return [];
  if (D <= 1) return [0.707];
  if (D <= 2) return [0.5, 0.866];
  if (D <= 4) return [0.408, 0.707, 0.913];
  if (D <= 4.5) return [0.354, 0.612, 0.791, 0.935];
  return [0.316, 0.548, 0.707, 0.837, 0.949];
};
// 흑백 선 그림(SVG 글자열): 굴뚝 단면, 등면적 원환(점선), 측정공, 지름 위 측정점과 번호, 안지름
export const buildStackSectionSvg = (diameter, width = 520, height = 460) => {
  const coeffsRaw = getStackPointCoeffs(diameter);
  if (coeffsRaw === null) return '';
  const cx = width / 2 + 30; const cy = height / 2 - 34; const R = Math.min(width, height) * 0.37;
  const coeffs = [...coeffsRaw].sort((a, b) => b - a); // 측정공(왼쪽)에 가까운 점부터 1번
  // 측정점은 화면과 같이 반경당 점 수만큼만: 측정공 쪽 반경 위에 1번부터
  const pts = coeffs.length === 0 ? [{ x: cx, n: 1 }] : coeffs.map((c, i) => ({ x: cx - c * R, n: i + 1 }));
  const ink = '#000';
  const parts = [
    `<svg xmlns="http://www.w3.org/2000/svg" width="${width}" height="${height}" viewBox="0 0 ${width} ${height}" font-family="Batang, 바탕, 'Malgun Gothic', serif">`,
    `<rect width="${width}" height="${height}" fill="#fff"/>`,
    `<circle cx="${cx}" cy="${cy}" r="${R}" fill="none" stroke="${ink}" stroke-width="2.4"/>`,
    ...coeffs.map((c) => `<circle cx="${cx}" cy="${cy}" r="${(c * R).toFixed(1)}" fill="none" stroke="${ink}" stroke-width="1" stroke-dasharray="5 6"/>`),
    `<line x1="${cx - R}" y1="${cy}" x2="${cx + R}" y2="${cy}" stroke="${ink}" stroke-width="1.4"/>`,
    `<line x1="${cx}" y1="${cy - 9}" x2="${cx}" y2="${cy + 9}" stroke="${ink}" stroke-width="1.4"/>`,
    `<rect x="${cx - R - 36}" y="${cy - 10}" width="36" height="20" fill="#fff" stroke="${ink}" stroke-width="1.8"/>`,
    `<text x="${cx - R - 26}" y="${cy + 46}" font-size="27" text-anchor="middle" fill="${ink}">측정공</text>`,
    ...pts.map((pt, i) => `<circle cx="${pt.x.toFixed(1)}" cy="${cy}" r="7.5" fill="${ink}"/><text x="${pt.x.toFixed(1)}" y="${cy + (i % 2 === 0 ? -18 : 40)}" font-size="28" text-anchor="middle" fill="${ink}">${pt.n}</text>`),
    `<line x1="${cx - R}" y1="${cy + R + 34}" x2="${cx + R}" y2="${cy + R + 34}" stroke="${ink}" stroke-width="1.2"/>`,
    `<line x1="${cx - R}" y1="${cy + R + 24}" x2="${cx - R}" y2="${cy + R + 44}" stroke="${ink}" stroke-width="1.2"/>`,
    `<line x1="${cx + R}" y1="${cy + R + 24}" x2="${cx + R}" y2="${cy + R + 44}" stroke="${ink}" stroke-width="1.2"/>`,
    `<text x="${cx - 15}" y="${cy + R + 72}" font-size="28" text-anchor="middle" fill="${ink}">안지름 ${Number(diameter)} m · 측정점 ${pts.length}점</text>`,
    '</svg>',
  ];
  return parts.join('');
};
// SVG 글자열을 PNG(Uint8Array)로 그린다. 화면(브라우저)에서만 동작
export const renderSvgToPng = (svg, width, height) => new Promise((resolve, reject) => {
  const img = new Image();
  img.onload = () => {
    try {
      const canvas = document.createElement('canvas');
      canvas.width = width; canvas.height = height;
      canvas.getContext('2d').drawImage(img, 0, 0, width, height);
      canvas.toBlob((blob) => {
        if (!blob) { reject(new Error('png encode failed')); return; }
        blob.arrayBuffer().then((buffer) => resolve(new Uint8Array(buffer)), reject);
      }, 'image/png');
    } catch (error) { reject(error); }
  };
  img.onerror = () => reject(new Error('svg load failed'));
  img.src = `data:image/svg+xml;charset=utf-8,${encodeURIComponent(svg)}`;
});

// ── 값 다듬기 ──
const toNum = (value) => {
  if (value === null || value === undefined) return null;
  const text = String(value).trim();
  if (text === '' || text === '-') return null;
  const n = Number(text);
  return Number.isFinite(n) ? n : null;
};
// 숫자로 읽히면 숫자, 아니면 글자 그대로 (빈칸은 빈칸)
const numOrText = (value) => {
  const n = toNum(value);
  if (n !== null) return n;
  const text = value === null || value === undefined ? '' : String(value).trim();
  return text === '-' ? '' : text;
};
const average = (values) => {
  const list = values.map(toNum).filter((v) => v !== null);
  return list.length ? list.reduce((a, b) => a + b, 0) / list.length : null;
};
const round = (value, digits) => (value === null ? '' : Math.round(value * 10 ** digits) / 10 ** digits);

const COLS = 12;

// ── 기록부(먼지·금속·수은·PAHs·유량) 공개 양식 ──
// report: 저장된 기록부 한 건(화면에서 계산한 결과 포함), meta: { title, isFlow, moistureLabel, impingerLabels }
export const buildRecordSheets = (report, meta) => {
  const isFlow = Boolean(meta.isFlow);
  const s = createSheet('세부 기록', { columns: Array(COLS).fill(11.5), landscape: true });
  let r = 1;
  const section = (title) => { r += 1; s.put(r, 0, title, 'section', COLS).height(r, 21); r += 1; };
  // 한 줄에 "이름표 1칸 + 값 2칸" 네 쌍
  const pairs = (items) => {
    items.forEach(([label, value, style], i) => {
      s.put(r, i * 3, label, 'label');
      s.put(r, i * 3 + 1, value === undefined ? '' : value, style || 'text', 2);
    });
    for (let i = items.length; i < 4; i++) { s.put(r, i * 3, '', 'label'); s.put(r, i * 3 + 1, '', 'text', 2); }
    s.height(r, items.some(([label]) => String(label).length > 14) ? 38 : 26);
    r += 1;
  };
  // 표 한 줄: spans = 칸 너비 목록
  const tableRow = (spans, values, style) => {
    let c = 0;
    spans.forEach((span, i) => { s.put(r, c, values[i] === undefined ? '' : values[i], Array.isArray(style) ? style[i] : style, span); c += span; });
    r += 1;
  };

  s.put(r, 0, '배출가스 시료채취 세부 기록', 'title', COLS).height(r, 32); r += 1;
  s.put(r, 0, `${meta.title} · 「대기오염공정시험기준」 ES 01301.1e(배출가스 중 먼지 – 반자동식 측정법) 표 3 "먼지시료채취 기록지"와 ES 01114.c(배출가스 유속 및 유량 측정방법)의 기록 항목을 따름`, 'subtitle', COLS); r += 1;

  section('1. 일반 사항 및 배출구 제원');
  pairs([['측정 일자', report.date || ''], ['사업장명', report.company || ''], ['배출구명', report.location || ''], ['측정자', report.sampler || '']]);
  pairs([['날씨', report.weather || ''], ['대기 온도 (℃)', numOrText(report.atmTemp)], ['대기압 (mmHg)', numOrText(report.atmPressure)], ['표준산소농도 Os (%)', numOrText(report.standardO2) === '' ? '해당 없음' : numOrText(report.standardO2)]]);
  const diameter = toNum(report.stackDiameter);
  pairs([['굴뚝 내경 D (m)', numOrText(report.stackDiameter)], ['단면적 A (㎡)', diameter === null ? '' : round(Math.PI * diameter * diameter / 4, 4), 'num4'], ['총 측정 깊이 (m)', numOrText(report.totalStackDepth)], ['플랜지 길이 (m)', numOrText(report.flangeLength)]]);

  section('2. 사용 장비');
  pairs([['피토관', report.pitotInfo || ''], ['피토관 계수 Cp', numOrText(report.pitotFactor)], ['가스미터 보정계수 Yd', numOrText(report.gasMeterFactor)], ['오리피스 계수 ΔH@', numOrText(report.deltaHAt)]]);
  if (!isFlow) {
    const leak = report.leakTest || {};
    const leakText = `장치 ${leak.device || 'O'} · 임핀저 ${leak.impinger || 'O'} · 채취부 ${leak.samplingPart || 'O'}`;
    const before = toNum(leak.meterBefore); const after = toNum(leak.meterAfter);
    pairs([['노즐 직경 Dn (mm)', numOrText(report.nozzleDiameterUsed !== undefined ? report.nozzleDiameterUsed : report.nozzleDiameter)], ['K-factor', numOrText(report.kFactorApplied !== undefined ? report.kFactorApplied : report.kFactor)], ['여과지 번호', report.filterId || ''], ['채취 시작 시각', report.samplingStartTime || '']]);
    pairs([['누출시험 (O: 이상 없음)', leakText], ['누출시험 가스미터 (L)', before !== null && after !== null ? `전 ${before} → 후 ${after}` : '']]);
  }

  section('3. 배출가스 조성');
  const gas = Array.isArray(report.gasAnalyzer) ? report.gasAnalyzer : [];
  const gasSpans = [1, 2, 2, 2, 2, 2, 1];
  tableRow(gasSpans, ['회차', '측정 시각', 'O₂ (%)', 'CO₂ (%)', 'CO (ppm)', 'NOx (ppm)', 'SOx (ppm)'], 'head');
  gas.forEach((g, i) => tableRow(gasSpans, [i + 1, g.time || '', numOrText(g.o2), numOrText(g.co2), numOrText(g.co), numOrText(g.nox), numOrText(g.sox)], 'num'));
  tableRow(gasSpans, ['평균', '', round(average(gas.map((g) => g.o2)), 2), round(average(gas.map((g) => g.co2)), 2), round(average(gas.map((g) => g.co)), 1), round(average(gas.map((g) => g.nox)), 1), round(average(gas.map((g) => g.sox)), 1)], 'calc');

  section('4. 배출가스 수분량');
  pairs([['측정 방법', meta.moistureLabel || ''], ['사전 수분량 Xw (%)', numOrText(report.moisturePre)], ['사후 수분량 Xw (%)', numOrText(report.moisturePost)], ['적용 수분량 Xw (%)', numOrText(report.moisturePercent), 'result']]);
  // 사전 수분량(흡습관법·임핀저법): 측정전·후 무게와 가스미터 값
  if (meta.preMoisture) {
    const pre = meta.preMoisture;
    const spans4 = [3, 3, 3, 3];
    tableRow(spans4, [`${pre.unitTitle} (사전 수분량)`, '측정전 무게 ma1 (g)', '측정후 무게 ma2 (g)', '흡습 수분량 (g)'], 'head');
    let ma = 0;
    pre.rows.forEach((row) => {
      const a = toNum(row.initial); const b = toNum(row.final);
      const gain = a !== null && b !== null ? b - a : null;
      if (gain !== null) ma += gain;
      tableRow(spans4, [row.label, numOrText(row.initial), numOrText(row.final), gain === null ? '' : round(gain, 2)], 'num');
    });
    tableRow(spans4, ['흡습 수분의 질량 ma', '', '', round(ma, 2)], 'calc');
    const start = toNum(pre.meterStart); const end = toNum(pre.meterEnd); const yd = toNum(pre.yd) || 1; const orifice = toNum(pre.orifice);
    const spans6 = [2, 2, 2, 2, 2, 2];
    s.height(r, 30);
    tableRow(spans6, ['가스미터 측정전 (L)', '가스미터 측정후 (L)', '가스채취량 Vm (L)', '가스미터 온도 θm (℃)', '측정공에서 기압 Pa (mmHg)', '가스미터압 Pm (mmHg)'], 'head');
    tableRow(spans6, [numOrText(pre.meterStart), numOrText(pre.meterEnd), start !== null && end !== null ? round((end - start) * yd, 2) : '', numOrText(pre.meterTemp), numOrText(pre.pressure), orifice === null ? 0 : round(orifice / 13.6, 2)], 'num');
  }
  const autoValues = (Array.isArray(report.moistureValues) ? report.moistureValues : []).map(toNum).filter((v) => v !== null);
  if (autoValues.length > 0) {
    s.put(r, 0, '자동측정기 측정값 (%)', 'label', 3);
    for (let i = 0; i < 5; i++) s.put(r, 3 + i, autoValues[i] === undefined ? '' : autoValues[i], 'num');
    s.put(r, 8, '평균', 'label'); s.put(r, 9, round(average(autoValues), 2), 'calc', 3);
    r += 1;
  }
  const impingers = (Array.isArray(report.impingers) ? report.impingers : []).filter((imp) => toNum(imp.initial) !== null || toNum(imp.final) !== null);
  if (impingers.length > 0) {
    const impSpans = [3, 3, 3, 3];
    tableRow(impSpans, ['임핀저 (사후 수분량, ES 01321.2)', '채취 전 ma1 (g)', '채취 후 ma2 (g)', '흡습 수분량 (g)'], 'head');
    let total = 0;
    impingers.forEach((imp, i) => {
      const a = toNum(imp.initial); const b = toNum(imp.final);
      const gain = a !== null && b !== null ? b - a : null;
      if (gain !== null) total += gain;
      tableRow(impSpans, [(meta.impingerLabels && meta.impingerLabels[i]) || `${i + 1}단`, numOrText(imp.initial), numOrText(imp.final), gain === null ? '' : round(gain, 2)], 'num');
    });
    tableRow(impSpans, ['흡습 수분의 질량 ma', '', '', round(total, 2)], 'calc');
  }

  section('5. 측정점별 동압 및 온도');
  const points = Array.isArray(report.points) ? report.points : [];
  const pointSpans = [2, 2, 2, 3, 3];
  tableRow(pointSpans, ['측정점', '전압 (mmH₂O)', '정압 (mmH₂O)', '동압 Δp (mmH₂O)', '배출가스 온도 Ts (℃)'], 'head');
  points.forEach((p, i) => tableRow(pointSpans, [i + 1, numOrText(p.tp), numOrText(p.sp), numOrText(p.dp), numOrText(p.ts)], 'num'));
  tableRow(pointSpans, ['평균', round(average(points.map((p) => p.tp)), 2), round(average(points.map((p) => p.sp)), 2), round(average(points.map((p) => p.dp)), 2), round(average(points.map((p) => p.ts)), 1)], 'calc');

  // 기록표: 누적 시간으로 적는다. 화면 입력이 구간(5,5,5…)이면 누적으로 바꾸고, 이미 누적(5,10,15…)이면 그대로
  const meters = (Array.isArray(meta.meterRows) ? meta.meterRows : (Array.isArray(report.gasMeters) ? report.gasMeters : [])).map((m) => ({ ...m }));
  {
    const times = meters.slice(1).map((m) => toNum(m.time)).filter((v) => v !== null && v > 0);
    const cumulative = times.length >= 2 && times.every((v, i) => i === 0 || v >= times[i - 1]);
    if (!cumulative) {
      let running = 0;
      meters.forEach((m, i) => { const t = toNum(m.time); if (i > 0 && t !== null) { running = Math.round((running + t) * 100) / 100; m.time = running; } });
    }
  }
  section(isFlow ? '6. 측정점별 동압·온도 기록' : '6. 시료채취 기록 (ES 01301.1e 표 3의 열 순서)');
  const measured = meters.slice(1);
  if (isFlow) {
    const spans = [2, 2, 2, 3, 3];
    tableRow(spans, ['채취점 번호', '측정 시간 (분)', '배출가스 온도 (℃)', '배출가스 동압 (mmH₂O)', '배출가스 정압 (mmH₂O)'], 'head');
    meters.forEach((m, i) => tableRow(spans, [m.pointNum || (i === 0 ? '시작' : ''), numOrText(m.time), numOrText(m.stackTemp), numOrText(m.dp), numOrText(m.sp)], 'num'));
    tableRow(spans, ['평균', '', round(average(measured.map((m) => m.stackTemp)), 1), round(average(measured.map((m) => m.dp)), 2), round(average(measured.map((m) => m.sp)), 2)], 'calc');
  } else {
    const spans = Array(COLS).fill(1);
    s.height(r, 54);
    tableRow(spans, ['채취점\n번호', '시료채취\n시간 (분)', '진공\n게이지압\n(mmHg)', '배출가스\n온도 (℃)', '배출가스\n동압\n(mmH₂O)', '배출가스\n정압\n(mmH₂O)', '오리피스\n압차\n(mmH₂O)', '가스미터\n지시값 (L)', '건식가스미터\n입구 온도\n(℃)', '건식가스미터\n출구 온도\n(℃)', '여과지홀더\n온도 (℃)', '최종 임핀저\n출구 온도\n(℃)'], 'head');
    meters.forEach((m, i) => tableRow(spans, [m.pointNum || (i === 0 ? '시작' : ''), numOrText(m.time), numOrText(m.vacuum), numOrText(m.stackTemp), numOrText(m.dp), numOrText(m.sp), numOrText(m.pressure), numOrText(m.volume), numOrText(m.tmIn), numOrText(m.tmOut), numOrText(m.filterTemp), numOrText(m.impingerTemp)], 'num'));
    const volumes = meters.map((m) => toNum(m.volume)).filter((v) => v !== null);
    const sampled = volumes.length >= 2 ? Math.round((volumes[volumes.length - 1] - volumes[0]) * 100) / 100 : '';
    const times = measured.map((m) => toNum(m.time)).filter((v) => v !== null);
    tableRow(spans, ['합계', times.length ? Math.max(...times) : '', '', '', '', '', '', sampled, '', '', '', ''], 'calc');
    tableRow(spans, ['평균', '', round(average(measured.map((m) => m.vacuum)), 1), round(average(measured.map((m) => m.stackTemp)), 1), round(average(measured.map((m) => m.dp)), 2), round(average(measured.map((m) => m.sp)), 2), round(average(measured.map((m) => m.pressure)), 2), '', round(average(measured.map((m) => m.tmIn)), 1), round(average(measured.map((m) => m.tmOut)), 1), round(average(measured.map((m) => m.filterTemp)), 1), round(average(measured.map((m) => m.impingerTemp)), 1)], 'calc');
    s.put(r, 0, '※ 시료채취 시간은 누적 분, 합계의 가스미터 지시값은 채취 종료값 − 시작값(보정계수 적용 전). 평균은 시작 행을 뺀 값. 여과지홀더 온도는 현장에서 직접 적는 칸입니다.', 'note', COLS); r += 1;
  }

  section('7. 산출 결과');
  if (isFlow) {
    pairs([['적용 수분량 Xw (%)', numOrText(report.moisturePercent)], ['평균 유속 V (m/s)', numOrText(report.avgVelocity), 'result'], ['습배출가스 유량 (Sm³/hr)', numOrText(report.wetFlowRate), 'result'], ['건조배출가스 유량 (Sm³/hr)', numOrText(report.dryFlowRate), 'result']]);
  } else {
    pairs([['실측 산소농도 Oa (%)', round(average(gas.map((g) => g.o2)), 2)], ['표준산소농도 Os (%)', numOrText(report.standardO2) === '' ? '해당 없음' : numOrText(report.standardO2)], ['배출가스 정압 (mmH₂O)', round(average(points.map((p) => p.sp)), 2)], ['적용 수분량 Xw (%)', numOrText(report.moisturePercent)]]);
    pairs([['총 채취 시간 θ (분)', numOrText(report.totalSamplingMinutes)], ['가스미터 채취량 (L, Yd 적용)', numOrText(report.gasMeterVolume)], ['표준상태 건조가스량 (SL)', numOrText(report.vmStdSL)], ['표준상태 건조가스량 (Sm³)', numOrText(report.vmStdSm3)]]);
    pairs([['평균 가스미터 온도 Tm (℃)', numOrText(report.avgTm)], ['평균 오리피스 압차 ΔH (mmH₂O)', numOrText(report.avgOrifice)], ['평균 유속 V (m/s)', numOrText(report.avgVelocity)], ['등속흡입계수 I (%)', numOrText(report.isokineticRate), 'result']]);
    pairs([['채취 전 무게 (mg)', numOrText(report.filterInitial)], ['채취 후 무게 (mg)', numOrText(report.filterFinal)], ['포집량 (mg)', numOrText(report.dustWeight)], ['등속흡입 판정 (90 %~110 %)', report.isokineticStatus || '', 'result']]);
    pairs([['실측 농도 Ca (mg/Sm³)', numOrText(report.actualConcentration), 'result'], ['표준산소 환산 농도 C (mg/Sm³)', numOrText(report.correctedConcentration), 'result'], ['습배출가스 유량 (Sm³/hr)', numOrText(report.wetFlowRate)], ['건조배출가스 유량 (Sm³/hr)', numOrText(report.dryFlowRate)]]);
  }

  section('8. 현장 특이사항 및 비고');
  s.put(r, 0, report.remarks || '', 'textLeft', COLS).height(r, 48); r += 2;
  s.put(r, 0, '측정자', 'label', 2); s.put(r, 2, report.sampler ? `${report.sampler}  (서명)` : '(서명)', 'text', 4);
  s.put(r, 6, '확인자', 'label', 2); s.put(r, 8, '(서명)', 'text', 4); s.height(r, 30);

  return [isFlow ? buildFlowTableSheet(report, meta, meters) : buildTable3Sheet(report, meta, meters), s, buildFormulaSheet(isFlow)];
};

// ── 「대기오염공정시험기준」 ES 01301.1e "표 3. 먼지시료채취 기록지" 양식 그대로 ──
// 원문에 있는 것만 둔다: 제목, 머리부(왼쪽 8칸·가운데 굴뚝단면 칸·오른쪽 8칸), 11열 표, 합계·평균 줄, 비고 1·2.
const buildTable3Sheet = (report, meta, meters) => {
  const s = createSheet('표 3. 먼지시료채취 기록지', { columns: [7, 7.5, 8.5, 7.5, 9, 9, 9, 8, 8, 8, 8.5] });
  const N = 11;
  let r = 1;
  s.put(r, 0, '표 3. 먼지시료채취 기록지', 'formTitle', N).height(r, 26); r += 1;

  const gas = Array.isArray(report.gasAnalyzer) ? report.gasAnalyzer : [];
  const points = Array.isArray(report.points) ? report.points : [];
  const o2 = average(gas.map((g) => g.o2));
  const staticMmH2O = average(points.map((p) => p.sp));
  const nozzle = toNum(report.nozzleDiameterUsed !== undefined ? report.nozzleDiameterUsed : report.nozzleDiameter);
  // 원문 배치: 왼쪽은 공장명~오리피스미터 △H, 한 줄 띄고 산소량·등속흡입계수 / 오른쪽은 피토관계수~배출가스정압, 한 줄 띄고 여과지 번호
  const left = [
    ['공장명', report.company || ''], ['측정대상명', report.location || ''], ['작성자명', report.sampler || ''], ['측정일', report.date || ''],
    ['측정번호', meta.measureNo || ''], ['오리피스미터  △H', numOrText(report.deltaHAt)], null, null,
    ['산소량 (%)', round(o2, 2)], ['등속흡입계수 (%)', numOrText(report.isokineticRate)],
  ];
  const right = [
    ['피토관계수', numOrText(report.pitotFactor)], ['기온, ℃', numOrText(report.atmTemp)], ['기압, mmHg', numOrText(report.atmPressure)], ['수분량, %', numOrText(report.moisturePercent)],
    ['흡입관 길이, m', meta.probeLengthM === null || meta.probeLengthM === undefined ? '' : meta.probeLengthM], ['흡입노즐 직경, cm', nozzle === null ? '' : round(nozzle / 10, 3)],
    null, ['배출가스정압, mmHg', staticMmH2O === null ? '' : round(staticMmH2O / 13.6, 2)], null, ['여과지 번호', report.filterId || ''],
  ];
  const top = r;
  left.forEach((item, i) => {
    const row = top + i;
    s.height(row, 21);
    if (item) { s.put(row, 0, item[0], 'formLabel', 2); s.put(row, 2, item[1], 'formValue', 2); }
    if (right[i]) { s.put(row, 7, right[i][0], 'formLabel', 2); s.put(row, 9, right[i][1], 'formValue', 2); }
  });
  // 가운데 네모: 굴뚝단면 및 측정점 배열 그림(굴뚝 안지름에 맞춰 그린 것). 그림이 없으면 빈 칸
  s.put(top, 4, '', 'formBox', 3, 7);
  if (meta.sectionPng) s.setImage(meta.sectionPng, 4, top - 1, 7, top - 1 + 7);
  s.put(top + 8, 4, '굴뚝단면 및 측정점\n배열', 'formCaption', 3, 2);
  r = top + left.length + 1;

  // 본문 표 머리 (원문 줄바꿈 그대로)
  s.height(r, 44); s.height(r + 1, 20);
  const heads = ['채취\n점\n번호', '시료\n채취\n시간\n(분)', '진공\n게이\n지압\n(mmHg)', '배출\n가스\n온도\n(℃)', '배출\n가스\n동압\n(mmH₂O)', '오리\n피스\n압차\n(mmH₂O)', '시료\n채취량\n(m³)'];
  heads.forEach((label, c) => s.put(r, c, label, 'formHead', 1, 2));
  s.put(r, 7, '건식가스미터\n에서의\n온도(℃)', 'formHead', 2); s.put(r + 1, 7, '입구', 'formHead'); s.put(r + 1, 8, '출구', 'formHead');
  s.put(r, 9, '여과지\n홀더\n온도\n(℃)', 'formHead', 1, 2); s.put(r, 10, '최종임\n핀저출\n구온도\n(℃)', 'formHead', 1, 2);
  r += 2;
  const m3 = (value) => { const n = toNum(value); return n === null ? '' : Math.round(n * 10) / 10000; };
  const dataRows = Math.max(meters.length, 8);
  for (let i = 0; i < dataRows; i++) {
    const m = meters[i];
    const values = m
      ? [m.pointNum || (i === 0 ? '시작' : ''), numOrText(m.time), numOrText(m.vacuum), numOrText(m.stackTemp), numOrText(m.dp), numOrText(m.pressure), m3(m.volume), numOrText(m.tmIn), numOrText(m.tmOut), numOrText(m.filterTemp), numOrText(m.impingerTemp)]
      : Array(N).fill('');
    values.forEach((value, c) => s.put(r, c, value, c === 6 ? 'formCell4' : 'formCell'));
    s.height(r, 20);
    r += 1;
  }
  const measured = meters.slice(1);
  const volumes = meters.map((m) => toNum(m.volume)).filter((v) => v !== null);
  const times = measured.map((m) => toNum(m.time)).filter((v) => v !== null);
  const avgIn = average(measured.map((m) => m.tmIn)); const avgOut = average(measured.map((m) => m.tmOut));
  const bothTemps = [avgIn, avgOut].filter((v) => v !== null);
  const withLabel = (label, value) => (value === '' || value === null ? label : `${label} ${value}`);
  // 합계 줄: 원문처럼 입구·출구 칸에 "평균"이 적혀 있다
  ['합계', times.length ? Math.max(...times) : '', '', '', '', '', volumes.length >= 2 ? Math.round((volumes[volumes.length - 1] - volumes[0]) * 10) / 10000 : '', withLabel('평균', round(avgIn, 1)), withLabel('평균', round(avgOut, 1)), '', '']
    .forEach((value, c) => s.put(r, c, value, c === 6 ? 'formCell4' : 'formCell'));
  s.height(r, 20); r += 1;
  // 평균 줄: 원문처럼 앞 두 칸과 가스미터 온도 두 칸이 합쳐져 있다
  s.put(r, 0, '평    균', 'formCell', 2);
  [round(average(measured.map((m) => m.vacuum)), 1), round(average(measured.map((m) => m.stackTemp)), 1), round(average(measured.map((m) => m.dp)), 2), round(average(measured.map((m) => m.pressure)), 2), '']
    .forEach((value, i) => s.put(r, 2 + i, value, 'formCell'));
  s.put(r, 7, withLabel('평  균', bothTemps.length ? round(bothTemps.reduce((a, b) => a + b, 0) / bothTemps.length, 1) : ''), 'formCell', 2);
  s.put(r, 9, round(average(measured.map((m) => m.filterTemp)), 1), 'formCell'); s.put(r, 10, round(average(measured.map((m) => m.impingerTemp)), 1), 'formCell');
  s.height(r, 20); r += 1;
  s.put(r, 0, '※비고: 1. 시료채취 중 임핀저 통과 후의 가스온도가 20 ℃가 넘을 경우엔 잘게 부순 얼음을 더 채우거나 소금을 첨가하도록 한다.', 'formNote', N).height(r, 30); r += 1;
  s.put(r, 0, '         2. 먼지가 채취됨에 따라 여과지에 전후의 압력이 너무 높아져 등속흡입을 유지하기가 어려울 경우에 새 여과지로 교환 (여과지홀더도 같이 교환하는 것이 좋다)한 후 시료채취를 계속한다. 이때 총채취먼지량은 각 여과지에 채취된 먼지량의 합이 된다.', 'formNote', N).height(r, 44);
  return s;
};

// ── ES 01114.c "배출가스 유량 측정표" 양식 (유량측정기록부) ──
const buildFlowTableSheet = (report, meta, meters) => {
  const s = createSheet('배출가스 유량 측정표', { columns: [12, 12, 15, 15, 15, 14] });
  const N = 6;
  let r = 1;
  s.put(r, 0, '배출가스 유량 측정표', 'formTitle', N).height(r, 24); r += 1;
  s.put(r, 0, '「대기오염공정시험기준」 ES 01114.c (배출가스 유속 및 유량 측정방법)의 측정표 양식', 'note', N); r += 2;
  const diameter = toNum(report.stackDiameter);
  const left = [['공장명', report.company || ''], ['측정대상명', report.location || ''], ['작성자명', report.sampler || ''], ['측정일', report.date || ''], ['측정번호', meta.measureNo || ''], ['굴뚝 단면적, ㎡', diameter === null ? '' : round(Math.PI * diameter * diameter / 4, 4)]];
  const right = [['기온, ℃', numOrText(report.atmTemp)], ['기압, mmHg', numOrText(report.atmPressure)], ['피토관 계수', numOrText(report.pitotFactor)], ['수분량, %', numOrText(report.moisturePercent)], ['배출가스 밀도, kg/m³', meta.gasDensity === null || meta.gasDensity === undefined ? '' : meta.gasDensity], ['평균 유속, m/s', numOrText(report.avgVelocity)]];
  left.forEach((item, i) => {
    s.put(r, 0, item[0], 'formLabel'); s.put(r, 1, item[1], 'formValue', 2);
    s.put(r, 3, right[i][0], 'formLabel'); s.put(r, 4, right[i][1], 'formValue', 2);
    s.height(r, 21); r += 1;
  });
  r += 1;
  const rows = meters.slice(1).length > 0 ? meters.slice(1) : (Array.isArray(report.points) ? report.points.map((p, i) => ({ pointNum: i + 1, dp: p.dp, stackTemp: p.ts, sp: p.sp })) : []);
  ['측정점 번호', '측정 시간 (분)', '배출가스 동압 h\n(mmH₂O)', '배출가스 온도 θs\n(℃)', '배출가스 정압 Ps\n(mmHg)', '√h'].forEach((label, c) => s.put(r, c, label, 'formHead'));
  s.height(r, 34); r += 1;
  const roots = [];
  rows.forEach((m) => {
    const h = toNum(m.dp); const sp = toNum(m.sp);
    const root = h !== null && h >= 0 ? Math.sqrt(h) : null;
    if (root !== null) roots.push(root);
    [m.pointNum || '', numOrText(m.time), numOrText(m.dp), numOrText(m.stackTemp), sp === null ? '' : round(sp / 13.6, 3), round(root, 3)].forEach((value, c) => s.put(r, c, value, 'num'));
    s.height(r, 20); r += 1;
  });
  const avgSp = average(rows.map((m) => m.sp));
  ['평균', '', round(average(rows.map((m) => m.dp)), 2), round(average(rows.map((m) => m.stackTemp)), 1), avgSp === null ? '' : round(avgSp / 13.6, 3), roots.length ? round(roots.reduce((a, b) => a + b, 0) / roots.length, 3) : '']
    .forEach((value, c) => s.put(r, c, value, 'num'));
  s.height(r, 20); r += 1;
  s.put(r, 0, '※ 배출시설의 조업 및 가동상태는 별지에 기록한다. 수분량·배출가스 조성·유량 산출 결과는 "세부 기록" 시트에 있습니다.', 'note', N).height(r, 28);
  return s;
};

// 계산 근거 시트: 화면이 쓰는 계산식을 「대기오염공정시험기준」의 식과 기호로 적어 둔다
const buildFormulaSheet = (isFlow) => {
  const s = createSheet('계산 근거', { columns: [26, 66, 50, 22] });
  let r = 1;
  s.put(r, 0, '계산 근거', 'title', 4).height(r, 30); r += 1;
  s.put(r, 0, '기록지의 산출 결과는 「대기오염공정시험기준」(국립환경과학원고시)의 아래 식으로 계산한 값입니다.', 'subtitle', 4); r += 2;
  s.put(r, 0, '항목', 'head'); s.put(r, 1, '계산식', 'head'); s.put(r, 2, '기호', 'head'); s.put(r, 3, '근거', 'head'); r += 1;
  const rows = [
    ['수분량 Xw (%)', 'Xw = (22.4/18 × ma) ÷ [ Vm × 273/(273+θm) × (Pa+Pm)/760 + 22.4/18 × ma ] × 100', 'ma: 흡습 수분의 질량(g), Vm: 흡입한 건조가스량(L), θm: 가스미터 온도(℃), Pa: 대기압(mmHg), Pm: 가스미터 게이지압(mmHg)', 'ES 01114.c 식 10\nES 01321.2'],
    ['배출가스 밀도 γ (kg/m³)', 'γ = γo × 273/(273+θs) × (Pa+Ps)/760', 'γo: 0 ℃·760 mmHg에서 습한 배출가스 밀도(kg/Sm³), θs: 배출가스 온도(℃), Ps: 배출가스 정압(mmHg, 측정값 mmH₂O ÷ 13.6)', 'ES 01114.c 식 11'],
    ['배출가스 유속 V (m/s)', 'V = C × √(2 g h / γ)', 'C: 피토관 계수, g: 중력가속도, h: 배출가스 동압(mmH₂O), γ: 습한 배출가스 밀도(kg/m³)', 'ES 01114.c 식 12'],
    ['건조 배출가스 유량 QN (Sm³/h)', 'QN = V × A × 273/(273+θs) × (Pa+Ps)/760 × (1 − Xw/100) × 3,600', 'A: 굴뚝 단면적(㎡). 습한 배출가스 유량은 (1 − Xw/100)을 곱하지 않은 값', 'ES 01114.c 식 13'],
  ];
  if (!isFlow) {
    rows.push(
      ['등속흡입계수 I (%)', 'I = Ts × [ 0.00346 × Vic + (V′m / Tm) × (Pa + ΔH/13.6) ] ÷ ( Ps × t × v × An ) × (1.667 × 10⁴)', 'Ts: 배출가스 평균 절대온도(K), Vic: 임핀저와 실리카젤에 채취된 물의 총량(mL), V′m: 건식가스미터에서 읽은 가스시료채취량(m³, 보정계수 적용), Tm: 건식가스미터 평균 절대온도(K), ΔH: 오리피스 압차(mmH₂O), Ps: 배출가스 압력(mmHg), t: 총 시료채취시간(분), v: 배출가스 유속(m/s), An: 노즐 단면적(cm²). 90 %~110 % 범위 밖이면 다시 채취', 'ES 01301.1e 식 8'],
      ['먼지 농도 Cn (mg/Sm³)', 'Cn = md ÷ [ V′m × 273/(273+θm) × (Pa + ΔH/13.6)/760 ]', 'md: 채취된 먼지량(mg), 분모: 표준상태(0 ℃, 760 mmHg) 건조가스 채취량(Sm³)', 'ES 01301.1e 식 9'],
      ['표준산소농도 환산 C (mg/Sm³)', 'C = Ca × (21 − Os) ÷ (21 − Oa)', 'Ca: 실측 농도, Os: 표준산소농도(%), Oa: 실측산소농도(%). 표준산소농도가 정해지지 않은 시설은 환산하지 않음', 'ES 01000.b 총칙 1.9'],
    );
  }
  rows.forEach(([name, formula, symbols, basis]) => {
    s.put(r, 0, name, 'label'); s.put(r, 1, formula, 'formula'); s.put(r, 2, symbols, 'formula'); s.put(r, 3, basis, 'text'); s.height(r, 66); r += 1;
  });
  r += 1;
  s.put(r, 0, '※ 이 기록지는 프로그램이 직접 만든 양식입니다. 공정시험기준은 기록지를 "양식의 한 예"로 제시하고 있어, 기준의 기록 항목을 모두 담되 배치는 달리했습니다.', 'note', 4);
  return s;
};

// ── 시료채취기록부(종합) 공개 양식 ──
// 「환경분야 시험·검사 업무처리규정」 별표 1의 시료채취기록부 기록 사항(시료채취일시, 시료채취자, 시료명·시료번호,
// 관리분야, 시험항목, 시료채취량, 현장측정결과, 기타 참고사항)과 측정기록부(별지 제21호서식)의 현장기상 항목을 담는다.
// data: { company, industry, facility, stackName, samplers: [{ label, name, yd }], days: [...], rows: [...], requestedCount }
export const buildSamplingLogSheets = (data) => {
  const widths = [25, 21, 11.5, 7.5, 7.5, 10.5, 10.5, 10.5, 9, 8, 9, 22];
  const s = createSheet('시료채취기록부', { columns: widths, landscape: true });
  const N = widths.length;
  let r = 1;
  const section = (title) => { r += 1; s.put(r, 0, title, 'section', N).height(r, 21); r += 1; };
  s.put(r, 0, '시료채취기록부', 'title', N).height(r, 32); r += 1;
  s.put(r, 0, '관리분야: 대기 · 「환경분야 시험·검사 업무처리규정」 별표 1의 시료채취기록부 기록 사항을 따름 (작성 후 3년 보관)', 'subtitle', N); r += 1;

  section('1. 일반 사항');
  const pair = (c, label, value, span) => { s.put(r, c, label, 'label'); s.put(r, c + 1, value === undefined || value === null ? '' : value, 'text', span); };
  pair(0, '사업장명', data.company, 3); pair(4, '업종', data.industry, 3); pair(8, '관리분야', '대기', 3); s.height(r, 26); r += 1;
  pair(0, '배출시설', data.facility, 3); pair(4, '배출구', data.stackName, 3); pair(8, '의뢰 항목 수', `${data.requestedCount || 0}종`, 3); s.height(r, 26); r += 1;
  (data.samplers || []).forEach((sampler, i) => {
    if (i % 2 === 0 && i > 0) { s.height(r, 26); r += 1; }
    const c = (i % 2) * 6;
    s.put(r, c, sampler.label, 'label'); s.put(r, c + 1, sampler.name || '', 'text', 3); s.put(r, c + 4, '보정계수 Yd', 'label'); s.put(r, c + 5, numOrText(sampler.yd), 'num');
  });
  if ((data.samplers || []).length > 0) { s.height(r, 26); r += 1; }

  section('2. 시료채취일 및 현장 기상');
  const daySpans = [1, 2, 2, 1, 1, 2, 1, 2];
  const dayRow = (values, style) => { let c = 0; daySpans.forEach((span, i) => { s.put(r, c, values[i] === undefined ? '' : values[i], style, span); c += span; }); r += 1; };
  dayRow(['일차', '시료채취일', '날씨', '기온 (℃)', '습도 (%)', '기압 (mmHg)', '풍향', '풍속 (m/s)'], 'head');
  (data.days || []).forEach((day, i) => {
    if (!day.date) return;
    dayRow([`${i + 1}일차`, day.date, day.weather || '', numOrText(day.temp), numOrText(day.humidity), numOrText(day.pressure), day.windDir || '', numOrText(day.windSpeed)], 'num');
  });

  section('3. 시료별 채취 기록');
  s.height(r, 44);
  ['시료번호', '시험항목', '시료채취일', '채취\n시작', '채취\n종료', '가스미터\n시작값 (L)', '가스미터\n종료값 (L)', '시료채취량\n(L)', '가스미터\n온도 (℃)', '압력', '흡인유량\n(L/min)', '기타 참고사항'].forEach((label, c) => s.put(r, c, label, 'head'));
  r += 1;
  (data.rows || []).forEach((row) => {
    [row.sampleNo || '', row.name || '', row.date || '', row.start || '', row.end || '', numOrText(row.meterStart), numOrText(row.meterEnd), row.volume === null || row.volume === undefined ? '' : row.volume, numOrText(row.temp), numOrText(row.pressure), numOrText(row.suctionRate), row.note || '']
      .forEach((value, c) => s.put(r, c, value, c <= 1 || c === 11 ? 'text' : 'num'));
    s.height(r, 30);
    r += 1;
  });
  if ((data.rows || []).length === 0) { s.put(r, 0, '의뢰 표시한 항목이 없습니다.', 'text', N); r += 1; }
  s.put(r, 0, '※ 시료채취량: 입자·가스 시료채취기는 (종료값 − 시작값) × 보정계수 Yd, 그 밖에는 종료값 − 시작값. 압력의 단위는 시료채취 장치에 따릅니다(오리피스 압차 mmH₂O 등).', 'note', N); r += 1;

  section('4. 현장 특이사항 (육안 관찰, 보존제 등)');
  s.put(r, 0, data.remarks || '', 'textLeft', N).height(r, 44); r += 2;
  s.put(r, 0, '시료채취자', 'label', 1, 2);
  ['성명', '직급', '서명'].forEach((label, i) => { s.put(r, 1 + i * 3, label, 'label'); s.put(r, 2 + i * 3, '', 'text', 2); });
  s.put(r, 10, '', 'text', 2);
  s.height(r, 28); r += 1;
  ['성명', '직급', '서명'].forEach((label, i) => { s.put(r, 1 + i * 3, label, 'label'); s.put(r, 2 + i * 3, '', 'text', 2); });
  s.put(r, 10, '', 'text', 2);
  s.height(r, 28);
  return [s];
};
