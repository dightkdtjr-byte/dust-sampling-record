// 엑셀 파일 안의 계산식 결과(저장값)를 프로그램이 직접 계산해 채워 넣는다.
//
// 왜 필요한가: 엑셀 데스크톱은 파일을 열 때 다시 계산하지만, 휴대폰 미리보기·엑셀 보호된 보기처럼
// 다시 계산하지 않는 화면은 파일에 저장된 값을 그대로 보여 준다. 양식을 고쳐 쓰면 그 저장값이
// 비거나 양식의 옛 숫자로 남아, 밀도가 안 보이거나 #DIV/0! 같은 에러가 보였다.
//
// 방법: 모든 시트의 값·계산식을 읽어 fast-formula-parser로 계산하고, 각 계산식 칸의 <v>를 바꿔 쓴다.
// 계산하지 못한 칸은 저장값을 지워(빈 칸) 옛 숫자가 보이지 않게 하고, 엑셀은 열 때 다시 계산한다.
import FormulaParserModule from 'fast-formula-parser';
import FormulaOperators from 'fast-formula-parser/formulas/operators.js';

// 엑셀처럼 빈 칸은 글자와 비교할 땐 ""로, 숫자와 비교할 땐 0으로 본다 (계산 도구는 늘 0으로 봐서 IF(A1="",...)가 틀렸다)
{
  const Infix = (FormulaOperators && (FormulaOperators.Infix || FormulaOperators.default?.Infix)) || null;
  if (Infix && !Infix.__blankAware) {
    const original = Infix.compareOp;
    Infix.compareOp = (value1, infix, value2, isArray1, isArray2) => {
      const first = (v, isArr) => (isArr && Array.isArray(v) ? v[0]?.[0] : v);
      const a = first(value1, isArray1); const b = first(value2, isArray2);
      if ((a === null || a === undefined) && typeof b === 'string') return original('', infix, b, false, false);
      if ((b === null || b === undefined) && typeof a === 'string') return original(a, infix, '', false, false);
      return original(value1, infix, value2, isArray1, isArray2);
    };
    Infix.__blankAware = true;
  }
}

const FormulaParser = FormulaParserModule.default || FormulaParserModule;
const { FormulaError } = FormulaParserModule.default ? FormulaParserModule.default : FormulaParserModule;

const colToNum = (letters) => {
  let n = 0;
  for (const ch of letters) n = n * 26 + (ch.charCodeAt(0) - 64);
  return n;
};
const numToCol = (n) => {
  let s = '';
  while (n > 0) { const m = (n - 1) % 26; s = String.fromCharCode(65 + m) + s; n = Math.floor((n - 1) / 26); }
  return s;
};
const decodeXml = (text) => String(text)
  .replace(/&lt;/g, '<').replace(/&gt;/g, '>').replace(/&quot;/g, '"').replace(/&apos;/g, "'")
  .replace(/&#(\d+);/g, (_, d) => String.fromCodePoint(Number(d)))
  .replace(/&#x([0-9a-fA-F]+);/g, (_, h) => String.fromCodePoint(parseInt(h, 16)))
  .replace(/&amp;/g, '&');
const encodeXml = (text) => String(text).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
const attr = (attrs, name) => {
  const m = new RegExp(`\\b${name}="([^"]*)"`).exec(attrs);
  return m ? m[1] : null;
};

// 공유 수식(shared formula): 기준 칸의 수식을 행·열 차이만큼 옮긴다. 문자열·$고정·시트 이름은 그대로 둔다
const shiftFormula = (formula, dRow, dCol) => {
  if (!dRow && !dCol) return formula;
  let out = '';
  let i = 0;
  while (i < formula.length) {
    const ch = formula[i];
    if (ch === '"') { // 문자열
      let j = i + 1;
      while (j < formula.length) { if (formula[j] === '"') { if (formula[j + 1] === '"') { j += 2; continue; } break; } j += 1; }
      out += formula.slice(i, j + 1); i = j + 1; continue;
    }
    if (ch === "'") { // 따옴표 시트 이름
      let j = i + 1;
      while (j < formula.length) { if (formula[j] === "'") { if (formula[j + 1] === "'") { j += 2; continue; } break; } j += 1; }
      out += formula.slice(i, j + 1); i = j + 1; continue;
    }
    const m = /^(\$?)([A-Z]{1,3})(\$?)(\d+)(?![\d(A-Za-z_])/.exec(formula.slice(i));
    const prev = i > 0 ? formula[i - 1] : '';
    if (m && !/[A-Za-z0-9_.ㄱ-힝]/.test(prev)) {
      const [, cAbs, col, rAbs, row] = m;
      const newCol = cAbs ? col : numToCol(colToNum(col) + dCol);
      const newRow = rAbs ? row : String(Number(row) + dRow);
      out += `${cAbs}${newCol}${rAbs}${newRow}`;
      i += m[0].length; continue;
    }
    out += ch; i += 1;
  }
  return out;
};

const isErr = (v) => v instanceof FormulaError || (v && typeof v === 'object' && typeof v.error === 'string' && v.error.startsWith('#'));
const errText = (v) => (v instanceof FormulaError ? v.error : (v && v.error) || '#VALUE!');

// AGGREGATE(15=SMALL / 14=LARGE / 4=MAX / 5=MIN / 9=SUM / 1=AVERAGE, 옵션 6=에러 무시, 배열, k)
const flatten = (value) => {
  if (Array.isArray(value)) return value.flat(Infinity);
  if (value && typeof value === 'object' && Array.isArray(value.value)) return value.value.flat(Infinity);
  if (value && typeof value === 'object' && 'value' in value) return [value.value];
  return [value];
};

// 함수 인자에서 값만 꺼낸다 (범위·배열은 펼친다). literal은 수식에 직접 적은 값
const argValues = (param) => {
  if (param === null || param === undefined) return { values: [null], literal: true };
  if (typeof param === 'object' && 'value' in param) {
    return { values: flatten(param.value), literal: !param.isRangeRef && !param.isCellRef && !Array.isArray(param.value) };
  }
  return { values: flatten(param), literal: true };
};
const numbersFrom = (params) => {
  const nums = [];
  for (const param of params) {
    const { values, literal } = argValues(param);
    for (const v of values) {
      if (isErr(v)) throw v;
      if (typeof v === 'number' && Number.isFinite(v)) nums.push(v);
      else if (literal && typeof v === 'boolean') nums.push(v ? 1 : 0);
      else if (literal && typeof v === 'string' && v.trim() !== '' && Number.isFinite(Number(v))) nums.push(Number(v));
      else if (literal && typeof v === 'string' && v.trim() !== '') throw FormulaError.VALUE;
    }
  }
  return nums;
};
const scalar = (param) => argValues(param).values[0];
// 엑셀 비교: 숫자 < 문자 < 논리값, 문자는 대소문자 무시
const compareCells = (a, b) => {
  const rank = (v) => (typeof v === 'number' ? 0 : typeof v === 'string' ? 1 : typeof v === 'boolean' ? 2 : 3);
  if (rank(a) !== rank(b)) return rank(a) - rank(b);
  if (typeof a === 'string') return a.toLowerCase().localeCompare(b.toLowerCase());
  return a < b ? -1 : a > b ? 1 : 0;
};
// 문자열·따옴표 시트 이름을 건너뛰며 i부터 짝이 맞는 닫는 괄호 위치를 찾는다
const skipQuoted = (text, i) => {
  const q = text[i];
  let j = i + 1;
  while (j < text.length) { if (text[j] === q) { if (text[j + 1] === q) { j += 2; continue; } return j; } j += 1; }
  return j;
};
const splitCall = (text, open) => { // open = '(' 위치. → { args: [..], end: ')' 위치 }
  const args = []; let depth = 0; let start = open + 1;
  for (let i = open; i < text.length; i++) {
    const ch = text[i];
    if (ch === '"' || ch === "'") { i = skipQuoted(text, i); continue; }
    if (ch === '(' || ch === '{') depth += 1;
    else if (ch === ')' || ch === '}') { depth -= 1; if (depth === 0) { args.push(text.slice(start, i)); return { args, end: i }; } }
    else if (ch === ',' && depth === 1) { args.push(text.slice(start, i)); start = i + 1; }
  }
  return null;
};
const RANGE_RE = /((?:'(?:[^']|'')+'|[^\s'!(),=<>+\-*/&^"{}:;]+)!)?(\$?)([A-Z]{1,3})(\$?)(\d+):(\$?)([A-Z]{1,3})(\$?)(\d+)/g;
// 배열식(범위끼리의 계산)을 행마다 한 칸짜리 식으로 푼다: 범위 A2:A9 → i번째 칸 A(2+i)
const liftRows = (expr) => {
  let n = 0;
  const strip = expr.replace(/"(?:[^"]|"")*"/g, (m) => ' '.repeat(m.length));
  for (const m of strip.matchAll(RANGE_RE)) n = Math.max(n, Number(m[9]) - Number(m[5]) + 1);
  if (n <= 1) return [expr];
  const out = [];
  for (let i = 0; i < n; i++) {
    let line = ''; let last = 0;
    for (const m of strip.matchAll(RANGE_RE)) {
      const rows = Number(m[9]) - Number(m[5]) + 1;
      line += expr.slice(last, m.index);
      line += rows > 1 ? `${m[1] || ''}${m[2]}${m[3]}${m[4]}${Number(m[5]) + Math.min(i, rows - 1)}` : expr.slice(m.index, m.index + m[0].length);
      last = m.index + m[0].length;
    }
    out.push(line + expr.slice(last));
  }
  return out;
};
const literalOf = (value) => {
  if (isErr(value)) return errText(value);
  if (typeof value === 'number') return Number.isFinite(value) ? (value < 0 ? `(${value})` : String(value)) : '#NUM!';
  if (typeof value === 'boolean') return value ? 'TRUE' : 'FALSE';
  if (value === null || value === undefined) return '0';
  return `"${String(value).replace(/"/g, '""')}"`;
};
const aggregateOf = (fn, opt, raw, kk) => {
  const ignoreErrors = [2, 3, 6, 7].includes(opt);
  const nums = [];
  for (const v of raw) {
    if (isErr(v)) { if (ignoreErrors) continue; return v; }
    if (typeof v === 'number' && Number.isFinite(v)) nums.push(v);
  }
  const sorted = [...nums].sort((a, b) => a - b);
  switch (fn) {
    case 15: return Number.isFinite(kk) && kk >= 1 && kk <= sorted.length ? sorted[kk - 1] : FormulaError.NUM;
    case 14: return Number.isFinite(kk) && kk >= 1 && kk <= sorted.length ? sorted[sorted.length - kk] : FormulaError.NUM;
    case 4: return sorted.length ? sorted[sorted.length - 1] : 0;
    case 5: return sorted.length ? sorted[0] : 0;
    case 9: return nums.reduce((a, b) => a + b, 0);
    case 1: return nums.length ? nums.reduce((a, b) => a + b, 0) / nums.length : FormulaError.DIV0;
    case 2: return nums.length;
    default: return FormulaError.VALUE;
  }
};
const NAME_WORD = /^[A-Za-z_\u0080-￿][A-Za-z0-9_.\u0080-￿]*/;
const CELL_WORD = /^[A-Z]{1,3}\d+$/;

export async function recalcWorkbookCaches(zip, { today = new Date() } = {}) {
  const readText = async (path) => { const file = zip.file(path); return file ? file.async('string') : null; };
  const workbookXml = await readText('xl/workbook.xml');
  const relsXml = await readText('xl/_rels/workbook.xml.rels');
  if (!workbookXml || !relsXml) return { computed: 0, failed: 0 };

  const rels = {};
  for (const m of relsXml.matchAll(/<Relationship\b([^>]*)\/?>/g)) {
    const id = attr(m[1], 'Id'); const target = attr(m[1], 'Target');
    if (id && target) rels[id] = target.startsWith('/') ? target.slice(1) : `xl/${target.replace(/^\.\//, '')}`;
  }
  const sheets = [];
  for (const m of workbookXml.matchAll(/<sheet\b([^>]*)\/>/g)) {
    const name = decodeXml(attr(m[1], 'name') || '');
    const rid = attr(m[1], 'r:id');
    if (rid && rels[rid]) sheets.push({ name, path: rels[rid] });
  }
  const sharedXml = await readText('xl/sharedStrings.xml');
  const shared = [];
  if (sharedXml) {
    for (const m of sharedXml.matchAll(/<si>([\s\S]*?)<\/si>/g)) {
      shared.push([...m[1].matchAll(/<t[^>]*>([\s\S]*?)<\/t>/g)].map((t) => decodeXml(t[1])).join(''));
    }
  }

  // 시트별 칸: key "row,col" → { value, formula, xmlCell }
  const books = new Map();
  const sheetXmls = new Map();
  for (const sheet of sheets) {
    const xml = await readText(sheet.path);
    if (!xml) continue;
    sheetXmls.set(sheet.name, xml);
    const cells = new Map();
    const sharedMasters = new Map();
    for (const m of xml.matchAll(/<c\b([^>]*?)(?:\/>|>([\s\S]*?)<\/c>)/g)) {
      const attrs = m[1]; const inner = m[2] || '';
      const ref = attr(attrs, 'r'); if (!ref) continue;
      const rm = /^([A-Z]+)(\d+)$/.exec(ref); if (!rm) continue;
      const row = Number(rm[2]); const col = colToNum(rm[1]);
      const t = attr(attrs, 't');
      const fm = /<f\b([^>]*?)(?:\/>|>([\s\S]*?)<\/f>)/.exec(inner);
      const vm = /<v>([\s\S]*?)<\/v>/.exec(inner);
      let value = null;
      if (t === 's' && vm) value = shared[Number(vm[1])] ?? '';
      else if (t === 'inlineStr') value = [...inner.matchAll(/<t[^>]*>([\s\S]*?)<\/t>/g)].map((x) => decodeXml(x[1])).join('');
      else if (t === 'str' && vm) value = decodeXml(vm[1]);
      else if (t === 'b' && vm) value = vm[1] === '1';
      else if (t === 'e' && vm) value = { error: decodeXml(vm[1]) };
      else if (vm && vm[1] !== '') value = Number(vm[1]);
      const cell = { row, col, value };
      if (fm) {
        const fAttrs = fm[1]; const fText = fm[2] !== undefined ? decodeXml(fm[2]) : '';
        const fType = attr(fAttrs, 't');
        if (fType === 'shared') {
          const si = attr(fAttrs, 'si');
          if (fText) { sharedMasters.set(si, { row, col, formula: fText }); cell.formula = fText; } else cell.sharedSi = si;
        } else if (fText) cell.formula = fText;
        cell.isArray = fType === 'array';
      }
      cells.set(`${row},${col}`, cell);
    }
    for (const cell of cells.values()) {
      if (cell.sharedSi !== undefined) {
        const master = sharedMasters.get(cell.sharedSi);
        if (master) cell.formula = shiftFormula(master.formula, cell.row - master.row, cell.col - master.col);
        delete cell.sharedSi;
      }
    }
    books.set(sheet.name, cells);
  }

  // 정의된 이름 (예: Ca → 계산식!$P$97)
  const names = new Map();
  for (const m of workbookXml.matchAll(/<definedName\b([^>]*)>([^<]*)<\/definedName>/g)) {
    const name = attr(m[1], 'name'); const text = decodeXml(m[2]);
    if (!name || attr(m[1], 'localSheetId') !== null) continue;
    const dm = /^'?([^'!]+)'?!\$?([A-Z]+)\$?(\d+)(?::\$?([A-Z]+)\$?(\d+))?$/.exec(text);
    if (!dm) continue;
    const ref = { sheet: dm[1], row: Number(dm[3]), col: colToNum(dm[2]) };
    names.set(name.toUpperCase(), dm[4] ? { sheet: dm[1], from: { row: ref.row, col: ref.col }, to: { row: Number(dm[5]), col: colToNum(dm[4]) } } : ref);
  }

  // 수식 안의 정의된 이름을 실제 칸 주소로 바꾼다 (계산 도구가 한 글자 이름 D 같은 것을 못 읽는다)
  const quoteSheet = (name) => `'${String(name).replace(/'/g, "''")}'`;
  const nameText = new Map([...names].map(([key, ref]) => [key, ref.from
    ? `${quoteSheet(ref.sheet)}!$${numToCol(ref.from.col)}$${ref.from.row}:$${numToCol(ref.to.col)}$${ref.to.row}`
    : `${quoteSheet(ref.sheet)}!$${numToCol(ref.col)}$${ref.row}`]));
  const expandNames = (formula) => {
    if (nameText.size === 0) return formula;
    let out = ''; let i = 0;
    while (i < formula.length) {
      const ch = formula[i];
      if (ch === '"' || ch === "'") {
        let j = i + 1;
        while (j < formula.length) { if (formula[j] === ch) { if (formula[j + 1] === ch) { j += 2; continue; } break; } j += 1; }
        out += formula.slice(i, j + 1); i = j + 1; continue;
      }
      const m = NAME_WORD.exec(formula.slice(i));
      if (m) {
        const word = m[0]; const next = formula[i + word.length] || ''; const prev = i > 0 ? formula[i - 1] : '';
        const key = word.toUpperCase();
        const isName = !CELL_WORD.test(word) && next !== '(' && next !== '!' && next !== ':' && prev !== '!' && prev !== '$' && prev !== ':' && nameText.has(key);
        out += isName ? nameText.get(key) : word;
        i += word.length; continue;
      }
      out += ch; i += 1;
    }
    return out;
  };

  // 계산 도구는 계산 도중에 다른 계산을 부르면 상태가 꼬인다. 그래서 한 칸을 계산하기 전에
  // 그 칸이 참조하는 칸들을 먼저 (재귀로) 계산해 두고, 계산 중의 칸 조회는 이미 계산된 값만 읽는다
  const state = new Map(); // "sheet|row,col" → 'busy' | { result }
  const maxRow = new Map([...books].map(([name, cells]) => [name, Math.max(0, ...[...cells.values()].map((c) => c.row))]));
  const depParser = new FormulaParser.DepParser({ onVariable: () => null });
  let parser;
  const peekCell = (sheetName, row, col) => {
    const cell = books.get(sheetName)?.get(`${row},${col}`);
    if (!cell) return null;
    if (!cell.formula) return cell.value;
    const known = state.get(`${sheetName}|${row},${col}`);
    return known && known !== 'busy' ? known.result : cell.value;
  };
  const evalCell = (sheetName, row, col) => {
    const cell = books.get(sheetName)?.get(`${row},${col}`);
    if (!cell) return null;
    if (!cell.formula) return cell.value;
    const key = `${sheetName}|${row},${col}`;
    const known = state.get(key);
    if (known === 'busy') return cell.value; // 순환 참조: 저장값을 쓴다
    if (known) return known.result;
    state.set(key, 'busy');
    if (cell.expr === undefined) cell.expr = expandNames(cell.formula.replace(/_xlfn\./g, '').replace(/_xlws\./g, ''));
    const position = { sheet: sheetName, row, col };
    let deps = [];
    try { deps = depParser.parse(cell.expr, position, true) || []; } catch { deps = []; }
    for (const dep of deps) {
      const depSheet = dep.sheet || sheetName;
      if (!books.has(depSheet)) continue;
      if (dep.from) {
        const lastRow = Math.min(dep.to.row, maxRow.get(depSheet) || 0);
        for (let r = dep.from.row; r <= lastRow; r++) for (let c = dep.from.col; c <= dep.to.col; c++) evalCell(depSheet, r, c);
      } else evalCell(depSheet, dep.row, dep.col);
    }
    let expr = cell.expr;
    // AGGREGATE(함수, 옵션, 배열식, k): 배열식을 행마다 계산해 결과 값으로 바꿔 넣는다 (계산 도구가 배열식을 못 푼다)
    for (let guard = 0; guard < 20; guard++) {
      const at = expr.toUpperCase().indexOf('AGGREGATE(');
      if (at < 0) break;
      const call = splitCall(expr, at + 'AGGREGATE'.length);
      if (!call || call.args.length < 3) break;
      const scalarOf = (text) => { try { const v = parser.parse(text, position); return Array.isArray(v) ? flatten(v)[0] : v; } catch (error) { return isErr(error) ? error : FormulaError.VALUE; } };
      const fn = Number(scalarOf(call.args[0]));
      const opt = Number(scalarOf(call.args[1]));
      const raw = liftRows(call.args[2]).map(scalarOf);
      const kk = call.args[3] !== undefined ? Number(scalarOf(call.args[3])) : NaN;
      expr = expr.slice(0, at) + literalOf(aggregateOf(fn, opt, raw, kk)) + expr.slice(call.end + 1);
    }
    let result;
    try {
      result = parser.parse(expr, position, cell.isArray);
      if (Array.isArray(result)) result = flatten(result)[0];
    } catch (error) {
      result = isErr(error) ? error : { error: '#VALUE!', failed: true };
    }
    state.set(key, { result });
    return result;
  };

  parser = new FormulaParser({
    onCell: ({ sheet, row, col }) => peekCell(sheet, row, col),
    onRange: (ref) => {
      const cells = books.get(ref.sheet);
      const out = [];
      const toRow = Math.min(ref.to.row, cells ? Math.max(0, ...[...cells.values()].map((c) => c.row)) : 0);
      for (let r = ref.from.row; r <= Math.max(ref.from.row, toRow); r++) {
        const line = [];
        for (let c = ref.from.col; c <= ref.to.col; c++) line.push(peekCell(ref.sheet, r, c));
        out.push(line);
      }
      return out;
    },
    onVariable: (name) => names.get(String(name).toUpperCase()) || null,
    // INDEX는 계산 도구에서 context를 받는 함수라 같은 자리(functionsNeedContext)에 넣어야 덮어쓴다
    functionsNeedContext: {
      // 계산 도구의 INDEX는 다른 시트 범위를 넣으면 시트를 잃어버려 값이 비므로 값으로 직접 돌려준다
      INDEX: (context, arrayArg, rowArg, colArg) => {
        // INDEX는 인자를 참조 그대로 받으므로 값으로 풀어 쓴다
        const take = (arg) => (arg === null || arg === undefined ? arg : { value: context.utils.extractRefValue(arg).val });
        const array = take(arrayArg); const rowNum = take(rowArg); const colNum = take(colArg);
        const raw = array && typeof array === 'object' && 'value' in array ? array.value : array;
        if (isErr(raw)) return raw;
        const grid = Array.isArray(raw) ? (Array.isArray(raw[0]) ? raw : [raw]) : [[raw]];
        let r = Number(scalar(rowNum));
        let c = colNum === undefined || colNum === null ? NaN : Number(scalar(colNum));
        if (isErr(scalar(rowNum))) return scalar(rowNum);
        if (!Number.isFinite(c)) {
          if (grid.length === 1 && grid[0].length > 1) { c = r; r = 1; } else c = 1;
        }
        r = Math.trunc(r); c = Math.trunc(c);
        if (r < 1 || c < 1) return FormulaError.VALUE;
        const line = grid[r - 1];
        if (!line && r > grid.length) return null; // 범위 안이지만 비어 있는 행
        const v = line ? line[c - 1] : null;
        return v === undefined ? null : v;
      },
    },
    functions: {
      MAX: (...params) => { const n = numbersFrom(params); return n.length ? Math.max(...n) : 0; },
      MIN: (...params) => { const n = numbersFrom(params); return n.length ? Math.min(...n) : 0; },
      COUNTA: (...params) => params.reduce((sum, param) => sum + argValues(param).values.filter((v) => v !== null && v !== undefined).length, 0),
      MATCH: (lookup, array, type) => {
        const want = scalar(lookup);
        if (isErr(want)) return want;
        const list = argValues(array).values;
        const mode = type === undefined || type === null ? 1 : Number(scalar(type));
        if (mode === 0) {
          const idx = list.findIndex((v) => v !== null && v !== undefined && typeof v === typeof want && compareCells(v, want) === 0);
          return idx >= 0 ? idx + 1 : FormulaError.NA;
        }
        let found = -1;
        for (let i = 0; i < list.length; i++) {
          const v = list[i];
          if (v === null || v === undefined || typeof v !== typeof want) continue;
          const c = compareCells(v, want);
          if (mode > 0 ? c <= 0 : c >= 0) found = i; else break;
        }
        return found >= 0 ? found + 1 : FormulaError.NA;
      },
      // 날짜 일련번호(엑셀 1900 기준: 1 = 1900-01-01, 60 = 없는 1900-02-29)
      DAY: (serial) => {
        const n = Number(scalar(serial) ?? 0);
        if (!Number.isFinite(n) || n < 0) return FormulaError.NUM;
        const d = Math.floor(n);
        if (d === 0) return 0;
        if (d <= 31) return d;
        if (d <= 59) return d - 31;
        if (d === 60) return 29;
        return new Date(Date.UTC(1899, 11, 30) + d * 86400000).getUTCDate();
      },
      // WORKDAY(시작일, 일수, [휴일]): 토·일과 휴일을 건너뛴다. 시작일이 비면 에러 값(던지지 않음)
      WORKDAY: (start, days, holidays) => {
        const s0 = scalar(start); const n0 = scalar(days);
        if (isErr(s0)) return s0; if (isErr(n0)) return n0;
        const startNum = Number(s0); const step = Math.trunc(Number(n0));
        if (s0 === null || s0 === '' || !Number.isFinite(startNum) || !Number.isFinite(step)) return FormulaError.VALUE;
        const skip = new Set(holidays === undefined || holidays === null ? [] : argValues(holidays).values.filter((v) => typeof v === 'number').map((v) => Math.floor(v)));
        const weekend = (serial) => { const wd = new Date(Date.UTC(1899, 11, 30) + serial * 86400000).getUTCDay(); return wd === 0 || wd === 6; };
        let cur = Math.floor(startNum); let left = Math.abs(step); const dir = step < 0 ? -1 : 1;
        while (left > 0) { cur += dir; if (!weekend(cur) && !skip.has(cur)) left -= 1; }
        return cur;
      },
      TODAY: () => {
        const base = Date.UTC(1899, 11, 30);
        const d = Date.UTC(today.getFullYear(), today.getMonth(), today.getDate());
        return Math.round((d - base) / 86400000);
      },
    },
  });

  let computed = 0; let failed = 0;
  for (const sheet of sheets) {
    const cells = books.get(sheet.name);
    if (!cells) continue;
    const results = new Map();
    for (const cell of cells.values()) {
      if (!cell.formula) continue;
      const result = evalCell(sheet.name, cell.row, cell.col);
      results.set(`${numToCol(cell.col)}${cell.row}`, result);
    }
    if (results.size === 0) continue;
    let xml = sheetXmls.get(sheet.name);
    xml = xml.replace(/<c\b([^>]*?)>(<f\b[\s\S]*?)<\/c>/g, (whole, attrs, inner) => {
      const ref = attr(attrs, 'r');
      if (!results.has(ref)) return whole;
      let result = results.get(ref);
      const fPart = /^<f\b[^>]*?(?:\/>|>[\s\S]*?<\/f>)/.exec(inner);
      if (!fPart) return whole;
      const cleanAttrs = attrs.replace(/\s+t="[^"]*"/, '');
      if (result && typeof result === 'object' && !isErr(result) && 'value' in result) result = result.value;
      if (result && result.failed) { failed += 1; return `<c${cleanAttrs}>${fPart[0]}</c>`; }
      computed += 1;
      if (result === null || result === undefined) return `<c${cleanAttrs}>${fPart[0]}<v>0</v></c>`;
      if (isErr(result)) return `<c${cleanAttrs} t="e">${fPart[0]}<v>${encodeXml(errText(result))}</v></c>`;
      if (typeof result === 'boolean') return `<c${cleanAttrs} t="b">${fPart[0]}<v>${result ? 1 : 0}</v></c>`;
      if (typeof result === 'number') {
        if (!Number.isFinite(result)) return `<c${cleanAttrs} t="e">${fPart[0]}<v>#NUM!</v></c>`;
        return `<c${cleanAttrs}>${fPart[0]}<v>${result}</v></c>`;
      }
      return `<c${cleanAttrs} t="str">${fPart[0]}<v>${encodeXml(String(result))}</v></c>`;
    });
    zip.file(sheet.path, xml);
  }
  return { computed, failed };
}
