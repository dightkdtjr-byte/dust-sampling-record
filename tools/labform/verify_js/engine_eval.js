// 검증 방식 5-2: LibreOffice가 아닌 다른 엔진(fast-formula-parser)으로 결과 수식을 직접 계산해 기대값과 견준다.
//   node engine_eval.js <audit 결과 폴더>
// 결과 칸·기준 시트·기준을 가리키는 칸·표준산소 식은 이 엔진으로 다시 계산하고, 그 밖의 칸은 LibreOffice 계산값을 쓴다.
const FP = require('fast-formula-parser');
const { FormulaHelpers: H, Types: T, FormulaError } = FP;
const fs = require('fs'), path = require('path');
const dir = process.argv[2];
const cases = JSON.parse(fs.readFileSync(path.join(dir, 'cases.json'), 'utf8'));
const maps = {};
const col = s => s.split('').reduce((a, ch) => a * 26 + ch.charCodeAt(0) - 64, 0);
const colL = n => { let s = ''; while (n > 0) { const m = (n - 1) % 26; s = String.fromCharCode(65 + m) + s; n = Math.floor((n - 1) / 26); } return s; };
// 이 엔진에 없는 엑셀 함수 둘: 엑셀 정의대로
const functions = {
  SUBSTITUTE: (text, oldT, newT) => { text = H.accept(text, T.STRING); oldT = H.accept(oldT, T.STRING); newT = H.accept(newT, T.STRING); return oldT === '' ? text : text.split(oldT).join(newT); },
  VALUE: (text) => { const v = H.accept(text); if (typeof v === 'number') return v; const s = String(v).trim(); if (s === '' || !/^[-+]?(\d+\.?\d*|\.\d+)([eE][-+]?\d+)?$/.test(s)) throw FormulaError.VALUE; return Number(s); },
};
// MAX는 이 엔진에서 칸 참조를 받는 함수 자리에 둔다
const functionsNeedContext = { MAX: (context, ...params) => { const vals = []; const walk = x => { if (Array.isArray(x)) x.forEach(walk); else if (x && typeof x === 'object' && 'value' in x) walk(x.value); else if (typeof x === 'number') vals.push(x); }; params.forEach(walk); return vals.length ? Math.max(...vals) : 0; } };
let ok = 0, bad = 0; const fails = [];
for (const c of cases) {
  const file = c.file.replace(/\.xls[xm]$/, '');
  maps[file] = maps[file] || JSON.parse(fs.readFileSync(path.join(dir, `cells_${file}.json`), 'utf8'));
  const cells = maps[file];
  const over = {}; over[`${c.rsh}!${c.rc}`] = { f: null, v: c.x };
  if (c.pc) over[`${c.sheet}!${c.pc}`] = { f: null, v: c.permit };
  const memo = {};
  // 이 엔진은 계산 중에 같은 계산기를 다시 부를 수 없어서, 안쪽 칸은 새 계산기로 계산한다
  const makeParser = () => new FP({
    functions, functionsNeedContext,
    onCell: ({ sheet, row, col: cc }) => get(sheet, `${colL(cc)}${row}`),
    onRange: ({ sheet, from, to }) => { const out = []; for (let r = from.row; r <= to.row; r++) { const line = []; for (let k = from.col; k <= to.col; k++) line.push(get(sheet, `${colL(k)}${r}`)); out.push(line); } return out; },
  });
  const parser = makeParser();
  function get(sheet, ref) {
    const key = `${sheet}!${ref}`;
    if (key in memo) return memo[key];
    const cell = over[key] || cells[key];
    let v = null;
    if (cell) {
      const f = cell.f;
      const evalHere = f && (over[key] || sheet === '기준' || f.includes("'기준'") || f.includes('VALUE(MID(TRIM(SUBSTITUTE('));
      if (evalHere) { memo[key] = null; v = makeParser().parse(f.slice(1), { sheet, row: Number(ref.match(/\d+/)[0]), col: col(ref.match(/[A-Z]+/)[0]) }); if (v && v.error) v = v; }
      else v = cell.v;
    }
    memo[key] = v; return v;
  }
  let got;
  try { got = parser.parse(c.formula.slice(1), { sheet: c.sheet, row: Number(c.result.match(/\d+/)[0]), col: col(c.result.match(/[A-Z]+/)[0]) }); }
  catch (e) { got = 'ERR ' + e.message; if (!globalThis.__shown) { globalThis.__shown = 1; console.log(e.stack.split('\n').slice(0,8).join('\n')); } }
  if (got && typeof got === 'object' && got.error) got = got.error;
  const g = String(got ?? '');
  if (g === c.expect) ok++; else { bad++; fails.push(`${c.item} ${c.result} x=${c.x} 허가=${c.permit}: 엔진 ${JSON.stringify(g)} ≠ 기대 ${JSON.stringify(c.expect)}`); }
}
console.log(`다른 엔진(fast-formula-parser) 계산: ${ok + bad}건 중 같음 ${ok}, 다름 ${bad}`);
fails.slice(0, 40).forEach(f => console.log('   ✖', f));
