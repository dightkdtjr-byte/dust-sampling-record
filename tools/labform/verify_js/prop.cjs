const { chromium } = require('playwright'); const fs = require('fs');
(async () => {
  const [file, casesF, outF] = process.argv.slice(2); const cases = JSON.parse(fs.readFileSync(casesF));
  const b = await chromium.launch(); const p = await b.newPage(); await p.goto('file://' + file); await p.waitForSelector('#home');
  const res = await p.evaluate(cs => cs.map(c => { const row = Object.assign({}, DATA[0], { calc: c.calc, final: c.final, mdl: c.mdl === null ? '-' : c.mdl, linked: null }); const e = evaluate(row, c.raw, ''); return e ? e.text : null; }), cases);
  fs.writeFileSync(outF, JSON.stringify(res)); await b.close();
})();
