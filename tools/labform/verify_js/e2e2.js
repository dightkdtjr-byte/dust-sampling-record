// 검증 방식 7: 프로그램에서 기준을 무작위로 바꾸고, 프로그램 판정과 내려받은 완성본(엑셀)의 판정을 견줄 자료를 만든다
const { chromium } = require('playwright'); const fs = require('fs'); const path = require('path');
(async () => {
  const [file, outDir] = process.argv.slice(2);
  const browser = await chromium.launch(); const ctx = await browser.newContext({ acceptDownloads: true });
  const page = await ctx.newPage(); const errors = []; page.on('pageerror', e => errors.push(String(e)));
  await page.goto('file://' + file); await page.waitForSelector('#home');
  const cases = await page.evaluate(() => {
    let seed = 20261126; const rnd = () => (seed = (seed * 1103515245 + 12345) % 2147483648) / 2147483648;
    const out = [];
    for (const r of DATA) {
      const final = 1 + Math.floor(rnd() * 4), calc = final + 1;
      const e = { calc, final };
      if (!r.linked) { const d = Math.max(final - (rnd() < 0.3 ? 0 : 1), 0); const m = Number(((1 + Math.floor(rnd() * 9)) * Math.pow(10, -Math.max(d, 1))).toFixed(Math.max(d, 1) + 1)); e.mdl = m; e.loq = Number((m * 3).toPrecision(2)); }
      edits[r.file + '|' + r.row] = e;
    }
    save();
    for (const r of DATA) {
      const c = cur(r); const mdl = Number(mdlValue(c)) || 0.1; const u = Math.pow(10, -judgeDigits(c));
      for (const raw of [mdl, mdl - 0.4 * u, mdl - 0.6 * u, mdl * (1 + rnd() * 30), rnd() * mdl * 3]) {
        const x = Number(Math.max(0, raw).toFixed(Number(c.calc) + 3));
        const L = Number((mdl * (1 + rnd() * 20)).toFixed(Number(c.final)));
        const permit = rnd() < 0.2 ? '해당없음' : String(L);
        const ev = evaluate(c, x, permit);
        out.push({ file: r.file, row: r.row, sheet: r.sheet, result: r.result, x, permit, program: ev ? ev.text : '' });
      }
    }
    render(); return out;
  });
  const [dl] = await Promise.all([page.waitForEvent('download', { timeout: 180000 }), page.click('#btnAll')]);
  const zp = path.join(outDir, 'golden.zip'); await dl.saveAs(zp);
  fs.writeFileSync(path.join(outDir, 'program_cases.json'), JSON.stringify(cases));
  console.log(JSON.stringify({ cases: cases.length, errors }));
  await browser.close();
})();
