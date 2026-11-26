const { chromium } = require('playwright');
const fs = require('fs'); const path = require('path');
(async () => {
  const file = process.argv[2], outDir = process.argv[3];
  const browser = await chromium.launch(); const ctx = await browser.newContext({ acceptDownloads: true, viewport: { width: 1500, height: 1000 } });
  const page = await ctx.newPage(); const errors = [];
  page.on('pageerror', e => errors.push(String(e))); page.on('console', m => { if (m.type() === 'error') errors.push(m.text()); });
  await page.goto('file://' + file); await page.waitForSelector('#summary .pill');
  const res = {};
  res.summary = await page.$eval('#summary', e => e.innerText.replace(/\n/g, ' '));
  await page.click('#btnExpand');
  res.rows = await page.$$eval('tr[data-k]:not(.detail)', t => t.length);
  // 2. 아세트알데히드 최종 4
  const k1 = '시험기록부_아세트알데하이드.xlsx|2';
  const sel = f => `tr[data-k="${k1}"]:not(.detail) [data-f="${f}"]`;
  await page.fill(sel('final'), '4'); await page.waitForTimeout(600);
  res.ald_badges = await page.$eval(`tr[data-k="${k1}"]:not(.detail) .badges`, e => e.innerText.replace(/\n/g, ' / '));
  await page.fill(sel('final'), '3'); await page.waitForTimeout(600);
  res.ald_back = await page.$eval(`tr[data-k="${k1}"]:not(.detail) .badges`, e => e.innerText.replace(/\n/g, ' / '));
  // 3. 벤젠 MDL 0.05
  const k2 = [...await page.$$eval('tr[data-k]:not(.detail)', t => t.map(x => [x.dataset.k, x.querySelector('.item').innerText]))].find(x => x[1].startsWith('벤젠'))[0];
  await page.fill(`tr[data-k="${k2}"]:not(.detail) [data-f="mdl"]`, '0.05'); await page.waitForTimeout(600);
  res.benz_badges = await page.$eval(`tr[data-k="${k2}"]:not(.detail) .badges`, e => e.innerText.replace(/\n/g, ' / '));
  await page.fill(`tr[data-k="${k2}"]:not(.detail) [data-f="mdl"]`, '0.03'); await page.waitForTimeout(600);
  // 4. 시험해 보기 (벤젠: 최종 2, MDL 0.03)
  await page.click(`tr[data-k="${k2}"]:not(.detail) [data-act="detail"]`); await page.waitForTimeout(300);
  await page.fill(`tr.detail[data-k="${k2}"] [data-try="raw"]`, '0.0295'); await page.fill(`tr.detail[data-k="${k2}"] [data-try="permit"]`, '0.05(12)'); await page.waitForTimeout(700);
  res.try1 = await page.$eval(`tr.detail[data-k="${k2}"] .res`, e => e.innerText);
  await page.fill(`tr.detail[data-k="${k2}"] [data-try="raw"]`, '0.06'); await page.waitForTimeout(700);
  res.try2 = await page.$eval(`tr.detail[data-k="${k2}"] .res`, e => e.innerText);
  await page.fill(`tr.detail[data-k="${k2}"] [data-try="raw"]`, '0.0249'); await page.waitForTimeout(700);
  res.try3 = await page.$eval(`tr.detail[data-k="${k2}"] .res`, e => e.innerText);
  // 5. 수은 정량한계 바꾸고 골든 카피 내려받기
  const k3 = '시험기록부_수은화합물.xlsx|2';
  await page.fill(`tr[data-k="${k3}"]:not(.detail) [data-f="loq"]`, '0.0006'); await page.waitForTimeout(600);
  const [dl] = await Promise.all([page.waitForEvent('download', { timeout: 120000 }), page.click('#btnGolden')]);
  const p = path.join(outDir, dl.suggestedFilename()); await dl.saveAs(p); res.download = p;
  res.toast = await page.$eval('#toast', e => e.innerText);
  await page.screenshot({ path: path.join(outDir, 'program.png'), fullPage: false });
  res.errors = errors;
  console.log(JSON.stringify(res, null, 1)); await browser.close();
})();
