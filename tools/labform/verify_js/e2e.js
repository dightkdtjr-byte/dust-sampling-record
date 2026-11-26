// 프로그램 브라우저 시험: 처음 화면 → 기록부 화면 → 기준 수정·원문 대조 → 결과 계산 → 원문 근거 → 기록부 하나·전체 내려받기 → 휴대폰 화면 넘침
const { chromium } = require('playwright');
const path = require('path');
(async () => {
  const file = process.argv[2], outDir = process.argv[3];
  const browser = await chromium.launch(); const ctx = await browser.newContext({ acceptDownloads: true, viewport: { width: 1440, height: 1000 } });
  const page = await ctx.newPage(); const errors = [];
  page.on('pageerror', e => errors.push(String(e))); page.on('console', m => { if (m.type() === 'error') errors.push(m.text()); });
  await page.goto('file://' + file); await page.waitForSelector('#home');
  const res = {};
  res.summary = await page.$eval('#summary', e => e.innerText.replace(/\s+/g, ' '));
  res.cards = await page.$$eval('.card', t => t.length);
  const open = async f => { await page.evaluate(f => { location.hash = '#f=' + encodeURIComponent(f); }, f); await page.waitForSelector('.item'); };
  const checksOf = k => page.$eval(`.item[data-k="${k}"] .checks`, e => [...e.querySelectorAll('li')].filter(li => !li.querySelector('.pill.ok')).map(li => li.innerText.replace(/\s+/g, ' ')).join(' / '));
  // 아세트알데하이드 최종 4 → 부적합, 3으로 되돌림
  await open('시험기록부_아세트알데하이드.xlsx');
  const k1 = '시험기록부_아세트알데하이드.xlsx|2';
  await page.fill(`.item[data-k="${k1}"] [data-f="final"]`, '4'); await page.waitForTimeout(600);
  res.ald_badges = await checksOf(k1);
  await page.fill(`.item[data-k="${k1}"] [data-f="final"]`, '3'); await page.waitForTimeout(600);
  res.ald_back = await checksOf(k1);
  // 벤젠 방법검출한계 0.05 → 원문과 다름
  await open('시험기록부_VOCs.xlsm');
  const k2 = await page.$$eval('.item', t => t.find(x => x.querySelector('h2').innerText.trim() === '벤젠').dataset.k);
  await page.fill(`.item[data-k="${k2}"] [data-f="mdl"]`, '0.05'); await page.waitForTimeout(600);
  res.benz_badges = await checksOf(k2);
  await page.fill(`.item[data-k="${k2}"] [data-f="mdl"]`, '0.03'); await page.waitForTimeout(600);
  // 결과 계산 (벤젠: 최종 2, 방법검출한계 0.03)
  const tr = w => `.item[data-k="${k2}"] [data-try="${w}"]`;
  await page.fill(tr('raw'), '0.0295'); await page.fill(tr('permit'), '0.05(12)'); await page.waitForTimeout(700);
  res.try1 = await page.$eval(`.item[data-k="${k2}"] .res`, e => e.innerText);
  await page.fill(tr('raw'), '0.06'); await page.waitForTimeout(700);
  res.try2 = await page.$eval(`.item[data-k="${k2}"] .res`, e => e.innerText);
  await page.fill(tr('raw'), '0.0249'); await page.waitForTimeout(700);
  res.try3 = await page.$eval(`.item[data-k="${k2}"] .res`, e => e.innerText);
  // 원문 근거 창
  await page.click(`.item[data-k="${k2}"] .item-h [data-ev]`); await page.waitForTimeout(400);
  res.evidence = await page.$$eval('#modal.open figure img', t => t.map(i => i.naturalWidth > 100).filter(Boolean).length);
  res.evidence_titles = await page.$$eval('#modal.open figcaption b', t => t.map(x => x.innerText));
  await page.screenshot({ path: path.join(outDir, 'evidence.png') });
  await page.keyboard.press('Escape');
  // 수은 정량한계 바꾸고 기록부 하나 내려받기
  await open('시험기록부_수은화합물.xlsx');
  const k3 = '시험기록부_수은화합물.xlsx|2';
  await page.fill(`.item[data-k="${k3}"] [data-f="loq"]`, '0.0006'); await page.waitForTimeout(600);
  const [d1] = await Promise.all([page.waitForEvent('download', { timeout: 120000 }), page.click('.fhead [data-dl]')]);
  const p1 = path.join(outDir, d1.suggestedFilename()); await d1.saveAs(p1); res.download_one = p1;
  // 전체 내려받기
  await page.evaluate(() => { location.hash = ''; }); await page.waitForSelector('#btnAll');
  const [dl] = await Promise.all([page.waitForEvent('download', { timeout: 180000 }), page.click('#btnAll')]);
  const p = path.join(outDir, dl.suggestedFilename()); await dl.saveAs(p); res.download = p;
  res.toast = await page.$eval('#toast', e => e.innerText);
  await page.screenshot({ path: path.join(outDir, 'program.png'), fullPage: false });
  // 휴대폰 너비: 가로 넘침·잘린 글자
  const m = await browser.newPage({ viewport: { width: 375, height: 800 } });
  m.on('pageerror', e => errors.push(String(e)));
  res.overflow = [];
  for (const h of ['', '#f=' + encodeURIComponent('시험기록부_비소화합물.xlsx'), '#f=' + encodeURIComponent('시험기록부_VOCs.xlsm')]) {
    await m.goto('file://' + file + h); await m.waitForTimeout(500);
    res.overflow.push(...await m.evaluate(() => { const r = []; if (document.documentElement.scrollWidth > innerWidth + 1) r.push('가로 넘침 ' + document.documentElement.scrollWidth);
      document.querySelectorAll('.pill,.btn,.chip,.msg,h1,h2,h3,label,.kicker,.facts span,td,th').forEach(e => { const pr = e.parentElement.getBoundingClientRect(), b = e.getBoundingClientRect(); if (b.width && b.right > pr.right + 2 && !e.closest('.seg')) r.push((e.className || e.tagName) + ': ' + e.textContent.trim().slice(0, 20)); });
      document.querySelectorAll('input[placeholder]').forEach(e => { const c = document.createElement('canvas').getContext('2d'); c.font = getComputedStyle(e).font; if (c.measureText(e.placeholder).width > e.clientWidth - 16) r.push('placeholder 잘림: ' + e.placeholder); });
      return r; }));
  }
  res.errors = errors;
  console.log(JSON.stringify(res, null, 1)); await browser.close();
})();
