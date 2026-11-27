// 사용: node sampling-log.cjs <outdir>   → 통합시료채취기록부에 견본을 넣고 회사 양식으로 추출 (+ 넣은 값 저장)
const { chromium } = require('playwright');
const fs = require('fs');
const out = process.argv[2];
let seed = Number(process.env.SEED || 1); const rnd = () => { seed = (seed * 1103515245 + 12345) % 2147483648; return seed / 2147483648; };
const r = (lo, hi, dec) => (lo + rnd() * (hi - lo)).toFixed(dec);
// 줄 번호 → 종류: 입자 샘플러(먼지·금속·수은), 가스 샘플러(흡수액), 음압상자·테들러백, 그 밖
// (thc: 현장측정이라 채취유량 없음. 23·31·53행은 양식의 일차별 시트에 시험검사기간이 빠져 있던 줄, 33·36행은 기간이 따로 없던 줄)
const ROWS = { 13: 'particle', 17: 'gas', 18: 'gas', 19: 'gas', 20: 'gas', 21: 'gas', 23: 'gas', 24: 'gas-liquid', 25: 'particle', 31: 'particle', 32: 'particle', 33: 'particle', 34: 'particle', 35: 'gas', 36: 'gas', 37: 'gas', 38: 'gas', 40: 'thc', 41: 'plain', 42: 'box', 43: 'box', 53: 'box', 57: 'box', 58: 'box' };
const log = {
  company: '견본산업(주)', industry: '폐기물 소각', facility: '소각로 1호기', stackName: 'S1', siteAbbr: 'GB',
  particleSamplerId: '', particleSamplerYd: '0.991', particleSamplerYdManual: true, gasSamplerId: '', gasSamplerYd: '1.0127', gasSamplerYdManual: true,
  days: [{ date: '2026-11-20', weather: '맑음', temp: '14.5', humidity: '48', pressure: '758.3', windDir: '북서', windSpeed: '2.4' },
    { date: '2026-11-21', weather: '흐림', temp: '11.2', humidity: '61', pressure: '761.9', windDir: '서', windSpeed: '3.1' }],
  rows: {},
};
Object.entries(ROWS).forEach(([row, kind], i) => {
  const start = Number(r(100, 900, 2)); const end = start + Number(r(8, 60, 2));
  log.rows[row] = { requested: true, dayIndex: i % 2, start: `1${i % 6}:05`, end: `1${i % 6}:45`, meterStart: start.toFixed(2), meterEnd: end.toFixed(2),
    temp: i % 6 === 5 ? '' : r(12, 30, 1), pressureH2O: i % 3 === 2 ? '' : r(-40, 40, 1), ...(kind === 'gas-liquid' ? { collector: '흡수액' } : {}) };
});
// 비소(가스상, 33행)는 비소(입자상, 32행)와 한 번에 채취한다 → 일차·시각·가스미터 값·온도·압력이 같다
log.rows[33] = { ...log.rows[32] };
(async () => {
  const b = await chromium.launch(); const ctx = await b.newContext({ viewport: { width: 1300, height: 1000 }, acceptDownloads: true }); const p = await ctx.newPage();
  p.on('dialog', (d) => { console.log('DIALOG', d.message().slice(0, 120)); d.accept(); });
  p.on('pageerror', (e) => console.log('PAGEERR', e.message));
  await p.goto('http://localhost:4179/'); await p.waitForTimeout(3000);
  await p.getByText('user 로그인', { exact: true }).click(); await p.waitForTimeout(1200);
  await p.goto('http://localhost:4179/#sheet=samplinglog'); await p.waitForTimeout(2500);
  await p.evaluate((v) => window.__stackpilotSetSamplingLogForTest(v), log); await p.waitForTimeout(1000);
  const dl = p.waitForEvent('download', { timeout: 60000 }).catch(() => null);
  await p.evaluate(() => window.__stackpilotExportSamplingLogForTest());
  const d = await dl;
  if (d) { await d.saveAs(`${out}/samplinglog.xlsx`); console.log('samplinglog ok'); } else console.log('samplinglog NO DOWNLOAD');
  const stats = await p.evaluate(() => window.__stackpilotSamplingLogStatsForTest());
  fs.writeFileSync(`${out}/samplinglog.input.json`, JSON.stringify({ log, kinds: ROWS, stats }, null, 1));
  await b.close();
})().catch((e) => { console.log('ERR', e.message); process.exit(1); });
