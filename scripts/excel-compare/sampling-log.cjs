// 사용: node sampling-log.cjs <outdir>   → 통합시료채취기록부에 견본을 넣고 회사 양식으로 추출 (+ 넣은 값 저장)
const { chromium } = require('playwright');
const fs = require('fs');
const out = process.argv[2];
let seed = Number(process.env.SEED || 1); const rnd = () => { seed = (seed * 1103515245 + 12345) % 2147483648; return seed / 2147483648; };
const r = (lo, hi, dec) => (lo + rnd() * (hi - lo)).toFixed(dec);
// 줄 번호 → 종류: 입자 샘플러(먼지·금속·수은), 가스 샘플러(흡수액), 음압상자·테들러백, 그 밖
const ROWS = { 13: 'particle', 19: 'gas', 20: 'gas', 21: 'gas', 24: 'gas-liquid', 25: 'particle', 34: 'particle', 35: 'gas', 38: 'gas', 41: 'plain', 42: 'box', 43: 'box' };
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
    temp: r(12, 30, 1), pressureH2O: i % 3 === 2 ? '' : r(-40, 40, 1), ...(kind === 'gas-liquid' ? { collector: '흡수액' } : {}) };
});
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
  fs.writeFileSync(`${out}/samplinglog.input.json`, JSON.stringify({ log, kinds: ROWS }, null, 1));
  await b.close();
})().catch((e) => { console.log('ERR', e.message); process.exit(1); });
