// 사용: node exp2.cjs <outdir>   → 여러 조건의 견본을 넣고 회사 양식으로 추출 (+ 화면 계산값 저장)
const { chromium } = require('playwright');
const fs = require('fs');
const out = process.argv[2]; const only = process.argv.slice(3);
let seed = Number(process.env.SEED || 1); const rnd = () => { seed = (seed * 1103515245 + 12345) % 2147483648; return seed / 2147483648; };
const r = (lo, hi, dec) => (lo + rnd() * (hi - lo)).toFixed(dec);
// 조건: sheet, method(auto/imp2/impinger), shape(circle/horizontal/rect/small/big), sp(same/diff/blank), o2(표준산소 유무), rowsPer(지점당 줄 수)
const S = [
  { n: 'dust-auto-c2', sheet: 'dust', method: 'auto', shape: 'circle', sp: 'diff', o2: true, rowsPer: 2 },
  { n: 'dust-imp2-c5', sheet: 'dust', method: 'imp2', shape: 'big', sp: 'diff', o2: true, rowsPer: 1 },
  { n: 'dust-impinger-rect', sheet: 'dust', method: 'impinger', shape: 'rect', sp: 'blank', o2: false, rowsPer: 1 },
  { n: 'dust-auto-small', sheet: 'dust', method: 'auto', shape: 'small', sp: 'same', o2: true, rowsPer: 3 },
  { n: 'dust-auto-horiz', sheet: 'dust', method: 'auto', shape: 'horizontal', sp: 'diff', o2: false, rowsPer: 1 },
  { n: 'metal-auto-c2', sheet: 'metal', method: 'auto', shape: 'circle', sp: 'diff', o2: true, rowsPer: 2 },
  { n: 'metal-imp2-c5', sheet: 'metal', method: 'imp2', shape: 'big', sp: 'blank', o2: false, rowsPer: 1 },
  { n: 'mercury-auto-c2', sheet: 'mercury', method: 'auto', shape: 'circle', sp: 'diff', o2: true, rowsPer: 2 },
  { n: 'mercury-impinger-c5', sheet: 'mercury', method: 'impinger', shape: 'big', sp: 'same', o2: false, rowsPer: 1 },
  { n: 'pahs-auto-rect', sheet: 'pahs', method: 'auto', shape: 'rect', sp: 'diff', o2: true, rowsPer: 1 },
  { n: 'pahs-imp2-c2', sheet: 'pahs', method: 'imp2', shape: 'circle', sp: 'diff', o2: true, rowsPer: 2 },
  { n: 'flow-auto-c2', sheet: 'flow', method: 'auto', shape: 'circle', sp: 'same', o2: true, rowsPer: 1 },
  { n: 'flow-imp2-c5', sheet: 'flow', method: 'imp2', shape: 'big', sp: 'same', o2: false, rowsPer: 1 },
  { n: 'flow-impinger-rect', sheet: 'flow', method: 'impinger', shape: 'rect', sp: 'same', o2: true, rowsPer: 1 },
  // 측정점이 한 줄에 6개 이상인 큰 사각 굴뚝(가로 6칸 × 세로 4칸), 세로로 긴 사각 굴뚝(가로 2칸 × 세로 6칸)
  { n: 'dust-auto-rect6', sheet: 'dust', method: 'auto', shape: 'rect6', sp: 'diff', o2: true, rowsPer: 1 },
  { n: 'dust-auto-rectwide', sheet: 'dust', method: 'auto', shape: 'rectwide', sp: 'blank', o2: false, rowsPer: 2 },
  { n: 'flow-auto-rect6', sheet: 'flow', method: 'auto', shape: 'rect6', sp: 'same', o2: true, rowsPer: 1 },
  { n: 'flow-auto-rect8', sheet: 'flow', method: 'auto', shape: 'rect8', sp: 'same', o2: false, rowsPer: 1 },
];
const build = (sc) => {
  const geo = { circle: { total: '1.45', flange: '0.25', shape: 'circle', orient: 'vertical', pts: 2 }, big: { total: '4.95', flange: '0.25', shape: 'circle', orient: 'vertical', pts: 5 },
    small: { total: '0.75', flange: '0.25', shape: 'circle', orient: 'vertical', pts: 1 }, horizontal: { total: '2.35', flange: '0.25', shape: 'circle', orient: 'horizontal', pts: 3 },
    rect: { total: '1.75', flange: '0.25', shape: 'rect', orient: 'vertical', side: '1.3', pts: 3 },
    rect6: { total: '6.05', flange: '0.25', shape: 'rect', orient: 'vertical', side: '4.0', pts: 5 },
    rect8: { total: '8.25', flange: '0.25', shape: 'rect', orient: 'vertical', side: '7.2', pts: 5 },
    rectwide: { total: '1.75', flange: '0.25', shape: 'rect', orient: 'vertical', side: '5.6', pts: 2 } }[sc.shape];
  const pts = Array.from({ length: geo.pts }, (_, i) => { const sp = Number(r(-6, -1, 1)); const dp = Number(r(4, 12, 1)); return { id: i + 1, tp: (sp + dp).toFixed(1), sp: String(sp), dp: String(dp), ts: r(120, 180, 1) }; });
  const meters = [{ id: 0, pointNum: '시작', time: '0', stackTemp: '', dp: '', sp: '', pressure: '', volume: '1250.37', tmIn: '', tmOut: '', vacuum: '', impingerTemp: '' }];
  let vol = 1250.37, t = 0, id = 1;
  pts.forEach((p, i) => { for (let k = 0; k < sc.rowsPer; k++) { t += 7; vol += Number(r(90, 140, 2)); meters.push({ id: id++, pointNum: String(i + 1), time: String(t), stackTemp: r(120, 180, 1), dp: r(4, 12, 1),
    sp: sc.sp === 'diff' ? r(-9, -2, 1) : sc.sp === 'same' ? p.sp : '', pressure: '', volume: vol.toFixed(2), tmIn: r(18, 30, 1), tmOut: r(16, 28, 1), vacuum: r(60, 120, 0), impingerTemp: r(8, 18, 1) }); } });
  const flow = sc.sheet === 'flow';
  const d = {
    date: '2026-11-20', company: '견본산업(주)', location: '소각로 1호기 배출구', sampler: '홍길동', weather: '맑음', atmTemp: '14.5', atmPressure: '758.3',
    totalStackDepth: geo.total, flangeLength: geo.flange, stackDiameter: (Number(geo.total) - Number(geo.flange)).toFixed(2), stackShape: geo.shape, stackSide: geo.side || '', stackOrientation: geo.orient,
    pitotFactor: '0.837', standardO2: sc.o2 ? '12' : '',
    gasAnalyzer: [0, 1, 2].map((i) => ({ time: `10:${i}0`, o2: r(9, 13, 2), co2: r(6, 9, 2), co: r(5, 400, 0), nox: r(20, 60, 0), sox: r(1, 9, 0) })),
    moistureValues: sc.method === 'auto' ? [0, 1, 2, 3, 4].map(() => r(4, 9, 2)) : ['', '', '', '', ''], moistureMethod: sc.method,
    preImpinger: sc.method === 'auto' ? undefined : { impingers: [1, 2, 3, 4].map((i) => (i <= (sc.method === 'imp2' ? 2 : 4) ? { id: i, initial: r(40, 60, 2), final: '' } : { id: i, initial: '', final: '' })),
      meterStart: '100.25', meterEnd: sc.method === 'imp2' ? '110.83' : '161.42', meterTemp: '22.4', pressure: '757.1', orifice: '3.2', gasSamplerId: '', yd: '0.9953', wetMeter: sc.method === 'imp2' },
    impingers: [600, 610, 500, 700, 650, 640, 630].map((v, i) => ({ id: i + 1, initial: (v + Number(r(0, 5, 2))).toFixed(2), final: '' })),
    traverseTmIn: '24.3', traverseTmOut: '22.1', points: pts,
    targetVolume: '1000', planSamplingTime: '65', samplingStartTime: '10:05',
    leakTest: { device: 'O', impinger: 'O', samplingPart: 'O', meterBefore: '1249.80', meterAfter: '1249.80', note: '' },
    gasMeterFactor: '0.991', deltaHAt: '47.6079', gasMeters: meters, filterId: 'F-101', filterInitial: '351.24', filterFinal: '368.87', remarks: '견본 자료',
  };
  if (d.preImpinger) d.preImpinger.impingers.forEach((imp, i) => { if (imp.initial) imp.final = (Number(imp.initial) + [0.47, 0.13, 0.05, 0.31][i] * (sc.method === 'imp2' ? 1 : 6)).toFixed(2); });
  d.impingers.forEach((imp, i) => { imp.final = (Number(imp.initial) + [22.4, 9.3, 2.1, 8.6, 1.2, 0.7, 0.4][i]).toFixed(2); });
  if (flow) Object.assign(d, { gasMeters: undefined, filterId: '', filterInitial: '', filterFinal: '', impingers: undefined });
  if (sc.sheet === 'mercury' || sc.sheet === 'pahs') Object.assign(d, { filterId: '', filterInitial: '', filterFinal: '' });
  return d;
};
(async () => {
  const b = await chromium.launch(); const ctx = await b.newContext({ viewport: { width: 1300, height: 1000 }, acceptDownloads: true }); const p = await ctx.newPage();
  p.on('dialog', (d) => { console.log('DIALOG', d.message().slice(0, 120)); d.accept(); });
  p.on('pageerror', (e) => console.log('PAGEERR', e.message));
  await p.goto('http://localhost:4179/'); await p.waitForTimeout(3000);
  await p.getByText('user 로그인', { exact: true }).click(); await p.waitForTimeout(1200);
  for (const sc of S) {
    if (only.length && !only.includes(sc.n)) { build(sc); continue; }
    await p.goto('http://localhost:4179/#sheet=' + sc.sheet); await p.waitForTimeout(2500);
    const data = build(sc);
    await p.evaluate((d) => window.__stackpilotLoadForTest(d), JSON.parse(JSON.stringify(data))); await p.waitForTimeout(1200);
    if (sc.sheet !== 'flow') {
      // 실제 사용처럼: 노즐 세트에서 하나를 골라 그 노즐의 K를 적용한다
      const der0 = await p.evaluate(() => window.__stackpilotDerivedForTest());
      const cand = der0.full.nozzles.filter((n) => n.k > 0 && n.dH > 1 && n.dH < 60); const nz = cand[Math.floor(cand.length / 2)] || der0.full.nozzles[2];
      Object.assign(data, { usedNozzleNum: String(nz.num), nozzleDiameter: nz.d.toFixed(2), kFactor: String(nz.k) });
      await p.evaluate((d) => window.__stackpilotLoadForTest(d), JSON.parse(JSON.stringify(data))); await p.waitForTimeout(1200);
    }
    const dl = p.waitForEvent('download', { timeout: 60000 }).catch(() => null);
    await p.getByRole('button', { name: '현재 내용 엑셀로 추출' }).click();
    const d = await dl;
    if (d) { await d.saveAs(`${out}/${sc.n}.xlsm`); console.log(sc.n, 'ok'); } else console.log(sc.n, 'NO DOWNLOAD');
    const der = await p.evaluate(() => window.__stackpilotDerivedForTest());
    fs.writeFileSync(`${out}/${sc.n}.derived.json`, JSON.stringify(der, null, 1));
    fs.writeFileSync(`${out}/${sc.n}.input.json`, JSON.stringify({ sc, data }, null, 1));
  }
  await b.close();
})().catch((e) => { console.log('ERR', e.message); process.exit(1); });
