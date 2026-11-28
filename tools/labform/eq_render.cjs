// 계산식 그림 만들기: KaTeX(LaTeX 조판)로 그려 4배 해상도 PNG로 찍는다
//   node tools/labform/eq_render.cjs <specs.json> <출력 폴더> <katex dist 폴더>
//   specs.json: [{id, tex}] → <출력 폴더>/<id>.png, sizes.json {id: [너비, 높이] (화면 점, 배율 1)}
const fs = require('fs'), path = require('path');
const { chromium } = require('playwright');
(async () => {
  const [specsPath, out, kdist] = process.argv.slice(2);
  const specs = JSON.parse(fs.readFileSync(specsPath, 'utf8'));
  const katex = fs.readFileSync(path.join(kdist, 'katex.min.js'), 'utf8');
  let css = fs.readFileSync(path.join(kdist, 'katex.min.css'), 'utf8');
  css = css.replace(/url\(fonts\//g, 'url(file://' + path.join(kdist, 'fonts') + '/');
  const browser = await chromium.launch();
  const page = await browser.newPage({ deviceScaleFactor: 4 });
  await page.setContent(`<!doctype html><html><head><meta charset="utf-8"><style>${css}
    body{margin:0;background:#fff}
    .eq{display:inline-block;padding:3px 4px;background:#fff;color:#000;font-size:20px}
    .katex{font-size:1em}
    .katex .text, .katex .mord.text{font-family:"Malgun Gothic","NanumBarunGothic",sans-serif}
    .lg{display:inline-block;padding:2px 4px;background:#fff;color:#111;font-family:"Malgun Gothic","NanumBarunGothic",sans-serif;font-size:15.5px;line-height:1.42}
    .lg table{border-collapse:collapse}
    .lg td{padding:0;vertical-align:baseline}
    .lg td.s{white-space:nowrap;text-align:left;padding-right:2px}
    .lg td.c{padding:0 7px 0 3px;color:#333}
    .lg td.d{word-break:keep-all;overflow-wrap:break-word}
    .lg .katex{font-size:1.1em}
    .lg sup,.lg sub{font-size:0.7em;line-height:0}
    .lg sup{vertical-align:0.55em}.lg sub{vertical-align:-0.25em}
  </style></head><body><div id="host"></div><script>${katex}</script></body></html>`);
  await page.evaluate(() => document.fonts.ready);
  const sizesPath = path.join(out, 'sizes.json');
  const sizes = fs.existsSync(sizesPath) ? JSON.parse(fs.readFileSync(sizesPath, 'utf8')) : {};
  for (const s of specs) {
    if (s.legend) {
      await page.evaluate(({ legend, maxw }) => {
        const h = document.getElementById('host');
        h.innerHTML = '';
        const box = document.createElement('div'); box.className = 'lg'; box.id = 'eq';
        const t = document.createElement('table'); box.appendChild(t);
        for (const [sym, desc] of legend) {
          const tr = document.createElement('tr');
          const a = document.createElement('td'); a.className = 's'; katex.render(sym, a, { throwOnError: true, strict: false });
          const c = document.createElement('td'); c.className = 'c'; c.textContent = ':';
          const d = document.createElement('td'); d.className = 'd';
          for (const part of desc.split(/([³²₂⁻])/)) {   // 위·아래첨자는 글자 대신 조판으로
            const map = { '³': ['sup', '3'], '²': ['sup', '2'], '₂': ['sub', '2'], '⁻': ['sup', '−'] };
            if (map[part]) { const e = document.createElement(map[part][0]); e.textContent = map[part][1]; d.appendChild(e); }
            else if (part) d.appendChild(document.createTextNode(part));
          }
          tr.append(a, c, d); t.appendChild(tr);
        }
        h.appendChild(box);
        // 한 줄로 들어가면 그대로, 넘치면 최대 너비 안에서 줄바꿈
        if (box.getBoundingClientRect().width > maxw) box.style.width = maxw + 'px';
      }, s);
    } else {
      await page.evaluate(({ tex }) => {
        const h = document.getElementById('host');
        h.innerHTML = '<div class="eq" id="eq"></div>';
        katex.render(tex, document.getElementById('eq'), { displayMode: false, throwOnError: true, strict: false });
      }, s);
    }
    await page.evaluate(() => document.fonts.ready);
    const el = await page.$('#eq');
    const box = await el.boundingBox();
    sizes[s.id] = [Math.ceil(box.width), Math.ceil(box.height)];
    await el.screenshot({ path: path.join(out, s.id + '.png'), omitBackground: false });
  }
  fs.writeFileSync(path.join(out, 'sizes.json'), JSON.stringify(sizes, null, 1));
  await browser.close();
})().catch(e => { console.error(e); process.exit(1); });
