/*
 * StackPilot 배경 (하늘 · 해/달 · 별 · 구름 · 공장 풍경 · 연기 애니메이션)
 * React 없이 쓰는 단독 스크립트. 사용법:
 *   <div id="bg"></div>
 *   <script src="stackpilot-skyline.js"></script>
 *   <script>StackpilotSkyline.mount(document.getElementById('bg'), { mode: 'day' });</script>
 * mode: 'day' | 'morning' | 'night' | 'auto'(시간대에 따라 자동)
 * 컨테이너는 position: relative(또는 fixed)로 두고 크기를 지정하세요. 반환값의 setMode(mode)로 바꿀 수 있습니다.
 */
(function (global) {
  'use strict';

const seededNoise = (seed) => {
  const x = Math.sin(seed * 12.9898) * 43758.5453;
  return x - Math.floor(x);
};

const SKYLINE_STARS = Array.from({ length: 46 }, (_, idx) => ({
  id: `star-${idx}`,
  left: `${2 + ((idx * 17) % 96)}%`,
  top: `${4 + ((idx * 13) % 44)}%`,
  size: 1.2 + (idx % 4) * 0.55,
  duration: 2.4 + (idx % 5) * 0.65,
  delay: (idx % 7) * 0.35,
}));

const SKYLINE_CLOUDS = [
  { id: 'cl-1', top: '6%', scale: 1.25, duration: 170, delay: -20, opacity: 0.95 },
  { id: 'cl-2', top: '14%', scale: 0.8, duration: 130, delay: -75, opacity: 0.8 },
  { id: 'cl-3', top: '22%', scale: 1.05, duration: 200, delay: -140, opacity: 0.9 },
  { id: 'cl-4', top: '10%', scale: 0.6, duration: 115, delay: -40, opacity: 0.7 },
  { id: 'cl-5', top: '30%', scale: 0.9, duration: 185, delay: -110, opacity: 0.75 },
  { id: 'cl-6', top: '18%', scale: 1.4, duration: 230, delay: -180, opacity: 0.85 },
  { id: 'cl-7', top: '26%', scale: 0.7, duration: 150, delay: -10, opacity: 0.65 },
];

const SKYLINE_PALETTES = {
  day: {
    hillFar: ['#dbe8f6', '#adc6e4'], hillMid: ['#b3cdbf', '#73a08c'], hillNear: ['#86b88f', '#4b8762'],
    ground: ['#5f9f71', '#2b5740'], road: 'rgba(255,255,255,0.35)', haze: 'rgba(236, 246, 255, 0.75)',
    tower: ['#f5f9fd', '#bdd0e4', '#7e97b6'], towerEdge: 'rgba(40, 60, 90, 0.18)',
    factory: ['#e9e2da', '#bcb0a6', '#7f746b'], roof: '#8c7f76',
    cooling: ['#f7f9fb', '#d2dbe5', '#9aa7b7'], stack: ['#f3f6fa', '#c2cedd', '#8495ad'], stackCap: '#7b8ca4', band: '#d9534f',
    window: 'rgba(40, 60, 95, 0.25)', windowLit: null, smoke: ['rgba(255,255,255,0.95)', 'rgba(205,215,228,0)'],
    tree: ['#93cb84', '#3d7f4e'], trunk: '#7a5a44', beacon: null, cityGlow: null,
  },
  morning: {
    hillFar: ['#ebdff0', '#c6b4db'], hillMid: ['#d3c1b2', '#95818a'], hillNear: ['#a9a97c', '#64704f'],
    ground: ['#84976a', '#3f4c39'], road: 'rgba(255, 240, 220, 0.35)', haze: 'rgba(255, 226, 190, 0.7)',
    tower: ['#fff4ea', '#dcc0b8', '#957f8b'], towerEdge: 'rgba(90, 60, 70, 0.18)',
    factory: ['#f3e1d2', '#c9aa9b', '#826d69'], roof: '#8f7570',
    cooling: ['#fff7ee', '#e4cfc3', '#a8918d'], stack: ['#fff5ea', '#dcc3b7', '#968085'], stackCap: '#9a848a', band: '#d9534f',
    window: 'rgba(90, 60, 70, 0.22)', windowLit: null, smoke: ['rgba(255,245,235,0.95)', 'rgba(230,205,200,0)'],
    tree: ['#c2c080', '#737f4f'], trunk: '#6f5143', beacon: null, cityGlow: null,
  },
  night: {
    hillFar: ['#2c3d6c', '#1a274a'], hillMid: ['#20304f', '#121c35'], hillNear: ['#172a34', '#0b151c'],
    ground: ['#13232b', '#04090f'], road: 'rgba(120, 150, 170, 0.25)', haze: 'rgba(16, 26, 56, 0.6)',
    tower: ['#3f4f75', '#263455', '#141c33'], towerEdge: 'rgba(150, 170, 220, 0.22)',
    factory: ['#4a4546', '#312d31', '#1b191e'], roof: '#2a2628',
    cooling: ['#5c687f', '#3c465b', '#232b3b'], stack: ['#51607c', '#363f58', '#1f2638'], stackCap: '#2b3448', band: '#a33f3b',
    window: 'rgba(255, 232, 170, 0.12)', windowLit: 'rgba(255, 224, 150, 0.95)', smoke: ['rgba(205,214,235,0.6)', 'rgba(205,214,235,0)'],
    tree: ['#1f4b3d', '#0e2a22'], trunk: '#2a1f19', beacon: '#ff5a4a', cityGlow: 'rgba(255, 190, 110, 0.18)',
  },
};

const SKYLINE_SVG_TOWERS = [
  { x: 40, w: 70, h: 190 }, { x: 125, w: 56, h: 150 }, { x: 196, w: 84, h: 250 }, { x: 296, w: 62, h: 170 },
  { x: 560, w: 76, h: 230 }, { x: 650, w: 58, h: 180 }, { x: 760, w: 92, h: 290 }, { x: 870, w: 60, h: 200 },
  { x: 1130, w: 72, h: 210 }, { x: 1215, w: 56, h: 160 }, { x: 1490, w: 66, h: 180 },
];
const SKYLINE_SVG_FACTORIES = [
  { x: 360, w: 150, h: 120, teeth: 4 }, { x: 930, w: 170, h: 140, teeth: 5 }, { x: 1300, w: 150, h: 110, teeth: 4 },
];
const SKYLINE_SVG_COOLING = [
  { x: 1180, w: 130, h: 200 }, { x: 1380, w: 110, h: 170 },
];
const SKYLINE_SVG_STACKS = [
  { x: 330, w: 36, h: 310, smoke: 7 }, { x: 515, w: 52, h: 400, smoke: 9 }, { x: 700, w: 46, h: 360, smoke: 8 },
  { x: 905, w: 32, h: 290, smoke: 6 }, { x: 1060, w: 50, h: 380, smoke: 9 }, { x: 1250, w: 30, h: 250, smoke: 6 },
];
const SKYLINE_SVG_TREES = Array.from({ length: 22 }, (_, idx) => ({
  x: 20 + ((idx * 73) % 1560), rx: 11 + ((idx * 7) % 9), ry: 16 + ((idx * 11) % 14), dy: (idx * 5) % 12,
}));


  const SKY = {
    day: { bg: 'linear-gradient(180deg, #dff3ff 0%, #bee5ff 50%, #ffe8be 100%)', overlay: 'linear-gradient(180deg, rgba(255,255,255,0.34) 0%, rgba(224,242,254,0.1) 50%, rgba(0,0,0,0) 100%)', cloud: 'rgba(255,255,255,0.93)', cloudShadow: 'rgba(120,140,170,0.2)' },
    morning: { bg: 'linear-gradient(180deg, #99ceff 0%, #8ab8ef 50%, #f1b7c8 100%)', overlay: 'linear-gradient(180deg, rgba(255,243,214,0.34) 0%, rgba(255,228,184,0.16) 50%, rgba(0,0,0,0) 100%)', cloud: 'rgba(255,243,230,0.9)', cloudShadow: 'rgba(240,150,120,0.2)' },
    night: { bg: 'linear-gradient(180deg, #050914 0%, #264284 50%, #7a5a8c 100%)', overlay: 'linear-gradient(180deg, rgba(4,9,21,0.42) 0%, rgba(29,47,100,0.16) 50%, rgba(0,0,0,0) 100%)', cloud: 'rgba(190,200,235,0.09)', cloudShadow: 'rgba(0,0,0,0)' },
  };

  const CSS = `
.sp-sky { position: absolute; inset: 0; overflow: hidden; pointer-events: none; font-family: system-ui, sans-serif; }
.sp-sky-overlay { position: absolute; inset: 0; }
.sp-sky-landscape { position: absolute; left: 0; right: 0; bottom: 0; height: 62%; min-height: 320px; }

@keyframes skylineTwinkle {
  0%, 100% { opacity: 0.26; transform: scale(0.9); }
  50% { opacity: 0.95; transform: scale(1.08); }
}
.skyline-star {
  position: absolute;
  border-radius: 9999px;
  background: rgba(241, 245, 249, 0.95);
  box-shadow: 0 0 6px rgba(224, 231, 255, 0.8), 0 0 14px rgba(191, 219, 254, 0.45);
  animation: skylineTwinkle var(--twinkle-duration, 3.2s) ease-in-out infinite;
  animation-delay: var(--twinkle-delay, 0s);
  will-change: opacity, transform;
}
.skyline-moon {
  position: absolute;
  left: 8%;
  top: 8%;
  width: clamp(72px, 9vw, 108px);
  height: clamp(72px, 9vw, 108px);
  border-radius: 9999px;
  background: radial-gradient(circle at 35% 35%, rgba(254, 249, 195, 1) 0 44%, rgba(253, 230, 138, 0.96) 68%, rgba(250, 204, 21, 0.92) 100%);
  -webkit-mask: radial-gradient(circle at 70% 50%, transparent 46%, #000 47%);
  mask: radial-gradient(circle at 70% 50%, transparent 46%, #000 47%);
  box-shadow: 0 0 24px rgba(250, 204, 21, 0.35), 0 0 54px rgba(165, 180, 252, 0.18);
  transform: rotate(-12deg);
}
.skyline-moon::before {
  content: none;
}
.skyline-moon::after {
  content: none;
}
.skyline-sun {
  position: absolute;
  left: 8%;
  top: 8%;
  width: clamp(74px, 9vw, 112px);
  height: clamp(74px, 9vw, 112px);
  border-radius: 9999px;
  background: radial-gradient(circle at 36% 34%, rgba(254, 249, 195, 1), rgba(253, 230, 138, 0.95) 48%, rgba(251, 191, 36, 0.88));
  box-shadow: 0 0 0 12px rgba(254, 240, 138, 0.25), 0 0 36px rgba(250, 204, 21, 0.35), 0 0 70px rgba(249, 115, 22, 0.2);
}
.skyline-sunrise-mountain {
  position: absolute;
  left: 80%;
  top: 33%;
  width: clamp(92px, 10vw, 136px);
  height: clamp(92px, 10vw, 136px);
  border-radius: 9999px;
  background: radial-gradient(circle at 50% 38%, rgba(254, 249, 195, 1), rgba(253, 224, 71, 0.96) 44%, rgba(251, 146, 60, 0.9));
  box-shadow: 0 0 0 11px rgba(254, 240, 138, 0.2), 0 0 36px rgba(251, 191, 36, 0.32), 0 0 66px rgba(251, 146, 60, 0.2);
}
.skyline-sunrise-mountain-sheet {
  position: absolute;
  left: 78%;
  top: 54%;
  width: clamp(56px, 7vw, 84px);
  height: clamp(56px, 7vw, 84px);
  border-radius: 9999px;
  background: radial-gradient(circle at 50% 38%, rgba(254, 249, 195, 0.98), rgba(253, 224, 71, 0.92) 44%, rgba(251, 146, 60, 0.86));
  box-shadow: 0 0 0 8px rgba(254, 240, 138, 0.18), 0 0 28px rgba(251, 191, 36, 0.28), 0 0 48px rgba(251, 146, 60, 0.18);
}
@keyframes skylineCloudDrift {
  from { transform: translateX(-30vw) scale(var(--cloud-scale, 1)); }
  to { transform: translateX(125vw) scale(var(--cloud-scale, 1)); }
}
.skyline-cloud, .skyline-cloud::before, .skyline-cloud::after {
  background: var(--cloud-color, rgba(255, 255, 255, 0.92));
  border-radius: 9999px;
}
.skyline-cloud {
  position: absolute;
  left: 0;
  width: 190px;
  height: 54px;
  transform-origin: left center;
  filter: blur(0.6px) drop-shadow(0 6px 10px var(--cloud-shadow, rgba(120, 140, 170, 0.18)));
  animation: skylineCloudDrift var(--cloud-duration, 160s) linear infinite;
  animation-delay: var(--cloud-delay, 0s);
  will-change: transform;
}
.skyline-cloud::before {
  content: '';
  position: absolute;
  left: 18%;
  bottom: 30%;
  width: 56%;
  height: 140%;
}
.skyline-cloud::after {
  content: '';
  position: absolute;
  left: 52%;
  bottom: 24%;
  width: 40%;
  height: 120%;
}
@keyframes skylineSunRays {
  from { transform: rotate(0deg); }
  to { transform: rotate(360deg); }
}
.skyline-sun::after {
  content: '';
  position: absolute;
  inset: -90%;
  border-radius: 9999px;
  background: repeating-conic-gradient(from 0deg, rgba(255, 236, 170, 0.22) 0deg 7deg, rgba(255, 236, 170, 0) 7deg 24deg);
  -webkit-mask: radial-gradient(circle, rgba(0, 0, 0, 0.9) 20%, rgba(0, 0, 0, 0) 68%);
  mask: radial-gradient(circle, rgba(0, 0, 0, 0.9) 20%, rgba(0, 0, 0, 0) 68%);
  animation: skylineSunRays 90s linear infinite;
}
.skyline-moon-halo {
  position: absolute;
  left: calc(8% - 40px);
  top: calc(8% - 40px);
  width: clamp(152px, 17vw, 190px);
  height: clamp(152px, 17vw, 190px);
  border-radius: 9999px;
  background: radial-gradient(circle, rgba(250, 240, 200, 0.22), rgba(199, 210, 254, 0.1) 45%, rgba(199, 210, 254, 0) 72%);
}
@keyframes skylineShootingStar {
  0% { opacity: 0; transform: translate(0, 0) rotate(-28deg); }
  3% { opacity: 1; }
  10% { opacity: 0; transform: translate(-300px, 160px) rotate(-28deg); }
  100% { opacity: 0; transform: translate(-300px, 160px) rotate(-28deg); }
}
.skyline-shooting-star {
  position: absolute;
  width: 150px;
  height: 2px;
  border-radius: 9999px;
  background: linear-gradient(90deg, rgba(255, 255, 255, 0), rgba(255, 255, 255, 0.95));
  box-shadow: 0 0 8px rgba(255, 255, 255, 0.6);
  opacity: 0;
  animation: skylineShootingStar var(--shoot-duration, 13s) ease-out infinite;
  animation-delay: var(--shoot-delay, 0s);
}
@keyframes skylineWindowGlow {
  0%, 100% { opacity: 0.55; }
  50% { opacity: 1; }
}
.skyline-window-night {
  animation: skylineWindowGlow var(--glow-duration, 5s) ease-in-out infinite;
  animation-delay: var(--glow-delay, 0s);
}
@keyframes skylineSvgSmoke {
  0% { transform: translate(0px, 0px) scale(0.45); opacity: 0; }
  12% { opacity: var(--max-opacity, 0.5); }
  100% { transform: translate(var(--drift, 80px), calc(-1 * var(--rise, 260px))) scale(var(--end-scale, 3)); opacity: 0; }
}
.skyline-svg-smoke {
  transform-box: fill-box;
  transform-origin: center;
  opacity: 0;
  animation: skylineSvgSmoke var(--rise-duration, 16s) ease-out infinite;
  animation-delay: var(--rise-delay, 0s);
}
@keyframes skylineBeacon {
  0%, 100% { opacity: 0.15; }
  50% { opacity: 1; }
}
.skyline-beacon {
  animation: skylineBeacon 2.4s ease-in-out infinite;
  animation-delay: var(--beacon-delay, 0s);
}
.skyline-svg {
  position: absolute;
  left: 0;
  right: 0;
  bottom: 0;
  width: 100%;
  height: 100%;
  display: block;
}
`;

  const SVG_TOWERS = [
    { x: 40, w: 70, h: 190 }, { x: 125, w: 56, h: 150 }, { x: 196, w: 84, h: 250 }, { x: 296, w: 62, h: 170 },
    { x: 560, w: 76, h: 230 }, { x: 650, w: 58, h: 180 }, { x: 760, w: 92, h: 290 }, { x: 870, w: 60, h: 200 },
    { x: 1130, w: 72, h: 210 }, { x: 1215, w: 56, h: 160 }, { x: 1490, w: 66, h: 180 },
  ];
  const SVG_FACTORIES = [
    { x: 360, w: 150, h: 120, teeth: 4 }, { x: 930, w: 170, h: 140, teeth: 5 }, { x: 1300, w: 150, h: 110, teeth: 4 },
  ];
  const SVG_COOLING = [{ x: 1180, w: 130, h: 200 }, { x: 1380, w: 110, h: 170 }];
  const SVG_STACKS = [
    { x: 330, w: 36, h: 310, smoke: 7 }, { x: 515, w: 52, h: 400, smoke: 9 }, { x: 700, w: 46, h: 360, smoke: 8 },
    { x: 905, w: 32, h: 290, smoke: 6 }, { x: 1060, w: 50, h: 380, smoke: 9 }, { x: 1250, w: 30, h: 250, smoke: 6 },
  ];
  const SVG_TREES = Array.from({ length: 22 }, (_, idx) => ({
    x: 20 + ((idx * 73) % 1560), rx: 11 + ((idx * 7) % 9), ry: 16 + ((idx * 11) % 14), dy: (idx * 5) % 12,
  }));

  const esc = (v) => String(v);
  const el = (tag, attrs, children) => {
    const a = Object.keys(attrs || {}).map((k) => ` ${k}="${esc(attrs[k])}"`).join('');
    return `<${tag}${a}>${(children || []).join('')}</${tag}>`;
  };

  function buildSvg(mode, scope) {
    const pal = SKYLINE_PALETTES[mode] || SKYLINE_PALETTES.day;
    const id = (n) => `${scope}-sky-${n}`;
    const G = 440;
    const defs = [];
    [['hillFar', pal.hillFar], ['hillMid', pal.hillMid], ['hillNear', pal.hillNear], ['ground', pal.ground], ['tree', pal.tree]].forEach(([n, c]) => {
      defs.push(el('linearGradient', { id: id(n), x1: 0, y1: 0, x2: 0, y2: 1 }, [el('stop', { offset: 0, 'stop-color': c[0] }), el('stop', { offset: 1, 'stop-color': c[1] })]));
    });
    [['tower', pal.tower], ['factory', pal.factory], ['cooling', pal.cooling], ['stack', pal.stack]].forEach(([n, c]) => {
      defs.push(el('linearGradient', { id: id(n), x1: 0, y1: 0, x2: 1, y2: 0 }, [el('stop', { offset: 0, 'stop-color': c[0] }), el('stop', { offset: 0.45, 'stop-color': c[1] }), el('stop', { offset: 1, 'stop-color': c[2] })]));
    });
    defs.push(el('linearGradient', { id: id('haze'), x1: 0, y1: 0, x2: 0, y2: 1 }, [el('stop', { offset: 0, 'stop-color': pal.haze, 'stop-opacity': 0 }), el('stop', { offset: 1, 'stop-color': pal.haze })]));
    defs.push(el('radialGradient', { id: id('smoke'), cx: 0.45, cy: 0.4, r: 0.55 }, [el('stop', { offset: 0, 'stop-color': pal.smoke[0] }), el('stop', { offset: 1, 'stop-color': pal.smoke[1] })]));
    if (pal.cityGlow) defs.push(el('radialGradient', { id: id('cityglow'), cx: 0.5, cy: 1, r: 0.6 }, [el('stop', { offset: 0, 'stop-color': pal.cityGlow }), el('stop', { offset: 1, 'stop-color': pal.cityGlow, 'stop-opacity': 0 })]));
    defs.push(el('filter', { id: id('blur'), x: '-50%', y: '-50%', width: '200%', height: '200%' }, [el('feGaussianBlur', { stdDeviation: 3 })]));
    defs.push(el('filter', { id: id('soft'), x: '-20%', y: '-20%', width: '140%', height: '140%' }, [el('feGaussianBlur', { stdDeviation: 1.2 })]));

    const parts = [el('defs', {}, defs)];
    parts.push(el('path', { d: 'M0,330 C140,286 260,300 400,268 C560,232 680,282 840,262 C1000,240 1120,290 1280,258 C1400,232 1500,262 1600,238 L1600,520 L0,520 Z', fill: `url(#${id('hillFar')})` }));
    parts.push(el('path', { d: 'M0,372 C120,340 260,352 420,330 C600,306 720,352 900,338 C1080,322 1180,362 1360,340 C1460,328 1540,350 1600,338 L1600,520 L0,520 Z', fill: `url(#${id('hillMid')})` }));
    parts.push(el('rect', { x: 0, y: 330, width: 1600, height: 120, fill: `url(#${id('haze')})` }));
    if (pal.cityGlow) parts.push(el('rect', { x: 0, y: 260, width: 1600, height: 190, fill: `url(#${id('cityglow')})` }));

    SVG_TOWERS.forEach((t, i) => {
      const g = [el('rect', { x: t.x, y: G - t.h, width: t.w, height: t.h, rx: 6, fill: `url(#${id('tower')})`, stroke: pal.towerEdge, 'stroke-width': 1 }),
        el('rect', { x: t.x + 4, y: G - t.h + 4, width: Math.max(4, t.w * 0.22), height: t.h - 8, rx: 3, fill: 'rgba(255,255,255,0.22)' })];
      if (i % 3 === 0) g.push(el('rect', { x: t.x + t.w / 2 - 1.5, y: G - t.h - 26, width: 3, height: 26, rx: 1.5, fill: pal.stackCap }));
      if (pal.beacon && i % 3 === 0) g.push(el('circle', { class: 'skyline-beacon', cx: t.x + t.w / 2, cy: G - t.h - 27, r: 3, fill: pal.beacon, style: `--beacon-delay:-${i * 0.7}s` }));
      const cols = Math.max(2, Math.floor(t.w / 16)), rows = Math.max(2, Math.floor((t.h - 24) / 20));
      for (let r = 0; r < rows; r += 1) for (let c = 0; c < cols; c += 1) {
        const seed = i * 131 + r * 17 + c * 7;
        const lit = pal.windowLit && seededNoise(seed) > 0.42;
        g.push(el('rect', { x: t.x + 8 + c * ((t.w - 16) / cols) + 1, y: G - t.h + 16 + r * ((t.h - 24) / rows), width: Math.max(3, (t.w - 16) / cols - 5), height: 8, rx: 1, fill: lit ? pal.windowLit : pal.window, class: lit ? 'skyline-window-night' : '', style: lit ? `--glow-duration:${4 + (seed % 5)}s;--glow-delay:-${seed % 7}s` : '' }));
      }
      parts.push(el('g', {}, g));
    });
    SVG_FACTORIES.forEach((f, fi) => {
      const tw = f.w / f.teeth;
      const roof = Array.from({ length: f.teeth }).map((_, i) => `L${f.x + i * tw},${G - f.h} L${f.x + i * tw + tw * 0.62},${G - f.h - 22} L${f.x + (i + 1) * tw},${G - f.h}`).join(' ');
      const g = [el('rect', { x: f.x, y: G - f.h, width: f.w, height: f.h, rx: 4, fill: `url(#${id('factory')})` }), el('path', { d: `M${f.x},${G - f.h} ${roof} Z`, fill: pal.roof })];
      [0.3, 0.58].forEach((t, i) => g.push(el('rect', { x: f.x + 12, y: G - f.h + f.h * t, width: f.w - 24, height: 7, rx: 2, fill: pal.windowLit || pal.window, opacity: pal.windowLit ? 0.8 : 1, class: pal.windowLit ? 'skyline-window-night' : '', style: pal.windowLit ? `--glow-duration:${5 + i + fi}s;--glow-delay:-${fi * 2 + i}s` : '' })));
      parts.push(el('g', {}, g));
    });
    SVG_COOLING.forEach((t, ci) => {
      const top = G - t.h, x0 = t.x, x1 = t.x + t.w, inset = t.w * 0.17, waist = G - t.h * 0.42;
      const d = `M${x0},${G} C${x0 + inset * 0.3},${waist} ${x0 + inset},${waist} ${x0 + inset},${top + 18} L${x1 - inset},${top + 18} C${x1 - inset},${waist} ${x1 - inset * 0.3},${waist} ${x1},${G} Z`;
      const g = [el('path', { d, fill: `url(#${id('cooling')})`, stroke: pal.towerEdge, 'stroke-width': 1 }), el('ellipse', { cx: t.x + t.w / 2, cy: top + 18, rx: t.w * 0.33, ry: 5, fill: pal.stackCap, opacity: 0.55 })];
      for (let i = 0; i < 5; i += 1) g.push(el('circle', { class: 'skyline-svg-smoke', cx: t.x + t.w / 2, cy: top + 14, r: 16 + (i % 3) * 5, fill: `url(#${id('smoke')})`, filter: `url(#${id('blur')})`, style: `--drift:${50 + seededNoise(ci * 31 + i) * 120}px;--rise:${140 + seededNoise(ci * 17 + i * 3) * 110}px;--end-scale:${2.4 + seededNoise(ci + i * 5) * 1.2};--max-opacity:${0.3 + seededNoise(ci * 7 + i) * 0.25};--rise-duration:${15 + seededNoise(ci * 3 + i * 11) * 9}s;--rise-delay:-${seededNoise(ci * 13 + i * 2) * 20}s` }));
      parts.push(el('g', {}, g));
    });
    SVG_STACKS.forEach((st, si) => {
      const h = st.h, top = G - h;
      const g = [el('rect', { x: st.x, y: top, width: st.w, height: h, rx: st.w * 0.16, fill: `url(#${id('stack')})`, stroke: pal.towerEdge, 'stroke-width': 1 }),
        el('rect', { x: st.x + 3, y: top + 6, width: Math.max(3, st.w * 0.2), height: h - 12, rx: 2, fill: 'rgba(255,255,255,0.25)' }),
        el('rect', { x: st.x - 3, y: top, width: st.w + 6, height: 8, rx: 3, fill: pal.stackCap }),
        el('rect', { x: st.x, y: top + 26, width: st.w, height: 9, fill: pal.band, opacity: 0.85 }),
        el('rect', { x: st.x, y: top + 52, width: st.w, height: 9, fill: pal.band, opacity: 0.85 })];
      if (pal.beacon) g.push(el('circle', { class: 'skyline-beacon', cx: st.x + st.w / 2, cy: top - 4, r: 3.2, fill: pal.beacon, style: `--beacon-delay:-${si * 0.9}s` }));
      for (let i = 0; i < st.smoke; i += 1) {
        const seed = si * 97 + i * 31;
        g.push(el('circle', { class: 'skyline-svg-smoke', cx: st.x + st.w / 2, cy: top - 2, r: st.w * 0.32 + (i % 3) * 3, fill: `url(#${id('smoke')})`, filter: `url(#${id('blur')})`, style: `--drift:${(seededNoise(seed + 2) > 0.2 ? 1 : -1) * (60 + seededNoise(seed + 3) * 150)}px;--rise:${170 + seededNoise(seed + 5) * 170}px;--end-scale:${2.6 + seededNoise(seed + 8) * 1.6};--max-opacity:${0.35 + seededNoise(seed + 6) * 0.3};--rise-duration:${14 + seededNoise(seed + 4) * 9}s;--rise-delay:-${seededNoise(seed + 11) * 23}s` }));
      }
      parts.push(el('g', {}, g));
    });
    parts.push(el('path', { d: 'M0,418 C180,392 300,428 480,404 C660,380 820,420 1000,398 C1180,376 1320,418 1600,392 L1600,520 L0,520 Z', fill: `url(#${id('hillNear')})` }));
    SVG_TREES.forEach((tr) => {
      const baseY = 436 + tr.dy;
      parts.push(el('g', {}, [
        el('rect', { x: tr.x - 1.5, y: baseY - tr.ry - 2, width: 3, height: tr.ry + 4, rx: 1.5, fill: pal.trunk }),
        el('ellipse', { cx: tr.x, cy: baseY - tr.ry - 6, rx: tr.rx, ry: tr.ry, fill: `url(#${id('tree')})` }),
        el('ellipse', { cx: tr.x - tr.rx * 0.3, cy: baseY - tr.ry - 10, rx: tr.rx * 0.45, ry: tr.ry * 0.5, fill: 'rgba(255,255,255,0.14)' }),
      ]));
    });
    parts.push(el('rect', { x: 0, y: G + 8, width: 1600, height: 80, fill: `url(#${id('ground')})` }));
    parts.push(el('path', { d: 'M0,470 C300,462 600,478 900,468 C1200,458 1400,474 1600,466', fill: 'none', stroke: pal.road, 'stroke-width': 3, 'stroke-linecap': 'round', filter: `url(#${id('soft')})` }));
    return `<svg class="skyline-svg" viewBox="0 0 1600 520" preserveAspectRatio="xMidYMax slice" aria-hidden="true">${parts.join('')}</svg>`;
  }

  function resolveMode(mode) {
    if (mode && mode !== 'auto') return mode;
    const h = new Date().getHours();
    if (h >= 6 && h < 11) return 'morning';
    if (h >= 11 && h < 16) return 'day';
    return 'night';
  }

  function ensureStyle() {
    if (document.getElementById('stackpilot-skyline-style')) return;
    const style = document.createElement('style');
    style.id = 'stackpilot-skyline-style';
    style.textContent = CSS;
    document.head.appendChild(style);
  }

  function render(container, mode, scope) {
    const sky = SKY[mode] || SKY.day;
    const isNight = mode === 'night';
    const isMorning = mode === 'morning';
    const html = [];
    html.push(`<div class="sp-sky-overlay" style="background:${sky.overlay}"></div>`);
    if (isNight) html.push('<div class="skyline-moon-halo"></div><div class="skyline-moon"></div>');
    else if (isMorning) html.push('<div class="skyline-sunrise-mountain"></div>');
    else html.push('<div class="skyline-sun"></div>');
    SKYLINE_CLOUDS.forEach((c) => {
      html.push(`<div class="skyline-cloud" style="top:${c.top};opacity:${c.opacity};--cloud-scale:${c.scale};--cloud-duration:${c.duration}s;--cloud-delay:${c.delay}s;--cloud-color:${sky.cloud};--cloud-shadow:${sky.cloudShadow}"></div>`);
    });
    if (isNight) {
      SKYLINE_STARS.forEach((s) => html.push(`<span class="skyline-star" style="left:${s.left};top:${s.top};width:${s.size}px;height:${s.size}px;--twinkle-duration:${s.duration}s;--twinkle-delay:-${s.delay}s"></span>`));
      html.push('<div class="skyline-shooting-star" style="top:9%;left:72%;--shoot-duration:14s;--shoot-delay:-3s"></div>');
      html.push('<div class="skyline-shooting-star" style="top:18%;left:38%;--shoot-duration:23s;--shoot-delay:-15s"></div>');
    }
    html.push(`<div class="sp-sky-landscape">${buildSvg(mode, scope)}</div>`);
    container.innerHTML = `<div class="sp-sky" style="background:${sky.bg}">${html.join('')}</div>`;
  }

  const StackpilotSkyline = {
    mount(container, options) {
      ensureStyle();
      const opts = options || {};
      const scope = opts.scope || `sp${Math.random().toString(36).slice(2, 7)}`;
      let mode = resolveMode(opts.mode || 'auto');
      render(container, mode, scope);
      return {
        setMode(next) { mode = resolveMode(next); render(container, mode, scope); },
        getMode() { return mode; },
        destroy() { container.innerHTML = ''; },
      };
    },
    buildSvg,
    palettes: SKYLINE_PALETTES,
  };
  global.StackpilotSkyline = StackpilotSkyline;
})(typeof window !== 'undefined' ? window : globalThis);
