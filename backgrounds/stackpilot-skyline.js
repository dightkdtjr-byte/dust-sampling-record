/*
 * StackPilot 배경 (겹겹이 오린 종이 디오라마: 먼 산 · 도시 · 공단 · 앞 땅, 연기·기차·차량 애니메이션)
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

// ── 겹겹이 오린 종이 디오라마: 깊이별 한 가지 색 실루엣으로 그린 공단 풍경 ──
const SKYLINE_PALETTES = {
  day: {
    far: '#cfdceb', farDetail: '#bccbdd', city: '#a7bbd1', cityWin: '#bfd0e2', haze1: 'rgba(222,236,250,0.65)', haze2: 'rgba(214,230,246,0.5)',
    plant: '#73899f', plantShade: '#5e7389', plantLight: '#8ea2b6', plantBand: '#9fb1c3', plantWin: '#95a9bd',
    near: '#4b7362', nearShade: '#3a5c4e', nearLight: '#64917b', road: '#587566', roadLine: 'rgba(235,245,238,0.75)', rail: '#34514a',
    smoke: ['rgba(255,255,255,0.85)', 'rgba(255,255,255,0)'], windowLit: null, beacon: null, lamp: null, headlight: null,
    bird: '#48607a', plane: '#f2f6fa', contrail: 'rgba(255,255,255,0.7)', cirrus: 'rgba(255,255,255,0.55)', sunHaze: 'rgba(255,238,200,0.22)',
  },
  morning: {
    far: '#e3cfdd', farDetail: '#d4bccd', city: '#c0a3b8', cityWin: '#d2bacb', haze1: 'rgba(255,228,214,0.6)', haze2: 'rgba(250,220,210,0.45)',
    plant: '#8f7489', plantShade: '#776076', plantLight: '#a58ca0', plantBand: '#b59fb0', plantWin: '#ad95a8',
    near: '#5a4f6a', nearShade: '#473d57', nearLight: '#6e6380', road: '#66597a', roadLine: 'rgba(255,235,225,0.65)', rail: '#3d3450',
    smoke: ['rgba(255,246,238,0.85)', 'rgba(255,246,238,0)'], windowLit: null, beacon: null, lamp: null, headlight: null,
    bird: '#5a4a5c', plane: '#fff6ee', contrail: 'rgba(255,236,220,0.7)', cirrus: 'rgba(255,240,228,0.5)', sunHaze: 'rgba(255,190,150,0.3)',
  },
  night: {
    far: '#1a2850', farDetail: '#162243', city: '#111c3a', cityWin: '#ffd47a', haze1: 'rgba(30,45,90,0.5)', haze2: 'rgba(20,32,70,0.5)',
    plant: '#0d1630', plantShade: '#080f22', plantLight: '#16223f', plantBand: '#5a2d35', plantWin: '#ffd47a',
    near: '#060a18', nearShade: '#04060f', nearLight: '#0b1224', road: '#0e1526', roadLine: 'rgba(255,230,180,0.3)', rail: '#03050c',
    smoke: ['rgba(200,210,235,0.5)', 'rgba(200,210,235,0)'], windowLit: '#ffd47a', beacon: '#ff5a4a', lamp: 'rgba(255,214,130,0.9)', headlight: 'rgba(255,240,200,0.9)',
    bird: null, plane: '#8e9ab5', contrail: null, cirrus: 'rgba(190,205,240,0.07)', sunHaze: 'rgba(255,170,90,0.14)',
  },
};
const SCENE_GROUND = 432;
const SCENE_CITY = Array.from({ length: 64 }, (_, i) => ({ x: -12 + i * 26 + ((i * 7) % 9), w: 14 + ((i * 5) % 20), h: 26 + ((i * 13) % 74), antenna: i % 9 === 4 }));
const SCENE_HALLS = [{ x: 40, w: 280, h: 72, teeth: 5 }, { x: 880, w: 320, h: 92, teeth: 6 }, { x: 1320, w: 220, h: 62, teeth: 4 }];
const SCENE_BLOCKS = [{ x: 600, w: 84, h: 196 }, { x: 1090, w: 72, h: 166 }, { x: 0, w: 56, h: 140 }, { x: 760, w: 60, h: 120 }, { x: 1560, w: 60, h: 150 }];
const SCENE_STACKS = [
  { x: 150, w: 34, h: 296, smoke: 7 }, { x: 362, w: 48, h: 392, smoke: 9 }, { x: 556, w: 40, h: 336, smoke: 8 },
  { x: 1004, w: 52, h: 380, smoke: 9 }, { x: 1186, w: 30, h: 256, smoke: 6 }, { x: 1486, w: 36, h: 300, smoke: 7 },
];
const SCENE_COOLING = [{ x: 690, w: 150, h: 208 }, { x: 1236, w: 118, h: 178 }];
const SCENE_TANKS = [{ x: 420, w: 72, h: 60 }, { x: 504, w: 72, h: 82 }, { x: 1398, w: 62, h: 50 }, { x: 1468, w: 62, h: 50 }];
const SCENE_SPHERE = { cx: 862, cy: 388, r: 40 };
const SCENE_CRANE = { x: 262, top: 146, jib: 230, counter: 72, hookX: 140 };
const SCENE_PYLONS = [{ x: 470, h: 150 }, { x: 940, h: 165 }, { x: 1300, h: 150 }];
const SCENE_TREE_CLUSTERS = [[40, 62, 88], [230, 250], [330, 352, 372, 390], [560, 580], [700, 724, 744], [880, 902], [1010, 1036, 1052, 1076], [1210, 1230], [1360, 1384, 1402], [1520, 1544, 1566, 1590]];
const SCENE_TREES = SCENE_TREE_CLUSTERS.flatMap((cluster, ci) => cluster.map((x, k) => ({ x, r: 15 + ((ci * 5 + k * 3) % 11) })));
const SCENE_LAMPS = [120, 400, 680, 960, 1240, 1500];
const SCENE_CARS = [{ dir: 1, dur: 36, delay: -6, len: 26 }, { dir: -1, dur: 44, delay: -21, len: 24 }, { dir: 1, dur: 52, delay: -33, len: 36 }, { dir: -1, dur: 40, delay: -13, len: 22 }];
const SCENE_TRAIN = { dur: 100, delay: -42, cars: [50, 42, 46, 46, 42, 46, 42, 46] };
const SCENE_BIRD_FLOCKS = [
  { path: 'M-80,150 C300,120 700,170 1700,130', dur: 52, delay: -8, birds: [[0, 0], [-16, 9], [-32, 18], [16, 9], [32, 18], [-48, 26]] },
  { path: 'M1700,200 C1200,170 600,230 -120,190', dur: 64, delay: -30, birds: [[0, 0], [-14, 8], [14, 8], [-28, 16]] },
];
const SCENE_PLANE = { path: 'M-160,58 L1760,92', dur: 78, delay: -26 };
const SCENE_ROAD_Y = 478;
const SCENE_RAIL_Y = 500;

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
@keyframes skylineFlap {
  from { transform: scaleY(1); }
  to { transform: scaleY(0.3); }
}
.skyline-bird {
  transform-box: fill-box;
  transform-origin: center;
  animation: skylineFlap var(--flap-duration, 0.8s) ease-in-out infinite alternate;
  animation-delay: var(--flap-delay, 0s);
}
@keyframes skylineFog {
  from { transform: translateX(-70px); }
  to { transform: translateX(70px); }
}
.skyline-fog {
  animation: skylineFog var(--fog-duration, 36s) ease-in-out infinite alternate;
  animation-delay: var(--fog-delay, 0s);
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

  const esc = (v) => String(v);
  const el = (tag, attrs, children) => {
    const a = Object.keys(attrs || {}).map((k) => ` ${k}="${esc(attrs[k])}"`).join('');
    return `<${tag}${a}>${(children || []).join('')}</${tag}>`;
  };

  function buildSvg(mode, scope) {
    const pal = SKYLINE_PALETTES[mode] || SKYLINE_PALETTES.day;
    const id = (n) => `${scope}-sky-${n}`;
    const G = SCENE_GROUND;
    const night = Boolean(pal.windowLit);
    const shadeW = (w) => Math.max(4, Math.round(w * 0.34));
    const grad = (n, stops, horizontal) => el(horizontal ? 'linearGradient' : 'linearGradient', { id: id(n), x1: 0, y1: 0, x2: horizontal ? 1 : 0, y2: horizontal ? 0 : 1 }, stops.map(([o, c, op]) => el('stop', op === undefined ? { offset: o, 'stop-color': c } : { offset: o, 'stop-color': c, 'stop-opacity': op })));
    const defs = [
      el('radialGradient', { id: id('smoke'), cx: 0.45, cy: 0.4, r: 0.55 }, [el('stop', { offset: 0, 'stop-color': pal.smoke[0] }), el('stop', { offset: 1, 'stop-color': pal.smoke[1] })]),
      grad('haze1', [[0, pal.haze1, 0], [1, pal.haze1]]), grad('haze2', [[0, pal.haze2, 0], [1, pal.haze2]]),
      grad('sunhaze', [[0, pal.sunHaze, 0], [0.8, pal.sunHaze], [1, pal.sunHaze, 0]]), grad('ground', [[0, pal.near], [1, pal.nearShade]]),
      el('filter', { id: id('blur'), x: '-50%', y: '-50%', width: '200%', height: '200%' }, [el('feGaussianBlur', { stdDeviation: 3 })]),
      el('filter', { id: id('wide'), x: '-20%', y: '-100%', width: '140%', height: '300%' }, [el('feGaussianBlur', { stdDeviation: 9 })]),
    ];
    if (pal.contrail) defs.push(grad('contrail', [[0, pal.contrail, 0], [1, pal.contrail]], true));
    if (pal.lamp) defs.push(el('radialGradient', { id: id('lampglow'), cx: 0.5, cy: 0, r: 0.75 }, [el('stop', { offset: 0, 'stop-color': pal.lamp, 'stop-opacity': 0.55 }), el('stop', { offset: 1, 'stop-color': pal.lamp, 'stop-opacity': 0 })]));

    const parts = [el('defs', {}, defs)];
    const box = (x, y, w, h, extra) => el('g', {}, [el('rect', { x, y, width: w, height: h, fill: pal.plant }), el('rect', { x: x + w - shadeW(w), y, width: shadeW(w), height: h, fill: pal.plantShade })].concat(extra || []));
    const windowsFor = (x, y, w, h, seedBase) => {
      const cols = Math.max(2, Math.floor(w / 14)), rows = Math.max(2, Math.floor((h - 16) / 18)), out = [];
      for (let r = 0; r < rows; r += 1) for (let c = 0; c < cols; c += 1) {
        const seed = seedBase + r * 17 + c * 7, lit = night && seededNoise(seed) > 0.45;
        if (night && !lit) continue;
        out.push(el('rect', { x: x + 6 + c * ((w - 12) / cols) + 1, y: y + 10 + r * ((h - 16) / rows), width: Math.max(3, (w - 12) / cols - 4), height: 6, fill: pal.plantWin, class: lit ? 'skyline-window-night' : '', style: lit ? `--glow-duration:${4 + (seed % 5)}s;--glow-delay:-${seed % 7}s` : '' }));
      }
      return out;
    };
    const smokeFor = (cx, cy, baseR, count, seedBase) => Array.from({ length: count }).map((_, i) => {
      const seed = seedBase + i * 31;
      return el('circle', { class: 'skyline-svg-smoke', cx, cy, r: baseR + (i % 3) * 3, fill: `url(#${id('smoke')})`, filter: `url(#${id('blur')})`, style: `--drift:${(seededNoise(seed + 2) > 0.2 ? 1 : -1) * (60 + seededNoise(seed + 3) * 150)}px;--rise:${170 + seededNoise(seed + 5) * 170}px;--end-scale:${2.6 + seededNoise(seed + 8) * 1.6};--max-opacity:${0.35 + seededNoise(seed + 6) * 0.3};--rise-duration:${14 + seededNoise(seed + 4) * 9}s;--rise-delay:-${seededNoise(seed + 11) * 23}s` });
    });
    const treeBlob = (list, yBase) => list.map((t) => `M${t.x - t.r},${yBase} A${t.r},${t.r} 0 0 1 ${t.x + t.r},${yBase} Z`).join(' ');

    // 하늘 아래 빛 띠와 권운
    parts.push(el('rect', { x: 0, y: 140, width: 1600, height: 190, fill: `url(#${id('sunhaze')})` }));
    [[120, 48, 420, 7], [640, 92, 300, 5], [980, 36, 520, 8], [1380, 118, 260, 5]].forEach(([x, y, w, h]) => parts.push(el('ellipse', { cx: x + w / 2, cy: y, rx: w / 2, ry: h / 2, fill: pal.cirrus, filter: `url(#${id('wide')})` })));
    // 1층 먼 산
    parts.push(el('path', { d: 'M0,296 L70,262 L130,280 L210,226 L290,268 L350,246 L430,286 L520,214 L600,262 L660,244 L740,282 L830,228 L900,268 L960,250 L1040,292 L1130,230 L1200,266 L1260,238 L1340,288 L1420,224 L1500,264 L1560,250 L1600,270 L1600,520 L0,520 Z', fill: pal.far }));
    parts.push(el('path', { d: 'M520,214 L498,236 L545,234 Z M830,228 L808,248 L852,246 Z M1420,224 L1398,246 L1442,244 Z M210,226 L190,246 L232,244 Z', fill: pal.farDetail, opacity: 0.9 }));
    parts.push(el('rect', { x: 0, y: 230, width: 1600, height: 120, fill: `url(#${id('haze1')})` }));
    // 2층 도시
    SCENE_CITY.forEach((b, i) => {
      const g = [el('rect', { x: b.x, y: 318 - b.h, width: b.w, height: b.h + 60, fill: pal.city })];
      if (b.antenna) g.push(el('rect', { x: b.x + b.w / 2 - 1, y: 318 - b.h - 18, width: 2, height: 18, fill: pal.city }));
      if (night ? i % 2 === 0 : i % 3 === 0) [0, 1, 2].forEach((k) => g.push(el('rect', { x: b.x + 3, y: 318 - b.h + 6 + k * 12, width: Math.max(4, b.w - 6), height: 2, fill: pal.cityWin, opacity: night ? 0.55 : 0.8 })));
      parts.push(el('g', {}, g));
    });
    parts.push(el('rect', { x: 0, y: 300, width: 1600, height: 150, fill: `url(#${id('haze2')})` }));
    // 3층 공단
    {
      const c = SCENE_CRANE;
      const g = [
        el('path', { d: `M${c.x - 7},${G} L${c.x - 7},${c.top} M${c.x + 7},${G} L${c.x + 7},${c.top}` }),
        el('path', { d: Array.from({ length: 11 }).map((_, i) => `M${c.x - 7},${c.top + 14 + i * 26} L${c.x + 7},${c.top + i * 26}`).join(' '), 'stroke-width': 1.2 }),
        el('path', { d: `M${c.x - c.counter},${c.top - 10} L${c.x + c.jib},${c.top - 10}`, 'stroke-width': 3 }),
        el('path', { d: `M${c.x - c.counter},${c.top - 10} L${c.x},${c.top - 36} L${c.x + c.jib},${c.top - 10} M${c.x},${c.top - 36} L${c.x},${c.top - 18}`, 'stroke-width': 1.2 }),
        el('rect', { x: c.x - c.counter + 6, y: c.top - 10, width: 24, height: 14, fill: pal.plantShade, stroke: 'none' }),
        el('path', { d: `M${c.x + c.hookX},${c.top - 10} L${c.x + c.hookX},${c.top + 150}`, 'stroke-width': 1 }),
        el('rect', { x: c.x + c.hookX - 10, y: c.top + 150, width: 20, height: 16, fill: pal.plantShade, stroke: 'none' }),
      ];
      if (pal.beacon) g.push(el('circle', { class: 'skyline-beacon', cx: c.x, cy: c.top - 38, r: 3, fill: pal.beacon, stroke: 'none', style: '--beacon-delay:-1.6s' }));
      parts.push(el('g', { stroke: pal.plantShade, 'stroke-width': 2, fill: 'none', 'stroke-linecap': 'round' }, g));
    }
    SCENE_BLOCKS.forEach((b, i) => parts.push(box(b.x, G - b.h, b.w, b.h, windowsFor(b.x, G - b.h, b.w, b.h, i * 131))));
    SCENE_HALLS.forEach((f, i) => {
      const tw = f.w / f.teeth;
      const roof = Array.from({ length: f.teeth }).map((_, k) => `L${f.x + k * tw},${G - f.h} L${f.x + k * tw + tw * 0.62},${G - f.h - 22} L${f.x + (k + 1) * tw},${G - f.h}`).join(' ');
      parts.push(box(f.x, G - f.h, f.w, f.h, [
        el('path', { d: `M${f.x},${G - f.h} ${roof} Z`, fill: pal.plantLight }),
        el('path', { d: Array.from({ length: f.teeth }).map((_, k) => `M${f.x + k * tw + tw * 0.62},${G - f.h - 22} L${f.x + (k + 1) * tw},${G - f.h} L${f.x + k * tw + tw * 0.62},${G - f.h} Z`).join(' '), fill: pal.plantShade }),
        el('rect', { x: f.x + 14, y: G - f.h * 0.62, width: f.w - 28, height: 6, fill: pal.plantWin, opacity: night ? 0.9 : 1, class: night ? 'skyline-window-night' : '', style: night ? `--glow-duration:${5 + i}s;--glow-delay:-${i * 2}s` : '' }),
      ]));
    });
    {
      const s = SCENE_SPHERE;
      parts.push(el('g', {}, [
        el('path', { d: `M${s.cx - s.r * 0.6},${G} L${s.cx - s.r * 0.45},${s.cy + s.r * 0.5} M${s.cx + s.r * 0.6},${G} L${s.cx + s.r * 0.45},${s.cy + s.r * 0.5} M${s.cx - s.r * 0.18},${G} L${s.cx - s.r * 0.14},${s.cy + s.r * 0.92} M${s.cx + s.r * 0.18},${G} L${s.cx + s.r * 0.14},${s.cy + s.r * 0.92}`, stroke: pal.plantShade, 'stroke-width': 2.5, fill: 'none' }),
        el('circle', { cx: s.cx, cy: s.cy, r: s.r, fill: pal.plant }),
        el('path', { d: `M${s.cx + s.r * 0.25},${s.cy - s.r * 0.97} A${s.r},${s.r} 0 0 1 ${s.cx + s.r * 0.25},${s.cy + s.r * 0.97} A${s.r * 0.55},${s.r} 0 0 0 ${s.cx + s.r * 0.25},${s.cy - s.r * 0.97} Z`, fill: pal.plantShade }),
        el('ellipse', { cx: s.cx, cy: s.cy, rx: s.r, ry: s.r * 0.16, fill: 'none', stroke: pal.plantLight, 'stroke-width': 1.5, opacity: 0.8 }),
      ]));
    }
    SCENE_TANKS.forEach((t) => parts.push(box(t.x, G - t.h, t.w, t.h, [
      el('ellipse', { cx: t.x + t.w / 2, cy: G - t.h, rx: t.w / 2, ry: 5, fill: pal.plantLight }),
      el('rect', { x: t.x, y: G - t.h * 0.5, width: t.w, height: 3, fill: pal.plantBand, opacity: 0.7 }),
    ])));
    SCENE_COOLING.forEach((t, i) => {
      const top = G - t.h, x0 = t.x, x1 = t.x + t.w, inset = t.w * 0.17, waist = G - t.h * 0.42;
      parts.push(el('g', {}, [
        el('path', { d: `M${x0},${G} C${x0 + inset * 0.3},${waist} ${x0 + inset},${waist} ${x0 + inset},${top + 18} L${x1 - inset},${top + 18} C${x1 - inset},${waist} ${x1 - inset * 0.3},${waist} ${x1},${G} Z`, fill: pal.plant }),
        el('path', { d: `M${t.x + t.w * 0.62},${G} C${t.x + t.w * 0.66},${waist} ${t.x + t.w * 0.72},${waist} ${t.x + t.w * 0.72},${top + 18} L${x1 - inset},${top + 18} C${x1 - inset},${waist} ${x1 - t.w * 0.05},${waist} ${x1},${G} Z`, fill: pal.plantShade }),
        el('ellipse', { cx: t.x + t.w / 2, cy: top + 18, rx: t.w * 0.33, ry: 5, fill: pal.plantShade }),
      ].concat(smokeFor(t.x + t.w / 2, top + 14, 16, 5, i * 53))));
    });
    SCENE_STACKS.forEach((st, i) => {
      const top = G - st.h;
      const g = [
        el('rect', { x: st.x, y: top, width: st.w, height: st.h, fill: pal.plant }),
        el('rect', { x: st.x + st.w - shadeW(st.w), y: top, width: shadeW(st.w), height: st.h, fill: pal.plantShade }),
        el('rect', { x: st.x - 3, y: top, width: st.w + 6, height: 7, fill: pal.plantShade }),
        el('rect', { x: st.x, y: top + 24, width: st.w, height: 8, fill: pal.plantBand }),
        el('rect', { x: st.x, y: top + 48, width: st.w, height: 8, fill: pal.plantBand }),
      ];
      if (pal.beacon) g.push(el('circle', { class: 'skyline-beacon', cx: st.x + st.w / 2, cy: top - 4, r: 3.2, fill: pal.beacon, style: `--beacon-delay:-${i * 0.9}s` }));
      parts.push(el('g', {}, g.concat(smokeFor(st.x + st.w / 2, top - 2, st.w * 0.32, st.smoke, i * 97))));
    });
    SCENE_PYLONS.forEach((py, i) => {
      const top = G - py.h, next = SCENE_PYLONS[i + 1];
      const g = [
        el('path', { d: `M${py.x - 16},${G} L${py.x - 4},${top} L${py.x + 4},${top} L${py.x + 16},${G}` }),
        el('path', { d: `M${py.x - 12},${G - 30} L${py.x + 12},${G - 30} M${py.x - 9},${G - 62} L${py.x + 9},${G - 62} M${py.x - 13},${top + 22} L${py.x + 13},${top + 22} M${py.x - 22},${top + 46} L${py.x + 22},${top + 46}` }),
      ];
      if (next) [[22, 46], [13, 22]].forEach(([dx, dy]) => { const sag = Math.max(top, G - next.h) + dy + 46; g.push(el('path', { d: `M${py.x + dx},${top + dy} Q${(py.x + next.x) / 2},${sag} ${next.x - dx},${G - next.h + dy} M${py.x - dx},${top + dy} Q${(py.x + next.x) / 2},${sag} ${next.x + dx},${G - next.h + dy}`, 'stroke-width': 1, opacity: 0.7 })); });
      parts.push(el('g', { stroke: pal.plantShade, 'stroke-width': 2, fill: 'none', 'stroke-linecap': 'round' }, g));
    });
    parts.push(el('rect', { x: 0, y: G - 40, width: 1600, height: 44, fill: `url(#${id('haze2')})` }));
    // 4층 앞 땅
    parts.push(el('path', { d: `M0,${G - 14} C180,${G - 36} 300,${G} 480,${G - 24} C660,${G - 48} 820,${G - 8} 1000,${G - 30} C1180,${G - 52} 1320,${G - 10} 1600,${G - 36} L1600,520 L0,520 Z`, fill: `url(#${id('ground')})` }));
    parts.push(el('path', { d: treeBlob(SCENE_TREES, G + 6), fill: pal.nearLight, transform: 'translate(0 -26)' }));
    parts.push(el('path', { d: treeBlob(SCENE_TREES.map((t) => ({ x: t.x + 3, r: t.r * 0.72 })), G + 6), fill: pal.nearShade, transform: 'translate(0 -22)', opacity: 0.55 }));
    parts.push(el('rect', { x: 0, y: G + 6, width: 1600, height: 100, fill: pal.near }));
    parts.push(el('rect', { x: 0, y: G + 6, width: 1600, height: 18, fill: pal.nearLight, opacity: 0.35 }));
    parts.push(el('rect', { x: 0, y: SCENE_ROAD_Y - 8, width: 1600, height: 16, fill: pal.road }));
    parts.push(el('path', { d: `M0,${SCENE_ROAD_Y} L1600,${SCENE_ROAD_Y}`, stroke: pal.roadLine, 'stroke-width': 1.3, 'stroke-dasharray': '14 12' }));
    SCENE_LAMPS.forEach((x) => {
      const g = [];
      if (pal.lamp) g.push(el('ellipse', { cx: x + 10, cy: SCENE_ROAD_Y - 34, rx: 36, ry: 46, fill: `url(#${id('lampglow')})` }));
      g.push(el('path', { d: `M${x},${SCENE_ROAD_Y - 10} L${x},${SCENE_ROAD_Y - 40} Q${x},${SCENE_ROAD_Y - 46} ${x + 6},${SCENE_ROAD_Y - 46} L${x + 12},${SCENE_ROAD_Y - 46}`, fill: 'none', stroke: pal.nearShade, 'stroke-width': 2, 'stroke-linecap': 'round' }));
      g.push(el('rect', { x: x + 8, y: SCENE_ROAD_Y - 48, width: 8, height: 4, fill: pal.lamp || pal.nearLight }));
      parts.push(el('g', {}, g));
    });
    SCENE_CARS.forEach((car) => {
      const lane = car.dir > 0 ? 3 : -4, L = car.len;
      const inner = [el('rect', { x: -L / 2, y: -5, width: L, height: 6, rx: 2, fill: pal.nearShade }), el('rect', { x: -L / 2 + 5, y: -9, width: L * 0.55, height: 5, rx: 2, fill: pal.nearShade })];
      if (pal.headlight) {
        inner.push(el('path', { d: `M${L / 2},-3 L${L / 2 + 34},-9 L${L / 2 + 34},4 Z`, fill: pal.headlight, opacity: 0.2 }));
        inner.push(el('rect', { x: L / 2 - 1.5, y: -3.5, width: 2.5, height: 2.5, fill: pal.headlight }));
        inner.push(el('rect', { x: -L / 2 - 1, y: -3.5, width: 2.5, height: 2.5, fill: '#ff4d3d' }));
      }
      parts.push(el('g', {}, [
        el('g', { transform: `translate(0 ${SCENE_ROAD_Y + lane}) scale(${car.dir} 1)` }, inner),
        el('animateMotion', { dur: `${car.dur}s`, begin: `${car.delay}s`, repeatCount: 'indefinite', path: 'M-60,0 L1660,0', keyPoints: car.dir > 0 ? '0;1' : '1;0', keyTimes: '0;1', calcMode: 'linear' }),
      ]));
    });
    parts.push(el('path', { d: `M0,${SCENE_RAIL_Y + 6} L1600,${SCENE_RAIL_Y + 6}`, stroke: pal.rail, 'stroke-width': 5, 'stroke-dasharray': '3 9' }));
    parts.push(el('path', { d: `M0,${SCENE_RAIL_Y + 3.5} L1600,${SCENE_RAIL_Y + 3.5} M0,${SCENE_RAIL_Y + 8.5} L1600,${SCENE_RAIL_Y + 8.5}`, stroke: pal.rail, 'stroke-width': 1.4 }));
    {
      let offset = 0;
      const cars = SCENE_TRAIN.cars.map((w, i) => {
        const x = offset; offset += w + 5;
        const h = i === 0 ? 16 : 13, y = SCENE_RAIL_Y + 3 - h;
        const g = [el('rect', { x, y, width: w, height: h, rx: i % 3 === 1 ? 6 : 1.5, fill: pal.nearShade })];
        if (i === 0) g.push(el('rect', { x: x + w - 16, y: y - 6, width: 14, height: 7, fill: pal.nearShade }));
        if (i === 0 && pal.headlight) g.push(el('path', { d: `M${x + w},${y + 4} L${x + w + 46},${y - 6} L${x + w + 46},${y + 14} Z`, fill: pal.headlight, opacity: 0.22 }));
        g.push(el('circle', { cx: x + 7, cy: SCENE_RAIL_Y + 3.5, r: 2.4, fill: pal.rail }));
        g.push(el('circle', { cx: x + w - 7, cy: SCENE_RAIL_Y + 3.5, r: 2.4, fill: pal.rail }));
        return el('g', {}, g);
      });
      parts.push(el('g', {}, [el('g', { transform: `translate(${-offset} 0)` }, cars), el('animateMotion', { dur: `${SCENE_TRAIN.dur}s`, begin: `${SCENE_TRAIN.delay}s`, repeatCount: 'indefinite', path: `M-60,0 L${1600 + offset + 60},0` })]));
    }
    if (pal.bird) SCENE_BIRD_FLOCKS.forEach((flock) => {
      const g = flock.birds.map(([bx, by], bi) => el('path', { class: 'skyline-bird', d: `M${bx - 6},${by} Q${bx - 3},${by - 5} ${bx},${by} Q${bx + 3},${by - 5} ${bx + 6},${by}`, style: `--flap-duration:${0.7 + (bi % 3) * 0.12}s;--flap-delay:-${(bi * 0.23) % 0.8}s` }));
      g.push(el('animateMotion', { dur: `${flock.dur}s`, begin: `${flock.delay}s`, repeatCount: 'indefinite', path: flock.path }));
      parts.push(el('g', { fill: 'none', stroke: pal.bird, 'stroke-width': 1.6, 'stroke-linecap': 'round' }, g));
    });
    {
      const g = [];
      if (pal.contrail) g.push(el('rect', { x: -230, y: -1.2, width: 212, height: 2.4, rx: 1.2, fill: `url(#${id('contrail')})` }));
      g.push(el('path', { d: 'M-14,0 L8,0 Q14,0 14,2.5 L-14,2.5 Z M-8,0 L-6,-7 L-2,-7 L2,0 Z M-10,2.5 L-14,8 L-9,8 L-3,2.5 Z', fill: pal.plane }));
      if (!pal.contrail) { g.push(el('circle', { class: 'skyline-beacon', cx: -12, cy: 0, r: 1.8, fill: '#ff5a4a' })); g.push(el('circle', { class: 'skyline-beacon', cx: -4, cy: -7, r: 1.8, fill: '#8bf0a0', style: '--beacon-delay:-1.2s' })); }
      g.push(el('animateMotion', { dur: `${SCENE_PLANE.dur}s`, begin: `${SCENE_PLANE.delay}s`, repeatCount: 'indefinite', path: SCENE_PLANE.path, rotate: 'auto' }));
      parts.push(el('g', {}, g));
    }
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
