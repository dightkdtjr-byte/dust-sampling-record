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
    mountain: ['#f1f6fc', '#cbdcef'], asphalt: '#707a89', roadLine: 'rgba(255,255,255,0.75)', wire: 'rgba(60,70,90,0.38)', pylon: '#8a94a6',
    tank: ['#eef1f5', '#c3ccd8', '#8d9aae'], fog: 'rgba(255,255,255,0.42)', bird: '#3b4252', plane: '#f3f4f6', contrail: 'rgba(255,255,255,0.75)',
    lampGlow: null, headlight: null, taillight: null, carColors: ['#d9534f', '#3b82f6', '#f5f5f5', '#f4c20d', '#6b7280'],
    farCity: ['#dde7f2', '#c3d3e5'], cirrus: 'rgba(255,255,255,0.6)', horizonGlow: 'rgba(255,236,200,0.5)',
    crane: '#c6462f', rail: '#7d8690', sleeper: '#a0917c', loco: '#2f4f6f', wagons: ['#8a6d4b', '#6f7d6c', '#9a4b3c', '#7a7f8a'], trainLight: null,
  },
  morning: {
    hillFar: ['#ebdff0', '#c6b4db'], hillMid: ['#d3c1b2', '#95818a'], hillNear: ['#a9a97c', '#64704f'],
    ground: ['#84976a', '#3f4c39'], road: 'rgba(255, 240, 220, 0.35)', haze: 'rgba(255, 226, 190, 0.7)',
    tower: ['#fff4ea', '#dcc0b8', '#957f8b'], towerEdge: 'rgba(90, 60, 70, 0.18)',
    factory: ['#f3e1d2', '#c9aa9b', '#826d69'], roof: '#8f7570',
    cooling: ['#fff7ee', '#e4cfc3', '#a8918d'], stack: ['#fff5ea', '#dcc3b7', '#968085'], stackCap: '#9a848a', band: '#d9534f',
    window: 'rgba(90, 60, 70, 0.22)', windowLit: null, smoke: ['rgba(255,245,235,0.95)', 'rgba(230,205,200,0)'],
    tree: ['#c2c080', '#737f4f'], trunk: '#6f5143', beacon: null, cityGlow: null,
    mountain: ['#f8e9f1', '#dcc6df'], asphalt: '#7a6f78', roadLine: 'rgba(255,240,220,0.75)', wire: 'rgba(80,60,70,0.38)', pylon: '#8d7f88',
    tank: ['#fbeee4', '#d6bcb1', '#9a8488'], fog: 'rgba(255,230,200,0.45)', bird: '#4b3f4a', plane: '#fff5ec', contrail: 'rgba(255,235,215,0.75)',
    lampGlow: null, headlight: null, taillight: null, carColors: ['#c94f4b', '#4f7fd1', '#f8efe6', '#e6b83a', '#6e6a73'],
    farCity: ['#e9d8e6', '#d0bdd3'], cirrus: 'rgba(255,240,225,0.55)', horizonGlow: 'rgba(255,190,150,0.55)',
    crane: '#b6473a', rail: '#8a7f86', sleeper: '#9c8a7c', loco: '#4a4f66', wagons: ['#8d6c52', '#6f7566', '#954a3f', '#7b7682'], trainLight: null,
  },
  night: {
    hillFar: ['#2c3d6c', '#1a274a'], hillMid: ['#20304f', '#121c35'], hillNear: ['#172a34', '#0b151c'],
    ground: ['#13232b', '#04090f'], road: 'rgba(120, 150, 170, 0.25)', haze: 'rgba(16, 26, 56, 0.6)',
    tower: ['#3f4f75', '#263455', '#141c33'], towerEdge: 'rgba(150, 170, 220, 0.22)',
    factory: ['#4a4546', '#312d31', '#1b191e'], roof: '#2a2628',
    cooling: ['#5c687f', '#3c465b', '#232b3b'], stack: ['#51607c', '#363f58', '#1f2638'], stackCap: '#2b3448', band: '#a33f3b',
    window: 'rgba(255, 232, 170, 0.12)', windowLit: 'rgba(255, 224, 150, 0.95)', smoke: ['rgba(205,214,235,0.6)', 'rgba(205,214,235,0)'],
    tree: ['#1f4b3d', '#0e2a22'], trunk: '#2a1f19', beacon: '#ff5a4a', cityGlow: 'rgba(255, 190, 110, 0.18)',
    mountain: ['#1c2c56', '#14203f'], asphalt: '#0f1620', roadLine: 'rgba(230,220,170,0.35)', wire: 'rgba(170,185,220,0.25)', pylon: '#3c4760',
    tank: ['#4b586f', '#2f3a4f', '#1b2232'], fog: 'rgba(120,140,190,0.16)', bird: null, plane: '#8e9ab5', contrail: null,
    lampGlow: 'rgba(255,214,130,0.5)', headlight: 'rgba(255,240,200,0.95)', taillight: '#ff4d3d', carColors: ['#5b6475', '#4a5366', '#6b7386', '#565e70', '#3f4756'],
    farCity: ['#1f2f58', '#182546'], cirrus: 'rgba(190,205,240,0.07)', horizonGlow: 'rgba(255,170,90,0.16)',
    crane: '#4c2b2f', rail: '#2b3340', sleeper: '#2a2420', loco: '#1c2a3a', wagons: ['#3a3129', '#2d3530', '#3f2622', '#30343c'], trainLight: 'rgba(255,240,200,0.9)',
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
  // 송전탑(전선으로 이어짐), 저장탱크, 가로등, 도로 위 차량, 새 떼, 비행기
  const SVG_PYLONS = [{ x: 470, h: 150 }, { x: 1020, h: 165 }, { x: 1345, h: 150 }];
  const SVG_TANKS = [{ x: 1500, w: 30, h: 58 }, { x: 1538, w: 30, h: 70 }, { x: 1576, w: 30, h: 58 }];
  const ROAD_PATH = 'M-40,471 C300,462 600,478 900,468 C1200,458 1400,474 1640,466';
  const SVG_LAMPS = [{ x: 150, y: 467 }, { x: 450, y: 470 }, { x: 750, y: 473 }, { x: 1050, y: 465 }, { x: 1350, y: 470 }];
  const SVG_CARS = [
    { dir: 1, dur: 34, delay: -5, len: 26, color: 0 }, { dir: -1, dur: 42, delay: -20, len: 24, color: 1 },
    { dir: 1, dur: 50, delay: -31, len: 34, color: 2 }, { dir: -1, dur: 38, delay: -12, len: 22, color: 3 },
    { dir: 1, dur: 44, delay: -17, len: 24, color: 4 },
  ];
  const BIRD_FLOCKS = [
    { path: 'M-80,150 C300,120 700,170 1700,130', dur: 52, delay: -8, birds: [[0, 0], [-16, 9], [-32, 18], [16, 9], [32, 18], [-48, 26]] },
    { path: 'M1700,200 C1200,170 600,230 -120,190', dur: 64, delay: -30, birds: [[0, 0], [-14, 8], [14, 8], [-28, 16]] },
  ];
  const PLANE = { path: 'M-160,58 L1760,92', dur: 78, delay: -26 };
  // 원경 도시, 권운, 타워크레인, 가스홀더, 철길과 화물열차
  const FAR_CITY = Array.from({ length: 46 }, (_, i) => ({ x: -10 + i * 35 + ((i * 13) % 9), w: 18 + ((i * 7) % 16), h: 14 + ((i * 11) % 34) }));
  const CIRRUS = [{ x: 120, y: 48, w: 420, h: 7 }, { x: 640, y: 92, w: 300, h: 5 }, { x: 980, y: 36, w: 520, h: 8 }, { x: 1380, y: 118, w: 260, h: 5 }];
  const CRANE = { x: 176, top: 128, jib: 230, counter: 78, hookX: 150 };
  const GAS_HOLDER = { cx: 1000, cy: 352, r: 44 };
  const RAIL_Y = 500;
  const TRAIN = { dur: 96, delay: -40, cars: [
    { w: 46, h: 16, kind: 'loco' }, { w: 40, h: 13, kind: 'tank' }, { w: 44, h: 14, kind: 'box' }, { w: 44, h: 14, kind: 'box' },
    { w: 40, h: 13, kind: 'tank' }, { w: 44, h: 11, kind: 'flat' }, { w: 44, h: 14, kind: 'box' }, { w: 40, h: 13, kind: 'tank' },
  ] };

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
    [['mountain', pal.mountain], ['farCity', pal.farCity], ['hillFar', pal.hillFar], ['hillMid', pal.hillMid], ['hillNear', pal.hillNear], ['ground', pal.ground], ['tree', pal.tree]].forEach(([n, c]) => {
      defs.push(el('linearGradient', { id: id(n), x1: 0, y1: 0, x2: 0, y2: 1 }, [el('stop', { offset: 0, 'stop-color': c[0] }), el('stop', { offset: 1, 'stop-color': c[1] })]));
    });
    [['tower', pal.tower], ['factory', pal.factory], ['cooling', pal.cooling], ['stack', pal.stack], ['tank', pal.tank]].forEach(([n, c]) => {
      defs.push(el('linearGradient', { id: id(n), x1: 0, y1: 0, x2: 1, y2: 0 }, [el('stop', { offset: 0, 'stop-color': c[0] }), el('stop', { offset: 0.45, 'stop-color': c[1] }), el('stop', { offset: 1, 'stop-color': c[2] })]));
    });
    defs.push(el('linearGradient', { id: id('haze'), x1: 0, y1: 0, x2: 0, y2: 1 }, [el('stop', { offset: 0, 'stop-color': pal.haze, 'stop-opacity': 0 }), el('stop', { offset: 1, 'stop-color': pal.haze })]));
    defs.push(el('radialGradient', { id: id('smoke'), cx: 0.45, cy: 0.4, r: 0.55 }, [el('stop', { offset: 0, 'stop-color': pal.smoke[0] }), el('stop', { offset: 1, 'stop-color': pal.smoke[1] })]));
    if (pal.cityGlow) defs.push(el('radialGradient', { id: id('cityglow'), cx: 0.5, cy: 1, r: 0.6 }, [el('stop', { offset: 0, 'stop-color': pal.cityGlow }), el('stop', { offset: 1, 'stop-color': pal.cityGlow, 'stop-opacity': 0 })]));
    defs.push(el('filter', { id: id('blur'), x: '-50%', y: '-50%', width: '200%', height: '200%' }, [el('feGaussianBlur', { stdDeviation: 3 })]));
    defs.push(el('filter', { id: id('soft'), x: '-20%', y: '-20%', width: '140%', height: '140%' }, [el('feGaussianBlur', { stdDeviation: 1.2 })]));
    defs.push(el('filter', { id: id('fogblur'), x: '-20%', y: '-100%', width: '140%', height: '300%' }, [el('feGaussianBlur', { stdDeviation: 9 })]));
    if (pal.contrail) defs.push(el('linearGradient', { id: id('contrail'), x1: 0, y1: 0, x2: 1, y2: 0 }, [el('stop', { offset: 0, 'stop-color': pal.contrail, 'stop-opacity': 0 }), el('stop', { offset: 1, 'stop-color': pal.contrail })]));
    if (pal.lampGlow) defs.push(el('radialGradient', { id: id('lampglow'), cx: 0.5, cy: 0, r: 0.75 }, [el('stop', { offset: 0, 'stop-color': pal.lampGlow }), el('stop', { offset: 1, 'stop-color': pal.lampGlow, 'stop-opacity': 0 })]));

    defs.push(el('linearGradient', { id: id('horizon'), x1: 0, y1: 0, x2: 0, y2: 1 }, [el('stop', { offset: 0, 'stop-color': pal.horizonGlow, 'stop-opacity': 0 }), el('stop', { offset: 0.85, 'stop-color': pal.horizonGlow }), el('stop', { offset: 1, 'stop-color': pal.horizonGlow, 'stop-opacity': 0 })]));

    const parts = [el('defs', {}, defs)];
    // 지평선 빛 띠와 권운
    parts.push(el('rect', { x: 0, y: 150, width: 1600, height: 170, fill: `url(#${id('horizon')})` }));
    CIRRUS.forEach((c) => parts.push(el('ellipse', { cx: c.x + c.w / 2, cy: c.y, rx: c.w / 2, ry: c.h / 2, fill: pal.cirrus, filter: `url(#${id('fogblur')})` })));
    // 먼 산맥 (대기 원근으로 가장 연함) + 낮에는 만년설
    parts.push(el('path', { d: 'M0,300 L90,250 L150,276 L230,214 L300,262 L360,236 L430,286 L520,206 L600,258 L660,238 L740,282 L830,222 L900,266 L960,244 L1040,292 L1130,226 L1200,264 L1260,232 L1340,288 L1420,218 L1500,262 L1560,248 L1600,270 L1600,520 L0,520 Z', fill: `url(#${id('mountain')})` }));
    if (pal.contrail) parts.push(el('path', { d: 'M230,214 L205,236 L255,232 Z M520,206 L498,228 L545,226 Z M830,222 L808,242 L852,240 Z M1130,226 L1108,248 L1152,246 Z M1420,218 L1398,240 L1442,238 Z', fill: 'rgba(255,255,255,0.55)' }));
    // 원경 도시: 산 앞에 납작한 건물 열
    FAR_CITY.forEach((b) => parts.push(el('rect', { x: b.x, y: 302 - b.h, width: b.w, height: b.h + 40, fill: `url(#${id('farCity')})` })));
    if (pal.windowLit) FAR_CITY.forEach((b, i) => { if (i % 3 !== 1) parts.push(el('rect', { x: b.x + 3, y: 302 - b.h + 4 + ((i * 7) % 10), width: b.w - 6, height: 2, fill: pal.windowLit, opacity: 0.35 })); });
    parts.push(el('path', { d: 'M0,330 C140,286 260,300 400,268 C560,232 680,282 840,262 C1000,240 1120,290 1280,258 C1400,232 1500,262 1600,238 L1600,520 L0,520 Z', fill: `url(#${id('hillFar')})` }));
    parts.push(el('path', { d: 'M0,372 C120,340 260,352 420,330 C600,306 720,352 900,338 C1080,322 1180,362 1360,340 C1460,328 1540,350 1600,338 L1600,520 L0,520 Z', fill: `url(#${id('hillMid')})` }));
    parts.push(el('rect', { x: 0, y: 330, width: 1600, height: 120, fill: `url(#${id('haze')})` }));
    if (pal.cityGlow) parts.push(el('rect', { x: 0, y: 260, width: 1600, height: 190, fill: `url(#${id('cityglow')})` }));

    // 타워크레인
    {
      const c = CRANE;
      const g = [
        el('path', { d: `M${c.x - 7},${G} L${c.x - 7},${c.top} M${c.x + 7},${G} L${c.x + 7},${c.top}` }),
        el('path', { d: Array.from({ length: 12 }).map((_, i) => `M${c.x - 7},${c.top + 14 + i * 26} L${c.x + 7},${c.top + i * 26}`).join(' '), 'stroke-width': 1.2 }),
        el('path', { d: `M${c.x - 7},${c.top} L${c.x + 7},${c.top} L${c.x + 7},${c.top - 18} L${c.x - 7},${c.top - 18} Z` }),
        el('path', { d: `M${c.x - c.counter},${c.top - 10} L${c.x + c.jib},${c.top - 10}`, 'stroke-width': 3 }),
        el('path', { d: `M${c.x - c.counter},${c.top - 10} L${c.x},${c.top - 36} L${c.x + c.jib},${c.top - 10} M${c.x},${c.top - 36} L${c.x},${c.top - 18}`, 'stroke-width': 1.2 }),
        el('path', { d: `M${c.x - c.counter + 6},${c.top - 10} L${c.x - c.counter + 6},${c.top + 4} L${c.x - c.counter + 30},${c.top + 4} L${c.x - c.counter + 30},${c.top - 10}`, fill: pal.crane }),
        el('path', { d: `M${c.x + c.hookX},${c.top - 10} L${c.x + c.hookX},${c.top + 150}`, 'stroke-width': 1 }),
        el('path', { d: `M${c.x + c.hookX - 10},${c.top + 150} L${c.x + c.hookX + 10},${c.top + 150} L${c.x + c.hookX + 10},${c.top + 166} L${c.x + c.hookX - 10},${c.top + 166} Z`, fill: pal.stackCap }),
      ];
      if (pal.beacon) g.push(el('circle', { class: 'skyline-beacon', cx: c.x, cy: c.top - 38, r: 3, fill: pal.beacon, style: '--beacon-delay:-1.6s' }));
      parts.push(el('g', { stroke: pal.crane, 'stroke-width': 2, fill: 'none', 'stroke-linecap': 'round', opacity: 0.92 }, g));
    }
    // 가스홀더
    {
      const g = GAS_HOLDER;
      parts.push(el('g', {}, [
        el('path', { d: `M${g.cx - g.r * 0.6},${G} L${g.cx - g.r * 0.45},${g.cy + g.r * 0.5} M${g.cx + g.r * 0.6},${G} L${g.cx + g.r * 0.45},${g.cy + g.r * 0.5} M${g.cx - g.r * 0.2},${G} L${g.cx - g.r * 0.15},${g.cy + g.r * 0.9} M${g.cx + g.r * 0.2},${G} L${g.cx + g.r * 0.15},${g.cy + g.r * 0.9}`, stroke: pal.pylon, 'stroke-width': 2.5, fill: 'none' }),
        el('circle', { cx: g.cx, cy: g.cy, r: g.r, fill: `url(#${id('tank')})`, stroke: pal.towerEdge, 'stroke-width': 1 }),
        el('ellipse', { cx: g.cx, cy: g.cy, rx: g.r, ry: g.r * 0.18, fill: 'none', stroke: pal.band, 'stroke-width': 2, opacity: 0.6 }),
        el('path', { d: `M${g.cx - g.r},${g.cy} A${g.r},${g.r} 0 0 1 ${g.cx + g.r},${g.cy}`, fill: 'none', stroke: 'rgba(255,255,255,0.22)', 'stroke-width': 6 }),
        el('path', { d: `M${g.cx + g.r * 0.95},${g.cy - 4} L${g.cx + g.r * 0.95},${g.cy - g.r - 10} L${g.cx + g.r * 0.6},${g.cy - g.r - 10}`, stroke: pal.pylon, 'stroke-width': 1.5, fill: 'none' }),
      ]));
    }
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
    SVG_PYLONS.forEach((py, i) => {
      const top = G - py.h, next = SVG_PYLONS[i + 1];
      const g = [
        el('path', { d: `M${py.x - 16},${G} L${py.x - 4},${top} L${py.x + 4},${top} L${py.x + 16},${G}` }),
        el('path', { d: `M${py.x - 12},${G - 30} L${py.x + 12},${G - 30} M${py.x - 9},${G - 62} L${py.x + 9},${G - 62} M${py.x - 13},${top + 22} L${py.x + 13},${top + 22} M${py.x - 22},${top + 46} L${py.x + 22},${top + 46}` }),
        el('path', { d: `M${py.x - 12},${G - 30} L${py.x + 9},${G - 62} M${py.x + 12},${G - 30} L${py.x - 9},${G - 62} M${py.x - 9},${G - 62} L${py.x + 7},${G - 94} M${py.x + 9},${G - 62} L${py.x - 7},${G - 94}`, 'stroke-width': 1.2 }),
      ];
      if (next) [[22, 46], [13, 22]].forEach(([dx, dy]) => {
        const sag = Math.max(top, G - next.h) + dy + 46;
        g.push(el('path', { d: `M${py.x + dx},${top + dy} Q${(py.x + next.x) / 2},${sag} ${next.x - dx},${G - next.h + dy} M${py.x - dx},${top + dy} Q${(py.x + next.x) / 2},${sag} ${next.x + dx},${G - next.h + dy}`, stroke: pal.wire, 'stroke-width': 1 }));
      });
      parts.push(el('g', { stroke: pal.pylon, 'stroke-width': 2, fill: 'none', 'stroke-linecap': 'round' }, g));
    });
    SVG_TANKS.forEach((tk) => {
      parts.push(el('g', {}, [
        el('rect', { x: tk.x, y: G - tk.h, width: tk.w, height: tk.h, rx: 3, fill: `url(#${id('tank')})`, stroke: pal.towerEdge, 'stroke-width': 1 }),
        el('ellipse', { cx: tk.x + tk.w / 2, cy: G - tk.h, rx: tk.w / 2, ry: 4, fill: pal.stackCap, opacity: 0.8 }),
        el('rect', { x: tk.x, y: G - tk.h * 0.55, width: tk.w, height: 3, fill: pal.band, opacity: 0.7 }),
        el('rect', { x: tk.x + tk.w - 6, y: G - tk.h + 6, width: 2, height: tk.h - 8, fill: pal.pylon, opacity: 0.7 }),
      ]));
    });
    parts.push(el('path', { d: 'M0,418 C180,392 300,428 480,404 C660,380 820,420 1000,398 C1180,376 1320,418 1600,392 L1600,520 L0,520 Z', fill: `url(#${id('hillNear')})` }));
    [{ cx: 380, rx: 420, dur: 38, delay: -6 }, { cx: 1180, rx: 480, dur: 46, delay: -21 }].forEach((fg) => {
      parts.push(el('ellipse', { class: 'skyline-fog', cx: fg.cx, cy: G - 6, rx: fg.rx, ry: 16, fill: pal.fog, filter: `url(#${id('fogblur')})`, style: `--fog-duration:${fg.dur}s;--fog-delay:${fg.delay}s` }));
    });
    SVG_TREES.forEach((tr) => {
      const baseY = 436 + tr.dy;
      parts.push(el('g', {}, [
        el('rect', { x: tr.x - 1.5, y: baseY - tr.ry - 2, width: 3, height: tr.ry + 4, rx: 1.5, fill: pal.trunk }),
        el('ellipse', { cx: tr.x, cy: baseY - tr.ry - 6, rx: tr.rx, ry: tr.ry, fill: `url(#${id('tree')})` }),
        el('ellipse', { cx: tr.x - tr.rx * 0.3, cy: baseY - tr.ry - 10, rx: tr.rx * 0.45, ry: tr.ry * 0.5, fill: 'rgba(255,255,255,0.14)' }),
      ]));
    });
    parts.push(el('rect', { x: 0, y: G + 8, width: 1600, height: 80, fill: `url(#${id('ground')})` }));
    parts.push(el('path', { d: ROAD_PATH, fill: 'none', stroke: pal.asphalt, 'stroke-width': 15, 'stroke-linecap': 'round' }));
    parts.push(el('path', { d: ROAD_PATH, fill: 'none', stroke: pal.roadLine, 'stroke-width': 1.4, 'stroke-dasharray': '14 12', opacity: 0.9 }));
    parts.push(el('path', { d: ROAD_PATH, fill: 'none', stroke: pal.road, 'stroke-width': 15, 'stroke-linecap': 'round', opacity: 0.25, filter: `url(#${id('soft')})` }));
    SVG_LAMPS.forEach((lp) => {
      const g = [];
      if (pal.lampGlow) g.push(el('ellipse', { cx: lp.x + 10, cy: lp.y - 32, rx: 34, ry: 44, fill: `url(#${id('lampglow')})` }));
      g.push(el('path', { d: `M${lp.x},${lp.y - 8} L${lp.x},${lp.y - 38} Q${lp.x},${lp.y - 44} ${lp.x + 6},${lp.y - 44} L${lp.x + 12},${lp.y - 44}`, fill: 'none', stroke: pal.pylon, 'stroke-width': 2, 'stroke-linecap': 'round' }));
      g.push(el('rect', { x: lp.x + 8, y: lp.y - 46, width: 8, height: 4, rx: 1.5, fill: pal.lampGlow ? '#ffe9b0' : pal.stackCap }));
      parts.push(el('g', {}, g));
    });
    SVG_CARS.forEach((car) => {
      const body = pal.carColors[car.color % pal.carColors.length];
      const lane = car.dir > 0 ? 3 : -3, L = car.len;
      const inner = [
        el('rect', { x: -L / 2, y: -5, width: L, height: 7, rx: 2.5, fill: body }),
        el('rect', { x: -L / 2 + 5, y: -10, width: L * 0.55, height: 6, rx: 2.5, fill: body }),
        el('rect', { x: -L / 2 + 7, y: -9, width: L * 0.22, height: 4, rx: 1, fill: 'rgba(200,225,245,0.75)' }),
        el('circle', { cx: -L / 2 + 5, cy: 2.5, r: 2.2, fill: '#1f2937' }),
        el('circle', { cx: L / 2 - 5, cy: 2.5, r: 2.2, fill: '#1f2937' }),
      ];
      if (pal.headlight) {
        inner.push(el('path', { d: `M${L / 2},-3.5 L${L / 2 + 34},-10 L${L / 2 + 34},4 Z`, fill: pal.headlight, opacity: 0.3 }));
        inner.push(el('rect', { x: L / 2 - 1.5, y: -3.5, width: 2.5, height: 3, rx: 1, fill: pal.headlight }));
      }
      if (pal.taillight) inner.push(el('rect', { x: -L / 2 - 1, y: -3.5, width: 2.5, height: 3, rx: 1, fill: pal.taillight }));
      parts.push(el('g', {}, [
        el('g', { transform: `translate(0 ${lane}) scale(${car.dir} 1)` }, inner),
        el('animateMotion', { dur: `${car.dur}s`, begin: `${car.delay}s`, repeatCount: 'indefinite', path: ROAD_PATH, keyPoints: car.dir > 0 ? '0;1' : '1;0', keyTimes: '0;1', calcMode: 'linear' }),
      ]));
    });
    // 철길과 화물열차
    parts.push(el('g', {}, [
      el('path', { d: `M0,${RAIL_Y + 6} L1600,${RAIL_Y + 6}`, stroke: pal.sleeper, 'stroke-width': 5, 'stroke-dasharray': '3 9', opacity: 0.9 }),
      el('path', { d: `M0,${RAIL_Y + 3.5} L1600,${RAIL_Y + 3.5} M0,${RAIL_Y + 8.5} L1600,${RAIL_Y + 8.5}`, stroke: pal.rail, 'stroke-width': 1.4 }),
    ]));
    {
      const gap = 5; let offset = 0;
      const cars = TRAIN.cars.map((car, i) => {
        const x = offset; offset += car.w + gap;
        const body = car.kind === 'loco' ? pal.loco : pal.wagons[i % pal.wagons.length];
        const y = RAIL_Y + 3 - car.h;
        const g = [];
        if (car.kind === 'tank') g.push(el('rect', { x, y: y + 2, width: car.w, height: car.h - 2, rx: (car.h - 2) / 2, fill: body }));
        else g.push(el('rect', { x, y, width: car.w, height: car.h, rx: 1.5, fill: body }));
        if (car.kind === 'loco') {
          g.push(el('rect', { x: x + car.w - 16, y: y - 6, width: 14, height: 7, rx: 1.5, fill: body }));
          g.push(el('rect', { x: x + car.w - 13, y: y - 5, width: 8, height: 4, rx: 0.8, fill: pal.trainLight || 'rgba(200,225,245,0.8)' }));
          if (pal.trainLight) g.push(el('path', { d: `M${x + car.w},${y + 4} L${x + car.w + 46},${y - 6} L${x + car.w + 46},${y + 14} Z`, fill: pal.trainLight, opacity: 0.22 }));
        }
        if (car.kind === 'box') g.push(el('rect', { x: x + car.w / 2 - 2, y: y + 3, width: 4, height: car.h - 6, fill: 'rgba(0,0,0,0.2)' }));
        if (car.kind === 'flat') [0, 1].forEach((k) => g.push(el('rect', { x: x + 4 + k * 20, y: y - 7, width: 16, height: 8, rx: 1, fill: pal.wagons[(i + k + 1) % pal.wagons.length] })));
        g.push(el('circle', { cx: x + 7, cy: RAIL_Y + 3.5, r: 2.4, fill: '#1f2937' }));
        g.push(el('circle', { cx: x + car.w - 7, cy: RAIL_Y + 3.5, r: 2.4, fill: '#1f2937' }));
        return el('g', {}, g);
      });
      parts.push(el('g', {}, [
        el('g', { transform: `translate(${-offset} 0)` }, cars),
        el('animateMotion', { dur: `${TRAIN.dur}s`, begin: `${TRAIN.delay}s`, repeatCount: 'indefinite', path: `M-60,0 L${1600 + offset + 60},0` }),
      ]));
    }
    if (pal.bird) BIRD_FLOCKS.forEach((flock) => {
      const g = flock.birds.map(([bx, by], bi) => el('path', { class: 'skyline-bird', d: `M${bx - 6},${by} Q${bx - 3},${by - 5} ${bx},${by} Q${bx + 3},${by - 5} ${bx + 6},${by}`, style: `--flap-duration:${0.7 + (bi % 3) * 0.12}s;--flap-delay:-${(bi * 0.23) % 0.8}s` }));
      g.push(el('animateMotion', { dur: `${flock.dur}s`, begin: `${flock.delay}s`, repeatCount: 'indefinite', path: flock.path }));
      parts.push(el('g', { fill: 'none', stroke: pal.bird, 'stroke-width': 1.6, 'stroke-linecap': 'round' }, g));
    });
    {
      const g = [];
      if (pal.contrail) g.push(el('rect', { x: -230, y: -1.2, width: 212, height: 2.4, rx: 1.2, fill: `url(#${id('contrail')})` }));
      g.push(el('path', { d: 'M-14,0 L8,0 Q14,0 14,2.5 L-14,2.5 Z M-8,0 L-6,-7 L-2,-7 L2,0 Z M-10,2.5 L-14,8 L-9,8 L-3,2.5 Z', fill: pal.plane }));
      if (!pal.contrail) {
        g.push(el('circle', { class: 'skyline-beacon', cx: -12, cy: 0, r: 1.8, fill: '#ff5a4a' }));
        g.push(el('circle', { class: 'skyline-beacon', cx: -4, cy: -7, r: 1.8, fill: '#8bf0a0', style: '--beacon-delay:-1.2s' }));
      }
      g.push(el('animateMotion', { dur: `${PLANE.dur}s`, begin: `${PLANE.delay}s`, repeatCount: 'indefinite', path: PLANE.path, rotate: 'auto' }));
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
