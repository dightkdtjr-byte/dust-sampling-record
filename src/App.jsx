import React, { useEffect, useRef, useState } from 'react';
import JSZip from 'jszip';
import mainLogoImage from './assets/mainlogo.png';

const DEFAULT_NOZZLE_SET = [
  { num: 4, d: 3.21 }, { num: 5, d: 3.97 }, { num: 6, d: 4.79 },
  { num: 7, d: 5.54 }, { num: 8, d: 6.28 }, { num: 9, d: 7.10 },
  { num: 10, d: 7.86 }, { num: 11, d: 8.40 }, { num: 12, d: 9.24 },
  { num: 13, d: 10.18 }, { num: 14, d: 10.87 }, { num: 15, d: 11.70 },
  { num: 16, d: 12.56 }, { num: 17, d: 13.48 }, { num: 18, d: 14.30 },
  { num: 20, d: 15.95 }, { num: 22, d: 17.51 }, { num: 26, d: 20.72 }
];

const DEFAULT_SAMPLERS = [
  { id: 1, name: '샘플러 1호기', yd: '0.991', dHAt: '47.6' },
  { id: 2, name: '샘플러 2호기', yd: '1.005', dHAt: '40.2' },
  { id: 3, name: '샘플러 3호기', yd: '0.988', dHAt: '39.5' },
  { id: 4, name: '샘플러 4호기', yd: '1.012', dHAt: '42.1' },
  { id: 5, name: '샘플러 5호기', yd: '0.998', dHAt: '41.5' }
];

const SHEET_MENU = [
  {
    id: 'dust',
    title: '먼지시료채취기록부',
    desc: '먼지시료채취 및 등속흡인 기록',
    toneClass: 'from-emerald-200/60 to-teal-200/20',
    hoverTextClass: 'group-hover:text-emerald-700',
  },
  {
    id: 'metal',
    title: '금속시료채취기록부',
    desc: '금속시료채취 및 등속흡인 기록',
    toneClass: 'from-yellow-200/70 to-amber-200/20',
    hoverTextClass: 'group-hover:text-yellow-700',
  },
  {
    id: 'mercury',
    title: '수은시료채취기록부',
    desc: '수은시료채취 및 등속흡인 기록',
    toneClass: 'from-blue-200/70 to-sky-200/20',
    hoverTextClass: 'group-hover:text-blue-700',
  },
  {
    id: 'pahs',
    title: 'PAHs시료채취기록부',
    desc: 'PAHs시료채취 및 등속흡인 기록',
    toneClass: 'from-purple-200/70 to-violet-200/20',
    hoverTextClass: 'group-hover:text-purple-700',
  },
];

const SKYLINE_FLYERS = [
  { id: 148, name: '신뇽', top: '40%', size: 54, direction: 'right', duration: 34, delay: 6, bob: 3.2 },
  { id: 149, name: '망나뇽', top: '46%', size: 62, direction: 'left', duration: 29, delay: 2, bob: 2.8 },
  { id: 146, name: '파이어', top: '28%', size: 66, direction: 'right', duration: 30, delay: 4, bob: 3.0 },
  { id: 145, name: '썬더', top: '22%', size: 64, direction: 'left', duration: 27, delay: 7, bob: 2.6 },
  { id: 144, name: '프리져', top: '12%', size: 64, direction: 'right', duration: 33, delay: 1, bob: 3.3 },
  { id: 142, name: '프테라', top: '33%', size: 60, direction: 'right', duration: 26, delay: 11, bob: 2.7 },
  { id: 6, name: '리자몽', top: '52%', size: 68, direction: 'left', duration: 32, delay: 5, bob: 3.1 },
  { id: 15, name: '독침붕', top: '36%', size: 50, direction: 'right', duration: 22, delay: 8, bob: 2.2 },
  { id: 18, name: '피존튜', top: '58%', size: 58, direction: 'left', duration: 28, delay: 13, bob: 2.5 },
  { id: 150, name: '뮤츠', top: '19%', size: 56, direction: 'right', duration: 31, delay: 10, bob: 2.9 },
];
const SKYLINE_FLYER_UNIFIED_SIZE = 96;

const LEGENDARY_SKYLINE_CHOICES = [
  {
    id: 249,
    name: '루기아',
    size: 420,
    url: 'https://raw.githubusercontent.com/PokeAPI/sprites/master/sprites/pokemon/other/home/249.png',
  },
  {
    id: 250,
    name: '칠색조',
    size: 400,
    url: 'https://raw.githubusercontent.com/PokeAPI/sprites/master/sprites/pokemon/other/home/250.png',
  },
  {
    id: 151,
    name: '뮤',
    size: 380,
    url: 'https://raw.githubusercontent.com/PokeAPI/sprites/master/sprites/pokemon/other/home/151.png',
  },
];

const pickRandomLegendaryFlight = () => {
  const chosen = LEGENDARY_SKYLINE_CHOICES[Math.floor(Math.random() * LEGENDARY_SKYLINE_CHOICES.length)];
  const duration = 13 + Math.random() * 4;
  return {
    ...chosen,
    direction: Math.random() < 0.5 ? 'right' : 'left',
    duration,
    // 산/굴뚝 라인(하단부)에서 시작하도록 위치를 아래쪽으로 고정
    top: `${58 + Math.random() * 12}%`,
    left: `${16 + Math.random() * 62}%`,
    startOffset: 0,
    key: `${Date.now()}-${Math.random().toString(36).slice(2, 7)}`,
  };
};

const getSkyPhaseByHour = (date = new Date()) => {
  const hour = date.getHours();
  return hour >= 9 && hour < 17 ? 'day' : 'night';
};

const seededNoise = (seed) => {
  const x = Math.sin(seed * 12.9898) * 43758.5453;
  return x - Math.floor(x);
};

const buildSmokeParticleStyle = (stackId, smokeIdx) => {
  const stackNum = parseInt(String(stackId).replace(/\D/g, ''), 10) || 1;
  const seedBase = stackNum * 97 + smokeIdx * 31;
  const riseDuration = 16.5 + seededNoise(seedBase + 4) * 8.5;
  const driftDir = seededNoise(seedBase + 2) > 0.5 ? 1 : -1;

  return {
    '--size': `${44 + seededNoise(seedBase + 1) * 58}px`,
    '--drift': `${driftDir * (22 + seededNoise(seedBase + 3) * 76)}px`,
    '--rise-duration': `${riseDuration}s`,
    '--rise-distance': `${132 + seededNoise(seedBase + 5) * 84}vh`,
    '--max-opacity': `${0.26 + seededNoise(seedBase + 6) * 0.2}`,
    '--start-scale': `${0.56 + seededNoise(seedBase + 7) * 0.18}`,
    '--end-scale': `${1.54 + seededNoise(seedBase + 8) * 0.54}`,
    '--tilt-start': `${-8 + seededNoise(seedBase + 9) * 6}deg`,
    '--tilt-end': `${4 + seededNoise(seedBase + 10) * 8}deg`,
    animationDelay: `-${seededNoise(seedBase + 11) * riseDuration}s`,
  };
};

const SKYLINE_STARS = Array.from({ length: 46 }, (_, idx) => ({
  id: `star-${idx}`,
  left: `${2 + ((idx * 17) % 96)}%`,
  top: `${4 + ((idx * 13) % 44)}%`,
  size: 1.2 + (idx % 4) * 0.55,
  duration: 2.4 + (idx % 5) * 0.65,
  delay: (idx % 7) * 0.35,
}));

const SKYLINE_MOUNTAINS = [
  { id: 'mt-1', left: '-8%', width: '38%', height: 182 },
  { id: 'mt-2', left: '20%', width: '34%', height: 154 },
  { id: 'mt-3', left: '42%', width: '36%', height: 196 },
  { id: 'mt-4', left: '68%', width: '34%', height: 166 },
];

const SKYLINE_TREE_BELT = Array.from({ length: 40 }, (_, idx) => ({
  id: `tree-${idx}`,
  left: `${(idx * 2.6) % 100}%`,
  width: 14 + ((idx * 7) % 10),
  height: 24 + ((idx * 11) % 26),
}));

const SKYLINE_FRONT_TREE_BELT = Array.from({ length: 32 }, (_, idx) => ({
  id: `front-tree-${idx}`,
  left: `${(idx * 3.05 + 1.2) % 100}%`,
  width: 18 + ((idx * 5) % 12),
  height: 28 + ((idx * 9) % 24),
}));

const PIXEL_TREE_VARIANTS = [
  {
    cells: [
      [2, 6], [3, 6],
      [1, 5], [2, 5], [3, 5], [4, 5],
      [0, 4], [1, 4], [2, 4], [3, 4], [4, 4], [5, 4],
      [1, 3], [2, 3], [3, 3], [4, 3],
      [2, 2], [3, 2],
    ],
    trunk: { x: 3, h: 2, baseX: 2, baseW: 3 },
  },
  {
    cells: [
      [3, 7],
      [2, 6], [3, 6], [4, 6],
      [1, 5], [2, 5], [3, 5], [4, 5], [5, 5],
      [0, 4], [1, 4], [2, 4], [3, 4], [4, 4], [5, 4], [6, 4],
      [1, 3], [2, 3], [3, 3], [4, 3], [5, 3],
      [2, 2], [3, 2], [4, 2],
    ],
    trunk: { x: 3, h: 2, baseX: 2, baseW: 3 },
  },
  {
    cells: [
      [3, 8],
      [2, 7], [3, 7], [4, 7],
      [2, 6], [3, 6], [4, 6],
      [1, 5], [2, 5], [3, 5], [4, 5], [5, 5],
      [1, 4], [2, 4], [3, 4], [4, 4], [5, 4],
      [0, 3], [1, 3], [2, 3], [3, 3], [4, 3], [5, 3], [6, 3],
      [2, 2], [3, 2], [4, 2],
    ],
    trunk: { x: 3, h: 3, baseX: 2, baseW: 3 },
  },
  {
    cells: [
      [3, 9],
      [3, 8],
      [2, 7], [3, 7], [4, 7],
      [2, 6], [3, 6], [4, 6],
      [1, 5], [2, 5], [3, 5], [4, 5], [5, 5],
      [2, 4], [3, 4], [4, 4],
      [3, 3],
    ],
    trunk: { x: 3, h: 3, baseX: 2, baseW: 3 },
  },
  {
    cells: [
      [2, 7], [3, 7], [4, 7],
      [1, 6], [2, 6], [3, 6], [4, 6], [5, 6],
      [0, 5], [1, 5], [2, 5], [3, 5], [4, 5], [5, 5], [6, 5],
      [0, 4], [1, 4], [2, 4], [3, 4], [4, 4], [5, 4], [6, 4],
      [1, 3], [2, 3], [3, 3], [4, 3], [5, 3],
      [2, 2], [3, 2], [4, 2],
    ],
    trunk: { x: 3, h: 2, baseX: 2, baseW: 3 },
  },
  {
    cells: [
      [2, 8], [3, 8],
      [2, 7], [3, 7], [4, 7],
      [1, 6], [2, 6], [3, 6], [4, 6], [5, 6],
      [1, 5], [2, 5], [3, 5], [4, 5], [5, 5],
      [0, 4], [1, 4], [2, 4], [3, 4], [4, 4],
      [1, 3], [2, 3], [3, 3],
    ],
    trunk: { x: 2, h: 3, baseX: 1, baseW: 3 },
  },
  {
    cells: [
      [2, 5], [3, 5],
      [1, 4], [2, 4], [3, 4], [4, 4],
      [0, 3], [1, 3], [2, 3], [3, 3], [4, 3], [5, 3],
      [1, 2], [2, 2], [3, 2], [4, 2],
    ],
    trunk: { x: 3, h: 2, baseX: 2, baseW: 2 },
  },
];

const PIXEL_TREE_PALETTES = {
  night: [
    { main: '#1f4d3b', alt: '#2a6b52', trunk: '#4f3c2a' },
    { main: '#2d4d73', alt: '#4672a4', trunk: '#4a3a2a' },
    { main: '#435226', alt: '#62763a', trunk: '#4b3a29' },
    { main: '#4c3f69', alt: '#6d58a0', trunk: '#4a3728' },
    { main: '#2f5f55', alt: '#43927e', trunk: '#5a402b' },
  ],
  day: [
    { main: '#4f8b6b', alt: '#79b58e', trunk: '#6f5339' },
    { main: '#5b88b1', alt: '#86b5dc', trunk: '#6c4f38' },
    { main: '#7c9a46', alt: '#a8c86a', trunk: '#755338' },
    { main: '#8c75b5', alt: '#b49ad9', trunk: '#6c4c36' },
    { main: '#5ea78a', alt: '#84d1b0', trunk: '#7a593d' },
  ],
};

const getPixelTreePalette = (isNightSky, seedIndex) => {
  const palettes = isNightSky ? PIXEL_TREE_PALETTES.night : PIXEL_TREE_PALETTES.day;
  return palettes[seedIndex % palettes.length];
};

const SKY_PREVIEW_OPTIONS = [
  { id: 'auto', label: '자동' },
  { id: 'day', label: '낮' },
  { id: 'night', label: '밤' },
];

const PIXEL_MOUNTAIN_CLIP_PATH = 'polygon(0 100%, 0 78%, 8% 78%, 8% 66%, 16% 66%, 16% 54%, 24% 54%, 24% 42%, 32% 42%, 32% 30%, 40% 30%, 40% 18%, 48% 18%, 48% 8%, 56% 8%, 56% 18%, 64% 18%, 64% 30%, 72% 30%, 72% 42%, 80% 42%, 80% 54%, 88% 54%, 88% 66%, 96% 66%, 96% 78%, 100% 78%, 100% 100%)';
const PIXEL_HILL_CLIP_PATH = 'polygon(0 100%, 0 72%, 6% 72%, 6% 64%, 14% 64%, 14% 56%, 22% 56%, 22% 62%, 30% 62%, 30% 52%, 38% 52%, 38% 44%, 46% 44%, 46% 50%, 54% 50%, 54% 42%, 62% 42%, 62% 50%, 70% 50%, 70% 60%, 78% 60%, 78% 54%, 86% 54%, 86% 66%, 94% 66%, 94% 74%, 100% 74%, 100% 100%)';

const getPixelMountainFill = (isNightSky) => (
  isNightSky
    ? 'linear-gradient(180deg, #5f79a9 0 12%, #4d6693 12% 24%, #3f557f 24% 36%, #34466a 36% 48%, #293958 48% 62%, #1f2c48 62% 100%)'
    : 'linear-gradient(180deg, #c9daf0 0 12%, #b4cbe8 12% 24%, #9fbce0 24% 36%, #8bacd3 36% 48%, #789ac4 48% 62%, #6789b4 62% 100%)'
);

const getPixelHillFill = (isNightSky) => (
  isNightSky
    ? 'linear-gradient(180deg, #36564a 0 18%, #2b4a3f 18% 36%, #233f36 36% 54%, #1a322b 54% 100%)'
    : 'linear-gradient(180deg, #6f9c7d 0 18%, #5f8b6f 18% 36%, #507c62 36% 54%, #446b55 54% 100%)'
);

const SKYLINE_BACK_TOWERS = [
  { id: 'bt-1', left: '3%', width: '5.4%', height: 86 },
  { id: 'bt-2', left: '9.2%', width: '5.9%', height: 98 },
  { id: 'bt-3', left: '16%', width: '6.8%', height: 112 },
  { id: 'bt-4', left: '24.1%', width: '7.4%', height: 128 },
  { id: 'bt-5', left: '32.7%', width: '6.6%', height: 96 },
  { id: 'bt-6', left: '40%', width: '8.2%', height: 138 },
  { id: 'bt-7', left: '49.5%', width: '6.8%', height: 122 },
  { id: 'bt-8', left: '57.4%', width: '7.6%', height: 104 },
  { id: 'bt-9', left: '66%', width: '7%', height: 114 },
  { id: 'bt-10', left: '74.2%', width: '8%', height: 132 },
  { id: 'bt-11', left: '83.4%', width: '7.2%', height: 102 },
];

const SKYLINE_FACTORY_BLOCKS = [
  { id: 'fb-1', left: '12%', width: '12%', height: 106, roof: 12, windows: 4 },
  { id: 'fb-2', left: '24%', width: '11%', height: 126, roof: 11, windows: 5 },
  { id: 'fb-3', left: '34.5%', width: '13%', height: 112, roof: 10, windows: 4 },
  { id: 'fb-4', left: '47%', width: '12.5%', height: 134, roof: 13, windows: 5 },
  { id: 'fb-5', left: '58.7%', width: '11%', height: 118, roof: 10, windows: 4 },
  { id: 'fb-6', left: '69%', width: '13%', height: 128, roof: 12, windows: 5 },
];

const SKYLINE_COOLING_TOWERS = [
  { id: 'ct-1', left: '79.5%', width: 92, height: 156 },
  { id: 'ct-2', left: '89%', width: 78, height: 140 },
];

const SKYLINE_STACKS = [
  { id: 'st-1', left: '18.5%', width: 36, height: 220, smokeCount: 10, variant: 'slate' },
  { id: 'st-2', left: '29.5%', width: 54, height: 286, smokeCount: 13, variant: 'brick' },
  { id: 'st-3', left: '43%', width: 48, height: 258, smokeCount: 12, variant: 'brick' },
  { id: 'st-4', left: '55.5%', width: 34, height: 206, smokeCount: 9, variant: 'slate' },
  { id: 'st-5', left: '64.8%', width: 50, height: 274, smokeCount: 12, variant: 'concrete' },
  { id: 'st-6', left: '73.8%', width: 30, height: 188, smokeCount: 9, variant: 'slate' },
];

const SKYLINE_FLYER_CSS = `
@keyframes skylineFlyRight {
  0% { transform: translateX(-8vw) scaleX(1); }
  100% { transform: translateX(108vw) scaleX(1); }
}
@keyframes skylineFlyLeft {
  0% { transform: translateX(108vw) scaleX(-1); }
  100% { transform: translateX(-8vw) scaleX(-1); }
}
@keyframes skylineBob {
  0% { transform: translateY(-6px); }
  100% { transform: translateY(6px); }
}
@keyframes skylineTwinkle {
  0%, 100% { opacity: 0.26; transform: scale(0.9); }
  50% { opacity: 0.95; transform: scale(1.08); }
}
@keyframes legendaryDiagonalRight {
  0% {
    transform: translateX(0) translateY(14vh) scaleX(1);
    opacity: 0;
  }
  12% {
    opacity: 0.94;
  }
  92% {
    opacity: 0.9;
  }
  100% {
    transform: translateX(34vw) translateY(-76vh) scaleX(1);
    opacity: 0;
  }
}
@keyframes legendaryDiagonalLeft {
  0% {
    transform: translateX(0) translateY(14vh) scaleX(-1);
    opacity: 0;
  }
  12% {
    opacity: 0.94;
  }
  92% {
    opacity: 0.9;
  }
  100% {
    transform: translateX(-34vw) translateY(-76vh) scaleX(-1);
    opacity: 0;
  }
}
.skyline-flyer {
  position: absolute;
  left: 0;
  will-change: transform;
}
.skyline-flyer-inner {
  animation: skylineBob var(--bob-duration, 3s) ease-in-out infinite alternate;
  filter: drop-shadow(0 8px 16px rgba(15, 23, 42, 0.4));
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
.legendary-flyer {
  position: absolute;
  left: 0;
  will-change: transform, opacity;
  pointer-events: none;
  opacity: 0;
}
.legendary-flyer img {
  display: block;
  filter: drop-shadow(0 18px 28px rgba(15, 23, 42, 0.32));
}
@keyframes stackHeartRise {
  0% {
    transform: translate(-50%, 0) scale(var(--start-scale, 0.58)) rotate(var(--tilt-start, -7deg));
    opacity: 0;
  }
  12% {
    opacity: var(--max-opacity, 0.24);
  }
  100% {
    transform: translate(calc(-50% + var(--drift, 0px)), calc(-1 * var(--rise-distance, 150vh))) scale(var(--end-scale, 1.86)) rotate(var(--tilt-end, 7deg));
    opacity: 0;
  }
}
.stack-heart-smoke {
  position: absolute;
  left: 50%;
  top: -38px;
  width: var(--size, 16px);
  height: var(--size, 16px);
  clip-path: polygon(50% 100%, 0 52%, 0 30%, 16% 12%, 34% 12%, 50% 27%, 66% 12%, 84% 12%, 100% 30%, 100% 52%);
  background: linear-gradient(135deg, rgba(248, 250, 252, 0.64), rgba(221, 214, 254, 0.48));
  border: 1px solid rgba(199, 210, 254, 0.48);
  opacity: 0;
  pointer-events: none;
  z-index: 1;
  mix-blend-mode: normal;
  filter: blur(0.45px) drop-shadow(0 0 14px rgba(167, 139, 250, 0.28));
  animation: stackHeartRise var(--rise-duration, 18s) ease-out infinite;
}
`;

const SHEET_THEMES = {
  dust: {
    selection: 'selection:bg-emerald-100',
    pageBg: 'bg-emerald-50',
    pageTint: 'bg-emerald-100/40',
    headerBorder: 'border-t-emerald-600',
    headerIcon: 'text-emerald-600',
    focusRing: 'focus:ring-emerald-500',
    primaryButton: 'bg-emerald-600 hover:bg-emerald-700',
    primaryButtonSoft: 'bg-emerald-600 hover:bg-emerald-500',
    accentText: 'text-emerald-400',
  },
  metal: {
    selection: 'selection:bg-yellow-100',
    pageBg: 'bg-yellow-50',
    pageTint: 'bg-yellow-100/50',
    headerBorder: 'border-t-yellow-500',
    headerIcon: 'text-yellow-600',
    focusRing: 'focus:ring-yellow-500',
    primaryButton: 'bg-yellow-600 hover:bg-yellow-700',
    primaryButtonSoft: 'bg-yellow-600 hover:bg-yellow-500',
    accentText: 'text-yellow-400',
  },
  mercury: {
    selection: 'selection:bg-blue-100',
    pageBg: 'bg-blue-50',
    pageTint: 'bg-blue-100/50',
    headerBorder: 'border-t-blue-600',
    headerIcon: 'text-blue-600',
    focusRing: 'focus:ring-blue-500',
    primaryButton: 'bg-blue-600 hover:bg-blue-700',
    primaryButtonSoft: 'bg-blue-600 hover:bg-blue-500',
    accentText: 'text-blue-300',
  },
  pahs: {
    selection: 'selection:bg-purple-100',
    pageBg: 'bg-purple-50',
    pageTint: 'bg-purple-100/50',
    headerBorder: 'border-t-purple-600',
    headerIcon: 'text-purple-600',
    focusRing: 'focus:ring-purple-500',
    primaryButton: 'bg-purple-600 hover:bg-purple-700',
    primaryButtonSoft: 'bg-purple-600 hover:bg-purple-500',
    accentText: 'text-purple-300',
  },
};

const THEME_OVERRIDE_CSS = `
.theme-metal [class*="bg-emerald-50"], .theme-metal [class*="bg-teal-50"], .theme-metal [class*="bg-cyan-50"] { background-color: #fefce8 !important; }
.theme-metal [class*="bg-emerald-100"], .theme-metal [class*="bg-teal-100"], .theme-metal [class*="bg-cyan-100"] { background-color: #fef3c7 !important; }
.theme-metal [class*="bg-emerald-200"], .theme-metal [class*="bg-teal-200"], .theme-metal [class*="bg-cyan-200"] { background-color: #fde68a !important; }
.theme-metal [class*="bg-emerald-600"], .theme-metal [class*="bg-teal-600"], .theme-metal [class*="bg-cyan-600"] { background-color: #ca8a04 !important; }
.theme-metal [class*="bg-emerald-700"], .theme-metal [class*="bg-teal-700"], .theme-metal [class*="bg-cyan-700"] { background-color: #a16207 !important; }
.theme-metal [class*="text-emerald-"], .theme-metal [class*="text-teal-"], .theme-metal [class*="text-cyan-"] { color: #a16207 !important; }
.theme-metal [class*="border-emerald-"], .theme-metal [class*="border-teal-"], .theme-metal [class*="border-cyan-"] { border-color: #facc15 !important; }
.theme-metal [class*="focus:ring-emerald-500"]:focus { --tw-ring-color: rgb(234 179 8 / 0.55) !important; }

.theme-mercury [class*="bg-emerald-50"], .theme-mercury [class*="bg-teal-50"], .theme-mercury [class*="bg-cyan-50"] { background-color: #eff6ff !important; }
.theme-mercury [class*="bg-emerald-100"], .theme-mercury [class*="bg-teal-100"], .theme-mercury [class*="bg-cyan-100"] { background-color: #dbeafe !important; }
.theme-mercury [class*="bg-emerald-200"], .theme-mercury [class*="bg-teal-200"], .theme-mercury [class*="bg-cyan-200"] { background-color: #bfdbfe !important; }
.theme-mercury [class*="bg-emerald-600"], .theme-mercury [class*="bg-teal-600"], .theme-mercury [class*="bg-cyan-600"] { background-color: #2563eb !important; }
.theme-mercury [class*="bg-emerald-700"], .theme-mercury [class*="bg-teal-700"], .theme-mercury [class*="bg-cyan-700"] { background-color: #1d4ed8 !important; }
.theme-mercury [class*="text-emerald-"], .theme-mercury [class*="text-teal-"], .theme-mercury [class*="text-cyan-"] { color: #1d4ed8 !important; }
.theme-mercury [class*="border-emerald-"], .theme-mercury [class*="border-teal-"], .theme-mercury [class*="border-cyan-"] { border-color: #60a5fa !important; }
.theme-mercury [class*="focus:ring-emerald-500"]:focus { --tw-ring-color: rgb(59 130 246 / 0.55) !important; }

.theme-pahs [class*="bg-emerald-50"], .theme-pahs [class*="bg-teal-50"], .theme-pahs [class*="bg-cyan-50"] { background-color: #faf5ff !important; }
.theme-pahs [class*="bg-emerald-100"], .theme-pahs [class*="bg-teal-100"], .theme-pahs [class*="bg-cyan-100"] { background-color: #f3e8ff !important; }
.theme-pahs [class*="bg-emerald-200"], .theme-pahs [class*="bg-teal-200"], .theme-pahs [class*="bg-cyan-200"] { background-color: #e9d5ff !important; }
.theme-pahs [class*="bg-emerald-600"], .theme-pahs [class*="bg-teal-600"], .theme-pahs [class*="bg-cyan-600"] { background-color: #9333ea !important; }
.theme-pahs [class*="bg-emerald-700"], .theme-pahs [class*="bg-teal-700"], .theme-pahs [class*="bg-cyan-700"] { background-color: #7e22ce !important; }
.theme-pahs [class*="text-emerald-"], .theme-pahs [class*="text-teal-"], .theme-pahs [class*="text-cyan-"] { color: #7e22ce !important; }
.theme-pahs [class*="border-emerald-"], .theme-pahs [class*="border-teal-"], .theme-pahs [class*="border-cyan-"] { border-color: #c084fc !important; }
.theme-pahs [class*="focus:ring-emerald-500"]:focus { --tw-ring-color: rgb(168 85 247 / 0.55) !important; }
`;

const STORAGE_KEYS = {
  nozzles: 'dust-sampling.nozzles.v1',
  samplers: 'dust-sampling.samplers.v1',
  secureUsers: 'dust-sampling.secure-users.v2',
  activeUser: 'dust-sampling.secure-active-user.v2',
  lastLoginId: 'dust-sampling.last-login-id.v1',
  sessionAuth: 'dust-sampling.session-auth.v1',
};

const persistSessionAuth = (userId, password) => {
  try {
    localStorage.setItem(
      STORAGE_KEYS.sessionAuth,
      JSON.stringify({ userId: String(userId || ''), password: String(password || '') })
    );
  } catch (error) {
    console.error('로그인 세션 저장 실패:', error);
  }
};

const readSessionAuth = () => {
  try {
    const raw = localStorage.getItem(STORAGE_KEYS.sessionAuth);
    if (!raw) return null;
    const parsed = JSON.parse(raw);
    const userId = String(parsed?.userId || '').trim();
    const password = typeof parsed?.password === 'string' ? parsed.password : '';
    if (!userId || !password) return null;
    return { userId, password };
  } catch (error) {
    console.error('로그인 세션 로딩 실패:', error);
    return null;
  }
};

const clearSessionAuth = () => {
  try {
    localStorage.removeItem(STORAGE_KEYS.sessionAuth);
  } catch (error) {
    console.error('로그인 세션 삭제 실패:', error);
  }
};

const cloneDefaultNozzles = () => DEFAULT_NOZZLE_SET.map(item => ({ ...item }));
const cloneDefaultSamplers = () => DEFAULT_SAMPLERS.map(item => ({ ...item }));

const sanitizeNozzleSet = (value) => {
  if (!Array.isArray(value)) return null;
  return value
    .filter(item => item && typeof item === 'object')
    .map(item => ({
      num: item.num === undefined || item.num === null ? '' : String(item.num),
      d: item.d === undefined || item.d === null ? '' : String(item.d),
    }));
};

const sanitizeSamplers = (value) => {
  if (!Array.isArray(value)) return null;
  return value
    .filter(item => item && typeof item === 'object')
    .map((item, idx) => {
      const parsedId = parseInt(item.id, 10);
      return {
        id: Number.isFinite(parsedId) ? parsedId : idx + 1,
        name: typeof item.name === 'string' && item.name.trim() ? item.name : `샘플러 ${idx + 1}호기`,
        yd: item.yd === undefined || item.yd === null ? '' : String(item.yd),
        dHAt: item.dHAt === undefined || item.dHAt === null ? '' : String(item.dHAt),
      };
    });
};

const sanitizeSecureUsers = (value) => {
  if (!value || typeof value !== 'object' || Array.isArray(value)) return null;
  const entries = Object.entries(value);
  if (entries.length === 0) return {};

  const result = {};
  entries.forEach(([user, record]) => {
    const userName = String(user || '').trim();
    if (!userName) return;
    if (!record || typeof record !== 'object' || Array.isArray(record)) return;

    const normalized = {
      salt: typeof record.salt === 'string' ? record.salt : '',
      verifierIv: typeof record.verifierIv === 'string' ? record.verifierIv : '',
      verifierCipher: typeof record.verifierCipher === 'string' ? record.verifierCipher : '',
      dataIv: typeof record.dataIv === 'string' ? record.dataIv : '',
      dataCipher: typeof record.dataCipher === 'string' ? record.dataCipher : '',
      updatedAt: typeof record.updatedAt === 'string' ? record.updatedAt : '',
      hintQuestion: typeof record.hintQuestion === 'string' ? record.hintQuestion : '',
      hintAnswerHash: typeof record.hintAnswerHash === 'string' ? record.hintAnswerHash : '',
      avatarUrl: typeof record.avatarUrl === 'string' ? record.avatarUrl : '',
      nickname: typeof record.nickname === 'string' ? record.nickname : '',
    };

    if (!normalized.salt || !normalized.verifierIv || !normalized.verifierCipher || !normalized.dataIv || !normalized.dataCipher) return;
    result[userName] = normalized;
  });

  return result;
};

const PASSWORD_RULE = /^(?=.*[A-Za-z])(?=.*\d)(?=.*[^A-Za-z0-9\s])\S{8,64}$/;
const normalizeHintText = (value) => String(value || '').trim().toLowerCase();
const ADMIN_USER_ID = 'dightkdtjr';
const ADMIN_PASSWORD = 'ss22127016!';
const CHARIZARD_POKEMON_ID = 6;
const BETA_USER_ID = 'user';
const BETA_USER_PASSWORD = 'beta-user';
const LEGACY_BETA_USER_ID = 'user1';
const LEGACY_BETA_USER_PASSWORD = 'beta-user1';
const MAGIKARP_POKEMON_ID = 129;

const textEncoder = new TextEncoder();
const textDecoder = new TextDecoder();

const bytesToBase64 = (bytes) => {
  const arr = bytes instanceof Uint8Array ? bytes : new Uint8Array(bytes);
  let binary = '';
  const chunkSize = 0x8000;
  for (let i = 0; i < arr.length; i += chunkSize) {
    binary += String.fromCharCode(...arr.subarray(i, i + chunkSize));
  }
  return btoa(binary);
};

const base64ToBytes = (b64) => {
  const binary = atob(b64);
  const bytes = new Uint8Array(binary.length);
  for (let i = 0; i < binary.length; i++) {
    bytes[i] = binary.charCodeAt(i);
  }
  return bytes;
};

const deriveVaultKey = async (password, saltB64) => {
  const baseKey = await crypto.subtle.importKey('raw', textEncoder.encode(password), 'PBKDF2', false, ['deriveKey']);
  return crypto.subtle.deriveKey(
    {
      name: 'PBKDF2',
      salt: base64ToBytes(saltB64),
      iterations: 120000,
      hash: 'SHA-256',
    },
    baseKey,
    { name: 'AES-GCM', length: 256 },
    false,
    ['encrypt', 'decrypt']
  );
};

const encryptTextWithKey = async (key, plainText) => {
  const iv = crypto.getRandomValues(new Uint8Array(12));
  const encrypted = await crypto.subtle.encrypt(
    { name: 'AES-GCM', iv },
    key,
    textEncoder.encode(plainText)
  );
  return {
    iv: bytesToBase64(iv),
    cipher: bytesToBase64(new Uint8Array(encrypted)),
  };
};

const decryptTextWithKey = async (key, ivB64, cipherB64) => {
  const decrypted = await crypto.subtle.decrypt(
    { name: 'AES-GCM', iv: base64ToBytes(ivB64) },
    key,
    base64ToBytes(cipherB64)
  );
  return textDecoder.decode(decrypted);
};

const encryptJsonWithKey = async (key, value) => {
  return encryptTextWithKey(key, JSON.stringify(value));
};

const decryptJsonWithKey = async (key, ivB64, cipherB64) => {
  const text = await decryptTextWithKey(key, ivB64, cipherB64);
  return JSON.parse(text);
};

const hashHintAnswer = async (answerText) => {
  const normalized = normalizeHintText(answerText);
  if (!normalized) return '';
  const digest = await crypto.subtle.digest('SHA-256', textEncoder.encode(normalized));
  return bytesToBase64(new Uint8Array(digest));
};

const getFixedAvatarUrlForUser = (userId) => {
  const normalizedId = String(userId || '').trim().toLowerCase();
  if (normalizedId === ADMIN_USER_ID) {
    return `https://raw.githubusercontent.com/PokeAPI/sprites/master/sprites/pokemon/${CHARIZARD_POKEMON_ID}.png`;
  }
  if (normalizedId === BETA_USER_ID) {
    return `https://raw.githubusercontent.com/PokeAPI/sprites/master/sprites/pokemon/${MAGIKARP_POKEMON_ID}.png`;
  }

  if (!normalizedId) {
    return `https://raw.githubusercontent.com/PokeAPI/sprites/master/sprites/pokemon/1.png`;
  }

  let hash = 0;
  for (let i = 0; i < normalizedId.length; i++) {
    hash = (hash * 31 + normalizedId.charCodeAt(i)) % 2147483647;
  }
  const pokemonId = (Math.abs(hash) % 151) + 1;
  return `https://raw.githubusercontent.com/PokeAPI/sprites/master/sprites/pokemon/${pokemonId}.png`;
};

// 자바스크립트 부동소수점 오차 방지용 안전한 반올림 함수 (K * dp 오차 보정용)
const roundDH = (val) => {
  return (Math.round((val + Number.EPSILON) * 100) / 100).toFixed(2);
};

const calcExpectedOrificeDH = (kVal, dpVal) => {
  const kNum = parseFloat(kVal);
  const dpNum = parseFloat(dpVal);
  if (!Number.isFinite(kNum) || !Number.isFinite(dpNum)) return '-';
  // 엑셀과 동일하게 원값으로 곱한 뒤 최종값만 소수 둘째자리로 반올림
  return roundDH(kNum * dpNum);
};

const formatKFactorDisplay = (value) => {
  const num = parseFloat(value);
  return Number.isFinite(num) ? num.toFixed(2) : '-';
};

const SignatureBadge = () => (
  <div className="fixed right-4 bottom-4 z-30 px-3 py-1.5 rounded-full border border-slate-300 bg-white/95 backdrop-blur shadow-sm text-xs font-bold text-slate-700">
    sangseok
  </div>
);

export default function App() {
  const logoCandidates = [
    mainLogoImage,
    `${import.meta.env.BASE_URL}mainlogo.png`,
    `${import.meta.env.BASE_URL}app-logo.png`,
    `${import.meta.env.BASE_URL}mainlogo.jpg`,
    `${import.meta.env.BASE_URL}app-logo.jpg`,
  ];
  const [logoIndex, setLogoIndex] = useState(0);
  const mainLogoSrc = logoCandidates[Math.min(logoIndex, logoCandidates.length - 1)];
  const dustTemplateUrl = `${import.meta.env.BASE_URL}templates/dust-template.xlsm`;
  const [selectedSheet, setSelectedSheet] = useState('');
  const [nozzleSet, setNozzleSet] = useState(cloneDefaultNozzles);
  const [samplers, setSamplers] = useState(cloneDefaultSamplers);

  const [formData, setFormData] = useState({
    date: '', company: '', location: '', sampler: '',
    weather: '맑음', atmTemp: '20', atmPressure: '760',
    totalStackDepth: '', flangeLength: '', stackDiameter: '', pitotFactor: '0.84',
    standardO2: '',
    gasAnalyzer: [
      { time: '', o2: '21.00', co2: '0.00', co: '0.00', nox: '0.00', sox: '0.00' },
      { time: '', o2: '21.00', co2: '0.00', co: '0.00', nox: '0.00', sox: '0.00' },
      { time: '', o2: '21.00', co2: '0.00', co: '0.00', nox: '0.00', sox: '0.00' }
    ],
    moistureValues: ['', '', '', '', ''],
    impingers: [
      { id: 1, initial: '', final: '' }, { id: 2, initial: '', final: '' },
      { id: 3, initial: '', final: '' }, { id: 4, initial: '', final: '' },
    ],
    // 동정압 측정 시 가스미터 온도 (예비조사 Tm)
    traverseTmIn: '', traverseTmOut: '',
    points: [
      { id: 1, tp: '', sp: '', dp: '', ts: '' }, { id: 2, tp: '', sp: '', dp: '', ts: '' },
      { id: 3, tp: '', sp: '', dp: '', ts: '' }, { id: 4, tp: '', sp: '', dp: '', ts: '' },
      { id: 5, tp: '', sp: '', dp: '', ts: '' },
    ],
    samplerId: '', targetVolume: '1000', planSamplingTime: '65',
    gasMeterFactor: '0.991', deltaHAt: '47.6',
    recommendedNozzleNum: '', recommendedNozzleDia: '', usedNozzleNum: '',
    nozzleDiameter: '', kFactor: '',
    gasMeters: Array.from({ length: 6 }, (_, i) => ({
      id: i, pointNum: i === 0 ? '시작' : '', time: i === 0 ? '0' : '', stackTemp: '', 
      dp: '', pressure: '', volume: '', tmIn: '', tmOut: '', vacuum: '', impingerTemp: '' 
    })),
    filterId: '', filterInitial: '', filterFinal: '', remarks: ''
  });

  const [secureUsers, setSecureUsers] = useState({});
  const [activeUser, setActiveUser] = useState('');
  const [activeUserReports, setActiveUserReports] = useState([]);
  const [isUserUnlocked, setIsUserUnlocked] = useState(false);
  const [loginUserId, setLoginUserId] = useState(BETA_USER_ID);
  const [loginPassword, setLoginPassword] = useState('');
  const [newUserName, setNewUserName] = useState('');
  const [newUserPassword, setNewUserPassword] = useState('');
  const [newUserPasswordConfirm, setNewUserPasswordConfirm] = useState('');
  const [newUserHintQuestion, setNewUserHintQuestion] = useState('');
  const [newUserHintAnswer, setNewUserHintAnswer] = useState('');
  const [recoverHintAnswer, setRecoverHintAnswer] = useState('');
  const [recoveredUserIds, setRecoveredUserIds] = useState([]);
  const [recoverHintUserId, setRecoverHintUserId] = useState('');
  const [recoveredHintQuestion, setRecoveredHintQuestion] = useState('');
  const [profileAvatarUrl, setProfileAvatarUrl] = useState('');
  const [profileNickname, setProfileNickname] = useState('');
  const [authModal, setAuthModal] = useState('');
  const [isStorageHydrated, setIsStorageHydrated] = useState(false);
  const vaultKeyRef = useRef({});
  const [recommendations, setRecommendations] = useState(null);
  const [legendaryFlight, setLegendaryFlight] = useState(() => pickRandomLegendaryFlight());
  const [skyPhase, setSkyPhase] = useState(() => getSkyPhaseByHour());
  const [skyPreviewMode, setSkyPreviewMode] = useState('auto');
  const [lowSpecMode, setLowSpecMode] = useState(() => {
    if (typeof window === 'undefined') return false;
    return /iPhone|iPad|Android|Mobile/i.test(window.navigator.userAgent);
  });

  useEffect(() => {
    let cancelled = false;

    const bootstrap = async () => {
      try {
        const rawNozzles = localStorage.getItem(STORAGE_KEYS.nozzles);
        const rawSamplers = localStorage.getItem(STORAGE_KEYS.samplers);
        const rawSecureUsers = localStorage.getItem(STORAGE_KEYS.secureUsers);
        const rawLastLoginId = localStorage.getItem(STORAGE_KEYS.lastLoginId);

        if (rawNozzles) {
          const parsedNozzles = sanitizeNozzleSet(JSON.parse(rawNozzles));
          if (!cancelled && parsedNozzles) setNozzleSet(parsedNozzles);
        }

        if (rawSamplers) {
          const parsedSamplers = sanitizeSamplers(JSON.parse(rawSamplers));
          if (!cancelled && parsedSamplers && parsedSamplers.length > 0) setSamplers(parsedSamplers);
        }

        const parsedUsers = rawSecureUsers ? sanitizeSecureUsers(JSON.parse(rawSecureUsers)) : null;
        let normalizedUsers = parsedUsers || {};
        if (!cancelled) setSecureUsers(normalizedUsers);

        let restored = false;
        const savedSession = readSessionAuth();
        if (savedSession) {
          const { userId, password } = savedSession;
          const targetUser = normalizedUsers[userId];

          if (targetUser) {
            try {
              const key = await deriveVaultKey(password, targetUser.salt);
              const verifierText = await decryptTextWithKey(key, targetUser.verifierIv, targetUser.verifierCipher);
              if (verifierText === `vault:${userId}:verified`) {
                const decrypted = await decryptJsonWithKey(key, targetUser.dataIv, targetUser.dataCipher);
                const reports = Array.isArray(decrypted) ? decrypted.filter(item => item && typeof item === 'object') : [];
                const fixedAvatarUrl = getFixedAvatarUrlForUser(userId);

                if (targetUser.avatarUrl !== fixedAvatarUrl) {
                  normalizedUsers = {
                    ...normalizedUsers,
                    [userId]: {
                      ...targetUser,
                      avatarUrl: fixedAvatarUrl,
                      updatedAt: new Date().toISOString(),
                    },
                  };
                  if (!cancelled) setSecureUsers(normalizedUsers);
                }

                if (!cancelled) {
                  openUserSession(userId, key, reports, normalizedUsers[userId]);
                  restored = true;
                }
              }
            } catch (error) {
              console.error('자동 로그인 복원 실패:', error);
            }
          }
        }

        if (!restored && !cancelled) {
          clearSessionAuth();
          setActiveUser('');
          setLoginUserId(BETA_USER_ID);
          setIsUserUnlocked(false);
          setActiveUserReports([]);
          setLoginPassword('');
          setProfileAvatarUrl('');
          setProfileNickname('');
          setAuthModal('');
          setRecoveredUserIds([]);
          setRecoveredHintQuestion('');
        }
      } catch (error) {
        console.error('로컬 저장값 로딩 실패:', error);
      } finally {
        if (!cancelled) setIsStorageHydrated(true);
      }
    };

    bootstrap();
    return () => {
      cancelled = true;
    };
  }, []);

  useEffect(() => {
    try {
      localStorage.setItem(STORAGE_KEYS.nozzles, JSON.stringify(nozzleSet));
    } catch (error) {
      console.error('노즐 설정 저장 실패:', error);
    }
  }, [nozzleSet]);

  useEffect(() => {
    try {
      localStorage.setItem(STORAGE_KEYS.samplers, JSON.stringify(samplers));
    } catch (error) {
      console.error('샘플러 설정 저장 실패:', error);
    }
  }, [samplers]);

  useEffect(() => {
    if (!isStorageHydrated) return;
    try {
      localStorage.setItem(STORAGE_KEYS.secureUsers, JSON.stringify(secureUsers));
    } catch (error) {
      console.error('보안 사용자 저장 실패:', error);
    }
  }, [secureUsers, isStorageHydrated]);

  useEffect(() => {
    if (!isStorageHydrated) return;
    try {
      localStorage.setItem(STORAGE_KEYS.activeUser, activeUser);
    } catch (error) {
      console.error('활성 사용자 저장 실패:', error);
    }
  }, [activeUser, isStorageHydrated]);

  const savedData = [...activeUserReports].sort((a, b) => {
    // 측정일자(date)를 우선 정렬 기준으로 사용하고, 없을 때만 저장시각(savedAt)로 보조 정렬
    const aMeasured = a.date ? new Date(a.date).getTime() : NaN;
    const bMeasured = b.date ? new Date(b.date).getTime() : NaN;
    if (Number.isFinite(aMeasured) && Number.isFinite(bMeasured)) return bMeasured - aMeasured;

    const aTime = new Date(a.savedAt || 0).getTime();
    const bTime = new Date(b.savedAt || 0).getTime();
    return bTime - aTime;
  });
  const resolveReportSheetId = (report) => {
    const byId = String(report?.sheetId || '').trim();
    if (byId && SHEET_MENU.some(item => item.id === byId)) return byId;

    const byTitle = String(report?.sheetTitle || '').trim();
    const matched = SHEET_MENU.find(item => item.title === byTitle);
    if (matched) return matched.id;

    return 'dust';
  };
  const profileCombinedReports = savedData.map((report) => {
    const sheetId = resolveReportSheetId(report);
    const sheetMeta = SHEET_MENU.find(item => item.id === sheetId);
    return {
      ...report,
      _sheetId: sheetId,
      _sheetTitle: sheetMeta ? sheetMeta.title : (report.sheetTitle || '먼지시료채취기록부'),
    };
  });
  const profileSheetCounts = SHEET_MENU.reduce((acc, item) => {
    acc[item.id] = profileCombinedReports.filter(report => report._sheetId === item.id).length;
    return acc;
  }, {});
  const savedFitCount = savedData.filter(item => item.isokineticStatus === '적합').length;
  const savedFailCount = savedData.filter(item => item.isokineticStatus === '부적합').length;
  const savedAvgActualC = (() => {
    const nums = savedData
      .map(item => parseFloat(item.actualConcentration))
      .filter(v => Number.isFinite(v));
    if (nums.length === 0) return '-';
    return (nums.reduce((acc, cur) => acc + cur, 0) / nums.length).toFixed(2);
  })();

  const getConfiguredNozzles = () => {
    return nozzleSet
      .map((n) => ({ num: parseInt(n.num, 10), d: parseFloat(n.d) }))
      .filter((n) => !isNaN(n.num) && !isNaN(n.d) && n.d > 0)
      .sort((a, b) => a.num - b.num);
  };

  const handleNozzleConfigChange = (index, field, value) => {
    setNozzleSet(prev => prev.map((item, i) => (
      i === index ? { ...item, [field]: value } : item
    )));
  };

  const addNozzleConfig = () => {
    setNozzleSet(prev => [...prev, { num: '', d: '' }]);
  };

  const removeNozzleConfig = (index) => {
    setNozzleSet(prev => prev.filter((_, i) => i !== index));
  };

  const resetNozzleDefaults = () => {
    setNozzleSet(cloneDefaultNozzles());
  };

  const handleSamplerPresetChange = (index, field, value) => {
    setSamplers(prev => {
      const next = prev.map((item, i) => (i === index ? { ...item, [field]: value } : item));
      const selected = next.find(s => String(s.id) === formData.samplerId);
      if (selected) {
        setFormData(current => ({
          ...current,
          gasMeterFactor: selected.yd,
          deltaHAt: selected.dHAt
        }));
      }
      return next;
    });
  };

  const addSamplerPreset = () => {
    setSamplers(prev => {
      const maxId = prev.reduce((max, item) => Math.max(max, Number(item.id) || 0), 0);
      return [...prev, { id: maxId + 1, name: `샘플러 ${maxId + 1}호기`, yd: '1.000', dHAt: '40.0' }];
    });
  };

  const removeSamplerPreset = (index) => {
    setSamplers(prev => {
      const removed = prev[index];
      const next = prev.filter((_, i) => i !== index);
      if (removed && String(removed.id) === formData.samplerId) {
        setFormData(current => ({ ...current, samplerId: '' }));
      }
      return next;
    });
  };

  const resetSamplerDefaults = () => {
    setSamplers(cloneDefaultSamplers());
    setFormData(prev => ({ ...prev, samplerId: '' }));
  };

  const persistReportsEncrypted = async (userName, reports, keyOverride = null) => {
    const targetUser = secureUsers[userName];
    if (!targetUser) throw new Error('사용자 정보가 없습니다.');
    const key = keyOverride || vaultKeyRef.current[userName];
    if (!key) throw new Error('잠금 해제가 필요합니다.');

    const encryptedData = await encryptJsonWithKey(key, reports);
    setSecureUsers(prev => {
      const current = prev[userName];
      if (!current) return prev;
      return {
        ...prev,
        [userName]: {
          ...current,
          dataIv: encryptedData.iv,
          dataCipher: encryptedData.cipher,
          updatedAt: new Date().toISOString(),
        },
      };
    });
  };

  const createVaultUserRecord = async (userId, password, reports = [], hintQuestion = '', hintAnswer = '') => {
    const salt = bytesToBase64(crypto.getRandomValues(new Uint8Array(16)));
    const key = await deriveVaultKey(password, salt);
    const verifierText = `vault:${userId}:verified`;
    const verifier = await encryptTextWithKey(key, verifierText);
    const encryptedData = await encryptJsonWithKey(key, reports);
    const hintAnswerHash = hintAnswer ? await hashHintAnswer(hintAnswer) : '';

    return {
      key,
      record: {
        salt,
        verifierIv: verifier.iv,
        verifierCipher: verifier.cipher,
        dataIv: encryptedData.iv,
        dataCipher: encryptedData.cipher,
        hintQuestion: hintQuestion || '',
        hintAnswerHash,
        avatarUrl: getFixedAvatarUrlForUser(userId),
        nickname: '',
        updatedAt: new Date().toISOString(),
      },
    };
  };

  const openUserSession = (userId, key, reports, userRecord = null) => {
    const fixedAvatarUrl = getFixedAvatarUrlForUser(userId);
    vaultKeyRef.current[userId] = key;
    setActiveUser(userId);
    setLoginUserId(userId);
    localStorage.setItem(STORAGE_KEYS.lastLoginId, userId);
    setActiveUserReports(Array.isArray(reports) ? reports : []);
    setIsUserUnlocked(true);
    setLoginPassword('');
    setProfileAvatarUrl(fixedAvatarUrl);
    setProfileNickname((userRecord && userRecord.nickname) || '');
  };

  const updateActiveUserProfile = (profilePatch) => {
    if (!activeUser || !isUserUnlocked) return;
    setSecureUsers(prev => {
      const current = prev[activeUser];
      if (!current) return prev;
      return {
        ...prev,
        [activeUser]: {
          ...current,
          ...profilePatch,
          updatedAt: new Date().toISOString(),
        },
      };
    });
  };

  const handleSaveProfileNickname = () => {
    if (!activeUser || !isUserUnlocked) {
      alert('로그인 후 별명을 저장할 수 있습니다.');
      return false;
    }

    const nextNickname = profileNickname.trim().slice(0, 24);
    updateActiveUserProfile({ nickname: nextNickname });
    setProfileNickname(nextNickname);
    alert('별명이 저장되었습니다.');
    return true;
  };

  const handleFindUserIdsByHint = async () => {
    const answer = recoverHintAnswer.trim();
    if (!answer) {
      alert('힌트 정답을 입력해주세요.');
      return;
    }
    const answerHash = await hashHintAnswer(answer);
    const matched = Object.entries(secureUsers)
      .filter(([, user]) => user && user.hintAnswerHash && user.hintAnswerHash === answerHash)
      .map(([userId]) => userId);

    setRecoveredUserIds(matched);
    if (matched.length === 0) {
      alert('일치하는 아이디를 찾지 못했습니다.');
    }
  };

  const handleShowPasswordHint = () => {
    const userId = recoverHintUserId.trim();
    if (!userId) {
      alert('아이디를 입력해주세요.');
      return;
    }
    const targetUser = secureUsers[userId];
    if (!targetUser) {
      alert('해당 아이디를 찾지 못했습니다.');
      setRecoveredHintQuestion('');
      return;
    }
    if (!targetUser.hintQuestion) {
      alert('이 계정은 힌트가 설정되어 있지 않습니다.');
      setRecoveredHintQuestion('');
      return;
    }
    setRecoveredHintQuestion(targetUser.hintQuestion);
  };

  const handleBetaQuickLogin = async () => {
    const userId = BETA_USER_ID;
    const password = BETA_USER_PASSWORD;

    try {
      let targetUser = secureUsers[userId];

      if (!targetUser) {
        const legacyUser = secureUsers[LEGACY_BETA_USER_ID];
        if (legacyUser) {
          try {
            const legacyKey = await deriveVaultKey(LEGACY_BETA_USER_PASSWORD, legacyUser.salt);
            const legacyVerifier = await decryptTextWithKey(legacyKey, legacyUser.verifierIv, legacyUser.verifierCipher);
            if (legacyVerifier === `vault:${LEGACY_BETA_USER_ID}:verified`) {
              const legacyDecrypted = await decryptJsonWithKey(legacyKey, legacyUser.dataIv, legacyUser.dataCipher);
              const legacyReports = Array.isArray(legacyDecrypted) ? legacyDecrypted.filter(item => item && typeof item === 'object') : [];
              const migrated = await createVaultUserRecord(userId, password, legacyReports);
              setSecureUsers(prev => ({ ...prev, [userId]: migrated.record }));
              openUserSession(userId, migrated.key, legacyReports, migrated.record);
              persistSessionAuth(userId, password);
              setAuthModal('');
              return;
            }
          } catch (legacyError) {
            console.error('legacy beta migration failed:', legacyError);
          }
        }

        const created = await createVaultUserRecord(userId, password, []);
        targetUser = created.record;
        setSecureUsers(prev => ({ ...prev, [userId]: created.record }));
        openUserSession(userId, created.key, [], created.record);
        persistSessionAuth(userId, password);
        setAuthModal('');
        return;
      }

      const key = await deriveVaultKey(password, targetUser.salt);
      const verifierText = await decryptTextWithKey(key, targetUser.verifierIv, targetUser.verifierCipher);
      if (verifierText !== `vault:${userId}:verified`) throw new Error('Invalid beta credentials');

      const decrypted = await decryptJsonWithKey(key, targetUser.dataIv, targetUser.dataCipher);
      const reports = Array.isArray(decrypted) ? decrypted.filter(item => item && typeof item === 'object') : [];
      const fixedAvatarUrl = getFixedAvatarUrlForUser(userId);

      if (targetUser.avatarUrl !== fixedAvatarUrl) {
        setSecureUsers(prev => {
          const current = prev[userId];
          if (!current) return prev;
          return {
            ...prev,
            [userId]: {
              ...current,
              avatarUrl: fixedAvatarUrl,
              updatedAt: new Date().toISOString(),
            },
          };
        });
        targetUser = { ...targetUser, avatarUrl: fixedAvatarUrl };
      }

      openUserSession(userId, key, reports, targetUser);
      persistSessionAuth(userId, password);
      setAuthModal('');
    } catch (error) {
      console.error(error);
      alert('베타 user 로그인 중 오류가 발생했습니다.');
    }
  };

  const openAdminLoginModal = () => {
    setLoginUserId(ADMIN_USER_ID);
    setLoginPassword('');
    setAuthModal('admin');
  };

  const handleCreateUser = async () => {
    const trimmed = newUserName.trim();
    const normalizedId = trimmed.toLowerCase();
    const hintQuestion = newUserHintQuestion.trim();
    const hintAnswer = newUserHintAnswer.trim();
    if (!trimmed) {
      alert('사용자 이름을 입력해주세요.');
      return;
    }
    if (secureUsers[trimmed]) {
      alert('이미 존재하는 사용자 이름입니다.');
      return;
    }
    if (normalizedId === ADMIN_USER_ID && newUserPassword !== ADMIN_PASSWORD) {
      alert('관리자 계정 비밀번호는 지정된 값으로만 생성할 수 있습니다.');
      return;
    }
    if (newUserPassword !== newUserPasswordConfirm) {
      alert('비밀번호 확인이 일치하지 않습니다.');
      return;
    }
    if (!PASSWORD_RULE.test(newUserPassword)) {
      alert('비밀번호는 8자 이상, 영문/숫자/특수문자를 각각 1개 이상 포함해야 합니다.');
      return;
    }
    if (!hintQuestion || !hintAnswer) {
      alert('아이디/비밀번호 찾기용 힌트 질문과 정답을 입력해주세요.');
      return;
    }

    try {
      const { key, record } = await createVaultUserRecord(trimmed, newUserPassword, [], hintQuestion, hintAnswer);

      setSecureUsers(prev => ({
        ...prev,
        [trimmed]: record,
      }));
      openUserSession(trimmed, key, [], record);
      persistSessionAuth(trimmed, newUserPassword);
      setNewUserName('');
      setNewUserPassword('');
      setNewUserPasswordConfirm('');
      setNewUserHintQuestion('');
      setNewUserHintAnswer('');
      setAuthModal('');
      alert(`[${trimmed}] 사용자 생성 및 로그인 완료되었습니다.`);
    } catch (error) {
      console.error(error);
      alert('사용자 생성 중 오류가 발생했습니다.');
    }
  };

  const handleLoginUser = async () => {
    const userId = loginUserId.trim();
    if (!userId) {
      alert('아이디를 입력해주세요.');
      return;
    }
    if (!loginPassword) {
      alert('비밀번호를 입력해주세요.');
      return;
    }

    try {
      let targetUser = secureUsers[userId];
      if (!targetUser && userId.toLowerCase() === ADMIN_USER_ID && loginPassword === ADMIN_PASSWORD) {
        const created = await createVaultUserRecord(userId, ADMIN_PASSWORD, []);
        targetUser = created.record;
        setSecureUsers(prev => ({ ...prev, [userId]: created.record }));
        openUserSession(userId, created.key, [], created.record);
        persistSessionAuth(userId, loginPassword);
        setAuthModal('');
        return;
      }

      if (!targetUser) throw new Error('Invalid credentials');
      const key = await deriveVaultKey(loginPassword, targetUser.salt);
      const verifierText = await decryptTextWithKey(key, targetUser.verifierIv, targetUser.verifierCipher);
      if (verifierText !== `vault:${userId}:verified`) throw new Error('Invalid credentials');
      const decrypted = await decryptJsonWithKey(key, targetUser.dataIv, targetUser.dataCipher);
      const reports = Array.isArray(decrypted) ? decrypted.filter(item => item && typeof item === 'object') : [];

      const fixedAvatarUrl = getFixedAvatarUrlForUser(userId);
      if (targetUser.avatarUrl !== fixedAvatarUrl) {
        setSecureUsers(prev => {
          const current = prev[userId];
          if (!current) return prev;
          return {
            ...prev,
            [userId]: {
              ...current,
              avatarUrl: fixedAvatarUrl,
              updatedAt: new Date().toISOString(),
            },
          };
        });
        targetUser = { ...targetUser, avatarUrl: fixedAvatarUrl };
      }

      openUserSession(userId, key, reports, targetUser);
      persistSessionAuth(userId, loginPassword);
      setAuthModal('');
    } catch (error) {
      console.error(error);

      if (userId.toLowerCase() === ADMIN_USER_ID && loginPassword === ADMIN_PASSWORD) {
        try {
          const recreated = await createVaultUserRecord(userId, ADMIN_PASSWORD, []);
          setSecureUsers(prev => ({ ...prev, [userId]: recreated.record }));
          openUserSession(userId, recreated.key, [], recreated.record);
          persistSessionAuth(userId, loginPassword);
          setAuthModal('');
          alert('관리자 계정을 표준 비밀번호로 초기화하고 로그인했습니다.');
          return;
        } catch (adminError) {
          console.error(adminError);
        }
      }

      clearSessionAuth();
      setIsUserUnlocked(false);
      setActiveUserReports([]);
      setProfileAvatarUrl('');
      setProfileNickname('');
      alert('아이디 또는 비밀번호가 올바르지 않습니다.');
    }
  };

  const handleDeleteActiveUser = async () => {
    if (!activeUser || !isUserUnlocked) {
      alert('로그인 후 계정 삭제가 가능합니다.');
      return;
    }

    const willDelete = window.confirm(`[${activeUser}] 계정을 삭제할까요?\n저장된 모든 리포트가 함께 삭제됩니다.`);
    if (!willDelete) return;

    const verifyPassword = window.prompt('삭제 확인을 위해 비밀번호를 다시 입력하세요.');
    if (verifyPassword === null) return;
    if (!verifyPassword) {
      alert('비밀번호를 입력해주세요.');
      return;
    }

    try {
      const targetUser = secureUsers[activeUser];
      if (!targetUser) throw new Error('Missing user');
      const key = await deriveVaultKey(verifyPassword, targetUser.salt);
      const verifierText = await decryptTextWithKey(key, targetUser.verifierIv, targetUser.verifierCipher);
      if (verifierText !== `vault:${activeUser}:verified`) throw new Error('Invalid password');

      setSecureUsers(prev => {
        const next = { ...prev };
        delete next[activeUser];
        return next;
      });
      delete vaultKeyRef.current[activeUser];
      clearSessionAuth();

      setActiveUser('');
      setIsUserUnlocked(false);
      setActiveUserReports([]);
      setLoginUserId(BETA_USER_ID);
      setLoginPassword('');
      setProfileAvatarUrl('');
      setProfileNickname('');
      setSelectedSheet('');
      if (window.location.hash) {
        window.history.pushState(null, '', window.location.pathname + window.location.search);
      }

      alert('계정이 삭제되었습니다.');
    } catch (error) {
      console.error(error);
      alert('비밀번호가 올바르지 않아 계정을 삭제할 수 없습니다.');
    }
  };

  const handleLockActiveUser = () => {
    const currentId = activeUser;
    if (activeUser) {
      delete vaultKeyRef.current[activeUser];
    }
    clearSessionAuth();
    setActiveUser('');
    setIsUserUnlocked(false);
    setActiveUserReports([]);
    setLoginUserId(BETA_USER_ID);
    setLoginPassword('');
    setProfileAvatarUrl('');
    setProfileNickname('');
    setAuthModal('');
    setSelectedSheet('');
    if (window.location.hash) {
      window.history.pushState(null, '', window.location.pathname + window.location.search);
    }
  };

  const clearActiveUserReports = async () => {
    if (!activeUser) return;
    if (!isUserUnlocked) {
      alert('먼저 로그인해주세요.');
      return;
    }
    if (!window.confirm(`'${activeUser}' 사용자의 저장 리포트를 모두 삭제할까요?`)) return;

    try {
      await persistReportsEncrypted(activeUser, []);
      setActiveUserReports([]);
    } catch (error) {
      console.error(error);
      alert('리포트 삭제 중 오류가 발생했습니다.');
    }
  };

  const removeSavedReport = async (id, savedAt) => {
    if (!activeUser || !isUserUnlocked) {
      alert('로그인 후 삭제할 수 있습니다.');
      return;
    }

    const nextReports = (activeUserReports || []).filter((item) => {
      if (id) return item.id !== id;
      return item.savedAt !== savedAt;
    });

    try {
      await persistReportsEncrypted(activeUser, nextReports);
      setActiveUserReports(nextReports);
    } catch (error) {
      console.error(error);
      alert('리포트 삭제 중 오류가 발생했습니다.');
    }
  };

  const handleLoadSavedReport = (report) => {
    if (!report || typeof report !== 'object') return;
    if (!activeUser || !isUserUnlocked) {
      alert('로그인 후 불러올 수 있습니다.');
      return;
    }

    setFormData(prev => {
      const nextForm = { ...prev };
      Object.keys(nextForm).forEach((key) => {
        if (Object.prototype.hasOwnProperty.call(report, key)) {
          nextForm[key] = report[key];
        }
      });

      return nextForm;
    });

    setRecommendations(null);

    const targetSheet = resolveReportSheetId(report);
    if (selectedSheet !== targetSheet) {
      window.location.hash = `sheet=${targetSheet}`;
    }

    const targetTitle = SHEET_MENU.find(item => item.id === targetSheet)?.title || report.sheetTitle || '기록부';
    alert(`${targetTitle} 리포트를 불러왔습니다.`);
  };

  const handleChange = (e) => {
    const { name, value } = e.target;
    setFormData(prev => {
      const newData = { ...prev, [name]: value };
      if (name === 'gasMeterFactor' || name === 'deltaHAt') newData.samplerId = '';
      if (name === 'totalStackDepth' || name === 'flangeLength') {
        const total = parseFloat(newData.totalStackDepth);
        const flange = parseFloat(newData.flangeLength) || 0; 
        if (!isNaN(total)) newData.stackDiameter = Math.max(0, total - flange).toFixed(3);
      }
      return newData;
    });
  };

  const handleSamplerChange = (e) => {
    const id = e.target.value;
    if (id === '') { setFormData(prev => ({ ...prev, samplerId: '' })); return; }
    const sampler = samplers.find(s => s.id === parseInt(id, 10));
    if (sampler) setFormData(prev => ({ ...prev, samplerId: id, gasMeterFactor: sampler.yd, deltaHAt: sampler.dHAt }));
  };

  const handleGasAnalyzerChange = (index, field, value) => {
    setFormData(prev => {
      const newAnalyzer = prev.gasAnalyzer.map((item, i) => 
        i === index ? { ...item, [field]: value } : item
      );
      return { ...prev, gasAnalyzer: newAnalyzer };
    });
  };

  const handleMoistureChange = (index, value) => {
    setFormData(prev => {
      const newMoistures = [...prev.moistureValues];
      newMoistures[index] = value;
      return { ...prev, moistureValues: newMoistures };
    });
  };

  const handleImpingerChange = (index, field, value) => {
    setFormData(prev => {
      const newImpingers = prev.impingers.map((item, i) => 
        i === index ? { ...item, [field]: value } : item
      );
      return { ...prev, impingers: newImpingers };
    });
  };

  const handlePointChange = (index, field, value) => {
    setFormData(prev => {
      const newPoints = prev.points.map((item, i) => {
        if (i !== index) return item;
        const newItem = { ...item, [field]: value };
        if (field === 'tp' || field === 'sp') {
          const tp = parseFloat(newItem.tp);
          const sp = parseFloat(newItem.sp);
          if (!isNaN(tp) && !isNaN(sp)) newItem.dp = (tp - sp).toFixed(2);
        }
        return newItem;
      });
      return { ...prev, points: newPoints };
    });
  };

  const addPoint = () => setFormData(prev => ({ ...prev, points: [...prev.points, { id: prev.points.length + 1, tp: '', sp: '', dp: '', ts: '' }] }));
  const removePoint = (index) => {
    setFormData(prev => {
      const newPoints = [...prev.points];
      newPoints.splice(index, 1);
      return { ...prev, points: newPoints };
    });
  };

  const handleGasMeterChange = (index, field, value) => {
    setFormData(prev => {
      const newMeters = prev.gasMeters.map((item, i) => {
        if (i !== index) return item;
        const newItem = { ...item, [field]: value };
        if (field === 'dp') {
           const dpValue = parseFloat(value);
           const kVal = parseFloat(prev.kFactor);
           if (!isNaN(dpValue) && !isNaN(kVal)) newItem.pressure = roundDH(kVal * dpValue);
           else if (value === '') newItem.pressure = '';
        }
        return newItem;
      });
      return { ...prev, gasMeters: newMeters };
    });
  };
  
  const handleKFactorChange = (e) => {
      const kValStr = e.target.value;
      setFormData(prev => {
          const kNum = parseFloat(kValStr);
          const newMeters = prev.gasMeters.map((meter, idx) => {
              if (idx === 0) return meter;
              const dpValue = parseFloat(meter.dp);
              if (Number.isFinite(kNum) && !isNaN(dpValue)) return { ...meter, pressure: roundDH(kNum * dpValue) };
              return meter;
          });
          return { ...prev, kFactor: kValStr, gasMeters: newMeters };
      });
  };

  const addGasMeter = () => setFormData(prev => ({ ...prev, gasMeters: [...prev.gasMeters, { id: prev.gasMeters.length, pointNum: '', time: '', stackTemp: '', dp: '', pressure: '', volume: '', tmIn: '', tmOut: '', vacuum: '', impingerTemp: '' }] }));
  const removeGasMeter = (index) => {
    if (index === 0) return;
    setFormData(prev => {
      const newMeters = [...prev.gasMeters];
      newMeters.splice(index, 1);
      newMeters.forEach((m, i) => m.id = i);
      return { ...prev, gasMeters: newMeters };
    });
  };

  const getSamplingPoints = () => {
    const D = parseFloat(formData.stackDiameter), flange = parseFloat(formData.flangeLength) || 0;
    if (isNaN(D) || D <= 0) return null;
    const R = D / 2, area = Math.PI * Math.pow(R, 2);
    let rnCoeffs = [], isCenterOnly = false;

    if (area <= 0.25) isCenterOnly = true;
    else if (D <= 1) rnCoeffs = [0.707]; 
    else if (D <= 2) rnCoeffs = [0.500, 0.866]; 
    else if (D <= 4) rnCoeffs = [0.408, 0.707, 0.913]; 
    else if (D <= 4.5) rnCoeffs = [0.354, 0.612, 0.791, 0.935];
    else rnCoeffs = [0.316, 0.548, 0.707, 0.837, 0.949]; 
    
    let nearDistances = [], farDistances = [];
    if (isCenterOnly) {
      nearDistances.push(R); farDistances.push(R);
    } else {
      for (let i = rnCoeffs.length - 1; i >= 0; i--) {
        nearDistances.push(R - (rnCoeffs[i] * R));
        farDistances.push(R + (rnCoeffs[i] * R));
      }
    }
    return { perRadius: isCenterOnly ? 1 : rnCoeffs.length, isCenterOnly, area: area.toFixed(3), nearInsertion: nearDistances.map(d => (d + flange).toFixed(3)), farInsertion: farDistances.map(d => (d + flange).toFixed(3)) };
  };

  const applySamplingPointsToTable = () => {
    const spData = getSamplingPoints();
    if (!spData) return;
    const pointCount = spData.perRadius;
    const minMeasureRows = 5; // 시작(0분) 행 제외 최소 측정기록 칸
    const totalMeterRows = Math.max(pointCount + 1, minMeasureRows + 1);
    setFormData(prev => {
      const newPoints = Array.from({ length: pointCount }, (_, i) => prev.points[i] || { id: i + 1, tp: '', sp: '', dp: '', ts: '' });
      const newMeters = Array.from({ length: totalMeterRows }, (_, i) => {
        if (i === 0) return prev.gasMeters[0] || { id: 0, pointNum: '시작', time: '0', stackTemp: '', dp: '', pressure: '', volume: '', tmIn: '', tmOut: '', vacuum: '', impingerTemp: '' };
        return prev.gasMeters[i] || { id: i, pointNum: '', time: '', stackTemp: '', dp: '', pressure: '', volume: '', tmIn: '', tmOut: '', vacuum: '', impingerTemp: '' };
      });
      return { ...prev, points: newPoints, gasMeters: newMeters };
    });
    alert(`측정점 ${pointCount}개를 반영했고, 측정기록표는 시작행 포함 총 ${totalMeterRows}행(측정칸 최소 ${minMeasureRows}개)으로 생성했습니다.`);
  };

  const getGasComposition = () => {
    const o2 = formData.gasAnalyzer.reduce((acc, curr) => acc + (parseFloat(curr.o2) || 0), 0) / 3;
    const co2 = formData.gasAnalyzer.reduce((acc, curr) => acc + (parseFloat(curr.co2) || 0), 0) / 3;
    const co = formData.gasAnalyzer.reduce((acc, curr) => acc + (parseFloat(curr.co) || 0), 0) / 3;
    const sox = formData.gasAnalyzer.reduce((acc, curr) => acc + (parseFloat(curr.sox) || 0), 0) / 3;
    const nox = formData.gasAnalyzer.reduce((acc, curr) => acc + (parseFloat(curr.nox) || 0), 0) / 3;
    
    const coPercent = co / 10000; 
    const n2 = 100 - co2 - o2 - coPercent; 
    
    const M_O2 = 32, M_CO2 = 44, M_CO = 28, M_N2 = 28;
    const sumMx = (M_O2 * o2) + (M_CO2 * co2) + (M_CO * coPercent) + (M_N2 * n2);

    const Md = sumMx / 100;
    const validVals = formData.moistureValues.map(v => parseFloat(v)).filter(v => !isNaN(v));
    const Xw = validVals.length > 0 ? (validVals.reduce((a, b) => a + b, 0) / validVals.length) : 0;
    const Ms = Md * ((100 - Xw) / 100) + 18 * (Xw / 100);

    const r0 = (1 / (22.4 * 100)) * ( sumMx * ((100 - Xw) / 100) + 18 * Xw );

    return { 
      o2, co2, co, sox, nox, n2, 
      Md: Number(Md.toFixed(3)), 
      Ms: Number(Ms.toFixed(3)), 
      Xw, 
      r0: Number(r0.toFixed(3)) 
    };
  };

  const getRawAvgTp = () => {
    const validTps = formData.points.map(p => parseFloat(p.tp)).filter(v => !isNaN(v));
    return validTps.length === 0 ? 0 : validTps.reduce((a, b) => a + b, 0) / validTps.length;
  };

  const getRawAvgSp = () => {
    const validSps = formData.points.map(p => parseFloat(p.sp)).filter(v => !isNaN(v));
    return validSps.length === 0 ? 0 : validSps.reduce((a, b) => a + b, 0) / validSps.length;
  };

  const getRawAvgDp = () => {
    const validDps = formData.points.map(p => parseFloat(p.dp)).filter(v => !isNaN(v));
    return validDps.length === 0 ? 0 : validDps.reduce((a, b) => a + b, 0) / validDps.length;
  };

  const getRawAvgTs = () => {
    const validTs = formData.points.map(p => parseFloat(p.ts)).filter(v => !isNaN(v));
    return validTs.length === 0 ? 0 : validTs.reduce((a, b) => a + b, 0) / validTs.length;
  };

  const getRawStackPressure = () => {
    const Pa = parseFloat(formData.atmPressure);
    const validSps = formData.points.map(p => parseFloat(p.sp)).filter(v => !isNaN(v));
    const avgPs = validSps.length === 0 ? 0 : validSps.reduce((a, b) => a + b, 0) / validSps.length;
    return !isNaN(Pa) ? Pa + (avgPs / 13.6) : NaN;
  };

  const getRawGasVelocity = () => {
    const avgTs = getRawAvgTs(), C = parseFloat(formData.pitotFactor) || 0.84, P = getRawStackPressure(); 
    const dpAvg = getRawAvgDp();
    if (dpAvg === 0 || isNaN(avgTs) || isNaN(P)) return 0;
    
    const gasComp = getGasComposition();
    if (isNaN(gasComp.r0) || gasComp.r0 === 0) return 0;
    
    const r = Number((gasComp.r0 * (273 / (273 + avgTs)) * (P / 760)).toFixed(3));
    
    // 엑셀 수식: =D9*(2*9.81*J15/K15)^0.5
    // 각 점을 따로 계산하지 않고, '평균 동압(dpAvg)'을 식에 한 번에 넣어서 엑셀과 완전히 똑같이 계산합니다.
    return C * Math.pow((2 * 9.81 * dpAvg) / r, 0.5);
  };

  const getRawGasMeterVolDiff = () => {
    const validVols = formData.gasMeters.map(g => g.volume).filter(v => v !== '');
    if (validVols.length > 1) {
      const initial = parseFloat(validVols[0]), final = parseFloat(validVols[validVols.length - 1]);
      if (!isNaN(initial) && !isNaN(final)) return final - initial;
    }
    return 0;
  };

  const getRawAvgTm = () => {
    const validTemps = formData.gasMeters.filter((_, i) => i !== 0).map(g => {
      const inT = parseFloat(g.tmIn), outT = parseFloat(g.tmOut);
      if (!isNaN(inT) && !isNaN(outT)) return (inT + outT) / 2;
      if (!isNaN(inT)) return inT;
      if (!isNaN(outT)) return outT;
      return NaN;
    }).filter(v => !isNaN(v));
    return validTemps.length === 0 ? NaN : validTemps.reduce((a, b) => a + b, 0) / validTemps.length;
  };

  const getRawAvgOrifice = () => {
    const validPressures = formData.gasMeters.filter((_, i) => i !== 0).map(g => parseFloat(g.pressure)).filter(v => !isNaN(v));
    return validPressures.length === 0 ? NaN : validPressures.reduce((a, b) => a + b, 0) / validPressures.length;
  };

  const getRawTotalMoistureWeight = () => {
    let total = 0;
    formData.impingers.forEach(imp => {
      const init = parseFloat(imp.initial), fin = parseFloat(imp.final);
      if (!isNaN(init) && !isNaN(fin)) total += (fin - init);
    });
    return total;
  };

  const getRawPostMoisture = () => {
    const Wm = getRawTotalMoistureWeight(), Vm = getRawGasMeterVolDiff(), Tm = getRawAvgTm();
    const Pa = parseFloat(formData.atmPressure), Pm = getRawAvgOrifice() || 0;
    const Y = parseFloat(formData.gasMeterFactor) || 1.0;
    if (Wm > 0 && Vm > 0 && !isNaN(Tm) && !isNaN(Pa)) {
      const Vm_std = Vm * Y * (273 / (273 + Tm)) * ((Pa + Pm / 13.6) / 760);
      const Vw_std = Wm * (22.4 / 18.015);
      return (Vw_std / (Vm_std + Vw_std)) * 100;
    }
    return 0;
  };

  const getRawVmStd = () => {
    const Vm = getRawGasMeterVolDiff(), Tm = getRawAvgTm();
    const Pa = parseFloat(formData.atmPressure), Pm = getRawAvgOrifice() || 0;
    const Y = parseFloat(formData.gasMeterFactor) || 1.0;
    if (Vm > 0 && !isNaN(Tm) && !isNaN(Pa)) {
      return Vm * Y * (273 / (273 + Tm)) * ((Pa + Pm / 13.6) / 760);
    }
    return 0;
  };

  const getRawDustWeightDiff = () => {
    const initial = parseFloat(formData.filterInitial), final = parseFloat(formData.filterFinal);
    if (!isNaN(initial) && !isNaN(final)) return final - initial;
    return 0;
  };

  // -------------------------------------------------------------
  // 화면 표기(UI 표시) 전용 함수 모음
  // -------------------------------------------------------------
  const calcAvgTp = () => getRawAvgTp();
  const calcAvgSp = () => getRawAvgSp();
  const calcAvgDp = () => getRawAvgDp(); 
  const calcAvgTs = () => getRawAvgTs();
  const calcStackPressure = () => getRawStackPressure();
  const calcMoisture = () => getGasComposition().Xw.toFixed(2);
  const calcPostMoisture = () => getRawPostMoisture().toFixed(2);
  const calcGasVelocity = () => { const v = getRawGasVelocity(); return v === 0 ? '0.00' : v.toFixed(2); };
  const calcGasMeterVolDiff = () => { const v = getRawGasMeterVolDiff(); return v > 0 ? v.toFixed(2) : 0; };
  const calcAvgTm = () => { const v = getRawAvgTm(); return isNaN(v) ? '-' : v.toFixed(1); };
  const calcAvgOrifice = () => { const v = getRawAvgOrifice(); return isNaN(v) ? '-' : v.toFixed(1); };
  const calcTotalMoistureWeight = () => getRawTotalMoistureWeight();
  const getVmStd = () => { const v = getRawVmStd(); return v > 0 ? v.toFixed(3) : '0.000'; };
  const calcDustWeightDiff = () => { const v = getRawDustWeightDiff(); return v.toFixed(4); };

  const calcO2CorrectionFactor = () => {
    const o2_ref = parseFloat(formData.standardO2);
    const gasComp = getGasComposition();
    const o2_act = gasComp.o2;
    if (isNaN(o2_ref) || isNaN(o2_act) || o2_act >= 21) return 1.0;
    return (21 - o2_ref) / (21 - o2_act);
  };

  const getSamplingMinutes = () => formData.gasMeters.reduce((sum, meter, idx) => idx === 0 ? sum : sum + (isNaN(parseFloat(meter.time)) ? 0 : parseFloat(meter.time)), 0);

  const calcGasFlowRates = () => {
    const Vs = getRawGasVelocity(); 
    const Ts = getRawAvgTs();
    const Ps = getRawStackPressure();
    const D = parseFloat(formData.stackDiameter);
    
    if (Vs === 0 || isNaN(Ts) || isNaN(Ps) || isNaN(D)) return { dry: '-', wet: '-' };

    const A = Math.PI * Math.pow(D / 2, 2);
    const gasComp = getGasComposition(); 
    const Xw = gasComp.Xw;

    const Q_wet = Vs * A * (273 / (273 + Ts)) * (Ps / 760) * 3600;
    const Q_dry = Q_wet * (1 - Xw / 100);

    return { wet: Q_wet.toFixed(0), dry: Q_dry.toFixed(0) };
  };

  const generateRecommendation = (mode) => {
    const Vs = getRawGasVelocity(), dpAvg = getRawAvgDp();
    const targetVol = parseFloat(formData.targetVolume) || 1000;
    const configuredNozzles = getConfiguredNozzles();
    if (configuredNozzles.length === 0) return null;
    
    const tmIn = parseFloat(formData.traverseTmIn) || parseFloat(formData.atmTemp) || 25;
    const tmOut = parseFloat(formData.traverseTmOut) || parseFloat(formData.atmTemp) || 25;
    const Tm = (tmIn + tmOut) / 2;

    const Ts = getRawAvgTs();
    const Pm = parseFloat(formData.atmPressure); 
    
    const validSps = formData.points.map(p => parseFloat(p.sp)).filter(v => !isNaN(v));
    const avgSp = validSps.length === 0 ? 0 : validSps.reduce((a, b) => a + b, 0) / validSps.length;
    const Ps_ratio = (Pm + avgSp / 13.6) / Pm; 

    const gasComp = getGasComposition();
    const Md = gasComp.Md, Ms = gasComp.Ms, Xw = gasComp.Xw;

    const Cp = parseFloat(formData.pitotFactor) || 0.84;
    const Yd = parseFloat(formData.gasMeterFactor) || 1.0;
    const dHAt = parseFloat(formData.deltaHAt);

    let bestNozzle = null, maxDH = 0, closestDiffStandard = Infinity, closestDiffStable = Infinity;

    for (let n of configuredNozzles) {
        const Dn = n.d;
        const tempK = 8.001e-5 * Math.pow(Dn, 4) * dHAt * Math.pow(Cp, 2) * Math.pow(1 - Xw / 100, 2) * (Md / Ms) * ((Tm + 273) / (Ts + 273)) * Ps_ratio;
        const expectedDHCheck = parseFloat(calcExpectedOrificeDH(tempK, dpAvg));
        if (!Number.isFinite(expectedDHCheck)) continue;

        if (expectedDHCheck <= 50) {
            if (mode === 'fast') {
                if (expectedDHCheck > maxDH) { maxDH = expectedDHCheck; bestNozzle = n; }
            } else if (mode === 'standard') {
                const diff = Math.abs(expectedDHCheck - 25);
                if (diff < closestDiffStandard) { closestDiffStandard = diff; bestNozzle = n; }
            } else if (mode === 'stable') {
                const diff = Math.abs(expectedDHCheck - 10);
                if (diff < closestDiffStable) { closestDiffStable = diff; bestNozzle = n; }
            }
        }
    }

    if (!bestNozzle) bestNozzle = configuredNozzles[0];

    const Dn_best = bestNozzle.d;
    const An = Math.PI * Math.pow(Dn_best / 2000, 2);
    const Ps_abs = Pm + avgSp / 13.6;
    const bestQm = (Vs * An * 60 * 1000 * ((Tm + 273) / (Ts + 273)) * (Ps_abs / Pm) * (1 - Xw / 100)) / Yd;
    const Q_sl = bestQm * Yd * (273 / (Tm + 273)) * (Pm / 760);
    
    let fastestTime = Math.ceil(targetVol / Q_sl);
    if (fastestTime <= 0 || !isFinite(fastestTime)) fastestTime = 65;

    const K = 8.001e-5 * Math.pow(Dn_best, 4) * dHAt * Math.pow(Cp, 2) * Math.pow(1 - Xw / 100, 2) * (Md / Ms) * ((Tm + 273) / (Ts + 273)) * Ps_ratio;

    return { bestNozzleNum: bestNozzle.num, finalDn: Dn_best.toFixed(2), fastestTime, calculatedK: K, expectedDH: calcExpectedOrificeDH(K, dpAvg) };
  };

  const handleCalculateOptions = () => {
    const Vs = getRawGasVelocity(), Ts = getRawAvgTs(), Ps = getRawStackPressure(), dpAvg = getRawAvgDp();
    const Y = parseFloat(formData.gasMeterFactor), dHAt = parseFloat(formData.deltaHAt);
    const configuredNozzles = getConfiguredNozzles();

    if (isNaN(Y) || Y <= 0 || isNaN(dHAt) || dHAt <= 0) { alert('장비 교정 성적서에 명시된 [보정계수(Yd)]와 [오리피스 계수(ΔH@)]를 먼저 정확히 입력해주세요.'); return; }
    if (Vs === 0 || isNaN(Vs) || isNaN(Ts) || isNaN(Ps) || isNaN(dpAvg) || dpAvg === 0) { alert('오류: 4번 항목의 유속(동압 및 온도) 기초 데이터가 먼저 입력되어야 산정이 가능합니다.'); return; }
    if (configuredNozzles.length === 0) { alert('노즐 목록이 비어 있습니다. 노즐 설정에서 최소 1개 이상 입력해주세요.'); return; }

    setRecommendations({ fast: generateRecommendation('fast'), standard: generateRecommendation('standard'), stable: generateRecommendation('stable') });
  };

  const applyRecommendation = (opt) => {
    const appliedK = parseFloat(opt.calculatedK);
    setFormData(prev => {
        const newMeters = prev.gasMeters.map((meter, idx) => {
            if (idx === 0) return meter;
            const dpValue = parseFloat(meter.dp);
            if (Number.isFinite(appliedK) && !isNaN(dpValue)) return { ...meter, pressure: roundDH(appliedK * dpValue) };
            return meter;
        });

        return { 
          ...prev, 
          usedNozzleNum: String(opt.bestNozzleNum), 
          nozzleDiameter: opt.finalDn, 
          kFactor: Number.isFinite(appliedK) ? String(opt.calculatedK) : '', 
          planSamplingTime: String(opt.fastestTime), 
          recommendedNozzleNum: String(opt.bestNozzleNum), 
          recommendedNozzleDia: opt.finalDn, 
          gasMeters: newMeters 
        };
    });
  };

  const calcRowIsokineticRate = (idx) => {
    if (idx === 0) return '-';
    const meter = formData.gasMeters[idx];
    
    let prevVolStr = '';
    for (let i = idx - 1; i >= 0; i--) {
      if (formData.gasMeters[i].volume) { prevVolStr = formData.gasMeters[i].volume; break; }
    }
    
    const currVolStr = meter.volume, pressureStr = meter.pressure, inT = meter.tmIn, outT = meter.tmOut, timeStr = meter.time; 
    if (!prevVolStr || !currVolStr || !timeStr || !pressureStr || (inT === '' && outT === '')) return '-';
    
    const prevVol = parseFloat(prevVolStr), currVol = parseFloat(currVolStr), Vm = currVol - prevVol; 
    if (Vm <= 0 || isNaN(Vm)) return '-';
    
    let theta = parseFloat(timeStr);
    if (isNaN(theta) || theta <= 0) return '-';
    
    let Tm;
    if (inT !== '' && outT !== '') Tm = (parseFloat(inT) + parseFloat(outT)) / 2;
    else if (inT !== '') Tm = parseFloat(inT);
    else Tm = parseFloat(outT);
    if (isNaN(Tm)) return '-';
    
    let Ts = parseFloat(meter.stackTemp);
    if (isNaN(Ts)) Ts = getRawAvgTs();
    if (isNaN(Ts)) return '-';
    
    const Pa = parseFloat(formData.atmPressure);
    const Pm_val = parseFloat(pressureStr);
    const Ps = getRawStackPressure(), Cp = parseFloat(formData.pitotFactor) || 0.84;
    const Y = parseFloat(formData.gasMeterFactor) || 1.0;
    if (isNaN(Ps) || isNaN(Pa) || isNaN(Pm_val)) return '-';
    
    const Pm_abs = Pa + Pm_val / 13.6;
    const gasComp = getGasComposition();
    const r0 = gasComp.r0;
    const Xw = gasComp.Xw;
    
    const dpStr = meter.dp;
    if (!dpStr) return '-';
    const dp = parseFloat(dpStr);
    if (dp < 0 || isNaN(dp)) return '-';

    const r_row = Number((r0 * (273 / (273 + Ts)) * (Ps / 760)).toFixed(3));

    const Vs = Cp * Math.pow((2 * 9.81 * dp) / r_row, 0.5);
    if (Vs <= 0 || isNaN(Vs)) return '-';
    
    const Dn = parseFloat(formData.nozzleDiameter);
    if (isNaN(Dn) || Dn <= 0) return '-';
    const An = Math.PI * Math.pow(Dn / 2000, 2);
    
    const rate = (Vm * Y * (Ts + 273) * Pm_abs * 100) / ((Tm + 273) * Vs * An * 60 * 1000 * theta * Ps * (1 - Xw / 100));
    return rate.toFixed(1);
  };

  const calcIsokineticRate = (isPost = false) => {
    const Vm = getRawGasMeterVolDiff();
    const Ts = getRawAvgTs();
    const rawTm = getRawAvgTm();
    const Tm = isNaN(rawTm) ? Ts : rawTm; 
    const Dn = parseFloat(formData.nozzleDiameter);
    const Vs = getRawGasVelocity();
    const theta = getSamplingMinutes();
    const Y = parseFloat(formData.gasMeterFactor) || 1.0;
    const Pa = parseFloat(formData.atmPressure);
    const dH_avg = getRawAvgOrifice();
    const Ps = getRawStackPressure();

    if (Vm > 0 && Dn > 0 && Vs > 0 && theta > 0 && !isNaN(Ts) && !isNaN(Pa) && !isNaN(dH_avg) && !isNaN(Ps)) {
      const An = Math.PI * Math.pow(Dn / 2000, 2);
      const Pm_abs = Pa + dH_avg / 13.6;

      if (isPost) {
        const Vic = getRawTotalMoistureWeight(); 
        const Vm_m3 = Vm / 1000;
        const part1 = 0.00346 * Vic;
        const part2 = (Vm_m3 * Y / (Tm + 273)) * Pm_abs;
        const I = (Ts + 273) * (part1 + part2) * 1.667 / (theta * Vs * Ps * An);
        return I.toFixed(1);
      } else {
        const gasComp = getGasComposition();
        const Xw = gasComp.Xw;
        const I = (Vm * Y * (Ts + 273) * Pm_abs * 100) / ((Tm + 273) * Vs * An * 60 * 1000 * theta * Ps * (1 - Xw / 100));
        return I.toFixed(1);
      }
    }
    return '-';
  };

  const calcDustConcentrations = () => {
    const Vm_std = getRawVmStd();
    const Wm = getRawDustWeightDiff();
    let actualC = '-', correctedC = '-';

    if (Vm_std > 0 && Wm > 0) {
      const C = Wm / Vm_std; 
      actualC = C.toFixed(2);
      const K = calcO2CorrectionFactor();
      correctedC = (C * K).toFixed(2);
    }
    return { actualC, correctedC };
  };

  const handleSave = async (e) => {
    e.preventDefault();
    if (!activeUser) {
      alert('먼저 아이디/비밀번호로 로그인해주세요.');
      return;
    }
    if (!isUserUnlocked) {
      alert('저장하려면 먼저 로그인해주세요.');
      return;
    }

    const { actualC, correctedC } = calcDustConcentrations();
    const flowRates = calcGasFlowRates();
    const isokineticRate = calcIsokineticRate(true);
    const isokineticNum = parseFloat(isokineticRate);
    const isokineticStatus = !isNaN(isokineticNum) && isokineticNum >= 95 && isokineticNum <= 105 ? '적합' : '부적합';

    const result = {
      id: `${Date.now()}-${Math.random().toString(36).slice(2, 8)}`,
      savedAt: new Date().toISOString(),
      user: activeUser,
      sheetId: selectedSheet || 'dust',
      sheetTitle: activeSheet?.title || '먼지시료채취기록부',
      ...formData,
      sampler: formData.sampler || activeUser,
      moisturePre: calcMoisture(),
      moisturePost: calcPostMoisture(),
      moisturePercent: calcTotalMoistureWeight() > 0 ? calcPostMoisture() : calcMoisture(),
      avgVelocity: calcGasVelocity(),
      isokineticRate,
      isokineticStatus,
      dustWeight: calcDustWeightDiff(),
      actualConcentration: actualC,
      correctedConcentration: correctedC,
      wetFlowRate: flowRates.wet,
      dryFlowRate: flowRates.dry,
      gasMeterVolume: calcGasMeterVolDiff(),
      vmStd: getVmStd(),
      totalSamplingMinutes: String(getSamplingMinutes()),
      avgTm: calcAvgTm(),
      avgOrifice: calcAvgOrifice(),
      o2CorrectionFactor: calcO2CorrectionFactor().toFixed(3),
      nozzleUsed: formData.usedNozzleNum || formData.recommendedNozzleNum || '-',
      nozzleDiameterUsed: formData.nozzleDiameter || '-',
      kFactorApplied: formData.kFactor || '-',
    };

    const nextReports = [...activeUserReports, result];
    try {
      await persistReportsEncrypted(activeUser, nextReports);
      setActiveUserReports(nextReports);
      alert(`[${activeUser}] 데이터로 임시 저장되었습니다.`);
    } catch (error) {
      console.error(error);
      alert('암호화 저장 중 오류가 발생했습니다.');
    }
  };

  // Enter 키를 제출 대신 다음 입력 칸 이동(Tab 동작)으로 사용
  const handleFormKeyDown = (e) => {
    if (e.key !== 'Enter') return;

    const target = e.target;
    if (!target || !(target instanceof HTMLElement)) return;

    const tagName = target.tagName.toLowerCase();
    if (tagName !== 'input' && tagName !== 'select') return;
    if (target instanceof HTMLInputElement) {
      const inputType = (target.type || '').toLowerCase();
      if (inputType === 'submit' || inputType === 'button' || inputType === 'checkbox' || inputType === 'radio') {
        return;
      }
    }

    e.preventDefault();

    const form = e.currentTarget;
    if (!(form instanceof HTMLFormElement)) return;

    const focusable = Array.from(
      form.querySelectorAll('input:not([disabled]):not([type="hidden"]), select:not([disabled]), textarea:not([disabled]), button:not([disabled]), [tabindex]:not([tabindex="-1"])')
    ).filter((el) => el instanceof HTMLElement && (el.offsetParent !== null || el === document.activeElement));

    const currentIndex = focusable.indexOf(target);
    if (currentIndex === -1) return;

    const nextEl = focusable[currentIndex + 1] || focusable[0];
    if (nextEl && nextEl instanceof HTMLElement) {
      nextEl.focus();
      if (nextEl instanceof HTMLInputElement && ['text', 'number', 'search', 'email', 'tel', 'url', 'password', 'time', 'date'].includes(nextEl.type)) {
        nextEl.select?.();
      }
    }
  };

  const toNumber = (val) => {
    const n = typeof val === 'number' ? val : parseFloat(val);
    return Number.isFinite(n) ? n : null;
  };

  const toExcelDateSerial = (val) => {
    if (!val) return null;
    const dt = new Date(val);
    if (Number.isNaN(dt.getTime())) return null;
    return (Date.UTC(dt.getFullYear(), dt.getMonth(), dt.getDate()) - Date.UTC(1899, 11, 30)) / 86400000;
  };

  const toExcelTimeFraction = (val) => {
    if (!val || typeof val !== 'string' || !val.includes(':')) return null;
    const [h, m] = val.split(':').map(Number);
    if (!Number.isFinite(h) || !Number.isFinite(m)) return null;
    return ((h * 60) + m) / (24 * 60);
  };

  const parseCellRef = (ref) => {
    const match = /^([A-Z]+)(\d+)$/.exec(String(ref).toUpperCase());
    if (!match) return null;

    const [, colText, rowText] = match;
    let col = 0;
    for (let i = 0; i < colText.length; i++) {
      col = (col * 26) + (colText.charCodeAt(i) - 64);
    }

    const row = parseInt(rowText, 10);
    if (!Number.isFinite(row) || row <= 0) return null;
    return { ref: `${colText}${row}`, row, col };
  };

  const getChildElementsByLocalName = (node, localName) => {
    return Array.from(node.childNodes).filter(
      child => child.nodeType === 1 && child.localName === localName
    );
  };

  const findOrCreateRow = (sheetData, rowNum, ns) => {
    const rows = getChildElementsByLocalName(sheetData, 'row');
    const rowKey = String(rowNum);
    const existing = rows.find(row => row.getAttribute('r') === rowKey);
    if (existing) return existing;

    const newRow = sheetData.ownerDocument.createElementNS(ns, 'row');
    newRow.setAttribute('r', rowKey);

    const nextRow = rows.find((row) => {
      const current = parseInt(row.getAttribute('r'), 10);
      return Number.isFinite(current) && current > rowNum;
    });

    if (nextRow) sheetData.insertBefore(newRow, nextRow);
    else sheetData.appendChild(newRow);

    return newRow;
  };

  const findOrCreateCell = (rowEl, cellRef, colIdx, ns) => {
    const cells = getChildElementsByLocalName(rowEl, 'c');
    const existing = cells.find(cell => (cell.getAttribute('r') || '').toUpperCase() === cellRef);
    if (existing) return existing;

    const newCell = rowEl.ownerDocument.createElementNS(ns, 'c');
    newCell.setAttribute('r', cellRef);

    const nextCell = cells.find((cell) => {
      const parsed = parseCellRef(cell.getAttribute('r') || '');
      return parsed && parsed.col > colIdx;
    });

    if (nextCell) rowEl.insertBefore(newCell, nextCell);
    else rowEl.appendChild(newCell);

    return newCell;
  };

  const setXmlCellValue = (cell, value, ns) => {
    getChildElementsByLocalName(cell, 'f').forEach(node => cell.removeChild(node));
    getChildElementsByLocalName(cell, 'v').forEach(node => cell.removeChild(node));
    getChildElementsByLocalName(cell, 'is').forEach(node => cell.removeChild(node));

    if (value === null || value === undefined || value === '') {
      cell.removeAttribute('t');
      return;
    }

    if (typeof value === 'number' && Number.isFinite(value)) {
      cell.removeAttribute('t');
      const v = cell.ownerDocument.createElementNS(ns, 'v');
      v.textContent = String(value);
      cell.appendChild(v);
      return;
    }

    const text = String(value);
    cell.setAttribute('t', 'inlineStr');

    const isNode = cell.ownerDocument.createElementNS(ns, 'is');
    const tNode = cell.ownerDocument.createElementNS(ns, 't');
    if (/^\s|\s$/.test(text) || /\s{2,}/.test(text)) {
      tNode.setAttribute('xml:space', 'preserve');
    }
    tNode.textContent = text;
    isNode.appendChild(tNode);
    cell.appendChild(isNode);
  };

  const patchWorksheetXml = (sheetXml, updates) => {
    const parser = new DOMParser();
    const doc = parser.parseFromString(sheetXml, 'application/xml');
    if (doc.getElementsByTagName('parsererror').length > 0) {
      throw new Error('시트 XML 파싱에 실패했습니다.');
    }

    const worksheet = doc.documentElement;
    const ns = worksheet.namespaceURI || 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';
    const sheetData = worksheet.getElementsByTagNameNS('*', 'sheetData')[0];
    if (!sheetData) {
      throw new Error('시트 데이터(sheetData)를 찾지 못했습니다.');
    }

    updates.forEach(({ ref, value }) => {
      const parsed = parseCellRef(ref);
      if (!parsed) return;
      const row = findOrCreateRow(sheetData, parsed.row, ns);
      const cell = findOrCreateCell(row, parsed.ref, parsed.col, ns);
      setXmlCellValue(cell, value, ns);
    });

    return new XMLSerializer().serializeToString(doc);
  };

  const normalizeWorksheetPath = (targetPath) => {
    const cleaned = String(targetPath || '').trim();
    if (!cleaned) return null;
    if (cleaned.startsWith('/')) return cleaned.replace(/^\/+/, '');
    if (cleaned.startsWith('xl/')) return cleaned;
    return `xl/${cleaned.replace(/^\.?\//, '')}`;
  };

  const buildSheetPathMap = (workbookXml, relsXml) => {
    const parser = new DOMParser();
    const workbookDoc = parser.parseFromString(workbookXml, 'application/xml');
    const relsDoc = parser.parseFromString(relsXml, 'application/xml');

    if (workbookDoc.getElementsByTagName('parsererror').length > 0 || relsDoc.getElementsByTagName('parsererror').length > 0) {
      throw new Error('템플릿 메타데이터 파싱에 실패했습니다.');
    }

    const relMap = new Map();
    Array.from(relsDoc.getElementsByTagNameNS('*', 'Relationship')).forEach((rel) => {
      const id = rel.getAttribute('Id');
      const target = rel.getAttribute('Target');
      const normalized = normalizeWorksheetPath(target);
      if (id && normalized) relMap.set(id, normalized);
    });

    const sheetMap = new Map();
    Array.from(workbookDoc.getElementsByTagNameNS('*', 'sheet')).forEach((sheet) => {
      const name = sheet.getAttribute('name');
      const relId = sheet.getAttribute('r:id') || sheet.getAttributeNS('http://schemas.openxmlformats.org/officeDocument/2006/relationships', 'id');
      if (!name || !relId) return;
      const path = relMap.get(relId);
      if (path) sheetMap.set(name, path);
    });

    return sheetMap;
  };

  const exportToTemplateExcel = async () => {
    try {
      const response = await fetch(dustTemplateUrl);
      if (!response.ok) throw new Error('템플릿 파일을 불러오지 못했습니다.');

      const templateBuffer = await response.arrayBuffer();
      const zip = await JSZip.loadAsync(templateBuffer);

      const workbookFile = zip.file('xl/workbook.xml');
      const relsFile = zip.file('xl/_rels/workbook.xml.rels');
      if (!workbookFile || !relsFile) {
        throw new Error('템플릿 workbook 구조를 찾지 못했습니다.');
      }

      const workbookXml = await workbookFile.async('string');
      const workbookRelsXml = await relsFile.async('string');
      const sheetPathMap = buildSheetPathMap(workbookXml, workbookRelsXml);

      const requiredSheets = ['표지', '연소가스자료', '노즐산정', '수분량', '기록지(수분량자동측정기)'];
      requiredSheets.forEach((sheetName) => {
        if (!sheetPathMap.has(sheetName)) {
          throw new Error(`템플릿 시트 누락: ${sheetName}`);
        }
      });

      const updatesBySheet = new Map();
      const addUpdate = (sheetName, ref, value) => {
        if (!updatesBySheet.has(sheetName)) updatesBySheet.set(sheetName, []);
        updatesBySheet.get(sheetName).push({ ref, value });
      };

      const gasComp = getGasComposition();
      const moisturePercent = calcTotalMoistureWeight() > 0 ? toNumber(calcPostMoisture()) : toNumber(calcMoisture());
      const isokineticPercent = toNumber(calcIsokineticRate(true));
      const avgTm = getRawAvgTm();
      const avgOrifice = getRawAvgOrifice();
      const avgSp = getRawAvgSp();
      const avgTp = getRawAvgTp();
      const avgDp = getRawAvgDp();
      const avgTs = getRawAvgTs();
      const stackPressure = getRawStackPressure();
      const flowRates = calcGasFlowRates();

      const gasMeters = formData.gasMeters.slice(0, 10);
      const firstVolume = gasMeters.find(m => m.volume !== '' && m.volume !== undefined)?.volume ?? '';
      const lastVolume = [...gasMeters].reverse().find(m => m.volume !== '' && m.volume !== undefined)?.volume ?? '';

      addUpdate('표지', 'D19', formData.company);
      addUpdate('표지', 'D20', formData.location);
      addUpdate('표지', 'D21', toExcelDateSerial(formData.date));

      for (let i = 0; i < 3; i++) {
        const gas = formData.gasAnalyzer[i] || {};
        const row = 8 + i;
        addUpdate('연소가스자료', `A${row}`, toExcelTimeFraction(gas.time));
        addUpdate('연소가스자료', `B${row}`, toNumber(gas.o2));
        addUpdate('연소가스자료', `C${row}`, toNumber(gas.co2));
        addUpdate('연소가스자료', `D${row}`, toNumber(gas.co));
        addUpdate('연소가스자료', `E${row}`, toNumber(gas.nox));
        addUpdate('연소가스자료', `F${row}`, toNumber(gas.sox));
      }

      addUpdate('노즐산정', 'C6', toNumber(formData.samplerId));
      addUpdate('노즐산정', 'B9', toNumber(formData.atmTemp));
      addUpdate('노즐산정', 'C9', toNumber(formData.atmPressure));
      addUpdate('노즐산정', 'D9', toNumber(formData.pitotFactor));
      addUpdate('노즐산정', 'E9', toNumber(formData.deltaHAt));
      addUpdate('노즐산정', 'F9', toNumber(formData.gasMeterFactor));
      addUpdate('노즐산정', 'E15', toNumber(avgTs));
      addUpdate('노즐산정', 'F15', toNumber(formData.traverseTmIn));
      addUpdate('노즐산정', 'G15', toNumber(formData.traverseTmOut));
      addUpdate('노즐산정', 'H15', toNumber(avgTp));
      addUpdate('노즐산정', 'I15', toNumber(avgSp));
      addUpdate('노즐산정', 'J15', toNumber(avgDp));
      addUpdate('노즐산정', 'B39', toNumber(formData.planSamplingTime));
      addUpdate('노즐산정', 'K39', toNumber(formData.usedNozzleNum) || toNumber(formData.recommendedNozzleNum));

      for (let i = 0; i < 5; i++) {
        const point = formData.points[i] || {};
        const row = 17 + i;
        addUpdate('노즐산정', `H${row}`, toNumber(point.tp));
        addUpdate('노즐산정', `I${row}`, toNumber(point.sp));
        addUpdate('노즐산정', `J${row}`, toNumber(point.dp));
      }

      for (let i = 0; i < 18; i++) {
        const row = 39 + i;
        const nozzle = nozzleSet[i];
        addUpdate('노즐산정', `C${row}`, nozzle ? toNumber(nozzle.num) : null);
        addUpdate('노즐산정', `D${row}`, nozzle ? toNumber(nozzle.d) : null);
      }

      addUpdate('수분량', 'D6', 1);
      for (let i = 0; i < 5; i++) {
        addUpdate('수분량', `B${11 + i}`, i + 1);
        addUpdate('수분량', `C${11 + i}`, toNumber(formData.moistureValues[i]));
      }
      for (let i = 0; i < 4; i++) {
        const imp = formData.impingers[i] || {};
        addUpdate('수분량', `D${34 + i}`, toNumber(imp.initial));
        addUpdate('수분량', `E${34 + i}`, toNumber(imp.final));
      }
      addUpdate('수분량', 'B29', toNumber(firstVolume));
      addUpdate('수분량', 'C29', toNumber(lastVolume));
      addUpdate('수분량', 'E29', Number.isFinite(avgTm) ? avgTm : null);
      addUpdate('수분량', 'F29', toNumber(formData.atmPressure));
      addUpdate('수분량', 'G29', Number.isFinite(avgOrifice) ? avgOrifice : null);

      addUpdate('기록지(수분량자동측정기)', 'C5', formData.company);
      addUpdate('기록지(수분량자동측정기)', 'C6', formData.location);
      addUpdate('기록지(수분량자동측정기)', 'C7', formData.sampler);
      addUpdate('기록지(수분량자동측정기)', 'C8', toExcelDateSerial(formData.date));
      addUpdate('기록지(수분량자동측정기)', 'C9', formData.filterId || '먼지-1');
      addUpdate('기록지(수분량자동측정기)', 'J5', toNumber(formData.pitotFactor));
      addUpdate('기록지(수분량자동측정기)', 'J6', toNumber(formData.atmTemp));
      addUpdate('기록지(수분량자동측정기)', 'J7', toNumber(formData.atmPressure));
      addUpdate('기록지(수분량자동측정기)', 'J8', Number.isFinite(stackPressure) ? stackPressure : null);
      addUpdate('기록지(수분량자동측정기)', 'J9', moisturePercent);
      addUpdate('기록지(수분량자동측정기)', 'J10', toNumber(formData.totalStackDepth));
      addUpdate('기록지(수분량자동측정기)', 'J11', toNumber(formData.nozzleDiameter) ? toNumber(formData.nozzleDiameter) / 10 : null);
      addUpdate('기록지(수분량자동측정기)', 'J13', formData.filterId);
      addUpdate('기록지(수분량자동측정기)', 'J14', toNumber(avgSp));
      addUpdate('기록지(수분량자동측정기)', 'C11', toNumber(gasComp.o2));
      addUpdate('기록지(수분량자동측정기)', 'C12', toNumber(formData.kFactor));
      addUpdate('기록지(수분량자동측정기)', 'C13', isokineticPercent);
      addUpdate('기록지(수분량자동측정기)', 'B14', formData.standardO2 ? toNumber(formData.standardO2) : '-');
      addUpdate('기록지(수분량자동측정기)', 'D14', toNumber(formData.filterInitial));
      addUpdate('기록지(수분량자동측정기)', 'F14', toNumber(formData.filterFinal));

      for (let i = 0; i < 10; i++) {
        const row = 17 + i;
        addUpdate('기록지(수분량자동측정기)', `A${row}`, null);
        addUpdate('기록지(수분량자동측정기)', `B${row}`, null);
        addUpdate('기록지(수분량자동측정기)', `C${row}`, null);
        addUpdate('기록지(수분량자동측정기)', `D${row}`, null);
        addUpdate('기록지(수분량자동측정기)', `E${row}`, null);
        addUpdate('기록지(수분량자동측정기)', `F${row}`, null);
        addUpdate('기록지(수분량자동측정기)', `G${row}`, null);
        addUpdate('기록지(수분량자동측정기)', `H${row}`, null);
        addUpdate('기록지(수분량자동측정기)', `I${row}`, null);
        addUpdate('기록지(수분량자동측정기)', `J${row}`, null);
        addUpdate('기록지(수분량자동측정기)', `K${row}`, null);
      }

      gasMeters.forEach((meter, idx) => {
        const row = 17 + idx;
        addUpdate('기록지(수분량자동측정기)', `A${row}`, meter.pointNum || (idx === 0 ? '시작' : ''));
        addUpdate('기록지(수분량자동측정기)', `B${row}`, toNumber(meter.time));
        addUpdate('기록지(수분량자동측정기)', `C${row}`, toNumber(meter.vacuum));
        addUpdate('기록지(수분량자동측정기)', `D${row}`, toNumber(meter.stackTemp));
        addUpdate('기록지(수분량자동측정기)', `E${row}`, toNumber(meter.dp));
        addUpdate('기록지(수분량자동측정기)', `F${row}`, toNumber(meter.pressure));
        addUpdate('기록지(수분량자동측정기)', `G${row}`, toNumber(meter.volume));
        addUpdate('기록지(수분량자동측정기)', `H${row}`, toNumber(meter.tmIn));
        addUpdate('기록지(수분량자동측정기)', `I${row}`, toNumber(meter.tmOut));
        addUpdate('기록지(수분량자동측정기)', `K${row}`, toNumber(meter.impingerTemp));
      });

      if (sheetPathMap.has('먼지측정결과보고서')) {
        addUpdate('먼지측정결과보고서', 'C5', formData.company);
        addUpdate('먼지측정결과보고서', 'C6', formData.location);
        addUpdate('먼지측정결과보고서', 'A17', formData.remarks);
        addUpdate('먼지측정결과보고서', 'G13', toNumber(flowRates.dry));
      }

      for (const [sheetName, updates] of updatesBySheet.entries()) {
        const sheetPath = sheetPathMap.get(sheetName);
        if (!sheetPath) continue;
        const file = zip.file(sheetPath);
        if (!file) {
          throw new Error(`시트 파일 누락: ${sheetName} (${sheetPath})`);
        }
        const xml = await file.async('string');
        const patched = patchWorksheetXml(xml, updates);
        zip.file(sheetPath, patched);
      }

      const output = await zip.generateAsync({ type: 'arraybuffer', compression: 'DEFLATE' });
      const blob = new Blob([output], { type: 'application/vnd.ms-excel.sheet.macroEnabled.12' });
      const url = URL.createObjectURL(blob);

      const link = document.createElement('a');
      link.href = url;
      link.setAttribute('download', `${activeSheet?.title || '먼지시료채취기록부'}_${formData.date || new Date().toISOString().slice(0, 10)}.xlsm`);
      document.body.appendChild(link);
      link.click();
      document.body.removeChild(link);
      URL.revokeObjectURL(url);
    } catch (error) {
      console.error(error);
      alert('템플릿 엑셀 저장 중 오류가 발생했습니다. 템플릿 파일 또는 입력값을 확인해주세요.');
    }
  };

  const appGasComp = getGasComposition();
  const o2 = appGasComp.o2, co2 = appGasComp.co2, co = appGasComp.co, sox = appGasComp.sox, nox = appGasComp.nox, n2 = appGasComp.n2, Md = appGasComp.Md, Ms = appGasComp.Ms, r0 = appGasComp.r0;
  
  const currentTs = getRawAvgTs();
  const currentPs = getRawStackPressure();
  const r = (!isNaN(r0) && !isNaN(currentTs) && !isNaN(currentPs)) ? r0 * (273 / (273 + currentTs)) * (currentPs / 760) : NaN;
  const samplingPointsData = getSamplingPoints();
  const configuredNozzles = getConfiguredNozzles();
  
  const getExpectedValues = () => {
    const d = parseFloat(formData.nozzleDiameter), k = parseFloat(formData.kFactor), dp = getRawAvgDp(), Vs = getRawGasVelocity();

    let dH = '-', L = '-', SL = '-', Sm3 = '-', reqTime = '-';
    if (!isNaN(k) && !isNaN(dp)) dH = calcExpectedOrificeDH(k, dp);

    const Ts = getRawAvgTs(), Ps = getRawStackPressure(), Pm = parseFloat(formData.atmPressure);
    
    // 동정압 측정 시 입력한 예비조사 미터온도를 기반으로 산출
    const tmIn = parseFloat(formData.traverseTmIn) || parseFloat(formData.atmTemp) || 25;
    const tmOut = parseFloat(formData.traverseTmOut) || parseFloat(formData.atmTemp) || 25;
    const Tm = (tmIn + tmOut) / 2;

    const gasComp = getGasComposition();
    const Xw = gasComp.Xw;
    const Yd = parseFloat(formData.gasMeterFactor) || 1.0;
    const targetVol = parseFloat(formData.targetVolume) || 1000;

    if (!isNaN(d) && !isNaN(Vs) && Vs > 0 && !isNaN(Ts) && !isNaN(Ps) && Pm > 0 && !isNaN(Xw)) {
        const An = Math.PI * Math.pow(d / 2000, 2);
        const Q_m = (Vs * An * 60 * 1000 * ((Tm + 273) / (Ts + 273)) * (Ps / Pm) * (1 - Xw / 100)) / Yd;
        const Q_sl = Q_m * Yd * (273 / (273 + Tm)) * (Pm / 760);
        
        const calcTime = Math.ceil(targetVol / Q_sl);
        if (isFinite(calcTime) && calcTime > 0) {
            reqTime = calcTime;
            const Vm_val = Q_m * reqTime;
            L = Vm_val.toFixed(1);
            const Vsl_val = Q_sl * reqTime;
            SL = Vsl_val.toFixed(1);
            Sm3 = (Vsl_val / 1000).toFixed(3);
        }
    }
    return { dH, L, SL, Sm3, reqTime };
  };

  const expData = getExpectedValues();
  const activeSheet = SHEET_MENU.find(item => item.id === selectedSheet);
  const activeTheme = SHEET_THEMES[selectedSheet] || SHEET_THEMES.dust;
  const isDustSheet = selectedSheet === 'dust';
  const isNightSky = skyPreviewMode === 'auto' ? skyPhase === 'night' : skyPreviewMode === 'night';
  const sceneStarsMenu = lowSpecMode ? SKYLINE_STARS.slice(0, 12) : SKYLINE_STARS;
  const sceneStarsSheet = lowSpecMode ? SKYLINE_STARS.slice(0, 10) : SKYLINE_STARS.slice(0, 30);
  const sceneFlyers = lowSpecMode ? SKYLINE_FLYERS.slice(0, 4) : SKYLINE_FLYERS;
  const sceneMountains = lowSpecMode ? SKYLINE_MOUNTAINS.slice(0, 2) : SKYLINE_MOUNTAINS;
  const sceneTreeBelt = lowSpecMode ? SKYLINE_TREE_BELT.slice(0, 14) : SKYLINE_TREE_BELT;
  const sceneFrontTreeBelt = lowSpecMode ? SKYLINE_FRONT_TREE_BELT.slice(0, 12) : SKYLINE_FRONT_TREE_BELT;
  const sceneBackTowers = lowSpecMode ? SKYLINE_BACK_TOWERS.slice(0, 5) : SKYLINE_BACK_TOWERS;
  const sceneFactoryBlocks = lowSpecMode ? SKYLINE_FACTORY_BLOCKS.slice(0, 3) : SKYLINE_FACTORY_BLOCKS;
  const sceneCoolingTowers = lowSpecMode ? SKYLINE_COOLING_TOWERS.slice(0, 1) : SKYLINE_COOLING_TOWERS;
  const sceneStacks = lowSpecMode ? SKYLINE_STACKS.slice(0, 3) : SKYLINE_STACKS;
  const getSmokeParticleTotal = (stack) => (lowSpecMode ? Math.min(4, stack.smokeCount) : stack.smokeCount + 4);
  const navigateToMenu = () => {
    setSelectedSheet('');
    if (window.location.hash) {
      window.location.hash = '';
    }
  };
  const navigateToSheet = (sheetId) => {
    if (!activeUser || !isUserUnlocked) {
      alert('초기 화면에서 user 로그인 후 이용해주세요.');
      return;
    }
    if (selectedSheet === sheetId) return;
    window.location.hash = `sheet=${sheetId}`;
  };

  useEffect(() => {
    const applyHashRoute = () => {
      const rawHash = window.location.hash.startsWith('#') ? window.location.hash.slice(1) : '';
      const params = new URLSearchParams(rawHash);
      const hashSheet = params.get('sheet');
      if (!hashSheet || !SHEET_MENU.some(sheet => sheet.id === hashSheet)) {
        setSelectedSheet('');
        return;
      }
      if (!activeUser || !isUserUnlocked) {
        setSelectedSheet('');
        return;
      }
      setSelectedSheet(hashSheet);
    };

    applyHashRoute();
    window.addEventListener('hashchange', applyHashRoute);
    return () => window.removeEventListener('hashchange', applyHashRoute);
  }, [activeUser, isUserUnlocked]);

  useEffect(() => {
    if (!selectedSheet) return;
    window.requestAnimationFrame(() => {
      window.scrollTo({ top: 0, left: 0, behavior: 'auto' });
    });
  }, [selectedSheet]);

  useEffect(() => {
    if (selectedSheet) return undefined;
    const spawnLegendary = () => {
      setLegendaryFlight(pickRandomLegendaryFlight());
    };
    spawnLegendary();
    const timer = window.setInterval(spawnLegendary, 60 * 1000);
    return () => window.clearInterval(timer);
  }, [selectedSheet]);

  useEffect(() => {
    const refreshSkyPhase = () => {
      setSkyPhase(getSkyPhaseByHour());
    };
    refreshSkyPhase();
    const timer = window.setInterval(refreshSkyPhase, 60 * 1000);
    return () => window.clearInterval(timer);
  }, []);

  if (!selectedSheet) {
    return (
      <div className={`relative min-h-screen overflow-hidden p-6 md:p-10 font-sans text-slate-800 ${isNightSky ? 'bg-gradient-to-b from-[#050914] via-[#264284] to-[#7a5a8c]' : 'bg-gradient-to-b from-[#99ceff] via-[#8ab8ef] to-[#f1b7c8]'}`}>
        <style>{SKYLINE_FLYER_CSS}</style>
        <div className="fixed left-4 bottom-4 z-40 flex items-center gap-1 rounded-xl border border-white/70 bg-white/80 px-2 py-1 backdrop-blur">
          {SKY_PREVIEW_OPTIONS.map((opt) => (
            <button
              key={`menu-sky-${opt.id}`}
              type="button"
              onClick={() => setSkyPreviewMode(opt.id)}
              className={`px-2 py-1 rounded text-[11px] font-black transition-colors ${skyPreviewMode === opt.id ? 'bg-slate-800 text-white' : 'text-slate-700 hover:bg-slate-100'}`}
            >
              {opt.label}
            </button>
          ))}
          <button
            type="button"
            onClick={() => setLowSpecMode((prev) => !prev)}
            className={`px-2 py-1 rounded text-[11px] font-black transition-colors ${lowSpecMode ? 'bg-emerald-600 text-white' : 'text-slate-700 hover:bg-slate-100'}`}
          >
            경량
          </button>
        </div>
        <div className="pointer-events-none absolute inset-0 z-0">
          <div className={`absolute inset-0 ${isNightSky ? 'bg-gradient-to-b from-[#040915]/45 via-[#1d2f64]/15 to-transparent' : 'bg-gradient-to-b from-white/35 via-sky-100/10 to-transparent'}`} />
          <div className={`absolute -left-20 -top-24 h-80 w-80 rounded-full ${lowSpecMode ? 'blur-xl' : 'blur-3xl'} ${isNightSky ? 'bg-indigo-300/20' : 'bg-white/80'}`} />
          <div className={`absolute right-4 top-2 h-72 w-72 rounded-full ${lowSpecMode ? 'blur-xl' : 'blur-3xl'} ${isNightSky ? 'bg-violet-400/15' : 'bg-sky-200/65'}`} />
          <div className={isNightSky ? 'skyline-moon' : 'skyline-sun'} />
          {isNightSky && (
            <div className="absolute inset-0">
              {sceneStarsMenu.map((star) => (
                <span
                  key={star.id}
                  className="skyline-star"
                  style={{
                    left: star.left,
                    top: star.top,
                    width: `${star.size}px`,
                    height: `${star.size}px`,
                    '--twinkle-duration': `${star.duration}s`,
                    '--twinkle-delay': `-${star.delay}s`,
                  }}
                />
              ))}
            </div>
          )}
          <div className="absolute inset-0 overflow-hidden">
            {sceneFlyers.map((flyer, idx) => {
              const visibleStartOffset = Math.min(flyer.duration * (0.24 + (idx % 5) * 0.12), flyer.duration * 0.8);
              return (
                <div
                  key={flyer.id}
                  className="skyline-flyer opacity-90"
                  style={{
                    top: flyer.top,
                    animationName: flyer.direction === 'right' ? 'skylineFlyRight' : 'skylineFlyLeft',
                    animationDuration: `${flyer.duration}s`,
                    animationTimingFunction: 'linear',
                    animationIterationCount: 'infinite',
                    animationDelay: `-${visibleStartOffset}s`,
                  }}
                >
                  <div className="skyline-flyer-inner" style={{ '--bob-duration': `${flyer.bob}s` }}>
                    <img
                      src={`https://raw.githubusercontent.com/PokeAPI/sprites/master/sprites/pokemon/${flyer.id}.png`}
                      alt={flyer.name}
                      className="block"
                      style={{ width: `${SKYLINE_FLYER_UNIFIED_SIZE}px`, height: `${SKYLINE_FLYER_UNIFIED_SIZE}px`, imageRendering: 'pixelated' }}
                    />
                  </div>
                </div>
              );
            })}
            {legendaryFlight && (
              <div
                key={legendaryFlight.key}
                className="legendary-flyer"
                style={{
                  top: legendaryFlight.top,
                  left: legendaryFlight.left,
                  animationName: legendaryFlight.direction === 'right' ? 'legendaryDiagonalRight' : 'legendaryDiagonalLeft',
                  animationDuration: `${legendaryFlight.duration}s`,
                  animationTimingFunction: 'ease-out',
                  animationIterationCount: '1',
                  animationFillMode: 'forwards',
                  animationDelay: `-${legendaryFlight.startOffset}s`,
                }}
              >
                <img
                  src={legendaryFlight.url}
                  alt={legendaryFlight.name}
                  style={{ width: `${legendaryFlight.size}px`, height: `${legendaryFlight.size}px`, objectFit: 'contain' }}
                />
              </div>
            )}
          </div>
          <div className="absolute bottom-0 left-0 right-0">
            <div className="relative mx-auto h-[25rem] md:h-[34rem] max-w-7xl">
              <div className={`absolute bottom-0 left-0 right-0 h-24 md:h-32 ${isNightSky ? 'bg-gradient-to-b from-[#233a33] via-[#16282c] to-[#080f15]' : 'bg-gradient-to-b from-[#5e8c67] via-[#4e7458] to-[#2c3f37]'}`} />
              <div className={`absolute bottom-16 left-0 right-0 h-14 ${isNightSky ? 'bg-gradient-to-t from-[#0c1424]/60 to-transparent' : 'bg-gradient-to-t from-[#d3e6ff]/40 to-transparent'}`} />
              {sceneMountains.map((mountain) => (
                <div
                  key={mountain.id}
                  className={`absolute bottom-[6.4rem] rounded-t-[45%] ${isNightSky ? 'bg-gradient-to-b from-[#35446c]/80 to-[#1a243f]/75' : 'bg-gradient-to-b from-[#90a9c9]/78 to-[#6784aa]/72'}`}
                  style={{ left: mountain.left, width: mountain.width, height: `${mountain.height}px` }}
                />
              ))}
              {sceneTreeBelt.map((tree, treeIdx) => (
                <div
                  key={tree.id}
                  className="absolute bottom-[5.7rem]"
                  style={{
                    left: tree.left,
                    width: `${Math.max(12, Math.round(tree.width * 1.18))}px`,
                    height: `${Math.max(18, Math.round(tree.height * 1.18))}px`,
                  }}
                >
                  {(() => {
                    const pixel = Math.max(2, Math.round(tree.width / 6));
                    const variant = PIXEL_TREE_VARIANTS[(treeIdx * 3 + tree.width) % PIXEL_TREE_VARIANTS.length];
                    const palette = getPixelTreePalette(isNightSky, treeIdx + (tree.height % 5));
                    return (
                      <>
                        {variant.cells.map((cell, idx) => (
                          <span
                            key={`${tree.id}-cell-${idx}`}
                            className="absolute"
                            style={{
                              left: `${cell[0] * pixel}px`,
                              bottom: `${(cell[1] + 2) * pixel}px`,
                              width: `${pixel}px`,
                              height: `${pixel}px`,
                              background: idx % 4 === 0 ? palette.alt : palette.main,
                            }}
                          />
                        ))}
                        <span
                          className="absolute"
                          style={{
                            left: `${variant.trunk.x * pixel}px`,
                            bottom: `${pixel}px`,
                            width: `${pixel}px`,
                            height: `${variant.trunk.h * pixel}px`,
                            background: palette.trunk,
                          }}
                        />
                        <span
                          className="absolute"
                          style={{
                            left: `${variant.trunk.baseX * pixel}px`,
                            bottom: '0px',
                            width: `${variant.trunk.baseW * pixel}px`,
                            height: `${pixel}px`,
                            background: palette.trunk,
                          }}
                        />
                      </>
                    );
                  })()}
                </div>
              ))}

              {sceneBackTowers.map((tower) => (
                <div
                  key={tower.id}
                  className="absolute bottom-[5.4rem] rounded-t-[2px] border-2 border-[#2b3558]/80 bg-gradient-to-b from-[#6c7cab] via-[#495a8a] to-[#243052] shadow-[inset_-6px_-8px_0_rgba(15,23,42,0.36)]"
                  style={{ left: tower.left, width: tower.width, height: `${tower.height}px` }}
                >
                  <div className="absolute inset-x-[10%] top-[22%] h-[2px] bg-[#9bb0df]/45" />
                  <div className="absolute inset-x-[10%] top-[45%] h-[2px] bg-[#9bb0df]/34" />
                  <div className="absolute inset-x-[10%] top-[68%] h-[2px] bg-[#9bb0df]/26" />
                </div>
              ))}

              {sceneFactoryBlocks.map((block) => (
                <div
                  key={block.id}
                  className="absolute bottom-[4.6rem] rounded-t-[3px] border-2 border-[#2f2d45] bg-gradient-to-b from-[#8f7668] via-[#6e5850] to-[#3d3033] shadow-[inset_-8px_-8px_0_rgba(15,23,42,0.3)]"
                  style={{ left: block.left, width: block.width, height: `${block.height}px` }}
                >
                  <div className="absolute -top-[10px] inset-x-[4%] h-[10px] border-2 border-[#2f2d45] bg-[#6c5a58]" style={{ clipPath: `polygon(0 100%, ${block.roof}% 0, ${100 - block.roof}% 0, 100% 100%)` }} />
                  <div className="absolute inset-x-[8%] top-[20%] h-[2px] bg-[#e2d1be]/22" />
                  <div className="absolute inset-x-[8%] top-[44%] h-[2px] bg-[#e2d1be]/18" />
                  <div className="absolute inset-x-[8%] top-[68%] h-[2px] bg-[#e2d1be]/14" />
                  {Array.from({ length: block.windows }).map((_, windowIdx) => (
                    <div
                      key={`${block.id}-window-${windowIdx}`}
                      className="absolute h-[4px] w-[4px] bg-[#fff2bf]/82"
                      style={{
                        left: `${16 + windowIdx * 14}%`,
                        top: '60%',
                        boxShadow: '0 12px 0 rgba(255, 242, 191, 0.8)',
                      }}
                    />
                  ))}
                </div>
              ))}

              {sceneCoolingTowers.map((tower) => (
                <div
                  key={tower.id}
                  className="absolute bottom-[4.6rem] rounded-t-[42%] border-2 border-[#4c5a71] bg-gradient-to-b from-[#dde3ee] via-[#b7bfcd] to-[#70798d] shadow-[inset_-10px_-8px_0_rgba(15,23,42,0.22)]"
                  style={{ left: tower.left, width: `${tower.width}px`, height: `${tower.height}px` }}
                >
                  <div className="absolute inset-x-[13%] top-[14%] h-[5px] bg-[#6c7f97]/60 rounded-full" />
                </div>
              ))}

              {sceneStacks.map((stack) => (
                (() => {
                  const palette = stack.variant === 'brick'
                    ? {
                        body: 'from-[#cdbfb1] via-[#a3948a] to-[#655c63]',
                        cap: 'bg-[#5b5563]',
                        lineA: '#e8ddd4',
                        lineB: '#c9bbb0',
                      }
                    : stack.variant === 'concrete'
                      ? {
                          body: 'from-[#c9d7f2] via-[#9db5dc] to-[#647ea9]',
                          cap: 'bg-[#556d95]',
                          lineA: '#ebf2ff',
                          lineB: '#cddcf8',
                        }
                      : {
                          body: 'from-[#b9cce6] via-[#8ea9cf] to-[#5d769b]',
                          cap: 'bg-[#4f6489]',
                          lineA: '#e1ebff',
                          lineB: '#bed2ef',
                        };
                  return (
                    <div
                      key={stack.id}
                      className="absolute bottom-[4.6rem]"
                      style={{ left: stack.left, width: `${stack.width}px`, height: `${stack.height}px` }}
                    >
                      <div className={`absolute inset-0 border-2 border-[#2c3147] bg-gradient-to-b ${palette.body} shadow-[inset_-4px_-6px_0_rgba(15,23,42,0.26)]`} />
                      <div className={`absolute -top-[6px] left-1/2 h-[6px] w-[78%] -translate-x-1/2 border-2 border-[#2c3147] ${palette.cap}`} />
                      <div className="absolute left-[22%] top-[10%] w-[2px] h-[74%] bg-white/20" />
                      <div className="absolute left-[58%] top-[18%] w-[2px] h-[64%] bg-black/18" />
                      <div className="absolute inset-x-[14%] top-[16%] h-[4px]" style={{ background: palette.lineA }} />
                      <div className="absolute inset-x-[14%] top-[30%] h-[3px]" style={{ background: palette.lineB }} />
                      <div className="absolute inset-x-[14%] top-[44%] h-[4px]" style={{ background: palette.lineA }} />
                      <div className="absolute inset-x-[14%] top-[58%] h-[3px]" style={{ background: palette.lineB }} />
                      {Array.from({ length: getSmokeParticleTotal(stack) }).map((_, smokeIdx) => (
                        <span
                          key={`${stack.id}-smoke-${smokeIdx}`}
                          className="stack-heart-smoke"
                          style={buildSmokeParticleStyle(stack.id, smokeIdx)}
                        />
                      ))}
                    </div>
                  );
                })()
              ))}
              {sceneFrontTreeBelt.map((tree, treeIdx) => (
                <div
                  key={tree.id}
                  className="absolute bottom-[3.9rem]"
                  style={{
                    left: tree.left,
                    width: `${Math.max(14, Math.round(tree.width * 1.16))}px`,
                    height: `${Math.max(22, Math.round(tree.height * 1.12))}px`,
                  }}
                >
                  {(() => {
                    const pixel = Math.max(2, Math.round(tree.width / 6));
                    const variant = PIXEL_TREE_VARIANTS[(treeIdx * 5 + tree.height) % PIXEL_TREE_VARIANTS.length];
                    const palette = getPixelTreePalette(isNightSky, treeIdx + 2 + (tree.width % 5));
                    return (
                      <>
                        {variant.cells.map((cell, idx) => (
                          <span
                            key={`${tree.id}-cell-${idx}`}
                            className="absolute"
                            style={{
                              left: `${cell[0] * pixel}px`,
                              bottom: `${(cell[1] + 2) * pixel}px`,
                              width: `${pixel}px`,
                              height: `${pixel}px`,
                              background: idx % 4 === 0 ? palette.alt : palette.main,
                            }}
                          />
                        ))}
                        <span
                          className="absolute"
                          style={{
                            left: `${variant.trunk.x * pixel}px`,
                            bottom: `${pixel}px`,
                            width: `${pixel}px`,
                            height: `${variant.trunk.h * pixel}px`,
                            background: palette.trunk,
                          }}
                        />
                        <span
                          className="absolute"
                          style={{
                            left: `${variant.trunk.baseX * pixel}px`,
                            bottom: '0px',
                            width: `${variant.trunk.baseW * pixel}px`,
                            height: `${pixel}px`,
                            background: palette.trunk,
                          }}
                        />
                      </>
                    );
                  })()}
                </div>
              ))}
              <div className={`absolute bottom-0 left-0 right-0 h-[34px] ${isNightSky ? 'bg-gradient-to-b from-[#10222f] to-[#04080f]' : 'bg-gradient-to-b from-[#4f7058] to-[#27372e]'}`} />
            </div>
          </div>
        </div>

        <div className="relative z-30 max-w-6xl mx-auto">
          <div className="mb-6 p-2 text-center">
            <h1 className="text-slate-900 tracking-tight flex justify-center">
              <img
                src={mainLogoSrc}
                alt="STACKPILOT"
                onError={() => setLogoIndex((prev) => Math.min(prev + 1, logoCandidates.length - 1))}
                className="h-[2.6rem] md:h-[4rem] w-auto select-none"
              />
            </h1>
            {isUserUnlocked && (
              <p className="text-sm text-slate-600 mt-2">로그인 완료: 원하는 기록부 아이콘을 누르면 해당 기록부로 이동합니다.</p>
            )}
          </div>
          <div className="bg-white/85 backdrop-blur p-4 rounded-2xl shadow-lg border border-white/70 mb-6">
            {!isUserUnlocked ? (
              <div className="relative overflow-hidden border border-emerald-200 rounded-2xl p-5 bg-gradient-to-br from-emerald-50 via-white to-emerald-100/70 shadow-inner text-center">
                <div className="pointer-events-none absolute -right-10 -top-10 h-28 w-28 rounded-full bg-emerald-200/70 blur-2xl" />
                <p className="text-sm font-black text-emerald-900">베타 접속</p>
                <p className="text-xs text-emerald-800 mt-1">회원가입 없이 `user` 고정 계정으로 바로 사용합니다.</p>
                <div className="mt-3 flex flex-wrap items-center justify-center gap-2">
                  <button
                    type="button"
                    onClick={handleBetaQuickLogin}
                    className="px-4 py-2 bg-emerald-600 text-white rounded-lg font-bold text-sm hover:bg-emerald-700 shadow-sm"
                  >
                    user 로그인
                  </button>
                  <button
                    type="button"
                    onClick={openAdminLoginModal}
                    className="px-4 py-2 bg-slate-700 text-white rounded-lg font-bold text-sm hover:bg-slate-800 shadow-sm"
                  >
                    관리자 로그인
                  </button>
                </div>

                {authModal === 'admin' && (
                  <div className="fixed inset-0 z-50 bg-black/40 p-4 flex items-center justify-center" onClick={() => setAuthModal('')}>
                    <div className="w-full max-w-md bg-white rounded-2xl border border-slate-200 shadow-xl p-5" onClick={(e) => e.stopPropagation()}>
                      <div className="flex items-center justify-between mb-4">
                        <h3 className="text-base font-black text-slate-900">관리자 로그인</h3>
                        <button type="button" onClick={() => setAuthModal('')} className="px-2 py-1 rounded hover:bg-slate-100 text-slate-600 text-xs font-bold">
                          닫기
                        </button>
                      </div>
                      <div className="space-y-2">
                        <input
                          type="text"
                          value={loginUserId}
                          onChange={(e) => setLoginUserId(e.target.value)}
                          className="w-full p-2 border border-slate-300 rounded-lg bg-white outline-none focus:ring-2 focus:ring-emerald-500"
                          placeholder="아이디"
                        />
                        <input
                          type="password"
                          value={loginPassword}
                          onChange={(e) => setLoginPassword(e.target.value)}
                          onKeyDown={(e) => {
                            if (e.key === 'Enter') {
                              e.preventDefault();
                              handleLoginUser();
                            }
                          }}
                          className="w-full p-2 border border-slate-300 rounded-lg bg-white outline-none focus:ring-2 focus:ring-emerald-500"
                          placeholder="비밀번호"
                        />
                        <button
                          type="button"
                          onClick={handleLoginUser}
                          className="w-full py-2 bg-slate-800 text-white rounded-lg font-bold text-sm hover:bg-slate-900"
                        >
                          로그인
                        </button>
                      </div>
                    </div>
                  </div>
                )}
              </div>
            ) : (
              <div className="space-y-4 flex flex-col items-center">
                <div className="w-full max-w-3xl border border-slate-200 rounded-xl p-4 bg-slate-50">
                  <div className="flex flex-col items-center md:flex-row md:items-center gap-4 justify-center">
                    <div className="w-24 h-24 rounded-xl border border-slate-300 bg-white shadow-sm overflow-hidden shrink-0 flex items-center justify-center">
                      {profileAvatarUrl ? (
                        <img src={profileAvatarUrl} alt={`${activeUser} 프로필`} className="w-full h-full object-cover" />
                      ) : (
                        <span className="text-slate-400 text-xs font-bold">NO IMAGE</span>
                      )}
                    </div>
                    <div className="flex-1 text-center md:text-left">
                      <p className="text-sm text-slate-600">현재 사용자</p>
                      <p className="text-xl font-black text-slate-900">{profileNickname || activeUser}</p>
                      {profileNickname && (
                        <p className="text-xs text-slate-500 mt-0.5">아이디: {activeUser}</p>
                      )}
                      <p className="text-xs text-slate-600 mt-1">
                        상태: <span className="font-black text-emerald-700">로그인됨</span>
                        {' '}| 저장 건수: <span className="font-black text-emerald-700">{savedData.length}</span>건
                      </p>
                      <div className="mt-3 flex flex-wrap items-center justify-center md:justify-start gap-2">
                        <button
                          type="button"
                          onClick={() => setAuthModal('nickname')}
                          className="px-3 py-1.5 rounded border border-slate-300 bg-white text-slate-700 hover:bg-slate-100 font-bold text-xs"
                        >
                          별명 설정
                        </button>
                        <button
                          type="button"
                          onClick={handleLockActiveUser}
                          className="px-3 py-1.5 rounded bg-slate-700 text-white hover:bg-slate-800 font-bold text-xs"
                        >
                          로그아웃
                        </button>
                      </div>
                    </div>
                  </div>
                </div>

                {authModal === 'nickname' && (
                  <div className="fixed inset-0 z-50 bg-black/40 p-4 flex items-center justify-center" onClick={() => setAuthModal('')}>
                    <div className="w-full max-w-lg bg-white rounded-2xl border border-slate-200 shadow-xl p-5" onClick={(e) => e.stopPropagation()}>
                      <div className="flex items-center justify-between mb-4">
                        <h3 className="text-base font-black text-slate-900">별명 설정</h3>
                        <button type="button" onClick={() => setAuthModal('')} className="px-2 py-1 rounded hover:bg-slate-100 text-slate-600 text-xs font-bold">
                          닫기
                        </button>
                      </div>
                      <div className="flex gap-2">
                        <input
                          type="text"
                          value={profileNickname}
                          onChange={(e) => setProfileNickname(e.target.value)}
                          className="flex-1 p-2 border border-slate-300 rounded-lg bg-white outline-none focus:ring-2 focus:ring-emerald-500"
                          placeholder="별명 입력 (최대 24자)"
                          maxLength={24}
                        />
                        <button
                          type="button"
                          onClick={() => {
                            const saved = handleSaveProfileNickname();
                            if (saved) setAuthModal('');
                          }}
                          className="px-4 py-2 bg-emerald-600 text-white rounded-lg font-bold text-sm hover:bg-emerald-700"
                        >
                          저장
                        </button>
                      </div>
                    </div>
                  </div>
                )}

                <div className="flex flex-wrap justify-center gap-2">
                  <button
                    type="button"
                    onClick={clearActiveUserReports}
                    className="px-3 py-1.5 rounded border border-red-300 text-red-600 hover:bg-red-50 font-bold text-xs"
                  >
                    리포트 전체 삭제
                  </button>
                  <button
                    type="button"
                    onClick={handleDeleteActiveUser}
                    className="px-3 py-1.5 rounded bg-red-600 text-white hover:bg-red-700 font-bold text-xs"
                  >
                    현재 계정 삭제
                  </button>
                </div>
              </div>
            )}
          </div>
          {isUserUnlocked && (
            <div className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-4 gap-5">
              {SHEET_MENU.map((sheet) => {
                return (
                  <button
                    key={sheet.id}
                    type="button"
                    onClick={() => navigateToSheet(sheet.id)}
                    className="group relative overflow-hidden text-left bg-white border border-slate-200 rounded-2xl p-5 shadow-sm transition-all hover:shadow-lg hover:-translate-y-0.5"
                  >
                    <div className={`absolute -right-8 -top-8 w-28 h-28 rounded-full bg-gradient-to-br ${sheet.toneClass} blur-sm group-hover:scale-110 transition-transform`} />
                    <h2 className="text-base font-black text-slate-900">{sheet.title}</h2>
                    <p className="text-xs text-slate-600 mt-1">{sheet.desc}</p>
                    <div className={`mt-4 flex items-center text-[11px] font-bold text-slate-500 transition-colors ${sheet.hoverTextClass}`}>
                      열기
                    </div>
                  </button>
                );
              })}
            </div>
          )}

          {isUserUnlocked && (
            <div className="mt-6 bg-white p-4 rounded-xl shadow-sm border border-slate-200">
              <div className="flex items-center justify-between gap-2 mb-3">
                <h3 className="text-sm font-black text-slate-900">저장된 리포트 종합 (먼지/금속/수은/PAHs)</h3>
                <span className="text-xs font-bold text-slate-600">총 {profileCombinedReports.length}건</span>
              </div>

              <div className="grid grid-cols-2 md:grid-cols-4 gap-2 mb-3">
                {SHEET_MENU.map((sheet) => (
                  <div key={sheet.id} className="p-2 rounded-lg border border-slate-200 bg-slate-50 text-center">
                    <p className="text-[11px] font-bold text-slate-600">{sheet.title}</p>
                    <p className="text-lg font-black text-slate-900">{profileSheetCounts[sheet.id] || 0}</p>
                  </div>
                ))}
              </div>

              {profileCombinedReports.length === 0 ? (
                <div className="p-5 text-center text-xs text-slate-500 border border-dashed border-slate-300 rounded-lg">
                  저장된 통합 리포트가 없습니다.
                </div>
              ) : (
                <div className="overflow-x-auto rounded-lg border border-slate-200">
                  <table className="w-full text-xs min-w-[1080px]">
                    <thead className="bg-slate-100 text-slate-700">
                      <tr>
                        <th className="p-2 font-bold text-center">구분</th>
                        <th className="p-2 font-bold text-center">측정일자</th>
                        <th className="p-2 font-bold text-center">사업장</th>
                        <th className="p-2 font-bold text-center">배출구</th>
                        <th className="p-2 font-bold text-center">등속흡인율(%)</th>
                        <th className="p-2 font-bold text-center">실측농도</th>
                        <th className="p-2 font-bold text-center">보정농도</th>
                        <th className="p-2 font-bold text-center">불러오기</th>
                        <th className="p-2 font-bold text-center">삭제</th>
                      </tr>
                    </thead>
                    <tbody className="divide-y divide-slate-100">
                      {profileCombinedReports.map((data) => (
                        <tr key={data.id || data.savedAt} className="hover:bg-slate-50 cursor-pointer" onClick={() => handleLoadSavedReport(data)}>
                          <td className="p-2 text-center font-bold text-slate-800">{data._sheetTitle}</td>
                          <td className="p-2 text-center whitespace-nowrap">{data.date || '-'}</td>
                          <td className="p-2 text-center">{data.company || '-'}</td>
                          <td className="p-2 text-center">{data.location || '-'}</td>
                          <td className="p-2 text-center">{data.isokineticRate || '-'}</td>
                          <td className="p-2 text-center">{data.actualConcentration || '-'}</td>
                          <td className="p-2 text-center">{data.correctedConcentration || '-'}</td>
                          <td className="p-2 text-center">
                            <button
                              type="button"
                              onClick={(e) => {
                                e.stopPropagation();
                                handleLoadSavedReport(data);
                              }}
                              className="px-2 py-1 rounded border border-emerald-300 text-emerald-700 hover:bg-emerald-50 font-bold"
                            >
                              불러오기
                            </button>
                          </td>
                          <td className="p-2 text-center">
                            <button
                              type="button"
                              onClick={(e) => {
                                e.stopPropagation();
                                removeSavedReport(data.id, data.savedAt);
                              }}
                              className="px-2 py-1 rounded border border-red-300 text-red-700 hover:bg-red-50 font-bold"
                            >
                              삭제
                            </button>
                          </td>
                        </tr>
                      ))}
                    </tbody>
                  </table>
                </div>
              )}
            </div>
          )}
        </div>
        <SignatureBadge />
      </div>
    );
  }

  return (
    <div className={`theme-${selectedSheet} relative min-h-screen p-4 md:p-8 font-sans text-slate-800 ${activeTheme.selection} ${isNightSky ? 'bg-gradient-to-b from-[#050914] via-[#264284] to-[#7a5a8c]' : 'bg-gradient-to-b from-[#99ceff] via-[#8ab8ef] to-[#f1b7c8]'}`}>
      <style>{SKYLINE_FLYER_CSS}</style>
      <div className="fixed left-4 bottom-4 z-40 flex items-center gap-1 rounded-xl border border-white/70 bg-white/80 px-2 py-1 backdrop-blur">
        {SKY_PREVIEW_OPTIONS.map((opt) => (
          <button
            key={`sheet-sky-${opt.id}`}
            type="button"
            onClick={() => setSkyPreviewMode(opt.id)}
            className={`px-2 py-1 rounded text-[11px] font-black transition-colors ${skyPreviewMode === opt.id ? 'bg-slate-800 text-white' : 'text-slate-700 hover:bg-slate-100'}`}
          >
            {opt.label}
          </button>
        ))}
        <button
          type="button"
          onClick={() => setLowSpecMode((prev) => !prev)}
          className={`px-2 py-1 rounded text-[11px] font-black transition-colors ${lowSpecMode ? 'bg-emerald-600 text-white' : 'text-slate-700 hover:bg-slate-100'}`}
        >
          경량
        </button>
      </div>
      {selectedSheet !== 'dust' && <style>{THEME_OVERRIDE_CSS}</style>}
      {activeTheme.pageTint && <div className={`pointer-events-none fixed inset-0 z-0 ${activeTheme.pageTint}`} />}
      <div className="pointer-events-none fixed inset-0 z-0 overflow-hidden opacity-[0.4]">
        <div className={`absolute inset-0 ${isNightSky ? 'bg-gradient-to-b from-[#040915]/42 via-[#1d2f64]/16 to-transparent' : 'bg-gradient-to-b from-white/28 via-sky-100/12 to-transparent'}`} />
        <div className={isNightSky ? 'skyline-moon' : 'skyline-sun'} />
        {isNightSky && sceneStarsSheet.map((star) => (
          <span
            key={`sheet-${star.id}`}
            className="skyline-star"
            style={{
              left: star.left,
              top: star.top,
              width: `${star.size}px`,
              height: `${star.size}px`,
              '--twinkle-duration': `${star.duration}s`,
              '--twinkle-delay': `-${star.delay}s`,
            }}
          />
        ))}
        <div className="absolute bottom-0 left-0 right-0">
          <div className="relative mx-auto h-44 md:h-56 max-w-7xl">
            <div className={`absolute bottom-0 left-0 right-0 h-20 ${isNightSky ? 'bg-gradient-to-b from-[#223733] to-[#070d14]' : 'bg-gradient-to-b from-[#6b8f70] to-[#395447]'}`} />
            {sceneMountains.map((mountain) => (
              <div
                key={`sheet-${mountain.id}`}
                className={`absolute bottom-16 rounded-t-[45%] ${isNightSky ? 'bg-gradient-to-b from-[#35446c]/72 to-[#1a243f]/68' : 'bg-gradient-to-b from-[#90a9c9]/70 to-[#6784aa]/66'}`}
                style={{ left: mountain.left, width: mountain.width, height: `${Math.round(mountain.height * 0.38)}px` }}
              />
            ))}
            {sceneTreeBelt.map((tree, treeIdx) => (
              <div
                key={`sheet-tree-${tree.id}`}
                className="absolute bottom-[3.9rem]"
                style={{
                  left: tree.left,
                  width: `${Math.max(9, Math.round(tree.width * 0.86))}px`,
                  height: `${Math.max(12, Math.round(tree.height * 0.86))}px`,
                }}
              >
                {(() => {
                  const pixel = Math.max(2, Math.round(tree.width / 7));
                  const variant = PIXEL_TREE_VARIANTS[(treeIdx * 3 + tree.width) % PIXEL_TREE_VARIANTS.length];
                  const palette = getPixelTreePalette(isNightSky, treeIdx + (tree.height % 5));
                  return (
                    <>
                      {variant.cells.map((cell, idx) => (
                        <span
                          key={`sheet-${tree.id}-cell-${idx}`}
                          className="absolute"
                          style={{
                            left: `${cell[0] * pixel}px`,
                            bottom: `${(cell[1] + 2) * pixel}px`,
                            width: `${pixel}px`,
                            height: `${pixel}px`,
                            background: idx % 4 === 0 ? palette.alt : palette.main,
                          }}
                        />
                      ))}
                      <span
                        className="absolute"
                        style={{
                          left: `${variant.trunk.x * pixel}px`,
                          bottom: `${pixel}px`,
                          width: `${pixel}px`,
                          height: `${variant.trunk.h * pixel}px`,
                          background: palette.trunk,
                        }}
                      />
                      <span
                        className="absolute"
                        style={{
                          left: `${variant.trunk.baseX * pixel}px`,
                          bottom: '0px',
                          width: `${variant.trunk.baseW * pixel}px`,
                          height: `${pixel}px`,
                          background: palette.trunk,
                        }}
                      />
                    </>
                  );
                })()}
              </div>
            ))}
            {sceneBackTowers.map((tower) => (
              <div
                key={`sheet-${tower.id}`}
                className={`absolute bottom-14 rounded-t-[2px] border ${isNightSky ? 'border-[#2b3558]/70 bg-gradient-to-b from-[#6c7cab] via-[#495a8a] to-[#243052]' : 'border-[#51667e]/55 bg-gradient-to-b from-[#9cb4d1] via-[#7f9cbf] to-[#5f7d9f]'}`}
                style={{ left: tower.left, width: tower.width, height: `${Math.round(tower.height * 0.48)}px` }}
              />
            ))}
            {sceneFactoryBlocks.map((block) => (
              <div
                key={`sheet-${block.id}`}
                className={`absolute bottom-12 rounded-t-[3px] border ${isNightSky ? 'border-[#2f2d45]/80 bg-gradient-to-b from-[#8f7668] via-[#6e5850] to-[#3d3033]' : 'border-[#5a5b72]/65 bg-gradient-to-b from-[#b89c89] via-[#967b6f] to-[#6c5550]'}`}
                style={{ left: block.left, width: block.width, height: `${Math.round(block.height * 0.52)}px` }}
              />
            ))}
            {sceneStacks.map((stack) => (
              <div
                key={`sheet-${stack.id}`}
                className={`absolute bottom-12 rounded-t-[3px] border ${isNightSky ? 'border-[#2a3148]/80 bg-gradient-to-b from-[#a4afc0] via-[#7f8ba0] to-[#4d596e]' : 'border-[#51667a]/65 bg-gradient-to-b from-[#bfd0e0] via-[#98adc4] to-[#6f89a3]'}`}
                style={{ left: stack.left, width: `${Math.round(stack.width * 0.82)}px`, height: `${Math.round(stack.height * 0.52)}px` }}
              />
            ))}
            {sceneFrontTreeBelt.map((tree, treeIdx) => (
              <div
                key={`sheet-front-${tree.id}`}
                className="absolute bottom-[3rem]"
                style={{
                  left: tree.left,
                  width: `${Math.max(10, Math.round(tree.width * 0.92))}px`,
                  height: `${Math.max(14, Math.round(tree.height * 0.88))}px`,
                }}
              >
                {(() => {
                  const pixel = Math.max(2, Math.round(tree.width / 8));
                  const variant = PIXEL_TREE_VARIANTS[(treeIdx * 5 + tree.height) % PIXEL_TREE_VARIANTS.length];
                  const palette = getPixelTreePalette(isNightSky, treeIdx + 2 + (tree.width % 5));
                  return (
                    <>
                      {variant.cells.map((cell, idx) => (
                        <span
                          key={`sheet-front-${tree.id}-cell-${idx}`}
                          className="absolute"
                          style={{
                            left: `${cell[0] * pixel}px`,
                            bottom: `${(cell[1] + 2) * pixel}px`,
                            width: `${pixel}px`,
                            height: `${pixel}px`,
                            background: idx % 4 === 0 ? palette.alt : palette.main,
                          }}
                        />
                      ))}
                      <span
                        className="absolute"
                        style={{
                          left: `${variant.trunk.x * pixel}px`,
                          bottom: `${pixel}px`,
                          width: `${pixel}px`,
                          height: `${variant.trunk.h * pixel}px`,
                          background: palette.trunk,
                        }}
                      />
                      <span
                        className="absolute"
                        style={{
                          left: `${variant.trunk.baseX * pixel}px`,
                          bottom: '0px',
                          width: `${variant.trunk.baseW * pixel}px`,
                          height: `${pixel}px`,
                          background: palette.trunk,
                        }}
                      />
                    </>
                  );
                })()}
              </div>
            ))}
          </div>
        </div>
      </div>
      <div className="max-w-7xl mx-auto space-y-6 relative z-10">
        
        {/* 헤더 섹션 */}
        <div className={`bg-white p-6 rounded-xl shadow-md border-t-4 relative ${activeTheme.headerBorder}`}>
          <div className="flex flex-col items-center text-center">
            <h1 className="text-2xl font-black text-slate-900">
              <img
                src={mainLogoSrc}
                alt="STACKPILOT"
                onError={() => setLogoIndex((prev) => Math.min(prev + 1, logoCandidates.length - 1))}
                className="h-[2.2rem] md:h-[3.2rem] w-auto select-none"
              />
            </h1>
            <p className="text-slate-600 mt-2 text-sm font-medium">
              <span className="font-black text-slate-800">{activeSheet?.title || '먼지시료채취기록부'}</span>
              {' '}| 대기오염공정시험기준에 맞춘 전 항목 종합 연산 및 자동 기록 시트
            </p>
          </div>
          <button
            type="button"
            onClick={navigateToMenu}
            className="mt-3 md:mt-0 md:absolute md:right-4 md:top-4 px-3 py-2 text-xs font-bold rounded-lg border border-slate-300 bg-white hover:bg-slate-50"
          >
            기록부 선택
          </button>
        </div>

        <div className="bg-white p-3 rounded-xl shadow-sm border border-slate-200 flex flex-col md:flex-row md:items-center md:justify-between gap-2">
          <p className="text-xs text-slate-600">
            사용자: <span className="font-black text-slate-900">{activeUser || '미선택'}</span>
            {' '}| 상태: <span className={`font-black ${isUserUnlocked ? activeTheme.accentText : 'text-red-600'}`}>{isUserUnlocked ? '로그인됨' : '로그인 필요'}</span>
            {' '}| 계정/로그아웃은 상단 초기 화면에서 설정합니다.
          </p>
          {isUserUnlocked && (
            <button
              type="button"
              onClick={handleLockActiveUser}
              className="px-3 py-1.5 bg-slate-700 text-white rounded-lg font-bold text-xs hover:bg-slate-800 self-start md:self-auto"
            >
              로그아웃
            </button>
          )}
        </div>

        <form onSubmit={handleSave} onKeyDown={handleFormKeyDown} className="space-y-6">
          
          {/* 섹션 1: 일반 사항 및 배출구 제원 */}
          <div className="bg-white p-6 rounded-xl shadow-sm border border-slate-200">
            <h2 className="text-lg font-bold text-slate-800 mb-4 border-b border-slate-200 pb-2 flex items-center gap-2">
              <span className="bg-slate-800 text-white font-black w-6 h-6 rounded-full flex items-center justify-center text-sm">1</span>
              일반 사항 및 배출구 제원
            </h2>
            <div className="grid grid-cols-1 md:grid-cols-4 gap-4 mb-4">
              <div><label className="block text-xs font-bold text-slate-600 mb-1">측정 일자</label>
                <input type="date" name="date" value={formData.date} onChange={handleChange} className="w-full p-2 bg-emerald-50/50 border border-emerald-200 rounded focus:ring-2 focus:ring-emerald-500 outline-none transition-colors" /></div>
              <div><label className="block text-xs font-bold text-slate-600 mb-1">사업장명</label>
                <input type="text" name="company" value={formData.company} onChange={handleChange} className="w-full p-2 bg-emerald-50/50 border border-emerald-200 rounded focus:ring-2 focus:ring-emerald-500 outline-none transition-colors" /></div>
              <div><label className="block text-xs font-bold text-slate-600 mb-1">배출구명 (위치)</label>
                <input type="text" name="location" value={formData.location} onChange={handleChange} className="w-full p-2 bg-emerald-50/50 border border-emerald-200 rounded focus:ring-2 focus:ring-emerald-500 outline-none transition-colors" /></div>
              <div><label className="block text-xs font-bold text-slate-600 mb-1">측정자</label>
                <input type="text" name="sampler" value={formData.sampler} onChange={handleChange} className="w-full p-2 bg-emerald-50/50 border border-emerald-200 rounded focus:ring-2 focus:ring-emerald-500 outline-none transition-colors" /></div>
            </div>
            <div className="grid grid-cols-1 md:grid-cols-4 gap-4 bg-emerald-50/30 p-4 rounded-lg border border-emerald-100">
              <div><label className="block text-xs font-bold text-slate-600 mb-1">날씨</label>
                <input type="text" name="weather" value={formData.weather} onChange={handleChange} className="w-full p-2 border border-slate-300 rounded focus:ring-2 focus:ring-emerald-500 outline-none" /></div>
              <div><label className="block text-xs font-bold text-slate-600 mb-1">대기 온도 (℃)</label>
                <input type="number" step="0.1" name="atmTemp" value={formData.atmTemp} onChange={handleChange} className="w-full p-2 border border-slate-300 rounded focus:ring-2 focus:ring-emerald-500 outline-none" /></div>
              <div className="md:col-span-2"><label className="block text-xs font-bold text-slate-600 mb-1">대기압 (측정공 주변, mmHg)</label>
                <input type="number" step="1" name="atmPressure" value={formData.atmPressure} onChange={handleChange} className="w-full p-2 border border-slate-300 rounded focus:ring-2 focus:ring-emerald-500 outline-none" /></div>
            </div>

            {/* 굴뚝 제원 (직경 산출) */}
            <div className="mt-4 grid grid-cols-1 md:grid-cols-4 gap-4 bg-emerald-50 p-4 rounded-lg border border-emerald-200">
              <div>
                <label className="block text-xs font-bold text-emerald-800 mb-1">총 측정 깊이 (m)</label>
                <input type="number" step="0.001" name="totalStackDepth" value={formData.totalStackDepth} onChange={handleChange} placeholder="끝벽까지의 총 길이" className="w-full p-2 border border-emerald-300 rounded focus:ring-2 focus:ring-emerald-500 bg-white outline-none" />
              </div>
              <div>
                <label className="block text-xs font-bold text-emerald-800 mb-1">플랜지 길이 (m)</label>
                <input type="number" step="0.001" name="flangeLength" value={formData.flangeLength} onChange={handleChange} placeholder="외벽 돌출부" className="w-full p-2 border border-emerald-300 rounded focus:ring-2 focus:ring-emerald-500 bg-white outline-none" />
              </div>
              <div>
                <label className="block text-xs font-bold text-emerald-900 mb-1">순수 굴뚝 내경 (D, m)</label>
                <input type="number" step="0.001" name="stackDiameter" value={formData.stackDiameter} onChange={handleChange} className="w-full p-2 border border-emerald-400 rounded focus:ring-2 focus:ring-emerald-500 bg-emerald-100 font-black text-emerald-900 outline-none" placeholder="자동계산 (총 깊이 - 플랜지)" />
              </div>
              <div>
                <label className="block text-xs font-bold text-emerald-900 mb-1">단면적 (A, ㎡)</label>
                <input type="text" readOnly value={samplingPointsData ? samplingPointsData.area : ''} className="w-full p-2 border border-emerald-400 rounded bg-emerald-100 font-black text-emerald-900 outline-none" placeholder="자동계산" />
              </div>
            </div>

            {/* 굴뚝 내경에 따른 측정점 산출 패널 */}
            {samplingPointsData && (
              <div className="mt-4 bg-white p-4 rounded-lg border border-slate-200 shadow-sm">
                <div className="flex justify-between items-start mb-3">
                  <h3 className="text-sm font-bold text-slate-800">
                    굴뚝 내경({formData.stackDiameter}m) 연동 측정점 산출
                  </h3>
                  {samplingPointsData.isCenterOnly && (
                    <span className="bg-red-50 text-red-600 px-2 py-1 rounded text-xs font-bold animate-pulse border border-red-200">
                      단면적 0.25㎡ 이하 (단면중심 1점 측정)
                    </span>
                  )}
                </div>
                
                <div className="flex flex-col xl:flex-row gap-4 items-start">
                  <div className="flex flex-col gap-2 shrink-0">
                    <div className="bg-slate-100 px-3 py-2 rounded border border-slate-200 text-center shadow-sm">
                      <span className="block text-[10px] text-slate-600 font-bold">1개 측정공(Port) 기준</span>
                      <span className="font-black text-emerald-700 text-lg">총 {samplingPointsData.perRadius}개</span>
                    </div>
                  </div>
                  
                  <div className="flex-1 w-full overflow-x-auto bg-white rounded border border-slate-200">
                    <table className="w-full text-xs text-center border-collapse">
                      <thead>
                        <tr className="bg-slate-50 text-slate-700">
                          <th className="border-r border-b border-slate-200 p-2 whitespace-nowrap" rowSpan="2">삽입 깊이 (마킹 위치)</th>
                          <th className="border-b border-slate-200 p-2 whitespace-nowrap" colSpan={samplingPointsData.perRadius}>측정 포인트 (1~5)</th>
                        </tr>
                        <tr className="bg-slate-50 text-slate-700">
                          {samplingPointsData.nearInsertion.map((_, i) => (
                            <th key={i} className="border-b border-l border-slate-200 p-1">{i + 1}번</th>
                          ))}
                        </tr>
                      </thead>
                      <tbody>
                        <tr>
                          <td className="border-r border-b border-slate-200 p-2 font-bold text-slate-800 whitespace-nowrap bg-slate-50/50">
                            가까운 쪽 (정방향)<br/><span className="text-[9px] font-normal text-slate-600">기본 측정</span>
                          </td>
                          {samplingPointsData.nearInsertion.map((d, i) => (
                            <td key={i} className="border-l border-b border-slate-200 p-2 font-black text-emerald-700 bg-white">
                              {d} m
                            </td>
                          ))}
                        </tr>
                        {!samplingPointsData.isCenterOnly && (
                          <tr>
                            <td className="border-r border-slate-200 p-2 font-bold text-red-700 whitespace-nowrap bg-red-50">
                              깊은 쪽 (역방향)<br/><span className="text-[9px] font-normal text-red-500">반대편 측정공 대체시</span>
                            </td>
                            {samplingPointsData.farInsertion.map((d, i) => (
                              <td key={i} className="border-l border-slate-200 p-2 font-black text-red-600 bg-red-50/50">
                                {d} m
                              </td>
                            ))}
                          </tr>
                        )}
                      </tbody>
                    </table>
                  </div>
                </div>
                
                <div className="mt-4 flex flex-col md:flex-row justify-between items-center gap-3 border-t border-slate-100 pt-4">
                  <div className="text-[10px] text-slate-500 text-left w-full space-y-1">
                    <p>※ 반대편 측정공이 없어 깊게 찔러야 할 경우에도 <strong>기록표의 칸수는 변하지 않으며</strong>, 위 표의 <strong>[깊은 쪽] 삽입 깊이 수치</strong>만 보시고 1~5번에 맞춰 찌르시면 됩니다.</p>
                  </div>
                  <button type="button" onClick={applySamplingPointsToTable} className="w-full md:w-auto py-2 px-5 bg-emerald-600 text-white rounded-lg font-bold text-sm hover:bg-emerald-700 shadow-sm transition-colors shrink-0 border border-emerald-700">
                    하단 표 {samplingPointsData.perRadius}칸 자동 생성
                  </button>
                </div>
              </div>
            )}
          </div>

          {/* 섹션 2: 배출가스 조성 및 분자량 (가스분석기 자료) */}
          <div className="bg-white p-6 rounded-xl shadow-sm border border-slate-200">
            <div className="flex flex-col md:flex-row md:items-center justify-between mb-4 border-b border-slate-200 pb-2">
              <h2 className="text-lg font-bold text-slate-800 flex items-center gap-2">
                <span className="bg-teal-600 text-white w-6 h-6 rounded-full flex items-center justify-center text-sm font-black">2</span>
                배출가스 조성 및 밀도 (가스분석기 5분 간격 3회 측정)
              </h2>
            </div>
            
            <div className="grid grid-cols-1 lg:grid-cols-3 gap-6">
               <div className="lg:col-span-2 overflow-x-auto">
                  <table className="w-full text-sm text-center border-collapse">
                    <thead>
                      <tr className="bg-teal-50 text-teal-800 border-b border-teal-200">
                        <th className="p-2 border-r border-teal-100 font-bold whitespace-nowrap">회차 (시간)</th>
                        <th className="p-2 border-r border-teal-100 font-bold">O₂ (%)</th>
                        <th className="p-2 border-r border-teal-100 font-bold">CO₂ (%)</th>
                        <th className="p-2 border-r border-teal-100 font-bold">CO (ppm)</th>
                        {/* NOx, SOx 위치 변경 적용됨 */}
                        <th className="p-2 border-r border-teal-100 font-bold">NOx (ppm)</th>
                        <th className="p-2 font-bold">SOx (ppm)</th>
                      </tr>
                    </thead>
                    <tbody>
                      {formData.gasAnalyzer.map((gas, idx) => (
                        <tr key={idx} className="hover:bg-slate-50 border-b border-slate-100">
                          <td className="p-2 border-r border-slate-100 font-medium text-slate-700 flex justify-center items-center gap-2">
                            <span className="whitespace-nowrap">{idx + 1}회차</span>
                            <input type="time" value={gas.time || ''} onChange={(e) => handleGasAnalyzerChange(idx, 'time', e.target.value)} className="w-24 p-1 border border-slate-300 rounded text-center text-xs outline-none focus:ring-2 focus:ring-teal-500" />
                          </td>
                          <td className="p-1 border-r border-slate-100">
                            <input type="number" step="0.01" value={gas.o2} onChange={(e) => handleGasAnalyzerChange(idx, 'o2', e.target.value)} className="w-full p-1 border border-slate-300 rounded text-center outline-none focus:ring-2 focus:ring-teal-500" />
                          </td>
                          <td className="p-1 border-r border-slate-100">
                            <input type="number" step="0.01" value={gas.co2} onChange={(e) => handleGasAnalyzerChange(idx, 'co2', e.target.value)} className="w-full p-1 border border-slate-300 rounded text-center outline-none focus:ring-2 focus:ring-teal-500" />
                          </td>
                          <td className="p-1 border-r border-slate-100">
                            <input type="number" step="0.01" value={gas.co} onChange={(e) => handleGasAnalyzerChange(idx, 'co', e.target.value)} className="w-full p-1 border border-slate-300 rounded text-center outline-none focus:ring-2 focus:ring-teal-500" />
                          </td>
                          <td className="p-1 border-r border-slate-100">
                            <input type="number" step="0.01" value={gas.nox} onChange={(e) => handleGasAnalyzerChange(idx, 'nox', e.target.value)} className="w-full p-1 border border-slate-300 rounded text-center outline-none focus:ring-2 focus:ring-teal-500" />
                          </td>
                          <td className="p-1">
                            <input type="number" step="0.01" value={gas.sox} onChange={(e) => handleGasAnalyzerChange(idx, 'sox', e.target.value)} className="w-full p-1 border border-slate-300 rounded text-center outline-none focus:ring-2 focus:ring-teal-500" />
                          </td>
                        </tr>
                      ))}
                      <tr className="bg-slate-50 font-bold text-slate-800">
                        <td className="p-2 border-r border-slate-200">평균</td>
                        <td className="p-2 border-r border-slate-200 text-teal-600">{o2.toFixed(2)}</td>
                        <td className="p-2 border-r border-slate-200">{co2.toFixed(2)}</td>
                        <td className="p-2 border-r border-slate-200">{co.toFixed(2)}</td>
                        <td className="p-2 border-r border-slate-200">{nox.toFixed(2)}</td>
                        <td className="p-2">{sox.toFixed(2)}</td>
                      </tr>
                    </tbody>
                  </table>
               </div>
               
               <div className="lg:col-span-1 flex flex-col gap-4">
                 {/* O2 보정 계수 패널 */}
                 <div className="bg-white p-4 rounded-lg border border-teal-200 shadow-sm relative overflow-hidden">
                    <div className="absolute top-0 left-0 w-1.5 h-full bg-teal-500"></div>
                    <h3 className="text-xs font-bold text-teal-800 mb-3 border-b border-teal-100 pb-2">
                      산소(O₂) 보정 계수 산출
                    </h3>
                    <div className="flex justify-between items-center mb-2">
                       <div className="flex flex-col">
                         <span className="text-xs font-bold text-slate-700">표준 산소 농도 (O<sub>s</sub>, %)</span>
                         <span className="text-[9px] text-slate-500">보정 대상이 아니면 비워두세요</span>
                       </div>
                       <input type="number" step="0.1" name="standardO2" value={formData.standardO2} onChange={handleChange} className="w-20 p-1 border border-teal-300 rounded text-center text-sm font-black text-teal-700 bg-teal-50 focus:ring-2 focus:ring-teal-400 outline-none" placeholder="공란" />
                    </div>
                    <div className="flex justify-between items-center mb-2">
                       <span className="text-xs font-bold text-slate-700">실측 산소 농도 (O<sub>a</sub>, %)</span>
                       <span className="font-bold text-slate-800">{o2.toFixed(2)} %</span>
                    </div>
                    <div className="flex justify-between items-center pt-2 mt-1 border-t border-teal-100">
                       <span className="text-xs font-bold text-teal-700" title="(21 - 표준O2) / (21 - 실측O2)">산소 보정 계수 (K)</span>
                       <span className="font-black text-lg text-teal-600">{calcO2CorrectionFactor().toFixed(3)}</span>
                    </div>
                 </div>

                 {/* 분자량/밀도 패널 */}
                 <div className="bg-slate-50 p-4 rounded-lg border border-slate-200">
                    <div className="flex justify-between items-center mb-1">
                       <span className="text-[11px] font-bold text-slate-600" title="100 - O₂(%) - CO₂(%) - CO(%)">평균 N₂ (%)</span>
                       <span className="text-xs font-bold text-slate-700">{n2.toFixed(2)}</span>
                    </div>
                    <div className="flex justify-between items-center mb-1">
                       <span className="text-[11px] font-bold text-slate-600" title="Σ(M_i * x_i) / 100">건조가스분자량 (M<sub>d</sub>)</span>
                       <span className="text-xs font-black text-slate-700">{Md.toFixed(3)}</span>
                    </div>
                    <div className="flex justify-between items-center mb-1 pt-1 border-t border-slate-200/50">
                       <span className="text-[11px] font-bold text-slate-600" title="사전 수분량 반영됨">습가스분자량 (M<sub>s</sub>)</span>
                       <span className="text-sm font-black text-slate-700">{Ms.toFixed(3)}</span>
                    </div>
                    <div className="flex justify-between items-center pt-1 mt-1 border-t border-slate-200/50">
                       <span className="text-[11px] font-bold text-slate-600" title="1 / (22.4 * 100) * [ Σ(M_i*x_i) * (100 - Xw)/100 + 18 * Xw ]">배출가스밀도 (γ<sub>a</sub>)</span>
                       <span className="text-sm font-black text-teal-600">{r0.toFixed(3)}</span>
                    </div>
                    <div className="flex justify-between items-center pt-1 mt-1 border-t border-slate-200/50">
                       <span className="text-[11px] font-bold text-slate-600" title="온도 및 압력 보정된 실제 밀도">배출가스밀도 (r)</span>
                       <span className="text-sm font-black text-teal-600">{isNaN(r) ? '-' : r.toFixed(3)}</span>
                    </div>
                 </div>
               </div>
            </div>
          </div>

          <div className="grid grid-cols-1 lg:grid-cols-2 gap-6">
            {/* 섹션 3: 배출가스 수분량 */}
            <div className="bg-white p-6 rounded-xl shadow-sm border border-slate-200">
              <h2 className="text-lg font-bold text-slate-800 mb-4 border-b border-slate-200 pb-2 flex items-center gap-2">
                <span className="bg-cyan-600 text-white w-6 h-6 rounded-full flex items-center justify-center text-sm font-black">3</span>
                배출가스 수분량 (자동측정기법)
              </h2>
              <div className="space-y-4">
                <div>
                  <label className="block text-xs font-bold text-slate-600 mb-2">사전 수분량 (%) - 5회 측정값 입력</label>
                  <div className="grid grid-cols-5 gap-2">
                    {formData.moistureValues.map((val, idx) => (
                      <div key={idx}>
                        <label className="block text-[10px] text-slate-500 text-center mb-1">{idx + 1}회차</label>
                        <input 
                          type="number" step="0.1" value={val} 
                          onChange={(e) => handleMoistureChange(idx, e.target.value)} 
                          className="w-full p-2 border border-slate-300 rounded focus:ring-2 focus:ring-cyan-500 bg-slate-50 text-center text-sm outline-none" 
                        />
                      </div>
                    ))}
                  </div>
                </div>
                
                <div className="bg-cyan-50 p-4 rounded-lg flex justify-between items-center border border-cyan-100 mt-4 shadow-sm">
                  <div className="text-slate-800">
                    <span className="font-bold">사전 적용 수분량 (평균)</span>
                  </div>
                  <div className="text-2xl font-black text-cyan-700">{calcMoisture()}%</div>
                </div>
                
                <div className="mt-6 pt-4 border-t border-slate-200">
                  <label className="block text-sm font-bold text-slate-800 mb-2 flex items-center gap-1">
                    사후 수분량 (임핀저 무게법) <span className="text-slate-500 font-normal text-xs">- 최종 등속흡인계수 계산용</span>
                  </label>
                  <div className="overflow-x-auto">
                    <table className="w-full text-sm text-center border-collapse">
                      <thead>
                        <tr className="bg-slate-100 text-slate-700">
                          <th className="p-2 border border-slate-200 font-bold rounded-tl-lg">임핀저</th>
                          <th className="p-2 border border-slate-200 font-bold">채취 전 (g)</th>
                          <th className="p-2 border border-slate-200 font-bold">채취 후 (g)</th>
                          <th className="p-2 border border-slate-200 font-bold rounded-tr-lg">증가량 (g)</th>
                        </tr>
                      </thead>
                      <tbody>
                        {formData.impingers.map((imp, idx) => (
                          <tr key={imp.id} className="hover:bg-slate-50">
                            <td className="p-2 border border-slate-200 text-slate-700 font-bold">{imp.id}단{imp.id === 4 ? '(실리카겔)' : ''}</td>
                            <td className="p-1 border border-slate-200">
                              <input type="number" step="0.01" value={imp.initial} onChange={(e) => handleImpingerChange(idx, 'initial', e.target.value)} className="w-full p-1 border border-slate-300 rounded text-center outline-none focus:ring-2 focus:ring-emerald-500" />
                            </td>
                            <td className="p-1 border border-slate-200">
                              <input type="number" step="0.01" value={imp.final} onChange={(e) => handleImpingerChange(idx, 'final', e.target.value)} className="w-full p-1 border border-slate-300 rounded text-center outline-none focus:ring-2 focus:ring-emerald-500" />
                            </td>
                            <td className="p-2 border border-slate-200 font-black text-slate-800 bg-slate-50/50">
                              {!isNaN(parseFloat(imp.initial)) && !isNaN(parseFloat(imp.final)) ? (parseFloat(imp.final) - parseFloat(imp.initial)).toFixed(2) : '-'}
                            </td>
                          </tr>
                        ))}
                      </tbody>
                    </table>
                  </div>
                  <div className="flex flex-col gap-2 mt-3">
                    <div className="flex justify-between items-center bg-slate-50 p-3 rounded-lg border border-slate-200">
                      <span className="text-sm font-bold text-slate-700">포집된 수분 총량 (V<sub>lc</sub>)</span>
                      <span className="font-black text-slate-800">{calcTotalMoistureWeight().toFixed(2)} ml</span>
                    </div>
                    <div className="flex justify-between items-center bg-emerald-600 p-3 rounded-lg shadow-sm text-white">
                      <span className="text-sm font-bold">사후 수분량 (X<sub>w</sub>)</span>
                      <span className="font-black text-xl text-emerald-100">{calcPostMoisture()}%</span>
                    </div>
                  </div>
                </div>
              </div>
            </div>

            {/* ✅ 4번 섹션 동정압 측정 시의 미터온도(Tm) 입력칸 추가 */}
            <div className="bg-white p-6 rounded-xl shadow-sm border border-slate-200 flex flex-col h-full">
              <div className="flex flex-col md:flex-row md:justify-between md:items-center mb-4 border-b border-slate-200 pb-3 gap-3">
                <h2 className="text-lg font-bold text-slate-800 flex items-center gap-2">
                  <span className="bg-emerald-500 text-white w-6 h-6 rounded-full flex items-center justify-center text-sm font-black">4</span>
                  측정점별 동압 및 온도 (예비조사)
                </h2>
                <div className="flex flex-wrap items-center gap-2 md:gap-4">
                  <div className="flex items-center gap-1.5 bg-slate-50 px-2 py-1.5 rounded-lg border border-slate-200">
                    <label className="text-[11px] font-bold text-slate-700">피토계수</label>
                    <input type="number" step="0.01" name="pitotFactor" value={formData.pitotFactor} onChange={handleChange} className="w-14 p-1 border border-slate-300 rounded text-xs text-center outline-none focus:ring-2 focus:ring-emerald-500" />
                  </div>
                  <div className="flex items-center gap-1.5 bg-emerald-50 px-2 py-1.5 rounded-lg border border-emerald-200 shadow-sm">
                    <label className="text-[11px] font-bold text-emerald-800">미터온도(T<sub>m</sub>)</label>
                    <input type="number" step="0.1" name="traverseTmIn" value={formData.traverseTmIn} onChange={handleChange} placeholder="입구" className="w-14 p-1 border border-emerald-300 rounded text-xs text-center outline-none focus:ring-2 focus:ring-emerald-500" />
                    <span className="text-slate-400 font-black">/</span>
                    <input type="number" step="0.1" name="traverseTmOut" value={formData.traverseTmOut} onChange={handleChange} placeholder="출구" className="w-14 p-1 border border-emerald-300 rounded text-xs text-center outline-none focus:ring-2 focus:ring-emerald-500" />
                  </div>
                </div>
              </div>
              
              <div className="flex-1 overflow-auto bg-slate-50/50 p-2 rounded border border-slate-200 mb-4">
                <table className="w-full text-sm text-center">
                  <thead>
                    <tr className="text-slate-700 border-b border-slate-300 bg-slate-100/50">
                      <th className="pb-2 pt-2 rounded-tl-lg font-bold">측정점</th>
                      <th className="pb-2 pt-2 text-xs font-bold">전압<br/>(mmH2O)</th>
                      <th className="pb-2 pt-2 text-xs font-bold">정압<br/>(mmH2O)</th>
                      <th className="pb-2 pt-2 text-xs font-bold">동압<br/>(mmH2O)</th>
                      <th className="pb-2 pt-2 text-xs font-bold">온도(℃)</th>
                      <th className="pb-2 pt-2 rounded-tr-lg"></th>
                    </tr>
                  </thead>
                  <tbody>
                    {formData.points.map((point, idx) => (
                      <tr key={point.id} className="hover:bg-slate-100/30">
                        <td className="py-1 font-bold text-slate-700">{point.id}</td>
                        <td className="py-1"><input type="number" step="0.1" value={point.tp !== undefined ? point.tp : ''} onChange={(e) => handlePointChange(idx, 'tp', e.target.value)} className="w-16 p-1 border border-slate-300 rounded text-center text-sm outline-none focus:ring-2 focus:ring-emerald-500" /></td>
                        <td className="py-1"><input type="number" step="0.1" value={point.sp !== undefined ? point.sp : ''} onChange={(e) => handlePointChange(idx, 'sp', e.target.value)} className="w-16 p-1 border border-slate-300 rounded text-center text-sm outline-none focus:ring-2 focus:ring-emerald-500" /></td>
                        <td className="py-1"><input type="number" step="0.1" value={point.dp} onChange={(e) => handlePointChange(idx, 'dp', e.target.value)} className="w-16 p-1 border border-emerald-300 bg-emerald-50 rounded text-center font-black text-emerald-700 text-sm outline-none focus:ring-2 focus:ring-emerald-500" /></td>
                        <td className="py-1"><input type="number" step="0.1" value={point.ts} onChange={(e) => handlePointChange(idx, 'ts', e.target.value)} className="w-16 p-1 border border-slate-300 rounded text-center text-sm outline-none focus:ring-2 focus:ring-emerald-500" /></td>
                        <td className="py-1">
                          <button type="button" onClick={() => removePoint(idx)} className="text-red-500 hover:text-red-700 text-xs font-bold p-1 transition-colors">삭제</button>
                        </td>
                      </tr>
                    ))}
                    {/* 전압/정압/동압/온도 평균 행 표시 */}
                    <tr className="bg-emerald-50/50 border-t-2 border-emerald-200 text-emerald-900 font-bold">
                        <td className="py-1.5">평균</td>
                        <td className="py-1.5">{calcAvgTp() !== 0 ? calcAvgTp().toFixed(2) : '-'}</td>
                        <td className="py-1.5">{calcAvgSp() !== 0 ? calcAvgSp().toFixed(2) : '-'}</td>
                        <td className="py-1.5 text-emerald-700">{calcAvgDp() !== 0 ? calcAvgDp().toFixed(2) : '-'}</td>
                        <td className="py-1.5">{calcAvgTs() !== 0 ? calcAvgTs().toFixed(1) : '-'}</td>
                        <td className="py-1.5"></td>
                    </tr>
                  </tbody>
                </table>
                <button type="button" onClick={addPoint} className="w-full mt-2 py-1.5 text-emerald-700 bg-emerald-100 hover:bg-emerald-200 rounded text-xs font-bold border border-emerald-300 transition-colors shadow-sm">
                  측정점 추가
                </button>
              </div>

              <div className="bg-emerald-50 p-4 rounded-lg flex justify-between items-center border border-emerald-200 shadow-sm">
                <div className="text-slate-800">
                  <span className="font-bold">평균 유속 (Vs)</span>
                </div>
                <div className="text-2xl font-black text-emerald-700">{calcGasVelocity()} <span className="text-sm font-bold text-slate-700">m/s</span></div>
              </div>
            </div>
          </div>

          {/* 섹션 5: 시료채취 및 등속흡인 */}
          <div className="bg-white p-6 rounded-xl shadow-sm border border-slate-200">
            <h2 className="text-lg font-bold text-slate-800 mb-4 border-b border-slate-200 pb-2 flex items-center gap-2">
              <span className="bg-emerald-600 text-white w-6 h-6 rounded-full flex items-center justify-center text-sm font-black">5</span>
              시료 채취 기록 (적산유량계 및 등속흡인)
            </h2>

            {/* 1단계: 자동 산정 및 예측 결과 통합 패널 */}
            <div className="bg-emerald-50 border border-emerald-200 p-5 rounded-lg mb-6 shadow-sm">
               <div className="flex flex-col md:flex-row md:justify-between md:items-end mb-4 gap-2">
                 <h3 className="font-bold text-emerald-800">
                    1단계: 채취조건 정밀 산출 (기기 고유 계수 연동)
                 </h3>
                 <div className="flex items-center gap-2 bg-white px-3 py-1.5 rounded-lg border border-emerald-200 shadow-sm">
                    <label className="text-xs font-bold text-slate-700">장비 프리셋 선택</label>
                    <select 
                       value={formData.samplerId} 
                       onChange={handleSamplerChange}
                       className="p-1 border border-emerald-300 rounded text-sm font-bold text-emerald-800 bg-emerald-50 focus:ring-2 focus:ring-emerald-500 outline-none cursor-pointer"
                    >
                       <option value="">직접 입력</option>
                       {samplers.map(s => (
                          <option key={s.id} value={s.id}>{s.name}</option>
                       ))}
                    </select>
                 </div>
               </div>

               <details className="mb-3 bg-white border border-emerald-200 rounded-lg p-2">
                 <summary className="cursor-pointer text-[11px] font-bold text-emerald-800">
                   샘플러 프리셋 직접 편집 (Yd / ΔH@)
                 </summary>
                 <div className="mt-2 flex items-center gap-2">
                   <button type="button" onClick={addSamplerPreset} className="px-2 py-1 text-[11px] font-bold bg-emerald-100 text-emerald-800 rounded border border-emerald-300">행 추가</button>
                   <button type="button" onClick={resetSamplerDefaults} className="px-2 py-1 text-[11px] font-bold bg-slate-100 text-slate-700 rounded border border-slate-300">초기값 복원</button>
                 </div>
                 <div className="mt-2 overflow-x-auto">
                   <table className="w-full text-[11px] text-center border border-emerald-200">
                     <thead className="bg-emerald-100 text-slate-800">
                       <tr>
                         <th className="border border-emerald-200 p-1">샘플러명</th>
                         <th className="border border-emerald-200 p-1">Yd</th>
                         <th className="border border-emerald-200 p-1">ΔH@</th>
                         <th className="border border-emerald-200 p-1">삭제</th>
                       </tr>
                     </thead>
                     <tbody>
                       {samplers.map((s, idx) => (
                         <tr key={s.id} className="bg-white">
                           <td className="border border-emerald-200 p-1">
                             <input
                               type="text"
                               value={s.name}
                               onChange={(e) => handleSamplerPresetChange(idx, 'name', e.target.value)}
                               className="w-full p-1 border border-emerald-300 rounded text-center"
                             />
                           </td>
                           <td className="border border-emerald-200 p-1">
                             <input
                               type="number"
                               step="0.001"
                               value={s.yd}
                               onChange={(e) => handleSamplerPresetChange(idx, 'yd', e.target.value)}
                               className="w-full p-1 border border-emerald-300 rounded text-center"
                             />
                           </td>
                           <td className="border border-emerald-200 p-1">
                             <input
                               type="number"
                               step="0.1"
                               value={s.dHAt}
                               onChange={(e) => handleSamplerPresetChange(idx, 'dHAt', e.target.value)}
                               className="w-full p-1 border border-emerald-300 rounded text-center"
                             />
                           </td>
                           <td className="border border-emerald-200 p-1">
                             <button type="button" onClick={() => removeSamplerPreset(idx)} className="text-red-500 hover:text-red-700 text-[11px] font-bold p-1">
                               삭제
                             </button>
                           </td>
                         </tr>
                       ))}
                     </tbody>
                   </table>
                 </div>
               </details>
               
               <div className="grid grid-cols-1 md:grid-cols-3 gap-4 mb-3 items-end">
                 <div>
                    <label className="block text-xs font-bold text-emerald-800 mb-1">목표 채취량 (SL)</label>
                    <input type="number" step="1" name="targetVolume" value={formData.targetVolume} onChange={handleChange} className="w-full p-2 border border-emerald-300 rounded text-sm bg-white font-bold outline-none focus:ring-2 focus:ring-emerald-500" />
                 </div>
                 <div>
                    <label className="block text-xs font-bold text-emerald-800 mb-1" title="가스미터 보정계수">보정계수 (Y<sub>d</sub>)</label>
                    <input type="number" step="0.001" name="gasMeterFactor" value={formData.gasMeterFactor} onChange={handleChange} placeholder="Yd" className="w-full p-2 border border-emerald-300 rounded text-sm bg-white outline-none focus:ring-2 focus:ring-emerald-500 font-bold text-emerald-700" />
                 </div>
                 <div>
                    <label className="block text-xs font-bold text-emerald-800 mb-1" title="오리피스 계수">오리피스 계수 (ΔH<sub>@</sub>)</label>
                    <input type="number" step="0.1" name="deltaHAt" value={formData.deltaHAt} onChange={handleChange} placeholder="ΔH@" className="w-full p-2 border border-emerald-300 rounded text-sm bg-white outline-none focus:ring-2 focus:ring-emerald-500 font-bold text-emerald-700" />
                 </div>
               </div>

               <div className="mb-4">
                  <button type="button" onClick={handleCalculateOptions} className="w-full py-3 bg-emerald-600 text-white rounded-lg font-bold text-sm hover:bg-emerald-700 transition shadow-md border border-emerald-700">
                    K-Factor 및 예상 채취시간 자동 산출
                  </button>
               </div>
               
               {recommendations && (
                 <div className="grid grid-cols-1 md:grid-cols-3 gap-4 mb-5">
                    <div className="bg-white border-2 border-emerald-300 p-4 rounded-lg shadow-sm hover:border-emerald-500 transition-colors">
                       <h4 className="font-bold text-emerald-800 mb-3 flex items-center gap-1 border-b border-emerald-100 pb-2">빠른 채취 <span className="text-[10px] font-bold text-emerald-600 ml-auto">(최대치, 50 이하)</span></h4>
                       <ul className="text-sm text-slate-700 space-y-2 mb-4">
                          <li className="flex justify-between font-bold"><span>추천 노즐:</span> <span className="font-black text-lg text-slate-900">No.{recommendations.fast.bestNozzleNum} <span className="text-sm font-medium">({recommendations.fast.finalDn}mm)</span></span></li>
                          <li className="flex justify-between font-bold"><span>예상 시간:</span> <span className="font-black text-emerald-700">{recommendations.fast.fastestTime} 분</span></li>
                          <li className="flex justify-between font-bold"><span>K-Factor:</span> <span className="font-black">{formatKFactorDisplay(recommendations.fast.calculatedK)}</span></li>
                          <li className="flex justify-between items-center bg-emerald-50 p-1 rounded border border-emerald-100 font-bold"><span>예상 ΔH:</span> <span className="font-black text-emerald-700 text-lg">{recommendations.fast.expectedDH}</span></li>
                       </ul>
                       <button type="button" onClick={() => applyRecommendation(recommendations.fast)} className="w-full py-2 bg-emerald-600 text-white rounded-lg text-sm font-bold hover:bg-emerald-700 shadow-sm">적용하기</button>
                    </div>
                    <div className="bg-teal-50 border-2 border-teal-400 p-4 rounded-lg shadow-sm hover:border-teal-500 transition-colors">
                       <h4 className="font-bold text-teal-900 mb-3 flex items-center gap-1 border-b border-teal-200 pb-2">표준 채취 <span className="text-[10px] font-bold text-teal-600 ml-auto">(ΔH 25 내외)</span></h4>
                       <ul className="text-sm text-slate-700 space-y-2 mb-4">
                          <li className="flex justify-between font-bold"><span>추천 노즐:</span> <span className="font-black text-lg text-slate-900">No.{recommendations.standard.bestNozzleNum} <span className="text-sm font-medium">({recommendations.standard.finalDn}mm)</span></span></li>
                          <li className="flex justify-between font-bold"><span>예상 시간:</span> <span className="font-black text-teal-700">{recommendations.standard.fastestTime} 분</span></li>
                          <li className="flex justify-between font-bold"><span>K-Factor:</span> <span className="font-black">{formatKFactorDisplay(recommendations.standard.calculatedK)}</span></li>
                          <li className="flex justify-between items-center bg-white p-1 rounded border border-teal-200 font-bold"><span>예상 ΔH:</span> <span className="font-black text-teal-700 text-lg">{recommendations.standard.expectedDH}</span></li>
                       </ul>
                       <button type="button" onClick={() => applyRecommendation(recommendations.standard)} className="w-full py-2 bg-teal-600 text-white rounded-lg text-sm font-bold hover:bg-teal-700 shadow-sm">적용하기</button>
                    </div>
                    <div className="bg-cyan-50 border-2 border-cyan-300 p-4 rounded-lg shadow-sm hover:border-cyan-500 transition-colors">
                       <h4 className="font-bold text-cyan-900 mb-3 flex items-center gap-1 border-b border-cyan-200 pb-2">안정/장시간 <span className="text-[10px] font-bold text-cyan-600 ml-auto">(ΔH 10 내외)</span></h4>
                       <ul className="text-sm text-slate-700 space-y-2 mb-4">
                          <li className="flex justify-between font-bold"><span>추천 노즐:</span> <span className="font-black text-lg text-slate-900">No.{recommendations.stable.bestNozzleNum} <span className="text-sm font-medium">({recommendations.stable.finalDn}mm)</span></span></li>
                          <li className="flex justify-between font-bold"><span>예상 시간:</span> <span className="font-black text-cyan-700">{recommendations.stable.fastestTime} 분</span></li>
                          <li className="flex justify-between font-bold"><span>K-Factor:</span> <span className="font-black">{formatKFactorDisplay(recommendations.stable.calculatedK)}</span></li>
                          <li className="flex justify-between items-center bg-white p-1 rounded border border-cyan-200 font-bold"><span>예상 ΔH:</span> <span className="font-black text-cyan-700 text-lg">{recommendations.stable.expectedDH}</span></li>
                       </ul>
                       <button type="button" onClick={() => applyRecommendation(recommendations.stable)} className="w-full py-2 bg-cyan-600 text-white rounded-lg text-sm font-bold hover:bg-cyan-700 shadow-sm">적용하기</button>
                    </div>
                 </div>
               )}
               
               <p className="text-[11px] text-emerald-700 mb-4 font-bold">※ K-Factor 및 예상 오리피스압(ΔH) 산출 후, 해당 노즐의 유량(Q<sub>sl</sub>)을 바탕으로 <span className="text-emerald-900">목표 채취량({formData.targetVolume || 1000}SL)에 도달하기 위한 필요 채취시간</span>을 자동 역산합니다.</p>

               {/* 접기/펴기 엑셀 전체 리스트 */}
               <details className="mt-4 mb-4">
                  <summary className="cursor-pointer text-xs font-bold text-emerald-700 hover:text-emerald-900 flex items-center justify-between gap-2">
                     <span>전체 노즐 예측표 열어보기</span>
                     <span className="text-[10px] text-slate-500">표 안에서 번호/직경 직접 기입 가능</span>
                  </summary>
                  <div className="mt-2 flex items-center gap-2">
                    <button type="button" onClick={addNozzleConfig} className="px-2 py-1 text-[11px] font-bold bg-emerald-100 text-emerald-800 rounded border border-emerald-300">행 추가</button>
                    <button type="button" onClick={resetNozzleDefaults} className="px-2 py-1 text-[11px] font-bold bg-slate-100 text-slate-700 rounded border border-slate-300">초기값 복원</button>
                  </div>
                  <div className="mt-3 overflow-x-auto">
                     <table className="w-full text-xs text-center border-collapse border border-emerald-300">
                       <thead className="bg-emerald-100 text-slate-800">
                          <tr>
                             <th className="border border-emerald-200 p-1 font-bold">노즐번호</th>
                             <th className="border border-emerald-200 p-1 font-bold">노즐직경(mm)</th>
                             <th className="border border-emerald-200 p-1 font-bold">K-factor</th>
                             <th className="border border-emerald-200 p-1 font-bold">ΔH</th>
                             <th className="border border-emerald-200 p-1 font-bold text-emerald-800">예상시간(분)</th>
                             <th className="border border-emerald-200 p-1 font-bold">(L)</th>
                             <th className="border border-emerald-200 p-1 font-bold">(SL)</th>
                             <th className="border border-emerald-200 p-1 font-bold">(Sm³)</th>
                             <th className="border border-emerald-200 p-1 font-bold">삭제</th>
                          </tr>
                       </thead>
                       <tbody className="bg-white">
                          {nozzleSet.map((n, idx) => {
                             let k='-', dh='-', vl='-', vsl='-', vsm='-', reqTime='-';
                             const nozzleNum = parseInt(n.num, 10);
                             const Dn = parseFloat(n.d);
                             const vs = getRawGasVelocity(), cP = parseFloat(formData.pitotFactor)||0.84, ts = getRawAvgTs(), pm = parseFloat(formData.atmPressure), ps = getRawStackPressure();
                             
                             const tmIn = parseFloat(formData.traverseTmIn) || parseFloat(formData.atmTemp) || 25;
                             const tmOut = parseFloat(formData.traverseTmOut) || parseFloat(formData.atmTemp) || 25;
                             const tm = (tmIn + tmOut) / 2;
                             
                             const validSps = formData.points.map(p => parseFloat(p.sp)).filter(v => !isNaN(v));
                             const avgSp = validSps.length === 0 ? 0 : validSps.reduce((a, b) => a + b, 0) / validSps.length;
                             const ps_ratio = (pm + avgSp / 13.6) / pm;

                             const gasComp = getGasComposition();
                             const Md = gasComp.Md, Ms = gasComp.Ms, Xw = gasComp.Xw;
                             const y = parseFloat(formData.gasMeterFactor)||1.0;
                             const dHAt = parseFloat(formData.deltaHAt);

                             if(!isNaN(nozzleNum) && !isNaN(Dn) && Dn > 0 && vs>0 && ps && pm>0) {
                                const k_val = 8.001e-5 * Math.pow(Dn, 4) * dHAt * Math.pow(cP, 2) * Math.pow(1 - Xw / 100, 2) * (Md / Ms) * ((tm + 273) / (ts + 273)) * ps_ratio;
                                k = formatKFactorDisplay(k_val);
                                dh = calcExpectedOrificeDH(k_val, getRawAvgDp());
                                const an = Math.PI * Math.pow(Dn/2000, 2);
                                const q_m = (vs * an * 60000 * ((tm+273)/(ts+273)) * (ps/pm) * (1-Xw/100)) / y;
                                const q_sl = q_m * y * (273/(tm+273)) * (pm/760);
                                
                                let time = Math.ceil((parseFloat(formData.targetVolume) || 1000) / q_sl);
                                if (!isFinite(time) || time <= 0) time = 0;
                                reqTime = time;
                                
                                const v_val = q_m * time;
                                vl = v_val.toFixed(1);
                                vsl = (q_sl * time).toFixed(1);
                                vsm = ((q_sl * time) / 1000).toFixed(3);
                             }
                             const isSelected = !isNaN(nozzleNum) && formData.usedNozzleNum === String(nozzleNum);
                             return (
                               <tr key={`${idx}-${n.num}-${n.d}`} className={isSelected ? "bg-emerald-100 font-bold" : "hover:bg-emerald-50"}>
                                 <td className="border border-emerald-200 p-1 font-bold text-slate-800">
                                   <div className="flex items-center justify-center gap-1">
                                     <input
                                       type="number"
                                       step="1"
                                       value={n.num}
                                       onChange={(e) => handleNozzleConfigChange(idx, 'num', e.target.value)}
                                       className="w-16 p-1 border border-emerald-300 rounded text-center font-bold"
                                       placeholder="번호"
                                     />
                                   </div>
                                   {!isNaN(nozzleNum) && formData.recommendedNozzleNum === String(nozzleNum) && (
                                     <span className="ml-1 text-[10px] bg-teal-600 text-white px-1.5 py-0.5 rounded-full shadow-sm">추천</span>
                                   )}
                                 </td>
                                 <td className="border border-emerald-200 p-1 font-bold">
                                   <input
                                     type="number"
                                     step="0.01"
                                     value={n.d}
                                     onChange={(e) => handleNozzleConfigChange(idx, 'd', e.target.value)}
                                     className="w-20 p-1 border border-emerald-300 rounded text-center font-bold"
                                     placeholder="직경"
                                   />
                                 </td>
                                 <td className="border border-emerald-200 p-1 font-black text-slate-700">{k}</td>
                                 <td className="border border-emerald-200 p-1 font-black text-emerald-700">{dh}</td>
                                 <td className="border border-emerald-200 p-1 font-black text-teal-800 bg-teal-50/50">{reqTime}</td>
                                 <td className="border border-emerald-200 p-1">{vl}</td>
                                 <td className="border border-emerald-200 p-1">{vsl}</td>
                                 <td className="border border-emerald-200 p-1">{vsm}</td>
                                 <td className="border border-emerald-200 p-1">
                                   <button type="button" onClick={() => removeNozzleConfig(idx)} className="text-red-500 hover:text-red-700 text-[11px] font-bold p-1">
                                     삭제
                                   </button>
                                 </td>
                               </tr>
                             )
                          })}
                          {nozzleSet.length === 0 && (
                            <tr>
                              <td className="border border-emerald-200 p-2 text-slate-500" colSpan={9}>
                                노즐 행이 없습니다. "행 추가" 버튼으로 입력을 시작하세요.
                              </td>
                            </tr>
                          )}
                       </tbody>
                     </table>
                  </div>
               </details>

               {/* 산정 결과 요약창 */}
               <div className="bg-white p-4 border border-emerald-200 rounded-lg shadow-inner">
                  {formData.recommendedNozzleNum && (
                    <div className="mb-4 p-3 bg-teal-50 border border-teal-200 rounded-lg flex items-center justify-between shadow-sm">
                      <div className="flex items-center gap-2">
                        <span className="font-black text-teal-900">시스템 추천 적정 노즐:</span>
                        <span className="text-lg font-black text-emerald-900 bg-emerald-200 px-3 py-1 border border-emerald-300 rounded-lg">
                          No.{formData.recommendedNozzleNum} ({formData.recommendedNozzleDia} mm)
                        </span>
                      </div>
                    </div>
                  )}
                  
                  <h4 className="text-sm font-bold text-slate-800 mb-3 border-b border-slate-200 pb-2">노즐 선정 및 채취량 예측 정보</h4>
                  <div className="flex flex-wrap justify-center gap-2 text-center">
                      <div className="bg-slate-50 py-2 px-1 rounded-lg border border-slate-200 flex-1 min-w-[90px]">
                          <div className="text-[10px] text-slate-500 font-bold">가스유속(V<sub>s</sub>)</div>
                          <div className="font-black text-slate-800 mt-1">{calcGasVelocity()}</div>
                      </div>
                      <div className="bg-emerald-50 py-2 px-1 rounded-lg border border-emerald-200 flex-1 min-w-[90px]">
                          <div className="text-[10px] text-emerald-700 font-bold">노즐경(D<sub>n</sub>)</div>
                          <div className="font-black text-emerald-900 mt-1">{formData.nozzleDiameter || '-'}</div>
                      </div>
                      <div className="bg-emerald-50 py-2 px-1 rounded-lg border border-emerald-200 flex-1 min-w-[90px]">
                          <div className="text-[10px] text-emerald-700 font-bold">K-Factor</div>
                          <div className="font-black text-emerald-900 mt-1">{formatKFactorDisplay(formData.kFactor)}</div>
                      </div>
                      <div className="bg-teal-50 py-2 px-1 rounded-lg border border-teal-200 flex-1 min-w-[90px]">
                          <div className="text-[10px] text-teal-700 font-bold">예상 ΔH</div>
                          <div className="font-black text-teal-900 mt-1">{expData.dH}</div>
                      </div>
                      <div className="bg-emerald-600 py-2 px-1 rounded-lg border border-emerald-700 flex-1 min-w-[90px] shadow-sm">
                          <div className="text-[10px] text-emerald-100 font-bold">예상시간(분)</div>
                          <div className="font-black text-white mt-1">{expData.reqTime}</div>
                      </div>
                      <div className="bg-cyan-50 py-2 px-1 rounded-lg border border-cyan-200 flex-1 min-w-[90px]">
                          <div className="text-[10px] text-cyan-800 font-bold">건조채취량(V<sub>m</sub>)</div>
                          <div className="font-black text-cyan-900 mt-1">{expData.L}</div>
                      </div>
                      <div className="bg-cyan-50 py-2 px-1 rounded-lg border border-cyan-200 flex-1 min-w-[90px]">
                          <div className="text-[10px] text-cyan-800 font-bold">표준채취량(SL)</div>
                          <div className="font-black text-cyan-900 mt-1">{expData.SL}</div>
                      </div>
                      <div className="bg-cyan-50 py-2 px-1 rounded-lg border border-cyan-200 flex-1 min-w-[90px]">
                          <div className="text-[10px] text-cyan-800 font-bold">표준채취량(Sm³)</div>
                          <div className="font-black text-cyan-900 mt-1">{expData.Sm3}</div>
                      </div>
                  </div>
               </div>
            </div>
            
            {/* 기본 채취 조건 */}
            <div className="grid grid-cols-1 md:grid-cols-2 gap-4 mb-6 bg-slate-50 p-4 rounded-xl border border-slate-200">
              <div>
                <label className="block text-xs font-bold text-slate-700 mb-1">사용 노즐경 (mm)</label>
                <input 
                  type="number" step="0.01" list="nozzle-list" name="nozzleDiameter" 
                  value={formData.nozzleDiameter} onChange={handleChange} 
                  placeholder="목록 선택 또는 입력"
                  className="w-full p-2 border border-slate-300 rounded-lg focus:ring-2 focus:ring-emerald-500 bg-white outline-none" 
                />
                <datalist id="nozzle-list">
                  {configuredNozzles.map(n => <option key={n.num} value={n.d.toFixed(2)}>No.{n.num} ({n.d}mm)</option>)}
                </datalist>
              </div>
              <div><label className="block text-xs font-bold text-slate-700 mb-1">적용 K-Factor</label>
                <input type="number" step="any" name="kFactor" value={formData.kFactor} onChange={handleKFactorChange} className="w-full p-2 border border-slate-300 bg-white rounded-lg focus:ring-2 focus:ring-emerald-500 font-black outline-none text-slate-800" /></div>
            </div>

            {/* 적산유량계 기록표 */}
            <div className="mb-6">
              <div className="flex justify-between items-end mb-2">
                <h3 className="text-sm font-bold text-slate-800">2단계: 적산유량계 기록표 (대기오염공정시험기준 통합 양식)</h3>
              </div>
              
              <div className="overflow-x-auto border border-slate-300 rounded-xl shadow-sm">
                <table className="w-full text-xs text-center min-w-[900px]">
                  <thead className="bg-slate-100 text-slate-700 border-b border-slate-300">
                    <tr>
                      <th className="p-1 font-bold whitespace-nowrap">순번</th>
                      <th className="p-1 font-bold whitespace-nowrap">채취점</th>
                      <th className="p-1 font-bold whitespace-nowrap">시간<br/>(분)</th>
                      <th className="p-1 font-bold whitespace-nowrap">배출가스<br/>온도(T<sub>s</sub>, ℃)</th>
                      <th className="p-1 font-bold text-emerald-700 whitespace-nowrap">동압<br/>(Δp)</th>
                      <th className="p-1 font-bold text-teal-700 whitespace-nowrap">오리피스압<br/>(ΔH)</th>
                      <th className="p-1 font-bold whitespace-nowrap">적산유량<br/>(V<sub>m</sub>, L)</th>
                      <th className="p-1 font-bold whitespace-nowrap">미터온도(T<sub>m</sub>, ℃)<br/><span className="text-[10px] text-slate-500">입구 | 출구</span></th>
                      <th className="p-1 font-bold whitespace-nowrap">진공압<br/>(mmHg)</th>
                      <th className="p-1 font-bold whitespace-nowrap">임핀저<br/>온도(℃)</th>
                      <th className="p-1 font-black text-emerald-600 border-l border-slate-300 whitespace-nowrap">순간 등속<br/>(I%)</th>
                      <th className="p-1"></th>
                    </tr>
                  </thead>
                  <tbody className="bg-white">
                    {formData.gasMeters.map((meter, idx) => {
                      const isStartRow = idx === 0;
                      const rowRate = calcRowIsokineticRate(idx);
                      const isRateValid = rowRate !== '-' && parseFloat(rowRate) >= 95 && parseFloat(rowRate) <= 105;
                      
                      return (
                        <tr key={meter.id} className={`border-b border-slate-100 last:border-0 hover:bg-slate-50 transition-colors`}>
                          <td className="p-1 font-bold text-slate-500 w-8">{isStartRow ? '초기' : meter.id}</td>
                          <td className="p-1 w-12">
                            {isStartRow ? (
                              <span className="font-bold text-slate-400">시작</span>
                            ) : (
                              <input type="text" value={meter.pointNum} onChange={(e) => handleGasMeterChange(idx, 'pointNum', e.target.value)} placeholder="입력" className="w-full p-1 border border-slate-200 rounded text-center outline-none focus:ring-2 focus:ring-emerald-500" />
                            )}
                          </td>
                          <td className="p-1 w-12">
                            {isStartRow ? (
                              <span className="font-bold text-slate-400">0</span>
                            ) : (
                              <input type="number" step="0.1" value={meter.time} onChange={(e) => handleGasMeterChange(idx, 'time', e.target.value)} placeholder="5" className="w-full p-1 border border-slate-200 rounded text-center outline-none focus:ring-2 focus:ring-emerald-500" />
                            )}
                          </td>
                          <td className="p-1 w-14">
                            <input type="number" step="0.1" value={meter.stackTemp} onChange={(e) => handleGasMeterChange(idx, 'stackTemp', e.target.value)} placeholder="Ts" className={`w-full p-1 border border-slate-200 rounded text-center outline-none focus:ring-2 focus:ring-emerald-500 ${isStartRow ? 'bg-slate-50' : ''}`} />
                          </td>
                          
                          <td className="p-1 w-14">
                            <input type="number" step="0.1" value={meter.dp} onChange={(e) => handleGasMeterChange(idx, 'dp', e.target.value)} className="w-full p-1 border border-emerald-300 bg-emerald-50 rounded text-center font-bold outline-none focus:ring-2 focus:ring-emerald-500" />
                          </td>
                          <td className="p-1 w-14">
                            <input type="number" step="0.1" value={meter.pressure} onChange={(e) => handleGasMeterChange(idx, 'pressure', e.target.value)} className="w-full p-1 border border-teal-300 bg-teal-50 rounded text-center font-black text-teal-800 outline-none focus:ring-2 focus:ring-teal-500" />
                          </td>
                          <td className="p-1 w-20">
                            <input type="number" step="0.1" value={meter.volume} onChange={(e) => handleGasMeterChange(idx, 'volume', e.target.value)} className={`w-full p-1 border ${isStartRow ? 'border-slate-300 bg-slate-100 text-slate-900 font-bold' : 'border-slate-200 bg-white font-medium'} rounded text-center outline-none focus:ring-2 focus:ring-emerald-500`} placeholder={isStartRow ? '초기유량' : ''} />
                          </td>
                          
                          <td className="p-1 w-28">
                            <div className="flex items-center gap-1">
                              <input type="number" step="0.1" value={meter.tmIn} onChange={(e) => handleGasMeterChange(idx, 'tmIn', e.target.value)} placeholder="입구" className="w-1/2 p-1 border border-slate-200 rounded text-center outline-none focus:ring-2 focus:ring-emerald-500" />
                              <input type="number" step="0.1" value={meter.tmOut} onChange={(e) => handleGasMeterChange(idx, 'tmOut', e.target.value)} placeholder="출구" className="w-1/2 p-1 border border-slate-200 rounded text-center outline-none focus:ring-2 focus:ring-emerald-500" />
                            </div>
                          </td>
                          
                          <td className="p-1 w-14">
                            <input type="number" step="0.1" value={meter.vacuum} onChange={(e) => handleGasMeterChange(idx, 'vacuum', e.target.value)} placeholder="압" className={`w-full p-1 border border-slate-200 rounded text-center outline-none focus:ring-2 focus:ring-emerald-500 ${isStartRow ? 'bg-slate-50' : ''}`} />
                          </td>
                          <td className="p-1 w-14">
                            <input type="number" step="0.1" value={meter.impingerTemp} onChange={(e) => handleGasMeterChange(idx, 'impingerTemp', e.target.value)} placeholder="온도" className={`w-full p-1 border border-slate-200 rounded text-center outline-none focus:ring-2 focus:ring-emerald-500 ${isStartRow ? 'bg-slate-50' : ''}`} />
                          </td>
                          
                          <td className="p-1 border-l border-slate-200 w-16">
                            <span className={`inline-block w-full py-1 rounded-lg font-black ${rowRate === '-' ? 'text-slate-400 bg-slate-100' : isRateValid ? 'text-emerald-900 bg-emerald-200' : 'text-white bg-red-500'}`}>
                              {rowRate === '-' ? '-' : `${rowRate}%`}
                            </span>
                          </td>
                          <td className="p-1 w-8">
                            {!isStartRow && <button type="button" onClick={() => removeGasMeter(idx)} className="text-red-500 hover:text-red-700 text-xs font-bold p-1 transition-colors">삭제</button>}
                          </td>
                        </tr>
                      );
                    })}
                  </tbody>
                </table>
              </div>
              <button type="button" onClick={addGasMeter} className="w-full mt-2 py-2 text-emerald-700 bg-emerald-100 hover:bg-emerald-200 rounded-lg text-sm font-bold border border-emerald-300 transition-colors shadow-sm">
                측정 기록 칸 추가
              </button>
            </div>

            {/* 유량계 기록 요약표 */}
            <div className="grid grid-cols-1 md:grid-cols-4 gap-4 mb-6">
              <div className="bg-slate-50 border border-slate-200 p-3 rounded-xl flex flex-col items-center justify-center">
                <span className="text-xs text-slate-600 font-bold mb-1">총 채취 시간 (분)</span>
                <span className="text-lg font-black text-slate-800">{getSamplingMinutes()}</span>
              </div>
              <div className="bg-emerald-50 border border-emerald-200 p-3 rounded-xl flex flex-col items-center justify-center">
                <span className="text-xs text-emerald-700 font-bold mb-1">총 채취 가스량 (V<sub>m</sub>)</span>
                <span className="text-xl font-black text-emerald-900">{calcGasMeterVolDiff() > 0 ? calcGasMeterVolDiff() : '-'}</span>
              </div>
              <div className="bg-teal-50 border border-teal-200 p-3 rounded-xl flex flex-col items-center justify-center">
                <span className="text-xs text-teal-700 font-bold mb-1">평균 온도 (T<sub>m</sub>, ℃)</span>
                <span className="text-lg font-black text-teal-900">{calcAvgTm()}</span>
              </div>
              <div className="bg-cyan-50 border border-cyan-200 p-3 rounded-xl flex flex-col items-center justify-center">
                <span className="text-xs text-cyan-700 font-bold mb-1">평균 오리피스압 (ΔH)</span>
                <span className="text-lg font-black text-cyan-900">{calcAvgOrifice()}</span>
              </div>
            </div>

            <div className="bg-white p-5 rounded-2xl flex flex-col justify-between shadow-lg border-2 border-slate-300">
              <div className="mb-4">
                <div>
                  <h3 className="font-black text-lg text-slate-900 tracking-tight">전체 평균 등속흡인율 (Isokinetic Rate)</h3>
                  <p className="text-xs text-slate-600 mt-1">대기오염공정시험기준 유효범위 : 95% ~ 105% (수식 0.00346 적용됨)</p>
                </div>
              </div>
              
              <div className="grid grid-cols-1 md:grid-cols-2 gap-4 mt-2">
                <div className="bg-slate-100 p-4 rounded-xl border border-slate-300 flex flex-col items-center justify-center">
                  <span className="text-xs text-slate-700 mb-1 font-bold">사전 수분량 기준 (전체)</span>
                  <div className="text-3xl font-black text-slate-900">{calcIsokineticRate(false)}<span className="text-base ml-1 text-slate-600">%</span></div>
                </div>
                <div className="bg-slate-900 p-4 rounded-xl border border-slate-800 flex flex-col items-center justify-center shadow-md relative overflow-hidden">
                  <div className="absolute top-0 left-0 w-full h-1 bg-white/50"></div>
                  <span className="text-xs text-slate-200 mb-1 font-bold">사후 수분량 기준 (최종)</span>
                  <div className="text-4xl font-black text-white">{calcIsokineticRate(true)}<span className="text-xl ml-1 text-slate-200">%</span></div>
                </div>
              </div>

              <div className="mt-5 flex justify-end">
                {calcIsokineticRate(true) !== '-' && (
                  <div>
                    {(parseFloat(calcIsokineticRate(true)) >= 95 && parseFloat(calcIsokineticRate(true)) <= 105) ? 
                      <span className="flex items-center text-white text-sm font-bold bg-green-700 px-4 py-2 rounded-full shadow-sm">적합 (최종 유효 데이터)</span> : 
                      <span className="flex items-center text-white text-sm font-bold bg-red-600 px-4 py-2 rounded-full shadow-sm">부적합 (재측정 요망)</span>
                    }
                  </div>
                )}
              </div>
            </div>
          </div>

          {/* 섹션 6: 시료 분석 및 비고 */}
          <div className="bg-white p-6 rounded-xl shadow-sm border border-slate-200">
            <h2 className="text-lg font-bold text-slate-800 mb-4 border-b border-slate-200 pb-2 flex items-center gap-2">
              <span className="bg-teal-600 text-white w-6 h-6 rounded-full flex items-center justify-center text-sm font-black">6</span>
              {isDustSheet ? '먼지 시료 무게 및 최종 결과' : '최종 결과 및 비고'}
            </h2>
            
            {isDustSheet && (
              <div className="grid grid-cols-1 lg:grid-cols-3 gap-6 mb-6">
                <div className="lg:col-span-2 grid grid-cols-3 gap-4 border border-slate-200 p-4 rounded-xl bg-slate-50 h-fit">
                  <div><label className="block text-xs font-bold text-slate-700 mb-1">여과지 번호</label>
                    <input type="text" name="filterId" value={formData.filterId} onChange={handleChange} className="w-full p-2 border border-slate-300 rounded-lg focus:ring-2 focus:ring-teal-500 outline-none" /></div>
                  <div><label className="block text-xs font-bold text-slate-700 mb-1">채취 전 무게 (mg)</label>
                    <input type="number" step="0.0001" name="filterInitial" value={formData.filterInitial} onChange={handleChange} placeholder="0.0000" className="w-full p-2 border border-slate-300 rounded-lg focus:ring-2 focus:ring-teal-500 outline-none" /></div>
                  <div><label className="block text-xs font-bold text-slate-700 mb-1">채취 후 무게 (mg)</label>
                    <input type="number" step="0.0001" name="filterFinal" value={formData.filterFinal} onChange={handleChange} placeholder="0.0000" className="w-full p-2 border border-slate-300 rounded-lg focus:ring-2 focus:ring-teal-500 outline-none" /></div>
                </div>
                <div className="lg:col-span-1 bg-teal-50 p-4 rounded-xl border border-teal-200 flex flex-col justify-center items-center text-teal-900 shadow-inner">
                  <span className="text-xs font-bold mb-1">포집된 먼지 무게 (W<sub>d</sub>)</span>
                  <span className="text-3xl font-black text-teal-700">{calcDustWeightDiff()} <span className="text-sm font-normal">mg</span></span>
                </div>
              </div>
            )}

            {/* 최종 농도 산출 결과 패널 */}
            <div className="bg-slate-800 p-6 rounded-xl border border-slate-700 shadow-md text-white mb-6 relative overflow-hidden">
               <h3 className="text-sm font-bold text-slate-300 mb-4 border-b border-slate-600 pb-3">
                 최종 산출 결과 요약
               </h3>
               <div className={`grid grid-cols-2 ${isDustSheet ? 'md:grid-cols-4' : 'md:grid-cols-2'} gap-6 divide-y md:divide-y-0 md:divide-x divide-slate-600 text-center`}>
                  <div className="flex flex-col items-center justify-center p-2">
                      <span className="text-xs text-slate-400 mb-2 font-bold">표준 습가스 유량 (Q<sub>sw</sub>)</span>
                      <span className="text-2xl font-bold text-emerald-300">{calcGasFlowRates().wet} <span className="text-sm font-normal text-slate-400">Sm³/hr</span></span>
                  </div>
                  <div className="flex flex-col items-center justify-center p-2">
                      <span className="text-xs text-slate-400 mb-2 font-bold">표준 건조가스 유량 (Q<sub>s</sub>)</span>
                      <span className="text-2xl font-bold text-cyan-300">{calcGasFlowRates().dry} <span className="text-sm font-normal text-slate-400">Sm³/hr</span></span>
                  </div>
                  {isDustSheet && (
                    <div className="flex flex-col items-center justify-center p-2">
                        <span className="text-xs text-slate-400 mb-2 font-bold">실측 먼지 농도 (C)</span>
                        <span className="text-2xl font-bold text-teal-300">{calcDustConcentrations().actualC} <span className="text-sm font-normal text-slate-400">mg/Sm³</span></span>
                    </div>
                  )}
                  {isDustSheet && (
                    <div className="flex flex-col items-center justify-center p-3 bg-slate-700/50 rounded-lg md:border-none md:bg-transparent">
                        <span className="text-xs text-red-300 font-bold mb-2">
                          O₂ 보정 농도 (C<sub>c</sub>)
                          {!formData.standardO2 && <span className="text-[10px] font-normal text-slate-400 ml-1">(제외)</span>}
                        </span>
                        <span className="text-3xl font-black text-white">{calcDustConcentrations().correctedC} <span className="text-sm font-normal text-slate-300">mg/Sm³</span></span>
                    </div>
                  )}
               </div>
            </div>

            <div className="mt-2">
              <label className="block text-xs font-bold text-slate-700 mb-1">현장 특이사항 및 비고</label>
              <textarea name="remarks" value={formData.remarks} onChange={handleChange} rows="2" className="w-full p-3 border border-slate-300 rounded-xl focus:ring-2 focus:ring-emerald-500 outline-none bg-slate-50/50" placeholder="측정공 상태, 장비 특이점 등"></textarea>
            </div>
          </div>

          <div className="flex items-center justify-center gap-4 mt-8 pb-8 border-b-2 border-slate-200 border-dashed">
            <button type="button" onClick={() => window.location.reload()} className="px-8 py-3.5 bg-white border border-slate-300 text-slate-700 font-bold rounded-xl hover:bg-slate-50 transition-colors shadow-sm">
              새 기록지 작성
            </button>
            <button type="submit" className={`px-8 py-3.5 text-white font-bold rounded-xl transition-colors shadow-md ${activeTheme.primaryButton}`}>
              기록부 데이터 저장
            </button>
          </div>
        </form>

        <div className="bg-slate-800 p-6 rounded-2xl shadow-lg mt-8 text-white border border-slate-700">
          <div className="flex flex-col md:flex-row justify-between items-center mb-6 gap-4 border-b border-slate-700 pb-4">
            <h2 className={`text-xl font-bold ${activeTheme.accentText}`}>
              저장된 종합 리포트 ({activeUser})
            </h2>
            <button
              onClick={exportToTemplateExcel}
              className={`px-5 py-2.5 text-white font-bold rounded-xl transition-colors shadow-md text-sm ${activeTheme.primaryButtonSoft}`}
            >
              템플릿 엑셀로 추출
            </button>
          </div>

          <div className="grid grid-cols-1 md:grid-cols-4 gap-3 mb-4">
            <div className="bg-slate-700/60 border border-slate-600 rounded-lg p-3 text-center">
              <p className="text-[11px] text-slate-300 font-bold">저장 건수</p>
              <p className="text-2xl font-black text-white">{savedData.length}</p>
            </div>
            <div className="bg-slate-700/60 border border-slate-600 rounded-lg p-3 text-center">
              <p className={`text-[11px] font-bold ${activeTheme.accentText}`}>등속 적합</p>
              <p className={`text-2xl font-black ${activeTheme.accentText}`}>{savedFitCount}</p>
            </div>
            <div className="bg-red-700/30 border border-red-500/40 rounded-lg p-3 text-center">
              <p className="text-[11px] text-red-200 font-bold">등속 부적합</p>
              <p className="text-2xl font-black text-red-200">{savedFailCount}</p>
            </div>
            <div className="bg-slate-700/80 border border-slate-500 rounded-lg p-3 text-center">
              <p className="text-[11px] text-slate-200 font-bold">평균 실측농도</p>
              <p className="text-2xl font-black text-white">{savedAvgActualC}</p>
            </div>
          </div>

          {savedData.length === 0 ? (
            <div className="p-8 text-center rounded-xl border border-dashed border-slate-600 text-slate-300">
              {activeUser
                ? (isUserUnlocked ? '현재 사용자의 저장 리포트가 없습니다.' : '로그인하면 저장 리포트가 표시됩니다.')
                : '로그인하거나 새로 계정을 생성해주세요.'}
            </div>
          ) : (
            <div className="overflow-x-auto rounded-xl border border-slate-700">
              <table className="w-full text-xs text-center min-w-[1850px]">
                <thead className="bg-slate-700 text-slate-200">
                  <tr>
                    <th className="p-2 font-bold">측정일자</th>
                    <th className="p-2 font-bold">사업장</th>
                    <th className="p-2 font-bold">배출구</th>
                    <th className="p-2 font-bold">측정자</th>
                    <th className="p-2 font-bold">노즐(No./mm)</th>
                    <th className="p-2 font-bold">K-Factor</th>
                    <th className="p-2 font-bold">채취시간(분)</th>
                    <th className="p-2 font-bold">채취량 Vm(L)</th>
                    <th className="p-2 font-bold">Vm_std(Sm³)</th>
                    <th className="p-2 font-bold">수분량(%)</th>
                    <th className="p-2 font-bold">등속흡인율(%)</th>
                    <th className="p-2 font-bold">먼지무게(mg)</th>
                    <th className="p-2 font-bold">실측농도</th>
                    <th className="p-2 font-bold">보정농도</th>
                    <th className="p-2 font-bold">건조유량(Sm³/hr)</th>
                    <th className="p-2 font-bold">비고</th>
                    <th className="p-2 font-bold">불러오기</th>
                    <th className="p-2 font-bold">삭제</th>
                  </tr>
                </thead>
                <tbody className="divide-y divide-slate-600 bg-slate-800">
                  {savedData.map((data) => (
                    <tr key={data.id || data.savedAt} className="hover:bg-slate-700/30 transition-colors cursor-pointer" onClick={() => handleLoadSavedReport(data)}>
                      <td className="p-2 whitespace-nowrap">{data.date || '-'}</td>
                      <td className="p-2">{data.company || '-'}</td>
                      <td className="p-2">{data.location || '-'}</td>
                      <td className="p-2">{data.sampler || '-'}</td>
                      <td className="p-2 font-bold text-emerald-300">No.{data.nozzleUsed || '-'} / {data.nozzleDiameterUsed || '-'}mm</td>
                      <td className="p-2">{formatKFactorDisplay(data.kFactorApplied)}</td>
                      <td className="p-2">{data.totalSamplingMinutes || '-'}</td>
                      <td className="p-2">{data.gasMeterVolume || '-'}</td>
                      <td className="p-2">{data.vmStd || '-'}</td>
                      <td className="p-2 text-cyan-300">{data.moisturePercent || '-'}</td>
                      <td className="p-2">
                        <span className={data.isokineticStatus === '적합' ? 'text-emerald-900 font-bold bg-emerald-400 px-2 py-0.5 rounded-md' : 'text-white font-bold bg-red-500 px-2 py-0.5 rounded-md'}>
                          {data.isokineticRate || '-'}
                        </span>
                      </td>
                      <td className="p-2 text-teal-300">{data.dustWeight || '-'}</td>
                      <td className="p-2 font-bold text-white">{data.actualConcentration || '-'}</td>
                      <td className="p-2 text-red-200 font-bold">{data.correctedConcentration || '-'}</td>
                      <td className="p-2">{data.dryFlowRate || '-'}</td>
                      <td className="p-2 text-left max-w-[280px]">{data.remarks || '-'}</td>
                      <td className="p-2">
                        <button
                          type="button"
                          onClick={(e) => {
                            e.stopPropagation();
                            handleLoadSavedReport(data);
                          }}
                          className="px-2 py-1 rounded border border-emerald-400 text-emerald-200 hover:bg-emerald-600/20 font-bold"
                        >
                          불러오기
                        </button>
                      </td>
                      <td className="p-2">
                        <button
                          type="button"
                          onClick={(e) => {
                            e.stopPropagation();
                            removeSavedReport(data.id, data.savedAt);
                          }}
                          className="px-2 py-1 rounded border border-red-400 text-red-200 hover:bg-red-600/20 font-bold"
                        >
                          삭제
                        </button>
                      </td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
          )}
        </div>

      </div>
      <SignatureBadge />
    </div>
  );
}
