/* 시료채취 프로그램(StackPilot) 첫 화면의 「판화 포스터풍 공단 풍경」. 성적서 발급 시스템(apps/성적서발급시스템.html)에 들어 있는 것을 그대로 옮김.
   원본: 시료채취 프로그램 src/App.jsx (성적서 쪽 gen_scene.js가 뽑아 만듦). 직접 고치지 말고 다시 뽑을 것 */
/* StackPilot(시료채취 기록부) 화면의 「판화 포스터풍 공단 풍경」 — gen_scene.js 가 App.jsx 에서 뽑아 만든 파일. 직접 고치지 말 것 */
(function(root){

  var VOID = { path:1, rect:1, circle:1, ellipse:1, line:1, polygon:1, polyline:1, stop:1, use:1, img:1, br:1, animate:1, animateTransform:1, feTurbulence:1, feColorMatrix:1, feGaussianBlur:1, feOffset:1, feBlend:1, feComposite:1, feFlood:1, feMergeNode:1, feDisplacementMap:1 };
  var KEEP = { viewBox:1, preserveAspectRatio:1, gradientUnits:1, gradientTransform:1, patternUnits:1, patternTransform:1, clipPathUnits:1, maskUnits:1, stdDeviation:1, baseFrequency:1, numOctaves:1, attributeName:1, repeatCount:1, calcMode:1, keyTimes:1, keySplines:1, keyPoints:1, attributeType:1, repeatDur:1, requiredExtensions:1, systemLanguage:1, primitiveUnits:1, filterUnits:1, maskContentUnits:1, patternContentUnits:1, surfaceScale:1, specularConstant:1, specularExponent:1, diffuseConstant:1, limitingConeAngle:1, pointsAtX:1, pointsAtY:1, pointsAtZ:1, kernelMatrix:1, kernelUnitLength:1, edgeMode:1, targetX:1, targetY:1, preserveAlpha:1, zoomAndPan:1, refX:1, refY:1, markerWidth:1, markerHeight:1, textLength:1, lengthAdjust:1, startOffset:1, tableValues:1, xChannelSelector:1, yChannelSelector:1, spreadMethod:1, pathLength:1 };
  var UNITLESS = { opacity:1, zIndex:1, fontWeight:1, lineHeight:1, flex:1, flexGrow:1, flexShrink:1, order:1, zoom:1, fillOpacity:1, strokeOpacity:1, scale:1 };
  function esc(s){ return String(s).replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;'); }
  function hyph(k){ return k.replace(/[A-Z]/g, function(m){ return '-' + m.toLowerCase(); }); }
  function css(o){ var a=[]; for (var k in o){ var v=o[k]; if (v==null || v===false || v==='') continue; var n = k.indexOf('--')===0 ? k : hyph(k); if (typeof v==='number' && k.indexOf('--')!==0 && !UNITLESS[k]) v = v + 'px'; a.push(n + ':' + v); } return a.join(';'); }
  function flat(c, out){ if (Array.isArray(c)) c.forEach(function(x){ flat(x, out); }); else if (c==null || c===false || c===true) {} else out.push(typeof c==='object' && c.__html!=null ? c.__html : esc(c)); return out; }
  function __spFrag(p){ return p.children; }
  function __spH(tag, props){ var kids = flat(Array.prototype.slice.call(arguments, 2), []);
    if (typeof tag === 'function'){ var r = tag(Object.assign({}, props || {}, { children: kids.map(function(x){ return { __html: x }; }) })); return { __html: flat(r, []).join('') }; }
    var a = ''; props = props || {};
    for (var k in props){ var v = props[k]; if (k==='key' || k==='ref' || k==='children' || v==null || v===false) continue;
      if (k==='className') k='class'; else if (k==='style'){ v = css(v); if (!v) continue; } else if (k==='dangerouslySetInnerHTML'){ kids=[v.__html]; continue; } else if (k==='htmlFor') k='for'; else if (!KEEP[k] && k.indexOf('aria-')!==0 && k.indexOf('data-')!==0) k = hyph(k);
      a += ' ' + k + '="' + esc(v===true ? '' : v) + '"'; }
    return { __html: VOID[tag] && !kids.length ? '<' + tag + a + '/>' : '<' + tag + a + '>' + kids.join('') + '</' + tag + '>' }; }

"use strict";
const seededNoise = (seed) => {
    const x = Math.sin(seed * 12.9898) * 43758.5453;
    return x - Math.floor(x);
};
const getSkyPhaseByHour = (date = new Date()) => {
    const hour = date.getHours();
    if (hour >= 6 && hour < 11)
        return 'morning';
    if (hour >= 11 && hour < 16)
        return 'day';
    if (hour >= 16 && hour < 19)
        return 'dusk';
    return 'night';
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
const SKYLINE_FLYER_CSS = `
@keyframes skylineTwinkle {
  0%, 100% { opacity: 0.35; }
  50% { opacity: 0.9; }
}
.skyline-star {
  position: absolute;
  border-radius: 9999px;
  background: #e9dfc6;
  animation: skylineTwinkle var(--twinkle-duration, 3.2s) ease-in-out infinite;
  animation-delay: var(--twinkle-delay, 0s);
}
/* 해·달: 번짐 없는 납작한 색면 (판화의 한 판) */
.skyline-moon {
  position: absolute;
  left: 10%;
  top: 9%;
  width: clamp(44px, 5vw, 66px);
  height: clamp(44px, 5vw, 66px);
  border-radius: 9999px;
  background: #e9dfc6;
  -webkit-mask: radial-gradient(circle at 68% 40%, transparent 50%, #000 51%);
  mask: radial-gradient(circle at 68% 40%, transparent 50%, #000 51%);
  transform: rotate(-20deg);
}
.skyline-moon-halo { display: none; }
.skyline-sun {
  position: absolute;
  left: 9%;
  top: 9%;
  width: clamp(56px, 6.5vw, 86px);
  height: clamp(56px, 6.5vw, 86px);
  border-radius: 9999px;
  background: #f7f1df;
}
.skyline-sunrise-mountain, .skyline-sunrise-mountain-sheet {
  position: absolute;
  left: 64%;
  top: 44%;
  width: clamp(130px, 17vw, 250px);
  height: clamp(130px, 17vw, 250px);
  border-radius: 9999px;
  background: #d9603b;
}
.skyline-sunrise-mountain-sheet {
  top: 62%;
  width: clamp(90px, 11vw, 160px);
  height: clamp(90px, 11vw, 160px);
}
/* 구름: 밑이 평평하고 위가 올록볼록한 납작한 구름 (연기와 같은 판화 모양, 색은 하늘에 가깝게 옅게) */
@keyframes skylineCloudDrift {
  from { transform: translateX(-30vw) scale(var(--cloud-scale, 1)); }
  to { transform: translateX(125vw) scale(var(--cloud-scale, 1)); }
}
.skyline-cloud {
  position: absolute;
  left: 0;
  width: 200px;
  height: 56px;
  transform-origin: left center;
  animation: skylineCloudDrift calc(var(--cloud-duration, 160s) * 2) linear infinite;
  animation-delay: var(--cloud-delay, 0s);
  will-change: transform;
}
.skyline-cloud svg { display: block; width: 100%; height: 100%; overflow: visible; }
/* 종이 결: 화면 전체에 깔리는 고운 알갱이 */
.skyline-grain {
  position: absolute;
  inset: 0;
  background-image: url("data:image/svg+xml,%3Csvg xmlns='http://www.w3.org/2000/svg' width='240' height='240'%3E%3Cfilter id='n'%3E%3CfeTurbulence type='fractalNoise' baseFrequency='.8' numOctaves='2' stitchTiles='stitch'/%3E%3CfeColorMatrix values='0 0 0 0 .1 0 0 0 0 .1 0 0 0 0 .1 0 0 0 .7 -.2'/%3E%3C/filter%3E%3Crect width='100%25' height='100%25' filter='url(%23n)'/%3E%3C/svg%3E");
  opacity: var(--grain-opacity, 0.5);
  mix-blend-mode: multiply;
}
@keyframes skylinePlume {
  from { transform: skewX(2deg); }
  to { transform: skewX(-7deg); }
}
/* 플레어 불꽃: 밑동은 붙인 채 길어졌다 짧아졌다 좌우로 일렁인다 */
@keyframes skylineFlare {
  0%, 100% { transform: scale(1, 1) skewX(0deg); }
  14% { transform: scale(0.88, 1.28) skewX(-7deg); }
  30% { transform: scale(1.1, 0.86) skewX(5deg); }
  46% { transform: scale(0.92, 1.4) skewX(-3deg); }
  62% { transform: scale(1.06, 0.95) skewX(8deg); }
  80% { transform: scale(0.9, 1.2) skewX(-5deg); }
}
.skyline-flare { animation: skylineFlare 1.7s ease-in-out infinite; }
.skyline-flare-core { animation: skylineFlare 1.1s ease-in-out infinite reverse; }
.skyline-plume {
  animation: skylinePlume var(--plume-duration, 16s) ease-in-out infinite alternate;
  animation-delay: var(--plume-delay, 0s);
}
@keyframes skylineBeacon {
  0%, 45% { opacity: 1; }
  55%, 100% { opacity: 0.15; }
}
.skyline-beacon {
  animation: skylineBeacon 2.8s linear infinite;
  animation-delay: var(--beacon-delay, 0s);
}
/* 학: 세 마리가 줄지어 하늘을 오른쪽에서 왼쪽으로 가로지른다. 한 번 지나가면 한동안 쉰다 */
@keyframes skylineCraneFly {
  0% { transform: translate(112vw, 7vh); }
  15% { transform: translate(74vw, 0.6vh); }
  62%, 100% { transform: translate(-45vw, -0.6vh); }
}
@keyframes skylineCraneFlap {
  from { transform: scaleY(1); }
  to { transform: scaleY(-0.78); }
}
@keyframes skylineCraneBob {
  from { transform: translateY(-2%); }
  to { transform: translateY(3%); }
}
.skyline-cranes {
  position: absolute;
  left: 0;
  top: 0.6%;
  width: clamp(170px, 19vw, 280px);
  aspect-ratio: 3.4;
  animation: skylineCraneFly 64s linear infinite;
  animation-delay: -8s;
  will-change: transform;
}
.skyline-crane { position: absolute; animation: skylineCraneBob 0.5s ease-in-out infinite alternate; animation-delay: var(--crane-delay, 0s); }
.skyline-crane svg { display: block; width: 100%; height: auto; overflow: visible; }
.skyline-crane-wing { transform-origin: 60px 31px; animation: skylineCraneFlap 0.5s ease-in-out infinite alternate; animation-delay: var(--crane-delay, 0s); }
/* 기러기: 일곱 마리가 ㅅ자로 하늘을 왼쪽에서 오른쪽으로 건넌다 */
@keyframes skylineGeeseFly {
  0% { transform: translate(-32vw, 5vh); }
  22% { transform: translate(12vw, 0.8vh); }
  62%, 100% { transform: translate(118vw, -0.6vh); }
}
.skyline-geese {
  position: absolute;
  left: 0;
  top: 2.2%;
  width: clamp(170px, 18vw, 270px);
  aspect-ratio: 3.4;
  animation: skylineGeeseFly 76s linear infinite;
  animation-delay: -9s;
  will-change: transform;
}
/* 나비: 둘레를 맴도는 길 + 날갯짓 */
@keyframes skylineFlutter {
  0%, 100% { transform: translate(0, 0); }
  14% { transform: translate(16px, -13px); }
  30% { transform: translate(34px, -3px); }
  46% { transform: translate(22px, -24px); }
  62% { transform: translate(-4px, -15px); }
  80% { transform: translate(-20px, -4px); }
}
@keyframes skylineWingbeat {
  from { transform: scaleX(1); }
  to { transform: scaleX(0.16); }
}
.skyline-flutter { animation: skylineFlutter var(--flutter-duration, 12s) ease-in-out infinite; animation-delay: var(--flutter-delay, 0s); }
.skyline-wingbeat { transform-box: fill-box; transform-origin: 50% 50%; animation: skylineWingbeat var(--beat, 0.28s) ease-in-out infinite alternate; }
/* 잠자리: 한자리에 떠 있다가 휙 옮겨 간다 + 날개 떨림 */
@keyframes skylineDart {
  0%, 17% { transform: translate(0, 0); }
  21%, 38% { transform: translate(52px, -12px); }
  42%, 60% { transform: translate(22px, -30px); }
  64%, 82% { transform: translate(-34px, -9px); }
  86%, 100% { transform: translate(0, 0); }
}
@keyframes skylineShimmer {
  from { opacity: 0.9; }
  to { opacity: 0.45; }
}
.skyline-dart { animation: skylineDart var(--dart-duration, 14s) ease-in-out infinite; animation-delay: var(--dart-delay, 0s); }
.skyline-shimmer { animation: skylineShimmer 0.09s linear infinite alternate; }
/* 너구리·오소리 걸음: 다리 두 쌍이 엇갈려 앞뒤로 움직이고 몸이 살짝 들썩인다 */
@keyframes skylineTrot {
  from { transform: rotate(13deg); }
  to { transform: rotate(-13deg); }
}
@keyframes skylineTrotBob {
  from { transform: translateY(0); }
  to { transform: translateY(-0.9px); }
}
.skyline-trot-a, .skyline-trot-b { transform-box: fill-box; transform-origin: 50% 0%; animation: skylineTrot var(--trot, 0.34s) ease-in-out infinite alternate; }
.skyline-trot-b { animation-direction: alternate-reverse; }
.skyline-trot-bob { animation: skylineTrotBob calc(var(--trot, 0.34s) / 2) ease-in-out infinite alternate; }
/* 개구리 울음: 턱 밑 울음주머니가 세 번 부풀었다 꺼진다 */
@keyframes skylineFrogCroak {
  0%, 16%, 25%, 34%, 43%, 100% { transform: scale(0); }
  20.5%, 29.5%, 38.5% { transform: scale(1); }
}
.skyline-frog-sac {
  transform: scale(0);
  animation: skylineFrogCroak var(--frog-duration, 9s) ease-in-out infinite;
  animation-delay: var(--frog-delay, 0s);
}
.skyline-svg {
  position: absolute;
  left: 0;
  right: 0;
  bottom: 0;
  width: 100%;
  height: 100%;
  display: block;
  overflow: visible;
}
@media (prefers-reduced-motion: reduce) {
  .skyline-star, .skyline-cloud, .skyline-plume, .skyline-beacon, .skyline-flare, .skyline-flare-core, .skyline-frog-sac, .skyline-crane, .skyline-crane-wing, .skyline-flutter, .skyline-wingbeat, .skyline-dart, .skyline-shimmer, .skyline-trot-a, .skyline-trot-b, .skyline-trot-bob { animation: none; }
  .skyline-geese { animation: none; transform: translate(60vw, 1vh); }
  .skyline-cranes { animation: none; transform: translate(72vw, 1vh); }
}
/* ── 첫 화면의 패널·버튼·카드: 흰 면을 부드럽게. 둥근 모서리, 가는 테두리, 낮게 깔리는 그림자 ── */
.sp-menu, .sp-modal { --sp-paper: #ffffff; --sp-paper-2: #f3f4f3; --sp-ink: #2a3944; --sp-soft: #66747c; --sp-line: rgba(42, 57, 68, 0.22); --sp-shadow: 0 1px 0 rgba(42, 57, 68, 0.1), 0 10px 24px -14px rgba(20, 28, 34, 0.55); }
.sp-modal { --sp-accent: #c4523a; }
.sp-panel {
  background: var(--sp-paper);
  border: 1px solid var(--sp-line);
  border-radius: 16px;
  box-shadow: var(--sp-shadow);
  color: var(--sp-ink);
}
.sp-kicker { font-size: 0.74rem; font-weight: 600; color: var(--sp-soft); }
.sp-title { font-weight: 800; color: var(--sp-ink); letter-spacing: -0.01em; }
.sp-btn {
  display: inline-flex; align-items: center; justify-content: center;
  padding: 0.45rem 0.95rem;
  border: 1px solid var(--sp-line); border-radius: 999px;
  background: var(--sp-paper); color: var(--sp-ink);
  font-size: 0.76rem; font-weight: 700; line-height: 1.2; white-space: nowrap;
  transition: background-color 0.15s ease, border-color 0.15s ease, transform 0.15s ease;
}
.sp-btn:hover { background: var(--sp-paper-2); border-color: rgba(42, 57, 68, 0.4); }
.sp-btn:active { transform: scale(0.97); }
.sp-btn-fill { background: var(--sp-ink); color: var(--sp-paper); border-color: var(--sp-ink); }
.sp-btn-fill:hover { background: #1f2c35; border-color: #1f2c35; }
.sp-btn-lg { padding: 0.7rem 1.6rem; font-size: 0.9rem; }
.sp-btn:focus-visible, .sp-card:focus-visible { outline: 3px solid var(--sp-accent); outline-offset: 2px; }
.sp-avatar { border: 1px solid var(--sp-line) !important; border-radius: 999px !important; }
.sp-menu .menu-avatar-fallback { background: var(--sp-ink) !important; color: var(--sp-paper) !important; }
/* 기록부 카드: 이름 아래에 붓으로 끝까지 한 번 그은 듯한 색 획으로 종류를 구분한다 */
.sp-card {
  display: block; text-align: left;
  padding: 1.1rem 1.25rem 1rem;
  background: var(--sp-paper);
  border: 1px solid var(--sp-line); border-radius: 16px;
  box-shadow: var(--sp-shadow);
  color: var(--sp-ink);
  transition: transform 0.18s ease, box-shadow 0.18s ease, background-color 0.18s ease;
}
.sp-card:hover { transform: translateY(-3px); background: #ffffff; box-shadow: 0 1px 0 rgba(42, 57, 68, 0.1), 0 18px 30px -16px rgba(20, 28, 34, 0.6); }
.sp-card:active { transform: translateY(0); }
.sp-card-mark {
  display: block; width: 100%; height: 0.42rem; margin: 0.35rem 0 0.5rem;
  background: var(--sp-card-color, var(--sp-ink));
  -webkit-mask: url("data:image/svg+xml,%3Csvg xmlns='http://www.w3.org/2000/svg' viewBox='0 0 120 12' preserveAspectRatio='none'%3E%3Cpath d='M2,7 C12,3 26,4 40,5 C58,6 74,3 92,4 C104,4.5 114,5 118,6 C116,8.5 104,9 90,8.6 C70,8 52,10 34,9 C20,8.4 8,9.5 2,7 Z'/%3E%3C/svg%3E") center / 100% 100% no-repeat;
  mask: url("data:image/svg+xml,%3Csvg xmlns='http://www.w3.org/2000/svg' viewBox='0 0 120 12' preserveAspectRatio='none'%3E%3Cpath d='M2,7 C12,3 26,4 40,5 C58,6 74,3 92,4 C104,4.5 114,5 118,6 C116,8.5 104,9 90,8.6 C70,8 52,10 34,9 C20,8.4 8,9.5 2,7 Z'/%3E%3C/svg%3E") center / 100% 100% no-repeat;
}
.sp-card-open { margin-top: 0.85rem; font-size: 0.72rem; font-weight: 600; color: var(--sp-soft); display: flex; justify-content: space-between; gap: 0.5rem; }
.sp-card:hover .sp-card-open { color: var(--sp-ink); }
.sp-empty { padding: 1.4rem; text-align: center; font-size: 0.78rem; color: var(--sp-soft); background: var(--sp-paper-2); border-radius: 12px; }
/* 설정·별명·로그인 창: 같은 종이 패널 */
div.sp-modal.sp-modal {
  background: var(--sp-paper) !important; color: var(--sp-ink);
  border: 1px solid var(--sp-line) !important; border-radius: 18px !important;
  box-shadow: 0 30px 60px -20px rgba(10, 16, 20, 0.6) !important;
  backdrop-filter: none !important; -webkit-backdrop-filter: none !important;
}
.sp-modal div[class*="bg-slate-50"], .sp-modal div.bg-white, .sp-modal div[class*="bg-emerald-50"], .sp-modal div[class*="bg-blue-50"] {
  background: var(--sp-paper-2) !important; border-color: transparent !important; border-radius: 12px !important; box-shadow: none !important;
  backdrop-filter: none !important; -webkit-backdrop-filter: none !important;
}
.sp-modal div[class*="bg-red-50"] { border-radius: 12px !important; }
.sp-modal button { border-radius: 999px !important; background-image: none !important; box-shadow: none !important; transition: background-color 0.15s ease, transform 0.15s ease; }
.sp-modal button.bg-white, .sp-modal button[class*="border-slate"], .sp-modal label[class*="border-slate"] {
  background: var(--sp-paper) !important; color: var(--sp-ink) !important; border: 1px solid var(--sp-line) !important;
}
.sp-modal button.bg-white:hover, .sp-modal button[class*="border-slate"]:hover { background: #f7f8f7 !important; }
.sp-modal button[class*="bg-emerald-6"], .sp-modal button[class*="bg-slate-7"], .sp-modal button[class*="bg-slate-8"], .sp-modal button[class*="bg-blue-6"] {
  background: var(--sp-ink) !important; color: var(--sp-paper) !important; border: 1px solid var(--sp-ink) !important;
}
/* 고른(먹색) 버튼은 마우스를 올려도 먹색 그대로: 흰 바탕으로 바뀌면 흰 글자가 안 보였다 */
.sp-modal button[class*="bg-emerald-6"]:hover, .sp-modal button[class*="bg-slate-7"]:hover, .sp-modal button[class*="bg-slate-8"]:hover, .sp-modal button[class*="bg-blue-6"]:hover { background: #1f2c35 !important; }
.sp-modal button:not(:disabled):active { transform: scale(0.97); }
/* 고른 테마 카드: 먹색 굵은 테두리 */
.sp-modal button.sp-theme-card { border-radius: 14px !important; }
.sp-modal button.sp-theme-card[aria-pressed="true"] { border-color: var(--sp-ink) !important; box-shadow: 0 0 0 2px var(--sp-ink) !important; background: #ffffff !important; }
.sp-modal input:not([type="checkbox"]):not([type="radio"]):not([type="range"]):not([type="file"]), .sp-modal select, .sp-modal textarea {
  background: #ffffff !important; border: 1px solid var(--sp-line) !important; border-radius: 10px !important; color: var(--sp-ink);
}
.sp-modal input:focus, .sp-modal select:focus, .sp-modal textarea:focus { border-color: var(--sp-ink) !important; box-shadow: 0 0 0 3px rgba(196, 82, 58, 0.28) !important; }
.sp-modal h3 { color: var(--sp-ink); }
@media (prefers-reduced-motion: reduce) { .sp-btn, .sp-card, .sp-modal button { transition: none; } }
`;
// ── 판화 포스터풍 공단 풍경 ──
// 잉크 다섯 가지쯤만 쓴다. 그라데이션·번짐 없이 납작한 색면과 가는 선으로만 그리고,
// 가장자리는 손으로 오린 듯 살짝 흔들리게, 화면은 한쪽(가운데 오른편)으로 몰아 비대칭으로 짠다.
// ── 배경 테마 × 시간대 ──
// 테마(기본·안개·세피아·청사진)마다 아침·낮·노을·밤 네 벌이 있다. 한 벌은 하늘, 해·달, 구름, 풍경 잉크, 글자 로고 색을 한꺼번에 정한다.
//  sky: 하늘 위·아래 색,  orb: sun(작은 해) · bigsun(산 위의 큰 해) · moon(초승달) · none,  stars: 별,  grain: 종이 결 세기
//  pal: 풍경 잉크 (light·beacon이 있으면 창에 불이 켜지고 굴뚝 경고등이 깜박인다),  wm: 글자 로고
const scene = (sky, orb, orbColor, cloud, pal, wm, extra = {}) => ({ sky, orb, orbColor, cloud, pal, wm, grain: 0.45, ...extra });
const ink = (ridge1, ridge2, far, inkColor, mid, band, plume, plumeShade, water, stroke, bank, light = null, beacon = null) => ({ ridge1, ridge2, far, ink: inkColor, mid, band, plume, plumeShade, water, stroke, bank, light, beacon });
const mark = (inkColor, mid, accent, band, plume, plumeShade) => ({ ink: inkColor, mid, accent, band, plume, plumeShade });
const SCENE_THEMES = {
    basic: {
        morning: scene(['#e6c296', '#f4e1c3'], 'bigsun', '#d9603b', ['#f4dfbe', '#ecd0a8'], ink('#e0ae8b', '#bf8672', '#93676a', '#2b2436', '#54425c', '#d9603b', '#f9ecd5', '#e9c9a5', '#eecda3', '#f9ecd5', '#1b1624'), mark('#2b2436', '#54425c', '#d9603b', '#f9ecd5', '#f9ecd5', '#e9c9a5')),
        day: scene(['#c5d3d2', '#ece6d6'], 'sun', '#f7f1df', ['#e4e9e1', '#d5ddd6'], ink('#b9c8c4', '#93aaaa', '#7b949a', '#2c4150', '#4f6875', '#c4523a', '#f7f2e5', '#dcd9ca', '#d5dcd0', '#f7f2e5', '#1d2c36'), mark('#2c4150', '#4f6875', '#c4523a', '#f7f2e5', '#f7f2e5', '#dcd9ca')),
        dusk: scene(['#3a3358', '#e08a5e'], 'bigsun', '#f7c873', ['#8a5a72', '#74486a'], ink('#9a5a6e', '#6f3f5e', '#55304f', '#1d1630', '#3a2a4c', '#f0a35e', '#f2c79f', '#d89a7c', '#de9668', '#f6d3aa', '#130d1d', '#ffd9a3', '#e5503f'), mark('#f6e3c8', '#cdb79c', '#e2603f', '#3a3358', '#8a5a72', '#74486a'), { stars: true }),
        night: scene(['#0f1b2a', '#24384c'], 'moon', '#e9dfc6', ['#22364a', '#1c2e40'], ink('#1e3348', '#18293b', '#132334', '#060c14', '#101c2a', '#7c2f29', '#34495e', '#2a3c50', '#1a2b3c', '#3b536a', '#03070c', '#e9b867', '#e5503f'), mark('#e9dfc6', '#bdb39a', '#e5503f', '#1c2e40', '#3b536a', '#2f4459'), { stars: true, grain: 0.4 }),
    },
    // 안개: 색을 거의 뺀 잿빛 한 계열. 시간대는 밝기와 온기만 바뀐다
    fog: {
        morning: scene(['#e2dbd0', '#f3eee5'], 'bigsun', '#e7bfa8', ['#ece6dc', '#dfd8cd'], ink('#d6cfc6', '#c4bcb4', '#b0a8a2', '#5e5a5c', '#7d787a', '#b56a55', '#f8f5ee', '#e6e0d6', '#e6e0d6', '#f8f5ee', '#423f42'), mark('#423f42', '#6a6668', '#b56a55', '#f8f5ee', '#f8f5ee', '#e6e0d6'), { grain: 0.5 }),
        day: scene(['#d6dbd4', '#eef0ea'], 'sun', '#f6f5ec', ['#e6e9e3', '#d9ddd6'], ink('#cfd5cf', '#bcc6c0', '#a8b4af', '#55666a', '#75878b', '#b56a55', '#f6f7f2', '#e3e6e0', '#dfe4de', '#f6f7f2', '#3b494d'), mark('#3b494d', '#5f7074', '#b56a55', '#f6f7f2', '#f6f7f2', '#e3e6e0'), { grain: 0.5 }),
        dusk: scene(['#8d8894', '#cfc3bf'], 'bigsun', '#e8ccb0', ['#aaa2aa', '#9a929b'], ink('#a79fa8', '#8f8792', '#7b7480', '#3d3a45', '#58545f', '#c07a5e', '#ddd3d0', '#c5bab8', '#c2b6b4', '#e2d9d6', '#26242c', '#f3d6a8', '#d9604a'), mark('#f1ebe6', '#c9c1bd', '#c07a5e', '#58545f', '#aaa2aa', '#9a929b'), { grain: 0.5 }),
        night: scene(['#2a2f34', '#454c52'], 'moon', '#dcdcd2', ['#3a4147', '#333a40'], ink('#3a4147', '#333a40', '#2c3238', '#15191d', '#232a30', '#6d3a33', '#4b545b', '#40484f', '#353c42', '#566068', '#0c0f12', '#e6c48a', '#d9604a'), mark('#dcdcd2', '#aeb0a8', '#c07a5e', '#2a2f34', '#4b545b', '#40484f'), { stars: true, grain: 0.4 }),
    },
    // 세피아: 오래된 인쇄물 같은 갈색 한 계열에 붉은 띠
    sepia: {
        morning: scene(['#efd2a4', '#f8ecd0'], 'bigsun', '#c9612f', ['#f6e3bd', '#ecd3a4'], ink('#e3b98a', '#cf9764', '#b07a4e', '#4a2f1e', '#714e36', '#c2502f', '#fdf5e2', '#eeddb9', '#efd6a6', '#fdf5e2', '#2c1c12'), mark('#4a2f1e', '#714e36', '#c2502f', '#fdf5e2', '#fdf5e2', '#eeddb9'), { grain: 0.55 }),
        day: scene(['#e3d2ac', '#f4ebd3'], 'sun', '#c9612f', ['#efe3c4', '#e3d3ac'], ink('#d9bb8c', '#c49a64', '#a67f50', '#47311f', '#6d5137', '#b5452c', '#fbf4e0', '#e9dcbd', '#e2d1a5', '#fbf4e0', '#2b1e14'), mark('#47311f', '#6d5137', '#b5452c', '#fbf4e0', '#fbf4e0', '#e9dcbd'), { grain: 0.55 }),
        dusk: scene(['#7a4a34', '#e0a064'], 'bigsun', '#f3c873', ['#a8683f', '#93582f'], ink('#a8683f', '#86502f', '#6b3e26', '#2a170e', '#4a2c1c', '#e8a85a', '#edc48e', '#d3a070', '#d99a5e', '#f3d2a0', '#1a0e08', '#ffd9a3', '#d9482f'), mark('#f8e6c4', '#d3bd98', '#d9482f', '#7a4a34', '#a8683f', '#93582f'), { grain: 0.55 }),
        night: scene(['#24170f', '#4a3322'], 'moon', '#ecdcb8', ['#3b2a1c', '#322317'], ink('#3b2a1c', '#322317', '#2a1d13', '#120b06', '#20150d', '#7a3220', '#4f3a28', '#433021', '#33241a', '#5e4630', '#0b0603', '#e9b867', '#d9482f'), mark('#ecdcb8', '#bfae8c', '#c2502f', '#24170f', '#4f3a28', '#433021'), { stars: true, grain: 0.45 }),
    },
    // 청사진: 파란 바탕에 흰 잉크로 찍은 도면
    blueprint: {
        morning: scene(['#5b8fc0', '#9cc3e2'], 'bigsun', '#fbe7ad', ['#8db6da', '#7ea9d0'], ink('#79a8d2', '#6898c6', '#a9c9e4', '#ffffff', '#d5e5f2', '#f2b84b', '#c4dcf0', '#b0cde6', '#4f84b8', '#cfe2f2', '#2a5f94'), mark('#ffffff', '#d5e5f2', '#f2b84b', '#2a5f94', '#c4dcf0', '#b0cde6'), { grain: 0.35 }),
        day: scene(['#1d4b76', '#2d6698'], 'none', '#eaf2f8', ['#356fa0', '#2d6494'], ink('#2a5f90', '#245684', '#4f82b0', '#eaf2f8', '#b3cbe0', '#f2b84b', '#4f84b4', '#42759f', '#1a456e', '#6f9dc4', '#0f3152'), mark('#eaf2f8', '#b3cbe0', '#f2b84b', '#1d4b76', '#4f84b4', '#42759f'), { grain: 0.35 }),
        dusk: scene(['#2a2f6a', '#5a5aa8'], 'bigsun', '#f2a04b', ['#4a4e98', '#40448a'], ink('#3b3f86', '#33377a', '#5b60a8', '#eef0fb', '#b9bde6', '#f2a04b', '#5d62ad', '#5055a0', '#262a62', '#7f84c8', '#171a47', '#ffe0a0', '#ff7a5a'), mark('#eef0fb', '#b9bde6', '#f2a04b', '#2a2f6a', '#5d62ad', '#5055a0'), { stars: true, grain: 0.35 }),
        night: scene(['#0b1e36', '#16365a'], 'moon', '#eaf2f8', ['#1c4470', '#183b62'], ink('#173a60', '#143354', '#2a5586', '#cfe0ee', '#94b2cc', '#e0a43f', '#244d78', '#1f436b', '#0c2440', '#3f6c98', '#07182c', '#ffe0a0', '#ff7a5a'), mark('#cfe0ee', '#94b2cc', '#e0a43f', '#0b1e36', '#244d78', '#1f436b'), { stars: true, grain: 0.35 }),
    },
};
const getSceneTheme = (theme, time) => (SCENE_THEMES[theme] || SCENE_THEMES.basic)[time] || SCENE_THEMES.basic.day;
const SKY_THEME_OPTIONS = [
    { id: 'basic', label: '기본' },
    { id: 'fog', label: '안개' },
    { id: 'sepia', label: '세피아' },
    { id: 'blueprint', label: '청사진' },
];
const sanitizeSkyTheme = (value) => (SKY_THEME_OPTIONS.some((opt) => opt.id === value) ? value : 'basic');
// 물가(공단 바닥) 높이
const SCENE_GROUND = 452;
// 굴뚝: x는 중심, w는 밑동 너비. plat는 작업대 높이 비율, plume은 연기 크기 배율(0이면 연기 없음)
const SCENE_STACKS = [
    { x: 700, w: 44, h: 384, plat: [0.52, 0.82], plume: 1.25 },
    { x: 776, w: 34, h: 300, plat: [0.6], plume: 0.9 },
    { x: 918, w: 26, h: 214, plat: [0.62], plume: 0 },
    { x: 318, w: 22, h: 190, plat: [0.64], plume: 0.6 },
];
const SCENE_FAR_STACKS = [{ x: 1452, w: 10, h: 132 }, { x: 1486, w: 8, h: 104 }, { x: 96, w: 9, h: 96 }, { x: 128, w: 7, h: 74 }];
const SCENE_TANKS = [{ x: 1010, w: 96, h: 58 }, { x: 1118, w: 96, h: 58 }, { x: 1226, w: 70, h: 44 }];
const SCENE_COLUMNS = [{ x: 838, w: 14, h: 150 }, { x: 862, w: 18, h: 118 }, { x: 962, w: 13, h: 132 }];
const SCENE_HALLS = [{ x: 520, w: 150, h: 46, teeth: 4 }, { x: 352, w: 130, h: 30, teeth: 0 }];
const SCENE_SILOS = [{ x: 214, w: 22, h: 74 }, { x: 240, w: 22, h: 74 }, { x: 266, w: 22, h: 74 }];
const SCENE_FLARE = { x: 1336, h: 236 };
// 앞 둑의 윗선 (x, 물가에서 내려온 높이). 풀은 이 선을 따라 심는다
const SCENE_BANK = [[0, 10], [140, 16], [300, 34], [520, 44], [1000, 50], [1320, 46], [1600, 30]];
const sceneBankY = (x) => {
    for (let i = 1; i < SCENE_BANK.length; i += 1) {
        const [x0, y0] = SCENE_BANK[i - 1], [x1, y1] = SCENE_BANK[i];
        if (x <= x1)
            return SCENE_GROUND + y0 + ((y1 - y0) * (x - x0)) / (x1 - x0);
    }
    return SCENE_GROUND + SCENE_BANK[SCENE_BANK.length - 1][1];
};
// 풀잎 한 장: 밑동이 넓고 끝이 뾰족한 휜 잎
const sceneBlade = (x, y, h, lean, w) => `M${x - w / 2},${y} Q${x + lean * 0.2 - w / 2},${y - h * 0.62} ${x + lean},${y - h} Q${x + lean * 0.4 + w / 2},${y - h * 0.5} ${x + w / 2},${y} Z`;
// 풀포기: 한 뿌리에서 부채꼴로 퍼지는 잎 여러 장
const sceneTuft = (x, size, seed) => {
    const y = sceneBankY(x) + 5, n = 5 + Math.floor(seededNoise(seed) * 4);
    return Array.from({ length: n }).map((_, k) => {
        const t = n === 1 ? 0 : k / (n - 1) - 0.5;
        const h = size * (0.55 + seededNoise(seed + k * 3 + 1) * 0.5) * (1 - Math.abs(t) * 0.5);
        return sceneBlade(x + t * size * 0.22, y, h, t * size * (0.7 + seededNoise(seed + k * 5 + 2) * 0.5), 2.2 + size * 0.035);
    }).join(' ');
};
const SCENE_TUFTS = [
    ...Array.from({ length: 24 }, (_, i) => ({ x: 6 + i * 13 + seededNoise(i * 3 + 7) * 8, size: 26 + seededNoise(i * 5 + 1) * 34, back: i % 3 === 1 })),
    ...Array.from({ length: 16 }, (_, i) => ({ x: 340 + i * 62 + seededNoise(i * 7 + 4) * 40, size: 12 + seededNoise(i * 11 + 2) * 14, back: false })),
    ...Array.from({ length: 18 }, (_, i) => ({ x: 1370 + i * 13 + seededNoise(i * 3 + 9) * 8, size: 22 + seededNoise(i * 5 + 6) * 32, back: i % 3 === 0 })),
];
// 부들: 길고 가는 줄기 끝에 통통한 이삭
const SCENE_CATTAILS = [
    ...Array.from({ length: 9 }, (_, i) => ({ x: 14 + i * 27 + seededNoise(i * 9 + 5) * 12, h: 62 + seededNoise(i * 4 + 8) * 40, lean: seededNoise(i * 6 + 3) * 16 - 4 })),
    ...Array.from({ length: 6 }, (_, i) => ({ x: 1436 + i * 29 + seededNoise(i * 8 + 1) * 12, h: 54 + seededNoise(i * 5 + 4) * 38, lean: seededNoise(i * 7 + 2) * -16 + 4 })),
];
// 개구리: 앞 둑 위를 폴짝폴짝 뛰어 왼쪽 풀숲과 오른쪽 풀숲 사이를 오간다.
// 착지점마다 잠깐 앉아 쉬고(가끔 길게 쉬며 운다), 포물선으로 뛰어 다음 착지점으로 간다. 끝까지 가면 돌아선다.
const SCENE_FROG_RANGE = [258, 1412];
const buildFrogTrack = (seed) => {
    const [x0, x1] = SCENE_FROG_RANGE;
    const xs = [];
    for (let x = x0, k = 0; x < x1 - 30; k += 1) {
        xs.push(x);
        x += 46 + seededNoise(seed + k * 7) * 26;
    }
    xs.push(x1);
    const pts = xs.map((x) => [x, sceneBankY(x) + 2]);
    // 가는 길 + 오는 길
    const stops = [...pts, ...pts.slice(0, -1).reverse()];
    const arc = (a, b) => {
        const dist = Math.abs(b[0] - a[0]);
        return [(a[0] + b[0]) / 2, Math.min(a[1], b[1]) - (20 + dist * 0.42)];
    };
    const f = (v) => v.toFixed(1);
    let path = `M${f(stops[0][0])},${f(stops[0][1])}`;
    const lengths = [0];
    for (let i = 1; i < stops.length; i += 1) {
        const a = stops[i - 1], b = stops[i], c = arc(a, b);
        path += ` Q${f(c[0])},${f(c[1])} ${f(b[0])},${f(b[1])}`;
        let len = 0, px = a[0], py = a[1];
        for (let k = 1; k <= 16; k += 1) {
            const t = k / 16, u = 1 - t;
            const x = u * u * a[0] + 2 * u * t * c[0] + t * t * b[0], y = u * u * a[1] + 2 * u * t * c[1] + t * t * b[1];
            len += Math.hypot(x - px, y - py);
            px = x;
            py = y;
        }
        lengths.push(lengths[i - 1] + len);
    }
    const total = lengths[lengths.length - 1];
    const HOP = 0.5;
    const motion = [], sit = [[0, 1]], up = [[0, 0]], down = [[0, 0]];
    let t = 0, turnAt = 0;
    for (let i = 0; i < stops.length - 1; i += 1) {
        // 쉬는 시간: 보통 1~2초, 네댓 번에 한 번은 길게, 돌아설 때도 길게
        let rest = 0.9 + seededNoise(seed + i * 13 + 5) * 1.3;
        if (i % 5 === 3)
            rest += 2.6;
        if (i === pts.length - 1) {
            rest += 2.4;
            turnAt = t + rest * 0.5;
        }
        motion.push([t, lengths[i] / total]);
        t += rest;
        motion.push([t, lengths[i] / total]);
        up.push([t, 1]);
        sit.push([t, 0]);
        up.push([t + HOP * 0.5, 0]);
        down.push([t + HOP * 0.5, 1]);
        t += HOP;
        down.push([t, 0]);
        sit.push([t, 1]);
    }
    motion.push([t, 1]);
    const T = t + 2.2; // 제자리로 돌아와 잠깐 쉰 뒤 다시 출발
    motion.push([T, 1]);
    const times = (list) => list.map(([time]) => (time / T).toFixed(5)).join(';');
    const vals = (list) => list.map(([, v]) => (typeof v === 'number' && v % 1 !== 0 ? v.toFixed(5) : v)).join(';');
    return {
        path, dur: T, turn: (turnAt / T).toFixed(5),
        motion: { keyTimes: times(motion), keyPoints: vals(motion) },
        sit: { keyTimes: times(sit), values: vals(sit) },
        up: { keyTimes: times(up), values: vals(up) },
        down: { keyTimes: times(down), values: vals(down) },
    };
};
const SCENE_FROG_TRACKS = [
    { ...buildFrogTrack(3), size: 0.82, phase: 0.04, croak: 9 },
    { ...buildFrogTrack(11), size: 0.68, phase: 0.55, croak: 13 },
];
// 앉은 모습(오른쪽을 본다)과 뛰는 모습(몸을 쭉 펴고 뒷다리를 뻗는다)
const SCENE_FROG_SIT = 'M-16,0 L-17,-2 Q-17,-8 -11,-11 L-4,-14.5 L3,-19 a4,4 0 0 1 7.6,0.6 L17,-13.5 Q18.6,-12 16.5,-10.8 L12,-9.5 Q10.5,-6.5 11,-4 L12.5,-1.2 L16,0 L8.5,0 L8,-4 Q4,-2.5 1,-4 L2.5,-1 L5,0 Z';
const SCENE_FROG_LEAP = 'M-30,-3 L-27,-6.5 L-15,-9 Q-6,-14.5 5,-14 a4,4 0 0 1 7.4,0.8 L20,-9.5 Q21.6,-7.6 19,-6.4 L14,-5 L17.5,-0.5 L14.5,0.5 L10.5,-4 Q2,-2.4 -6,-3.4 L-14,-4.2 L-27,-1.2 L-31,0.6 Z';
// 둑 위를 걸어 다니는 동물(너구리·오소리): 오른쪽 풀숲에서 나와 둑을 건너 왼쪽 풀숲으로 들어갔다가, 한참 뒤 되돌아온다
const SCENE_WALK_TRACK = (() => {
    const xs = [];
    for (let x = 1452; x >= 214; x -= 38)
        xs.push(x);
    const f = (v) => v.toFixed(1);
    return {
        path: `M${xs.map((x) => `${f(x)},${f(sceneBankY(x) + 2)}`).join(' L')}`,
        // 구간: 풀숲 속(0~4%) → 왼쪽으로 건넘(~40%) → 왼쪽 풀숲 속(~54%) → 오른쪽으로 건넘(~92%) → 풀숲 속
        keyTimes: '0;0.04;0.4;0.54;0.92;1',
        keyPoints: '0;0;1;1;0;0',
        turn: '0.47',
    };
})();
// 짧고 통통한 다리 한 짝 (맨 위는 몸통 속에 있어 그곳을 축으로 앞뒤로 움직인다)
const sceneWalkerLeg = (x, h) => `M${x},-${h} L${x - 0.3},-1.6 Q${x - 0.3},0 ${x + 1.4},0 L${x + 4.2},0 Q${x + 5.8},0 ${x + 5.6},-1.6 L${x + 5.2},-${h} Z`;
// 너구리: 만화 캐릭터처럼. 큰 머리, 둥근 귀, 눈가의 검은 띠와 큰 눈, 밝은 주둥이, 줄무늬 꼬리, 밝은 배. 걷는 쪽을 보는 옆모습
const renderSceneRaccoon = (pal) => (__spH(__spFrag, null,
    __spH("path", { d: "M-12,-19 C-21,-25 -29,-36 -40,-32 C-48,-29 -47,-18 -40,-15 C-32,-12 -22,-10.6 -12,-11.6 Z", fill: pal.mid }),
    __spH("path", { d: "M-40,-32 C-48,-29 -47,-18 -40,-15 C-38,-20 -38,-27 -40,-32 Z M-31.6,-32.4 C-29.6,-26 -29.6,-18.6 -32,-13 L-27,-11.8 C-24.6,-17.6 -24.6,-24.6 -26.4,-29.8 Z", fill: pal.bank }),
    __spH("ellipse", { cx: "-3", cy: "-14.6", rx: "15.5", ry: "11", fill: pal.mid }),
    __spH("path", { d: "M-13,-9.6 C-8,-3.4 6,-3.4 11,-9.6 C5,-6.8 -7,-6.8 -13,-9.6 Z", fill: pal.plume }),
    __spH("circle", { cx: "20", cy: "-37.6", r: "4.2", fill: pal.bank }),
    __spH("path", { d: "M1,-25 C1,-33 7,-38.6 14.6,-38.6 C21,-38.6 26,-34.6 27.6,-29 C28,-27.6 28.2,-26.4 28.2,-25.2 L28.2,-18 C26.6,-13.6 22,-10.8 15.6,-10.8 C11.6,-10.8 8.4,-12 6,-14 L0.6,-12.4 L3,-18.4 C1.6,-20.4 1,-22.6 1,-25 Z", fill: pal.mid }),
    __spH("circle", { cx: "6.6", cy: "-35", r: "4.8", fill: pal.bank }),
    __spH("circle", { cx: "7", cy: "-34.6", r: "2.3", fill: pal.plume, opacity: "0.85" }),
    __spH("path", { d: "M27,-29.4 C22.6,-32.8 15.6,-31.6 11.4,-27.2 C8.6,-24.2 7.6,-20.8 8.4,-18.2 C12.8,-17.8 17.4,-19.6 21,-22.6 C23.6,-24.6 25.8,-27 27,-29.4 Z", fill: pal.bank }),
    __spH("path", { d: "M19.6,-22.6 C24.4,-24 30,-22.6 32.8,-19.6 C34,-18.2 33.6,-16.4 32,-15.8 C28,-14.4 23.4,-13.8 19.6,-14.8 C18,-17 18,-20.6 19.6,-22.6 Z", fill: pal.plume }),
    __spH("ellipse", { cx: "32", cy: "-19", rx: "2.1", ry: "1.7", fill: pal.bank }),
    __spH("ellipse", { cx: "19.2", cy: "-26.2", rx: "2.9", ry: "3.2", fill: pal.plume }),
    __spH("circle", { cx: "19.9", cy: "-26", r: "1.9", fill: pal.bank }),
    __spH("circle", { cx: "20.6", cy: "-26.9", r: "0.65", fill: pal.plume })));
// 오소리: 낮고 긴 몸, 흰 머리에 코에서 눈을 지나 귀로 가는 검은 줄, 검은 다리와 배, 짧은 꼬리
const renderSceneBadger = (pal) => (__spH(__spFrag, null,
    __spH("path", { d: "M-22,-17 C-30,-21 -36,-15 -33,-9.6 C-29,-8 -25,-8.6 -21,-10 Z", fill: pal.mid }),
    __spH("ellipse", { cx: "-4", cy: "-13", rx: "20.5", ry: "10", fill: pal.mid }),
    __spH("path", { d: "M-21,-8.6 C-12,-3 8,-3 17,-8.6 C8,-5.8 -12,-5.8 -21,-8.6 Z", fill: pal.bank }),
    __spH("circle", { cx: "15", cy: "-28", r: "3.6", fill: pal.bank }),
    __spH("path", { d: "M12.4,-30.4 C13.8,-32 16.2,-32 17.6,-30.4", stroke: pal.plume, strokeWidth: "1.2", fill: "none", strokeLinecap: "round" }),
    __spH("path", { d: "M8.6,-21 C10.6,-27.4 18,-30.4 25,-28 C31,-26 36.4,-20.4 39.2,-15.2 C40.2,-13.4 39.2,-11.6 37.2,-11.6 L29,-10 C22,-8 14,-8.6 9.6,-11.6 C7.6,-14.6 7.6,-18 8.6,-21 Z", fill: pal.plume }),
    __spH("path", { d: "M38.8,-15.8 C33,-20 26,-24.4 17.4,-27.8 C15,-26 13.4,-23.6 12.8,-21 C21,-20 28.4,-17 35.8,-12.6 C37.2,-13.2 38.2,-14.4 38.8,-15.8 Z", fill: pal.bank }),
    __spH("circle", { cx: "25.6", cy: "-21.4", r: "1.6", fill: pal.plume }),
    __spH("circle", { cx: "26", cy: "-21.2", r: "0.9", fill: pal.bank }),
    __spH("ellipse", { cx: "38.6", cy: "-13.6", rx: "2", ry: "1.7", fill: pal.bank }),
    __spH("path", { d: "M29,-10 C22,-8 14,-8.6 9.6,-11.6 C15,-12.6 23,-12.8 31,-12.4 Z", fill: pal.bank, opacity: "0.85" })));
// 나비(봄): 풀숲과 들꽃 둘레를 팔랑팔랑 맴돈다.  잠자리(여름·가을): 물 위에 떠 있다가 휙 자리를 옮긴다
const SCENE_BUTTERFLIES = [
    { x: 166, up: 58, dur: 11, delay: 2, beat: 0.26 }, { x: 322, up: 38, dur: 14, delay: 6, beat: 0.31 },
    { x: 640, up: 26, dur: 16, delay: 9, beat: 0.24 }, { x: 1404, up: 52, dur: 12, delay: 4, beat: 0.29 },
];
const SCENE_DRAGONFLIES = [
    { x: 438, up: 34, dur: 13, delay: 3, dir: 1 }, { x: 912, up: 52, dur: 17, delay: 8, dir: -1 }, { x: 1248, up: 30, dur: 15, delay: 11, dir: 1 },
];
// 기러기(가을): 일곱 마리가 ㅅ자 대형으로 하늘을 가로지른다 (맨 앞이 오른쪽)
const SCENE_GEESE = [
    { left: '84%', top: '38%', delay: 0 },
    { left: '67%', top: '20%', delay: -0.14 }, { left: '67%', top: '58%', delay: -0.31 },
    { left: '50%', top: '4%', delay: -0.22 }, { left: '50%', top: '76%', delay: -0.08 },
    { left: '33%', top: '-10%', delay: -0.37 }, { left: '16%', top: '-24%', delay: -0.19 },
];
const renderSceneGoose = (pal) => (__spH("svg", { viewBox: "0 0 120 64", "aria-hidden": "true" },
    __spH("g", { transform: "translate(120 0) scale(-1 1)" },
        __spH("g", { className: "skyline-crane-wing" },
            __spH("path", { d: "M56,31 Q62,15 79,8 Q82,18 77,31 Z", fill: pal.mid })),
        __spH("path", { d: "M44,31 Q54,25 68,27 Q79,28 88,32 Q74,37 58,36 Q49,35 44,31 Z", fill: pal.ink }),
        __spH("path", { d: "M47,29 L27,26 L27,29.2 L46,33 Z", fill: pal.ink }),
        __spH("ellipse", { cx: "25", cy: "27.6", rx: "4", ry: "2.9", fill: pal.ink }),
        __spH("path", { d: "M22,26.6 L14,28 L22,29 Z", fill: pal.ink }),
        __spH("g", { className: "skyline-crane-wing" },
            __spH("path", { d: "M49,31 Q52,16 70,3 Q77,12 73,31 Z", fill: pal.ink })))));
// 학(두루미) 무리: 맨 앞이 가장 크고, 뒤따르는 두 마리는 조금 작게 위·아래로 어긋나 난다
const SCENE_CRANES = [
    { left: '0%', top: '20%', width: '44%', delay: 0 },
    { left: '40%', top: '0%', width: '36%', delay: -0.22 },
    { left: '66%', top: '42%', width: '32%', delay: -0.41 },
];
// 학 한 마리: 흰 몸, 검은 목·날개 끝·다리, 붉은 정수리. 목은 앞으로, 다리는 뒤로 곧게 뻗는다
const renderSceneCrane = (pal) => (__spH("svg", { viewBox: "0 0 120 64", "aria-hidden": "true" },
    __spH("g", { className: "skyline-crane-wing" },
        __spH("path", { d: "M56,31 Q62,13 80,5 Q83,17 78,31 Z", fill: pal.plumeShade })),
    __spH("path", { d: "M76,32 L112,36 M112,36 l6,-1.2 M112,36 l6,2", stroke: pal.ink, strokeWidth: "1.3", strokeLinecap: "round", fill: "none" }),
    __spH("path", { d: "M44,31 Q54,25 68,27 Q78,28 84,32 Q72,37 58,36 Q49,35 44,31 Z", fill: pal.plume }),
    __spH("path", { d: "M70,28.5 Q81,28.5 90,33.5 Q79,35.5 72,34 Z", fill: pal.ink }),
    __spH("path", { d: "M47,29.2 L20,25.6 L20,28.6 L46,33.2 Z", fill: pal.ink }),
    __spH("ellipse", { cx: "18", cy: "27", rx: "4.2", ry: "3.1", fill: pal.plume }),
    __spH("circle", { cx: "18.6", cy: "24.6", r: "1.8", fill: pal.band }),
    __spH("path", { d: "M15,25.8 L2,28 L15,28.8 Z", fill: pal.ink }),
    __spH("g", { className: "skyline-crane-wing" },
        __spH("path", { d: "M49,31 Q51,15 71,0 Q78,10 73,31 Z", fill: pal.plume }),
        __spH("path", { d: "M71,0 Q78,10 73,31 L67,31 Q71,16 64,6 Z", fill: pal.ink }))));
const renderSkylineSvg = ({ pal, scope, compact = false, season = 'summer' }) => {
    const id = (name) => `${scope}-sky-${name}`;
    const warm = season === 'spring' || season === 'summer'; // 개구리·새끼 오리가 나오는 철
    const stillScene = typeof window !== 'undefined' && window.matchMedia ? window.matchMedia('(prefers-reduced-motion: reduce)').matches : false;
    const G = SCENE_GROUND;
    const stackScale = compact ? 0.66 : 1;
    const night = Boolean(pal.light);
    const cut = `url(#${id('cut')})`;
    const halfAt = (st, t) => (st.w / 2) * (1 - 0.36 * t);
    // 창: 낮에는 한 단계 밝은 잉크, 밤에는 몇 칸만 불이 켜진다
    const windowFill = (seed) => (night ? (seededNoise(seed) > 0.42 ? pal.light : pal.mid) : pal.mid);
    // 연기: 목판화 구름처럼 세 번 굽이치는 납작한 덩어리 + 아랫면 그늘 한 판
    // 밑동은 굴뚝 구멍 너비(b)만큼 벌리고 굴뚝 꼭대기 테 안쪽에서 시작한다 → 연기가 굴뚝 입구에 꼭 맞는다
    // [a, c, t]: 첫째·둘째 봉우리가 솟는 정도와 꼬리가 늘어나는 정도. 이 값을 바꿔 가며 연기가 피어오르게 한다
    const plumePath = (cx, top, s, b, [a, c, t] = [0, 0, 0]) => {
        const p = (dx, dy) => `${cx + dx * s},${top - dy * s}`;
        return `M${cx - b},${top + 5} L${cx - b},${top} C${p(-12, 30)} ${p(8, 60)} ${p(44, 62)} C${p(50, 90 + a)} ${p(98, 100 + a)} ${p(122, 78 + (a + c) / 4)} C${p(152, 106 + c)} ${p(216, 106 + c)} ${p(230 + t / 2, 72 + c / 4)} C${p(266 + t, 86)} ${p(306 + t, 64)} ${p(300 + t, 42)} C${p(252 + t / 2, 28)} ${p(192, 40)} ${p(150, 33)} C${p(100, 25)} ${p(40, 34)} ${cx + b},${top} L${cx + b},${top + 5} Z`;
    };
    const plumeShadePath = (cx, top, s, [, , t] = [0, 0, 0]) => {
        const p = (dx, dy) => `${cx + dx * s},${top - dy * s}`;
        return `M${p(40, 44)} C${p(90, 52)} ${p(120, 44)} ${p(150, 52)} C${p(200, 64)} ${p(250 + t / 2, 52)} ${p(300 + t, 42)} C${p(252 + t / 2, 28)} ${p(192, 40)} ${p(150, 33)} C${p(100, 25)} ${p(60, 32)} ${p(40, 44)} Z`;
    };
    // 꼬리 끝에서 떨어져 나가는 작은 연기 덩이
    const plumePuffPath = (cx, top, s) => {
        const x = cx + 286 * s, y = top - 44 * s, q = (v) => v * s;
        return `M${x},${y} c${q(-10)},0 ${q(-10)},${q(-12)} ${q(2)},${q(-12)} c${q(3)},${q(-12)} ${q(24)},${q(-14)} ${q(31)},${q(-4)} c${q(12)},${q(-5)} ${q(24)},${q(5)} ${q(17)},${q(12)} c${q(5)},${q(2)} ${q(2)},${q(7)} ${q(-5)},${q(7)} Z`;
    };
    const PLUME_FRAMES = [[0, 0, 0], [13, -9, 18], [-7, 12, 6], [0, 0, 0]];
    const plumeAnim = (values, dur) => (__spH("animate", { attributeName: "d", dur: `${dur}s`, repeatCount: "indefinite", calcMode: "spline", keyTimes: "0;0.34;0.68;1", keySplines: "0.4 0 0.6 1;0.4 0 0.6 1;0.4 0 0.6 1", values: values.join(';') }));
    const stacks = SCENE_STACKS.map((st) => ({ ...st, hh: Math.round(st.h * stackScale) }));
    const tufts = compact ? SCENE_TUFTS.filter((_, i) => i % 2 === 0) : SCENE_TUFTS;
    const cattails = compact ? SCENE_CATTAILS.filter((_, i) => i % 2 === 0) : SCENE_CATTAILS;
    const flareTop = G - SCENE_FLARE.h * stackScale;
    return (__spH("svg", { className: "skyline-svg", viewBox: "0 0 1600 520", preserveAspectRatio: "xMidYMax slice", "aria-hidden": "true" },
        __spH("defs", null,
            __spH("filter", { id: id('cut'), x: "-2%", y: "-2%", width: "104%", height: "104%" },
                __spH("feTurbulence", { type: "fractalNoise", baseFrequency: "0.02", numOctaves: "2", seed: "7", result: "n" }),
                __spH("feDisplacementMap", { in: "SourceGraphic", in2: "n", scale: compact ? 1.4 : 2, xChannelSelector: "R", yChannelSelector: "G" }))),
        __spH("g", { filter: cut },
            __spH("path", { d: "M0,356 L90,330 L170,342 L260,300 L340,328 L420,312 L520,344 L640,322 L760,350 L900,318 L1010,340 L1120,296 L1230,332 L1330,306 L1450,338 L1540,320 L1600,330 L1600,520 L0,520 Z", fill: pal.ridge1 }),
            __spH("path", { d: "M0,392 L120,372 L250,384 L360,360 L470,382 L600,396 L1010,396 L1130,372 L1260,350 L1380,366 L1480,344 L1600,358 L1600,520 L0,520 Z", fill: pal.ridge2 }),
            __spH("g", { fill: pal.far },
                SCENE_FAR_STACKS.map((st, i) => __spH("rect", { key: `far-${i}`, x: st.x - st.w / 2, y: G - st.h, width: st.w, height: st.h })),
                __spH("path", { d: `M1400,${G} L1400,${G - 30} L1440,${G - 40} L1440,${G - 30} L1530,${G - 30} L1530,${G - 18} L1600,${G - 18} L1600,${G} Z M40,${G} L40,${G - 22} L84,${G - 22} L84,${G - 30} L150,${G - 30} L150,${G} Z` }),
                [1540, 1566].map((x) => __spH("path", { key: x, d: `M${x},${G} L${x},${G - 24} Q${x + 11},${G - 32} ${x + 22},${G - 24} L${x + 22},${G} Z` })))),
        !compact && (__spH("g", { stroke: pal.far, fill: "none", strokeWidth: "1.2", strokeLinecap: "round" },
            [[1180, 66], [1272, 60]].map(([x, h]) => (__spH("path", { key: x, d: `M${x - 9},${G - 22} L${x - 2},${G - 22 - h} L${x + 2},${G - 22 - h} L${x + 9},${G - 22} M${x - 12},${G - 22 - h + 12} L${x + 12},${G - 22 - h + 12} M${x - 9},${G - 22 - h + 24} L${x + 9},${G - 22 - h + 24} M${x - 6},${G - 40} L${x + 6},${G - 52} M${x + 6},${G - 40} L${x - 6},${G - 52}` }))),
            __spH("path", { d: `M1090,${G - 62} Q1136,${G - 60} 1168,${G - 76} M1192,${G - 76} Q1230,${G - 60} 1260,${G - 70} M1284,${G - 70} Q1340,${G - 52} 1400,${G - 58}`, strokeWidth: "0.8" }))),
        stacks.filter((st) => st.plume > 0).map((st, i) => {
            const s = st.plume * (compact ? 0.7 : 1), top = G - st.hh, b = Math.max(3, halfAt(st, 1) - 1.5);
            return (__spH("g", { key: `plume-${i}`, className: "skyline-plume", style: { '--plume-duration': `${6 + i * 2}s`, '--plume-delay': `-${i * 3}s`, transformOrigin: `${st.x}px ${top}px` } },
                __spH("g", { opacity: "0" },
                    __spH("path", { d: plumePuffPath(st.x, top, s), fill: pal.plume }),
                    __spH("animateTransform", { attributeName: "transform", type: "translate", values: `0 0;${150 * s} ${-18 * s}`, dur: `${9 + i * 2}s`, begin: `-${i * 4}s`, repeatCount: "indefinite" }),
                    __spH("animate", { attributeName: "opacity", values: "0;0.9;0.75;0", keyTimes: "0;0.12;0.55;1", dur: `${9 + i * 2}s`, begin: `-${i * 4}s`, repeatCount: "indefinite" })),
                __spH("path", { d: plumePath(st.x, top, s, b), fill: pal.plume }, plumeAnim(PLUME_FRAMES.map((v) => plumePath(st.x, top, s, b, v)), 7 + i * 2)),
                __spH("path", { d: plumeShadePath(st.x, top, s), fill: pal.plumeShade }, plumeAnim(PLUME_FRAMES.map((v) => plumeShadePath(st.x, top, s, v)), 7 + i * 2))));
        }),
        __spH("g", { filter: cut },
            __spH("path", { d: `M664,${G - 72} L684,${G - 72} L684,${G - 54} L664,${G - 54} Z M744,${G - 66} L762,${G - 66} L762,${G - 50} L744,${G - 50} Z M326,${G - 27} L354,${G - 27} L354,${G - 15} L326,${G - 15} Z`, fill: pal.ink }),
            SCENE_SILOS.map((sl, i) => (__spH("g", { key: `silo-${i}` },
                __spH("path", { d: `M${sl.x},${G} L${sl.x},${G - sl.h + 10} L${sl.x + sl.w / 2},${G - sl.h} L${sl.x + sl.w},${G - sl.h + 10} L${sl.x + sl.w},${G} Z`, fill: pal.ink }),
                __spH("path", { d: `M${sl.x},${G} L${sl.x},${G - sl.h + 10} L${sl.x + 6},${G - sl.h + 6} L${sl.x + 6},${G} Z`, fill: pal.mid })))),
            SCENE_HALLS.map((f, i) => {
                const tw = f.teeth ? f.w / f.teeth : 0;
                const roof = f.teeth
                    ? Array.from({ length: f.teeth }).map((_, k) => `L${f.x + k * tw},${G - f.h - 16} L${f.x + (k + 1) * tw},${G - f.h}`).join(' ')
                    : `L${f.x},${G - f.h} L${f.x + f.w},${G - f.h}`;
                return __spH("path", { key: `hall-${i}`, d: `M${f.x},${G} L${f.x},${G - f.h} ${roof} L${f.x + f.w},${G} Z`, fill: pal.ink });
            }),
            __spH("path", { d: `M372,${G - 30} L372,${G - 38} L386,${G - 38} L386,${G - 30} Z M440,${G - 30} L440,${G - 52} L447,${G - 52} L447,${G - 30} Z M436,${G - 54} L451,${G - 54} L451,${G - 51} L436,${G - 51} Z`, fill: pal.ink }),
            SCENE_TANKS.map((t, i) => (__spH("g", { key: `tank-${i}` },
                __spH("path", { d: `M${t.x},${G} L${t.x},${G - t.h + 8} Q${t.x + t.w / 2},${G - t.h - 8} ${t.x + t.w},${G - t.h + 8} L${t.x + t.w},${G} Z`, fill: pal.ink }),
                __spH("path", { d: `M${t.x},${G} L${t.x},${G - t.h + 8} Q${t.x + t.w * 0.14},${G - t.h + 2} ${t.x + t.w * 0.3},${G - t.h - 1} L${t.x + t.w * 0.3},${G} Z`, fill: pal.mid })))),
            SCENE_COLUMNS.map((c, i) => (__spH("g", { key: `col-${i}` },
                __spH("path", { d: `M${c.x},${G} L${c.x},${G - c.h + c.w / 2} A${c.w / 2},${c.w / 2} 0 0 1 ${c.x + c.w},${G - c.h + c.w / 2} L${c.x + c.w},${G} Z`, fill: pal.ink }),
                __spH("path", { d: `M${c.x},${G} L${c.x},${G - c.h + c.w / 2} Q${c.x + 1},${G - c.h + 3} ${c.x + 4},${G - c.h + 1} L${c.x + 4},${G} Z`, fill: pal.mid })))),
            stacks.map((st, i) => {
                const top = G - st.hh, ht = halfAt(st, 1);
                return (__spH("g", { key: `stack-${i}` },
                    __spH("path", { d: `M${st.x - st.w / 2 - 4},${G} L${st.x - st.w / 2 - 4},${G - 10} L${st.x + st.w / 2 + 4},${G - 10} L${st.x + st.w / 2 + 4},${G} Z`, fill: pal.ink }),
                    __spH("path", { d: `M${st.x - st.w / 2},${G} L${st.x - ht},${top} L${st.x + ht},${top} L${st.x + st.w / 2},${G} Z`, fill: pal.ink }),
                    __spH("path", { d: `M${st.x - st.w / 2},${G} L${st.x - ht},${top} L${st.x - ht * 0.3},${top} L${st.x - st.w * 0.16},${G} Z`, fill: pal.mid })));
            })),
        stacks.map((st, i) => {
            const top = G - st.hh, ht = halfAt(st, 1);
            return (__spH("g", { key: `stack-band-${i}` },
                [14, 34].map((dy) => {
                    const h0 = halfAt(st, 1 - dy / st.hh) + 0.8, h1 = halfAt(st, 1 - (dy + 9) / st.hh) + 0.8;
                    return __spH("path", { key: dy, d: `M${st.x - h0},${top + dy} L${st.x + h0},${top + dy} L${st.x + h1},${top + dy + 9} L${st.x - h1},${top + dy + 9} Z`, fill: pal.band });
                }),
                __spH("rect", { x: st.x - ht - 3.5, y: top - 2, width: ht * 2 + 7, height: "6", fill: pal.ink })));
        }),
        __spH("g", null,
            Array.from({ length: 9 }).map((_, k) => __spH("rect", { key: `wa-${k}`, x: 528 + k * 15, y: G - 34, width: "9", height: "7", fill: windowFill(k * 3 + 1) })),
            __spH("rect", { x: "532", y: G - 18, width: "22", height: "18", fill: pal.mid }),
            __spH("path", { d: `M543,${G - 18} L543,${G}`, stroke: pal.ink, strokeWidth: "1" }),
            Array.from({ length: 6 }).map((_, k) => __spH("rect", { key: `wb-${k}`, x: 362 + k * 14, y: G - 22, width: "7", height: "5", fill: windowFill(k * 5 + 2) })),
            __spH("rect", { x: "456", y: G - 14, width: "16", height: "14", fill: pal.mid })),
        __spH("g", { stroke: pal.ink, fill: "none", strokeLinecap: "round" },
            stacks.map((st, i) => {
                const topF = st.plat[st.plat.length - 1];
                return (__spH("g", { key: `st-line-${i}` },
                    st.plat.map((f) => {
                        const y = G - st.hh * f, hw = halfAt(st, f) + 9;
                        return (__spH("g", { key: f },
                            __spH("path", { d: `M${st.x - hw},${y} L${st.x + hw},${y}`, strokeWidth: "3" }),
                            __spH("path", { d: `M${st.x - hw},${y - 9} L${st.x + hw},${y - 9} M${st.x - hw},${y - 4.5} L${st.x + hw},${y - 4.5}`, strokeWidth: "0.9" }),
                            __spH("path", { d: Array.from({ length: Math.floor((hw * 2) / 7) + 1 }).map((_, k) => `M${st.x - hw + k * 7},${y} L${st.x - hw + k * 7},${y - 9}`).join(' '), strokeWidth: "0.9" }),
                            __spH("path", { d: `M${st.x - hw},${y + 2} L${st.x - hw + 9},${y + 13} M${st.x + hw},${y + 2} L${st.x + hw - 9},${y + 13}`, strokeWidth: "1.2" })));
                    }),
                    __spH("path", { d: `M${st.x + st.w / 2 + 2},${G - 10} L${st.x + halfAt(st, topF) + 2},${G - st.hh * topF}`, strokeWidth: "0.9" }),
                    !compact && __spH("path", { d: Array.from({ length: Math.floor((st.hh * topF) / 14) }).map((_, k) => { const t = (k * 14 + 16) / st.hh; return `M${st.x + halfAt(st, t) + 2},${G - st.hh * t} l4,0`; }).join(' '), strokeWidth: "0.8" }),
                    !compact && [0.2, 0.36].map((f) => __spH("path", { key: f, d: `M${st.x - halfAt(st, f)},${G - st.hh * f} L${st.x + halfAt(st, f)},${G - st.hh * f}`, stroke: pal.mid, strokeWidth: "1" })),
                    __spH("path", { d: `M${st.x + halfAt(st, 1) - 2},${G - st.hh} l0,-12`, strokeWidth: "0.9" })));
            }),
            SCENE_TANKS.map((t, i) => (__spH("g", { key: `tank-line-${i}` },
                __spH("path", { d: `M${t.x + 2},${G - t.h + 1} L${t.x + t.w - 2},${G - t.h + 1}`, strokeWidth: "0.9" }),
                !compact && __spH("path", { d: Array.from({ length: Math.floor(t.w / 8) }).map((_, k) => `M${t.x + 4 + k * 8},${G - t.h + 1} l0,${k < 2 || k > t.w / 8 - 3 ? 6 : 3}`).join(' '), strokeWidth: "0.7" }),
                __spH("path", { d: `M${t.x + t.w * 0.36},${G} L${t.x + t.w - 3},${G - t.h + 10}`, stroke: pal.mid, strokeWidth: "1.6" }),
                __spH("path", { d: `M${t.x + t.w * 0.3},${G - t.h * 0.45} L${t.x + t.w},${G - t.h * 0.45}`, stroke: pal.mid, strokeWidth: "0.8" })))),
            SCENE_COLUMNS.map((c, i) => (__spH("g", { key: `col-line-${i}` },
                [0.34, 0.62, 0.86].map((f) => (__spH("path", { key: f, d: `M${c.x - 5},${G - c.h * f} L${c.x + c.w + 5},${G - c.h * f} M${c.x - 5},${G - c.h * f - 6} L${c.x + c.w + 5},${G - c.h * f - 6} M${c.x - 5},${G - c.h * f} l0,-6 M${c.x + c.w + 5},${G - c.h * f} l0,-6`, strokeWidth: 1 }))),
                __spH("path", { d: `M${c.x + c.w / 2},${G - c.h} l0,-8 q0,-5 6,-5 l${c.w / 2 + 2},0 q5,0 5,5 L${c.x + c.w + 13},${G - c.h * 0.3}`, strokeWidth: "1.3" })))),
            __spH("path", { d: `M216,${G - 74.5} L286,${G - 74.5}`, strokeWidth: "2.2" }),
            __spH("path", { d: `M216,${G - 82} L286,${G - 82} ${[0, 1, 2, 3, 4, 5, 6, 7].map((k) => `M${216 + k * 10},${G - 74} l0,-8`).join(' ')}`, strokeWidth: "0.9" }),
            !compact && __spH("path", { d: `M211,${G} L211,${G - 76} L216,${G - 76} ${[0, 1, 2, 3, 4, 5, 6, 7].map((k) => `M211,${G - 8 - k * 9} l3,0`).join(' ')}`, strokeWidth: "0.8" }),
            !compact && __spH("path", { d: `M586,${G - 34} L690,${G - 34} M586,${G - 28} L690,${G - 28} ${[0, 1, 2, 3].map((k) => `M${596 + k * 28},${G - 34} L${596 + k * 28},${G}`).join(' ')}`, strokeWidth: "1.6" }),
            !compact && __spH("path", { d: `M880,${G - 22} L1010,${G - 22} M880,${G - 17} L1010,${G - 17} ${[0, 1, 2, 3, 4].map((k) => `M${888 + k * 28},${G - 22} l0,22`).join(' ')} M936,${G - 22} l0,-9 l14,0 l0,9`, strokeWidth: "1.3" }),
            !compact && __spH("path", { d: `M${SCENE_STACKS[0].x + 60},${G} L${SCENE_STACKS[0].x + halfAt(SCENE_STACKS[0], 0.82) + 6},${G - SCENE_STACKS[0].h * 0.82}`, strokeWidth: "0.8" }),
            __spH("path", { d: `M${SCENE_FLARE.x - 12},${G} L${SCENE_FLARE.x - 2},${flareTop} M${SCENE_FLARE.x + 12},${G} L${SCENE_FLARE.x + 2},${flareTop} ${Array.from({ length: compact ? 5 : 9 }).map((_, k, arr) => {
                    const n = arr.length, y0 = G - ((G - flareTop) * k) / n, y1 = G - ((G - flareTop) * (k + 1)) / n;
                    const w0 = 12 - (10 * k) / n, w1 = 12 - (10 * (k + 1)) / n;
                    return `M${SCENE_FLARE.x - w0},${y0} L${SCENE_FLARE.x + w1},${y1} M${SCENE_FLARE.x + w0},${y0} L${SCENE_FLARE.x - w1},${y1} M${SCENE_FLARE.x - w1},${y1} L${SCENE_FLARE.x + w1},${y1}`;
                }).join(' ')}`, strokeWidth: "1" }),
            __spH("path", { d: `M${SCENE_FLARE.x},${G} L${SCENE_FLARE.x},${flareTop - 8}`, strokeWidth: "2.2" }),
            !compact && __spH("path", { d: `M190,${G - 9} L1310,${G - 9} ${Array.from({ length: 57 }).map((_, k) => `M${190 + k * 20},${G} l0,-9`).join(' ')}`, strokeWidth: "0.7", opacity: "0.8" }),
            !compact && [498, 804, 1000, 1306].map((x) => __spH("path", { key: x, d: `M${x},${G} l0,-26 q0,-4 5,-4 l4,0`, strokeWidth: "1.1" }))),
        __spH("path", { className: "skyline-flare", style: { transformOrigin: `${SCENE_FLARE.x}px ${flareTop - 8}px` }, d: `M${SCENE_FLARE.x},${flareTop - 8} q-8,-10 0,-24 q2,8 6,11 q4,6 -6,13 Z`, fill: night ? pal.light : pal.band }),
        __spH("path", { className: "skyline-flare-core", style: { transformOrigin: `${SCENE_FLARE.x}px ${flareTop - 8}px` }, d: `M${SCENE_FLARE.x + 0.5},${flareTop - 8} q-3.5,-5 0,-11 q1,3 2.6,5 q1.6,3 -2.6,6 Z`, fill: pal.plume }),
        night && (__spH("g", { fill: pal.light },
            [[1024, 30], [1060, 30], [1132, 30], [1190, 30], [1240, 22], [844, 60], [868, 44], [968, 52], [1416, 16], [1500, 12], [70, 10], [224, 40], [276, 40]].map(([x, dy], i) => (__spH("rect", { key: `win-${i}`, x: x, y: G - dy, width: "4", height: "4" }))),
            !compact && [498, 804, 1000, 1306].map((x) => __spH("rect", { key: `lamp-${x}`, x: x + 6, y: G - 31, width: "5", height: "3" })),
            stacks.map((st, i) => st.plat.map((f) => __spH("rect", { key: `pl-${i}-${f}`, x: st.x - halfAt(st, f) - 9, y: G - st.hh * f - 14, width: "3", height: "3" }))))),
        pal.beacon && stacks.map((st, i) => (__spH("circle", { key: `beacon-${i}`, className: "skyline-beacon", cx: st.x, cy: G - st.hh - 6, r: "2.6", fill: pal.beacon, style: { '--beacon-delay': `-${i * 0.7}s` } }))),
        __spH("rect", { x: "0", y: G, width: "1600", height: "70", fill: pal.water }),
        __spH("g", { fill: pal.ink, opacity: "0.3" },
            stacks.map((st, i) => __spH("path", { key: `refl-${i}`, d: `M${st.x - st.w / 2},${G} L${st.x + st.w / 2},${G} L${st.x + st.w * 0.3},${G + Math.min(40, st.hh * 0.11)} L${st.x - st.w * 0.3},${G + Math.min(40, st.hh * 0.11)} Z` })),
            SCENE_TANKS.map((t, i) => __spH("path", { key: `reflt-${i}`, d: `M${t.x},${G} L${t.x + t.w},${G} L${t.x + t.w - 6},${G + 9} L${t.x + 6},${G + 9} Z` })),
            SCENE_HALLS.map((f, i) => __spH("path", { key: `reflh-${i}`, d: `M${f.x},${G} L${f.x + f.w},${G} L${f.x + f.w - 5},${G + 7} L${f.x + 5},${G + 7} Z` }))),
        __spH("g", { stroke: pal.stroke, strokeLinecap: "round", fill: "none" },
            __spH("path", { d: `M120,${G + 9} L330,${G + 9} M560,${G + 15} L860,${G + 15} M980,${G + 8} L1190,${G + 8} M1280,${G + 18} L1470,${G + 18}`, strokeWidth: "2.2" }),
            __spH("path", { d: `M640,${G + 23} L760,${G + 23} M260,${G + 24} L420,${G + 24} M1040,${G + 27} L1150,${G + 27} M420,${G + 5} L500,${G + 5} M1350,${G + 6} L1420,${G + 6}`, strokeWidth: "1.2" })),
        [
            { y: G + 21, dur: 170, delay: -60, dir: 1, sizes: warm ? [1, 0.55, 0.55, 0.5, 0.55] : [1, 0.9, 0.86] },
            { y: G + 12, dur: 210, delay: -150, dir: -1, sizes: [0.8, 0.8] },
        ].slice(0, compact ? 1 : 2).map((flock, fi) => (__spH("g", { key: `ducks-${fi}` },
            __spH("g", { transform: `translate(0 ${flock.y}) scale(${flock.dir} 1)` }, flock.sizes.map((sz, k) => {
                const x = -flock.sizes.slice(0, k).reduce((sum, v) => sum + 14 + v * 22, 0);
                return (__spH("g", { key: k, transform: `translate(${x} 0) scale(${sz})` },
                    __spH("path", { d: "M-34,3.5 L-17,3.5 M-28,0.5 L-18,1.5", stroke: pal.stroke, strokeWidth: 1.4 / sz, strokeLinecap: "round", fill: "none" }),
                    __spH("path", { d: "M-12,0 L-16,-6 L-9,-2.5 Q-2,-4.5 4,-2.5 L4,-9 a4.3,4.3 0 1 1 6.6,3.6 L10.6,-2 Q12,3.5 6,4 L-8,4 Q-12,3.2 -12,0 Z", fill: pal.bank }),
                    __spH("path", { d: "M11.4,-10 l5.6,1.2 l-5.4,1.8 Z", fill: pal.band })));
            })),
            __spH("animateMotion", { dur: `${flock.dur}s`, begin: `${flock.delay}s`, repeatCount: "indefinite", path: "M-140,0 L1740,0", keyPoints: flock.dir > 0 ? '0;1' : '1;0', keyTimes: "0;1", calcMode: "linear" })))),
        __spH("path", { d: tufts.filter((t) => t.back).map((t, i) => sceneTuft(t.x, t.size * 1.25, i * 17 + 5)).join(' '), fill: pal.ink, transform: "translate(0 -3)" }),
        __spH("g", { filter: cut },
            __spH("path", { d: `M${SCENE_BANK.map(([x, dy]) => `${x},${G + dy}`).join(' L')} L1600,580 L0,580 Z`, fill: pal.bank })),
        [
            { key: 'raccoon', show: true, dur: 168, begin: compact ? -70 : -9, scale: 0.62, trot: 0.34, legH: 15, legs: [[-13, 'a'], [-6.6, 'b'], [2.4, 'b'], [8.4, 'a']], art: renderSceneRaccoon, stillX: 1010 },
            { key: 'badger', show: season !== 'winter' && !compact, dur: 214, begin: -150, scale: 0.6, trot: 0.46, legH: 13, legs: [[-17, 'a'], [-10.6, 'b'], [4.4, 'b'], [10.6, 'a']], art: renderSceneBadger, stillX: 560 },
        ].filter((walker) => walker.show).map((walker) => {
            const timing = { dur: `${walker.dur}s`, begin: `${walker.begin}s`, repeatCount: 'indefinite' };
            const figure = (__spH("g", { style: { '--trot': `${walker.trot}s` } },
                walker.legs.map(([x, pair]) => (__spH("path", { key: x, className: stillScene ? undefined : `skyline-trot-${pair}`, d: sceneWalkerLeg(x, walker.legH), fill: pal.bank }))),
                __spH("g", { className: stillScene ? undefined : 'skyline-trot-bob' },
                    __spH("g", { transform: "translate(0 -3.4)" }, walker.art(pal)))));
            if (stillScene)
                return __spH("g", { key: walker.key, transform: `translate(${walker.stillX} ${sceneBankY(walker.stillX) + 2}) scale(-${walker.scale} ${walker.scale})` }, figure);
            return (__spH("g", { key: walker.key },
                __spH("animateMotion", { ...timing, calcMode: "linear", path: SCENE_WALK_TRACK.path, keyPoints: SCENE_WALK_TRACK.keyPoints, keyTimes: SCENE_WALK_TRACK.keyTimes }),
                __spH("animate", { attributeName: "opacity", calcMode: "discrete", values: "0;1;0;1;0", keyTimes: "0;0.04;0.4;0.54;0.92", ...timing }),
                __spH("g", { transform: `scale(${walker.scale})` },
                    __spH("g", null,
                        __spH("animateTransform", { attributeName: "transform", type: "scale", calcMode: "discrete", values: "-1 1;1 1", keyTimes: `0;${SCENE_WALK_TRACK.turn}`, ...timing }),
                        figure))));
        }),
        warm && SCENE_FROG_TRACKS.slice(0, compact ? 1 : SCENE_FROG_TRACKS.length).map((frog, i) => {
            const timing = { dur: `${frog.dur.toFixed(2)}s`, begin: `-${(frog.dur * frog.phase).toFixed(2)}s`, repeatCount: 'indefinite' };
            const sitting = (__spH("g", null,
                __spH("g", { transform: "translate(11.5 -7.5)" },
                    __spH("circle", { className: "skyline-frog-sac", cx: "3.6", cy: "1.2", r: "5.2", fill: pal.plume, style: { '--frog-duration': `${frog.croak}s`, '--frog-delay': `-${i * 3}s` } })),
                __spH("path", { d: SCENE_FROG_SIT, fill: pal.bank }),
                __spH("path", { d: "M-13,-1.5 Q-10,-10 -3,-7.5 Q1,-6 1,-4", stroke: pal.water, strokeWidth: "0.9", strokeLinecap: "round", fill: "none", opacity: "0.55" }),
                __spH("circle", { cx: "6.8", cy: "-19", r: "1.5", fill: pal.water })));
            const leaping = (__spH(__spFrag, null,
                __spH("path", { d: SCENE_FROG_LEAP, fill: pal.bank }),
                __spH("circle", { cx: "8.6", cy: "-14.4", r: "1.5", fill: pal.water })));
            if (stillScene)
                return __spH("g", { key: `frog-${i}`, transform: `translate(${i === 0 ? 692 : 1186} ${sceneBankY(i === 0 ? 692 : 1186) + 2}) scale(${frog.size})` }, sitting);
            return (__spH("g", { key: `frog-${i}` },
                __spH("animateMotion", { ...timing, calcMode: "linear", path: frog.path, keyPoints: frog.motion.keyPoints, keyTimes: frog.motion.keyTimes }),
                __spH("g", { transform: `scale(${frog.size})` },
                    __spH("g", null,
                        __spH("animateTransform", { attributeName: "transform", type: "scale", calcMode: "discrete", values: "1 1;-1 1", keyTimes: `0;${frog.turn}`, ...timing }),
                        __spH("g", null,
                            sitting,
                            __spH("animate", { attributeName: "opacity", calcMode: "discrete", values: frog.sit.values, keyTimes: frog.sit.keyTimes, ...timing })),
                        __spH("g", { opacity: "0", transform: "rotate(-26)" },
                            leaping,
                            __spH("animate", { attributeName: "opacity", calcMode: "discrete", values: frog.up.values, keyTimes: frog.up.keyTimes, ...timing })),
                        __spH("g", { opacity: "0", transform: "rotate(22)" },
                            leaping,
                            __spH("animate", { attributeName: "opacity", calcMode: "discrete", values: frog.down.values, keyTimes: frog.down.keyTimes, ...timing }))))));
        }),
        __spH("g", { fill: pal.bank },
            __spH("path", { d: tufts.filter((t) => !t.back).map((t, i) => sceneTuft(t.x, t.size, i * 13 + 3)).join(' ') }),
            !compact && [[468, 9, 4], [486, 6, 3], [826, 10, 4], [1108, 8, 3.5], [1126, 5, 2.5]].map(([x, rx, ry]) => (__spH("ellipse", { key: `stone-${x}`, cx: x, cy: sceneBankY(x) + 1, rx: rx, ry: ry })))),
        __spH("g", { stroke: pal.bank, fill: "none", strokeLinecap: "round" },
            cattails.map((c, i) => {
                const y = sceneBankY(c.x) + 4, tx = c.x + c.lean, ty = y - c.h;
                return (__spH("g", { key: `cattail-${i}` },
                    __spH("path", { d: `M${c.x},${y} Q${c.x + c.lean * 0.2},${y - c.h * 0.6} ${tx},${ty}`, strokeWidth: "1.4" }),
                    __spH("path", { d: `M${c.x + c.lean * 0.62},${y - c.h * 0.84} L${c.x + c.lean * 0.86},${y - c.h * 0.95}`, strokeWidth: "4.6" })));
            }),
            !compact && [[150, 44], [176, 36], [1402, 40], [330, 26]].map(([x, h], i) => {
                const y = sceneBankY(x) + 4;
                return (__spH("g", { key: `umbel-${i}` },
                    __spH("path", { d: `M${x},${y} L${x + 3},${y - h} M${x + 3},${y - h} l-6,-7 M${x + 3},${y - h} l1,-9 M${x + 3},${y - h} l7,-6`, strokeWidth: "1" }),
                    __spH("path", { d: `M${x - 3},${y - h - 7} l0,0 M${x + 4},${y - h - 9} l0,0 M${x + 10},${y - h - 6} l0,0`, strokeWidth: "3.4" })));
            })),
        season === 'spring' && SCENE_BUTTERFLIES.slice(0, compact ? 2 : SCENE_BUTTERFLIES.length).map((bf, i) => (__spH("g", { key: `butterfly-${i}`, transform: `translate(${bf.x} ${sceneBankY(bf.x) - bf.up})` },
            __spH("g", { className: "skyline-flutter", style: { '--flutter-duration': `${bf.dur}s`, '--flutter-delay': `-${bf.delay}s` } },
                __spH("g", { className: "skyline-wingbeat", style: { '--beat': `${bf.beat}s` } },
                    __spH("path", { d: "M0,0 Q-6,-10 -9,-4 Q-10,2 -1,1 Q-6,6 -2,6 Q0,4 0,1 Q0,4 2,6 Q6,6 1,1 Q10,2 9,-4 Q6,-10 0,0 Z", fill: pal.plume, stroke: pal.bank, strokeWidth: "0.5" })),
                __spH("path", { d: "M0,-3 L0,4", stroke: pal.bank, strokeWidth: "1.1", strokeLinecap: "round" }))))),
        (season === 'summer' || season === 'autumn') && SCENE_DRAGONFLIES.slice(0, compact ? 1 : SCENE_DRAGONFLIES.length).map((df, i) => (__spH("g", { key: `dragonfly-${i}`, transform: `translate(${df.x} ${sceneBankY(df.x) - df.up}) scale(${df.dir} 1)` },
            __spH("g", { className: "skyline-dart", style: { '--dart-duration': `${df.dur}s`, '--dart-delay': `-${df.delay}s` } },
                __spH("path", { className: "skyline-shimmer", d: "M1,-0.6 L-6,-8.5 L-2.5,-9 Z M2.5,-0.6 L6,-9.5 L9.5,-8.5 Z", fill: pal.plume }),
                __spH("path", { d: "M-4,0 L12,0.6", stroke: pal.band, strokeWidth: "1.5", strokeLinecap: "round" }),
                __spH("circle", { cx: "-4.4", cy: "-0.2", r: "1.7", fill: pal.band })))))));
};
const spSeasonByDate = (date) => { const m = (date || new Date()).getMonth() + 1; return m >= 3 && m <= 5 ? 'spring' : m >= 6 && m <= 8 ? 'summer' : m >= 9 && m <= 11 ? 'autumn' : 'winter'; };
const renderScene = (themeId, time, scope, opts) => {
    const o = opts || {};
    const skySeason = ['spring', 'summer', 'autumn', 'winter'].includes(o.season) ? o.season : spSeasonByDate();
    const skyTheme = getSceneTheme(themeId, time);
    const isNightSky = Boolean(skyTheme.stars);
    const isMorningSky = skyTheme.orb === 'bigsun';
    const skyOrbClass = skyTheme.orb === 'moon' ? 'skyline-moon' : (skyTheme.orb === 'sun' ? 'skyline-sun' : '');
    const skyOrbStyle = { background: skyTheme.orbColor };
    const clouds = SKYLINE_CLOUDS.slice(0, scope === 'menu' ? 4 : 5);
    const stars = scope === 'menu' ? SKYLINE_STARS : SKYLINE_STARS.slice(0, 30);
    return (__spH("div", { className: 'sp-scene sp-scene-' + scope, style: { position: 'absolute', inset: 0, overflow: 'hidden', background: 'linear-gradient(to bottom, ' + skyTheme.sky[0] + ', ' + skyTheme.sky[1] + ')' } },
        __spH("div", { style: { position: 'absolute', inset: 0, opacity: scope === 'menu' ? 1 : (o.opacity == null ? 0.45 : o.opacity) } },
            skyOrbClass && __spH("div", { className: skyOrbClass, style: skyOrbStyle }),
            __spH("div", { className: "skyline-grain", style: { '--grain-opacity': skyTheme.grain } }),
            isMorningSky && __spH("div", { className: scope === 'menu' ? 'skyline-sunrise-mountain' : 'skyline-sunrise-mountain-sheet', style: skyOrbStyle }),
            scope === 'menu' && skySeason === 'autumn' && typeof SCENE_GEESE !== 'undefined' && (__spH("div", { className: "skyline-geese" }, SCENE_GEESE.map((goose, i) => (__spH("div", { className: "skyline-crane", style: { left: goose.left, top: goose.top, width: '15%', '--crane-delay': goose.delay + 's' } }, renderSceneGoose(skyTheme.pal)))))),
            scope === 'menu' && skySeason === 'winter' && (__spH("div", { className: "skyline-cranes" }, SCENE_CRANES.map((crane, i) => (__spH("div", { className: "skyline-crane", style: { left: crane.left, top: crane.top, width: crane.width, '--crane-delay': crane.delay + 's' } }, renderSceneCrane(skyTheme.pal)))))),
            clouds.map((cloud) => (__spH("div", { className: "skyline-cloud", style: { top: cloud.top, opacity: cloud.opacity, '--cloud-scale': cloud.scale, '--cloud-duration': cloud.duration + 's', '--cloud-delay': cloud.delay + 's' } },
                __spH("svg", { viewBox: "0 0 200 56", "aria-hidden": "true" },
                    __spH("path", { d: "M14,50 C1,50 2,36 16,35 C15,22 36,15 50,24 C56,7 86,3 100,18 C112,7 139,10 142,26 C158,19 177,28 172,38 C189,36 199,50 184,50 Z", fill: skyTheme.cloud[0] }),
                    __spH("path", { d: "M44,50 C66,43 92,47 120,42 C148,37 170,43 190,46 C191,48 188,50 184,50 Z", fill: skyTheme.cloud[1] }))))),
            isNightSky && (__spH("div", { style: { position: 'absolute', inset: 0 } }, stars.map((star) => (__spH("span", { className: "skyline-star", style: { left: star.left, top: star.top, width: star.size + 'px', height: star.size + 'px', '--twinkle-duration': star.duration + 's', '--twinkle-delay': '-' + star.delay + 's' } }))))),
            __spH("div", { className: "sp-scene-land", style: { position: 'absolute', bottom: 0, left: 0, right: 0, height: scope === 'menu' ? 'min(32rem, 62vh)' : 'min(19rem, 42vh)' } }, renderSkylineSvg({ pal: skyTheme.pal, scope: scope === 'menu' ? 'menu' : 'sheet', compact: scope !== 'menu', season: skySeason })))));
};

  var EXTRA_CSS = '.sp-scene svg{display:block}.sp-scene .sp-scene-land>svg{width:100%;height:100%}';
  function ensureCss(){ if (document.getElementById('sp-scene-css')) return; var st=document.createElement('style'); st.id='sp-scene-css'; st.textContent = SKYLINE_FLYER_CSS + EXTRA_CSS; document.head.appendChild(st); }
  function html(theme, time, scope, opts){ return renderScene(sanitizeSkyTheme(theme), time, scope || 'sheet', opts).__html; }
  root.StackpilotScene = {
    themes: SKY_THEME_OPTIONS, phase: function(d){ return getSkyPhaseByHour(d || new Date()); },
    theme: function(theme, time){ return getSceneTheme(sanitizeSkyTheme(theme), time); },
    html: html, css: function(){ return SKYLINE_FLYER_CSS + EXTRA_CSS; },
    mount: function(el, o){ o = o || {}; ensureCss(); var st = { theme: sanitizeSkyTheme(o.theme), mode: o.mode || 'auto', scope: o.scope || 'sheet', opts: o };
      function time(){ return st.mode === 'auto' ? getSkyPhaseByHour(new Date()) : st.mode; }
      var last = '';
      function draw(force){ var key = st.theme + '|' + time() + '|' + st.scope + '|' + (st.opts.season || 'auto') + '|' + spSeasonByDate(); if (!force && key === last) return; last = key; el.innerHTML = html(st.theme, time(), st.scope, st.opts); el.dataset.spTime = time(); el.dataset.spTheme = st.theme; }
      draw(true); var timer = setInterval(function(){ draw(false); }, 60000);
      return { setTheme: function(t){ st.theme = sanitizeSkyTheme(t); draw(true); }, setMode: function(m){ st.mode = m || 'auto'; draw(true); }, setSeason: function(x){ st.opts.season = x || 'auto'; draw(true); }, getSeason: function(){ return st.opts.season || 'auto'; }, getMode: function(){ return st.mode; }, getTheme: function(){ return st.theme; }, time: time, info: function(){ return getSceneTheme(st.theme, time()); }, destroy: function(){ clearInterval(timer); el.innerHTML=''; } }; }
  };
})(typeof window !== 'undefined' ? window : globalThis);

