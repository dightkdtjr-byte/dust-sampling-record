/* 글자 로고 (시료채취 프로그램 STACKPILOT 로고와 같은 꾸밈): 굵은 글자 + 엇갈린 색 그림자 층 + 손으로 오린 듯한 가장자리. 색은 풍경(StackpilotScene)의 시간대 글자 로고 색을 따라간다.
   SPWordmark.mount(요소, '제목', { scene: StackpilotScene.mount(...)의 결과 }) */
(function () {
  var BASIC = { ink: '#2c4150', accent: '#c4523a' };
  var FONT = '"Pretendard","IBM Plex Sans KR","Apple SD Gothic Neo","맑은 고딕","Malgun Gothic",sans-serif';
  var n = 0;
  function esc(s) { return String(s).replace(/[&<>"]/g, function (c) { return { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]; }); }

  // ── 글자 그림 (STACKPILOT 로고와 같은 규칙: 획 굵기 13, 각진 이음, 높이 0~53.5) ──
  // S·T·A·C·K·P·L·O는 시료채취 프로그램 로고 그대로, B·V·U는 같은 규칙으로 그림
  var G = {
    S: 'M38,0 H14 a12,12 0 0 0 0,24 H26 a12,12 0 0 1 0,24 H2',
    T: 'M0,0 H40 M20,0 V53.5',
    A: 'M3,53.5 L20,-2 L37,53.5 M10,34 H30',
    C: 'M38,0 H24 a24,24 0 0 0 0,48 H38',
    K: 'M5.5,-5.5 V53.5 M39,-5.5 L8,27 M18,18 L39,53.5',
    P: 'M5.5,53.5 V0 H25 a13.5,13.5 0 0 1 0,27 H5.5',
    L: 'M5.5,-5.5 V48 H38',
    O: 'M17,0 H23 a15,15 0 0 1 15,15 V33 a15,15 0 0 1 -15,15 H17 a15,15 0 0 1 -15,-15 V15 a15,15 0 0 1 15,-15 Z',
    B: 'M5.5,-6.5 V54.5 M5.5,0 H23 a12,12 0 0 1 0,24 H5.5 M5.5,24 H25.5 a12,12 0 0 1 0,24 H5.5',
    V: 'M3,-2 L20,53.5 L37,-2',
    U: 'M5.5,-5.5 V31 a14.5,17 0 0 0 29,0 V-5.5',
    I: 'M20,-5.5 V53.5',
  };
  var EXT = { S: [-4.5, 44.5], T: [0, 40], A: [-3, 43], C: [-6.5, 38], K: [-1, 45], P: [-1, 45], I: [13.5, 26.5], L: [-1, 38], O: [-4.5, 44.5], B: [-1, 44], V: [-3, 43], U: [-1, 41] };
  // 짝마다 눈에 보이는 틈이 같도록 당김 (한쪽이 트이거나 기운 글자 옆은 좁히고, 곧은 획끼리는 넓힘)
  var KERN = { ST: 3, TA: 9, AC: 3, CK: 5, KP: 4, PI: -2, IL: -6, LO: 5, OT: 2, LA: 10, AB: 3, BV: 3, VA: 12, AU: 3, UL: -3, LT: 10 };
  function glyphSvg(text, c, id) {
    var letters = text.split(''), y0 = 30, gap = 21, offs = [];
    var pen = 6;
    letters.forEach(function (ch, i) { var k = i > 0 ? (KERN[letters[i - 1] + ch] || 0) : 0; offs.push(pen - k - EXT[ch][0]); pen = pen - k + (EXT[ch][1] - EXT[ch][0]) + gap; });
    var last = letters.length - 1, total = offs[last] + EXT[letters[last]][1] + 6;
    var layer = function (col) {
      return '<g stroke="' + col + '" stroke-width="13" fill="none" stroke-linejoin="miter" stroke-miterlimit="2">' +
        letters.map(function (ch, i) { return '<path d="' + G[ch] + '" transform="translate(' + offs[i].toFixed(2) + ' ' + y0 + ')"/>'; }).join('') + '</g>';
    };
    return '<svg class="sp-wordmark" role="img" aria-label="' + esc(text) + '" viewBox="0 -24 ' + (total + 3).toFixed(1) + ' 112">' +
      '<defs><filter id="' + id + '" x="-2%" y="-5%" width="104%" height="110%"><feTurbulence type="fractalNoise" baseFrequency="0.03" numOctaves="2" seed="4" result="n"/>' +
      '<feDisplacementMap in="SourceGraphic" in2="n" scale="1.6" xChannelSelector="R" yChannelSelector="G"/></filter></defs>' +
      '<g filter="url(#' + id + ')"><g transform="translate(3 3)">' + layer(c.accent) + '</g>' + layer(c.ink) + '</g></svg>';
  }
  var canGlyph = function (t) { return /^[A-Z]+$/.test(t) && t.split('').every(function (ch) { return G[ch]; }); };
  function svg(text, c, id) {
    if (canGlyph(text)) return glyphSvg(text, c, id);
    var size = 72, base = 84;
    var tx = 6;
    var word = function (col) {
      return '<text x="' + tx + '" y="' + (base - 4) + '" font-family=\'' + FONT + '\' font-size="' + size + '" font-weight="900" letter-spacing="' + (/^[\x00-\x7f]+$/.test(text) ? 4 : -1.5) + '" ' +
        'fill="' + col + '" stroke="' + col + '" stroke-width="2.6" stroke-linejoin="miter" paint-order="stroke">' + esc(text) + '</text>';
    };
    return '<svg class="sp-wordmark" role="img" aria-label="' + esc(text) + '" viewBox="0 4 ' + (tx + text.length * size * 0.98 + 10) + ' 92" preserveAspectRatio="xMidYMid meet">' +
      '<defs><filter id="' + id + '" x="-2%" y="-5%" width="104%" height="110%"><feTurbulence type="fractalNoise" baseFrequency="0.03" numOctaves="2" seed="4" result="n"/>' +
      '<feDisplacementMap in="SourceGraphic" in2="n" scale="1.6" xChannelSelector="R" yChannelSelector="G"/></filter></defs>' +
      '<g filter="url(#' + id + ')"><g transform="translate(3.5 3.5)">' + word(c.accent) + '</g>' + word(c.ink) + '</g>' +
      '</svg>';
  }
  function fit(host) {   // 글자 실제 너비로 viewBox를 맞춘다 (보이지 않을 때는 어림값 그대로)
    var s = host.querySelector('svg'), t = s && s.querySelector('text:last-of-type');   // 글자 그림 로고는 맞출 필요 없음
    if (!t || !t.getBBox) return;
    try { var b = t.getBBox(); if (b.width > 0) { var vb = s.viewBox.baseVal; s.setAttribute('viewBox', '0 ' + vb.y + ' ' + Math.ceil(b.x + b.width + 8) + ' ' + vb.height); } } catch (e) {}
  }
  function mount(host, text, o) {
    o = o || {}; var id = 'spwm' + (++n), last = '';
    function colors() { try { var sc = typeof o.scene === 'function' ? o.scene() : o.scene, i = sc && sc.info && sc.info(); if (i && i.wm) return i.wm; } catch (e) {} return BASIC; }
    function draw() { var c = colors(), key = JSON.stringify(c) + text; if (key === last) return; last = key; host.innerHTML = svg(text, c, id); fit(host); }
    draw(); var timer = setInterval(draw, 30000);
    if (window.ResizeObserver) new ResizeObserver(function () { fit(host); }).observe(host);
    return { redraw: function () { last = ''; draw(); }, destroy: function () { clearInterval(timer); } };
  }
  window.SPWordmark = { mount: mount, svg: svg };
})();
