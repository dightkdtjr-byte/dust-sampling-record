/* 글자 로고 (시료채취 프로그램 STACKPILOT 로고와 같은 꾸밈): 굵은 글자 + 엇갈린 색 그림자 층 + 손으로 오린 듯한 가장자리,
   앞에 연기 나는 굴뚝. 색은 풍경(StackpilotScene)의 시간대 글자 로고 색을 따라간다.
   SPWordmark.mount(요소, '제목', { scene: StackpilotScene.mount(...)의 결과 }) */
(function () {
  var BASIC = { ink: '#2c4150', mid: '#4f6875', accent: '#c4523a', band: '#f7f2e5', plume: '#f7f2e5', plumeShade: '#dcd9ca' };
  var FONT = '"Pretendard","IBM Plex Sans KR","Apple SD Gothic Neo","맑은 고딕","Malgun Gothic",sans-serif';
  var n = 0;
  function esc(s) { return String(s).replace(/[&<>"]/g, function (c) { return { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]; }); }
  function svg(text, c, id) {
    var size = 72, base = 84, ix = 30, top = -8;
    var p = function (dx, dy) { return (ix + dx * 0.36).toFixed(1) + ',' + (top - 1 - dy * 0.36).toFixed(1); };
    var plume = 'M' + (ix - 7) + ',' + (top + 2) + ' L' + (ix - 7) + ',' + (top - 1) + ' C' + p(-12, 30) + ' ' + p(8, 60) + ' ' + p(44, 62) +
      ' C' + p(50, 90) + ' ' + p(98, 100) + ' ' + p(122, 78) + ' C' + p(152, 106) + ' ' + p(216, 106) + ' ' + p(230, 72) +
      ' C' + p(266, 86) + ' ' + p(306, 64) + ' ' + p(300, 42) + ' C' + p(252, 28) + ' ' + p(192, 40) + ' ' + p(150, 33) +
      ' C' + p(100, 25) + ' ' + p(40, 34) + ' ' + (ix + 7) + ',' + (top - 1) + ' L' + (ix + 7) + ',' + (top + 2) + ' Z';
    var shade = 'M' + p(40, 44) + ' C' + p(90, 52) + ' ' + p(120, 44) + ' ' + p(150, 52) + ' C' + p(200, 64) + ' ' + p(250, 52) + ' ' + p(300, 42) +
      ' C' + p(252, 28) + ' ' + p(192, 40) + ' ' + p(150, 33) + ' C' + p(100, 25) + ' ' + p(60, 32) + ' ' + p(40, 44) + ' Z';
    var pillar = 'M' + (ix - 16) + ',' + base + ' L' + (ix - 12) + ',' + top + ' L' + (ix + 12) + ',' + top + ' L' + (ix + 16) + ',' + base + ' Z';
    var tx = ix + 16 + 16;
    var word = function (col) {
      return '<path d="' + pillar + '" fill="' + col + '"/>' +
        '<text x="' + tx + '" y="' + (base - 4) + '" font-family=\'' + FONT + '\' font-size="' + size + '" font-weight="900" letter-spacing="-1.5" ' +
        'fill="' + col + '" stroke="' + col + '" stroke-width="2.6" stroke-linejoin="miter" paint-order="stroke">' + esc(text) + '</text>';
    };
    return '<svg class="sp-wordmark" role="img" aria-label="' + esc(text) + '" viewBox="0 -48 ' + (tx + text.length * size * 0.98 + 10) + ' 142" preserveAspectRatio="xMidYMid meet">' +
      '<defs><filter id="' + id + '" x="-2%" y="-5%" width="104%" height="110%"><feTurbulence type="fractalNoise" baseFrequency="0.03" numOctaves="2" seed="4" result="n"/>' +
      '<feDisplacementMap in="SourceGraphic" in2="n" scale="1.6" xChannelSelector="R" yChannelSelector="G"/></filter></defs>' +
      '<path d="' + plume + '" fill="' + c.plume + '"/><path d="' + shade + '" fill="' + c.plumeShade + '"/>' +
      '<g filter="url(#' + id + ')"><g transform="translate(3.5 3.5)">' + word(c.accent) + '</g>' + word(c.ink) + '</g>' +
      '<path d="M' + (ix - 12.3) + ',' + (top + 14) + ' h24.6 l0.4,10 h-25.4 Z M' + (ix - 13.4) + ',' + (top + 34) + ' h26.8 l0.4,10 h-27.6 Z" fill="' + c.band + '"/>' +
      '<rect x="' + (ix - 16) + '" y="' + (top - 5) + '" width="32" height="9" fill="' + c.ink + '"/></svg>';
  }
  function fit(host) {   // 글자 실제 너비로 viewBox를 맞춘다 (보이지 않을 때는 어림값 그대로)
    var s = host.querySelector('svg'), t = s && s.querySelector('text:last-of-type');
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
