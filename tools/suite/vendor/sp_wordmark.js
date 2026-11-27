/* 글자 로고 (시료채취 프로그램 STACKPILOT 로고와 같은 꾸밈): 굵은 글자 + 엇갈린 색 그림자 층 + 손으로 오린 듯한 가장자리. 색은 풍경(StackpilotScene)의 시간대 글자 로고 색을 따라간다.
   SPWordmark.mount(요소, '제목', { scene: StackpilotScene.mount(...)의 결과 }) */
(function () {
  var BASIC = { ink: '#2c4150', accent: '#c4523a' };
  var FONT = '"Pretendard","IBM Plex Sans KR","Apple SD Gothic Neo","맑은 고딕","Malgun Gothic",sans-serif';
  var n = 0;
  function esc(s) { return String(s).replace(/[&<>"]/g, function (c) { return { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]; }); }
  function svg(text, c, id) {
    var size = 72, base = 84;
    var tx = 6;
    var word = function (col) {
      return '<text x="' + tx + '" y="' + (base - 4) + '" font-family=\'' + FONT + '\' font-size="' + size + '" font-weight="900" letter-spacing="-1.5" ' +
        'fill="' + col + '" stroke="' + col + '" stroke-width="2.6" stroke-linejoin="miter" paint-order="stroke">' + esc(text) + '</text>';
    };
    return '<svg class="sp-wordmark" role="img" aria-label="' + esc(text) + '" viewBox="0 4 ' + (tx + text.length * size * 0.98 + 10) + ' 92" preserveAspectRatio="xMidYMid meet">' +
      '<defs><filter id="' + id + '" x="-2%" y="-5%" width="104%" height="110%"><feTurbulence type="fractalNoise" baseFrequency="0.03" numOctaves="2" seed="4" result="n"/>' +
      '<feDisplacementMap in="SourceGraphic" in2="n" scale="1.6" xChannelSelector="R" yChannelSelector="G"/></filter></defs>' +
      '<g filter="url(#' + id + ')"><g transform="translate(3.5 3.5)">' + word(c.accent) + '</g>' + word(c.ink) + '</g>' +
      '</svg>';
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
