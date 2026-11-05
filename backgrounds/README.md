# StackPilot 배경 (코드)

React 없이 어떤 웹페이지에서든 쓸 수 있는 배경 스크립트입니다.

- `stackpilot-skyline.js` : 배경 전체(하늘 그라데이션, 해/달, 별, 구름, 공장 풍경, 연기·창문 애니메이션)를 그려 주는 단독 스크립트
- `stackpilot-background.html` : 사용 예시 (버튼으로 낮/아침/밤 전환)
- `stackpilot-background-*.png`, `stackpilot-skyline-*.svg` : 같은 배경의 이미지 버전

## 사용법

```html
<div id="bg" style="position:fixed;inset:0"></div>
<script src="stackpilot-skyline.js"></script>
<script>
  var sky = StackpilotSkyline.mount(document.getElementById('bg'), { mode: 'auto' });
  // sky.setMode('night');  // 'day' | 'morning' | 'night' | 'auto'
</script>
```

배경 위에 올릴 내용은 `position: relative; z-index: 1` 로 두면 됩니다.
색을 바꾸려면 `stackpilot-skyline.js` 안의 `SKY`(하늘)와 `SKYLINE_PALETTES`(풍경) 값을 고치면 됩니다.
