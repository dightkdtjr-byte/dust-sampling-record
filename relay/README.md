# 모바일 날씨 중계 설정 (Cloudflare Workers, 무료)

모바일(사파리)에서 "기상청 자료 가져오기"와 "장소 검색"을 쓰려면, 기상청 자료를 대신 받아 주는 작은 중계 서버가 하나 필요합니다.
한 번만 만들어 두면 됩니다. 비용은 들지 않습니다(무료 요금제: 하루 10만 번).

1. https://dash.cloudflare.com 에 가입하고 로그인합니다.
2. 왼쪽 메뉴 **Compute (Workers)** → **Workers & Pages** → **Create** → **Start with Hello World!** → 이름을 `stackpilot-weather`처럼 정하고 **Deploy**.
3. **Edit code**를 눌러 나오는 코드를 모두 지우고, 이 폴더의 `kma-relay.js` 내용을 그대로 붙여 넣은 뒤 **Deploy**.
4. 화면 위쪽에 나오는 주소(예: `https://stackpilot-weather.내아이디.workers.dev`)를 복사합니다.
5. 모바일 StackPilot → 첫 화면 **⚙ 설정** → **날씨 중계 주소**에 붙여 넣고 **저장**.

이 중계는 기상청 날씨누리의 정해진 네 페이지만 받아 오고, StackPilot 웹사이트(dightkdtjr-byte.github.io)에서 온 요청에만 답합니다. 아무것도 저장하지 않습니다.
