// StackPilot 날씨 중계 (Cloudflare Workers용)
// 모바일(사파리)에서는 브라우저가 기상청 사이트에 직접 접속하지 못하게 막는다(CORS).
// 이 작은 서버가 기상청 날씨누리의 정해진 페이지만 대신 받아서 그대로 돌려준다.
// - 기상청 www.weather.go.kr 의 아래 세 경로만 허용한다 (다른 주소는 거절 → 아무 사이트나 중계하는 데 쓰일 수 없음)
// - 요청한 사이트가 아래 ALLOWED_ORIGINS 에 있을 때만 응답한다
// - 아무것도 저장하지 않는다

const ALLOWED_ORIGINS = [
  'https://dightkdtjr-byte.github.io',
];
const ALLOWED_LOCAL = /^http:\/\/(localhost|127\.0\.0\.1)(:\d+)?$/; // 개발·시험용
const ALLOWED_PATHS = [
  '/w/weather/land/city-obs.do',                    // 도시별관측 (지난 날짜 포함 정시 관측값)
  '/w/renew2021/rest/main/place-search.do',         // 장소 검색
  '/w/wnuri-fct2021/main/current-weather.do',       // 동네 현재 날씨
  '/w/wnuri-fct2021/main/digital-forecast.do',      // 동네예보 (지금 하늘 상태)
];

const corsHeaders = (origin) => ({
  'Access-Control-Allow-Origin': origin,
  'Access-Control-Allow-Methods': 'GET, OPTIONS',
  'Access-Control-Max-Age': '600',
  Vary: 'Origin',
});

export default {
  async fetch(request) {
    const origin = request.headers.get('Origin') || '';
    const originOk = ALLOWED_ORIGINS.includes(origin) || ALLOWED_LOCAL.test(origin);
    if (!originOk) return new Response('forbidden origin', { status: 403 });
    if (request.method === 'OPTIONS') return new Response(null, { status: 204, headers: corsHeaders(origin) });
    if (request.method !== 'GET') return new Response('method not allowed', { status: 405, headers: corsHeaders(origin) });

    let target;
    try {
      target = new URL(new URL(request.url).searchParams.get('u') || '');
    } catch {
      return new Response('bad url', { status: 400, headers: corsHeaders(origin) });
    }
    if (target.protocol !== 'https:' || target.hostname !== 'www.weather.go.kr' || !ALLOWED_PATHS.includes(target.pathname)) {
      return new Response('not allowed', { status: 403, headers: corsHeaders(origin) });
    }

    const upstream = await fetch(target.toString(), {
      headers: { 'User-Agent': 'Mozilla/5.0', Referer: 'https://www.weather.go.kr/w/index.do' },
    });
    return new Response(await upstream.text(), {
      status: upstream.status,
      headers: {
        ...corsHeaders(origin),
        'Content-Type': upstream.headers.get('Content-Type') || 'text/plain; charset=utf-8',
        'Cache-Control': 'no-store',
      },
    });
  },
};
