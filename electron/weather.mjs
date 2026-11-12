// 설치형(electron 메인 프로세스)과 웹(브라우저, 중계 서버 경유)이 함께 쓰는 기상청 자료 읽기
// 기상청 날씨누리 "도시별관측" 표를 읽어 지점별 관측값으로 바꾼다.
// 인증키가 필요 없는 공식 웹 페이지이며, 전국 종관(ASOS) 지점의 매 정시 관측값을 준다.
// 주의: 기압은 해면기압만 나온다(현지기압 아님).
const PAGE_URL = 'https://www.weather.go.kr/w/weather/land/city-obs.do';

const stripTags = (html) => String(html)
  .replace(/<script[^>]*>[\s\S]*?<\/script>/g, '')
  .replace(/<[^>]+>/g, '')
  .replace(/&nbsp;/g, ' ')
  .replace(/&amp;/g, '&')
  .replace(/\s+/g, ' ')
  .trim();
const toNumber = (text) => {
  const m = /-?\d+(?:\.\d+)?/.exec(String(text || ''));
  return m ? Number(m[0]) : null;
};

// 표 한 줄 → 관측값. 열 순서: 지점, 현재일기, 시정, 운량, 중하운량, 기온, 이슬점, 체감, 일강수, 적설, 습도, 풍향, 풍속, 해면기압
const parseCityObservations = (html) => {
  const rows = [];
  const rowPattern = /<tr[^>]*>([\s\S]*?)<\/tr>/g;
  let match;
  while ((match = rowPattern.exec(html)) !== null) {
    const body = match[1];
    const stn = /[?&;]stn=(\d+)/.exec(body);
    const cells = [...body.matchAll(/<td[^>]*>([\s\S]*?)<\/td>/g)].map((m) => m[1]);
    if (!stn || cells.length < 14) continue;
    const wind = /writeWindSpeed\('([^']*)'/.exec(cells[12]);
    rows.push({
      stn: stn[1],
      name: stripTags(cells[0]),
      weather: stripTags(cells[1]),
      cloud: toNumber(stripTags(cells[3])),
      temp: toNumber(stripTags(cells[5])),
      humidity: toNumber(stripTags(cells[10])),
      windDir: stripTags(cells[11]),
      windSpeed: wind ? toNumber(wind[1]) : toNumber(stripTags(cells[12])),
      seaPressure: toNumber(stripTags(cells[13])),
    });
  }
  return rows;
};

// tm: 'YYYY.MM.DD.HH:00' (한국 시각)
const fetchCityObservations = async (tm, fetchImpl = fetch) => {
  const url = `${PAGE_URL}?auto_man=m&stn=0&type=t99&reg=100&tm=${encodeURIComponent(tm)}`;
  const response = await fetchImpl(url, { headers: { 'User-Agent': 'Mozilla/5.0' } });
  if (!response.ok) throw new Error(`기상청 응답 오류 (${response.status})`);
  const html = await response.text();
  const shown = /id="select-tm"[^>]*value="([^"]*)"/.exec(html);
  return { tm: shown ? shown[1] : tm, stations: parseCityObservations(html) };
};

// ── 동네(읍·면·동) 단위: 기상청 날씨누리의 장소 검색과 "현재 날씨" ──
const SITE = 'https://www.weather.go.kr';
const HEADERS = { 'User-Agent': 'Mozilla/5.0', Referer: `${SITE}/w/index.do` };

// 장소 검색: 시·군·구와 읍·면·동을 먼저, 모자라면 시설·지명까지. 결과: [{ title, address, dongCode, lat, lon }]
const searchPlaces = async (query, fetchImpl = fetch) => {
  const text = String(query || '').trim();
  if (!text) return [];
  const results = [];
  // A3: 시·군·구, A1: 읍·면·동, A2: 시설·지명 순으로 모은다
  for (const src of ['A3', 'A1', 'A2']) {
    const url = `${SITE}/w/renew2021/rest/main/place-search.do?query=${encodeURIComponent(text)}&start=1&src=${src}`;
    try {
      const response = await fetchImpl(url, { headers: HEADERS });
      if (!response.ok) continue;
      const list = JSON.parse(await response.text());
      (Array.isArray(list) ? list : []).forEach((item) => {
        const lat = Number(item.latitude); const lon = Number(item.longitude);
        if (!item.dongCode || !Number.isFinite(lat) || !Number.isFinite(lon)) return;
        results.push({ title: String(item.title || '').trim(), address: String(item.address || '').trim(), dongCode: String(item.dongCode), lat, lon });
      });
    } catch (error) { /* 다음 종류로 */ }
    if (src !== 'A3' && results.length >= 6) break;
  }
  const seen = new Set();
  return results.filter((item) => { const key = `${item.title}|${item.address}`; if (seen.has(key)) return false; seen.add(key); return true; }).slice(0, 12);
};

const parseCurrentWeather = (html) => {
  const pick = (pattern) => { const m = pattern.exec(html); return m ? stripTags(m[1]) : ''; };
  const wind = pick(/ic-wind[\s\S]*?<span class="val">([\s\S]*?)<\/span>\s*<\/li>/);
  const windMatch = /^(\S+)\s+(-?\d+(?:\.\d+)?)/.exec(wind);
  return {
    updated: pick(/class="updated-at"[^>]*>\s*<span>([^<]*)<\/span>/),
    temp: toNumber(pick(/class="tmp">\s*(-?\d+(?:\.\d+)?)/)),
    humidity: toNumber(pick(/ic-hm[\s\S]*?<span class="val">([\s\S]*?)<\/span>\s*<\/li>/)),
    windDir: windMatch ? windMatch[1] : '',
    windSpeed: windMatch ? Number(windMatch[2]) : null,
  };
};
// 동네예보 시간별 표의 첫 칸(지금 시각대) 하늘 상태
const parseForecastSky = (html) => {
  const m = /<ul class="item[^"]*"[^>]*data-time="[^"]*"[\s\S]*?class="wic[^"]*"\s+title="([^"]*)"/.exec(html);
  return m ? m[1].replace(/\s+/g, '') : '';
};
// 동(행정동 코드)의 현재 날씨: 가까운 관측값(습도·바람)과 동네예보의 지금 하늘 상태
const fetchLocalWeather = async (dongCode, fetchImpl = fetch) => {
  const code = String(dongCode || '');
  if (!/^\d{10}$/.test(code)) throw new Error('위치 코드가 올바르지 않습니다.');
  const current = await fetchImpl(`${SITE}/w/wnuri-fct2021/main/current-weather.do?code=${code}&unit=m%2Fs&aws=N`, { headers: HEADERS });
  if (!current.ok) throw new Error(`기상청 응답 오류 (${current.status})`);
  const result = parseCurrentWeather(await current.text());
  let sky = '';
  try {
    const forecast = await fetchImpl(`${SITE}/w/wnuri-fct2021/main/digital-forecast.do?code=${code}&unit=m%2Fs&hr1=Y`, { headers: HEADERS });
    if (forecast.ok) sky = parseForecastSky(await forecast.text());
  } catch (error) { /* 하늘 상태 없이 진행 */ }
  return { ...result, sky };
};

export { parseCityObservations, fetchCityObservations, searchPlaces, fetchLocalWeather, parseCurrentWeather, parseForecastSky };
