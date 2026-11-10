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

module.exports = { parseCityObservations, fetchCityObservations };
