const { contextBridge, ipcRenderer } = require('electron');

// Minimal bridge: lets the page ask the main process to re-focus the window.
// Needed because, on Windows, inputs stop receiving keyboard input after a
// native alert()/confirm() dialog closes (Electron focus bug).
contextBridge.exposeInMainWorld('stackpilotDesktop', {
  refocus: () => ipcRenderer.send('desktop:refocus'),
  // 화면의 창 닫기 버튼도 X 버튼과 같은 닫기 경로(저장 확인 포함)를 쓰게 한다
  closeWindow: () => ipcRenderer.send('desktop:close-window'),
  // blob 다운로드는 Electron이 파일 이름을 못 받는 경우가 있어 화면에서 지정한 이름을 전달
  noteDownloadName: (name) => ipcRenderer.send('desktop:download-name', String(name || '')),
  // 엑셀 저장 시 파일을 넣을 묶음 폴더 이름(날짜_사업장명_#배출구)
  noteDownloadFolder: (name) => ipcRenderer.send('desktop:download-folder', String(name || '')),
  // 기상청 도시별관측 자료 (tm: 'YYYY.MM.DD.HH:00', 한국 시각)
  fetchWeather: (tm) => ipcRenderer.invoke('desktop:fetch-weather', String(tm || '')),
  // 엑셀 저장 위치: 지금 위치 보기, 폴더 고르기, 다운로드 폴더로 되돌리기
  getExportBase: () => ipcRenderer.invoke('desktop:get-export-base'),
  chooseExportBase: () => ipcRenderer.invoke('desktop:choose-export-base'),
  resetExportBase: () => ipcRenderer.invoke('desktop:reset-export-base'),
  // 기상청 장소 검색과 동네 현재 날씨
  searchPlaces: (query) => ipcRenderer.invoke('desktop:search-places', String(query || '')),
  fetchLocalWeather: (dongCode) => ipcRenderer.invoke('desktop:fetch-local-weather', String(dongCode || '')),
});
