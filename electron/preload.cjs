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
});
