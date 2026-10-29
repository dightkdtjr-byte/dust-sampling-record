const { contextBridge, ipcRenderer } = require('electron');

// Minimal bridge: lets the page ask the main process to re-focus the window.
// Needed because, on Windows, inputs stop receiving keyboard input after a
// native alert()/confirm() dialog closes (Electron focus bug).
contextBridge.exposeInMainWorld('stackpilotDesktop', {
  refocus: () => ipcRenderer.send('desktop:refocus'),
  // blob 다운로드는 Electron이 파일 이름을 못 받는 경우가 있어 화면에서 지정한 이름을 전달
  noteDownloadName: (name) => ipcRenderer.send('desktop:download-name', String(name || '')),
});
