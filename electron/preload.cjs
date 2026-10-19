const { contextBridge, ipcRenderer } = require('electron');

// Minimal bridge: lets the page ask the main process to re-focus the window.
// Needed because, on Windows, inputs stop receiving keyboard input after a
// native alert()/confirm() dialog closes (Electron focus bug).
contextBridge.exposeInMainWorld('stackpilotDesktop', {
  refocus: () => ipcRenderer.send('desktop:refocus'),
});
