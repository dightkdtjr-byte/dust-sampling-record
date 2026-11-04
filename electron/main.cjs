const { app, BrowserWindow, Menu, protocol, net, shell, ipcMain, session } = require('electron');
const path = require('node:path');
const fs = require('node:fs');
const { pathToFileURL } = require('node:url');

const APP_SCHEME = 'app';
const DIST_DIR = path.join(__dirname, '..', 'dist');
const DEV_SERVER_URL = process.env.VITE_DEV_SERVER_URL;

// Serve the built site over a custom "app://" origin instead of file://.
// This keeps fetch() (used for the Excel template) and localStorage working.
protocol.registerSchemesAsPrivileged([
  {
    scheme: APP_SCHEME,
    privileges: { standard: true, secure: true, supportFetchAPI: true, corsEnabled: true },
  },
]);

function registerAppProtocol() {
  protocol.handle(APP_SCHEME, (request) => {
    const url = new URL(request.url);
    let pathname = decodeURIComponent(url.pathname);
    if (!pathname || pathname === '/') pathname = '/index.html';
    const target = path.normalize(path.join(DIST_DIR, pathname));
    if (!target.startsWith(DIST_DIR)) {
      return new Response('Forbidden', { status: 403 });
    }
    if (!fs.existsSync(target)) {
      return new Response('Not Found', { status: 404 });
    }
    return net.fetch(pathToFileURL(target).toString());
  });
}

const WEB_PREFERENCES = {
  preload: path.join(__dirname, 'preload.cjs'),
  contextIsolation: true,
  nodeIntegration: false,
  sandbox: true,
};

const isAppUrl = (url) => (
  url.startsWith(`${APP_SCHEME}://`) || (DEV_SERVER_URL && url.startsWith(DEV_SERVER_URL))
);

function createWindow() {
  const win = new BrowserWindow({
    width: 1280,
    height: 860,
    minWidth: 900,
    minHeight: 600,
    title: 'StackPilot',
    icon: path.join(__dirname, '..', 'build', 'icon.png'),
    autoHideMenuBar: true,
    webPreferences: WEB_PREFERENCES,
  });
  attachWindowBehaviors(win);

  if (DEV_SERVER_URL) {
    win.loadURL(DEV_SERVER_URL);
  } else {
    win.loadURL(`${APP_SCHEME}://bundle/index.html`);
  }
}

// 모든 창(메인 창, 통합 시료채취기록부 별도 창)에 공통으로 적용되는 동작
function attachWindowBehaviors(win) {
  // 앱 내부 주소(window.open)는 새 창으로, 외부 링크는 시스템 브라우저로.
  win.webContents.setWindowOpenHandler(({ url }) => {
    if (isAppUrl(url)) {
      return {
        action: 'allow',
        overrideBrowserWindowOptions: {
          width: 1180,
          height: 920,
          minWidth: 800,
          minHeight: 500,
          title: 'StackPilot - 통합 시료채취기록부',
          icon: path.join(__dirname, '..', 'build', 'icon.png'),
          autoHideMenuBar: true,
          webPreferences: WEB_PREFERENCES,
        },
      };
    }
    if (/^https?:/i.test(url)) shell.openExternal(url);
    return { action: 'deny' };
  });
  win.webContents.on('did-create-window', (childWindow) => {
    attachWindowBehaviors(childWindow);
  });

  // Wrap native dialogs so the window regains keyboard focus after they close.
  win.webContents.on('dom-ready', () => {
    win.webContents.executeJavaScript(`(() => {
      if (window.__stackpilotDialogsWrapped) return;
      window.__stackpilotDialogsWrapped = true;
      const refocus = () => {
        try { window.stackpilotDesktop && window.stackpilotDesktop.refocus(); } catch (e) {}
        try { const el = document.activeElement; if (el && el.blur) { el.blur(); setTimeout(() => el.focus(), 0); } } catch (e) {}
      };
      for (const name of ['alert', 'confirm']) {
        const original = window[name].bind(window);
        window[name] = (...args) => { const result = original(...args); refocus(); return result; };
      }
      const originalClick = HTMLAnchorElement.prototype.click;
      HTMLAnchorElement.prototype.click = function (...args) {
        try { if (this.download && window.stackpilotDesktop) window.stackpilotDesktop.noteDownloadName(this.download); } catch (e) {}
        return originalClick.apply(this, args);
      };
    })();`).catch(() => {});
  });
}

let lastDownloadName = '';
ipcMain.on('desktop:download-name', (event, name) => {
  lastDownloadName = String(name || '').replace(/[\\/:*?"<>|]/g, '_');
});

ipcMain.on('desktop:refocus', (event) => {
  const win = BrowserWindow.fromWebContents(event.sender);
  if (!win || win.isDestroyed()) return;
  win.blur();
  win.focus();
  event.sender.focus();
});

app.whenReady().then(() => {
  Menu.setApplicationMenu(null);
  registerAppProtocol();

  // 엑셀 추출 등 다운로드: 파일 이름을 채운 "다른 이름으로 저장" 창을 띄움
  session.defaultSession.on('will-download', (event, item) => {
    try {
      const reported = item.getFilename();
      const fileName = (!reported || reported === 'download' || !reported.includes('.')) && lastDownloadName ? lastDownloadName : reported;
      lastDownloadName = '';
      item.setSaveDialogOptions({
        title: '파일 저장',
        defaultPath: path.join(app.getPath('downloads'), fileName),
      });
    } catch (error) {
      console.error('download dialog setup failed', error);
    }
  });

  createWindow();

  app.on('activate', () => {
    if (BrowserWindow.getAllWindows().length === 0) createWindow();
  });
});

app.on('window-all-closed', () => {
  if (process.platform !== 'darwin') app.quit();
});
