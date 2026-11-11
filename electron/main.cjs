const { app, BrowserWindow, Menu, protocol, net, shell, ipcMain, session, dialog } = require('electron');
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
  // 닫기 전에 저장하지 않은 입력이 있으면 저장 여부를 묻는다 (X 버튼, 창 닫기 버튼, 앱 종료 모두)
  let closeConfirmed = false;
  let closePrompting = false;
  win.on('close', (event) => {
    if (closeConfirmed || win.webContents.isDestroyed()) return;
    event.preventDefault();
    if (closePrompting) return;
    closePrompting = true;
    const finishClose = () => { closeConfirmed = true; closePrompting = false; if (!win.isDestroyed()) win.close(); };
    win.webContents
      .executeJavaScript('Boolean(window.__stackpilotHasUnsaved && window.__stackpilotHasUnsaved())', true)
      .catch(() => false)
      .then(async (hasUnsaved) => {
        if (!hasUnsaved) { finishClose(); return; }
        const { response } = await dialog.showMessageBox(win, {
          type: 'question',
          title: 'StackPilot',
          message: '저장하지 않은 내용이 있습니다.',
          detail: '닫기 전에 저장할까요?',
          buttons: ['저장', '저장 안 함', '취소'],
          defaultId: 0,
          cancelId: 2,
          noLink: true,
        });
        if (response === 2) { closePrompting = false; return; }
        if (response === 0) {
          const saved = await win.webContents
            .executeJavaScript('window.__stackpilotSaveBeforeClose ? window.__stackpilotSaveBeforeClose() : false', true)
            .catch(() => false);
          if (!saved) { closePrompting = false; return; }
        }
        finishClose();
      });
  });

  // 앱 내부 주소(window.open)는 새 창으로, 외부 링크는 시스템 브라우저로.
  win.webContents.setWindowOpenHandler(({ url }) => {
    if (isAppUrl(url)) {
      return {
        action: 'allow',
        overrideBrowserWindowOptions: {
          width: 1280,
          height: 900,
          minWidth: 800,
          minHeight: 500,
          title: 'StackPilot',
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
// 엑셀 저장 시 파일을 넣을 묶음 폴더 이름(날짜_사업장명_#배출구). 화면에서 내려받기 직전에 알려 준다
let lastDownloadFolder = '';
ipcMain.on('desktop:download-folder', (event, name) => {
  lastDownloadFolder = String(name || '').replace(/[\\/:*?"<>|]/g, '_').replace(/[. ]+$/, '').trim();
});
// 묶음 폴더를 만들 상위 위치: 처음엔 다운로드 폴더, 이후엔 마지막으로 저장한 위치를 기억
const exportBaseFile = () => path.join(app.getPath('userData'), 'export-base.json');
const readExportBase = () => {
  try {
    const dir = JSON.parse(fs.readFileSync(exportBaseFile(), 'utf8')).dir;
    if (dir && fs.existsSync(dir) && fs.statSync(dir).isDirectory()) return dir;
  } catch (error) { /* 처음 실행이거나 폴더가 없어졌으면 기본값 */ }
  return app.getPath('downloads');
};
const writeExportBase = (dir) => {
  try { fs.writeFileSync(exportBaseFile(), JSON.stringify({ dir })); } catch (error) { console.error('export base save failed', error); }
};

// 기상청 날씨누리 도시별관측 자료 가져오기 (인증키 없이 볼 수 있는 공식 페이지). 화면에서는 CORS로 막히므로 여기서 받는다
ipcMain.handle('desktop:fetch-weather', async (event, tm) => {
  const text = String(tm || '');
  if (!/^\d{4}\.\d{2}\.\d{2}\.\d{2}:00$/.test(text)) throw new Error('시각 형식이 올바르지 않습니다.');
  const { fetchCityObservations } = require('./weather.cjs');
  return fetchCityObservations(text, (url, options) => net.fetch(url, options));
});

// 기상청 장소 검색(동·주소·지명)과 그 동네의 현재 날씨
ipcMain.handle('desktop:search-places', async (event, query) => {
  const { searchPlaces } = require('./weather.cjs');
  return searchPlaces(String(query || '').slice(0, 60), (url, options) => net.fetch(url, options));
});
ipcMain.handle('desktop:fetch-local-weather', async (event, dongCode) => {
  const { fetchLocalWeather } = require('./weather.cjs');
  return fetchLocalWeather(String(dongCode || ''), (url, options) => net.fetch(url, options));
});

// 화면 안의 "창 닫기" 버튼: 창의 close 이벤트를 거치게 해서 저장 확인이 뜨도록 한다
ipcMain.on('desktop:close-window', (event) => {
  const win = BrowserWindow.fromWebContents(event.sender);
  if (win && !win.isDestroyed()) win.close();
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
      const folderName = lastDownloadFolder;
      lastDownloadFolder = '';
      // 묶음 폴더(날짜_사업장명_#배출구)를 미리 만들어 저장 창이 그 안에서 열리게 한다
      let targetDir = readExportBase();
      let createdDir = '';
      if (folderName) {
        const dir = path.join(targetDir, folderName);
        try {
          if (!fs.existsSync(dir)) { fs.mkdirSync(dir, { recursive: true }); createdDir = dir; }
          targetDir = dir;
        } catch (error) {
          console.error('export folder create failed', error);
        }
      }
      item.setSaveDialogOptions({
        title: '파일 저장',
        defaultPath: path.join(targetDir, fileName),
      });
      // 저장이 끝나면 저장된 파일이 있는 폴더를 열어 파일을 선택해 보여줌
      item.once('done', (_event, state) => {
        try {
          const savedPath = state === 'completed' ? item.getSavePath() : '';
          // 저장을 취소했거나 다른 곳에 저장했으면, 방금 만든 빈 폴더는 지운다
          if (createdDir && (!savedPath || path.dirname(savedPath) !== createdDir)) {
            try { if (fs.readdirSync(createdDir).length === 0) fs.rmdirSync(createdDir); } catch (error) { /* 비어 있지 않으면 둔다 */ }
          }
          if (!savedPath) return;
          if (folderName) {
            const savedDir = path.dirname(savedPath);
            writeExportBase(path.basename(savedDir) === folderName ? path.dirname(savedDir) : savedDir);
          }
          shell.showItemInFolder(savedPath);
        } catch (error) {
          console.error('open folder failed', error);
        }
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
