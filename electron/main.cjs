const { app, BrowserWindow, Menu, protocol, net, shell, ipcMain } = require('electron');
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

function createWindow() {
  const win = new BrowserWindow({
    width: 1280,
    height: 860,
    minWidth: 900,
    minHeight: 600,
    title: 'StackPilot',
    icon: path.join(__dirname, '..', 'build', 'icon.png'),
    autoHideMenuBar: true,
    webPreferences: {
      preload: path.join(__dirname, 'preload.cjs'),
      contextIsolation: true,
      nodeIntegration: false,
      sandbox: true,
    },
  });

  // Open external links (e.g. GitHub token pages) in the system browser.
  win.webContents.setWindowOpenHandler(({ url }) => {
    if (/^https?:/i.test(url)) shell.openExternal(url);
    return { action: 'deny' };
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
    })();`).catch(() => {});
  });

  if (DEV_SERVER_URL) {
    win.loadURL(DEV_SERVER_URL);
  } else {
    win.loadURL(`${APP_SCHEME}://bundle/index.html`);
  }
}

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
  createWindow();

  app.on('activate', () => {
    if (BrowserWindow.getAllWindows().length === 0) createWindow();
  });
});

app.on('window-all-closed', () => {
  if (process.platform !== 'darwin') app.quit();
});
