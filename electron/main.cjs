const { app, BrowserWindow, Menu, protocol, net, shell } = require('electron');
const path = require('node:path');
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

  if (DEV_SERVER_URL) {
    win.loadURL(DEV_SERVER_URL);
  } else {
    win.loadURL(`${APP_SCHEME}://bundle/index.html`);
  }
}

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
