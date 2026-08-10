/**
 * InkFlow — Electron 主进程
 * ------------------------------------------------------------
 * 职责：
 *   1. 创建 BrowserWindow 并加载渲染进程（开发模式走 Vite dev server，
 *      生产模式加载 dist/index.html）。
 *   2. 通过 IPC 暴露「打开 / 保存 DOCX」的原生文件对话框。
 *   3. 注册原生应用菜单（文件菜单：新建 / 打开 / 保存 / 另存为），
 *      并把对应的菜单动作以事件推送给渲染进程。
 *   4. 支持通过命令行参数直接打开一个 .docx（双击文件关联）。
 *
 * 安全：使用 contextIsolation + preload 的 contextBridge，渲染进程不
 * 直接接触 Node 能力，只通过 window.inkflow 调用白名单里的 IPC。
 */
const { app, BrowserWindow, ipcMain, dialog, Menu, shell } = require('electron');
const path = require('node:path');
const fs = require('node:fs');

const DOCX_FILTERS = [
  { name: 'DOCX 文档', extensions: ['docx'] },
  { name: '所有文件', extensions: ['*'] },
];

let mainWindow = null;

function createWindow() {
  mainWindow = new BrowserWindow({
    width: 1440,
    height: 920,
    minWidth: 960,
    minHeight: 620,
    show: false,
    backgroundColor: '#f7f8fa',
    webPreferences: {
      preload: path.join(__dirname, 'preload.cjs'),
      contextIsolation: true,
      nodeIntegration: false,
      sandbox: false, // 允许 preload 使用 require('electron')
    },
  });

  // 打开外部链接（例如编辑器里的超链接）交给系统浏览器，不在应用内新开窗口
  mainWindow.webContents.setWindowOpenHandler(({ url }) => {
    if (/^https?:/i.test(url)) shell.openExternal(url);
    return { action: 'deny' };
  });

  mainWindow.once('ready-to-show', () => mainWindow?.show());

  const devServerUrl = process.env.VITE_DEV_SERVER_URL;
  if (devServerUrl) {
    mainWindow.loadURL(devServerUrl);
    mainWindow.webContents.openDevTools({ mode: 'detach' });
  } else {
    mainWindow.loadFile(path.join(__dirname, '..', 'dist', 'index.html'));
  }

  mainWindow.on('closed', () => {
    mainWindow = null;
  });
}

/** 构造 window.inkflow 期望的「打开结果」：文件内容以可克隆的字节数组返回 */
function readDocxBuffer(filePath) {
  const buf = fs.readFileSync(filePath);
  // 跨 contextBridge 传递：转成普通的 Uint8Array 视图即可被 structured clone
  const data = new Uint8Array(buf.buffer, buf.byteOffset, buf.byteLength);
  return {
    name: path.basename(filePath),
    filePath,
    data,
    size: buf.byteLength,
  };
}

/* ------------------------------------------------------------------ *
 * IPC —— 渲染进程（window.inkflow）可调用的原生对话框
 * ------------------------------------------------------------------ */

/** 选取一个有效的、未被销毁的窗口作为对话框父窗口。
 *  点击原生菜单时主窗口可能短暂失焦，getFocusedWindow() 可能返回 null；
 *  这里兜底到 mainWindow；两者都没有就返回 undefined（走无父窗口的对话框），
 *  避免把 null 传进 dialog API 导致异常/不弹出。 */
function getDialogParent() {
  const focused = BrowserWindow.getFocusedWindow();
  if (focused && !focused.isDestroyed()) return focused;
  if (mainWindow && !mainWindow.isDestroyed()) return mainWindow;
  return undefined;
}

// 打开 DOCX：弹出系统文件选择框，读取后返回 { name, filePath, data, size }
ipcMain.handle('inkflow:open', async () => {
  const parent = getDialogParent();
  const options = {
    title: '打开 DOCX 文档',
    properties: ['openFile'],
    filters: DOCX_FILTERS,
  };
  const result = parent
    ? await dialog.showOpenDialog(parent, options)
    : await dialog.showOpenDialog(options);
  if (result.canceled || result.filePaths.length === 0) return null;
  try {
    return readDocxBuffer(result.filePaths[0]);
  } catch (err) {
    return { error: err instanceof Error ? err.message : String(err) };
  }
});

// 保存当前打开的文档的原始文件路径（供「保存」覆盖用）
let currentFilePath = null;
ipcMain.handle('inkflow:set-current-file', (_e, filePath) => {
  currentFilePath = filePath ? String(filePath) : null;
  return true;
});

// 保存 DOCX：
//   saveAs=true  → 总是弹出「另存为」对话框
//   saveAs=false & 已打开过文件（currentFilePath 存在）→ 直接覆盖保存
//   否则 → 弹出「另存为」对话框
// 返回 { filePath, name }，取消则返回 null
ipcMain.handle(
  'inkflow:save',
  async (_event, { defaultName, data, saveAs = false } = {}) => {
    const baseName =
      String(defaultName || 'document.docx').replace(/\.docx$/i, '') + '.docx';

    let filePath = !saveAs ? currentFilePath : null;
    if (!filePath) {
      const parent = getDialogParent();
      const options = {
        title: '导出 DOCX',
        defaultPath: baseName,
        filters: DOCX_FILTERS,
      };
      const result = parent
        ? await dialog.showSaveDialog(parent, options)
        : await dialog.showSaveDialog(options);
      if (result.canceled || !result.filePath) return null;
      filePath = result.filePath;
    }
    // data 是从渲染进程传回的 Uint8Array
    fs.writeFileSync(filePath, Buffer.from(data));
    currentFilePath = filePath;
    return { filePath, name: path.basename(filePath) };
  }
);

/* ------------------------------------------------------------------ *
 * 原生应用菜单 + 快捷键 → 推送给渲染进程
 * ------------------------------------------------------------------ */
function sendToRenderer(channel, payload) {
  mainWindow?.webContents.send(channel, payload);
}

function buildMenu() {
  const isMac = process.platform === 'darwin';
  const template = [
    ...(isMac
      ? [{ role: 'appMenu' }]
      : []),
    {
      label: '文件',
      submenu: [
        { id: 'new', label: '新建空白文档', accelerator: 'CmdOrCtrl+N', click: () => sendToRenderer('inkflow:menu', 'new') },
        { type: 'separator' },
        { id: 'open', label: '打开 DOCX…', accelerator: 'CmdOrCtrl+O', click: () => sendToRenderer('inkflow:menu', 'open') },
        { type: 'separator' },
        { id: 'save', label: '保存', accelerator: 'CmdOrCtrl+S', click: () => sendToRenderer('inkflow:menu', 'save') },
        { id: 'saveAs', label: '另存为…', accelerator: 'CmdOrCtrl+Shift+S', click: () => sendToRenderer('inkflow:menu', 'saveAs') },
        { type: 'separator' },
        isMac ? { role: 'close' } : { role: 'quit', label: '退出' },
      ],
    },
    { role: 'editMenu', label: '编辑' },
    { role: 'viewMenu', label: '视图' },
    { role: 'windowMenu', label: '窗口' },
  ];
  Menu.setApplicationMenu(Menu.buildFromTemplate(template));
}

/* ------------------------------------------------------------------ *
 * 生命周期
 * ------------------------------------------------------------------ */
app.whenReady().then(() => {
  buildMenu();
  createWindow();

  // 从命令行参数 / 文件关联打开的 .docx：等窗口就绪后推给渲染进程
  const fileArg = process.argv
    .slice(app.isPackaged ? 1 : 2)
    .find((a) => !a.startsWith('-') && a.toLowerCase().endsWith('.docx'));
  if (fileArg) {
    mainWindow?.webContents.once('did-finish-load', () => {
      try {
        sendToRenderer('inkflow:open-result', readDocxBuffer(fileArg));
      } catch {
        /* 文件读取失败则忽略，静默回落到上传页 */
      }
    });
  }

  app.on('activate', () => {
    if (BrowserWindow.getAllWindows().length === 0) createWindow();
  });
});

app.on('window-all-closed', () => {
  if (process.platform !== 'darwin') app.quit();
});
