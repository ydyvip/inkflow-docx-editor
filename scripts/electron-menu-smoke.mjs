/**
 * 临时诊断脚本 —— 验证「文件 → 打开 DOCX…」菜单项 / Ctrl+O
 * 是否真的触发原生打开对话框。
 * 做法：把主进程 dialog.showOpenDialog 包一层（记录是否被调用并以"取消"返回），
 *       然后通过 app 菜单点击「打开 DOCX…」，看链条是否贯通。
 * 用法：node scripts/electron-menu-smoke.mjs
 */
import { _electron as electron } from 'playwright';

const app = await electron.launch({ args: ['.'] });
const page = await app.firstWindow();

const rendererLogs = [];
page.on('console', (m) => rendererLogs.push(`[console:${m.type()}] ${m.text()}`));
page.on('pageerror', (e) => rendererLogs.push(`[pageerror] ${e.message}`));

// 1) 在渲染进程里监听菜单事件是否真的到达
await page.evaluate(() => {
  window.__menuEvents = 0;
  window.inkflow.onMenu((action) => {
    window.__menuEvents++;
    window.__lastMenuAction = action;
  });
});

// 2) 包一层 dialog.showOpenDialog，记录是否被调用并自动取消
await app.evaluate(({ dialog }) => {
  global.__dialogCalled = 0;
  global.__dialogParent = '(none)';
  const orig = dialog.showOpenDialog.bind(dialog);
  dialog.showOpenDialog = async (win, opts) => {
    global.__dialogCalled++;
    global.__dialogParent =
      win && win.constructor && win.constructor.name
        ? win.constructor.name
        : '(null)';
    return { canceled: true, filePaths: [] };
  };
});

// 3) 等页面就绪，通过真实应用菜单触发「打开 DOCX…」
await page.waitForTimeout(1200);
await app.evaluate(({ Menu }) => {
  const menu = Menu.getApplicationMenu();
  if (!menu) {
    global.__menuFound = false;
    return;
  }
  global.__menuFound = true;
  const item = menu.getMenuItemById('open');
  global.__itemFound = !!item;
  item?.click(); // 触发与用户点击完全相同的 handler
});
await page.waitForTimeout(1500);

// 4) 汇总
const result = await app.evaluate(() => ({
  menuFound: global.__menuFound,
  itemFound: global.__itemFound,
  dialogCalled: global.__dialogCalled,
  dialogParent: global.__dialogParent,
}));
const renderer = await page.evaluate(() => ({
  menuEvents: window.__menuEvents,
  lastAction: window.__lastMenuAction,
}));

console.log('app.menuFound        =', result.menuFound);
console.log('app.itemFound        =', result.itemFound);
console.log('app.dialogCalled     =', result.dialogCalled, '(0=菜单没触达渲染进程或没调用对话框)');
console.log('app.dialogParent     =', result.dialogParent);
console.log('renderer.menuEvents  =', renderer.menuEvents, '(渲染进程收到的 inkflow:menu 次数)');
console.log('renderer.lastAction  =', renderer.lastAction);
console.log('--- renderer console ---');
for (const l of rendererLogs) console.log(l);

await app.close();
