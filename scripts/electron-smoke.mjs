/**
 * 临时诊断脚本 —— 用 Playwright 启动真实 Electron 应用，点击「📂 打开 DOCX…」
 * 按钮，捕获渲染进程 console / pageerror，判断点击后到底发生了什么。
 * 用法：node scripts/electron-smoke.mjs
 */
import { _electron as electron } from 'playwright';

console.log('launching electron...');
const app = await electron.launch({ args: ['.'] });

const logs = [];
app.process().stderr.on('data', (d) => console.log('[main-stderr]', String(d).trim()));

const page = await app.firstWindow();
page.on('console', (msg) => logs.push(`[console:${msg.type()}] ${msg.text()}`));
page.on('pageerror', (err) => logs.push(`[pageerror] ${err.message}`));

// 1) 桥是否存在
const bridgeType = await page.evaluate(() =>
  window.inkflow ? Object.keys(window.inkflow) : 'NO_BRIDGE'
);
console.log('window.inkflow =', JSON.stringify(bridgeType), '| isElectron =', await page.evaluate(() => !!window.inkflow));

// 2) 上传页里的「打开 DOCX」按钮是否存在
const btns = await page.locator('button', { hasText: '打开 DOCX' }).count();
console.log('“打开 DOCX” 按钮数量 =', btns);

if (btns > 0) {
  console.log('点击按钮...');
  await page.locator('button', { hasText: '打开 DOCX' }).first().click();
  // 按钮点击后立刻抓状态文本，看是否进入「正在解析」或被 catch 掉
  await page.waitForTimeout(1500);
  const body = await page.locator('body').innerText();
  const statusHint = body.includes('正在解析') ? 'PARSING(挂起=已弹原生对话框)' : body.includes('解析失败') ? 'ERROR' : 'IDLE(无反应)';
  console.log('点击后状态 =', statusHint);
}

console.log('--- 渲染进程日志 ---');
for (const l of logs) console.log(l);

await app.close();
