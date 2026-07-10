import { chromium } from 'playwright';
import { readFileSync } from 'fs';

const BASE = process.env.BASE_URL || 'http://localhost:4300';
const pageErrors = [];

const browser = await chromium.launch({
  executablePath: '/opt/pw-browsers/chromium-1194/chrome-linux/chrome',
});
const page = await browser.newPage();
page.on('pageerror', (err) => pageErrors.push(String(err)));

const log = (msg) => console.log('• ' + msg);

async function moveCursorToDocEnd() {
  await page.locator('.ProseMirror > *').last().click();
  await page.waitForTimeout(150);
  await page.keyboard.press('End');
  await page.waitForTimeout(50);
}
async function newLineAtEnd() {
  await moveCursorToDocEnd();
  await page.keyboard.press('Enter');
  await page.waitForTimeout(50);
}
async function selectLastNChars(n) {
  await page.keyboard.press('Home');
  await page.keyboard.down('Shift');
  for (let i = 0; i < n; i++) await page.keyboard.press('ArrowRight');
  await page.keyboard.up('Shift');
}

try {
  await page.goto(BASE, { waitUntil: 'networkidle' });
  await page.locator('.dropzone, .ghost-btn').last().click();
  await page.waitForSelector('.editor-shell', { timeout: 10000 });
  await page.locator('.ProseMirror').click();
  log('started from a blank document');

  // ---- 1. 格式清除 ----
  await page.keyboard.type('清除格式测试文字');
  await selectLastNChars(8);
  await page
    .locator('.toolbar-format .toolbar-select')
    .first()
    .selectOption('Arial');
  await page.waitForTimeout(80);
  await page.getByTitle('加粗').click();
  await page.waitForTimeout(80);
  const beforeClear = await page
    .locator('.ProseMirror .docx-style-run')
    .first()
    .getAttribute('style')
    .catch(() => null);
  const boldBefore = await page.locator('.ProseMirror strong').count();
  if (!beforeClear || boldBefore < 1)
    throw new Error('清除格式前置状态不对：未能先施加样式');

  await selectLastNChars(8);
  await page.getByTitle('清除字符与段落格式（不影响批注/链接）').click();
  await page.waitForTimeout(150);
  const boldAfter = await page.locator('.ProseMirror strong').count();
  const styleRunAfter = await page
    .locator('.ProseMirror .docx-style-run')
    .count();
  if (boldAfter > 0 || styleRunAfter > 0)
    throw new Error('清除格式未生效，仍残留 strong/docx-style-run');
  log('清除格式：加粗 + 字体样式被正确移除');

  // 段落级属性也应被清除格式重置
  await newLineAtEnd();
  await page.keyboard.type('对齐缩进清除测试');
  await page.getByTitle('居中').click();
  await page.waitForTimeout(80);
  await page.getByTitle('增加缩进').click();
  await page.waitForTimeout(80);
  const paraBefore = await page
    .locator('.ProseMirror > p')
    .last()
    .getAttribute('style');
  if (!paraBefore || !paraBefore.includes('center'))
    throw new Error('对齐未生效，无法验证清除格式对段落属性的重置');
  await page.getByTitle('清除字符与段落格式（不影响批注/链接）').click();
  await page.waitForTimeout(150);
  const paraAfter = await page
    .locator('.ProseMirror > p')
    .last()
    .getAttribute('style');
  if (paraAfter && paraAfter.includes('center'))
    throw new Error('清除格式未重置段落对齐属性: ' + paraAfter);
  log('清除格式：段落对齐/缩进属性也被正确重置');

  // ---- 2. 超链接：新增 / 编辑 / 移除 ----
  await newLineAtEnd();
  await page.keyboard.type('这是一个链接文字');
  await selectLastNChars(8);
  page.once('dialog', (d) => d.accept('https://example.com'));
  await page.getByTitle('插入/编辑/移除链接').click();
  await page.waitForTimeout(150);
  const hrefAfterInsert = await page
    .locator('.ProseMirror a')
    .first()
    .getAttribute('href');
  if (hrefAfterInsert !== 'https://example.com')
    throw new Error('插入链接失败: ' + hrefAfterInsert);
  log('超链接：插入成功 href=' + hrefAfterInsert);

  await selectLastNChars(8);
  page.once('dialog', (d) => {
    if (!d.defaultValue().includes('example.com'))
      throw new Error('编辑链接时未回填已有地址: ' + d.defaultValue());
    d.accept('https://changed.example.com');
  });
  await page.getByTitle('插入/编辑/移除链接').click();
  await page.waitForTimeout(150);
  const hrefAfterEdit = await page
    .locator('.ProseMirror a')
    .first()
    .getAttribute('href');
  if (hrefAfterEdit !== 'https://changed.example.com')
    throw new Error('编辑链接失败: ' + hrefAfterEdit);
  log('超链接：编辑回填并更新成功 href=' + hrefAfterEdit);

  await selectLastNChars(8);
  page.once('dialog', (d) => d.accept(''));
  await page.getByTitle('插入/编辑/移除链接').click();
  await page.waitForTimeout(150);
  const linkCountAfterRemove = await page.locator('.ProseMirror a').count();
  if (linkCountAfterRemove > 0) throw new Error('清空地址后未能移除链接');
  log('超链接：清空地址后成功移除');

  // ---- 3. 项目符号与编号样式 ----
  await newLineAtEnd();
  await page.keyboard.type('列表项一');
  await page
    .locator('select[title="项目符号与编号"]')
    .selectOption('bullet:circle');
  await page.waitForTimeout(100);
  let ulStyle = await page
    .locator('.ProseMirror ul')
    .last()
    .getAttribute('style');
  if (!ulStyle || !ulStyle.includes('circle'))
    throw new Error('圆点列表样式未生效: ' + ulStyle);
  log('项目符号：圆点(circle)样式生效 style=' + ulStyle);

  await page
    .locator('select[title="项目符号与编号"]')
    .selectOption('ordered:upper-roman');
  await page.waitForTimeout(100);
  const olStyle = await page
    .locator('.ProseMirror ol')
    .last()
    .getAttribute('style');
  if (!olStyle || !olStyle.includes('upper-roman'))
    throw new Error('大写罗马数字编号样式未生效: ' + olStyle);
  log('编号：大写罗马数字(upper-roman)样式生效 style=' + olStyle);

  // ---- 4. 表格属性：对齐 / 单元格垂直对齐 / 文字方向 / 边框 / 底纹 ----
  await newLineAtEnd();
  await page.getByTitle('▦ 表格', { exact: false }).click();
  await page.waitForSelector('.ProseMirror table', { timeout: 5000 });
  const cells = page.locator('.ProseMirror table td, .ProseMirror table th');
  await cells.first().click();
  await page.waitForSelector('.toolbar-table', { timeout: 3000 });

  await page
    .getByTitle('表格对齐 / 单元格对齐方向 / 文字方向 / 边框底纹')
    .click();
  await page.waitForSelector('.table-props-modal', { timeout: 3000 });
  await page.locator('input[name="tableAlign"][value="center"]').check();
  await page.locator('input[name="cellVAlign"][value="middle"]').check();
  await page.locator('input[name="textDirection"][value="vertical"]').check();
  await page.locator('.table-props-modal select').selectOption('dashed');
  await page.locator('.table-props-number').fill('2');
  await page
    .locator('.table-props-row input[type="color"]')
    .nth(0)
    .fill('#ab2222');
  await page
    .locator('.table-props-row input[type="color"]')
    .nth(1)
    .fill('#ddeeff');
  await page.getByRole('button', { name: '应用' }).click();
  await page.waitForTimeout(200);

  const tableWrapperStyle = await page
    .locator('.ProseMirror .tableWrapper')
    .last()
    .getAttribute('style');
  if (!tableWrapperStyle || !tableWrapperStyle.includes('auto'))
    throw new Error('表格居中对齐未生效: ' + tableWrapperStyle);
  const firstCellStyle = await cells.first().getAttribute('style');
  if (
    !firstCellStyle ||
    !firstCellStyle.includes('vertical-align') ||
    !firstCellStyle.includes('writing-mode') ||
    !firstCellStyle.includes('border') ||
    !firstCellStyle.includes('221')
  ) {
    throw new Error(
      '单元格垂直对齐/文字方向/边框未正确应用: ' + firstCellStyle
    );
  }
  log(
    '表格属性：对齐/垂直对齐/文字方向/边框/底纹全部生效 cellStyle=' +
      firstCellStyle
  );

  // ---- 5. 导出并核对 XML ----
  const downloadPromise = page.waitForEvent('download');
  await page.getByRole('button', { name: /导出 DOCX/ }).click();
  const download = await downloadPromise;
  const outPath = '/tmp/exported-advanced.docx';
  await download.saveAs(outPath);
  const bytes = readFileSync(outPath);
  if (bytes.subarray(0, 2).toString('latin1') !== 'PK')
    throw new Error('导出文件不是合法 docx');
  log('导出完成 ' + outPath + ` size=${bytes.length}`);

  if (pageErrors.length)
    throw new Error('页面运行时异常:\n' + pageErrors.join('\n'));

  console.log('\n✅ ALL ADVANCED CHECKS PASSED');
} catch (err) {
  console.error('\n❌ TEST FAILED:', err.message);
  await page.screenshot({ path: '/tmp/failure-advanced.png', fullPage: true });
  console.error('screenshot saved to /tmp/failure-advanced.png');
  console.error('page errors so far:', pageErrors);
  process.exitCode = 1;
} finally {
  await browser.close();
}
