import { chromium } from "playwright";
import { readFileSync, writeFileSync } from "fs";

const BASE = process.env.BASE_URL || "http://localhost:4300";
const consoleErrors = [];
const pageErrors = [];

const browser = await chromium.launch({
  executablePath: "/opt/pw-browsers/chromium-1194/chrome-linux/chrome",
});
const page = await browser.newPage();
page.on("console", (msg) => {
  if (msg.type() === "error") consoleErrors.push(msg.text());
});
page.on("pageerror", (err) => pageErrors.push(String(err)));

const log = (msg) => console.log("• " + msg);

try {
  await page.goto(BASE, { waitUntil: "networkidle" });
  log("page loaded");

  // ---- 1. Upload real docx ----
  const fileInput = page.locator('input[type="file"]');
  await fileInput.setInputFiles("/tmp/sample.docx");
  log("uploaded sample.docx");

  await page.waitForSelector(".editor-shell", { timeout: 15000 });
  log("editor mounted after parse");

  const bodyText = await page.locator(".ProseMirror").innerText();
  if (!bodyText.includes("测试文档标题")) throw new Error("解析结果缺少标题文本: " + bodyText.slice(0, 200));
  if (!bodyText.includes("加粗")) throw new Error("解析结果缺少正文文本");
  if (!bodyText.includes("第一条列表项")) throw new Error("解析结果缺少列表文本");
  log("parsed DOCX -> ProseMirror JSON content verified (heading/paragraph/list present)");

  // ---- 2. Edit: place cursor at end, type new heading + paragraph ----
  await page.locator(".ProseMirror").click();
  await page.keyboard.press("Control+End");
  await page.keyboard.press("Enter");
  await page.locator(".toolbar-select-heading").selectOption("h2");
  await page.keyboard.type("插入的新章节");
  await page.keyboard.press("Enter");
  await page.locator(".toolbar-select-heading").selectOption("paragraph");
  await page.keyboard.type("这是通过编辑器新增的段落，用来验证 transaction 流程。");
  log("typed new heading + paragraph via toolbar select + keyboard");

  // bold toggle test
  await page.keyboard.press("Enter");
  await page.keyboard.down("Shift");
  // select nothing special; just toggle bold on new text
  await page.keyboard.up("Shift");
  await page.getByTitle("加粗").click();
  await page.keyboard.type("加粗测试文字");
  await page.getByTitle("加粗").click();
  log("toggled bold mark and typed text");

  // ---- 3. Insert table ----
  await page.keyboard.press("Enter");
  await page.getByTitle("▦ 表格", { exact: false }).click();
  await page.waitForSelector(".ProseMirror table", { timeout: 5000 });
  log("inserted table node");

  // ---- 4. Insert callout via plugin toolbar button ----
  await page.keyboard.press("Control+End");
  await page.keyboard.press("Enter");
  const calloutBtn = page.locator(".toolbar-btn", { hasText: "提示块" });
  await calloutBtn.click();
  await page.keyboard.type("这是插件系统提供的 Callout 节点。");
  await page.waitForSelector(".ProseMirror .callout", { timeout: 5000 });
  log("inserted plugin-provided callout node");

  // ---- 5. AI outline patch ----
  await page.getByRole("button", { name: /生成大纲/ }).click();
  await page.waitForTimeout(500);
  const afterAI = await page.locator(".ProseMirror").innerText();
  if (!afterAI.includes("自动生成的大纲")) throw new Error("AI 扩展接口未生效");
  log("applyAIPatch (AI extension interface) generated outline block");

  // ---- 6. Preview mode ----
  await page.getByRole("tab", { name: "预览" }).click();
  await page.waitForSelector(".preview-page .ProseMirror", { timeout: 5000 });
  const previewText = await page.locator(".preview-page .ProseMirror").innerText();
  if (!previewText.includes("插入的新章节")) throw new Error("预览内容与编辑内容不一致");
  log("preview mode mirrors edited content");
  await page.getByRole("tab", { name: "编辑" }).click();

  // ---- 7. Export DOCX and validate binary ----
  const downloadPromise = page.waitForEvent("download");
  await page.getByRole("button", { name: /导出 DOCX/ }).click();
  const download = await downloadPromise;
  const outPath = "/tmp/exported.docx";
  await download.saveAs(outPath);
  log("export triggered, file saved to " + outPath);

  const bytes = readFileSync(outPath);
  const magic = bytes.subarray(0, 2).toString("latin1");
  if (magic !== "PK") throw new Error("导出文件不是有效的 zip/docx (magic=" + magic + ")");
  if (bytes.length < 1000) throw new Error("导出文件过小，可能为空文档: " + bytes.length + " bytes");
  log(`exported DOCX is a valid zip (PK header), size=${bytes.length} bytes`);

  if (pageErrors.length) {
    throw new Error("页面运行时出现异常:\n" + pageErrors.join("\n"));
  }
  const seriousConsoleErrors = consoleErrors.filter((e) => !/favicon/i.test(e));
  if (seriousConsoleErrors.length) {
    console.log("⚠ console errors (non-fatal, review):\n" + seriousConsoleErrors.join("\n"));
  }

  console.log("\n✅ ALL CHECKS PASSED");
} catch (err) {
  console.error("\n❌ TEST FAILED:", err.message);
  await page.screenshot({ path: "/tmp/failure.png", fullPage: true });
  console.error("screenshot saved to /tmp/failure.png");
  console.error("console errors so far:", consoleErrors);
  console.error("page errors so far:", pageErrors);
  process.exitCode = 1;
} finally {
  await browser.close();
}
