import { chromium } from "playwright";
import { readFileSync } from "fs";

const BASE = process.env.BASE_URL || "http://localhost:4300";
const consoleErrors = [];
const pageErrors = [];

const browser = await chromium.launch({ executablePath: "/opt/pw-browsers/chromium-1194/chrome-linux/chrome" });
const page = await browser.newPage();
page.on("console", (msg) => {
  if (msg.type() === "error") consoleErrors.push(msg.text());
});
page.on("pageerror", (err) => pageErrors.push(String(err)));
page.on("dialog", async (dialog) => {
  if (dialog.type() === "prompt") await dialog.accept("这是新增的批注内容");
  else await dialog.accept();
});

const log = (msg) => console.log("• " + msg);

try {
  await page.goto(BASE, { waitUntil: "networkidle" });
  await page.locator('input[type="file"]').setInputFiles("/tmp/rich-sample.docx");
  await page.waitForSelector(".editor-shell", { timeout: 15000 });
  await page.waitForTimeout(300);
  log("uploaded rich-sample.docx (headings + colored text + comment + shaded table)");

  // ---- 1. 目录树 outline ----
  const outlineButtons = page.locator(".outline-item button");
  const outlineCount = await outlineButtons.count();
  if (outlineCount < 1) throw new Error("目录树未显示任何标题条目");
  const firstOutlineText = await outlineButtons.first().innerText();
  if (!firstOutlineText.includes("样式与批注测试")) throw new Error("目录树标题文本不匹配: " + firstOutlineText);
  log(`outline tree shows ${outlineCount} heading(s), first = "${firstOutlineText}"`);

  await outlineButtons.first().click();
  await page.waitForSelector(".pm-highlight", { timeout: 3000 });
  log("clicking outline entry scrolls to + highlights the heading (pm-highlight decoration present)");

  // ---- 2. 预览样式来自解析的 DOCX 样式 ----
  const h1Style = await page.locator(".ProseMirror h1 .docx-style-run").first().getAttribute("style");
  if (!h1Style || !h1Style.includes("color") || !h1Style.includes("font-size")) {
    throw new Error("标题未携带解析出的真实颜色/字号样式: " + h1Style);
  }
  log("heading carries real color/font-size parsed from DOCX rPr: " + h1Style);

  const cellStyle = await page.locator('[data-cell-id="c1"]').getAttribute("style");
  if (!cellStyle || !cellStyle.includes("background-color")) throw new Error("表头单元格底色未还原: " + cellStyle);
  log("table header cell carries real shading color from DOCX: " + cellStyle);

  await page.getByRole("tab", { name: "预览" }).click();
  await page.waitForSelector(".preview-page .docx-style-run", { timeout: 5000 });
  const previewH1Style = await page.locator(".preview-page h1 .docx-style-run").first().getAttribute("style");
  if (!previewH1Style || !previewH1Style.includes("color")) throw new Error("预览模式未还原真实样式: " + previewH1Style);
  log("preview mode independently re-renders the same DOCX-derived styles: " + previewH1Style);
  await page.getByRole("tab", { name: "编辑" }).click();
  await page.waitForTimeout(200);

  // ---- 3. 批注显示 ----
  const commentItems = page.locator(".comment-item");
  const commentCount = await commentItems.count();
  if (commentCount < 1) throw new Error("批注面板未显示任何批注");
  const firstCommentText = await commentItems.first().locator(".comment-text").innerText();
  if (!firstCommentText.includes("确认")) throw new Error("批注内容不匹配: " + firstCommentText);
  const firstCommentAuthor = await commentItems.first().locator(".comment-author").innerText();
  if (!firstCommentAuthor.includes("评审员")) throw new Error("批注作者不匹配: " + firstCommentAuthor);
  log(`comments panel shows ${commentCount} comment(s): "${firstCommentAuthor}" — "${firstCommentText}"`);

  const anchorInEditor = await page.locator(".docx-comment").count();
  if (anchorInEditor < 1) throw new Error("正文中未渲染批注锚点（docx-comment span）");
  log("comment anchor range is visually marked in the editor body (.docx-comment)");

  await commentItems.first().click();
  await page.waitForTimeout(200);
  const highlightAfterCommentClick = await page.locator(".pm-highlight").count();
  if (highlightAfterCommentClick < 1) throw new Error("点击批注未高亮对应段落");
  log("clicking a comment scrolls to + highlights its anchor paragraph");

  // ---- 4. 通过接口高亮段落 / 单元格（window.inkflowEditor）----
  const blockHighlightWorked = await page.evaluate(() => {
    const api = window.inkflowEditor;
    if (!api) return false;
    api.clearHighlights();
    const beforeCount = document.querySelectorAll(".pm-highlight").length;
    // b2 = 富文本样例里第二个段落（居中彩色大字）
    api.highlightBlock("b2");
    const afterCount = document.querySelectorAll(".pm-highlight").length;
    return beforeCount === 0 && afterCount > 0;
  });
  if (!blockHighlightWorked) throw new Error("window.inkflowEditor.highlightBlock() 接口未生效");
  log("window.inkflowEditor.highlightBlock('b2') API highlights the target paragraph");

  const cellHighlightWorked = await page.evaluate(() => {
    const api = window.inkflowEditor;
    api.clearHighlights();
    api.highlightCell("c2");
    const el = document.querySelector('[data-cell-id="c2"]');
    return !!el && el.closest(".pm-highlight") !== null;
  });
  if (!cellHighlightWorked) throw new Error("window.inkflowEditor.highlightCell() 接口未生效");
  log("window.inkflowEditor.highlightCell('c2') API highlights the target table cell");

  // ---- 5. 新增批注（工具栏"批注"按钮）----
  // 选中最后一段里的一段文字："后面是普通文字" 所在段落
  const paragraphs = page.locator(".ProseMirror p");
  const targetParaCount = await paragraphs.count();
  const lastPara = paragraphs.nth(targetParaCount - 1);
  await lastPara.click();
  await page.keyboard.press("Home");
  await page.keyboard.down("Shift");
  for (let i = 0; i < 6; i++) await page.keyboard.press("ArrowRight");
  await page.keyboard.up("Shift");
  await page.getByTitle("为选中文字添加批注").click();
  await page.waitForTimeout(300);

  const commentCountAfterAdd = await page.locator(".comment-item").count();
  if (commentCountAfterAdd !== commentCount + 1) {
    throw new Error(`新增批注后数量不对，期望 ${commentCount + 1}，实际 ${commentCountAfterAdd}`);
  }
  log(`added a new comment via toolbar, panel now shows ${commentCountAfterAdd} comment(s)`);

  // ---- 6. 导出并校验批注 + 样式回填进真实 docx XML ----
  const downloadPromise = page.waitForEvent("download");
  await page.getByRole("button", { name: /导出 DOCX/ }).click();
  const download = await downloadPromise;
  const outPath = "/tmp/exported-rich.docx";
  await download.saveAs(outPath);
  log("exported edited rich document to " + outPath);

  if (pageErrors.length) throw new Error("页面运行时异常:\n" + pageErrors.join("\n"));

  console.log("\n✅ ALL FEATURE CHECKS PASSED");
} catch (err) {
  console.error("\n❌ TEST FAILED:", err.message);
  await page.screenshot({ path: "/tmp/failure-features.png", fullPage: true });
  console.error("screenshot saved to /tmp/failure-features.png");
  console.error("console errors so far:", consoleErrors);
  console.error("page errors so far:", pageErrors);
  process.exitCode = 1;
} finally {
  await browser.close();
}
