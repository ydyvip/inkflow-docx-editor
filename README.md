# InkFlow — 在线 DOCX 编辑系统

一个以 **ProseMirror JSON** 为唯一真相源的结构化 DOCX 在线编辑系统。
`DOCX → JSON → Editor → JSON → DOCX` 单向数据流，不使用 HTML 作为最终数据结构，
不直接操作 DOCX 二进制。UI 层基于 **Solid.js**。

## 快速开始

```bash
npm install
npm run dev        # 开发模式，默认 http://localhost:5173
npm run build       # 生产构建（tsc + vite build）
npm run preview     # 预览生产构建
```

打开页面后：
1. 拖拽 / 选择一个 `.docx` 文件上传，或点击"从空白文档开始"
2. 左侧是**文档目录树**（根据标题自动生成，点击跳转）
3. 中间是编辑器：工具栏支持标题 / 加粗 / 列表 / 表格 / 链接 / 图片 / 提示块 / **批注**
4. 右侧是**批注面板**（DOCX 原有批注 + 编辑器内新增的批注，点击跳转并高亮）
5. 切换到"预览"标签页查看只读渲染——颜色、字体、字号、对齐方式均来自解析出的真实 DOCX 样式
6. 点击"导出 DOCX"下载编辑结果（含批注、颜色、字号、对齐方式的完整回填）
7. 点击"✨ 生成大纲"体验 AI 扩展接口（`applyAIPatch`）

## ⚠️ 与最初方案的一个重要差异：解析器不再基于 mammoth.js

最初的方案要求用 mammoth.js 做 DOCX 解析。但后续新增的四个需求——

- 目录树需要标题的稳定结构
- **预览样式必须来自 DOCX 解析出的真实样式**（颜色/字体/字号/对齐）
- **通过接口高亮指定段落/单元格** 需要稳定可寻址的 ID
- **显示批注内容** 需要 `word/comments.xml` + `commentRangeStart/End`

——和 mammoth 的设计哲学（"语义优先，直接丢弃颜色/字体等视觉样式，且不解析批注"）根本冲突。
mammoth 的 HTML 中间态里没有这些信息，事后也补不回来。

因此解析器被替换为**直接读取 DOCX 原始 XML 的自研 OOXML 解析器**
（`src/parser/ooxml.ts`，用 `JSZip` 解包 + 浏览器原生 `DOMParser` 遍历
`word/document.xml` / `styles.xml` / `numbering.xml` / `comments.xml` / 关系文件），
不再经过 mammoth 或任何 HTML 中间态。`mammoth` 依赖已从 `package.json` 移除。
`DOCX → JSON` 的单向数据流原则（§2.2）依然成立，只是"怎么产出 JSON"这一步换了实现。

## 目录结构

```
src/
 ├── upload/     UploadPanel.tsx          — 拖拽/选择、.docx 校验、转 ArrayBuffer
 ├── parser/
 │    ├── parseDocx.ts                    — 编排：File → Worker → JSON + 批注
 │    └── ooxml.ts                        — 核心：直接解析 DOCX 原始 XML（见上）
 ├── editor/
 │    ├── EditorPane.tsx                  — 编辑器 + 目录树 + 批注面板 + 高亮 API
 │    ├── highlightPlugin.ts              — ProseMirror Decoration 高亮插件
 │    ├── commands.ts / pluginsSetup.ts / tableUtils.ts
 │    └── editor.css
 ├── outline/    OutlineTree.tsx          — 文档目录树组件
 ├── comments/   CommentsPanel.tsx        — 批注面板组件
 ├── preview/    PreviewPane.tsx          — 只读渲染（复用同一份 schema，样式天然一致）
 ├── export/     exportDocx.ts            — JSON(+批注) → docx.js → DOCX Blob
 ├── schema/     index.ts                 — 唯一 Schema：blockId/cellId/styleName/align + docxStyle/comment marks
 ├── plugins/    registry.ts / calloutPlugin.ts — 插件系统（Node/Mark/InputRule/Keymap）
 ├── worker/     parse.worker.ts          — 大文件 Worker 化解析（zip 解包 + XML 遍历都在 worker 里）
 └── utils/      applyAIPatch.ts          — AI 扩展接口预留
```

## 四个新增能力的实现方式

### 1. 文档目录树
`EditorPane` 每次 transaction 后遍历 `view.state.doc`，收集所有 `heading` 节点的
`level` + 文本 + `blockId`，交给 `OutlineTree` 渲染。点击条目通过 `blockId` 在文档里
定位对应节点位置，设置光标并 `scrollIntoView()`。

### 2. 预览样式来自解析的 DOCX 样式
`ooxml.ts` 解析每个 run 的 `w:rPr`（颜色/字体/字号/高亮底色），连同段落的
`w:pStyle`（命名样式，解析 `styles.xml` 里的继承链得到最终颜色/字号）一并写入
一个新 mark：`docxStyle`。这个 mark 的 `toDOM` 直接输出对应的 inline
`style="color:...;font-family:...;font-size:...pt"`。由于编辑器和预览用的是
**同一份 schema**，两边会自动渲染出一致的、真实还原的样式（inline style 的 CSS
优先级天然高于本应用主题的类选择器）。段落级的 `align`（对齐方式）同理。

### 3. 通过接口高亮段落 / 单元格
`schema` 里给 `paragraph`/`heading` 增加了 `blockId`，给 `table_cell`/`table_header`
增加了 `cellId`（解析时分配的稳定 ID，渲染时同时输出为 `data-block-id`/`data-cell-id`
方便调试）。`src/editor/highlightPlugin.ts` 是一个不修改文档内容、只叠加渲染层
Decoration 的 ProseMirror 插件。`EditorPane` 通过 `onReady` 回调暴露一组 API：

```ts
interface EditorApi {
  highlightBlock(blockId: string): void;
  highlightCell(cellId: string): void;
  clearHighlights(): void;
  scrollToBlock(id: string): void; // 定位 + 高亮
  getComments(): CommentItem[];
}
```

为方便直接调试/集成，同一个 API 对象也挂在 `window.inkflowEditor` 上，可以直接在
控制台里调用，例如：

```js
window.inkflowEditor.highlightBlock("b2");
window.inkflowEditor.highlightCell("c3");
window.inkflowEditor.clearHighlights();
```

### 4. 显示批注内容
解析 `word/comments.xml`（作者/日期/内容）+ `document.xml` 里的
`w:commentRangeStart`/`w:commentRangeEnd`（标出批注锚定的文字范围，映射为新增的
`comment` mark）。`CommentsPanel` 展示所有批注，点击可跳转并高亮锚点所在段落。
编辑器工具栏的"✍ 批注"按钮支持对选中文字新增批注（存内存里，随文档一起管理）。
导出时会把批注（含新增的）重新写回真实的 `w:commentRangeStart/End` +
`w:commentReference` + `word/comments.xml`（用 `docx.js` 的 `Comments` API）。

## 技术选型

| 能力 | 选型 |
| --- | --- |
| UI 框架 | Solid.js（细粒度响应式，无虚拟 DOM）|
| 编辑器内核 | `prosemirror-*`（schema/state/view/transform/commands/keymap/history/inputrules）|
| 列表 | `prosemirror-schema-list` |
| 表格 | `prosemirror-tables`（含列宽拖拽）|
| DOCX 解析 | 自研 OOXML 解析器（`jszip` 解包 + 浏览器 `DOMParser`）|
| DOCX 导出 | `docx`（`docx.js`）|
| 下载 | `file-saver` |

## 数据流

```
上传 .docx
   → File.arrayBuffer()
   → Worker: JSZip 解包 + 遍历 document.xml/styles.xml/numbering.xml/comments.xml/关系文件
   → ProseMirror JSON（含 blockId/cellId/styleName/align + docxStyle/comment marks）+ 批注列表
                                            ← 系统唯一真相源
   → EditorState.create({ schema, doc })
   → 每次 transaction 后 doc.toJSON() 上抛给 App；目录树/工具栏状态随 version 信号重算
   → 预览 = 用同一份 JSON 新建一个 readonly EditorView（同一 schema ⇒ 样式天然一致）
   → 导出 = 遍历 JSON，映射为 docx.js 对象树（含颜色/字号/对齐/批注范围）→ Packer.toBlob()
```

## 插件系统

`src/plugins/registry.ts` 定义了 `DocxPlugin` 接口：`nodes`/`marks`（Node/Mark plugin）、
`inputRules`、`keymap`、`toolbar`、`exportNode`。`src/plugins/calloutPlugin.ts` 是完整示例
（提示块 / callout），演示了全部扩展点。新增插件需在 `src/schema/index.ts` 里
`pluginRegistry.register(yourPlugin)`（必须在 schema 构建之前完成）。

## 质量验证

`scripts/` 目录包含两套端到端冒烟测试（Playwright + Chromium）：

```bash
npm run make-sample          # 生成基础测试用 docx（标题/正文/列表）
npm run make-rich-sample     # 生成含"彩色文字/居中/批注/表格底色"的富样式测试 docx
npx playwright install chromium
npm run build && npm run preview -- --port 4300 &
npm run e2e                  # 基础闭环：上传→编辑→表格→插件节点→AI大纲→预览→导出
npm run e2e:features         # 新增能力：目录树跳转、预览样式还原、高亮接口、批注显示与新增、导出回填
```

两套测试均已通过，并额外用 Python 校验过导出文件的原始 XML（`word/comments.xml`
真实包含批注、`word/document.xml` 真实包含 `w:color`/`w:sz`/`commentRangeStart`/
`w:jc="center"`），确认不是"看起来对但导出是空壳"。

## 已知限制

- **还原误差**：页眉页脚、分栏、修订标记的精确渲染、脚注/尾注内容、复杂编号重启逻辑不处理，
  遵循"结构优先、样式次之、还原度最后"原则（§2.3）。
- **表格**：支持 `gridSpan`（合并列）与单元格底色，不支持 `vMerge`（合并行）。
- **批注范围**：假定批注锚点不跨段落（真实 Word 文档里跨段落批注较少见）。
- **修订标记**：`w:ins` 内容展开为正常文本，`w:del` 内容直接丢弃，不保留标记状态。
- **协同编辑**：未实现，是预留而非已完成的能力（`buildEditorPlugins` 预留了
  `proseMirrorPlugins` 扩展点，接入 `y-prosemirror` 时不需要改动 Editor 模块本身）。

## AI 扩展接口的真实接入方式

`src/utils/applyAIPatch.ts` 默认是本地示例（生成大纲），不发网络请求。接入真实 LLM
只需替换其内部实现（输入输出都是 ProseMirror JSON），Editor/App 层不需要改动。
