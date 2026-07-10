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
2. 左侧是**文档目录树**（H1-H9 全部支持，自动生成，点击跳转+高亮）
3. 中间是编辑器，工具栏分三组，尽量贴近 Office「开始」选项卡：
   - **结构行**：正文/标题 1-9 下拉、**项目符号与编号样式下拉**（实心圆点/空心圆/方块、
     数字/小写字母/大写字母/小写罗马/大写罗马数字）、引用、链接（支持编辑已有链接地址、
     清空即可移除）、图片/表格/提示块、批注、面板开关
   - **字体与段落行**：字体、字号、文字颜色、高亮底色、加粗/斜体/下划线/删除线/行内代码、
     左中右两端对齐、增减缩进、行距、**清除格式**（不影响批注与链接，行为对齐 Word 的
     "清除格式"）
   - **表格上下文行**（光标进入表格时才出现）：插入/删除行列、合并/拆分单元格、
     表头行/列切换、单元格底色、**表格属性面板**（表格对齐 / 单元格垂直对齐方向 /
     文字方向水平-垂直 / 边框线型-粗细-颜色 / 底纹）
4. 右侧是**批注面板**（DOCX 原有批注 + 编辑器内新增的批注，点击跳转并高亮精确到被批注的
   那段文字，而不是整段/整块；新增批注后面板会自动滚动到并短暂高亮刚加入的那一条）
5. 切换到"预览"标签页——同样带**目录树**，只读渲染；颜色、字体、字号、对齐方式均来自
   解析出的真实 DOCX 样式，与编辑器完全一致
6. 点击"导出 DOCX"下载编辑结果（颜色/字号/对齐/缩进/行距/批注/表格结构/边框底纹/
   项目符号与编号样式完整回填）
7. 点击"✨ 生成大纲"体验 AI 扩展接口（`applyAIPatch`）

页面全局引入了一份标准的现代 CSS reset（Andy Bell 的 "A (more) Modern CSS Reset"，`src/reset.css`，
未做定制修改），作为所有页面样式的统一起点。

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
 │    ├── EditorPane.tsx                  — 编辑器 + 目录树 + 批注面板 + 高亮 API + 三行工具栏
 │    ├── highlightPlugin.ts              — ProseMirror Decoration 高亮插件
 │    ├── ensureIdsPlugin.ts              — 编辑过程中新节点自动补 blockId/cellId
 │    ├── formatOptions.ts                — 字体/字号/高亮/行距下拉选项
 │    ├── commands.ts / pluginsSetup.ts / tableUtils.ts
 │    └── editor.css
 ├── outline/    OutlineTree.tsx / computeOutline.ts — 文档目录树组件 + 共享的遍历/定位逻辑
 ├── comments/   CommentsPanel.tsx        — 批注面板组件
 ├── preview/    PreviewPane.tsx          — 只读渲染 + 自己的目录树（复用同一份 schema，样式天然一致）
 ├── export/     exportDocx.ts            — JSON(+批注) → docx.js → DOCX Blob
 ├── schema/     index.ts                 — 唯一 Schema：blockId/cellId/styleName/align/indent/lineSpacing + docxStyle/underline/strike/comment marks
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
window.inkflowEditor.highlightBlock('b2');
window.inkflowEditor.highlightCell('c3');
window.inkflowEditor.clearHighlights();
```

### 4. 显示批注内容

解析 `word/comments.xml`（作者/日期/内容）+ `document.xml` 里的
`w:commentRangeStart`/`w:commentRangeEnd`（标出批注锚定的文字范围，映射为新增的
`comment` mark）。`CommentsPanel` 展示所有批注，点击会跳转并**精确高亮被批注的
那段文字本身**（而不是整段/整块——`highlightPlugin` 支持按精确文档区间做行内
高亮装饰，不只是按 `blockId` 整块高亮）。编辑器工具栏的"✍ 批注"按钮支持对选中
文字新增批注（存内存里，随文档一起管理）；新增后如果已有批注很多、新的一条恰好
在可视区域之外，面板会自动滚动到最新一条并做一次短暂高亮，确保"刚刚新增的批注"
肉眼可见，不只是数据层面同步了。导出时会把批注（含新增的）重新写回真实的
`w:commentRangeStart/End` + `w:commentReference` + `word/comments.xml`（用
`docx.js` 的 `Comments` API）。

## Office 级编辑能力

### H1-H9 全部标题层级

`heading` 节点的 `level` 属性本来就没有上限校验（`prosemirror-schema-basic` 只是
"H"+level 拼标签），真正的限制在于 HTML 没有 `<h7>`-`<h9>` 标签。方案：
仍然渲染成 `<h7>`/`<h8>`/`<h9>`（浏览器允许任意标签名），配合 `editor.css` 里的
`display:block` 规则强制块级显示。导出时 1-6 级用 `docx.js` 的具名 `HeadingLevel`
样式（`Heading1`...`Heading6`），7-9 级由于 `docx.js` 没有对应常量、且直接引用
未在生成的 `styles.xml` 里定义的样式 ID 有被 Word 忽略的风险，改为"结构不依赖具名
样式"的直接加粗+递减字号格式兜底——标题层级语义仍完整保存在我们自己的 JSON 模型里
（`level` 属性），只是导出时的视觉呈现方式不同。

### 字体 / 段落格式完全开放

`docxStyle` mark（颜色/字体/字号/高亮）不再只是解析时只读展示，工具栏可以直接
写入：选中文字后改字体/字号/颜色/高亮，会对选区内**逐个文字节点**合并新属性
（保留该节点原有的其它样式，不会把混合格式的选区强行拉平成一份 attrs——这正是
Word「选区内格式不一致，只改一个属性」的行为）；光标无选区时写入 `storedMarks`，
影响接下来要输入的文字。对齐方式、缩进（`increase/decreaseIndent`，在列表项里时
优先走 `sinkListItem`/`liftListItem` 缩进列表层级，否则调整段落 `indent` 属性）、
行距都是直接修改 `paragraph`/`heading` 节点的 attrs（`setNodeMarkup`，不改变节点
大小，可以在一个事务里安全地连续处理选区覆盖到的多个块）。

### 表格编辑完全开放

光标进入表格时，工具栏下方会出现表格上下文行（`isInTable(state)` 驱动的响应式
显示），直接复用 `prosemirror-tables` 的完整命令集：`addRowBefore/After`、
`addColumnBefore/After`、`deleteRow/Column/Table`、`mergeCells`、`splitCell`、
`toggleHeaderRow/Column`、`setCellAttr("background", …)`。单元格合并产生的
`rowspan`/`colspan` 会在导出时回填为 docx.js 的 `rowSpan`/`columnSpan`。

### 编辑过程中新节点的 ID 保障

`blockId`/`cellId` 在解析阶段由 `ooxml.ts` 统一分配，但编辑时新产生的节点（回车
新建的段落、表格插入的新行/列……）默认没有这两个属性。`src/editor/ensureIdsPlugin.ts`
是一个 `appendTransaction` 插件，每次文档变化后自动给缺失 ID 的节点补上，保证目录树
和"通过接口高亮"在自由编辑之后依然可靠，不仅仅在刚解析完的那一刻有效。

### 清除格式

对应 Word 的"清除格式"（橡皮擦图标）：移除选区内所有格式类 mark（加粗/斜体/
下划线/删除线/行内代码/`docxStyle`），并把覆盖到的段落/标题的 `align`/`indent`/
`lineSpacing`/`styleName` 重置为默认值。**批注和链接不算"格式"，予以保留**——
这与 Word 的实际行为一致（Word 的清除格式同样不会把超链接拆掉）。

### 超链接：插入 / 编辑 / 移除

`insertLink` 会先用 `nodesBetween` 扫描选区内的文字节点找现有的 `link` mark：
如果选中的文字已经是链接，弹出的输入框会回填当前地址（可以直接改掉；早期实现
用 `resolve(pos).marks()` 判断"是否已有链接"，但那个 API 在节点边界位置的取值
语义比较微妙，选区恰好等于整个链接范围时经常判断不到——改成扫描文字节点后就
稳定了）。地址清空后确定，等于移除链接。

### 项目符号与编号

一个下拉框覆盖 Word 常见的列表样式：实心圆点 / 空心圆 / 方块（`bulletStyle`），
数字 / 小写字母 / 大写字母 / 小写罗马 / 大写罗马数字（`numberFormat`）。样式存在
`bullet_list`/`ordered_list` 节点的 attrs 上，用 CSS 原生的 `list-style-type`
关键字渲染（`disc`/`circle`/`square`/`decimal`/`lower-alpha`/`upper-alpha`/
`lower-roman`/`upper-roman` 本来就是标准 CSS 值，不需要手工画符号）。导出时每种
样式对应一份独立的 `docx.js` numbering 定义（项目符号用实际 Unicode 字符如
"●"/"○"/"▪"，编号用对应的 `LevelFormat` 枚举）。解析 DOCX 时也会从
`numbering.xml` 尽量还原具体样式（编号格式读 `w:numFmt`；项目符号的具体形状——
圆点/空心圆/方块——Word 是靠 `w:lvlText` 的实际字符区分的，这里做了启发式识别，
认不出的字符统一按实心圆点处理，是一个已知的近似）。

### 表格属性：对齐 / 单元格垂直对齐 / 文字方向 / 边框底纹

光标进入表格后，表格上下文行里的"⚙ 表格属性"按钮会打开一个面板，覆盖：

- 表格对齐（整张表格在页面里靠左/居中/靠右）
- 单元格垂直对齐（顶端/居中/底端，`vertical-align`）
- 文字方向（水平 / 垂直，`writing-mode: vertical-rl`）
- 边框（线型 + 磅值 + 颜色，一次性存成结构化的 `cellBorder` attr，而不是拆成三个
  独立属性各自拼 `style` 字符串互相覆盖）
- 底纹（背景色，复用已有的单元格底色机制）

这几项都落在 `table_cell`/`table_header` 的 attrs 上，导出时映射到 `docx.js` 的
`TableCell.verticalAlign`/`textDirection`/`borders`。

**踩过的一个坑**：表格对齐一开始怎么设置都没有视觉效果。原因是
`columnResizing()` 插件会给 `table` 节点装一个自定义 `NodeView`（`TableView`），
直接用 `document.createElement` 手搭 DOM 来管理列宽拖拽，**完全绕过了 schema 的
`toDOM`**——也就是说 `table.attrs.align` 的渲染逻辑写了但从来没被调用过。
解决方式是继承 `TableView` 写一个 `AlignAwareTableView`（`src/editor/AlignAwareTableView.ts`），
在它自己的列宽计算跑完之后再补上对齐样式；同时因为表格默认宽度是 100%，居中/
靠右在视觉上不会有任何变化，所以对齐生效时顺带让外层 wrapper 收缩到内容宽度
（`width: fit-content`）。这个问题只影响"表格"这一层节点本身——单元格的 valign/
textDirection/border/背景色都是通过 `cellAttributes` 机制走正常的 toDOM，不受
影响；Preview 模式也没用 `columnResizing`，同样不受影响。

## 技术选型

| 能力       | 选型                                                                              |
| ---------- | --------------------------------------------------------------------------------- |
| UI 框架    | Solid.js（细粒度响应式，无虚拟 DOM）                                              |
| 编辑器内核 | `prosemirror-*`（schema/state/view/transform/commands/keymap/history/inputrules） |
| 列表       | `prosemirror-schema-list`                                                         |
| 表格       | `prosemirror-tables`（含列宽拖拽）                                                |
| DOCX 解析  | 自研 OOXML 解析器（`jszip` 解包 + 浏览器 `DOMParser`）                            |
| DOCX 导出  | `docx`（`docx.js`）                                                               |
| 下载       | `file-saver`                                                                      |

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

`scripts/` 目录包含三套端到端冒烟测试（Playwright + Chromium）：

```bash
npm run make-sample          # 生成基础测试用 docx（标题/正文/列表）
npm run make-rich-sample     # 生成含"彩色文字/居中/批注/表格底色"的富样式测试 docx
npx playwright install chromium
npm run build && npm run preview -- --port 4300 &
npm run e2e                  # 基础闭环：上传→编辑→表格→插件节点→AI大纲→预览→导出
npm run e2e:features         # 解析相关能力：目录树跳转、预览样式还原、高亮接口、批注显示与新增、导出回填
npm run e2e:office           # Office 级编辑：H1-H9、字体/颜色/高亮/下划线/删除线、对齐/缩进/行距、
                              # 表格插入/删除行列/合并/拆分/表头/底色，全部校验到导出后的原始 XML
npm run e2e:advanced         # 清除格式、超链接插入/编辑/移除、项目符号与编号样式、
                              # 表格属性（对齐/垂直对齐/文字方向/边框/底纹），同样校验到导出 XML
```

四套测试均已通过，并额外用 Python 校验过导出文件的原始 XML（`word/comments.xml`
真实包含批注、`word/document.xml` 真实包含 `w:color`/`w:sz`/`w:u`/`w:strike`/
`w:jc`/`w:ind`/`w:spacing`/`w:gridSpan`/`w:tcBorders`/`w:vAlign`/`w:textDirection`/
`commentRangeStart` 等，`word/numbering.xml` 真实包含对应的 `w:numFmt`/项目符号
字符），确认不是"看起来对但导出是空壳"。

调试这套测试时踩过一个值得记录的坑：Playwright 的合成键盘/鼠标事件如果不加任何
间隔连续触发，可能跑在 ProseMirror 把浏览器原生选区同步回自身模型之前，导致
"点击 A 位置、按 Enter"这类操作实际作用在了旧的选区位置上。真实用户的手速远
达不到这个量级，所以这只是自动化测试要考虑的时序问题，不是应用本身的 bug——
但排查过程本身值得记录：加一点点等待（`e2e:office` 里的 `moveCursorToDocEnd`
辅助函数）就能稳定复现正确结果。

## 已知限制

- **还原误差**：页眉页脚、分栏、修订标记的精确渲染、脚注/尾注内容、复杂编号重启逻辑不处理，
  遵循"结构优先、样式次之、还原度最后"原则（§2.3）。
- **表格**：支持 `gridSpan`（合并列）、`rowSpan`（合并行，来自 `prosemirror-tables` 编辑时
  产生的 rowspan 属性）与单元格底色；不支持单元格内嵌套的复杂垂直合并继承样式。
- **H7-H9 标题**：由于 docx.js 没有对应的具名样式常量，导出时使用直接加粗+字号格式兜底，
  而不是引用 Word 的 `Heading7`-`Heading9` 样式——如果原始 DOCX 本身就带这几级标题的具名样式，
  解析时已经把该样式的颜色/字号还原到了 `docxStyle` mark 上，因此往返导出的视觉效果基本一致，
  只是不通过"具名样式引用"这条路径。
- **批注范围**：假定批注锚点不跨段落（真实 Word 文档里跨段落批注较少见）。
- **修订标记**：`w:ins` 内容展开为正常文本，`w:del` 内容直接丢弃，不保留标记状态。
- **项目符号形状识别**：解析已有 DOCX 时，编号格式（数字/字母/罗马数字）读取准确；
  但项目符号具体是圆点/空心圆/方块，Word 内部靠 `w:lvlText` 的实际字符区分，这里
  只做了启发式识别，认不出的字符统一按实心圆点处理。
- **协同编辑**：未实现，是预留而非已完成的能力（`buildEditorPlugins` 预留了
  `proseMirrorPlugins` 扩展点，接入 `y-prosemirror` 时不需要改动 Editor 模块本身）。

## AI 扩展接口的真实接入方式

`src/utils/applyAIPatch.ts` 默认是本地示例（生成大纲），不发网络请求。接入真实 LLM
只需替换其内部实现（输入输出都是 ProseMirror JSON），Editor/App 层不需要改动。
