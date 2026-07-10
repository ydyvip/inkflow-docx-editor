/**
 * Schema 模块
 * ------------------------------------------------------------
 * 系统的"唯一真相源"不是某个具体文件，而是由本模块产出的
 * ProseMirror Schema —— 它定义了 JSON 文档模型允许出现的所有
 * Node / Mark（§2.1 统一数据模型原则）。
 *
 * 相比初版，本模块扩展了三类信息，全部来自对 DOCX 原始 XML 的解析
 * （见 src/parser/ooxml.ts），而不是 mammoth 的语义化 HTML 中间态：
 *
 *   1. 块级节点新增 blockId（paragraph/heading）/ cellId（表格单元格）
 *      —— 稳定的可寻址 ID，供"通过接口高亮段落/单元格"使用
 *   2. 块级节点新增 styleName / align —— 保留 Word 命名样式与对齐方式
 *   3. 新增 underline / strike / docxStyle / comment 四个 Mark
 *      —— docxStyle 携带 color/fontFamily/fontSize，真正做到
 *         "预览样式来自解析的 DOCX 样式"；comment 携带批注 ID，
 *         用于渲染批注高亮和批注面板联动
 *
 * 组成：
 *   1. prosemirror-schema-basic  → paragraph / heading / text /
 *                                    image / hard_break / marks(strong,em,link,code)
 *   2. prosemirror-schema-list   → bullet_list / ordered_list / list_item
 *   3. prosemirror-tables        → table / table_row / table_cell / table_header
 *   4. 插件注册的自定义节点（如 callout）
 */
import { Schema } from 'prosemirror-model';
import type { NodeSpec, MarkSpec } from 'prosemirror-model';
import { schema as basicSchema } from 'prosemirror-schema-basic';
import { addListNodes } from 'prosemirror-schema-list';
import { tableNodes } from 'prosemirror-tables';
import { pluginRegistry } from '../plugins/registry';
import { calloutPlugin } from '../plugins/calloutPlugin';

// 注册内置插件（应用启动时一次性完成，早于 schema 构建）
pluginRegistry.register(calloutPlugin);

export interface CellBorder {
  style: 'none' | 'single' | 'dashed' | 'double';
  width: number; // 磅（pt）
  color: string; // #rrggbb
}

// ---- 1. 给 paragraph / heading 扩展 blockId / styleName / align / indent / lineSpacing ----

function withBlockAttrs(
  base: NodeSpec,
  extraParseDOM: NodeSpec['parseDOM'] = []
): NodeSpec {
  return {
    ...base,
    attrs: {
      ...(base.attrs ?? {}),
      blockId: { default: null },
      styleName: { default: null },
      align: { default: null },
      indent: { default: 0 }, // 缩进级别 0-8，一级约等于 Word "增加缩进" 一次（0.5in / 720 twips）
      lineSpacing: { default: null }, // 行距倍数，如 1 / 1.15 / 1.5 / 2
    },
    parseDOM: [...(base.parseDOM ?? []), ...extraParseDOM],
    toDOM(node) {
      const baseArr = base.toDOM
        ? (base.toDOM(node) as unknown as any[])
        : ['p', 0];
      const tag = baseArr[0];
      const domAttrs: Record<string, string> = {};
      if (node.attrs.blockId) domAttrs['data-block-id'] = node.attrs.blockId;
      if (node.attrs.styleName)
        domAttrs['data-style-name'] = node.attrs.styleName;

      const style: string[] = [];
      if (node.attrs.align && node.attrs.align !== 'left')
        style.push(`text-align:${node.attrs.align}`);
      if (node.attrs.indent)
        style.push(`margin-left:${Number(node.attrs.indent) * 36}pt`);
      if (node.attrs.lineSpacing)
        style.push(`line-height:${node.attrs.lineSpacing}`);
      if (style.length) domAttrs.style = style.join(';');

      return [tag, domAttrs, 0];
    },
  };
}

// 标题 7-9 没有原生 HTML 标签（<h7>/<h8>/<h9> 不是标准元素），
// 但浏览器允许渲染任意标签名；配合 editor.css 里的 display:block 规则即可正常显示为块级标题。
const HEADING_EXTRA_PARSE_DOM = [
  { tag: 'h7', attrs: { level: 7 } },
  { tag: 'h8', attrs: { level: 8 } },
  { tag: 'h9', attrs: { level: 9 } },
];

const paragraphSpec = withBlockAttrs(basicSchema.spec.nodes.get('paragraph')!);
const headingSpec = withBlockAttrs(
  basicSchema.spec.nodes.get('heading')!,
  HEADING_EXTRA_PARSE_DOM
);

const nodesWithBlockIds = basicSchema.spec.nodes
  .update('paragraph', paragraphSpec)
  .update('heading', headingSpec);

// 1. 基础 nodes + 列表 nodes
const nodesWithLists = addListNodes(
  nodesWithBlockIds,
  'paragraph block*',
  'block'
);

// ---- 1b. 项目符号 / 编号样式（§ 项目符号与编号）----
// bullet_list 新增 bulletStyle（对应 CSS list-style-type 的 disc/circle/square），
// ordered_list 新增 numberFormat（decimal/lower-alpha/upper-alpha/lower-roman/upper-roman，
// 与 CSS list-style-type 关键字一致，导出时映射到 docx.js 的 LevelFormat）。
function withListStyleAttr(
  base: NodeSpec,
  attrName: string,
  defaultValue: string
): NodeSpec {
  return {
    ...base,
    attrs: { ...(base.attrs ?? {}), [attrName]: { default: defaultValue } },
    toDOM(node) {
      const baseArr = (base.toDOM
        ? base.toDOM(node)
        : ['ul', 0]) as unknown as any[];
      const tag = baseArr[0];
      const hole = baseArr[baseArr.length - 1];
      const domAttrs: Record<string, any> =
        baseArr.length === 3 ? { ...baseArr[1] } : {};
      const styleVal = node.attrs[attrName];
      if (styleVal && styleVal !== defaultValue) {
        domAttrs.style =
          (domAttrs.style ? domAttrs.style + ';' : '') +
          `list-style-type:${styleVal}`;
      }
      return [tag, domAttrs, hole];
    },
  };
}

const bulletListSpec = withListStyleAttr(
  nodesWithLists.get('bullet_list')!,
  'bulletStyle',
  'disc'
);
const orderedListSpec = withListStyleAttr(
  nodesWithLists.get('ordered_list')!,
  'numberFormat',
  'decimal'
);
const nodesWithListStyles = nodesWithLists
  .update('bullet_list', bulletListSpec)
  .update('ordered_list', orderedListSpec);

// 2. 表格 nodes（§6.4 表格必须插件化实现）
//    - table 新增 align（表格在页面中的左/中/右对齐）
//    - 单元格新增 cellId / background（已有）+ valign（垂直对齐）/ textDirection（文字方向）/
//      cellBorder（边框：style+width+color 一体存储，避免多属性拼 style 字符串时互相覆盖）
const rawTableNodes = tableNodes({
  tableGroup: 'block',
  cellContent: 'block+',
  cellAttributes: {
    background: {
      default: null,
      getFromDOM: (dom) => (dom as HTMLElement).style.backgroundColor || null,
      setDOMAttr: (value, attrs) => {
        if (value)
          attrs.style = (attrs.style || '') + `background-color: ${value};`;
      },
    },
    cellId: {
      default: null,
      getFromDOM: (dom) => (dom as HTMLElement).getAttribute('data-cell-id'),
      setDOMAttr: (value, attrs) => {
        if (value) attrs['data-cell-id'] = value;
      },
    },
    valign: {
      default: null, // 'middle' | 'bottom' | null(=top，不写内联样式)
      getFromDOM: (dom) => {
        const v = (dom as HTMLElement).style.verticalAlign;
        return v === 'middle' || v === 'bottom' ? v : null;
      },
      setDOMAttr: (value, attrs) => {
        if (value && value !== 'top')
          attrs.style = (attrs.style || '') + `vertical-align:${value};`;
      },
    },
    textDirection: {
      default: null, // 'vertical' | null(=horizontal)
      getFromDOM: (dom) =>
        (dom as HTMLElement).getAttribute('data-text-direction'),
      setDOMAttr: (value, attrs) => {
        if (value === 'vertical') {
          attrs.style =
            (attrs.style || '') +
            `writing-mode:vertical-rl;text-orientation:upright;`;
          attrs['data-text-direction'] = 'vertical';
        }
      },
    },
    cellBorder: {
      default: null as CellBorder | null,
      setDOMAttr: (value: any, attrs: Record<string, any>) => {
        const border = value as CellBorder | null;
        if (border && border.style && border.style !== 'none') {
          const cssStyle =
            border.style === 'double'
              ? 'double'
              : border.style === 'dashed'
                ? 'dashed'
                : 'solid';
          attrs.style =
            (attrs.style || '') +
            `border:${border.width || 1}pt ${cssStyle} ${border.color || '#000000'};`;
        }
      },
    },
  },
});

const tableSpecWithAlign: NodeSpec = {
  ...rawTableNodes.table,
  attrs: { ...(rawTableNodes.table.attrs ?? {}), align: { default: null } },
  toDOM(node) {
    const domAttrs: Record<string, string> = {};
    if (node.attrs.align === 'center')
      domAttrs.style = 'margin-left:auto;margin-right:auto;';
    else if (node.attrs.align === 'right')
      domAttrs.style = 'margin-left:auto;margin-right:0;';
    return ['table', domAttrs, ['tbody', 0]];
  },
};

const nodesWithTables = nodesWithListStyles.append({
  ...rawTableNodes,
  table: tableSpecWithAlign,
});

// 3. 插件自定义 nodes
const nodesWithPlugins = nodesWithTables.append(pluginRegistry.nodes());

// ---- 4. 新增 Mark：underline / strike / docxStyle / comment ----

const underlineMark: MarkSpec = {
  parseDOM: [{ tag: 'u' }, { style: 'text-decoration=underline' }],
  toDOM: () => ['u', 0],
};

const strikeMark: MarkSpec = {
  parseDOM: [
    { tag: 's' },
    { tag: 'strike' },
    { style: 'text-decoration=line-through' },
  ],
  toDOM: () => ['s', 0],
};

const HIGHLIGHT_COLORS: Record<string, string> = {
  yellow: '#FFFF00',
  green: '#00FF00',
  cyan: '#00FFFF',
  magenta: '#FF00FF',
  blue: '#0000FF',
  red: '#FF0000',
  darkBlue: '#00008B',
  darkCyan: '#008B8B',
  darkGreen: '#006400',
  darkMagenta: '#8B008B',
  darkRed: '#8B0000',
  darkYellow: '#808000',
  darkGray: '#A9A9A9',
  lightGray: '#D3D3D3',
  black: '#000000',
  white: '#FFFFFF',
};

/**
 * docxStyle：承载从 OOXML runProperties（w:rPr）解析出的真实样式
 * （字体颜色 / 字体 / 字号 / 高亮底色），toDOM 直接输出 inline style，
 * 这样编辑器和预览（同一份 schema）都会"如实"显示原始 DOCX 的样式，
 * 而不是本应用的通用主题样式（内联 style 的 CSS 优先级天然高于类选择器）。
 */
const docxStyleMark: MarkSpec = {
  attrs: {
    color: { default: null },
    fontFamily: { default: null },
    sizeHalfPt: { default: null }, // 原始单位：半磅（OOXML w:sz），与 docx.js TextRun.size 完全一致，导出时可直接回填
    highlight: { default: null },
  },
  parseDOM: [], // 仅由 OOXML 解析器生成，不从粘贴的 HTML 还原
  toDOM(mark) {
    const style: string[] = [];
    if (mark.attrs.color) style.push(`color:${mark.attrs.color}`);
    if (mark.attrs.fontFamily)
      style.push(`font-family:'${mark.attrs.fontFamily}'`);
    if (mark.attrs.sizeHalfPt)
      style.push(`font-size:${Number(mark.attrs.sizeHalfPt) / 2}pt`);
    if (mark.attrs.highlight)
      style.push(
        `background-color:${HIGHLIGHT_COLORS[mark.attrs.highlight] ?? '#FFFF00'}`
      );
    return ['span', { style: style.join(';'), class: 'docx-style-run' }, 0];
  },
};

/** comment：批注锚点。id 对应批注面板 / comments.xml 里的批注记录 */
const commentMark: MarkSpec = {
  attrs: { id: { default: 0 } },
  parseDOM: [],
  toDOM(mark) {
    return [
      'span',
      { class: 'docx-comment', 'data-comment-id': String(mark.attrs.id) },
      0,
    ];
  },
};

const marksWithExtras = basicSchema.spec.marks
  .append({
    underline: underlineMark,
    strike: strikeMark,
    docxStyle: docxStyleMark,
    comment: commentMark,
  })
  .append(pluginRegistry.marks());

export const docSchema = new Schema({
  nodes: nodesWithPlugins,
  marks: marksWithExtras,
});

export type DocxSchema = typeof docSchema;

/** §5.1 文档根结构的空文档，用于新建 / 兜底 */
export const EMPTY_DOC = {
  type: 'doc',
  content: [{ type: 'paragraph', content: [] }],
};

export function highlightColorCss(name: string): string {
  return HIGHLIGHT_COLORS[name] ?? '#FFFF00';
}
