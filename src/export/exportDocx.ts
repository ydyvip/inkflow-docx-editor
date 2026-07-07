/**
 * Export 模块（关键）
 * ------------------------------------------------------------
 * 输入：ProseMirror JSON Document（唯一真相源）+ 批注列表
 * 输出：DOCX 二进制（Blob）
 *
 * 映射表（§6.5，含本次扩展新增的样式/批注部分）：
 *   paragraph/heading → Paragraph（align 属性 → alignment）
 *   text              → TextRun（strong/em/underline/strike/docxStyle → 对应格式）
 *   bullet_list       → Paragraph(bullet)
 *   ordered_list      → Paragraph(numbering)
 *   table             → Table（cellId/background 已在解析时写入，这里回填底色）
 *   image(inline)     → ImageRun
 *   comment mark      → CommentRangeStart/End + CommentReference + Document.comments
 *   自定义插件节点     → 由插件的 exportNode() 提供映射（§11 扩展点）
 *
 * 任何转换都不直接操作 DOCX 二进制细节，而是通过 docx.js 的对象模型
 * 组装，保持"结构优先，样式其次"的原则（§2.3）。
 */
import {
  Document,
  Packer,
  Paragraph,
  TextRun,
  ExternalHyperlink,
  ImageRun,
  Table,
  TableRow,
  TableCell,
  HeadingLevel,
  LevelFormat,
  AlignmentType,
  BorderStyle,
  WidthType,
  ShadingType,
  LineRuleType,
  CommentRangeStart,
  CommentRangeEnd,
  CommentReference,
  type ParagraphChild,
  type ICommentOptions,
} from 'docx';
import { saveAs } from 'file-saver';
import { pluginRegistry } from '../plugins/registry';
import {
  docxImageType,
  getImageDimensions,
  parseDataUrl,
  remoteImageToDataUrl,
} from './imageUtils';
import type { DocxComment } from '../parser/ooxml';

const HEADING_MAP: Record<
  number,
  (typeof HeadingLevel)[keyof typeof HeadingLevel]
> = {
  1: HeadingLevel.HEADING_1,
  2: HeadingLevel.HEADING_2,
  3: HeadingLevel.HEADING_3,
  4: HeadingLevel.HEADING_4,
  5: HeadingLevel.HEADING_5,
  6: HeadingLevel.HEADING_6,
};

// docx.js 的 HeadingLevel 常量只到 6（Word 默认模板虽然存在 Heading7-9 样式，
// 但直接引用未在生成的 styles.xml 里定义的样式 id 有被 Word 忽略的风险）。
// 因此 7-9 级标题改为"结构不依赖具名样式"的直接加粗+字号格式作为兜底外观，
// 语义层面的标题层级仍然完整保留在我们自己的 JSON 模型（level 属性）里。
const HEADING_FALLBACK_RUN: Record<
  number,
  { bold: boolean; italics: boolean; sizeHalfPt: number }
> = {
  7: { bold: true, italics: false, sizeHalfPt: 26 },
  8: { bold: true, italics: true, sizeHalfPt: 24 },
  9: { bold: true, italics: true, sizeHalfPt: 22 },
};

function mapIndent(indentLevel: number | null | undefined) {
  if (!indentLevel) return undefined;
  return { left: Number(indentLevel) * 720 };
}

function mapSpacing(lineSpacing: number | null | undefined) {
  if (!lineSpacing) return undefined;
  return {
    line: Math.round(Number(lineSpacing) * 240),
    lineRule: LineRuleType.AUTO,
  };
}

const ALIGN_MAP: Record<
  string,
  (typeof AlignmentType)[keyof typeof AlignmentType]
> = {
  left: AlignmentType.LEFT,
  center: AlignmentType.CENTER,
  right: AlignmentType.RIGHT,
  justify: AlignmentType.JUSTIFIED,
};

function mapAlign(align: string | null | undefined) {
  return align ? ALIGN_MAP[align] : undefined;
}

const ORDERED_LIST_REF = 'docflow-ordered-list';

const NUMBERING_CONFIG = {
  config: [
    {
      reference: ORDERED_LIST_REF,
      levels: [
        {
          level: 0,
          format: LevelFormat.DECIMAL,
          text: '%1.',
          alignment: AlignmentType.START,
          style: { paragraph: { indent: { left: 720, hanging: 260 } } },
        },
        {
          level: 1,
          format: LevelFormat.LOWER_LETTER,
          text: '%2)',
          alignment: AlignmentType.START,
          style: { paragraph: { indent: { left: 1440, hanging: 260 } } },
        },
        {
          level: 2,
          format: LevelFormat.LOWER_ROMAN,
          text: '%3.',
          alignment: AlignmentType.START,
          style: { paragraph: { indent: { left: 2160, hanging: 260 } } },
        },
      ],
    },
  ],
};

function extractPlainText(node: any): string {
  if (node.type === 'text') return node.text ?? '';
  return (node.content ?? []).map(extractPlainText).join('');
}

interface RunWithComments {
  child: ParagraphChild;
  commentIds: number[];
}

/** 行内内容 → 带批注 ID 标记的 docx.js ParagraphChild 列表（内部中间态）
 *  fallbackRun：当某个文字节点自身没有显式加粗/字号时使用的默认外观
 *  （用于 7-9 级标题——它们不依赖 Word 具名样式，需要一个直接格式兜底）*/
async function buildInlineItems(
  content: any[],
  fallbackRun?: { bold: boolean; italics: boolean; sizeHalfPt: number }
): Promise<RunWithComments[]> {
  const items: RunWithComments[] = [];

  for (const node of content ?? []) {
    if (node.type === 'text') {
      const marks: any[] = node.marks ?? [];
      const bold =
        marks.some((m) => m.type === 'strong') || !!fallbackRun?.bold;
      const italics =
        marks.some((m) => m.type === 'em') || !!fallbackRun?.italics;
      const isCode = marks.some((m) => m.type === 'code');
      const underline = marks.some((m) => m.type === 'underline');
      const strike = marks.some((m) => m.type === 'strike');
      const link = marks.find((m) => m.type === 'link');
      const styleMark = marks.find((m) => m.type === 'docxStyle');
      const commentIds = marks
        .filter((m) => m.type === 'comment')
        .map((m) => Number(m.attrs?.id))
        .filter((id) => !Number.isNaN(id));

      const run = new TextRun({
        text: node.text ?? '',
        bold,
        italics,
        strike,
        underline: underline ? {} : undefined,
        font: isCode
          ? 'Courier New'
          : (styleMark?.attrs?.fontFamily ?? undefined),
        size:
          styleMark?.attrs?.sizeHalfPt ?? fallbackRun?.sizeHalfPt ?? undefined,
        color: styleMark?.attrs?.color
          ? String(styleMark.attrs.color).replace('#', '')
          : undefined,
        shading: isCode
          ? { type: ShadingType.CLEAR, fill: 'F3F3F3', color: 'auto' }
          : undefined,
        highlight:
          !isCode && styleMark?.attrs?.highlight
            ? (styleMark.attrs.highlight as any)
            : undefined,
      });

      items.push({
        child: link
          ? new ExternalHyperlink({
              link: link.attrs?.href ?? '#',
              children: [run],
            })
          : run,
        commentIds,
      });
      continue;
    }

    if (node.type === 'image') {
      const src: string = node.attrs?.src ?? '';
      const dataUrl = src.startsWith('data:')
        ? src
        : ((await remoteImageToDataUrl(src)) ?? '');
      const parsed = parseDataUrl(dataUrl);
      if (parsed) {
        const dims = await getImageDimensions(dataUrl);
        items.push({
          child: new ImageRun({
            type: docxImageType(parsed.mime),
            data: parsed.base64,
            transformation: dims,
          } as any),
          commentIds: [],
        });
      }
      continue;
    }

    if (node.type === 'hard_break') {
      items.push({
        child: new TextRun({ text: '', break: 1 }),
        commentIds: [],
      });
    }
  }

  return items;
}

/**
 * 把带批注 ID 的中间态整理成最终 ParagraphChild[]，在批注范围的
 * 起止边界插入 CommentRangeStart/CommentRangeEnd + CommentReference。
 * 采用"相邻同 ID 合并成一个范围"的策略，避免逐字符生成大量冗余标记。
 */
function injectCommentRanges(items: RunWithComments[]): ParagraphChild[] {
  const out: ParagraphChild[] = [];
  let active: number[] = [];

  for (let i = 0; i < items.length; i++) {
    const ids = items[i].commentIds;
    for (const id of ids) {
      if (!active.includes(id)) out.push(new CommentRangeStart(id));
    }
    active = [...new Set([...active, ...ids])];

    out.push(items[i].child);

    const nextIds = i + 1 < items.length ? items[i + 1].commentIds : [];
    for (const id of [...active]) {
      if (!nextIds.includes(id)) {
        out.push(new CommentRangeEnd(id), new CommentReference(id));
        active = active.filter((a) => a !== id);
      }
    }
  }
  return out;
}

async function finalizeChildren(
  content: any[],
  fallbackRun?: { bold: boolean; italics: boolean; sizeHalfPt: number }
): Promise<ParagraphChild[]> {
  return injectCommentRanges(await buildInlineItems(content, fallbackRun));
}

async function listToDocx(
  listNode: any,
  depth: number,
  ordered: boolean,
  overrides: Record<string, any>
): Promise<Paragraph[]> {
  const out: Paragraph[] = [];
  for (const item of listNode.content ?? []) {
    for (const child of item.content ?? []) {
      if (child.type === 'paragraph') {
        const children = await finalizeChildren(child.content ?? []);
        out.push(
          new Paragraph({
            ...overrides,
            alignment: mapAlign(child.attrs?.align),
            indent: mapIndent(child.attrs?.indent),
            spacing: mapSpacing(child.attrs?.lineSpacing),
            ...(ordered
              ? {
                  numbering: {
                    reference: ORDERED_LIST_REF,
                    level: Math.min(depth, 2),
                  },
                }
              : { bullet: { level: Math.min(depth, 8) } }),
            children,
          })
        );
      } else if (
        child.type === 'bullet_list' ||
        child.type === 'ordered_list'
      ) {
        out.push(
          ...(await listToDocx(
            child,
            depth + 1,
            child.type === 'ordered_list',
            overrides
          ))
        );
      } else {
        out.push(...((await blockNodeToDocx(child, overrides)) as Paragraph[]));
      }
    }
  }
  return out;
}

async function tableToDocx(node: any): Promise<Table> {
  const rows: TableRow[] = [];
  for (const rowNode of node.content ?? []) {
    const cells: TableCell[] = [];
    for (const cellNode of rowNode.content ?? []) {
      const isHeader = cellNode.type === 'table_header';
      const cellBlocks: (Paragraph | Table)[] = [];
      for (const child of cellNode.content ?? []) {
        cellBlocks.push(...(await blockNodeToDocx(child)));
      }
      // 单元格底色：优先使用从 DOCX 解析出的真实底色（background），
      // 没有的话表头默认浅灰底，回到导出前的原始外观
      const background = cellNode.attrs?.background
        ? String(cellNode.attrs.background).replace('#', '')
        : isHeader
          ? 'EFEFEF'
          : undefined;
      cells.push(
        new TableCell({
          children: cellBlocks.length ? cellBlocks : [new Paragraph({})],
          shading: background
            ? { type: ShadingType.CLEAR, fill: background, color: 'auto' }
            : undefined,
          columnSpan: cellNode.attrs?.colspan ?? undefined,
          rowSpan:
            cellNode.attrs?.rowspan && cellNode.attrs.rowspan > 1
              ? cellNode.attrs.rowspan
              : undefined,
        })
      );
    }
    rows.push(new TableRow({ children: cells }));
  }
  return new Table({ rows, width: { size: 100, type: WidthType.PERCENTAGE } });
}

/** 提示块底色（对应 calloutPlugin 的 tone 属性）*/
function calloutFill(tone: string): string {
  switch (tone) {
    case 'warning':
      return 'FCEFC7';
    case 'danger':
      return 'FBE0E0';
    default:
      return 'E4F0EF'; // info
  }
}

async function blockNodeToDocx(
  node: any,
  overrides: Record<string, any> = {}
): Promise<(Paragraph | Table)[]> {
  switch (node.type) {
    case 'paragraph':
      return [
        new Paragraph({
          ...overrides,
          alignment: mapAlign(node.attrs?.align),
          indent: mapIndent(node.attrs?.indent),
          spacing: mapSpacing(node.attrs?.lineSpacing),
          children: await finalizeChildren(node.content ?? []),
        }),
      ];

    case 'heading': {
      const level = node.attrs?.level ?? 1;
      const fallback = HEADING_FALLBACK_RUN[level];
      return [
        new Paragraph({
          ...overrides,
          heading: fallback
            ? undefined
            : (HEADING_MAP[level] ?? HeadingLevel.HEADING_1),
          alignment: mapAlign(node.attrs?.align),
          indent: mapIndent(node.attrs?.indent),
          spacing: mapSpacing(node.attrs?.lineSpacing),
          children: await finalizeChildren(node.content ?? [], fallback),
        }),
      ];
    }

    case 'blockquote': {
      const out: Paragraph[] = [];
      for (const child of node.content ?? []) {
        out.push(
          ...((await blockNodeToDocx(child, {
            ...overrides,
            indent: { left: 720 },
            border: {
              left: {
                color: 'C9C2B4',
                space: 8,
                style: BorderStyle.SINGLE,
                size: 12,
              },
            },
          })) as Paragraph[])
        );
      }
      return out;
    }

    case 'horizontal_rule':
      return [
        new Paragraph({
          ...overrides,
          border: {
            bottom: {
              color: '999999',
              space: 1,
              style: BorderStyle.SINGLE,
              size: 6,
            },
          },
        }),
      ];

    case 'code_block':
      return [
        new Paragraph({
          ...overrides,
          shading: { type: ShadingType.CLEAR, fill: 'F3F3F3', color: 'auto' },
          children: [
            new TextRun({ text: extractPlainText(node), font: 'Courier New' }),
          ],
        }),
      ];

    case 'bullet_list':
      return listToDocx(node, 0, false, overrides);

    case 'ordered_list':
      return listToDocx(node, 0, true, overrides);

    case 'table':
      return [await tableToDocx(node)];

    default: {
      // §11 扩展点：交给插件的 exportNode() 处理自定义节点（如 callout）
      const plugin = pluginRegistry
        .all()
        .find((p) => p.name === node.type && p.exportNode);
      if (plugin?.exportNode) {
        const mapped = plugin.exportNode(node, () => []);
        if (mapped?.paragraphs) {
          const out: Paragraph[] = [];
          for (const child of mapped.paragraphs) {
            out.push(
              ...((await blockNodeToDocx(child, {
                ...overrides,
                shading: {
                  type: ShadingType.CLEAR,
                  fill: calloutFill(mapped.__calloutTone),
                  color: 'auto',
                },
                border: {
                  left: {
                    color: '3E7C7C',
                    space: 8,
                    style: BorderStyle.SINGLE,
                    size: 16,
                  },
                },
              })) as Paragraph[])
            );
          }
          return out;
        }
      }
      // 未知节点兜底：至少保留纯文本，保证"结构优先"不丢内容
      const text = extractPlainText(node);
      return text
        ? [new Paragraph({ ...overrides, children: [new TextRun(text)] })]
        : [];
    }
  }
}

function buildCommentOptions(comments: DocxComment[]): ICommentOptions[] {
  return comments.map((c) => ({
    id: c.id,
    author: c.author,
    date: c.date ? new Date(c.date) : new Date(),
    children: [new Paragraph(c.text)],
  }));
}

/** ProseMirror JSON → DOCX Blob */
export async function jsonToDocxBlob(
  docJson: any,
  comments: DocxComment[] = []
): Promise<Blob> {
  const blocks: (Paragraph | Table)[] = [];
  for (const node of docJson.content ?? []) {
    blocks.push(...(await blockNodeToDocx(node)));
  }
  const doc = new Document({
    numbering: NUMBERING_CONFIG,
    comments: comments.length
      ? { children: buildCommentOptions(comments) }
      : undefined,
    sections: [{ children: blocks.length ? blocks : [new Paragraph({})] }],
  });
  return Packer.toBlob(doc);
}

/** 导出并触发浏览器下载 */
export async function exportAndDownloadDocx(
  docJson: any,
  filename = 'document.docx',
  comments: DocxComment[] = []
): Promise<void> {
  const blob = await jsonToDocxBlob(docJson, comments);
  saveAs(blob, filename.endsWith('.docx') ? filename : `${filename}.docx`);
}
