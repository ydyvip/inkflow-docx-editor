/**
 * Pagination —— 把 ProseMirror 顶层块节点贪婪打包成固定高度页面。
 * ------------------------------------------------------------
 * 依赖 Layout：通过隐藏的只读 EditorView 渲染文档，读每个顶层块的真实矩形
 * （getBoundingClientRect，含外边距），按 pageHeightPx 贪婪分页。
 *
 * 多页友好处理：
 *  - 超高 code_block：按行预拆分成若干个"≤ 一页"的 code_block，从而可跨多页
 *    连续渲染（不会因一整段超长代码被裁切或占满一张超高页）。
 *    预拆分后的文档记为「展示文档 Pagination.doc」，分页与按页切片都以它为
 *    准；其余块（段落/标题/表格/图片等）保持原子。
 *  - 表格：维持原子 + 可变高度页（超高表格独占一页并自动加高），不裁切。
 *  - 图片/visio：原子 + CSS max-height 钳制在一页内展示（见 index.css）。
 */
import type { Node as PMNode } from 'prosemirror-model';
import { EditorState } from 'prosemirror-state';
import { EditorView } from 'prosemirror-view';
import { docSchema } from '../schema';
import { pluginRegistry } from '../plugins/registry';
import { highlightPlugin } from '../editor/highlightPlugin';

export interface Pagination {
  /** 分页所依据的"展示文档"：内含被预拆分的超高代码块。按页切片必须用它。 */
  doc: PMNode;
  /**
   * 每页包含的顶层块索引（针对展示文档 doc）。
   * contentH：该页实际占用的内容高度（px，含块间真实边距去重的累计高度）。
   * 常规页 ≤ CONTENT_HEIGHT；若该页仅含一个超高块（如超长表格），contentH 会超过
   * 页高——渲染端据此把该页（纸张）加高，保证超高内容不被裁切。
   * section：该页所属章节的索引（用于按节取页眉/页脚、页码与纸张几何）。
   */
  pages: { indices: number[]; contentH: number; section: number }[];
  /** blockId（paragraph/heading）→ 所在页 */
  blockToPage: Map<string, number>;
  /** cellId（表格单元格）→ 所在页 */
  cellToPage: Map<string, number>;
  /** comment id → 所在页 */
  commentToPage: Map<number, number>;
}

/** 单节（子文档）分页结果：页不含 section，由逐节拼接时统一指派 */
interface SubPagination {
  doc: PMNode;
  pages: { indices: number[]; contentH: number }[];
  blockToPage: Map<string, number>;
  cellToPage: Map<string, number>;
  commentToPage: Map<number, number>;
}

/** 打包安全余量（px）：补偿测量与最终渲染之间的微小差异，避免页面内容被裁切 */
const PACK_SAFETY = 24;
/** code_block 拆块时，每块额外预留的行/间距余量（px） */
const CODE_SPLIT_MARGIN = 8;

/** 从"展示文档"中切出一页：内容仅为给定顶层子节点 */
export function slicePage(doc: PMNode, indices: number[]): PMNode {
  if (indices.length === 0) return doc.type.create({});
  const children = indices.map((i) => doc.child(i));
  return doc.type.create({}, children);
}

/** 挂载一个隐藏的只读测量视图，返回其内容 DOM */
function mountMeasure(doc: PMNode, widthPx: number) {
  const container = document.createElement('div');
  container.style.cssText =
    'position:fixed;left:-99999px;top:0;visibility:hidden;';
  container.style.width = `${widthPx}px`;
  container.className = 'editor-page preview-mode';
  const host = document.createElement('div');
  container.appendChild(host);
  document.body.appendChild(container);

  const state = EditorState.create({
    schema: docSchema,
    doc,
    plugins: [highlightPlugin()],
  });
  const view = new EditorView(host, {
    state,
    editable: () => false,
    nodeViews: pluginRegistry.nodeViews(docSchema),
  });
  const contentDom =
    (view.dom.querySelector('.ProseMirror') as HTMLElement | null) ??
    view.dom;
  return { container, view, contentDom };
}

/** 读取每个顶层块的相对 top/bottom（含真实外边距/留白） */
function measureTops(childDoms: Element[]) {
  const rects = childDoms.map((d) =>
    d instanceof HTMLElement ? d.getBoundingClientRect() : null
  );
  const base = rects[0]?.top ?? 0;
  const tops = rects.map((r) =>
    r ? r.top - base : (rects.find((x) => x)?.top ?? 0) - base
  );
  const bottoms = rects.map((r) =>
    r ? r.bottom - base : tops[tops.length - 1] + 100
  );
  return { tops, bottoms };
}

/**
 * 把超高 code_block 按行预拆成多个 ≤ 一页的 code_block，
 * 以支持代码块跨多页连续渲染。其它块原样保留。
 */
function splitOversizedCodeBlocks(
  doc: PMNode,
  tops: number[],
  bottoms: number[],
  limit: number
): PMNode {
  let changed = false;
  const children: PMNode[] = [];
  const total = doc.content.childCount;
  for (let i = 0; i < total; i++) {
    const child = doc.content.child(i);
    const h = (bottoms[i] ?? 0) - (tops[i] ?? 0);
    if (child.type.name === 'code_block' && h > limit) {
      const lines = child.textContent.split('\n');
      const lineCount = lines.length || 1;
      // 代码块通常均匀行高，用总高/行数估算行高，再按页高切块
      const lineHeight = h / lineCount;
      const per = Math.max(1, Math.floor((limit - CODE_SPLIT_MARGIN) / lineHeight));
      if (per < lineCount) {
        changed = true;
        for (let s = 0; s < lineCount; s += per) {
          const chunk = lines.slice(s, s + per).join('\n');
          children.push(
            child.type.create(child.attrs, doc.type.schema.text(chunk))
          );
        }
        continue;
      }
    }
    children.push(child);
  }
  return changed ? doc.type.create({}, children) : doc;
}

/** 按测量出的 tops/bottoms 贪婪打包，并收集 blockId/cellId/commentId → 页映射 */
function packPages(
  doc: PMNode,
  tops: number[],
  bottoms: number[],
  limit: number
) {
  const pages: { indices: number[]; contentH: number }[] = [];
  const blockToPage = new Map<string, number>();
  const cellToPage = new Map<string, number>();
  const commentToPage = new Map<number, number>();

  let currentIndices: number[] = [];
  let pageTop = 0; // 当前页第一个块的相对基准
  let pageContentH = 0; // 当前页已占内容高度

  const flush = () => {
    if (currentIndices.length)
      pages.push({ indices: currentIndices, contentH: pageContentH });
    currentIndices = [];
    pageTop = 0;
    pageContentH = 0;
  };

  const total = doc.content.childCount;
  for (let i = 0; i < total; i++) {
    const block = doc.content.child(i);
    const topI = tops[i] ?? pageTop;
    const bottomI = bottoms[i] ?? topI + 100;
    const contentHeight = bottomI - pageTop;

    // 放不下且当前页已有 ≥1 块 → 换页。单个超高块（contentHeight > limit 但
    // 当前页为空）仍留本页并记录其真实高度，渲染端会把该页加高以免裁切。
    if (currentIndices.length > 0 && contentHeight > limit) {
      flush();
      pageTop = topI;
    }
    const pageIndex = pages.length;
    currentIndices.push(i);
    pageContentH = contentHeight;

    if (block.attrs.blockId) {
      blockToPage.set(block.attrs.blockId, pageIndex);
    }
    if (block.isBlock) {
      if (block.attrs.cellId) cellToPage.set(block.attrs.cellId, pageIndex);
      block.descendants((desc) => {
        const cellId = desc.attrs?.cellId;
        if (cellId) cellToPage.set(cellId, pageIndex);
        if (desc.isText) {
          for (const m of desc.marks) {
            if (m.type.name === 'comment') {
              commentToPage.set(Number(m.attrs.id), pageIndex);
            }
          }
        }
        return true;
      });
    }
  }
  flush();

  return { pages, blockToPage, cellToPage, commentToPage };
}

/** 同步测量 + 贪婪打包。调用方需确保容器布局已就绪。 */
export function paginateDoc(
  doc: PMNode,
  opts: { contentWidthPx: number; pageHeightPx: number }
): SubPagination {
  const limit = opts.pageHeightPx - PACK_SAFETY;

  // ---- 阶段 1：测量原文档，识别需要拆分的超高代码块 ----
  const m1 = mountMeasure(doc, opts.contentWidthPx);
  const childDoms1 = Array.from(m1.contentDom.children);
  const { tops: t1, bottoms: b1 } = measureTops(childDoms1);

  const displayDoc = splitOversizedCodeBlocks(doc, t1, b1, limit);

  // 没有需要拆分的超高代码块 → 直接复用阶段 1 的测量结果，避免二次隐藏渲染
  if (displayDoc === doc) {
    const packed = packPages(doc, t1, b1, limit);
    m1.view.destroy();
    m1.container.remove();
    return { doc, ...packed };
  }

  m1.view.destroy();
  m1.container.remove();

  // ---- 阶段 2：对展示文档（含拆分后的代码块）测量 + 贪婪分页 ----
  const m2 = mountMeasure(displayDoc, opts.contentWidthPx);
  const childDoms = Array.from(m2.contentDom.children);
  const { tops, bottoms } = measureTops(childDoms);

  const packed = packPages(displayDoc, tops, bottoms, limit);
  m2.view.destroy();
  m2.container.remove();

  return { doc: displayDoc, ...packed };
}

// ---------------------------------------------------------------------------
// 逐节分页：每节用自己的纸型/页边距（宽度、可用高度）独立测量与打包，
// 再把各节结果按顺序拼接成统一的展示文档 + 全局页序/映射。
// ---------------------------------------------------------------------------

export interface SectionGeom {
  /** 节的正文块区间（针对传入分页的原文档 childIndex） */
  startBlock: number;
  endBlock: number;
  contentWidthPx: number;
  contentHeightPx: number;
}

/** 把原文档按节切成若干子文档，逐节调 paginateDoc，再拼接成统一结果 */
export function paginatePerSection(
  doc: PMNode,
  sectionList: SectionGeom[]
): Pagination {
  const total = doc.content.childCount;

  // 规范化区间：保证覆盖 [0,total) 且不越界
  const norm: SectionGeom[] = [];
  let cursor = 0;
  for (const s of sectionList) {
    const start = Math.max(cursor, Math.min(s.startBlock, total));
    const end = Math.min(s.endBlock, total);
    if (end > start) norm.push({ ...s, startBlock: start, endBlock: end });
    cursor = Math.max(cursor, end);
  }
  // 若节未覆盖到文档末尾（异常数据），兜底补一节到底
  if (cursor < total && norm.length) {
    const last = norm[norm.length - 1];
    norm[norm.length - 1] = { ...last, endBlock: total };
  } else if (!norm.length && total > 0) {
    norm.push({
      startBlock: 0,
      endBlock: total,
      contentWidthPx: sectionList[0]?.contentWidthPx ?? 600,
      contentHeightPx: sectionList[0]?.contentHeightPx ?? 800,
    });
  }

  const globalChildren: PMNode[] = [];
  const pages: { indices: number[]; contentH: number; section: number }[] = [];
  const blockToPage = new Map<string, number>();
  const cellToPage = new Map<string, number>();
  const commentToPage = new Map<number, number>();

  let childOffset = 0;
  let pageOffset = 0;
  norm.forEach((s, si) => {
    const children: PMNode[] = [];
    for (let i = s.startBlock; i < s.endBlock; i++) children.push(doc.child(i));
    const subDoc = doc.type.create({}, children);
    // 逐节测量 + 打包（内部会处理该节内超高代码块拆分）
    const sub = paginateDoc(subDoc, {
      contentWidthPx: s.contentWidthPx,
      pageHeightPx: s.contentHeightPx,
    });
    const subPageOffsets: number[] = [];
    sub.pages.forEach(() => {
      subPageOffsets.push(pageOffset);
      pageOffset++;
    });
    for (let k = 0; k < sub.doc.content.childCount; k++)
      globalChildren.push(sub.doc.content.child(k));
    for (const pg of sub.pages) {
      pages.push({
        indices: pg.indices.map((ix) => ix + childOffset),
        contentH: pg.contentH,
        section: si,
      });
    }
    for (const [k, v] of sub.blockToPage) blockToPage.set(k, subPageOffsets[v]);
    for (const [k, v] of sub.cellToPage) cellToPage.set(k, subPageOffsets[v]);
    for (const [k, v] of sub.commentToPage)
      commentToPage.set(k, subPageOffsets[v]);
    childOffset = globalChildren.length;
  });

  const displayDoc = doc.type.create({}, globalChildren);
  return { doc: displayDoc, pages, blockToPage, cellToPage, commentToPage };
}

/** 异步包装逐节分页 */
export async function computePerSectionPagination(
  doc: PMNode,
  sectionList: SectionGeom[]
): Promise<Pagination> {
  if (typeof document !== 'undefined' && document.fonts?.ready) {
    await document.fonts.ready;
  }
  await new Promise<void>((resolve) => requestAnimationFrame(() => resolve()));
  return paginatePerSection(doc, sectionList);
}
