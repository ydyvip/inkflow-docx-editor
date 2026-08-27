/**
 * VirtualPagedPreview —— 虚拟化分页文档预览
 * ------------------------------------------------------------
 * 将 PreviewPane 中间的单滚动条连续文档渲染，替换为若干个固定尺寸
 * 的"纸张"页面容器，虚拟 DOM 化惰性加载：仅挂载可见页 ±3 页，
 * 未挂载的页渲染轻量页码占位符以保持滚动条稳定；每个已挂载页在
 * 短时模拟延迟后创建自己的只读 EditorView（复用同一 schema 与
 * highlightPlugin，样式通过 .editor-page .ProseMirror 选择器生效）。
 *
 * 页面几何（纸张大小 + 页边距）来自 DOCX 的 sectPr（用 pageSetup 传入），
 * 不再是硬编码常量，因此纸张/页数/页码与真实 DOCX 页面区域对齐。
 * 页眉/页脚作为覆盖层渲染在上下页边距带内，每页重复，并支持
 * PAGE / NUMPAGES 域（按当前页/总页数替换）。
 */
import {
  createSignal,
  createEffect,
  onMount,
  onCleanup,
  createMemo,
  Show,
  For,
} from 'solid-js';
import type { Node as PMNode } from 'prosemirror-model';
import { EditorState } from 'prosemirror-state';
import { EditorView } from 'prosemirror-view';
import { docSchema } from '../schema';
import { pluginRegistry } from '../plugins/registry';
import { highlightPlugin, highlightPluginKey } from '../editor/highlightPlugin';
import {
  FIELD_PAGE,
  FIELD_NUMPAGES,
  FIELD_SECTIONPAGES,
  type DocxPageSetup,
  type DocxSection,
} from '../parser/ooxml';
import {
  computePerSectionPagination,
  slicePage,
  type SectionGeom,
  type Pagination,
} from './paginateDoc';

// ---- 滚动布局常量 ----
export const GAP = 28; // 页面之间的间距
export const PAD = 40; // 滚动容器上下留白

// twips → px（96dpi：1 英寸 = 1440 twips = 96px）
const TW_PER_PX = 1440 / 96;

/** 缺省页面几何（无 sectPr 时保留旧视觉：A4 纸 + 1inch 近似边距） */
const FALLBACK_GEOMETRY = {
  pageWidthPx: 760,
  pageHeightPx: 1056,
  marginTopPx: 72,
  marginBottomPx: 72,
  marginLeftPx: 76,
  marginRightPx: 76,
};

function setupToGeometry(ps: DocxPageSetup | null | undefined) {
  if (!ps || typeof ps.pageWidthTw !== 'number' || !Number.isFinite(ps.pageWidthTw))
    return FALLBACK_GEOMETRY;
  const px = (tw: number) => tw / TW_PER_PX;
  return {
    pageWidthPx: px(ps.pageWidthTw),
    pageHeightPx: px(ps.pageHeightTw),
    marginTopPx: px(ps.marginTopTw),
    marginBottomPx: px(ps.marginBottomTw),
    marginLeftPx: px(ps.marginLeftTw),
    marginRightPx: px(ps.marginRightTw),
  };
}

/** ±预载页数 */
const PRELOAD = 3;

export interface ScrollOpts {
  highlightBlockId?: string;
  highlightCommentId?: number;
}

export interface ActiveViewInfo {
  view: EditorView;
  /** 该页第一个块的全局文档起始偏移，用于把本地选区位置映射到全局 */
  globalStart: number;
}

export interface VirtualPagedPreviewApi {
  scrollToPage: (index: number, o?: ScrollOpts) => void;
  getPageCount: () => number;
  pageCount: () => number;
  blockToPage: () => Map<string, number>;
  commentToPage: () => Map<number, number>;
}

interface VirtualPagedPreviewProps {
  doc: PMNode;
  commentsVersion: number;
  pageSetup: DocxPageSetup | null;
  header: any | null;
  footer: any | null;
  sections: DocxSection[];
  onApiReady?: (api: VirtualPagedPreviewApi) => void;
  onActiveViewChange?: (info: ActiveViewInfo) => void;
}

function simulateAsyncDelay(index: number): number {
  return 120 + (index % 4) * 90 + Math.random() * 120;
}

/** 把整数转成罗马数字（upper=true 大写 I,II,III；false 小写 i,ii,iii） */
function toRoman(n: number, upper: boolean): string {
  const map: [number, string][] = [
    [1000, 'M'], [900, 'CM'], [500, 'D'], [400, 'CD'], [100, 'C'],
    [90, 'XC'], [50, 'L'], [40, 'XL'], [10, 'X'], [9, 'IX'],
    [5, 'V'], [4, 'IV'], [1, 'I'],
  ];
  let s = '';
  let x = n;
  for (const [v, c] of map) {
    while (x >= v) {
      s += c;
      x -= v;
    }
  }
  return upper ? s : s.toLowerCase();
}

/** 把整数转成字母序号（A,B,...,Z,AA,AB,...） */
function toLetters(n: number): string {
  let s = '';
  let x = n;
  while (x > 0) {
    const r = (x - 1) % 26;
    s = String.fromCharCode(65 + r) + s;
    x = Math.floor((x - 1) / 26);
  }
  return s;
}

/** 按 OOXML 的 w:pgNumType w:format 把页码渲染成显示文本 */
function formatPageNumber(n: number, format: string): string {
  switch (format) {
    case 'upperRoman':
      return toRoman(n, true);
    case 'lowerRoman':
      return toRoman(n, false);
    case 'upperLetter':
      return toLetters(n);
    case 'lowerLetter':
      return toLetters(n).toLowerCase();
    case 'none':
    case 'bullet':
      return '';
    default: // decimal 及其它未识别格式按十进制
      return String(n);
  }
}

/** 渲染某页页眉/页脚所需的页码信息（参考 DocumentServer 的 SectionPageNumInfo 思路） */
interface HfNumInfo {
  /** 显示页码（已按节起始/重启计算，未格式化） */
  pageNum: number;
  /** 本页所在节的页码格式（w:pgNumType format） */
  format: string;
  /** 文档总页数 */
  totalPages: number;
  /** 本页所在节的页数（SECTIONPAGES 域用） */
  sectionPages: number;
}

/** 把页眉/页脚 JSON 里的 PAGE / NUMPAGES / SECTIONPAGES 域占位替换成当前页码/总页数/节页数 */
function resolveHfFields(json: any, info: HfNumInfo): any {
  if (Array.isArray(json))
    return json.map((c: any) => resolveHfFields(c, info));
  if (json && typeof json === 'object') {
    if (json.type === 'text' && typeof json.text === 'string') {
      if (json.text === FIELD_PAGE)
        return { ...json, text: formatPageNumber(info.pageNum, info.format) };
      if (json.text === FIELD_NUMPAGES)
        return { ...json, text: String(info.totalPages) };
      if (json.text === FIELD_SECTIONPAGES)
        return { ...json, text: String(info.sectionPages) };
    }
    if (json.type && json.content) {
      return { ...json, content: resolveHfFields(json.content, info) };
    }
    return json;
  }
  return json;
}

/** 在一个容器元素内挂载一个只读 EditorView（页眉/页脚都用它渲染） */
function makeReadonlyView(
  host: HTMLElement,
  doc: PMNode,
  onFocus?: (view: EditorView) => void
): EditorView {
  const state = EditorState.create({
    schema: docSchema,
    doc,
    plugins: [highlightPlugin()],
  });
  let v: EditorView;
  v = new EditorView(host, {
    state,
    editable: () => false,
    nodeViews: pluginRegistry.nodeViews(docSchema),
    dispatchTransaction(tr) {
      if (v) v.updateState(v.state.apply(tr));
    },
  });
  if (onFocus) {
    v.dom.addEventListener('focus', () => onFocus(v));
    v.dom.addEventListener('mouseup', () => onFocus(v));
  }
  return v;
}

interface VirtualPageProps {
  index: number;
  doc: PMNode;
  indices: number[];
  globalStart: number;
  commentsVersion: number;
  header: any | null;
  footer: any | null;
  numInfo: HfNumInfo;
  pageWidthPx: number;
  marginLeftPx: number;
  marginRightPx: number;
  marginTopPx: number;
  marginBottomPx: number;
  /** 纸张总高（= 内容区高度 + 上下页边距）；超高内容页会更大 */
  sheetH: number;
  /** 内容区高度（= sheetH - marginTop - marginBottom） */
  contentAreaH: number;
  onViewFocus: (globalStart: number, view: EditorView) => void;
  onLoaded: (index: number, view: EditorView) => void;
}

/** 单页：固定尺寸纸张容器，内嵌只读 EditorView（正文）+ 页眉/页脚渲染区 */
function VirtualPage(props: VirtualPageProps) {
  let contentEl: HTMLDivElement | undefined;
  let headerEl: HTMLDivElement | undefined;
  let footerEl: HTMLDivElement | undefined;
  let bodyView: EditorView | undefined;
  let headerView: EditorView | undefined;
  let footerView: EditorView | undefined;
  let timer: ReturnType<typeof setTimeout> | undefined;
  const [loaded, setLoaded] = createSignal(false);

  onMount(() => {
    timer = setTimeout(() => {
      setLoaded(true);
    }, simulateAsyncDelay(props.index));
    onCleanup(() => {
      if (timer) clearTimeout(timer);
      bodyView?.destroy();
      headerView?.destroy();
      footerView?.destroy();
      bodyView = undefined;
      headerView = undefined;
      footerView = undefined;
    });
  });

  // 把页眉/页脚 JSON（含 PAGE/NUMPAGES 占位）解析成只读 EditorView
  const mountHf = (el: HTMLDivElement | undefined, json: any | null) => {
    if (!el || !json) return;
    const doc = docSchema.nodeFromJSON(resolveHfFields(json, props.numInfo));
    return makeReadonlyView(el, doc);
  };

  // 页面加载完成 / commentsVersion 变化时（重新）构建正文 + 页眉 + 页脚视图
  createEffect(() => {
    if (!loaded()) return;
    void props.commentsVersion;
    if (contentEl) {
      bodyView?.destroy();
      bodyView = makeReadonlyView(
        contentEl,
        slicePage(props.doc, props.indices),
        (v) => props.onViewFocus(props.globalStart, v)
      );
      props.onLoaded(props.index, bodyView);
    }
    headerView?.destroy();
    headerView = mountHf(headerEl, props.header);
    footerView?.destroy();
    footerView = mountHf(footerEl, props.footer);
  });

  const pageNum = props.index + 1;

  return (
    <div
      class="vp-page preview-page editor-page preview-mode"
      style={`width:${props.pageWidthPx}px;height:${props.sheetH}px;padding-top:${props.marginTopPx}px;padding-right:${props.marginRightPx}px;padding-bottom:${props.marginBottomPx}px;padding-left:${props.marginLeftPx}px;`}
    >
      <Show when={props.header != null}>
        <div
          class="vp-page-header vp-hf"
          ref={headerEl}
          style={{
            height: `${props.marginTopPx}px`,
            left: `${props.marginLeftPx}px`,
            right: `${props.marginRightPx}px`,
          }}
        />
      </Show>
      <div
        class="vp-page-content"
        ref={contentEl}
        style={{ height: `${props.contentAreaH}px` }}
      />
      <Show when={props.footer != null}>
        <div
          class="vp-page-footer vp-hf"
          ref={footerEl}
          style={{
            height: `${props.marginBottomPx}px`,
            left: `${props.marginLeftPx}px`,
            right: `${props.marginRightPx}px`,
          }}
        />
      </Show>
      <Show when={!loaded()}>
        <div class="vp-loading">
          <span class="vp-spinner" />
          <div class="vp-loading-text">正在加载第 {pageNum} 页…</div>
        </div>
      </Show>
      {/* 应用自带的页码指示：仅在文档无页脚时显示，避免与真实页脚重复/重叠 */}
      <Show when={props.footer == null}>
        <div class="vp-page-num">{pageNum}</div>
      </Show>
    </div>
  );
}

export function VirtualPagedPreview(props: VirtualPagedPreviewProps) {
  let scrollEl: HTMLDivElement | undefined;
  const [pagination, setPagination] = createSignal<Pagination | null>(null);
  const [scrollTop, setScrollTop] = createSignal(0);
  const [viewHeight, setViewHeight] = createSignal(0);

  const pageViewsRef: Record<number, { view: EditorView; globalStart: number }> =
    {};
  const pendingRef: Record<number, ScrollOpts> = {};

  /** 逐节页眉/页脚 + 纸张设置；无节信息时退回一个覆盖全文档的隐式单节（用 props.header/footer/pageSetup） */
  const sectionList = createMemo<DocxSection[]>(() => {
    const s = props.sections;
    if (s && s.length) return s;
    return [
      {
        startBlock: 0,
        endBlock: Number.MAX_SAFE_INTEGER,
        header: props.header ?? null,
        footer: props.footer ?? null,
        pageStart: null,
        pageFormat: 'decimal',
        pageSetup: props.pageSetup ?? ({} as DocxPageSetup),
      },
    ];
  });

  /** 取某一节（索引）的几何：纸型 + 页边距 → px，以及内容宽/高 */
  const sectionGeom = (si: number) => {
    const sects = sectionList();
    const sect = sects[si];
    const ps = sect?.pageSetup ?? props.pageSetup;
    const g = setupToGeometry(ps);
    return {
      geometry: g,
      contentWidthPx: Math.max(1, g.pageWidthPx - g.marginLeftPx - g.marginRightPx),
      contentHeightPx: Math.max(1, g.pageHeightPx - g.marginTopPx - g.marginBottomPx),
    };
  };

  createEffect(() => {
    const doc = props.doc;
    if (!doc) return;
    const sects = sectionList();
    const geoms: SectionGeom[] = sects.map((_: DocxSection, i: number) => {
      const g = sectionGeom(i);
      return {
        startBlock: sects[i].startBlock,
        endBlock: sects[i].endBlock,
        contentWidthPx: g.contentWidthPx,
        contentHeightPx: g.contentHeightPx,
      };
    });
    let cancelled = false;
    computePerSectionPagination(doc, geoms).then((p) => {
      if (!cancelled) setPagination(p);
    });
    onCleanup(() => {
      cancelled = true;
    });
  });

  onMount(() => {
    if (!scrollEl) return;
    const update = () => {
      setScrollTop(scrollEl!.scrollTop);
      setViewHeight(scrollEl!.clientHeight);
    };
    update();
    const ro = new ResizeObserver(update);
    ro.observe(scrollEl);
    scrollEl.addEventListener('scroll', update);
    onCleanup(() => {
      ro.disconnect();
      scrollEl?.removeEventListener('scroll', update);
    });
  });

  const pageCount = () => pagination()?.pages.length ?? 0;

  const pageIndices = createMemo(() =>
    Array.from({ length: pageCount() }, (_, i) => i)
  );

  const sectionsMemo = createMemo(() => sectionList());

  /** 返回某页所属节的索引（分页时已按块区间指派，见 paginatePerSection） */
  const sectionOfPage = (i: number): number =>
    pagination()?.pages[i]?.section ?? 0;

  /** 取某页所属节的页眉/页脚 */
  const sectionAt = (
    i: number
  ): { header: any | null; footer: any | null } => {
    const s = sectionsMemo();
    const sect = s[sectionOfPage(i)];
    if (!sect) return { header: props.header, footer: props.footer };
    return { header: sect.header, footer: sect.footer };
  };

  /**
   * 逐页页码计算（参考 DocumentServer 的 SectionPageNumType 思路）：
   * - 每页按所属节（分页已指派）计算；
   * - 节的 pageStart 非空 → 页码从该值重新起算；否则沿用上一节连续编号；
   * - sectionPages = 该节实际页数（供 SECTIONPAGES 域用）。
   */
  const numbering = createMemo(() => {
    const n = pageCount();
    const sects = sectionList();
    const nums: number[] = new Array(n).fill(1);
    const sectionPages: number[] = new Array(n).fill(1);
    if (!n) return { nums, sectionPages };
    const sectOfPage: number[] = [];
    for (let i = 0; i < n; i++) sectOfPage.push(sectionOfPage(i));
    const countPer: number[] = new Array(sects.length).fill(0);
    for (const si of sectOfPage) if (si >= 0) countPer[si]++;
    let counter = 1;
    let lastSectionKey = -1;
    for (let i = 0; i < n; i++) {
      const si = sectOfPage[i];
      // 仅在进入新节时重置页码（该节声明了重启点）；同一节内的后续页沿续递增
      if (si !== lastSectionKey) {
        if (si >= 0 && sects[si].pageStart != null) counter = sects[si].pageStart!;
        lastSectionKey = si;
      }
      nums[i] = counter;
      counter++;
      sectionPages[i] = si >= 0 ? countPer[si] : 1;
    }
    return { nums, sectionPages };
  });

  const numInfo = (i: number): HfNumInfo => {
    const { nums, sectionPages } = numbering();
    const sects = sectionList();
    const si = sectionOfPage(i);
    const fmt = si >= 0 ? sects[si]?.pageFormat : 'decimal';
    return {
      pageNum: nums[i] ?? i + 1,
      format: fmt || 'decimal',
      totalPages: pageCount(),
      sectionPages: sectionPages[i] ?? 1,
    };
  };

  const startOffsets = createMemo(() => {
    const pages = pagination()?.pages ?? [];
    const doc = pagination()?.doc ?? props.doc;
    const offsets: number[] = [];
    let acc = 0;
    for (const page of pages) {
      offsets.push(acc);
      for (const i of page.indices) {
        acc += doc.child(i).nodeSize;
      }
    }
    return offsets;
  });

  const pagesMeta = createMemo(() => {
    const p = pagination();
    if (!p)
      return {
        contentAreas: [] as number[],
        sheetHs: [] as number[],
        offsets: [] as number[],
        geos: [] as ReturnType<typeof setupToGeometry>[],
        total: 0,
      };
    const geos = p.pages.map((pg) => sectionGeom(pg.section).geometry);
    const contentAreas = p.pages.map((pg) =>
      Math.max(sectionGeom(pg.section).contentHeightPx, pg.contentH)
    );
    const sheetHs = p.pages.map(
      (_, i) => contentAreas[i] + geos[i].marginTopPx + geos[i].marginBottomPx
    );
    const offsets: number[] = [];
    let acc = PAD;
    for (let i = 0; i < sheetHs.length; i++) {
      offsets.push(acc);
      acc += sheetHs[i] + GAP;
    }
    return { contentAreas, sheetHs, offsets, geos, total: acc - GAP + PAD };
  });

  const isActive = (i: number): boolean => {
    const { sheetHs, offsets } = pagesMeta();
    const n = offsets.length;
    if (n === 0) return false;
    const top = scrollTop();
    const bot = top + viewHeight();
    let first = n - 1;
    for (let idx = 0; idx < n; idx++) {
      if (offsets[idx] + sheetHs[idx] > top) {
        first = idx;
        break;
      }
    }
    let last = 0;
    for (let idx = 0; idx < n; idx++) {
      if (offsets[idx] < bot) last = idx;
    }
    if (last < first) last = first;
    return i >= first - PRELOAD && i <= last + PRELOAD;
  };

  const applyHighlight = (v: EditorView, opts: ScrollOpts) => {
    if (!v) return;
    if (opts.highlightBlockId) {
      v.dispatch(
        v.state.tr.setMeta(highlightPluginKey, {
          ids: [opts.highlightBlockId],
          mode: 'replace',
        })
      );
      return;
    }
    if (opts.highlightCommentId !== undefined) {
      const id = opts.highlightCommentId;
      let from: number | null = null;
      let to: number | null = null;
      v.state.doc.descendants((node, pos) => {
        if (
          node.isText &&
          node.marks.some(
            (m) =>
              m.type.name === 'comment' && Number(m.attrs.id) === Number(id)
          )
        ) {
          if (from === null || pos < from) from = pos;
          const end = pos + node.nodeSize;
          if (to === null || end > to) to = end;
        }
        return true;
      });
      if (from !== null && to !== null) {
        v.dispatch(
          v.state.tr.setMeta(highlightPluginKey, {
            ranges: [{ from, to }],
            mode: 'replace',
          })
        );
      }
    }
  };

  const scrollToPage = (index: number, o?: ScrollOpts) => {
    const { offsets } = pagesMeta();
    const n = offsets.length;
    if (n === 0) return;
    const idx = Math.max(0, Math.min(n - 1, index));
    const target = offsets[idx];
    if (scrollEl) scrollEl.scrollTop = target;
    if (o) {
      const existing = pageViewsRef[idx];
      if (existing?.view) {
        applyHighlight(existing.view, o);
      } else {
        pendingRef[idx] = o;
      }
    }
  };

  const handleLoaded = (index: number, view: EditorView) => {
    pageViewsRef[index] = { view, globalStart: startOffsets()[index] ?? 0 };
    const pending = pendingRef[index];
    if (pending) {
      applyHighlight(view, pending);
      delete pendingRef[index];
    }
  };

  const handleViewFocus = (globalStart: number, view: EditorView) => {
    props.onActiveViewChange?.({ view, globalStart });
  };

  let apiSent = false;
  createEffect(() => {
    if (pagination() && !apiSent) {
      apiSent = true;
      props.onApiReady?.({
        scrollToPage,
        getPageCount: pageCount,
        pageCount,
        blockToPage: () => pagination()?.blockToPage ?? new Map(),
        commentToPage: () => pagination()?.commentToPage ?? new Map(),
      });
    }
  });

  const totalHeight = () => pagesMeta().total;

  return (
    <div class="flex-1 min-w-0 h-full min-h-0 flex flex-col bg-canvas px-6 pt-10 pb-12">
      <div class="vh-vp-scroll flex-1 min-h-0" ref={scrollEl}>
        <Show
          when={pagination()}
          fallback={
            <div class="flex items-center justify-center h-full">
              <span class="vp-spinner" />
            </div>
          }
        >
          <div style={{ position: 'relative', height: `${totalHeight()}px` }}>
            <For each={pageIndices()}>
              {(i) => {
                const meta = pagesMeta();
                const top = meta.offsets[i];
                const sheetH = meta.sheetHs[i];
                const contentAreaH = meta.contentAreas[i];
                const geo = meta.geos[i];
                const hf = sectionAt(i);
                return (
                  <div
                    class="vh-vp-spacer"
                    style={{ top: `${top}px`, height: `${sheetH}px` }}
                  >
                    <Show
                      when={isActive(i)}
                      fallback={<div class="vp-page-num">{i + 1}</div>}
                    >
                      <VirtualPage
                        index={i}
                        doc={pagination()!.doc}
                        indices={pagination()!.pages[i].indices}
                        globalStart={startOffsets()[i] ?? 0}
                        commentsVersion={props.commentsVersion}
                        header={hf.header}
                        footer={hf.footer}
                        numInfo={numInfo(i)}
                        pageWidthPx={geo.pageWidthPx}
                        marginLeftPx={geo.marginLeftPx}
                        marginRightPx={geo.marginRightPx}
                        marginTopPx={geo.marginTopPx}
                        marginBottomPx={geo.marginBottomPx}
                        sheetH={sheetH}
                        contentAreaH={contentAreaH}
                        onViewFocus={handleViewFocus}
                        onLoaded={handleLoaded}
                      />
                    </Show>
                  </div>
                );
              }}
            </For>
          </div>
        </Show>
      </div>
    </div>
  );
}
