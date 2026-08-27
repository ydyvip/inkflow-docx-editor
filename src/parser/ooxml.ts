/**
 * OOXML 解析器（取代 mammoth 的 HTML 中间态）
 * ------------------------------------------------------------
 * mammoth 的设计哲学是"语义优先，丢弃视觉样式"，这与新的需求冲突：
 *   - 目录树需要标题的真实层级（mammoth 有，可用，但样式丢失）
 *   - 预览需要还原 DOCX 里的真实颜色/字体/字号 —— mammoth 默认不转换
 *   - 高亮段落/单元格需要稳定可寻址的 ID —— HTML 中间态没有这个概念
 *   - 批注（comments.xml + commentRangeStart/End）—— mammoth 完全不解析
 *
 * 因此这里直接读 DOCX 的原始 XML（word/document.xml、styles.xml、
 * numbering.xml、comments.xml、关系文件），产出 ProseMirror JSON。
 * 仍然遵守"结构优先，样式次之，还原度最后"的原则（§2.3）：
 * 复杂的修订标记、页眉页脚、分栏、行间距细节不处理。
 */
import JSZip from 'jszip';

const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const A_NS = 'http://schemas.openxmlformats.org/drawingml/2006/main';
const WP_NS =
  'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing';

/** EMU（OOXML 绘图单位）→ CSS 像素（96dpi）：英寸内有 914400 EMU，96 像素 */
const EMU_TO_PX = 96 / 914400;

/** twips（1/1440 英寸）→ CSS 像素（96dpi）：1 英寸 = 1440 twips = 96px */
const TW_PER_PX = 1440 / 96;

export interface DocxComment {
  id: number;
  author: string;
  date: string | null;
  text: string;
}

/** 页面几何参数（twips，1/1440 英寸）：纸张大小 + 页边距 + 页眉/页脚距离 */
export interface DocxPageSetup {
  pageWidthTw: number;
  pageHeightTw: number;
  marginTopTw: number;
  marginBottomTw: number;
  marginLeftTw: number;
  marginRightTw: number;
  headerDistTw: number;
  footerDistTw: number;
}

/** 页眉/页脚里 PAGE / NUMPAGES / SECTIONPAGES 域的占位文本标记（见 parseHfInline） */
export const FIELD_PAGE = '\uE000PAGE\uE001';
export const FIELD_NUMPAGES = '\uE000NUMPAGES\uE001';
export const FIELD_SECTIONPAGES = '\uE000SECTIONPAGES\uE001';

/** 一节：覆盖的正文块区间 + 该节默认页眉/页脚（已按"链接到前一节"继承解析）+ 页码/纸张设置 */
export interface DocxSection {
  /** 该节覆盖的正文 flat 块区间 [startBlock, endBlock)；末节 endBlock == 正文块总数 */
  startBlock: number;
  endBlock: number;
  header: any | null;
  footer: any | null;
  /** 该节页码起始值（w:pgNumType @w:start）；未重启页码时为 null（沿用上一节连续编号） */
  pageStart: number | null;
  /** 页码格式（w:pgNumType @w:format，缺省十进制 decimal） */
  pageFormat: string;
  /** 该节纸张规格 + 页边距（w:pgSz + w:pgMar），用于按节还原真实页面几何 */
  pageSetup: DocxPageSetup;
}

export interface ParsedDocx {
  json: any;
  comments: DocxComment[];
  warnings: string[];
  /** 实际纸张大小 + 页边距（twips），供分页预览还原真实页面区域 */
  pageSetup: DocxPageSetup;
  /** 默认页眉的 ProseMirror JSON 文档；无页眉时为 null（= 正文级/末节默认页眉） */
  header: any | null;
  /** 默认页脚同 header */
  footer: any | null;
  /** 逐节页眉/页脚（支持 21 节这类每节独立页眉页脚的文档），供按页渲染 */
  sections: DocxSection[];
}

// ---------------------------------------------------------------------------
// 小工具：按 localName 取直接子元素（忽略命名空间前缀差异）
// ---------------------------------------------------------------------------

function children(el: Element | null, localName: string): Element[] {
  if (!el) return [];
  const out: Element[] = [];
  for (const c of Array.from(el.children)) {
    if (c.localName === localName) out.push(c);
  }
  return out;
}

function child(el: Element | null, localName: string): Element | null {
  if (!el) return null;
  for (const c of Array.from(el.children)) {
    if (c.localName === localName) return c;
  }
  return null;
}

function wAttr(el: Element | null, name: string): string | null {
  if (!el) return null;
  return el.getAttribute('w:' + name);
}

/** 读取属性并转为整数，缺省/非法时返回 null */
function wAttrInt(el: Element | null, name: string): number | null {
  const v = wAttr(el, name);
  if (v == null) return null;
  const n = Number(v);
  return Number.isFinite(n) ? n : null;
}

function normalizeAlign(val: string | null): string | null {
  if (!val) return null;
  if (val === 'both') return 'justify';
  if (
    val === 'left' ||
    val === 'right' ||
    val === 'center' ||
    val === 'justify'
  )
    return val;
  return null;
}

/** w:ind 的 w:left（缇，1/20 pt）→ 我们的缩进级别（每级 720 缇 ≈ Word "增加缩进" 一次）*/
function parseIndentLevel(pPr: Element | null): number {
  const indEl = child(pPr, 'ind');
  const leftTwips = Number(
    indEl?.getAttribute('w:left') ?? indEl?.getAttribute('w:start') ?? '0'
  );
  if (!leftTwips || Number.isNaN(leftTwips)) return 0;
  return Math.max(0, Math.min(8, Math.round(leftTwips / 720)));
}

/** w:spacing 的 w:line（仅 lineRule=auto 时是"倍数×240"）→ 行距倍数，如 1 / 1.5 / 2 */
function parseLineSpacing(pPr: Element | null): number | null {
  const spacingEl = child(pPr, 'spacing');
  if (!spacingEl) return null;
  const lineRule = spacingEl.getAttribute('w:lineRule');
  const line = Number(spacingEl.getAttribute('w:line') ?? '0');
  if (!line || Number.isNaN(line)) return null;
  if (lineRule && lineRule !== 'auto') return null; // exact/atLeast 是绝对尺寸，暂不处理
  const multiplier = Math.round((line / 240) * 100) / 100;
  return multiplier > 0 ? multiplier : null;
}

async function readZipText(zip: JSZip, path: string): Promise<string | null> {
  const entry = zip.file(path);
  if (!entry) return null;
  return entry.async('text');
}

function parseXml(
  parser: DOMParser,
  text: string,
  label: string,
  warnings: string[]
): Document | null {
  const doc = parser.parseFromString(text, 'application/xml');
  if (doc.getElementsByTagName('parsererror').length > 0) {
    warnings.push(`[warning] ${label} 存在 XML 格式问题，相关信息可能不完整`);
    return null;
  }
  return doc;
}

// ---------------------------------------------------------------------------
// styles.xml —— 命名样式（含继承链解析）
// ---------------------------------------------------------------------------

interface RunProps {
  bold?: boolean;
  italic?: boolean;
  underline?: boolean;
  strike?: boolean;
  color?: string;
  sizeHalfPt?: number;
  fontFamily?: string;
  highlight?: string;
}

interface StyleInfo {
  name: string;
  basedOn: string | null;
  headingLevel: number | null;
  rPr: RunProps;
  align: string | null;
  /** 段落样式自身的编号定义（ Word 多级列表常定义在样式上） */
  numPr: Element | null;
  /** 段落样式自身携带的缩进级别（Word 目录 TOC1/TOC2… 的层级缩进定义在样式上） */
  indent: number | null;
}

function extractRunProps(rPrEl: Element | null): RunProps {
  const props: RunProps = {};
  if (!rPrEl) return props;

  const bEl = child(rPrEl, 'b');
  if (bEl && wAttr(bEl, 'val') !== 'false' && wAttr(bEl, 'val') !== '0')
    props.bold = true;

  const iEl = child(rPrEl, 'i');
  if (iEl && wAttr(iEl, 'val') !== 'false' && wAttr(iEl, 'val') !== '0')
    props.italic = true;

  const uEl = child(rPrEl, 'u');
  const uVal = wAttr(uEl, 'val');
  if (uEl && uVal && uVal !== 'none') props.underline = true;

  const strikeEl = child(rPrEl, 'strike');
  if (
    strikeEl &&
    wAttr(strikeEl, 'val') !== 'false' &&
    wAttr(strikeEl, 'val') !== '0'
  )
    props.strike = true;

  const colorEl = child(rPrEl, 'color');
  const colorVal = wAttr(colorEl, 'val');
  if (colorVal && colorVal !== 'auto' && /^[0-9a-fA-F]{6}$/.test(colorVal))
    props.color = `#${colorVal}`;

  const szEl = child(rPrEl, 'sz');
  const szVal = wAttr(szEl, 'val');
  if (szVal && !Number.isNaN(Number(szVal))) props.sizeHalfPt = Number(szVal);

  const fontsEl = child(rPrEl, 'rFonts');
  const fam =
    fontsEl?.getAttribute('w:ascii') ??
    fontsEl?.getAttribute('w:eastAsia') ??
    fontsEl?.getAttribute('w:hAnsi');
  if (fam) props.fontFamily = fam;

  const hlEl = child(rPrEl, 'highlight');
  const hlVal = wAttr(hlEl, 'val');
  if (hlVal && hlVal !== 'none') props.highlight = hlVal;

  return props;
}

function parseStylesXml(doc: Document): Map<string, StyleInfo> {
  const raw = new Map<
    string,
    {
      name: string;
      basedOn: string | null;
      headingLevel: number | null;
      rPr: RunProps;
      align: string | null;
      numPr: Element | null;
      indent: number | null;
    }
  >();

  for (const styleEl of Array.from(doc.getElementsByTagNameNS(W, 'style'))) {
    const id = styleEl.getAttribute('w:styleId');
    const type = styleEl.getAttribute('w:type');
    if (!id || (type !== 'paragraph' && type !== 'character')) continue;

    const name = wAttr(child(styleEl, 'name'), 'val') ?? id;
    const basedOn = wAttr(child(styleEl, 'basedOn'), 'val');
    const headingMatch = /^heading\s*(\d)/i.exec(name);
    const stylePPr = child(styleEl, 'pPr');
    const align = normalizeAlign(wAttr(child(stylePPr, 'jc'), 'val'));
    const numPr = child(stylePPr, 'numPr');
    // 仅当样式显式声明了非零左缩进时才视为"携带缩进"（null 表示未指定，便于继承）
    const indEl = child(stylePPr, 'ind');
    const indLeft = Number(
      indEl?.getAttribute('w:left') ?? indEl?.getAttribute('w:start') ?? '0'
    );
    const hasIndent =
      !!indEl && !!indLeft && !Number.isNaN(indLeft) && indLeft > 0;

    raw.set(id, {
      name,
      basedOn,
      headingLevel: headingMatch
        ? Math.min(9, Number(headingMatch[1]))
        : id === 'Title'
          ? 1
          : null,
      rPr: extractRunProps(child(styleEl, 'rPr')),
      align,
      numPr,
      indent: hasIndent ? Math.max(0, Math.min(8, Math.round(indLeft / 720))) : null,
    });
  }

  const resolved = new Map<string, StyleInfo>();
  const resolving = new Set<string>();

  function resolve(id: string): StyleInfo {
    const cached = resolved.get(id);
    if (cached) return cached;
    const info = raw.get(id);
    if (!info) {
      const fallback: StyleInfo = {
        name: id,
        basedOn: null,
        headingLevel: null,
        rPr: {},
        align: null,
        numPr: null,
        indent: null,
      };
      resolved.set(id, fallback);
      return fallback;
    }
    if (resolving.has(id)) {
      // 循环继承兜底，避免死循环
      const flat: StyleInfo = {
        name: info.name,
        basedOn: null,
        headingLevel: info.headingLevel,
        rPr: info.rPr,
        align: info.align,
        numPr: info.numPr,
        indent: info.indent,
      };
      resolved.set(id, flat);
      return flat;
    }
    resolving.add(id);
    const base = info.basedOn ? resolve(info.basedOn) : null;
    const merged: StyleInfo = {
      name: info.name,
      basedOn: info.basedOn,
      headingLevel: info.headingLevel,
      rPr: { ...(base?.rPr ?? {}), ...info.rPr },
      align: info.align ?? base?.align ?? null,
      // 样式的编号通常直接挂在具体样式上；按继承链兜底取最先出现的 numPr
      numPr: info.numPr ?? base?.numPr ?? null,
      indent: info.indent ?? base?.indent ?? null,
    };
    resolving.delete(id);
    resolved.set(id, merged);
    return merged;
  }

  for (const id of raw.keys()) resolve(id);
  return resolved;
}

// ---------------------------------------------------------------------------
// numbering.xml —— 列表格式（区分有序/无序，并识别具体样式）
// ---------------------------------------------------------------------------

export interface NumberingLevelInfo {
  ordered: boolean;
  numberFormat: string; // 'decimal' | 'lower-alpha' | 'upper-alpha' | 'lower-roman' | 'upper-roman'
  bulletStyle: string; // 'disc' | 'circle' | 'square'
  lvlText: string | null; // 原始 w:lvlText（如 "%1." / "%1.%2."），用于标题编号文本
  noMarker: boolean; // numFmt="none"：该级不显示任何项目符号/编号，仅保留缩进层级
  /** 布局：w:ind w:left（段落左缩进，twips） */
  indLeftTw: number;
  /** 布局：w:ind w:hanging（悬挂缩进，twips；会替换 w:left 作为真正的左缩进/对齐基准） */
  indHangingTw: number;
  /** 字体特征：编号/符号标记自身的 rPr（Symbol 等字体 + 字号，twips 半磅） */
  font: { fontFamily: string | null; sizeHalfPt: number | null } | null;
  /** 本级别起始值（w:start，restart 基准） */
  start: number;
}

const OOXML_NUMFMT_MAP: Record<string, string> = {
  decimal: 'decimal',
  lowerLetter: 'lower-alpha',
  upperLetter: 'upper-alpha',
  lowerRoman: 'lower-roman',
  upperRoman: 'upper-roman',
};

/** 项目符号的具体形状（圆点/空心圆/方块）由 w:lvlText 的实际字符决定，这里做启发式识别 */
function detectBulletStyle(lvlText: string | null): string {
  if (!lvlText) return 'disc';
  if (/[○ｏo]/.test(lvlText)) return 'circle';
  if (/[▪■□§]/.test(lvlText)) return 'square';
  return 'disc';
}

function parseNumberingXml(
  doc: Document
): Map<string, Map<number, NumberingLevelInfo>> {
  const abstractFormats = new Map<string, Map<number, NumberingLevelInfo>>();

  for (const abs of Array.from(doc.getElementsByTagNameNS(W, 'abstractNum'))) {
    const absId = abs.getAttribute('w:abstractNumId');
    if (absId == null) continue;
    const levelMap = new Map<number, NumberingLevelInfo>();
    for (const lvl of children(abs, 'lvl')) {
      const ilvl = Number(lvl.getAttribute('w:ilvl') ?? '0');
      const fmt = wAttr(child(lvl, 'numFmt'), 'val') ?? 'bullet';
      const ordered = fmt !== 'bullet' && fmt !== 'none';
      const lvlText = wAttr(child(lvl, 'lvlText'), 'val');
      // 布局：w:ind w:left / w:hanging（twips）
      const ind = child(lvl, 'ind');
      const indLeftTw = wAttrInt(ind, 'left') ?? wAttrInt(ind, 'start') ?? 0;
      const indHangingTw = wAttrInt(ind, 'hanging') ?? 0;
      // 字体特征：级别自身的 rPr（通常 Symbol 字体 + 字号），用于编号/符号渲染
      const rPr = child(lvl, 'rPr');
      const rFonts = rPr ? child(rPr, 'rFonts') : null;
      const fontFamily =
        rFonts?.getAttribute('w:ascii') ||
        rFonts?.getAttribute('w:hAnsi') ||
        rFonts?.getAttribute('w:eastAsia') ||
        null;
      const sz = rPr ? child(rPr, 'sz') : null;
      const sizeHalfPt = sz ? wAttrInt(sz, 'val') : null;
      levelMap.set(ilvl, {
        ordered,
        numberFormat: OOXML_NUMFMT_MAP[fmt] ?? 'decimal',
        bulletStyle: detectBulletStyle(lvlText),
        lvlText,
        noMarker: fmt === 'none',
        indLeftTw,
        indHangingTw,
        font:
          fontFamily || sizeHalfPt
            ? { fontFamily, sizeHalfPt }
            : null,
        start: wAttrInt(child(lvl, 'start'), 'val') ?? 1,
      });
    }
    abstractFormats.set(absId, levelMap);
  }

  const result = new Map<string, Map<number, NumberingLevelInfo>>();
  for (const num of Array.from(doc.getElementsByTagNameNS(W, 'num'))) {
    const numId = num.getAttribute('w:numId');
    const absId = wAttr(child(num, 'abstractNumId'), 'val');
    if (numId == null || absId == null) continue;
    const baseMap = abstractFormats.get(absId);
    if (!baseMap) continue;
    // 每个 numId 克隆一份 level 映射，避免把 startOverride 写进共享的 abstract 映射；
    // 再套用 <w:lvlOverride><w:startOverride w:val=…/> 作为该 num 实例各级别的起始值。
    const clone = new Map<number, NumberingLevelInfo>();
    for (const [lvl, info] of baseMap) clone.set(lvl, { ...info });
    for (const lo of Array.from(num.getElementsByTagNameNS(W, 'lvlOverride'))) {
      const lvl = wAttrInt(lo, 'ilvl');
      const so = child(lo, 'startOverride');
      const val = so ? wAttrInt(so, 'val') : null;
      if (lvl == null || val == null) continue;
      const cur = clone.get(lvl);
      if (cur) clone.set(lvl, { ...cur, start: val });
    }
    result.set(numId, clone);
  }
  return result;
}

// ---------------------------------------------------------------------------
// 关系文件 + 媒体资源
// ---------------------------------------------------------------------------

function parseRelsXml(doc: Document): Map<string, string> {
  const out = new Map<string, string>();
  for (const el of Array.from(doc.getElementsByTagName('Relationship'))) {
    const id = el.getAttribute('Id');
    const target = el.getAttribute('Target');
    if (id && target) out.set(id, target);
  }
  return out;
}

async function readMedia(
  zip: JSZip,
  rels: Map<string, string>
): Promise<Map<string, string>> {
  const out = new Map<string, string>();
  for (const [rId, target] of rels) {
    if (!/\.(png|jpe?g|gif|bmp)$/i.test(target)) continue;
    const path = target.startsWith('media/')
      ? `word/${target}`
      : target.replace(/^\.\.\//, 'word/');
    const entry = zip.file(path);
    if (!entry) continue;
    const base64 = await entry.async('base64');
    const ext = (path.split('.').pop() ?? 'png').toLowerCase();
    const mime =
      ext === 'jpg' || ext === 'jpeg'
        ? 'image/jpeg'
        : ext === 'gif'
          ? 'image/gif'
          : ext === 'bmp'
            ? 'image/bmp'
            : 'image/png';
    out.set(rId, `data:${mime};base64,${base64}`);
  }
  return out;
}

// ---------------------------------------------------------------------------
// comments.xml
// ---------------------------------------------------------------------------

function extractPlainText(el: Element): string {
  const parts: string[] = [];
  for (const t of Array.from(el.getElementsByTagNameNS(W, 't')))
    parts.push(t.textContent ?? '');
  return parts.join('');
}

/**
 * 递归将 OMML 元素转换为 LaTeX 字符串。
 * 支持：sub/sup/sSub/sSup 下标上标、f 分数、nary 累加/求和、bar 上划线、rad 根号等常用结构。
 */
function ommlToLatex(el: Element): string {
  const tag = el.localName;

  // 文本节点（递归时也会遇到纯文本）
  if (el.nodeType === 3) return el.textContent ?? '';

  // 普通 OMML 运行
  if (tag === 'r') {
    return Array.from(el.children)
      .map((c) => ommlToLatex(c))
      .join('');
  }

  // 纯文本
  if (tag === 't') return el.textContent ?? '';

  // 下标元素 (subscript) — 通常出现在 sSub 内部
  if (tag === 'sub') {
    const sub = child(el, 'sub');
    if (sub) return `_{${ommlToLatex(sub)}}`;
    // 无子 sub 时，遍历当前元素的所有子节点作为兜底
    return `_{${Array.from(el.children).map(ommlToLatex).join('')}}`;
  }

  // 上标元素 (superscript)
  if (tag === 'sup') {
    const sup = child(el, 'sup');
    if (sup) return `^{${ommlToLatex(sup)}}`;
    return `^{${Array.from(el.children).map(ommlToLatex).join('')}}`;
  }

  // 组合下标（base + sub）
  if (tag === 'sSub') {
    const e = child(el, 'e');
    const sub = child(el, 'sub');
    const base = e ? ommlToLatex(e) : '';
    const subscript = sub ? ommlToLatex(sub) : '';
    return `${base}_{${subscript}}`;
  }

  // 组合上标（base + sup）
  if (tag === 'sSup') {
    const e = child(el, 'e');
    const sup = child(el, 'sup');
    const base = e ? ommlToLatex(e) : '';
    const superscript = sup ? ommlToLatex(sup) : '';
    return `${base}^{${superscript}}`;
  }

  // 上下标同时
  if (tag === 'sSubSup') {
    const e = child(el, 'e');
    const sub = child(el, 'sub');
    const sup = child(el, 'sup');
    const base = e ? ommlToLatex(e) : '';
    const subscript = sub ? ommlToLatex(sub) : '';
    const superscript = sup ? ommlToLatex(sup) : '';
    return `${base}_{${subscript}}^{${superscript}}`;
  }

  // 分数
  if (tag === 'f') {
    const num = child(el, 'num');
    const den = child(el, 'den');
    const numerator = num ? ommlToLatex(num) : '';
    const denominator = den ? ommlToLatex(den) : '';
    return `\\frac{${numerator}}{${denominator}}`;
  }

  // 上划线
  if (tag === 'bar') {
    const e = child(el, 'e');
    return `\\overline{${e ? ommlToLatex(e) : ''}}`;
  }

  // 根号
  if (tag === 'rad') {
    const deg = child(el, 'deg');
    const e = child(el, 'e');
    const degree = deg ? ommlToLatex(deg) : '';
    const radicand = e ? ommlToLatex(e) : '';
    if (degree && degree !== '2') {
      return `\\sqrt[${degree}]{${radicand}}`;
    }
    return `\\sqrt{${radicand}}`;
  }

  // 极限
  if (tag === 'lim') {
    const limEl = child(el, 'lim');
    const e = child(el, 'e');
    const limit = limEl ? ommlToLatex(limEl) : '';
    const expr = e ? ommlToLatex(e) : '';
    return `\\lim_{${limit}}{${expr}}`;
  }

  // n-ary 运算符（∑ ∑ ∏ ∫ 等）
  if (tag === 'nary') {
    const sub = child(el, 'sub');
    const sup = child(el, 'sup');
    const e = child(el, 'e');
    const limLow = sub ? ommlToLatex(sub) : '';
    const limUpp = sup ? ommlToLatex(sup) : '';
    const body = e ? ommlToLatex(e) : '';
    // 通常 nary 包含一个 char 元素表示运算符，但我们直接拼接
    return `${body}_{${limLow}}^{${limUpp}}`;
  }

  // 括号/包围
  if (tag === 'd') {
    // delimiter wrapper, just extract inner
    const e = child(el, 'e');
    return e ? ommlToLatex(e) : '';
  }

  // 其他容器逐个处理子元素
  const children = Array.from(el.children);
  if (children.length > 0) {
    return children.map((c) => ommlToLatex(c)).join('');
  }

  // 兜底：返回文本内容
  return el.textContent ?? '';
}

function parseCommentsXml(doc: Document): DocxComment[] {
  const out: DocxComment[] = [];
  for (const el of Array.from(doc.getElementsByTagNameNS(W, 'comment'))) {
    const id = Number(el.getAttribute('w:id'));
    if (Number.isNaN(id)) continue;
    out.push({
      id,
      author: el.getAttribute('w:author') || '匿名',
      date: el.getAttribute('w:date'),
      text: extractPlainText(el).trim(),
    });
  }
  return out;
}

// ---------------------------------------------------------------------------
// document.xml —— 主体解析
// ---------------------------------------------------------------------------

interface ParseCtx {
  styles: Map<string, StyleInfo>;
  numbering: Map<string, Map<number, NumberingLevelInfo>>;
  /** 标题编号计数器：numId -> 各级别当前值（仅用于生成标题前缀编号） */
  numberingCounters: Map<string, number[]>;
  rels: Map<string, string>;
  media: Map<string, string>;
  warnings: string[];
  nextBlockId: () => string;
  nextCellId: () => string;
  /** 段落级分节符：sectPr 所在的段落是某一节的最后一块，endBlock 为该节结束(exclusive)的 flat 块序号 */
  sectBreaks: { sectPr: Element; endBlock: number }[];
}

type FlatBlock =
  | { kind: 'block'; node: any }
  | {
      kind: 'listItem';
      level: number;
      ordered: boolean;
      numberFormat: string;
      bulletStyle: string;
      numId: string;
      lvlText: string | null;
      start: number;
      indLeftTw: number;
      indHangingTw: number;
      font: { fontFamily: string | null; sizeHalfPt: number | null } | null;
      node: any;
      /** 连续编号的语义分组键：仅"表/图题注"这类带字面前缀（如 表%1 / 图%1.）的有序列表
       *  会被赋予键，从而在整个文档范围内跨组连续编号（表1,表2…图1,图2…）；普通列表为 null，
       *  仍按"每个独立组从 w:start 重新开始"。 */
      seqGroup: string | null;
    };

function buildMarksFromRunProps(p: RunProps, activeComments: number[]): any[] {
  const marks: any[] = [];
  if (p.bold) marks.push({ type: 'strong' });
  if (p.italic) marks.push({ type: 'em' });
  if (p.underline) marks.push({ type: 'underline' });
  if (p.strike) marks.push({ type: 'strike' });
  if (p.color || p.fontFamily || p.sizeHalfPt || p.highlight) {
    marks.push({
      type: 'docxStyle',
      attrs: {
        color: p.color ?? null,
        fontFamily: p.fontFamily ?? null,
        sizeHalfPt: p.sizeHalfPt ?? null,
        highlight: p.highlight ?? null,
      },
    });
  }
  for (const id of activeComments)
    marks.push({ type: 'comment', attrs: { id } });
  return marks;
}

function parseDrawing(drawingEl: Element, ctx: ParseCtx): any | null {
  const blip = drawingEl.getElementsByTagNameNS(A_NS, 'blip')[0];
  const rId = blip?.getAttribute('r:embed');
  if (!rId) return null;
  const dataUrl = ctx.media.get(rId);
  if (!dataUrl) {
    ctx.warnings.push('[warning] 存在未能解析的内嵌图片');
    return null;
  }
  // 按 OOXML 图片实际放置尺寸（wp:extent，单位 EMU）渲染，
  // 而不是用图片的分辨率原尺寸。仅取宽（高度按真实像素等比，保证不变形且不撑破页面）。
  let width: number | null = null;
  const extent = drawingEl.getElementsByTagNameNS(WP_NS, 'extent')[0];
  if (extent) {
    const cx = extent.getAttribute('cx');
    if (cx) {
      const px = Math.round(Number(cx) * EMU_TO_PX);
      if (Number.isFinite(px) && px > 0) width = px;
    }
  }
  const attrs: Record<string, unknown> = { src: dataUrl, alt: '' };
  if (width != null) attrs.width = width;
  return { type: 'image', attrs };
}

function parseRun(
  rEl: Element,
  ctx: ParseCtx,
  inheritedRPr: RunProps,
  activeComments: number[]
): any[] {
  const ownProps = extractRunProps(child(rEl, 'rPr'));
  const merged: RunProps = { ...inheritedRPr, ...ownProps };
  const marks = buildMarksFromRunProps(merged, activeComments);

  const out: any[] = [];
  for (const node of Array.from(rEl.children)) {
    switch (node.localName) {
      case 't': {
        const text = node.textContent ?? '';
        if (text)
          out.push({
            type: 'text',
            text,
            marks: marks.length ? marks : undefined,
          });
        break;
      }
      case 'br':
        out.push({ type: 'hard_break' });
        break;
      case 'tab': {
        const text = '\t';
        out.push({
          type: 'text',
          text,
          marks: marks.length ? marks : undefined,
        });
        break;
      }
      case 'drawing': {
        const img = parseDrawing(node, ctx);
        if (img) out.push(img);
        break;
      }
      default:
        break;
    }
  }
  return out;
}

function parseInlineContent(
  containerEl: Element,
  ctx: ParseCtx,
  inheritedRPr: RunProps
): any[] {
  const out: any[] = [];
  const activeComments: number[] = [];

  const walk = (el: Element) => {
    for (const node of Array.from(el.children)) {
      switch (node.localName) {
        case 'r':
          out.push(...parseRun(node, ctx, inheritedRPr, activeComments));
          break;
        case 'hyperlink': {
          const rId = node.getAttribute('r:id');
          const anchor = node.getAttribute('w:anchor');
          const href = rId
            ? (ctx.rels.get(rId) ?? null)
            : anchor
              ? `#${anchor}`
              : null;
          const before = out.length;
          for (const inner of Array.from(node.children)) {
            if (inner.localName === 'r')
              out.push(...parseRun(inner, ctx, inheritedRPr, activeComments));
          }
          if (href) {
            for (let i = before; i < out.length; i++) {
              const n = out[i];
              if (n.type === 'text')
                n.marks = [
                  ...(n.marks ?? []),
                  { type: 'link', attrs: { href, title: null } },
                ];
            }
          }
          break;
        }
        case 'commentRangeStart': {
          const id = Number(node.getAttribute('w:id'));
          if (!Number.isNaN(id)) activeComments.push(id);
          break;
        }
        case 'commentRangeEnd': {
          const id = Number(node.getAttribute('w:id'));
          const idx = activeComments.indexOf(id);
          if (idx >= 0) activeComments.splice(idx, 1);
          break;
        }
        case 'ins':
          walk(node); // 修订-插入：展开为正常内容
          break;
        case 'del':
          break; // 修订-删除：不进入最终结构（结构优先原则下的合理简化）
        case 'sdt': {
          const c = child(node, 'sdtContent');
          if (c) walk(c);
          break;
        }
        case 'smartTag':
          walk(node);
          break;
        case 'oMath': {
          const latex = ommlToLatex(node);
          if (latex) {
            out.push({ type: 'math_inline', attrs: { latex } });
          }
          break;
        }
        case 'oMathPara': {
          const oMath = child(node, 'oMath');
          const latex = oMath ? ommlToLatex(oMath) : '';
          if (latex) {
            out.push({ type: 'math_block', attrs: { latex } });
          }
          break;
        }
        default:
          break;
      }
    }
  };

  walk(containerEl);
  return out;
}

// ---------------------------------------------------------------------------
// 标题自动编号：维护每份 numId 的计数器并格式化为可读文本
// ---------------------------------------------------------------------------

function toRoman(num: number): string {
  if (num <= 0 || num > 3999) return String(num);
  const values = [1000, 900, 500, 400, 100, 90, 50, 40, 10, 9, 5, 4, 1];
  const symbols = [
    'M',
    'CM',
    'D',
    'CD',
    'C',
    'XC',
    'L',
    'XL',
    'X',
    'IX',
    'V',
    'IV',
    'I',
  ];
  let out = '';
  for (let i = 0; i < values.length; i++) {
    while (num >= values[i]) {
      num -= values[i];
      out += symbols[i];
    }
  }
  return out;
}

function formatListCounter(format: string, value: number): string {
  const n = value > 0 ? value : 1;
  switch (format) {
    case 'lower-alpha':
      return String.fromCharCode(97 + ((n - 1) % 26));
    case 'upper-alpha':
      return String.fromCharCode(65 + ((n - 1) % 26));
    case 'lower-roman':
      return toRoman(n).toLowerCase();
    case 'upper-roman':
      return toRoman(n).toUpperCase();
    case 'decimal':
    default:
      return String(n);
  }
}

/**
 * 根据多级列表的 w:lvlText 模板生成标题前缀编号。
 * 例如 lvlText="%1.%2." 的二级标题会渲染成 "1.2."。
 * 计数器按文档顺序维护，适合常见的连续标题编号场景；
 * 复杂重启/继续规则按 "结构优先" 原则做近似处理。
 */
function buildHeadingNumberText(
  ctx: ParseCtx,
  numId: string,
  ilvl: number,
  lvlInfo: NumberingLevelInfo | undefined
): string | null {
  if (!lvlInfo || !lvlInfo.ordered) return null;

  let counters = ctx.numberingCounters.get(numId);
  if (!counters) {
    counters = [];
    ctx.numberingCounters.set(numId, counters);
  }
  while (counters.length <= ilvl) counters.push(0);
  counters[ilvl] = (counters[ilvl] ?? 0) + 1;
  counters.splice(ilvl + 1);

  const text = lvlInfo.lvlText ?? (ilvl === 0 ? '%1.' : '%1.%2.');
  return text.replace(/%(\d+)/g, (_, nStr) => {
    const refLevel = Number(nStr) - 1;
    if (refLevel < 0 || refLevel > ilvl) return `%${nStr}`;
    const refInfo = ctx.numbering.get(numId)?.get(refLevel);
    const fmt = refInfo?.numberFormat ?? 'decimal';
    const val = counters[refLevel] ?? 1;
    return formatListCounter(fmt, val);
  });
}

function isCaptionLevel(lvlText: string | null): string | null {
  if (!lvlText) return null;
  // 去掉 %n 占位符得到字面骨架：仅当包含 图/表 字面前缀（表%1 / 图%1.）视为题注。
  const skeleton = lvlText.replace(/%\d+/g, '').trim();
  if (!skeleton) return null;
  if (skeleton.includes('图') || skeleton.includes('表')) return `caption:${skeleton}`;
  return null;
}

function foldListBlocks(
  flat: FlatBlock[],
  counters?: Map<string, number>
): any[] {
  const out: any[] = [];
  // 题注（表/图）的全局连续序列：同一 seqGroup 在整个文档范围内跨组延续。
  const captionSeq = new Map<string, number>();
  const stack: {
    level: number;
    ordered: boolean;
    numberFormat: string;
    bulletStyle: string;
    numId: string;
    lvlText: string | null;
    start: number;
    indLeftTw: number;
    indHangingTw: number;
    font: { fontFamily: string | null; sizeHalfPt: number | null } | null;
    n: number; // 本组（顶层）当前的列表计数器（仅有序列表使用）
    seqGroup: string | null;
    items: any[];
  }[] = [];

  const closeTop = () => {
    const top = stack.pop();
    if (!top) return;
    // 题注（表/图）：渲染为居中的普通段落，编号文本内联在标题前（"图1 系统功能清单"），
    // 而不是用悬挂缩进的原生 marker 盒子（那会让编号脱离到最左侧，标题却居中）。
    if (top.ordered && top.seqGroup) {
      const blocks = top.items.map((item: any) => {
        const inner = item.content?.[0] ?? { type: 'paragraph', attrs: {}, content: [] };
        const markerText: string | undefined = item.attrs?.marker;
        let content = inner.content ?? [];
        if (markerText != null) {
          content = [{ type: 'text', text: `${markerText} ` }, ...content];
        }
        return { ...inner, content };
      });
      if (stack.length) {
        stack[stack.length - 1].items[stack[stack.length - 1].items.length - 1].content.push(...blocks);
      } else {
        out.push(...blocks);
      }
      if (counters && top.ordered) {
        counters.set(top.numId, top.n);
      }
      captionSeq.set(top.seqGroup, top.n);
      return;
    }
    const listNode = top.ordered
      ? {
          type: 'ordered_list',
          attrs: {
            numberFormat: top.numberFormat,
            literal: true, // 用计算出的编号文本渲染（贴合 lvlText 模板），非原生 marker
            indent: top.indLeftTw,
            hanging: top.indHangingTw,
            font: top.font,
          },
          content: top.items,
        }
      : {
          type: 'bullet_list',
          attrs: {
            bulletStyle: top.bulletStyle,
            indent: top.indLeftTw,
            hanging: top.indHangingTw,
            font: top.font,
          },
          content: top.items,
        };
    if (stack.length) {
      const parentItems = stack[stack.length - 1].items;
      parentItems[parentItems.length - 1].content.push(listNode);
    } else {
      out.push(listNode);
    }
    // 文档级列表计数器：同一 numId 视为同一列表实例，持续递增；写回以便跨组/跨表格延续。
    if (counters && top.ordered) {
      counters.set(top.numId, top.n);
    }
    // 题注（表/图）完整序列：同一 seqGroup 跨组持续递增。
    if (top.ordered && top.seqGroup) {
      captionSeq.set(top.seqGroup, top.n);
    }
  };

  /** 依据 lvlText 模板生成单个有序条目的显示编号文本（如 "%1." → "1."；"%1)" → "1)"）。
   *  %k 引用第 k 层编号；单层列表常用 %1。传 ancestors 为当前层级链编号。 */
  const markerFor = (
    b: FlatBlock & { kind: 'listItem' },
    ancestors: number[]
  ): string => {
    const template = b.lvlText ?? '%1.';
    const fmt = b.numberFormat;
    return template.replace(/%(\d+)/g, (_s, dRaw: string) => {
      const idx = Number(dRaw) - 1;
      const val = ancestors[idx] ?? ancestors[0] ?? b.start;
      return formatListCounter(fmt, val);
    });
  };

  /** 桌表格按文档顺序就地展开：整个表格作为独立编号上下文（语义容器），
   *  为它单独建一份计数器，单元格/跨行同 numId 延续；表格之间/相对正文则互不串号
   *  （即每个表格的编号都从自身 w:start 重新开始）。 */
  const expandTable = (tbl: any): any => {
    // Column-aware list continuation: a numbered run only continues down the
    // SAME column of consecutive rows (and within a cell). A numbered list in a
    // different cell/column (e.g. 先决条件 vs 测试规程's 序号 column — both
    // share a numId but are conceptually independent groups) resets to its
    // w:start instead of merging into one Table-wide sequence. State is keyed
    // per column so processing other cells in a row never clobbers the column
    // that must keep running into the next row.
    const colState = new Map<number, { row: number; numId: string; next: number }>();
    return {
      type: 'table',
      content: (tbl.content as any[]).map((row: any, r: number) => ({
        type: 'table_row',
        content: (row.content as any[]).map((cell: any) => {
          const col = (cell.colStart as number) ?? 0;
          const counter = new Map<string, number>();
          const flat = cell.content as FlatBlock[];
          const firstLi = flat.find(
            (b): b is FlatBlock & { kind: 'listItem' } => b.kind === 'listItem',
          );
          const fromPrev = colState.get(col);
          // Continue only if the previous row's SAME column ended an open run
          // with the same numId.
          if (
            firstLi &&
            fromPrev &&
            r === fromPrev.row + 1 &&
            firstLi.numId === fromPrev.numId
          ) {
            counter.set(firstLi.numId, fromPrev.next);
          }
          const content = foldListBlocks(flat, counter);
          // Re-evaluate this column's run state: continuation applies only when
          // this cell ENDS with a list item (run still open toward next row).
          let lastLi: FlatBlock & { kind: 'listItem' } | undefined;
          for (let i = flat.length - 1; i >= 0; i--) {
            if (flat[i].kind === 'listItem') {
              lastLi = flat[i] as FlatBlock & { kind: 'listItem' };
              break;
            }
          }
          if (lastLi) {
            colState.set(col, {
              row: r,
              numId: lastLi.numId,
              next: counter.get(lastLi.numId) ?? lastLi.start,
            });
          } else {
            colState.delete(col);
          }
          return { type: cell.type, attrs: cell.attrs, content };
        }),
      })),
    };
  };

  const seedFor = (b: FlatBlock & { kind: 'listItem' }): number => {
    if (!b.ordered) return b.start;
    // 题注优先：同一 seqGroup 延续（表1,表2…）；否则回退到表格列计数器；否则从 w:start。
    if (b.seqGroup && captionSeq.has(b.seqGroup)) return captionSeq.get(b.seqGroup)!;
    if (counters && counters.has(b.numId)) return counters.get(b.numId)!;
    return b.start;
  };

  for (const b of flat) {
    if (b.kind !== 'listItem') {
      while (stack.length) closeTop();
      out.push(b.node && b.node.__rawTable ? expandTable(b.node) : b.node);
      continue;
    }
    while (stack.length && b.level < stack[stack.length - 1].level) closeTop();
    if (!stack.length || b.level > stack[stack.length - 1].level) {
      const seedNum = seedFor(b);
      stack.push({
        level: b.level,
        ordered: b.ordered,
        numberFormat: b.numberFormat,
        bulletStyle: b.bulletStyle,
        numId: b.numId,
        lvlText: b.lvlText,
        start: b.start,
        indLeftTw: b.indLeftTw,
        indHangingTw: b.indHangingTw,
        font: b.font,
        seqGroup: b.seqGroup,
        n: seedNum,
        items: [],
      });
    } else if (
      stack[stack.length - 1].ordered !== b.ordered ||
      stack[stack.length - 1].numId !== b.numId
    ) {
      closeTop();
      const seedNum = seedFor(b);
      stack.push({
        level: b.level,
        ordered: b.ordered,
        numberFormat: b.numberFormat,
        bulletStyle: b.bulletStyle,
        numId: b.numId,
        lvlText: b.lvlText,
        start: b.start,
        indLeftTw: b.indLeftTw,
        indHangingTw: b.indHangingTw,
        font: b.font,
        seqGroup: b.seqGroup,
        n: seedNum,
        items: [],
      });
    }
    const top = stack[stack.length - 1];
    // 计算该条目显示用的编号文本（有序列表；项目符号不在此处理）
    let marker: string | undefined;
    if (top.ordered) {
      const ancestors = stack.map((s) => s.n);
      marker = markerFor(b, ancestors);
      top.n += 1;
    }
    top.items.push({
      type: 'list_item',
      attrs: marker != null ? { marker } : undefined,
      content: [b.node],
    });
  }
  while (stack.length) closeTop();
  return out;
}

function parseParagraphEl(pEl: Element, ctx: ParseCtx): FlatBlock {
  const pPr = child(pEl, 'pPr');
  const styleId = wAttr(child(pPr, 'pStyle'), 'val');
  const styleInfo = styleId ? (ctx.styles.get(styleId) ?? null) : null;

  // 段落自身的 numPr 优先；否则继承样式链上的编号定义（Word 多级列表常见）
  const directNumPr = child(pPr, 'numPr');
  const effectiveNumPr = directNumPr ?? styleInfo?.numPr ?? null;
  const ilvl = effectiveNumPr
    ? Number(wAttr(child(effectiveNumPr, 'ilvl'), 'val') ?? '0')
    : null;
  const numId = effectiveNumPr
    ? wAttr(child(effectiveNumPr, 'numId'), 'val')
    : null;

  const directAlign = normalizeAlign(wAttr(child(pPr, 'jc'), 'val'));
  const align = directAlign ?? styleInfo?.align ?? null;
  // 缩进取"段落直写 w:ind"优先；否则回退到段落样式的缩进（Word 目录 TOC1/TOC2…
  // 的层级缩进定义在样式上，若不回退，内容里的目录会渲染成平铺无层次）。
  // 列表段落（numId/ilvl）的缩进由列表结构决定，不再叠加样式缩进以免双重缩进。
  const directIndent = parseIndentLevel(pPr);
  // numFmt="none" 的列表级：Word 不显示任何项目符号/编号，仅保留缩进层级。
  // 这类段落不应进入 <ul>/<ol>，而是渲染成带层级缩进的普通段落。
  const noMarkerLvl = numId != null && ilvl != null
    ? ctx.numbering.get(numId)?.get(ilvl)?.noMarker === true
    : false;
  const inList = numId != null && ilvl != null && !noMarkerLvl;
  // noMarker 列表段落保留其缩进层级（按 ilvl），否则就会在正文里丢失嵌套观感。
  const indent = noMarkerLvl
    ? (ilvl ?? 0)
    : directIndent || (!inList ? styleInfo?.indent ?? 0 : 0);
  const lineSpacing = parseLineSpacing(pPr);

  let inline = parseInlineContent(pEl, ctx, styleInfo?.rPr ?? {});
  const headingLevel = styleInfo?.headingLevel ?? null;
  const blockId = ctx.nextBlockId();

  // 标题若带自动编号，把编号文本前置显示（schema 的 list_item 不能包含 heading）
  if (headingLevel && numId != null && ilvl != null) {
    const lvlInfo = ctx.numbering.get(numId)?.get(ilvl);
    const numberText = buildHeadingNumberText(ctx, numId, ilvl, lvlInfo);
    if (numberText) {
      inline = [{ type: 'text', text: `${numberText} ` }, ...inline];
    }
  }

  const node = headingLevel
    ? {
        type: 'heading',
        attrs: {
          level: headingLevel,
          blockId,
          styleName: styleInfo?.name ?? null,
          align,
          indent,
          lineSpacing,
        },
        content: inline,
      }
    : {
        type: 'paragraph',
        attrs: {
          blockId,
          styleName: styleInfo?.name ?? null,
          align,
          indent,
          lineSpacing,
        },
        content: inline,
      };

  if (numId != null && ilvl != null && !headingLevel && !noMarkerLvl) {
    const info = ctx.numbering.get(numId)?.get(ilvl) ?? {
      ordered: false,
      numberFormat: 'decimal',
      bulletStyle: 'disc',
      lvlText: null,
      noMarker: false,
      indLeftTw: 0,
      indHangingTw: 0,
      font: null,
      start: 1,
    };
    return {
      kind: 'listItem',
      level: ilvl,
      ordered: info.ordered,
      numberFormat: info.numberFormat,
      bulletStyle: info.bulletStyle,
      numId,
      lvlText: info.lvlText,
      start: info.start,
      indLeftTw: info.indLeftTw,
      indHangingTw: info.indHangingTw,
      font: info.font,
      seqGroup: isCaptionLevel(info.lvlText),
      node,
    };
  }
  return { kind: 'block', node };
}

function parseTableEl(tblEl: Element, ctx: ParseCtx): any {
  const rows = children(tblEl, 'tr');

  // ---- 列宽模型（colwidth）----
  // 读取 <w:tblGrid> 的 gridCol 宽度（twips）。仅当网格列数与实际列数一致、且宽度像
  // "真实"列宽（≥ MIN_REAL_TW，排除全为占位假宽度）时，才把列宽写进单元格 colwidth：
  // prosemirror-tables 因此按真实列宽固定渲染（table.style.width = 真实总宽），而不是
  // 让每列落到 defaultCellMinWidth（N×100px）兜底。
  // 否则视为 Word 自动宽度（autofit）表格（如假 grid + 无 tcW/tblW）：不写 colwidth，
  // 由 CSS width:100% 铺满编辑区，配合较小的 defaultCellMinWidth（见 pluginsSetup），
  // 不再被强制撑到 N×100px 的最小宽度。
  const tblGrid = child(tblEl, 'tblGrid');
  const gridWidthsTw = (tblGrid ? children(tblGrid, 'gridCol') : [])
    .map((g) => wAttrInt(g, 'w'))
    .filter((w): w is number => w != null && w > 0);
  const gridSize = gridWidthsTw.length;

  // 实际列数 = 各行列槽数（Σ gridSpan）的最大值（处理 grid 被压缩/与 cell 不符的情况）
  let actualCols = 0;
  for (const tr of rows) {
    let slots = 0;
    for (const tc of children(tr, 'tc')) {
      const gs = wAttrInt(child(child(tc, 'tcPr'), 'gridSpan'), 'val');
      slots += gs && gs > 0 ? gs : 1;
    }
    actualCols = Math.max(actualCols, slots);
  }

  // 真实列宽判定：列数与实际一致，且至少一格宽度 ≥ 360twips（≈24px），排除占位假宽度
  const MIN_REAL_TW = 360;
  const colWidthsTw: number[] | null =
    gridSize === actualCols &&
    gridSize > 0 &&
    gridWidthsTw.some((w) => w >= MIN_REAL_TW)
      ? gridWidthsTw
      : null;

  const rowNodes = rows.map((tr, rowIndex) => {
    const cellEls = children(tr, 'tc');
    const cellNodes: any[] = [];
    let colStart = 0;
    for (const tc of cellEls) {
      const tcPr = child(tc, 'tcPr');
      const shd = child(tcPr, 'shd');
      const fill = shd?.getAttribute('w:fill');
      const background =
        fill && fill !== 'auto' && /^[0-9a-fA-F]{6}$/.test(fill)
          ? `#${fill}`
          : null;
      const gridSpanVal = wAttr(child(tcPr, 'gridSpan'), 'val');
      const colspan = gridSpanVal ? Number(gridSpanVal) : undefined;
      const span = colspan ?? 1;

      // colwidth = 该单元格覆盖的各列宽度（px）；autofit 表格不写、维持 CSS 自适应
      let colwidth: number[] | undefined;
      if (colWidthsTw) {
        colwidth = colWidthsTw
          .slice(colStart, colStart + span)
          .map((tw) => Math.round(tw / TW_PER_PX));
      }
      colStart += span;

      const flat: FlatBlock[] = [];
      for (const c of Array.from(tc.children)) {
        if (c.localName === 'p') flat.push(parseParagraphEl(c, ctx));
        else if (c.localName === 'tbl')
          flat.push({ kind: 'block', node: parseTableEl(c, ctx) });
      }
      // 单元格内容暂不折叠：表格整体作为一个"待展开"的原始块交给 foldListBlocks，
      // 让表格内外的列表在真正文档顺序里共享同一个编号计数器（否则表格在 walk 时
      // 提前消耗计数器，正文在其后的列表编号会错位）。
      if (!flat.length)
        flat.push({
          kind: 'block',
          node: {
            type: 'paragraph',
            attrs: { blockId: ctx.nextBlockId() },
            content: [],
          },
        });
      const isHeader = rowIndex === 0;

      const attrs: Record<string, any> = {
        cellId: ctx.nextCellId(),
        background,
        colspan,
      };
      if (colwidth) attrs.colwidth = colwidth;

      cellNodes.push({
        type: isHeader ? 'table_header' : 'table_cell',
        attrs,
        colStart,
        content: flat,
      });
    }
    return { type: 'table_row', content: cellNodes };
  });
  // 标记为"原始表格"，由 foldListBlocks 就地按文档顺序展开其单元格内容
  return { type: 'table', __rawTable: true, content: rowNodes };
}

function walkBodyChildren(
  containerEl: Element,
  ctx: ParseCtx,
  flat: FlatBlock[]
) {
  for (const el of Array.from(containerEl.children)) {
    if (el.localName === 'p') {
      flat.push(parseParagraphEl(el, ctx));
      // 段落级 sectPr（w:pPr 里的分节符）标记某一节的结束
      const sect = child(child(el, 'pPr'), 'sectPr');
      if (sect) ctx.sectBreaks.push({ sectPr: sect, endBlock: flat.length });
    } else if (el.localName === 'tbl') {
      flat.push({ kind: 'block', node: parseTableEl(el, ctx) });
    } else if (el.localName === 'sdt') {
      const c = child(el, 'sdtContent');
      if (c) walkBodyChildren(c, ctx, flat);
    }
    // sectPr / bookmarkStart 等结构性/元信息元素：忽略
  }
}

// ---------------------------------------------------------------------------
// 页眉页脚 + 页面几何
// ---------------------------------------------------------------------------

/** 解析单个节的页面几何（纸张大小 + 页边距 + 页眉/页脚距离） */
function parseSectPageSetup(sectPr: Element | null): DocxPageSetup {
  const num = (el: Element | null, attr: string, def: number) => {
    if (!el) return def;
    const v = Number(el.getAttribute(`w:${attr}`));
    return Number.isFinite(v) && v > 0 ? v : def;
  };
  const pgSz = sectPr ? child(sectPr, 'pgSz') : null;
  const pgMar = sectPr ? child(sectPr, 'pgMar') : null;
  return {
    pageWidthTw: num(pgSz, 'w', 11906),
    pageHeightTw: num(pgSz, 'h', 16838),
    marginTopTw: num(pgMar, 'top', 1417),
    marginBottomTw: num(pgMar, 'bottom', 1134),
    marginLeftTw: num(pgMar, 'left', 1417),
    marginRightTw: num(pgMar, 'right', 1417),
    headerDistTw: num(pgMar, 'header', 708),
    footerDistTw: num(pgMar, 'footer', 708),
  };
}

/** 解析正文末尾 sectPr 中的页面几何（旧接口，仅读末节） */
function parsePageSetup(body: Element): DocxPageSetup {
  return parseSectPageSetup(child(body, 'sectPr'));
}

/**
 * 把页眉/页脚域指令（instr / fldSimple）归为 PAGE / NUMPAGES / SECTIONPAGES。
 * 注意顺序：NUMPAGES、SECTIONPAGES 都要先于 PAGE 判断（前缀不含交叉，但保持明确）。
 */
function classifyHfField(
  instr: string
): 'PAGE' | 'NUMPAGES' | 'SECTIONPAGES' | null {
  if (instr.startsWith('NUMPAGES')) return 'NUMPAGES';
  if (instr.startsWith('SECTIONPAGES')) return 'SECTIONPAGES';
  if (instr.startsWith('PAGE')) return 'PAGE';
  return null;
}

/**
 * 处理页眉/页脚里一个 <w:r> 运行，跟踪 PAGE / NUMPAGES 域并把域结果替换成占位文本。
 * 标准域结构：begin → instrText(" PAGE ") → separate → 缓存结果(t) → end。
 * instr 在 begin 处清空；缓存结果那一 run 的 t 会被换成占位符。
 * 若该域从未被更新过（没有缓存 <w:t>），则在 end 处补一个占位符，
 * 保证渲染端始终能按当前页码/总页数输出数字。
 */
function processHfRun(
  rEl: Element,
  inheritedRPr: RunProps,
  fieldState: { instr: string | null; emitted: boolean },
  out: any[]
): void {
  const ownProps = extractRunProps(child(rEl, 'rPr'));
  const merged: RunProps = { ...inheritedRPr, ...ownProps };
  const marks = buildMarksFromRunProps(merged, []);
  let sawFldSimple = false;
  // 第一遍：扫描域指令（fldChar / instrText / fldSimple），更新 fieldState
  for (const c of Array.from(rEl.children)) {
    if (c.localName === 'fldChar') {
      const t = c.getAttribute('w:fldCharType');
      if (t === 'begin') {
        fieldState.instr = null;
        fieldState.emitted = false;
      } else if (t === 'end') {
        // 该域没有任何缓存值输出过一个数字 → 在这里补占位符
        if (fieldState.instr && !fieldState.emitted) {
          out.push({
            type: 'text',
            text:
              fieldState.instr === 'PAGE'
                ? FIELD_PAGE
                : fieldState.instr === 'SECTIONPAGES'
                  ? FIELD_SECTIONPAGES
                  : FIELD_NUMPAGES,
            marks: marks.length ? marks : undefined,
          });
        }
        fieldState.instr = null;
        fieldState.emitted = false;
      }
      // separate：保留 instr，缓存值随后输出
    } else if (c.localName === 'instrText') {
      const instr = (c.textContent ?? '').trim().toUpperCase();
      fieldState.instr = classifyHfField(instr);
      fieldState.emitted = false;
    } else if (c.localName === 'fldSimple') {
      sawFldSimple = true;
      const instr = (c.getAttribute('w:instr') ?? '').trim().toUpperCase();
      fieldState.instr = classifyHfField(instr);
    }
  }
  // 第二遍：产出文本（普通文本 / 域缓存值 / fldSimple 自闭合占位）
  let hasText = false;
  for (const c of Array.from(rEl.children)) {
    if (c.localName === 't' || c.localName === 'tab') {
      hasText = true;
      let text = c.localName === 'tab' ? '\t' : (c.textContent ?? '');
      if (fieldState.instr === 'PAGE') {
        text = FIELD_PAGE;
        fieldState.emitted = true;
      } else if (fieldState.instr === 'NUMPAGES') {
        text = FIELD_NUMPAGES;
        fieldState.emitted = true;
      } else if (fieldState.instr === 'SECTIONPAGES') {
        text = FIELD_SECTIONPAGES;
        fieldState.emitted = true;
      }
      if (text)
        out.push({
          type: 'text',
          text,
          marks: marks.length ? marks : undefined,
        });
    }
  }
  for (const c of Array.from(rEl.children)) {
    if (c.localName === 'fldSimple' && !hasText && fieldState.instr) {
      fieldState.emitted = true;
      out.push({
        type: 'text',
        text:
          fieldState.instr === 'PAGE' ? FIELD_PAGE : FIELD_NUMPAGES,
        marks: marks.length ? marks : undefined,
      });
    }
  }
  // fldSimple 是自包含的一体域，处理完后复位 fieldState，避免污染其后的普通文本
  if (sawFldSimple) {
    fieldState.instr = null;
    fieldState.emitted = false;
  }
}

/** 解析页眉/页脚段落的行级内容（复用 run 样式，支持 PAGE/NUMPAGES 域） */
function parseHfInline(
  containerEl: Element,
  inheritedRPr: RunProps
): any[] {
  const out: any[] = [];
  const fieldState: { instr: string | null; emitted: boolean } = {
    instr: null,
    emitted: false,
  };
  const walk = (el: Element) => {
    for (const node of Array.from(el.children)) {
      if (node.localName === 'r') {
        processHfRun(node, inheritedRPr, fieldState, out);
      } else if (node.localName === 'hyperlink') {
        for (const inner of Array.from(node.children)) {
          if (inner.localName === 'r')
            processHfRun(inner, inheritedRPr, fieldState, out);
        }
      } else if (node.localName === 'fldSimple') {
        const instr = (node.getAttribute('w:instr') ?? '').trim().toUpperCase();
        const saved = fieldState.instr;
        fieldState.instr = instr.startsWith('NUMPAGES')
          ? 'NUMPAGES'
          : instr.startsWith('PAGE')
            ? 'PAGE'
            : null;
        const before = out.length;
        for (const inner of Array.from(node.children)) {
          if (inner.localName === 'r')
            processHfRun(inner, inheritedRPr, fieldState, out);
        }
        // 域内没有产出任何文本（例如 docx.js SimpleField 无缓存值时），仍插入占位
        if (out.length === before && fieldState.instr) {
          out.push({
            type: 'text',
            text:
              fieldState.instr === 'PAGE' ? FIELD_PAGE : FIELD_NUMPAGES,
          });
        }
        fieldState.instr = saved;
      } else if (node.localName === 'ins') {
        walk(node);
      } else if (node.localName === 'del') {
        /* 修订-删除：忽略 */
      } else if (node.localName === 'sdt') {
        const c = child(node, 'sdtContent');
        if (c) walk(c);
      }
    }
  };
  walk(containerEl);
  return out;
}

/** 解析页眉/页脚里的一个段落 */
function parseHfParagraph(pEl: Element, ctx: ParseCtx): any {
  const pPr = child(pEl, 'pPr');
  const styleId = wAttr(child(pPr, 'pStyle'), 'val');
  const styleInfo = styleId ? (ctx.styles.get(styleId) ?? null) : null;
  const align =
    normalizeAlign(wAttr(child(pPr, 'jc'), 'val')) ??
    styleInfo?.align ??
    null;
  const indent = parseIndentLevel(pPr) || (styleInfo?.indent ?? 0);
  const lineSpacing = parseLineSpacing(pPr);
  const inline = parseHfInline(pEl, styleInfo?.rPr ?? {});
  return {
    type: 'paragraph',
    attrs: {
      blockId: ctx.nextBlockId(),
      styleName: styleInfo?.name ?? null,
      align,
      indent,
      lineSpacing,
    },
    content: inline,
  };
}

/** 判断一个段落是否只是"文本框"的载体（其真实内容在 wps:txbx / v:textbox 内层） */
function containsTextbox(p: Element): boolean {
  const nodes = Array.from(p.getElementsByTagName('*'));
  return nodes.some(
    (n) => n.localName === 'txbx' || n.localName === 'textbox'
  );
}

/**
 * 递归收集页眉/页脚里的块级元素（<w:p> / <w:tbl>），保持文档顺序。
 * 真实文档（尤其 WPS）常把页眉/页脚内容（如页码）包在文本框里：
 *   <w:p><w:r><mc:AlternateContent><w:drawing><wp:anchor>…<wps:txbx><w:txbxContent><w:p>…</w:p>
 * 这类"包装段落"只承载文本框、自身没有文本，应展开其内层段落而不是作为空段落输出。
 */
function collectHfBlocks(el: Element, out: Element[]): void {
  for (const c of Array.from(el.children)) {
    if (c.localName === 'p') {
      if (containsTextbox(c)) collectHfBlocks(c, out);
      else out.push(c);
    } else if (c.localName === 'tbl') {
      out.push(c);
    } else if (c.localName === 'AlternateContent') {
      // WPS/Word 用 AlternateContent 包装文本框：Choice 是首选渲染，Fallback 是兜底。
      // 两者内容相同，只取 Choice，避免页码等重复显示。
      for (const choice of Array.from(c.children)) {
        if (choice.localName === 'Choice') collectHfBlocks(choice, out);
      }
    } else {
      collectHfBlocks(c, out);
    }
  }
}

/** 解析页眉/页脚根元素（<w:hdr>/<w:ftr>）为 PM JSON 文档 */
function parseHfRoot(rootEl: Element, ctx: ParseCtx): any {
  const blocks: Element[] = [];
  collectHfBlocks(rootEl, blocks);
  const content: any[] = [];
  for (const el of blocks) {
    if (el.localName === 'p') content.push(parseHfParagraph(el, ctx));
    else if (el.localName === 'tbl') content.push(parseTableEl(el, ctx));
  }
  return content.length ? { type: 'doc', content } : null;
}

/** 关系 target（相对 word/）→ 完整 zip 内路径 */
function resolveZipPath(target: string): string {
  const t = target.replace(/^\//, '');
  if (t.startsWith('word/')) return t;
  return `word/${t}`;
}

/** 读取单个节 sectPr 引用的默认页眉/页脚，解析为 PM JSON 文档；无默认引用则返回 null */
async function loadHfFromSect(
  zip: JSZip,
  parser: DOMParser,
  rels: Map<string, string>,
  sectPr: Element | null,
  kind: 'headerReference' | 'footerReference',
  warnings: string[]
): Promise<any | null> {
  if (!sectPr) return null;
  const ref = children(sectPr, kind).find(
    (e) => (wAttr(e, 'type') || 'default') === 'default'
  );
  if (!ref) return null;
  const target = rels.get(ref.getAttribute('r:id') ?? '');
  if (!target) return null;
  const str = await readZipText(zip, resolveZipPath(target));
  if (!str) return null;
  const xml = parseXml(parser, str, target, warnings);
  if (!xml || !xml.documentElement) return null;
  const ctx: ParseCtx = {
    styles: new Map(),
    numbering: new Map(),
    numberingCounters: new Map(),
    rels,
    media: new Map(),
    warnings,
    nextBlockId: (() => {
      let n = 0;
      return () => `hb${++n}`;
    })(),
    nextCellId: (() => {
      let n = 0;
      return () => `hc${++n}`;
    })(),
    sectBreaks: [],
  };
  return parseHfRoot(xml.documentElement, ctx);
}

/** 读取正文 sectPr 引用的默认页眉/页脚，解析为 PM JSON 文档（旧接口，仅取单一默认节） */
async function readHeaderFooter(
  zip: JSZip,
  parser: DOMParser,
  rels: Map<string, string>,
  body: Element,
  kind: 'headerReference' | 'footerReference',
  warnings: string[]
): Promise<any | null> {
  // 收集所有节的 sectPr：正文级（末节）+ 段落级（分节符处的各非末节）。
  // 真实 Word 文档里中间各节的页眉/页脚定义在 <w:pPr><w:sectPr> 里，
  // 只读正文级会漏掉它们。优先取正文级（末节）默认引用，否则取第一个带默认引用的节。
  const sects: Element[] = [];
  const bodySect = child(body, 'sectPr');
  if (bodySect) sects.push(bodySect);
  for (const p of Array.from(body.getElementsByTagNameNS(W, 'p'))) {
    const s = child(child(p, 'pPr'), 'sectPr');
    if (s && !sects.includes(s)) sects.push(s);
  }
  for (const s of sects) {
    const r = children(s, kind).find(
      (e) => (wAttr(e, 'type') || 'default') === 'default'
    );
    if (r) return loadHfFromSect(zip, parser, rels, s, kind, warnings);
  }
  return null;
}

/**
 * 组装逐节页眉/页脚。一节 = 一块连续的正文区间。
 * - 段落级 <w:pPr><w:sectPr> 标记某一节的结束（该节区间 = [上一节结束, 本节结束)），
 *   且该 sectPr 定义这一节自己的页眉/页脚。
 * - 正文级 <w:sectPr>（末节）定义最后一节。
 * - "链接到前一节"：某节没有属于自己的默认引用时继承前一节。
 */
async function buildSections(
  ctx: ParseCtx,
  zip: JSZip,
  parser: DOMParser,
  rels: Map<string, string>,
  body: Element,
  blockCount: number,
  warnings: string[]
): Promise<DocxSection[]> {
  const bodySect = child(body, 'sectPr');
  const ranges: { sectPr: Element | null; start: number; end: number }[] = [];
  let start = 0;
  for (const br of ctx.sectBreaks) {
    ranges.push({ sectPr: br.sectPr, start, end: br.endBlock });
    start = br.endBlock;
  }
  ranges.push({ sectPr: bodySect, start, end: blockCount });

  const out: DocxSection[] = [];
  let prevHeader: any = null;
  let prevFooter: any = null;
  for (const r of ranges) {
    const header =
      (await loadHfFromSect(zip, parser, rels, r.sectPr, 'headerReference', warnings)) ??
      prevHeader;
    const footer =
      (await loadHfFromSect(zip, parser, rels, r.sectPr, 'footerReference', warnings)) ??
      prevFooter;
    prevHeader = header;
    prevFooter = footer;
    const { start, format } = parseSectionPageNum(r.sectPr);
    out.push({
      startBlock: r.start,
      endBlock: r.end,
      header,
      footer,
      pageStart: start,
      pageFormat: format,
      pageSetup: parseSectPageSetup(r.sectPr),
    });
  }
  return out;
}

/** 读取节的页码设置（w:pgNumType）：w:start 页码起始值、w:format 页码格式 */
function parseSectionPageNum(
  sectPr: Element | null
): { start: number | null; format: string } {
  if (!sectPr) return { start: null, format: 'decimal' };
  const pgn = child(sectPr, 'pgNumType');
  if (!pgn) return { start: null, format: 'decimal' };
  const startRaw = wAttr(pgn, 'start');
  const start =
    startRaw !== null && startRaw !== '' ? Number(startRaw) : null;
  const format = wAttr(pgn, 'format') || 'decimal';
  return {
    start: Number.isFinite(start ?? NaN) ? start : null,
    format,
  };
}

// ---------------------------------------------------------------------------
// 入口
// ---------------------------------------------------------------------------

export async function parseDocxFile(
  arrayBuffer: ArrayBuffer
): Promise<ParsedDocx> {
  const warnings: string[] = [];
  const zip = await JSZip.loadAsync(arrayBuffer);

  const documentXmlStr = await readZipText(zip, 'word/document.xml');
  if (!documentXmlStr) {
    throw new Error('不是有效的 DOCX 文件（缺少 word/document.xml）');
  }

  const parser = new DOMParser();
  const documentXml = parseXml(
    parser,
    documentXmlStr,
    'document.xml',
    warnings
  );
  if (!documentXml) throw new Error('document.xml 解析失败，文件可能已损坏');

  const stylesXmlStr = await readZipText(zip, 'word/styles.xml');
  const stylesXml = stylesXmlStr
    ? parseXml(parser, stylesXmlStr, 'styles.xml', warnings)
    : null;
  const styles = stylesXml
    ? parseStylesXml(stylesXml)
    : new Map<string, StyleInfo>();

  const numberingXmlStr = await readZipText(zip, 'word/numbering.xml');
  const numberingXml = numberingXmlStr
    ? parseXml(parser, numberingXmlStr, 'numbering.xml', warnings)
    : null;
  const numbering = numberingXml
    ? parseNumberingXml(numberingXml)
    : new Map<string, Map<number, NumberingLevelInfo>>();

  const relsXmlStr = await readZipText(zip, 'word/_rels/document.xml.rels');
  const relsXml = relsXmlStr
    ? parseXml(parser, relsXmlStr, 'document.xml.rels', warnings)
    : null;
  const rels = relsXml ? parseRelsXml(relsXml) : new Map<string, string>();

  const media = await readMedia(zip, rels);

  const commentsXmlStr = await readZipText(zip, 'word/comments.xml');
  const commentsXml = commentsXmlStr
    ? parseXml(parser, commentsXmlStr, 'comments.xml', warnings)
    : null;
  const comments = commentsXml ? parseCommentsXml(commentsXml) : [];

  let blockIdCounter = 1;
  let cellIdCounter = 1;
  const ctx: ParseCtx = {
    styles,
    numbering,
    numberingCounters: new Map<string, number[]>(),
    rels,
    media,
    warnings,
    nextBlockId: () => `b${blockIdCounter++}`,
    nextCellId: () => `c${cellIdCounter++}`,
    sectBreaks: [],
  };

  const body = documentXml.getElementsByTagNameNS(W, 'body')[0];
  if (!body) throw new Error('DOCX 缺少文档主体（word:body）');

  const pageSetup = parsePageSetup(body);
  const header = await readHeaderFooter(
    zip,
    parser,
    rels,
    body,
    'headerReference',
    warnings
  );
  const footer = await readHeaderFooter(
    zip,
    parser,
    rels,
    body,
    'footerReference',
    warnings
  );

  const flat: FlatBlock[] = [];
  walkBodyChildren(body, ctx, flat);
  const content = foldListBlocks(flat);

  const sections = await buildSections(
    ctx,
    zip,
    parser,
    rels,
    body,
    flat.length,
    warnings
  );

  const json = {
    type: 'doc',
    content: content.length
      ? content
      : [
          {
            type: 'paragraph',
            attrs: { blockId: ctx.nextBlockId() },
            content: [],
          },
        ],
  };

  return { json, comments, warnings, pageSetup, header, footer, sections };
}
