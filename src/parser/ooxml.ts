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
import JSZip from "jszip";

const W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const A_NS = "http://schemas.openxmlformats.org/drawingml/2006/main";

export interface DocxComment {
  id: number;
  author: string;
  date: string | null;
  text: string;
}

export interface ParsedDocx {
  json: any;
  comments: DocxComment[];
  warnings: string[];
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
  return el.getAttribute("w:" + name);
}

function normalizeAlign(val: string | null): string | null {
  if (!val) return null;
  if (val === "both") return "justify";
  if (val === "left" || val === "right" || val === "center" || val === "justify") return val;
  return null;
}

/** w:ind 的 w:left（缇，1/20 pt）→ 我们的缩进级别（每级 720 缇 ≈ Word "增加缩进" 一次）*/
function parseIndentLevel(pPr: Element | null): number {
  const indEl = child(pPr, "ind");
  const leftTwips = Number(indEl?.getAttribute("w:left") ?? indEl?.getAttribute("w:start") ?? "0");
  if (!leftTwips || Number.isNaN(leftTwips)) return 0;
  return Math.max(0, Math.min(8, Math.round(leftTwips / 720)));
}

/** w:spacing 的 w:line（仅 lineRule=auto 时是"倍数×240"）→ 行距倍数，如 1 / 1.5 / 2 */
function parseLineSpacing(pPr: Element | null): number | null {
  const spacingEl = child(pPr, "spacing");
  if (!spacingEl) return null;
  const lineRule = spacingEl.getAttribute("w:lineRule");
  const line = Number(spacingEl.getAttribute("w:line") ?? "0");
  if (!line || Number.isNaN(line)) return null;
  if (lineRule && lineRule !== "auto") return null; // exact/atLeast 是绝对尺寸，暂不处理
  const multiplier = Math.round((line / 240) * 100) / 100;
  return multiplier > 0 ? multiplier : null;
}

async function readZipText(zip: JSZip, path: string): Promise<string | null> {
  const entry = zip.file(path);
  if (!entry) return null;
  return entry.async("text");
}

function parseXml(parser: DOMParser, text: string, label: string, warnings: string[]): Document | null {
  const doc = parser.parseFromString(text, "application/xml");
  if (doc.getElementsByTagName("parsererror").length > 0) {
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
}

function extractRunProps(rPrEl: Element | null): RunProps {
  const props: RunProps = {};
  if (!rPrEl) return props;

  const bEl = child(rPrEl, "b");
  if (bEl && wAttr(bEl, "val") !== "false" && wAttr(bEl, "val") !== "0") props.bold = true;

  const iEl = child(rPrEl, "i");
  if (iEl && wAttr(iEl, "val") !== "false" && wAttr(iEl, "val") !== "0") props.italic = true;

  const uEl = child(rPrEl, "u");
  const uVal = wAttr(uEl, "val");
  if (uEl && uVal && uVal !== "none") props.underline = true;

  const strikeEl = child(rPrEl, "strike");
  if (strikeEl && wAttr(strikeEl, "val") !== "false" && wAttr(strikeEl, "val") !== "0") props.strike = true;

  const colorEl = child(rPrEl, "color");
  const colorVal = wAttr(colorEl, "val");
  if (colorVal && colorVal !== "auto" && /^[0-9a-fA-F]{6}$/.test(colorVal)) props.color = `#${colorVal}`;

  const szEl = child(rPrEl, "sz");
  const szVal = wAttr(szEl, "val");
  if (szVal && !Number.isNaN(Number(szVal))) props.sizeHalfPt = Number(szVal);

  const fontsEl = child(rPrEl, "rFonts");
  const fam = fontsEl?.getAttribute("w:ascii") ?? fontsEl?.getAttribute("w:eastAsia") ?? fontsEl?.getAttribute("w:hAnsi");
  if (fam) props.fontFamily = fam;

  const hlEl = child(rPrEl, "highlight");
  const hlVal = wAttr(hlEl, "val");
  if (hlVal && hlVal !== "none") props.highlight = hlVal;

  return props;
}

function parseStylesXml(doc: Document): Map<string, StyleInfo> {
  const raw = new Map<string, { name: string; basedOn: string | null; headingLevel: number | null; rPr: RunProps; align: string | null }>();

  for (const styleEl of Array.from(doc.getElementsByTagNameNS(W, "style"))) {
    const id = styleEl.getAttribute("w:styleId");
    const type = styleEl.getAttribute("w:type");
    if (!id || (type !== "paragraph" && type !== "character")) continue;

    const name = wAttr(child(styleEl, "name"), "val") ?? id;
    const basedOn = wAttr(child(styleEl, "basedOn"), "val");
    const headingMatch = /^heading\s*(\d)/i.exec(name);
    const align = normalizeAlign(wAttr(child(child(styleEl, "pPr"), "jc"), "val"));

    raw.set(id, {
      name,
      basedOn,
      headingLevel: headingMatch ? Math.min(9, Number(headingMatch[1])) : id === "Title" ? 1 : null,
      rPr: extractRunProps(child(styleEl, "rPr")),
      align,
    });
  }

  const resolved = new Map<string, StyleInfo>();
  const resolving = new Set<string>();

  function resolve(id: string): StyleInfo {
    const cached = resolved.get(id);
    if (cached) return cached;
    const info = raw.get(id);
    if (!info) {
      const fallback: StyleInfo = { name: id, basedOn: null, headingLevel: null, rPr: {}, align: null };
      resolved.set(id, fallback);
      return fallback;
    }
    if (resolving.has(id)) {
      // 循环继承兜底，避免死循环
      const flat: StyleInfo = { name: info.name, basedOn: null, headingLevel: info.headingLevel, rPr: info.rPr, align: info.align };
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
    };
    resolving.delete(id);
    resolved.set(id, merged);
    return merged;
  }

  for (const id of raw.keys()) resolve(id);
  return resolved;
}

// ---------------------------------------------------------------------------
// numbering.xml —— 列表格式（区分有序/无序）
// ---------------------------------------------------------------------------

function parseNumberingXml(doc: Document): Map<string, Map<number, boolean>> {
  const abstractFormats = new Map<string, Map<number, boolean>>();

  for (const abs of Array.from(doc.getElementsByTagNameNS(W, "abstractNum"))) {
    const absId = abs.getAttribute("w:abstractNumId");
    if (absId == null) continue;
    const levelMap = new Map<number, boolean>();
    for (const lvl of children(abs, "lvl")) {
      const ilvl = Number(lvl.getAttribute("w:ilvl") ?? "0");
      const fmt = wAttr(child(lvl, "numFmt"), "val") ?? "bullet";
      levelMap.set(ilvl, fmt !== "bullet" && fmt !== "none");
    }
    abstractFormats.set(absId, levelMap);
  }

  const numToAbstract = new Map<string, string>();
  for (const num of Array.from(doc.getElementsByTagNameNS(W, "num"))) {
    const numId = num.getAttribute("w:numId");
    const absId = wAttr(child(num, "abstractNumId"), "val");
    if (numId != null && absId != null) numToAbstract.set(numId, absId);
  }

  const result = new Map<string, Map<number, boolean>>();
  for (const [numId, absId] of numToAbstract) {
    const lvlMap = abstractFormats.get(absId);
    if (lvlMap) result.set(numId, lvlMap);
  }
  return result;
}

// ---------------------------------------------------------------------------
// 关系文件 + 媒体资源
// ---------------------------------------------------------------------------

function parseRelsXml(doc: Document): Map<string, string> {
  const out = new Map<string, string>();
  for (const el of Array.from(doc.getElementsByTagName("Relationship"))) {
    const id = el.getAttribute("Id");
    const target = el.getAttribute("Target");
    if (id && target) out.set(id, target);
  }
  return out;
}

async function readMedia(zip: JSZip, rels: Map<string, string>): Promise<Map<string, string>> {
  const out = new Map<string, string>();
  for (const [rId, target] of rels) {
    if (!/\.(png|jpe?g|gif|bmp)$/i.test(target)) continue;
    const path = target.startsWith("media/") ? `word/${target}` : target.replace(/^\.\.\//, "word/");
    const entry = zip.file(path);
    if (!entry) continue;
    const base64 = await entry.async("base64");
    const ext = (path.split(".").pop() ?? "png").toLowerCase();
    const mime = ext === "jpg" || ext === "jpeg" ? "image/jpeg" : ext === "gif" ? "image/gif" : ext === "bmp" ? "image/bmp" : "image/png";
    out.set(rId, `data:${mime};base64,${base64}`);
  }
  return out;
}

// ---------------------------------------------------------------------------
// comments.xml
// ---------------------------------------------------------------------------

function extractPlainText(el: Element): string {
  const parts: string[] = [];
  for (const t of Array.from(el.getElementsByTagNameNS(W, "t"))) parts.push(t.textContent ?? "");
  return parts.join("");
}

function parseCommentsXml(doc: Document): DocxComment[] {
  const out: DocxComment[] = [];
  for (const el of Array.from(doc.getElementsByTagNameNS(W, "comment"))) {
    const id = Number(el.getAttribute("w:id"));
    if (Number.isNaN(id)) continue;
    out.push({
      id,
      author: el.getAttribute("w:author") || "匿名",
      date: el.getAttribute("w:date"),
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
  numbering: Map<string, Map<number, boolean>>;
  rels: Map<string, string>;
  media: Map<string, string>;
  warnings: string[];
  nextBlockId: () => string;
  nextCellId: () => string;
}

type FlatBlock = { kind: "block"; node: any } | { kind: "listItem"; level: number; ordered: boolean; node: any };

function buildMarksFromRunProps(p: RunProps, activeComments: number[]): any[] {
  const marks: any[] = [];
  if (p.bold) marks.push({ type: "strong" });
  if (p.italic) marks.push({ type: "em" });
  if (p.underline) marks.push({ type: "underline" });
  if (p.strike) marks.push({ type: "strike" });
  if (p.color || p.fontFamily || p.sizeHalfPt || p.highlight) {
    marks.push({
      type: "docxStyle",
      attrs: {
        color: p.color ?? null,
        fontFamily: p.fontFamily ?? null,
        sizeHalfPt: p.sizeHalfPt ?? null,
        highlight: p.highlight ?? null,
      },
    });
  }
  for (const id of activeComments) marks.push({ type: "comment", attrs: { id } });
  return marks;
}

function parseDrawing(drawingEl: Element, ctx: ParseCtx): any | null {
  const blip = drawingEl.getElementsByTagNameNS(A_NS, "blip")[0];
  const rId = blip?.getAttribute("r:embed");
  if (!rId) return null;
  const dataUrl = ctx.media.get(rId);
  if (!dataUrl) {
    ctx.warnings.push("[warning] 存在未能解析的内嵌图片");
    return null;
  }
  return { type: "image", attrs: { src: dataUrl, alt: "" } };
}

function parseRun(rEl: Element, ctx: ParseCtx, inheritedRPr: RunProps, activeComments: number[]): any[] {
  const ownProps = extractRunProps(child(rEl, "rPr"));
  const merged: RunProps = { ...inheritedRPr, ...ownProps };
  const marks = buildMarksFromRunProps(merged, activeComments);

  const out: any[] = [];
  for (const node of Array.from(rEl.children)) {
    switch (node.localName) {
      case "t": {
        const text = node.textContent ?? "";
        if (text) out.push({ type: "text", text, marks: marks.length ? marks : undefined });
        break;
      }
      case "br":
        out.push({ type: "hard_break" });
        break;
      case "tab": {
        const text = "\t";
        out.push({ type: "text", text, marks: marks.length ? marks : undefined });
        break;
      }
      case "drawing": {
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

function parseInlineContent(containerEl: Element, ctx: ParseCtx, inheritedRPr: RunProps): any[] {
  const out: any[] = [];
  const activeComments: number[] = [];

  const walk = (el: Element) => {
    for (const node of Array.from(el.children)) {
      switch (node.localName) {
        case "r":
          out.push(...parseRun(node, ctx, inheritedRPr, activeComments));
          break;
        case "hyperlink": {
          const rId = node.getAttribute("r:id");
          const anchor = node.getAttribute("w:anchor");
          const href = rId ? ctx.rels.get(rId) ?? null : anchor ? `#${anchor}` : null;
          const before = out.length;
          for (const inner of Array.from(node.children)) {
            if (inner.localName === "r") out.push(...parseRun(inner, ctx, inheritedRPr, activeComments));
          }
          if (href) {
            for (let i = before; i < out.length; i++) {
              const n = out[i];
              if (n.type === "text") n.marks = [...(n.marks ?? []), { type: "link", attrs: { href, title: null } }];
            }
          }
          break;
        }
        case "commentRangeStart": {
          const id = Number(node.getAttribute("w:id"));
          if (!Number.isNaN(id)) activeComments.push(id);
          break;
        }
        case "commentRangeEnd": {
          const id = Number(node.getAttribute("w:id"));
          const idx = activeComments.indexOf(id);
          if (idx >= 0) activeComments.splice(idx, 1);
          break;
        }
        case "ins":
          walk(node); // 修订-插入：展开为正常内容
          break;
        case "del":
          break; // 修订-删除：不进入最终结构（结构优先原则下的合理简化）
        case "sdt": {
          const c = child(node, "sdtContent");
          if (c) walk(c);
          break;
        }
        case "smartTag":
          walk(node);
          break;
        default:
          break;
      }
    }
  };

  walk(containerEl);
  return out;
}

function foldListBlocks(flat: FlatBlock[]): any[] {
  const out: any[] = [];
  const stack: { level: number; ordered: boolean; items: any[] }[] = [];

  const closeTop = () => {
    const top = stack.pop();
    if (!top) return;
    const listNode = { type: top.ordered ? "ordered_list" : "bullet_list", content: top.items };
    if (stack.length) {
      const parentItems = stack[stack.length - 1].items;
      parentItems[parentItems.length - 1].content.push(listNode);
    } else {
      out.push(listNode);
    }
  };

  for (const b of flat) {
    if (b.kind !== "listItem") {
      while (stack.length) closeTop();
      out.push(b.node);
      continue;
    }
    while (stack.length && b.level < stack[stack.length - 1].level) closeTop();
    if (!stack.length || b.level > stack[stack.length - 1].level) {
      stack.push({ level: b.level, ordered: b.ordered, items: [] });
    } else if (stack[stack.length - 1].ordered !== b.ordered) {
      closeTop();
      stack.push({ level: b.level, ordered: b.ordered, items: [] });
    }
    stack[stack.length - 1].items.push({ type: "list_item", content: [b.node] });
  }
  while (stack.length) closeTop();
  return out;
}

function parseParagraphEl(pEl: Element, ctx: ParseCtx): FlatBlock {
  const pPr = child(pEl, "pPr");
  const styleId = wAttr(child(pPr, "pStyle"), "val");
  const styleInfo = styleId ? ctx.styles.get(styleId) ?? null : null;

  const numPr = child(pPr, "numPr");
  const ilvl = numPr ? Number(wAttr(child(numPr, "ilvl"), "val") ?? "0") : null;
  const numId = numPr ? wAttr(child(numPr, "numId"), "val") : null;

  const directAlign = normalizeAlign(wAttr(child(pPr, "jc"), "val"));
  const align = directAlign ?? styleInfo?.align ?? null;
  const indent = parseIndentLevel(pPr);
  const lineSpacing = parseLineSpacing(pPr);

  const inline = parseInlineContent(pEl, ctx, styleInfo?.rPr ?? {});
  const headingLevel = styleInfo?.headingLevel ?? null;
  const blockId = ctx.nextBlockId();

  const node = headingLevel
    ? { type: "heading", attrs: { level: headingLevel, blockId, styleName: styleInfo?.name ?? null, align, indent, lineSpacing }, content: inline }
    : { type: "paragraph", attrs: { blockId, styleName: styleInfo?.name ?? null, align, indent, lineSpacing }, content: inline };

  if (numId != null && ilvl != null && !headingLevel) {
    const ordered = ctx.numbering.get(numId)?.get(ilvl) ?? false;
    return { kind: "listItem", level: ilvl, ordered, node };
  }
  return { kind: "block", node };
}

function parseTableEl(tblEl: Element, ctx: ParseCtx): any {
  const rows = children(tblEl, "tr");
  const rowNodes = rows.map((tr, rowIndex) => {
    const cells = children(tr, "tc");
    const cellNodes = cells.map((tc) => {
      const tcPr = child(tc, "tcPr");
      const shd = child(tcPr, "shd");
      const fill = shd?.getAttribute("w:fill");
      const background = fill && fill !== "auto" && /^[0-9a-fA-F]{6}$/.test(fill) ? `#${fill}` : null;
      const gridSpanVal = wAttr(child(tcPr, "gridSpan"), "val");
      const colspan = gridSpanVal ? Number(gridSpanVal) : undefined;

      const flat: FlatBlock[] = [];
      for (const c of Array.from(tc.children)) {
        if (c.localName === "p") flat.push(parseParagraphEl(c, ctx));
        else if (c.localName === "tbl") flat.push({ kind: "block", node: parseTableEl(c, ctx) });
      }
      const content = foldListBlocks(flat);
      const isHeader = rowIndex === 0;

      return {
        type: isHeader ? "table_header" : "table_cell",
        attrs: { cellId: ctx.nextCellId(), background, colspan },
        content: content.length ? content : [{ type: "paragraph", attrs: { blockId: ctx.nextBlockId() }, content: [] }],
      };
    });
    return { type: "table_row", content: cellNodes };
  });
  return { type: "table", content: rowNodes };
}

function walkBodyChildren(containerEl: Element, ctx: ParseCtx, flat: FlatBlock[]) {
  for (const el of Array.from(containerEl.children)) {
    if (el.localName === "p") {
      flat.push(parseParagraphEl(el, ctx));
    } else if (el.localName === "tbl") {
      flat.push({ kind: "block", node: parseTableEl(el, ctx) });
    } else if (el.localName === "sdt") {
      const c = child(el, "sdtContent");
      if (c) walkBodyChildren(c, ctx, flat);
    }
    // sectPr / bookmarkStart 等结构性/元信息元素：忽略
  }
}

// ---------------------------------------------------------------------------
// 入口
// ---------------------------------------------------------------------------

export async function parseDocxFile(arrayBuffer: ArrayBuffer): Promise<ParsedDocx> {
  const warnings: string[] = [];
  const zip = await JSZip.loadAsync(arrayBuffer);

  const documentXmlStr = await readZipText(zip, "word/document.xml");
  if (!documentXmlStr) {
    throw new Error("不是有效的 DOCX 文件（缺少 word/document.xml）");
  }

  const parser = new DOMParser();
  const documentXml = parseXml(parser, documentXmlStr, "document.xml", warnings);
  if (!documentXml) throw new Error("document.xml 解析失败，文件可能已损坏");

  const stylesXmlStr = await readZipText(zip, "word/styles.xml");
  const stylesXml = stylesXmlStr ? parseXml(parser, stylesXmlStr, "styles.xml", warnings) : null;
  const styles = stylesXml ? parseStylesXml(stylesXml) : new Map<string, StyleInfo>();

  const numberingXmlStr = await readZipText(zip, "word/numbering.xml");
  const numberingXml = numberingXmlStr ? parseXml(parser, numberingXmlStr, "numbering.xml", warnings) : null;
  const numbering = numberingXml ? parseNumberingXml(numberingXml) : new Map<string, Map<number, boolean>>();

  const relsXmlStr = await readZipText(zip, "word/_rels/document.xml.rels");
  const relsXml = relsXmlStr ? parseXml(parser, relsXmlStr, "document.xml.rels", warnings) : null;
  const rels = relsXml ? parseRelsXml(relsXml) : new Map<string, string>();

  const media = await readMedia(zip, rels);

  const commentsXmlStr = await readZipText(zip, "word/comments.xml");
  const commentsXml = commentsXmlStr ? parseXml(parser, commentsXmlStr, "comments.xml", warnings) : null;
  const comments = commentsXml ? parseCommentsXml(commentsXml) : [];

  let blockIdCounter = 1;
  let cellIdCounter = 1;
  const ctx: ParseCtx = {
    styles,
    numbering,
    rels,
    media,
    warnings,
    nextBlockId: () => `b${blockIdCounter++}`,
    nextCellId: () => `c${cellIdCounter++}`,
  };

  const body = documentXml.getElementsByTagNameNS(W, "body")[0];
  if (!body) throw new Error("DOCX 缺少文档主体（word:body）");

  const flat: FlatBlock[] = [];
  walkBodyChildren(body, ctx, flat);
  const content = foldListBlocks(flat);

  const json = {
    type: "doc",
    content: content.length ? content : [{ type: "paragraph", attrs: { blockId: ctx.nextBlockId() }, content: [] }],
  };

  return { json, comments, warnings };
}
