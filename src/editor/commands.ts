import type { Schema, Node as PMNode } from "prosemirror-model";
import type { EditorState, Transaction } from "prosemirror-state";
import { toggleMark, setBlockType, wrapIn } from "prosemirror-commands";
import { wrapInList, sinkListItem, liftListItem } from "prosemirror-schema-list";
import {
  addColumnAfter,
  addColumnBefore,
  addRowAfter,
  addRowBefore,
  deleteColumn,
  deleteRow,
  deleteTable,
  mergeCells,
  splitCell,
  toggleHeaderRow,
  toggleHeaderColumn,
  setCellAttr,
  isInTable,
} from "prosemirror-tables";
import { insertTable } from "./tableUtils";

type Cmd = (state: EditorState, dispatch?: (tr: Transaction) => void) => boolean;

export function markActive(state: EditorState, markType: any): boolean {
  const { from, $from, to, empty } = state.selection;
  if (empty) return !!markType.isInSet(state.storedMarks || $from.marks());
  return state.doc.rangeHasMark(from, to, markType);
}

export function blockActive(state: EditorState, nodeType: any, attrs: Record<string, any> = {}): boolean {
  const { $from, to, node } = state.selection as any;
  if (node) return node.hasMarkup(nodeType, attrs);
  return to <= $from.end() && $from.parent.hasMarkup(nodeType, attrs);
}

/** 读取选区起点所在段落/标题当前的属性值（用于让工具栏控件回显当前状态）*/
export function currentBlockAttr(state: EditorState, key: string): any {
  const node = state.selection.$from.parent;
  return node?.attrs?.[key];
}

/** 读取选区起点文字上 docxStyle mark 的某个属性（用于字体/字号/颜色控件回显）*/
export function currentDocxStyleAttr(state: EditorState, markType: any, key: string): any {
  const marks = state.storedMarks || state.selection.$from.marks();
  const mark = markType.isInSet(marks);
  return mark?.attrs?.[key] ?? null;
}

export function insertImage(schema: Schema): Cmd {
  return (state, dispatch) => {
    const src = window.prompt("图片地址（URL 或粘贴 base64）：");
    if (!src) return false;
    if (dispatch) {
      const node = schema.nodes.image.create({ src, alt: "" });
      dispatch(state.tr.replaceSelectionWith(node).scrollIntoView());
    }
    return true;
  };
}

export function insertLink(schema: Schema): Cmd {
  return (state, dispatch) => {
    const { from, to, empty } = state.selection;
    if (empty) {
      window.alert("请先选中要添加链接的文字");
      return false;
    }
    const href = window.prompt("链接地址：", "https://");
    if (!href) return false;
    if (dispatch) {
      dispatch(state.tr.addMark(from, to, schema.marks.link.create({ href })));
    }
    return true;
  };
}

/** 通用：把选区覆盖到的每个 paragraph/heading 节点的属性做一次 patch（setNodeMarkup 不改变节点大小，可在同一个事务里安全地对多个位置连续调用）*/
function updateBlockAttrs(getPatch: (node: PMNode) => Record<string, any> | null): Cmd {
  return (state, dispatch) => {
    const { from, to } = state.selection;
    let tr = state.tr;
    let changed = false;
    state.doc.nodesBetween(from, to, (node, pos) => {
      if (node.type.name !== "paragraph" && node.type.name !== "heading") return;
      const patch = getPatch(node);
      if (!patch) return;
      tr = tr.setNodeMarkup(pos, undefined, { ...node.attrs, ...patch });
      changed = true;
    });
    if (changed && dispatch) dispatch(tr);
    return changed;
  };
}

export function setAlign(align: string | null): Cmd {
  return updateBlockAttrs(() => ({ align }));
}

export function setLineSpacing(value: number | null): Cmd {
  return updateBlockAttrs(() => ({ lineSpacing: value }));
}

/** "增加缩进"：在列表项里时走列表嵌套（sinkListItem），否则调整段落 indent 属性 */
export function increaseIndent(schema: Schema): Cmd {
  return (state, dispatch) => {
    const sink = sinkListItem(schema.nodes.list_item);
    if (sink(state, dispatch)) return true;
    return updateBlockAttrs((node) => ({ indent: Math.min(8, (node.attrs.indent ?? 0) + 1) }))(state, dispatch);
  };
}

/** "减少缩进"：在列表项里时走列表提升（liftListItem），否则调整段落 indent 属性 */
export function decreaseIndent(schema: Schema): Cmd {
  return (state, dispatch) => {
    const lift = liftListItem(schema.nodes.list_item);
    if (lift(state, dispatch)) return true;
    return updateBlockAttrs((node) => ({ indent: Math.max(0, (node.attrs.indent ?? 0) - 1) }))(state, dispatch);
  };
}

export interface DocxStylePatch {
  color?: string | null;
  fontFamily?: string | null;
  sizeHalfPt?: number | null;
  highlight?: string | null;
}

/**
 * 字体/字号/颜色/高亮控件的统一落地命令。
 * 有选区时：逐个文字节点合并 patch（保留该节点原有的其它样式属性，不会把整个选区
 * 强行拉成同一份 attrs——这正是 Word「选区内混合格式，只改一个属性」的行为）。
 * 无选区（光标）时：写入 storedMarks，影响接下来要输入的文字，对应 Word 光标处设置格式。
 */
export function setDocxStyle(schema: Schema, patch: DocxStylePatch): Cmd {
  const markType = schema.marks.docxStyle;
  return (state, dispatch) => {
    const { from, to, empty } = state.selection;
    if (empty) {
      if (dispatch) {
        const marks = state.storedMarks || state.selection.$from.marks();
        const current = markType.isInSet(marks);
        const merged = { ...(current?.attrs ?? {}), ...patch };
        const nextMarks = [...marks.filter((m) => m.type !== markType), markType.create(merged)];
        dispatch(state.tr.setStoredMarks(nextMarks));
      }
      return true;
    }
    if (dispatch) {
      let tr = state.tr;
      state.doc.nodesBetween(from, to, (node, pos) => {
        if (!node.isText) return;
        const start = Math.max(pos, from);
        const end = Math.min(pos + node.nodeSize, to);
        if (start >= end) return;
        const existing = markType.isInSet(node.marks);
        const merged = { ...(existing?.attrs ?? {}), ...patch };
        tr = tr.removeMark(start, end, markType);
        tr = tr.addMark(start, end, markType.create(merged));
      });
      dispatch(tr);
    }
    return true;
  };
}

export function buildToolbar(schema: Schema) {
  return {
    bold: toggleMark(schema.marks.strong),
    italic: toggleMark(schema.marks.em),
    underline: toggleMark(schema.marks.underline),
    strike: toggleMark(schema.marks.strike),
    code: toggleMark(schema.marks.code),
    paragraph: setBlockType(schema.nodes.paragraph),
    heading: (level: number) => setBlockType(schema.nodes.heading, { level }),
    bulletList: wrapInList(schema.nodes.bullet_list),
    orderedList: wrapInList(schema.nodes.ordered_list),
    blockquote: wrapIn(schema.nodes.blockquote),
    image: insertImage(schema),
    link: insertLink(schema),
    table: insertTable(3, 3),
    alignLeft: setAlign("left"),
    alignCenter: setAlign("center"),
    alignRight: setAlign("right"),
    alignJustify: setAlign("justify"),
    indentMore: increaseIndent(schema),
    indentLess: decreaseIndent(schema),
  };
}

/** 表格上下文工具栏：光标在表格内时展示（§ Office 式表格编辑）*/
export function buildTableToolbar() {
  return {
    addRowBefore,
    addRowAfter,
    addColumnBefore,
    addColumnAfter,
    deleteRow,
    deleteColumn,
    deleteTable,
    mergeCells,
    splitCell,
    toggleHeaderRow,
    toggleHeaderColumn,
    setCellBackground: (color: string | null) => setCellAttr("background", color),
  };
}

export { isInTable };
