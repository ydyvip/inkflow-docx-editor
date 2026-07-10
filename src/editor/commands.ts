import type { Schema, Node as PMNode } from 'prosemirror-model';
import type { EditorState, Transaction } from 'prosemirror-state';
import { toggleMark, setBlockType, wrapIn } from 'prosemirror-commands';
import {
  wrapInList,
  sinkListItem,
  liftListItem,
} from 'prosemirror-schema-list';
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
} from 'prosemirror-tables';
import { insertTable } from './tableUtils';
import type { CellBorder } from '../schema';

type Cmd = (
  state: EditorState,
  dispatch?: (tr: Transaction) => void
) => boolean;

export function markActive(state: EditorState, markType: any): boolean {
  const { from, $from, to, empty } = state.selection;
  if (empty) return !!markType.isInSet(state.storedMarks || $from.marks());
  return state.doc.rangeHasMark(from, to, markType);
}

export function blockActive(
  state: EditorState,
  nodeType: any,
  attrs: Record<string, any> = {}
): boolean {
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
export function currentDocxStyleAttr(
  state: EditorState,
  markType: any,
  key: string
): any {
  const marks = state.storedMarks || state.selection.$from.marks();
  const mark = markType.isInSet(marks);
  return mark?.attrs?.[key] ?? null;
}

/** 找到选区所在最近的 list_item 祖先里的 bullet_list/ordered_list 节点（用于列表样式下拉框回显）*/
export function currentListInfo(
  state: EditorState
): { kind: 'bullet' | 'ordered'; style: string } | null {
  const $from = state.selection.$from;
  for (let d = $from.depth; d >= 0; d--) {
    const node = $from.node(d);
    if (node.type.name === 'bullet_list')
      return { kind: 'bullet', style: node.attrs.bulletStyle ?? 'disc' };
    if (node.type.name === 'ordered_list')
      return { kind: 'ordered', style: node.attrs.numberFormat ?? 'decimal' };
  }
  return null;
}

/** 找到选区所在的 table 节点（用于表格属性面板回显对齐方式等）*/
export function currentTableNode(state: EditorState): PMNode | null {
  const $from = state.selection.$from;
  for (let d = $from.depth; d >= 0; d--) {
    const node = $from.node(d);
    if (node.type.name === 'table') return node;
  }
  return null;
}

export function insertImage(schema: Schema): Cmd {
  return (state, dispatch) => {
    const src = window.prompt('图片地址（URL 或粘贴 base64）：');
    if (!src) return false;
    if (dispatch) {
      const node = schema.nodes.image.create({ src, alt: '' });
      dispatch(state.tr.replaceSelectionWith(node).scrollIntoView());
    }
    return true;
  };
}

/**
 * 插入/编辑/移除超链接。选中已带链接的文字时会读出当前地址回填到输入框，
 * 清空后确定即可移除链接——对应 Word 里"编辑超链接"/"取消超链接"。
 */
export function insertLink(schema: Schema): Cmd {
  return (state, dispatch) => {
    const { from, to, empty } = state.selection;
    if (empty) {
      window.alert('请先选中要添加或编辑链接的文字');
      return false;
    }
    const markType = schema.marks.link;
    // 用 nodesBetween 直接扫描选区内的文字节点找现有链接的 href，
    // 比 resolve(pos).marks() 更可靠——后者在节点边界位置的取值语义比较微妙，
    // 选区恰好等于整个链接范围时容易判断不到已有的 mark。
    let currentHref: string | null = null;
    state.doc.nodesBetween(from, to, (node) => {
      if (currentHref) return false;
      if (node.isText) {
        const mark = markType.isInSet(node.marks);
        if (mark) currentHref = mark.attrs.href;
      }
      return true;
    });
    const href = window.prompt(
      '链接地址（清空后确定可移除链接）：',
      currentHref ?? 'https://'
    );
    if (href === null) return false; // 用户取消
    if (dispatch) {
      let tr = state.tr.removeMark(from, to, markType);
      if (href.trim()) {
        tr = tr.addMark(
          from,
          to,
          markType.create({ href: href.trim(), title: null })
        );
      }
      dispatch(tr);
    }
    return true;
  };
}

/** 通用：把选区覆盖到的每个 paragraph/heading 节点的属性做一次 patch（setNodeMarkup 不改变节点大小，可在同一个事务里安全地对多个位置连续调用）*/
function updateBlockAttrs(
  getPatch: (node: PMNode) => Record<string, any> | null
): Cmd {
  return (state, dispatch) => {
    const { from, to } = state.selection;
    let tr = state.tr;
    let changed = false;
    state.doc.nodesBetween(from, to, (node, pos) => {
      if (node.type.name !== 'paragraph' && node.type.name !== 'heading')
        return;
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
    return updateBlockAttrs((node) => ({
      indent: Math.min(8, (node.attrs.indent ?? 0) + 1),
    }))(state, dispatch);
  };
}

/** "减少缩进"：在列表项里时走列表提升（liftListItem），否则调整段落 indent 属性 */
export function decreaseIndent(schema: Schema): Cmd {
  return (state, dispatch) => {
    const lift = liftListItem(schema.nodes.list_item);
    if (lift(state, dispatch)) return true;
    return updateBlockAttrs((node) => ({
      indent: Math.max(0, (node.attrs.indent ?? 0) - 1),
    }))(state, dispatch);
  };
}

/**
 * 清除格式：对应 Word 的"清除格式"（Ctrl+空格 / 橡皮擦图标）。
 * 移除选区内所有"格式类" mark（加粗/斜体/下划线/删除线/代码/字体样式），
 * 并把覆盖到的段落/标题的排版属性（对齐/缩进/行距/命名样式）重置为默认值。
 * 批注（comment）和超链接（link）不算"格式"，予以保留——与 Word 行为一致。
 */
const CLEARABLE_MARK_NAMES = [
  'strong',
  'em',
  'underline',
  'strike',
  'code',
  'docxStyle',
];

export function clearFormatting(schema: Schema): Cmd {
  return (state, dispatch) => {
    const { from, to, empty } = state.selection;
    if (!dispatch) return true;

    if (empty) {
      let tr = state.tr.setStoredMarks([]);
      const $pos = state.doc.resolve(from);
      for (let d = $pos.depth; d >= 0; d--) {
        const node = $pos.node(d);
        if (node.type.name === 'paragraph' || node.type.name === 'heading') {
          const pos = $pos.before(d);
          tr = tr.setNodeMarkup(pos, undefined, {
            ...node.attrs,
            align: null,
            indent: 0,
            lineSpacing: null,
            styleName: null,
          });
          break;
        }
      }
      dispatch(tr);
      return true;
    }

    let tr = state.tr;
    for (const name of CLEARABLE_MARK_NAMES) {
      const markType = (schema.marks as Record<string, any>)[name];
      if (markType) tr = tr.removeMark(from, to, markType);
    }
    state.doc.nodesBetween(from, to, (node, pos) => {
      if (node.type.name === 'paragraph' || node.type.name === 'heading') {
        tr = tr.setNodeMarkup(pos, undefined, {
          ...node.attrs,
          align: null,
          indent: 0,
          lineSpacing: null,
          styleName: null,
        });
      }
    });
    dispatch(tr);
    return true;
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
        const nextMarks = [
          ...marks.filter((m) => m.type !== markType),
          markType.create(merged),
        ];
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

/**
 * 项目符号 / 编号样式（§ 项目符号与编号）。
 * 已经在列表里时直接改当前列表的类型/样式属性；不在列表里时包一层新列表。
 * kind="none" 表示"取消列表"。
 */
export function setListStyle(
  schema: Schema,
  kind: 'none' | 'bullet' | 'ordered',
  styleValue?: string
): Cmd {
  return (state, dispatch) => {
    const $from = state.selection.$from;
    for (let d = $from.depth; d >= 0; d--) {
      const node = $from.node(d);
      if (
        node.type.name === 'bullet_list' ||
        node.type.name === 'ordered_list'
      ) {
        if (kind === 'none') {
          return liftListItem(schema.nodes.list_item)(state, dispatch);
        }
        const pos = $from.before(d);
        const newType =
          kind === 'bullet'
            ? schema.nodes.bullet_list
            : schema.nodes.ordered_list;
        const attrKey = kind === 'bullet' ? 'bulletStyle' : 'numberFormat';
        if (dispatch) {
          const tr = state.tr;
          if (node.type === newType) {
            tr.setNodeMarkup(pos, undefined, {
              ...node.attrs,
              [attrKey]: styleValue,
            });
          } else {
            tr.setNodeMarkup(pos, newType, { [attrKey]: styleValue });
          }
          dispatch(tr);
        }
        return true;
      }
    }
    if (kind === 'none') return false;
    const listType =
      kind === 'bullet' ? schema.nodes.bullet_list : schema.nodes.ordered_list;
    const attrs =
      kind === 'bullet'
        ? { bulletStyle: styleValue }
        : { numberFormat: styleValue };
    return wrapInList(listType, attrs)(state, dispatch);
  };
}

/** 表格在页面中的对齐方式（左/中/右）*/
export function setTableAlign(align: string | null): Cmd {
  return (state, dispatch) => {
    const $from = state.selection.$from;
    for (let d = $from.depth; d >= 0; d--) {
      const node = $from.node(d);
      if (node.type.name === 'table') {
        if (dispatch) {
          const pos = $from.before(d);
          dispatch(
            state.tr.setNodeMarkup(pos, undefined, { ...node.attrs, align })
          );
        }
        return true;
      }
    }
    return false;
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
    alignLeft: setAlign('left'),
    alignCenter: setAlign('center'),
    alignRight: setAlign('right'),
    alignJustify: setAlign('justify'),
    indentMore: increaseIndent(schema),
    indentLess: decreaseIndent(schema),
    clearFormatting: clearFormatting(schema),
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
    setCellBackground: (color: string | null) =>
      setCellAttr('background', color),
    setCellVAlign: (value: string | null) => setCellAttr('valign', value),
    setCellTextDirection: (value: string | null) =>
      setCellAttr('textDirection', value),
    setCellBorder: (value: CellBorder | null) =>
      setCellAttr('cellBorder', value),
  };
}

export { isInTable };
