import type { Schema } from "prosemirror-model";
import type { EditorState, Transaction } from "prosemirror-state";
import { toggleMark, setBlockType, wrapIn } from "prosemirror-commands";
import { wrapInList } from "prosemirror-schema-list";
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

export function buildToolbar(schema: Schema) {
  return {
    bold: toggleMark(schema.marks.strong),
    italic: toggleMark(schema.marks.em),
    code: toggleMark(schema.marks.code),
    paragraph: setBlockType(schema.nodes.paragraph),
    heading: (level: number) => setBlockType(schema.nodes.heading, { level }),
    bulletList: wrapInList(schema.nodes.bullet_list),
    orderedList: wrapInList(schema.nodes.ordered_list),
    blockquote: wrapIn(schema.nodes.blockquote),
    image: insertImage(schema),
    link: insertLink(schema),
    table: insertTable(3, 3),
  };
}
