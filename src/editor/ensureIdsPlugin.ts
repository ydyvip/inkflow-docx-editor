/**
 * ensureIds 插件
 * ------------------------------------------------------------
 * 目录树 / 高亮接口都依赖 blockId（paragraph/heading）和 cellId
 * （table_cell/table_header）。这些 ID 在 OOXML 解析阶段由
 * ooxml.ts 统一分配，但编辑过程中新产生的节点（回车新建的段落、
 * 表格插入的新行/列……）默认 attrs 里这两个字段是 null。
 *
 * 用 appendTransaction 在每次文档变化后检查一遍：谁没有 ID 就
 * 补一个，保证"通过接口高亮段落/单元格"和目录树在自由编辑之后
 * 依然可靠，而不仅仅在刚解析完的那一刻有效。
 */
import { Plugin } from "prosemirror-state";

function randomId(prefix: string): string {
  return `${prefix}${Math.random().toString(36).slice(2, 9)}`;
}

export function ensureIdsPlugin(): Plugin {
  return new Plugin({
    appendTransaction(transactions, _oldState, newState) {
      if (!transactions.some((tr) => tr.docChanged)) return null;
      let tr = newState.tr;
      let changed = false;
      newState.doc.descendants((node, pos) => {
        if ((node.type.name === "paragraph" || node.type.name === "heading") && !node.attrs.blockId) {
          tr = tr.setNodeMarkup(pos, undefined, { ...node.attrs, blockId: randomId("b") });
          changed = true;
        } else if ((node.type.name === "table_cell" || node.type.name === "table_header") && !node.attrs.cellId) {
          tr = tr.setNodeMarkup(pos, undefined, { ...node.attrs, cellId: randomId("c") });
          changed = true;
        }
        return true;
      });
      return changed ? tr : null;
    },
  });
}
