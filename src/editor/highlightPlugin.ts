/**
 * Highlight 插件
 * ------------------------------------------------------------
 * 需求："支持能通过接口高亮其中的段落、单元格"。
 *
 * 实现方式：不修改文档内容，而是用 ProseMirror Decoration 在渲染层
 * 叠加高亮样式——这是 ProseMirror 里"高亮但不产生 transaction /
 * 不进历史记录"的标准做法。查找节点时依据 schema 里新增的
 * `blockId`（paragraph/heading）和 `cellId`（table_cell/table_header）
 * 属性（见 src/schema/index.ts，来自 OOXML 解析器分配的稳定 ID）。
 *
 * 对外通过 EditorPane 暴露的 API（highlightBlock/highlightCell/
 * clearHighlights）调用，内部通过 transaction 的 meta 更新插件状态。
 */
import { Plugin, PluginKey } from "prosemirror-state";
import { Decoration, DecorationSet } from "prosemirror-view";
import type { Node as PMNode } from "prosemirror-model";

export interface HighlightMeta {
  ids: string[]; // 要高亮的 blockId / cellId 集合（并集判断，不区分类型）
  mode: "replace" | "add" | "clear";
}

export const highlightPluginKey = new PluginKey<Set<string>>("highlight");

function buildDecorations(doc: PMNode, ids: Set<string>): DecorationSet {
  if (ids.size === 0) return DecorationSet.empty;
  const decorations: Decoration[] = [];
  doc.descendants((node, pos) => {
    const id = node.attrs?.blockId ?? node.attrs?.cellId;
    if (id && ids.has(id)) {
      decorations.push(Decoration.node(pos, pos + node.nodeSize, { class: "pm-highlight" }));
    }
    return true;
  });
  return DecorationSet.create(doc, decorations);
}

export function highlightPlugin(): Plugin<Set<string>> {
  return new Plugin<Set<string>>({
    key: highlightPluginKey,
    state: {
      init: () => new Set<string>(),
      apply(tr, prevIds) {
        const meta = tr.getMeta(highlightPluginKey) as HighlightMeta | undefined;
        if (!meta) return prevIds;
        if (meta.mode === "clear") return new Set();
        if (meta.mode === "replace") return new Set(meta.ids);
        return new Set([...prevIds, ...meta.ids]);
      },
    },
    props: {
      decorations(state) {
        const ids = highlightPluginKey.getState(state) ?? new Set<string>();
        return buildDecorations(state.doc, ids);
      },
    },
  });
}
