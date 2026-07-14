/**
 * Highlight 插件
 * ------------------------------------------------------------
 * 需求："支持能通过接口高亮其中的段落、单元格"，以及
 * "点击批注要精确高亮被批注的那段内容"（而不是整段/整块）。
 *
 * 实现方式：不修改文档内容，而是用 ProseMirror Decoration 在渲染层
 * 叠加高亮样式——这是 ProseMirror 里"高亮但不产生 transaction /
 * 不进历史记录"的标准做法。支持两种粒度：
 *
 *   - ids：按 blockId（paragraph/heading）/ cellId（表格单元格）整块高亮
 *   - ranges：按精确的文档位置区间（from-to）做行内高亮，用于批注这类
 *     "只高亮被批注的那几个字，而不是整段"的场景
 *
 * ranges 是原始文档位置，文档发生变化时需要用 tr.mapping 跟着重新映射，
 * 否则后续编辑会让高亮位置错位。
 *
 * 对外通过 EditorPane 暴露的 API（highlightBlock/highlightCell/
 * clearHighlights）调用，内部通过 transaction 的 meta 更新插件状态。
 */
import { Plugin, PluginKey } from 'prosemirror-state';
import { Decoration, DecorationSet } from 'prosemirror-view';
import type { Node as PMNode } from 'prosemirror-model';

export interface HighlightRange {
  from: number;
  to: number;
}

export interface HighlightMeta {
  ids?: string[]; // 要高亮的 blockId / cellId 集合（并集判断，不区分类型）
  ranges?: HighlightRange[]; // 要高亮的精确文档区间（用于批注等"部分内容"高亮）
  mode: 'replace' | 'add' | 'clear' | 'clear-ranges';
}

interface HighlightState {
  ids: Set<string>;
  ranges: HighlightRange[];
}

export const highlightPluginKey = new PluginKey<HighlightState>('highlight');

function buildDecorations(doc: PMNode, state: HighlightState): DecorationSet {
  const decorations: Decoration[] = [];

  if (state.ids.size > 0) {
    doc.descendants((node, pos) => {
      const id = node.attrs?.blockId ?? node.attrs?.cellId;
      if (id && state.ids.has(id)) {
        decorations.push(
          Decoration.node(pos, pos + node.nodeSize, { class: 'pm-highlight' })
        );
      }
      return true;
    });
  }

  for (const range of state.ranges) {
    if (range.to > range.from) {
      decorations.push(
        Decoration.inline(range.from, range.to, { class: 'pm-highlight-range' })
      );
    }
  }

  return decorations.length
    ? DecorationSet.create(doc, decorations)
    : DecorationSet.empty;
}

export function highlightPlugin(): Plugin<HighlightState> {
  return new Plugin<HighlightState>({
    key: highlightPluginKey,
    state: {
      init: () => ({ ids: new Set<string>(), ranges: [] }),
      apply(tr, prev) {
        const meta = tr.getMeta(highlightPluginKey) as
          HighlightMeta | undefined;

        // 文档变化时，已有的精确区间高亮要跟着映射，否则位置会错位
        let ranges = tr.docChanged
          ? prev.ranges.map((r) => ({
              from: tr.mapping.map(r.from),
              to: tr.mapping.map(r.to),
            }))
          : prev.ranges;
        let ids = prev.ids;

        if (!meta) return { ids, ranges };
        if (meta.mode === 'clear') return { ids: new Set(), ranges: [] };
        if (meta.mode === 'clear-ranges') return { ids, ranges: [] };
        if (meta.mode === 'replace') {
          return { ids: new Set(meta.ids ?? []), ranges: meta.ranges ?? [] };
        }
        // add
        return {
          ids: new Set([...ids, ...(meta.ids ?? [])]),
          ranges: [...ranges, ...(meta.ranges ?? [])],
        };
      },
    },
    props: {
      decorations(state) {
        const hs = highlightPluginKey.getState(state) ?? {
          ids: new Set<string>(),
          ranges: [],
        };
        return buildDecorations(state.doc, hs);
      },
    },
  });
}
