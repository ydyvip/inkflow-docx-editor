import type { Node as PMNode } from 'prosemirror-model';
import type { EditorView } from 'prosemirror-view';
import type { OutlineItem } from './OutlineTree';

/** 遍历文档收集所有 heading 节点，生成目录树数据（H1-H9 全部支持）*/
export function computeOutline(doc: PMNode): OutlineItem[] {
  const items: OutlineItem[] = [];
  doc.descendants((node) => {
    if (node.type.name === 'heading') {
      items.push({
        blockId: node.attrs.blockId ?? '',
        level: node.attrs.level,
        text: node.textContent,
      });
    }
    return true;
  });
  return items;
}

/** 按 blockId / cellId 查找节点位置 */
export function findNodeById(
  doc: PMNode,
  id: string
): { pos: number; node: PMNode } | null {
  let result: { pos: number; node: PMNode } | null = null;
  doc.descendants((node, pos) => {
    if (result) return false;
    if (node.attrs?.blockId === id || node.attrs?.cellId === id) {
      result = { pos, node };
      return false;
    }
    return true;
  });
  return result;
}

/**
 * 手动把文档中某个位置滚动到可见区域，不依赖 ProseMirror 内置的
 * scrollToSelection()。之所以自己实现：PM 的 scrollToSelection() 内部会检查
 * "浏览器当前的真实 DOM 选区是否落在编辑器 DOM 内"，而只读视图
 * （Preview，editable:false）的 DOM 节点 tabIndex 是 -1，
 * 天生就没法真正拿到焦点、选区也就落不进去，导致滚动静默失效。
 * 这里直接用 view.coordsAtPos() 拿到目标位置的视口坐标，手动计算并
 * 滚动最近的可滚动祖先容器——对可编辑和只读视图都同样可靠。
 */
export function scrollPosIntoView(
  view: EditorView,
  container: HTMLElement,
  pos: number,
  margin = 96
) {
  let coords: { top: number; bottom: number; left: number; right: number };
  try {
    coords = view.coordsAtPos(pos);
  } catch {
    return;
  }
  const rect = container.getBoundingClientRect();
  const offsetWithinContent = coords.top - rect.top + container.scrollTop;
  const target = Math.max(0, offsetWithinContent - margin);
  container.scrollTo({ top: target, behavior: 'smooth' });
}
