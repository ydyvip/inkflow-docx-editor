import type { Node as PMNode } from "prosemirror-model";
import type { OutlineItem } from "./OutlineTree";

/** 遍历文档收集所有 heading 节点，生成目录树数据（H1-H9 全部支持）*/
export function computeOutline(doc: PMNode): OutlineItem[] {
  const items: OutlineItem[] = [];
  doc.descendants((node) => {
    if (node.type.name === "heading") {
      items.push({ blockId: node.attrs.blockId ?? "", level: node.attrs.level, text: node.textContent });
    }
    return true;
  });
  return items;
}

/** 按 blockId / cellId 查找节点位置 */
export function findNodeById(doc: PMNode, id: string): { pos: number; node: PMNode } | null {
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
