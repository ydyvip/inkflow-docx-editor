import { For, Show } from "solid-js";
import "./outline.css";

export interface OutlineItem {
  blockId: string;
  level: number;
  text: string;
}

interface OutlineTreeProps {
  items: OutlineItem[];
  activeId?: string | null;
  onJump: (blockId: string) => void;
}

/**
 * Outline 模块 —— 文档目录树
 * 数据来源：EditorPane 每次 transaction 后重新遍历 doc 收集 heading 节点
 * （level + 文本 + blockId），点击条目通过 blockId 定位到编辑器中的对应位置。
 * 这是"编辑器要能显示 docx 文档的目录树"需求的实现。
 */
export function OutlineTree(props: OutlineTreeProps) {
  return (
    <div class="outline-panel">
      <div class="side-panel-title">文档目录</div>
      <Show when={props.items.length > 0} fallback={<div class="side-panel-empty">暂无标题，添加 H1–H3 后会显示在这里</div>}>
        <ul class="outline-list">
          <For each={props.items}>
            {(item) => (
              <li class={`outline-item outline-level-${item.level}${props.activeId === item.blockId ? " is-active" : ""}`}>
                <button type="button" onClick={() => props.onJump(item.blockId)} title={item.text}>
                  {item.text || "（空标题）"}
                </button>
              </li>
            )}
          </For>
        </ul>
      </Show>
    </div>
  );
}
