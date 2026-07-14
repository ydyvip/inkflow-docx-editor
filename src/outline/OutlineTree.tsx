import { For, Show } from 'solid-js';

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
 * 点击条目通过 blockId 定位到编辑器中的对应位置。
 */
export function OutlineTree(props: OutlineTreeProps) {
  const levelPadding = (level: number) => {
    if (level === 1) return 'pl-2';
    if (level === 2) return 'pl-4';
    if (level === 3) return 'pl-7';
    if (level === 4) return 'pl-10';
    if (level === 5) return 'pl-12';
    if (level === 6) return 'pl-14';
    return 'pl-16';
  };

  const levelStyle = (level: number) => {
    if (level === 1) return 'font-semibold text-ink-1';
    if (level === 2) return 'text-ink-2';
    if (level === 3) return 'text-ink-2 text-xs';
    if (level === 4) return 'text-ink-3 text-xs';
    return 'text-ink-3 text-[11px] italic';
  };

  return (
    <div class="w-60 flex-shrink-0 border-r border-line bg-surface-1 overflow-y-auto px-2.5 py-3.5">
      <div class="font-mono text-[11px] tracking-wider uppercase text-ink-3 px-2 pb-2.5">
        文档目录
      </div>
      <Show
        when={props.items.length > 0}
        fallback={
          <div class="text-xs text-ink-3 px-2 py-2 leading-relaxed">
            暂无标题，添加 H1–H3 后会显示在这里
          </div>
        }
      >
        <ul class="list-none m-0 p-0">
          <For each={props.items}>
            {(item) => (
              <li>
                <button
                  type="button"
                  onClick={() => props.onJump(item.blockId)}
                  title={item.text}
                  class={`block w-full text-left bg-transparent border-0 rounded-md py-1.5 px-2 text-[13px] cursor-pointer whitespace-nowrap overflow-hidden text-ellipsis transition-colors hover:bg-surface-2 hover:text-ink-1 ${levelPadding(item.level)} ${levelStyle(item.level)} ${props.activeId === item.blockId ? 'bg-accent-wash text-accent-ink font-semibold' : ''}`}
                >
                  {item.text || '（空标题）'}
                </button>
              </li>
            )}
          </For>
        </ul>
      </Show>
    </div>
  );
}
