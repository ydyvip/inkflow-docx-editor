import { For, Show, createSignal, createEffect } from 'solid-js';

export interface CommentItem {
  id: number;
  author: string;
  date: string | null;
  text: string;
}

interface CommentsPanelProps {
  comments: CommentItem[];
  onJump: (commentId: number) => void;
}

function formatDate(d: string | null): string {
  if (!d) return '';
  const parsed = new Date(d);
  if (Number.isNaN(parsed.getTime())) return d;
  return parsed.toLocaleDateString('zh-CN', {
    year: 'numeric',
    month: '2-digit',
    day: '2-digit',
  });
}

/**
 * Comments 模块 —— 批注面板
 */
export function CommentsPanel(props: CommentsPanelProps) {
  let listEl: HTMLUListElement | undefined;
  let prevCount = props.comments.length;
  const [flashId, setFlashId] = createSignal<number | null>(null);

  createEffect(() => {
    const list = props.comments;
    if (list.length > prevCount) {
      const newest = list[list.length - 1];
      setFlashId(newest.id);
      queueMicrotask(() => {
        listEl?.lastElementChild?.scrollIntoView({
          behavior: 'smooth',
          block: 'nearest',
        });
      });
      setTimeout(
        () => setFlashId((id) => (id === newest.id ? null : id)),
        1600
      );
    }
    prevCount = list.length;
  });

  return (
    <div class="w-60 flex-shrink-0 border-l border-line bg-surface-1 overflow-y-auto px-2.5 py-3.5">
      <div class="font-mono text-[11px] tracking-wider uppercase text-ink-3 px-2 pb-2.5">
        批注（{props.comments.length}）
      </div>
      <Show
        when={props.comments.length > 0}
        fallback={
          <div class="text-xs text-ink-3 px-2 py-2 leading-relaxed">
            暂无批注，选中文字后点击工具栏"批注"可以新增
          </div>
        }
      >
        <ul class="list-none m-0 p-0 flex flex-col gap-2" ref={listEl}>
          <For each={props.comments}>
            {(c) => (
              <li
                class={`bg-paper border border-line rounded-lg p-2.5 px-3 cursor-pointer transition-all hover:border-accent-soft hover:shadow-md ${flashId() === c.id ? 'border-accent animate-[comment-flash_1.6s_ease-out_1]' : ''}`}
                onClick={() => props.onJump(c.id)}
              >
                <div class="flex justify-between items-baseline mb-1">
                  <span class="text-xs font-semibold text-ink-1">{c.author}</span>
                  <span class="font-mono text-[10.5px] text-ink-3">{formatDate(c.date)}</span>
                </div>
                <div class="text-xs text-ink-2 leading-relaxed">{c.text}</div>
              </li>
            )}
          </For>
        </ul>
      </Show>
    </div>
  );
}
