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
  onUpdate?: (id: number, text: string) => void;
  onDelete?: (id: number) => void;
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
  const [activeMenuId, setActiveMenuId] = createSignal<number | null>(null);
  const [editingId, setEditingId] = createSignal<number | null>(null);
  const [editText, setEditText] = createSignal('');

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

  const startEdit = (c: CommentItem, e: MouseEvent) => {
    e.stopPropagation();
    setActiveMenuId(null);
    setEditText(c.text);
    setEditingId(c.id);
  };

  const saveEdit = (id: number) => {
    props.onUpdate?.(id, editText());
    setEditingId(null);
  };

  const cancelEdit = () => {
    setEditingId(null);
    setEditText('');
  };

  const handleDelete = (id: number, e: MouseEvent) => {
    e.stopPropagation();
    setActiveMenuId(null);
    if (window.confirm('确定要删除这条批注吗？')) {
      props.onDelete?.(id);
    }
  };

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
                class={`bg-paper border border-line rounded-lg p-2.5 px-3 cursor-pointer transition-all hover:border-accent-soft hover:shadow-md relative ${flashId() === c.id ? 'border-accent animate-[comment-flash_1.6s_ease-out_1]' : ''}`}
                onClick={() => props.onJump(c.id)}
              >
                <div class="flex justify-between items-start mb-1">
                  <span class="text-xs font-semibold text-ink-1 pt-0.5">{c.author}</span>
                  <div class="relative">
                    <button
                      type="button"
                      class="text-ink-3 hover:text-ink-1 text-xs px-1 py-0.5 rounded hover:bg-surface-2 transition-all"
                      onClick={(e) => {
                        e.stopPropagation();
                        setActiveMenuId(activeMenuId() === c.id ? null : c.id);
                      }}
                      title="更多"
                    >
                      ···
                    </button>
                    <Show when={activeMenuId() === c.id}>
                      <div class="absolute right-0 top-6 z-10 bg-paper border border-line rounded-lg shadow-lg py-1 min-w-[72px]">
                        <button
                          type="button"
                          class="block w-full text-left px-3 py-1.5 text-xs text-ink-2 hover:bg-accent-wash hover:text-accent-ink transition-all"
                          onClick={(e) => startEdit(c, e as unknown as MouseEvent)}
                        >
                          编辑
                        </button>
                        <button
                          type="button"
                           class="block w-full text-left px-3 py-1.5 text-xs text-danger hover:bg-danger/10 transition-all"
                          onClick={(e) => handleDelete(c.id, e as unknown as MouseEvent)}
                        >
                          删除
                        </button>
                      </div>
                    </Show>
                  </div>
                </div>
                <div class="font-mono text-[10.5px] text-ink-3 mb-1">
                  {formatDate(c.date)}
                </div>
                <Show
                  when={editingId() === c.id}
                  fallback={<div class="text-xs text-ink-2 leading-relaxed">{c.text}</div>}
                >
                  <textarea
                    class="w-full min-h-[60px] border border-line-strong rounded-md p-1.5 text-xs text-ink-1 bg-paper resize-y focus:border-accent focus:outline-none"
                    value={editText()}
                    onInput={(e) => setEditText(e.currentTarget.value)}
                    onClick={(e) => e.stopPropagation()}
                    onKeyDown={(e) => {
                      e.stopPropagation();
                      if (e.key === 'Enter' && e.metaKey) saveEdit(c.id);
                      if (e.key === 'Escape') cancelEdit();
                    }}
                  />
                  <div class="flex justify-end gap-1.5 mt-1.5">
                    <button
                      type="button"
                      class="px-2 py-1 text-[11px] rounded border border-line-strong bg-paper text-ink-2 hover:bg-surface-2"
                      onClick={(e) => { e.stopPropagation(); cancelEdit(); }}
                    >
                      取消
                    </button>
                    <button
                      type="button"
                      class="px-2 py-1 text-[11px] rounded border border-accent bg-accent text-white hover:bg-accent-ink"
                      onClick={(e) => { e.stopPropagation(); saveEdit(c.id); }}
                    >
                      保存
                    </button>
                  </div>
                </Show>
              </li>
            )}
          </For>
        </ul>
      </Show>
    </div>
  );
}
