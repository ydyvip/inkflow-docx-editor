import { For, Show, createSignal, createEffect } from 'solid-js';
import '../outline/outline.css';
import './comments.css';

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
 * 数据来源：OOXML 解析出的 word/comments.xml（初始批注）+ 编辑器内新增的批注，
 * 由 EditorPane 统一管理并通过 props 下发。点击某条批注会跳转并高亮
 * 其在正文中对应的批注锚点范围（comment mark），对应"支持显示批注内容"需求。
 *
 * 新增批注后列表会追加到末尾——如果已有很多条批注，新的一条可能刚好在
 * 可视区域之外，看起来像是"没有刷新"。这里主动检测批注数量变多，
 * 自动把面板滚动到最新一条并做一次短暂高亮，确保"刚刚新增的批注"
 * 肉眼可见，而不只是数据层面同步了。
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
    <div class="comments-panel">
      <div class="side-panel-title">批注（{props.comments.length}）</div>
      <Show
        when={props.comments.length > 0}
        fallback={
          <div class="side-panel-empty">
            暂无批注，选中文字后点击工具栏"批注"可以新增
          </div>
        }
      >
        <ul class="comments-list" ref={listEl}>
          <For each={props.comments}>
            {(c) => (
              <li
                class={`comment-item${flashId() === c.id ? ' is-new' : ''}`}
                onClick={() => props.onJump(c.id)}
              >
                <div class="comment-meta">
                  <span class="comment-author">{c.author}</span>
                  <span class="comment-date">{formatDate(c.date)}</span>
                </div>
                <div class="comment-text">{c.text}</div>
              </li>
            )}
          </For>
        </ul>
      </Show>
    </div>
  );
}
