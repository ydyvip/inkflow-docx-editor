import { For, Show } from "solid-js";
import "../outline/outline.css";
import "./comments.css";

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
  if (!d) return "";
  const parsed = new Date(d);
  if (Number.isNaN(parsed.getTime())) return d;
  return parsed.toLocaleDateString("zh-CN", { year: "numeric", month: "2-digit", day: "2-digit" });
}

/**
 * Comments 模块 —— 批注面板
 * 数据来源：OOXML 解析出的 word/comments.xml（初始批注）+ 编辑器内新增的批注，
 * 由 EditorPane 统一管理并通过 props 下发。点击某条批注会跳转并高亮
 * 其在正文中对应的批注锚点范围（comment mark），对应"支持显示批注内容"需求。
 */
export function CommentsPanel(props: CommentsPanelProps) {
  return (
    <div class="comments-panel">
      <div class="side-panel-title">批注（{props.comments.length}）</div>
      <Show when={props.comments.length > 0} fallback={<div class="side-panel-empty">暂无批注，选中文字后点击工具栏"批注"可以新增</div>}>
        <ul class="comments-list">
          <For each={props.comments}>
            {(c) => (
              <li class="comment-item" onClick={() => props.onJump(c.id)}>
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
