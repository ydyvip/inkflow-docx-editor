import { createSignal, createEffect, createMemo, Show } from 'solid-js';
import { EditorState } from 'prosemirror-state';
import type { Node as PMNode } from 'prosemirror-model';
import { docSchema } from '../schema';
import { highlightPlugin } from '../editor/highlightPlugin';
import { OutlineTree, type OutlineItem } from '../outline/OutlineTree';
import { computeOutline } from '../outline/computeOutline';
import { CommentsPanel, type CommentItem } from '../comments/CommentsPanel';
import { FloatingToolbar } from '../editor/FloatingToolbar';
import {
  VirtualPagedPreview,
  type VirtualPagedPreviewApi,
  type ActiveViewInfo,
} from './VirtualPagedPreview';
import type { DocxPageSetup, DocxSection } from '../parser/ooxml';

interface PreviewPaneProps {
  docJson: any;
  initialComments?: CommentItem[];
  pageSetup?: DocxPageSetup | null;
  header?: any | null;
  footer?: any | null;
  sections?: DocxSection[];
  onCommentsChange?: (comments: CommentItem[]) => void;
}

/**
 * Preview 模块 —— 虚拟化分页只读预览
 */
export function PreviewPane(props: PreviewPaneProps) {
  let apiRef: VirtualPagedPreviewApi | undefined;
  const [dirDoc, setDirDoc] = createSignal<PMNode | null>(null);
  const [commentsVersion, setCommentsVersion] = createSignal(0);
  const [showOutline, setShowOutline] = createSignal(true);
  const [comments, setComments] = createSignal<CommentItem[]>(
    props.initialComments ?? []
  );
  const [activeView, setActiveView] = createSignal<ActiveViewInfo | null>(null);

  // 全局唯一的 doc 真相源：把 JSON 反序列化一次，之后增删批注都重建它
  createEffect(() => {
    const json = props.docJson;
    if (json == null) return;
    setDirDoc(docSchema.nodeFromJSON(json));
  });

  const outlineItems = createMemo<OutlineItem[]>(() => {
    const doc = dirDoc();
    return doc ? computeOutline(doc) : [];
  });

  /** 预览模式下添加批注：作用于全局 doc，重建后 bump version 让分页重渲染 */
  const addCommentOnRange = (from: number, to: number) => {
    const doc = dirDoc();
    const info = activeView();
    if (!doc || !info) return;
    const text = window.prompt('输入批注内容：');
    if (!text) return;
    const existingIds = comments().map((c) => c.id);
    const newId = (existingIds.length ? Math.max(...existingIds) : 0) + 1;
    const globalFrom = info.globalStart + from;
    const globalTo = info.globalStart + to;
    const state = EditorState.create({
      schema: docSchema,
      doc,
      plugins: [highlightPlugin()],
    });
    const tr = state.tr.addMark(
      globalFrom,
      globalTo,
      docSchema.marks.comment.create({ id: newId })
    );
    setDirDoc(tr.doc);
    const updated = [
      ...comments(),
      { id: newId, author: '我', date: new Date().toISOString(), text },
    ];
    setComments(updated);
    setCommentsVersion((v) => v + 1);
    props.onCommentsChange?.(updated);
  };

  const jumpToBlock = (blockId: string) => {
    const api = apiRef;
    if (!api) return;
    api.scrollToPage(api.blockToPage().get(blockId) ?? 0, {
      highlightBlockId: blockId,
    });
  };

  const jumpToComment = (commentId: number) => {
    const api = apiRef;
    if (!api) return;
    api.scrollToPage(api.commentToPage().get(commentId) ?? 0, {
      highlightCommentId: commentId,
    });
  };

  const handleDelete = (id: number) => {
    const doc = dirDoc();
    const updated = comments().filter((c) => c.id !== id);
    setComments(updated);
    setCommentsVersion((v) => v + 1);
    props.onCommentsChange?.(updated);
    if (!doc) return;
    const commentMarkType = docSchema.marks.comment;
    const ranges: Array<{ from: number; to: number }> = [];
    doc.descendants((node, pos) => {
      if (!node.isText) return true;
      node.marks.forEach((m) => {
        if (
          m.type === commentMarkType &&
          Number(m.attrs.id) === Number(id)
        ) {
          ranges.push({ from: pos, to: pos + node.nodeSize });
        }
      });
      return true;
    });
    if (ranges.length === 0) return;
    const state = EditorState.create({ schema: docSchema, doc });
    const tr = state.tr;
    for (const r of ranges) tr.removeMark(r.from, r.to, commentMarkType);
    if (tr.docChanged) setDirDoc(tr.doc);
  };

  return (
    <div class="flex flex-col h-full min-h-0">
      <div class="flex items-center gap-1 flex-wrap px-3.5 py-2.5 bg-surface-1 border-b border-line sticky top-0 z-[5]" role="toolbar" aria-label="预览工具栏">
        <button
          type="button"
          class={`px-2.5 py-1.5 rounded-md text-[13px] font-semibold transition-all border border-transparent hover:bg-surface-2 ${showOutline() ? 'bg-accent-wash text-accent-ink border-accent-soft' : 'text-ink-2'}`}
          onClick={() => setShowOutline((v) => !v)}
          title="显示/隐藏文档目录"
        >
          目录
        </button>
        <span class="flex-1" />
        <span class="text-xs text-ink-3 whitespace-nowrap mr-0.5">只读预览 —— 样式来自解析出的原始 DOCX</span>
      </div>
      <FloatingToolbar
        view={() => activeView()?.view}
        schema={() => docSchema}
        onAddComment={addCommentOnRange}
        showFontControls={false}
      />
      <div class="flex-1 min-h-0 flex overflow-hidden">
        <Show when={showOutline()}>
          <OutlineTree items={outlineItems()} onJump={jumpToBlock} />
        </Show>
        <VirtualPagedPreview
          doc={dirDoc()!}
          commentsVersion={commentsVersion()}
          pageSetup={props.pageSetup ?? null}
          header={props.header ?? null}
          footer={props.footer ?? null}
          sections={props.sections ?? []}
          onApiReady={(api) => {
            apiRef = api;
          }}
          onActiveViewChange={(info) => setActiveView(info)}
        />
        <Show when={comments().length > 0}>
          <CommentsPanel
            comments={comments()}
            onJump={jumpToComment}
            onUpdate={(id, text) => {
              const updated = comments().map((c) =>
                c.id === id ? { ...c, text } : c
              );
              setComments(updated);
              props.onCommentsChange?.(updated);
            }}
            onDelete={handleDelete}
          />
        </Show>
      </div>
    </div>
  );
}
