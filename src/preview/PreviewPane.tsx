import { createSignal, createEffect, onCleanup, Show } from 'solid-js';
import { EditorState, TextSelection } from 'prosemirror-state';
import { EditorView } from 'prosemirror-view';
import { docSchema } from '../schema';
import { pluginRegistry } from '../plugins/registry';
import {
  highlightPluginKey,
  highlightPlugin,
  type HighlightMeta,
} from '../editor/highlightPlugin';
import { OutlineTree, type OutlineItem } from '../outline/OutlineTree';
import {
  computeOutline,
  findNodeById,
  scrollPosIntoView,
} from '../outline/computeOutline';
import { CommentsPanel, type CommentItem } from '../comments/CommentsPanel';
import { FloatingToolbar } from '../editor/FloatingToolbar';

interface PreviewPaneProps {
  docJson: any;
  initialComments?: CommentItem[];
  onCommentsChange?: (comments: CommentItem[]) => void;
}

/**
 * Preview 模块 —— readonly EditorView
 */
export function PreviewPane(props: PreviewPaneProps) {
  let hostEl: HTMLDivElement | undefined;
  let view: EditorView | undefined;
  const [version, setVersion] = createSignal(0);
  const [showOutline, setShowOutline] = createSignal(true);
  const [comments, setComments] = createSignal<CommentItem[]>(
    props.initialComments ?? []
  );

  createEffect(() => {
    const json = props.docJson;
    if (!hostEl) return;

    const doc = docSchema.nodeFromJSON(json);
    const state = EditorState.create({
      schema: docSchema,
      doc,
      plugins: [highlightPlugin()],
    });
    view = new EditorView(hostEl, {
      state,
      editable: () => false,
      nodeViews: pluginRegistry.nodeViews(docSchema),
      dispatchTransaction(tr) {
        if (!view) return;
        view.updateState(view.state.apply(tr));
        setVersion((n) => n + 1);
      },
    });
    setVersion((n) => n + 1);

    onCleanup(() => {
      view?.destroy();
      view = undefined;
    });
  });

  const outlineItems = (): OutlineItem[] => {
    version();
    return view ? computeOutline(view.state.doc) : [];
  };

  const scrollToBlock = (id: string) => {
    if (!view || !hostEl) return;
    const found = findNodeById(view.state.doc, id);
    if (!found) return;
    const selPos = Math.min(found.pos + 1, view.state.doc.content.size);
    const tr = view.state.tr
      .setSelection(TextSelection.near(view.state.doc.resolve(selPos)))
      .setMeta(highlightPluginKey, {
        ids: [id],
        mode: 'replace',
      } as HighlightMeta);
    view.dispatch(tr);
    scrollPosIntoView(view, hostEl.parentElement ?? hostEl, selPos);
  };

  /** 预览模式下添加批注 */
  const addCommentOnRange = (from: number, to: number) => {
    if (!view) return;
    const text = window.prompt('输入批注内容：');
    if (!text) return;
    const existingIds = comments().map((c) => c.id);
    const newId = (existingIds.length ? Math.max(...existingIds) : 0) + 1;
    view.dispatch(
      view.state.tr.addMark(
        from,
        to,
        docSchema.marks.comment.create({ id: newId })
      )
    );
    const updated = [
      ...comments(),
      { id: newId, author: '我', date: new Date().toISOString(), text },
    ];
    setComments(updated);
    props.onCommentsChange?.(updated);
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
        view={() => view}
        schema={() => docSchema}
        onAddComment={addCommentOnRange}
        showFontControls={false}
      />
      <div class="flex-1 min-h-0 flex overflow-hidden">
        <Show when={showOutline()}>
          <OutlineTree items={outlineItems()} onJump={scrollToBlock} />
        </Show>
        <div class="flex-1 min-w-0 overflow-y-auto px-6 pb-24 pt-10 bg-canvas">
          <div class="max-w-[760px] min-h-[900px] mx-auto bg-paper shadow-[0_1px_2px_rgba(23,26,33,0.06),0_12px_32px_rgba(23,26,33,0.08)] rounded-[3px] px-[76px] py-[72px] border-l-[3px] border-l-accent editor-page preview-mode" ref={hostEl} />
        </div>
        <Show when={comments().length > 0}>
          <CommentsPanel
            comments={comments()}
            onJump={(id) => {
              if (!view) return;
              let from: number | null = null;
              let to: number | null = null;
              view.state.doc.descendants((node, pos) => {
                if (
                  node.isText &&
                  node.marks.some((m) => m.type.name === 'comment' && Number(m.attrs.id) === Number(id))
                ) {
                  if (from === null || pos < from) from = pos;
                  const end = pos + node.nodeSize;
                  if (to === null || end > to) to = end;
                }
                return true;
              });
              if (from === null || to === null) return;
              const tr = view.state.tr
                .setSelection(TextSelection.near(view.state.doc.resolve(from)))
                .setMeta(highlightPluginKey, {
                  ranges: [{ from, to }],
                  mode: 'replace',
                } as HighlightMeta);
              view.dispatch(tr);
              scrollPosIntoView(view, hostEl!.parentElement ?? hostEl!, from);
            }}
            onUpdate={(id, text) => {
              const updated = comments().map((c) =>
                c.id === id ? { ...c, text } : c
              );
              setComments(updated);
              props.onCommentsChange?.(updated);
            }}
            onDelete={(id) => {
              if (!view) return;

              const updated = comments().filter((c) => c.id !== id);
              setComments(updated);
              props.onCommentsChange?.(updated);

              const commentMarkType = docSchema.marks.comment;
              if (!commentMarkType) return;

              const tr = view.state.tr;
              const ranges: Array<{ from: number; to: number }> = [];

              view.state.doc.descendants((node, pos) => {
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

              for (const r of ranges) {
                tr.removeMark(r.from, r.to, commentMarkType);
              }

              if (tr.docChanged) {
                view.dispatch(tr);
                view.dispatch(
                  view.state.tr.setMeta(highlightPluginKey, {
                    mode: 'clear-ranges',
                  } as HighlightMeta)
                );
              }
            }}
          />
        </Show>
      </div>
    </div>
  );
}
