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

interface PreviewPaneProps {
  docJson: any;
}

/**
 * Preview 模块 —— readonly EditorView
 */
export function PreviewPane(props: PreviewPaneProps) {
  let hostEl: HTMLDivElement | undefined;
  let view: EditorView | undefined;
  const [version, setVersion] = createSignal(0);
  const [showOutline, setShowOutline] = createSignal(true);

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
      <div class="flex-1 min-h-0 flex overflow-hidden">
        <Show when={showOutline()}>
          <OutlineTree items={outlineItems()} onJump={scrollToBlock} />
        </Show>
        <div class="flex-1 min-w-0 overflow-y-auto px-6 pb-24 pt-10 bg-canvas">
          <div class="max-w-[760px] min-h-[900px] mx-auto bg-paper shadow-[0_1px_2px_rgba(23,26,33,0.06),0_12px_32px_rgba(23,26,33,0.08)] rounded-[3px] px-[76px] py-[72px] border-l-[3px] border-l-accent editor-page preview-mode" ref={hostEl} />
        </div>
      </div>
    </div>
  );
}
