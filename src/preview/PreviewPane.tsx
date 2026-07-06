import { createSignal, createEffect, onCleanup, Show } from "solid-js";
import { EditorState, TextSelection } from "prosemirror-state";
import { EditorView } from "prosemirror-view";
import { docSchema } from "../schema";
import { highlightPluginKey, highlightPlugin, type HighlightMeta } from "../editor/highlightPlugin";
import { OutlineTree, type OutlineItem } from "../outline/OutlineTree";
import { computeOutline, findNodeById } from "../outline/computeOutline";
import "../editor/editor.css";
import "../outline/outline.css";
import "./preview.css";

interface PreviewPaneProps {
  docJson: any;
}

/**
 * Preview 模块 —— Mode 1（推荐）：readonly EditorView
 * 直接用当前 JSON 重新创建一个不可编辑的 EditorView，
 * 保证预览与编辑器使用完全相同的渲染规则（同一份 schema，
 * 因此颜色/字体/字号/对齐等真实 DOCX 样式在预览里天然一致）。
 *
 * 同时也带一份目录树（与编辑器共用 computeOutline），满足
 * "预览也需要显示目录"的要求——只读视图依然可以点击跳转+高亮，
 * 因为 dispatch 程序化 transaction 不受 editable:false 限制
 * （editable 只拦截用户直接在 DOM 里编辑的输入）。
 */
export function PreviewPane(props: PreviewPaneProps) {
  let hostEl: HTMLDivElement | undefined;
  let view: EditorView | undefined;
  const [version, setVersion] = createSignal(0);
  const [showOutline, setShowOutline] = createSignal(true);

  createEffect(() => {
    const json = props.docJson; // 建立对 docJson 的响应式依赖
    if (!hostEl) return;

    const doc = docSchema.nodeFromJSON(json);
    const state = EditorState.create({ schema: docSchema, doc, plugins: [highlightPlugin()] });
    view = new EditorView(hostEl, {
      state,
      editable: () => false,
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
    if (!view) return;
    const found = findNodeById(view.state.doc, id);
    if (!found) return;
    const selPos = Math.min(found.pos + 1, view.state.doc.content.size);
    const tr = view.state.tr
      .setSelection(TextSelection.near(view.state.doc.resolve(selPos)))
      .setMeta(highlightPluginKey, { ids: [id], mode: "replace" } as HighlightMeta)
      .scrollIntoView();
    view.dispatch(tr);
  };

  return (
    <div class="editor-shell">
      <div class="toolbar" role="toolbar" aria-label="预览工具栏">
        <button
          type="button"
          class={`toolbar-btn${showOutline() ? " is-active" : ""}`}
          onClick={() => setShowOutline((v) => !v)}
          title="显示/隐藏文档目录"
        >
          🗂 目录
        </button>
        <span class="toolbar-spacer" />
        <span class="toolbar-label">只读预览 —— 样式来自解析出的原始 DOCX</span>
      </div>
      <div class="editor-body">
        <Show when={showOutline()}>
          <OutlineTree items={outlineItems()} onJump={scrollToBlock} />
        </Show>
        <div class="editor-page-wrap preview-wrap">
          <div class="editor-page preview-page" ref={hostEl} />
        </div>
      </div>
    </div>
  );
}
