import { createEffect, onCleanup } from "solid-js";
import { EditorState } from "prosemirror-state";
import { EditorView } from "prosemirror-view";
import { docSchema } from "../schema";
import "../editor/editor.css";
import "./preview.css";

interface PreviewPaneProps {
  docJson: any;
}

/**
 * Preview 模块 —— Mode 1（推荐）：readonly EditorView
 * 直接用当前 JSON 重新创建一个不可编辑的 EditorView，
 * 保证预览与编辑器使用完全相同的渲染规则（同一份 schema）。
 *
 * createEffect 会追踪 props.docJson：每次它变化时，先清理上一次的
 * EditorView（通过内部 onCleanup 注册），再用最新 JSON 重建一个。
 */
export function PreviewPane(props: PreviewPaneProps) {
  let hostEl: HTMLDivElement | undefined;

  createEffect(() => {
    const json = props.docJson; // 建立对 docJson 的响应式依赖
    if (!hostEl) return;

    const doc = docSchema.nodeFromJSON(json);
    const state = EditorState.create({ schema: docSchema, doc });
    const view = new EditorView(hostEl, {
      state,
      editable: () => false,
    });

    onCleanup(() => view.destroy());
  });

  return (
    <div class="editor-page-wrap preview-wrap">
      <div class="editor-page preview-page" ref={hostEl} />
    </div>
  );
}
