import { createSignal, onMount, onCleanup, For, Show } from "solid-js";
import { EditorState, TextSelection } from "prosemirror-state";
import { EditorView } from "prosemirror-view";
import { baseKeymap } from "prosemirror-commands";
import { keymap } from "prosemirror-keymap";
import { docSchema } from "../schema";
import { buildEditorPlugins } from "./pluginsSetup";
import { buildToolbar, markActive, blockActive } from "./commands";
import { pluginRegistry } from "../plugins/registry";
import { highlightPluginKey, type HighlightMeta } from "./highlightPlugin";
import { OutlineTree, type OutlineItem } from "../outline/OutlineTree";
import { CommentsPanel, type CommentItem } from "../comments/CommentsPanel";
import "./editor.css";

export interface EditorApi {
  /** 通过 blockId 高亮一个段落/标题（不移动光标、不滚动）*/
  highlightBlock: (blockId: string) => void;
  /** 通过 cellId 高亮一个表格单元格 */
  highlightCell: (cellId: string) => void;
  /** 清除所有高亮 */
  clearHighlights: () => void;
  /** 滚动到指定 blockId/cellId 并高亮 */
  scrollToBlock: (id: string) => void;
  /** 当前批注列表（含解析出的 + 编辑器内新增的）*/
  getComments: () => CommentItem[];
}

interface EditorPaneProps {
  initialDoc: any;
  initialComments?: CommentItem[];
  onChange: (json: any) => void;
  onCommentsChange?: (comments: CommentItem[]) => void;
  onReady?: (api: EditorApi) => void;
}

function findNodeById(doc: any, id: string): { pos: number; node: any } | null {
  let result: { pos: number; node: any } | null = null;
  doc.descendants((node: any, pos: number) => {
    if (result) return false;
    if (node.attrs?.blockId === id || node.attrs?.cellId === id) {
      result = { pos, node };
      return false;
    }
    return true;
  });
  return result;
}

/**
 * Editor 模块
 * 约束：只能操作 ProseMirror JSON；所有变更必须通过 transaction（§6.3）。
 * 本组件本身不持有"业务状态"，每次 transaction 后把最新的
 * doc.toJSON() 上抛给父组件——JSON 才是唯一真相源。
 *
 * 新增三块能力（均建立在 schema 里新增的 blockId/cellId/comment 之上）：
 *   - 左侧目录树：遍历 heading 节点生成，点击跳转
 *   - 右侧批注面板：来自 OOXML 解析 + 编辑器内新增，点击跳转并高亮锚点
 *   - highlightPlugin 驱动的"通过接口高亮段落/单元格" API（onReady 暴露）
 */
export function EditorPane(props: EditorPaneProps) {
  let hostEl: HTMLDivElement | undefined;
  let view: EditorView | undefined;
  const [version, setVersion] = createSignal(0);
  const [comments, setComments] = createSignal<CommentItem[]>(props.initialComments ?? []);
  const [showOutline, setShowOutline] = createSignal(true);
  const [showComments, setShowComments] = createSignal((props.initialComments?.length ?? 0) > 0);

  const dispatchHighlight = (meta: HighlightMeta) => {
    if (!view) return;
    view.dispatch(view.state.tr.setMeta(highlightPluginKey, meta));
  };

  const highlightBlock = (blockId: string) => dispatchHighlight({ ids: [blockId], mode: "replace" });
  const highlightCell = (cellId: string) => dispatchHighlight({ ids: [cellId], mode: "replace" });
  const clearHighlights = () => dispatchHighlight({ ids: [], mode: "clear" });

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
    view.focus();
  };

  const jumpToComment = (commentId: number) => {
    if (!view) return;
    let targetPos: number | null = null;
    view.state.doc.descendants((node, pos) => {
      if (targetPos !== null) return false;
      if (node.isText && node.marks.some((m) => m.type.name === "comment" && m.attrs.id === commentId)) {
        targetPos = pos;
        return false;
      }
      return true;
    });
    if (targetPos === null) return;
    const $pos = view.state.doc.resolve(targetPos);
    let blockId: string | null = null;
    for (let d = $pos.depth; d >= 0; d--) {
      const n = $pos.node(d);
      if (n.attrs?.blockId) {
        blockId = n.attrs.blockId;
        break;
      }
    }
    const tr = view.state.tr.setSelection(TextSelection.near($pos)).scrollIntoView();
    if (blockId) tr.setMeta(highlightPluginKey, { ids: [blockId], mode: "replace" } as HighlightMeta);
    view.dispatch(tr);
    view.focus();
  };

  const addCommentOnSelection = () => {
    if (!view) return;
    const { from, to, empty } = view.state.selection;
    if (empty) {
      window.alert("请先在正文中选中要批注的文字");
      return;
    }
    const text = window.prompt("输入批注内容：");
    if (!text) return;
    const existingIds = comments().map((c) => c.id);
    const newId = (existingIds.length ? Math.max(...existingIds) : 0) + 1;
    view.dispatch(view.state.tr.addMark(from, to, docSchema.marks.comment.create({ id: newId })));
    const updated = [...comments(), { id: newId, author: "我", date: new Date().toISOString(), text }];
    setComments(updated);
    props.onCommentsChange?.(updated);
    setShowComments(true);
  };

  onMount(() => {
    if (!hostEl) return;

    const doc = docSchema.nodeFromJSON(props.initialDoc);
    const state = EditorState.create({
      schema: docSchema,
      doc,
      plugins: [...buildEditorPlugins(docSchema), keymap(baseKeymap)],
    });

    view = new EditorView(hostEl, {
      state,
      dispatchTransaction(tr) {
        if (!view) return;
        const newState = view.state.apply(tr);
        view.updateState(newState);
        if (tr.docChanged) {
          props.onChange(newState.doc.toJSON());
        }
        setVersion((n) => n + 1);
      },
    });

    const api: EditorApi = { highlightBlock, highlightCell, clearHighlights, scrollToBlock, getComments: () => comments() };
    props.onReady?.(api);
    props.onCommentsChange?.(comments());
    // 便于外部（宿主页面 / 控制台 / e2e 测试）直接调用高亮接口
    (window as any).inkflowEditor = api;

    // 初次挂载时没有任何 transaction 触发，version 信号不会自动变化；
    // 手动 bump 一次，让依赖 version() 的目录树/工具栏状态在 view 就绪后重新计算
    setVersion((n) => n + 1);
  });

  onCleanup(() => {
    view?.destroy();
    view = undefined;
    if ((window as any).inkflowEditor) delete (window as any).inkflowEditor;
  });

  const runCommand = (cmd: (state: any, dispatch: any) => boolean) => {
    if (!view) return;
    cmd(view.state, view.dispatch);
    view.focus();
  };

  const runPluginToolbarItem = (run: (v: { state: any; dispatch: any }) => void) => {
    if (!view) return;
    run({ state: view.state, dispatch: view.dispatch });
    view.focus();
  };

  const isMarkActive = (markType: any) => {
    version();
    return view ? markActive(view.state, markType) : false;
  };

  const isBlockActive = (nodeType: any, attrs?: Record<string, any>) => {
    version();
    return view ? blockActive(view.state, nodeType, attrs) : false;
  };

  const outlineItems = (): OutlineItem[] => {
    version();
    if (!view) return [];
    const items: OutlineItem[] = [];
    view.state.doc.descendants((node) => {
      if (node.type.name === "heading") {
        items.push({ blockId: node.attrs.blockId ?? "", level: node.attrs.level, text: node.textContent });
      }
      return true;
    });
    return items;
  };

  const toolbar = buildToolbar(docSchema);

  interface BtnDef {
    id: string;
    label: string;
    onClick: () => void;
    active?: () => boolean;
    title?: string;
  }

  const Btn = (def: BtnDef) => (
    <button
      type="button"
      class={`toolbar-btn${def.active?.() ? " is-active" : ""}`}
      onMouseDown={(e) => {
        e.preventDefault();
        def.onClick();
      }}
      title={def.title ?? def.label}
    >
      {def.label}
    </button>
  );

  const headingButtons: BtnDef[] = [1, 2, 3].map((level) => ({
    id: `h${level}`,
    label: `H${level}`,
    onClick: () => runCommand(toolbar.heading(level)),
    active: () => isBlockActive(docSchema.nodes.heading, { level }),
    title: `标题 ${level}`,
  }));

  const pluginButtons: BtnDef[] = pluginRegistry.all().flatMap((p) =>
    (p.toolbar ? p.toolbar(docSchema) : []).map((item) => ({
      id: item.id,
      label: item.label,
      onClick: () => runPluginToolbarItem(item.run),
      active: () => {
        version();
        return view && item.isActive ? item.isActive(view.state) : false;
      },
    }))
  );

  return (
    <div class="editor-shell">
      <div class="toolbar" role="toolbar" aria-label="格式工具栏">
        <Btn id="bold" label="B" onClick={() => runCommand(toolbar.bold)} active={() => isMarkActive(docSchema.marks.strong)} title="加粗" />
        <Btn id="italic" label="I" onClick={() => runCommand(toolbar.italic)} active={() => isMarkActive(docSchema.marks.em)} title="斜体" />
        <Btn id="code" label="</>" onClick={() => runCommand(toolbar.code)} active={() => isMarkActive(docSchema.marks.code)} title="行内代码" />
        <span class="toolbar-divider" />
        <Btn id="p" label="正文" onClick={() => runCommand(toolbar.paragraph)} active={() => isBlockActive(docSchema.nodes.paragraph)} />
        <For each={headingButtons}>{(def) => <Btn {...def} />}</For>
        <span class="toolbar-divider" />
        <Btn id="ul" label="• 列表" onClick={() => runCommand(toolbar.bulletList)} />
        <Btn id="ol" label="1. 列表" onClick={() => runCommand(toolbar.orderedList)} />
        <Btn id="quote" label="❝ 引用" onClick={() => runCommand(toolbar.blockquote)} />
        <span class="toolbar-divider" />
        <Btn id="link" label="🔗 链接" onClick={() => runCommand(toolbar.link)} />
        <Btn id="image" label="🖼 图片" onClick={() => runCommand(toolbar.image)} />
        <Btn id="table" label="▦ 表格" onClick={() => runCommand(toolbar.table)} />
        <For each={pluginButtons}>{(def) => <Btn {...def} />}</For>
        <span class="toolbar-divider" />
        <Btn id="add-comment" label="✍ 批注" onClick={addCommentOnSelection} title="为选中文字添加批注" />
        <span class="toolbar-spacer" />
        <Btn id="toggle-outline" label="🗂 目录" onClick={() => setShowOutline((v) => !v)} active={() => showOutline()} title="显示/隐藏文档目录" />
        <Btn id="toggle-comments" label="💬 批注列表" onClick={() => setShowComments((v) => !v)} active={() => showComments()} title="显示/隐藏批注面板" />
      </div>
      <div class="editor-body">
        <Show when={showOutline()}>
          <OutlineTree items={outlineItems()} onJump={scrollToBlock} />
        </Show>
        <div class="editor-page-wrap">
          <div class="editor-page" ref={hostEl} />
        </div>
        <Show when={showComments()}>
          <CommentsPanel comments={comments()} onJump={jumpToComment} />
        </Show>
      </div>
    </div>
  );
}
