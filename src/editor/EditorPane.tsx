import { createSignal, onMount, onCleanup, For, Show } from "solid-js";
import { EditorState, TextSelection } from "prosemirror-state";
import { EditorView } from "prosemirror-view";
import { baseKeymap } from "prosemirror-commands";
import { keymap } from "prosemirror-keymap";
import { docSchema } from "../schema";
import { buildEditorPlugins } from "./pluginsSetup";
import {
  buildToolbar,
  buildTableToolbar,
  markActive,
  currentBlockAttr,
  currentDocxStyleAttr,
  setDocxStyle,
  setLineSpacing,
  isInTable,
} from "./commands";
import { pluginRegistry } from "../plugins/registry";
import { highlightPluginKey, type HighlightMeta } from "./highlightPlugin";
import { OutlineTree, type OutlineItem } from "../outline/OutlineTree";
import { computeOutline, findNodeById } from "../outline/computeOutline";
import { CommentsPanel, type CommentItem } from "../comments/CommentsPanel";
import { FONT_FAMILIES, FONT_SIZES_PT, LINE_SPACING_OPTIONS, HIGHLIGHT_OPTIONS, HEADING_LEVELS } from "./formatOptions";
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

/**
 * Editor 模块
 * 约束：只能操作 ProseMirror JSON；所有变更必须通过 transaction（§6.3）。
 * 本组件本身不持有"业务状态"，每次 transaction 后把最新的
 * doc.toJSON() 上抛给父组件——JSON 才是唯一真相源。
 *
 * 工具栏分三组，对应 Office 的功能划分：
 *   - Row1 结构：正文/标题1-9 选择、列表、引用、插入（链接/图片/表格/插件）、批注、面板开关
 *   - Row2 字体与段落：字体/字号/颜色/高亮/加粗斜体下划线删除线、对齐、缩进、行距
 *   - Row3（条件显示）：光标进入表格时出现的表格上下文工具栏
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

  const outlineItems = (): OutlineItem[] => {
    version();
    return view ? computeOutline(view.state.doc) : [];
  };

  // ---- 段落/字体控件的"当前值"读取（驱动下拉框/颜色选择器回显）----

  const currentAlign = () => {
    version();
    return view ? currentBlockAttr(view.state, "align") || "left" : "left";
  };

  const currentHeadingValue = () => {
    version();
    if (!view) return "paragraph";
    const node = view.state.selection.$from.parent;
    return node.type.name === "heading" ? `h${node.attrs.level}` : "paragraph";
  };

  const currentFontFamily = () => {
    version();
    return view ? currentDocxStyleAttr(view.state, docSchema.marks.docxStyle, "fontFamily") ?? "" : "";
  };

  const currentFontSizePt = () => {
    version();
    if (!view) return "";
    const half = currentDocxStyleAttr(view.state, docSchema.marks.docxStyle, "sizeHalfPt");
    return half ? String(Number(half) / 2) : "";
  };

  const currentColor = () => {
    version();
    return view ? currentDocxStyleAttr(view.state, docSchema.marks.docxStyle, "color") || "#1c2321" : "#1c2321";
  };

  const currentHighlight = () => {
    version();
    return view ? currentDocxStyleAttr(view.state, docSchema.marks.docxStyle, "highlight") ?? "" : "";
  };

  const currentLineSpacing = () => {
    version();
    return view ? String(currentBlockAttr(view.state, "lineSpacing") ?? "") : "";
  };

  const inTable = () => {
    version();
    return view ? isInTable(view.state) : false;
  };

  // ---- 命令绑定 ----

  const toolbar = buildToolbar(docSchema);
  const tableToolbar = buildTableToolbar();

  const onHeadingSelect = (value: string) => {
    if (value === "paragraph") runCommand(toolbar.paragraph);
    else runCommand(toolbar.heading(Number(value.slice(1))));
  };
  const onFontFamilyChange = (value: string) => runCommand(setDocxStyle(docSchema, { fontFamily: value || null }));
  const onFontSizeChange = (ptStr: string) => {
    if (!ptStr) return;
    runCommand(setDocxStyle(docSchema, { sizeHalfPt: Math.round(Number(ptStr) * 2) }));
  };
  const onColorChange = (hex: string) => runCommand(setDocxStyle(docSchema, { color: hex }));
  const onHighlightChange = (value: string) => runCommand(setDocxStyle(docSchema, { highlight: value || null }));
  const onLineSpacingChange = (value: string) => runCommand(setLineSpacing(value ? Number(value) : null));

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
      {/* Row 1：结构 —— 标题级别 / 列表 / 插入 / 批注 / 面板开关 */}
      <div class="toolbar" role="toolbar" aria-label="结构工具栏">
        <select
          class="toolbar-select toolbar-select-heading"
          value={currentHeadingValue()}
          onChange={(e) => onHeadingSelect(e.currentTarget.value)}
          title="段落样式"
        >
          <option value="paragraph">正文</option>
          <For each={HEADING_LEVELS}>{(level) => <option value={`h${level}`}>标题 {level}</option>}</For>
        </select>
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

      {/* Row 2：字体 & 段落格式 —— 尽量贴近 Office「开始」选项卡 */}
      <div class="toolbar toolbar-format" role="toolbar" aria-label="字体与段落工具栏">
        <select class="toolbar-select" value={currentFontFamily()} onChange={(e) => onFontFamilyChange(e.currentTarget.value)} title="字体">
          <For each={FONT_FAMILIES}>{(f) => <option value={f.value}>{f.label}</option>}</For>
        </select>
        <select
          class="toolbar-select toolbar-select-narrow"
          value={currentFontSizePt()}
          onChange={(e) => onFontSizeChange(e.currentTarget.value)}
          title="字号"
        >
          <option value="">字号</option>
          <For each={FONT_SIZES_PT}>{(pt) => <option value={pt}>{pt}</option>}</For>
        </select>
        <Btn id="bold" label="B" onClick={() => runCommand(toolbar.bold)} active={() => isMarkActive(docSchema.marks.strong)} title="加粗" />
        <Btn id="italic" label="I" onClick={() => runCommand(toolbar.italic)} active={() => isMarkActive(docSchema.marks.em)} title="斜体" />
        <Btn id="underline" label="U" onClick={() => runCommand(toolbar.underline)} active={() => isMarkActive(docSchema.marks.underline)} title="下划线" />
        <Btn id="strike" label="S" onClick={() => runCommand(toolbar.strike)} active={() => isMarkActive(docSchema.marks.strike)} title="删除线" />
        <Btn id="code" label="</>" onClick={() => runCommand(toolbar.code)} active={() => isMarkActive(docSchema.marks.code)} title="行内代码" />
        <label class="toolbar-color" title="文字颜色">
          A
          <input type="color" value={currentColor()} onInput={(e) => onColorChange(e.currentTarget.value)} />
        </label>
        <select
          class="toolbar-select toolbar-select-narrow"
          value={currentHighlight()}
          onChange={(e) => onHighlightChange(e.currentTarget.value)}
          title="高亮"
        >
          <For each={HIGHLIGHT_OPTIONS}>{(h) => <option value={h.value}>{h.label}</option>}</For>
        </select>
        <span class="toolbar-divider" />
        <Btn id="align-left" label="≡L" onClick={() => runCommand(toolbar.alignLeft)} active={() => currentAlign() === "left"} title="左对齐" />
        <Btn id="align-center" label="≡C" onClick={() => runCommand(toolbar.alignCenter)} active={() => currentAlign() === "center"} title="居中" />
        <Btn id="align-right" label="≡R" onClick={() => runCommand(toolbar.alignRight)} active={() => currentAlign() === "right"} title="右对齐" />
        <Btn
          id="align-justify"
          label="≡J"
          onClick={() => runCommand(toolbar.alignJustify)}
          active={() => currentAlign() === "justify"}
          title="两端对齐"
        />
        <span class="toolbar-divider" />
        <Btn id="indent-less" label="⇤缩进" onClick={() => runCommand(toolbar.indentLess)} title="减少缩进" />
        <Btn id="indent-more" label="缩进⇥" onClick={() => runCommand(toolbar.indentMore)} title="增加缩进" />
        <select
          class="toolbar-select toolbar-select-narrow"
          value={currentLineSpacing()}
          onChange={(e) => onLineSpacingChange(e.currentTarget.value)}
          title="行距"
        >
          <For each={LINE_SPACING_OPTIONS}>{(l) => <option value={l.value}>{l.label}</option>}</For>
        </select>
      </div>

      {/* Row 3（条件显示）：光标位于表格内时出现，Office 式表格上下文工具栏 */}
      <Show when={inTable()}>
        <div class="toolbar toolbar-table" role="toolbar" aria-label="表格工具栏">
          <span class="toolbar-label">表格：</span>
          <Btn id="row-before" label="↑插入行" onClick={() => runCommand(tableToolbar.addRowBefore)} />
          <Btn id="row-after" label="↓插入行" onClick={() => runCommand(tableToolbar.addRowAfter)} />
          <Btn id="col-before" label="←插入列" onClick={() => runCommand(tableToolbar.addColumnBefore)} />
          <Btn id="col-after" label="→插入列" onClick={() => runCommand(tableToolbar.addColumnAfter)} />
          <span class="toolbar-divider" />
          <Btn id="del-row" label="删除行" onClick={() => runCommand(tableToolbar.deleteRow)} />
          <Btn id="del-col" label="删除列" onClick={() => runCommand(tableToolbar.deleteColumn)} />
          <Btn id="del-table" label="删除表格" onClick={() => runCommand(tableToolbar.deleteTable)} />
          <span class="toolbar-divider" />
          <Btn id="merge" label="合并单元格" onClick={() => runCommand(tableToolbar.mergeCells)} />
          <Btn id="split" label="拆分单元格" onClick={() => runCommand(tableToolbar.splitCell)} />
          <Btn id="header-row" label="表头行" onClick={() => runCommand(tableToolbar.toggleHeaderRow)} />
          <Btn id="header-col" label="表头列" onClick={() => runCommand(tableToolbar.toggleHeaderColumn)} />
          <label class="toolbar-color" title="单元格底色">
            底色
            <input type="color" value="#ffffff" onInput={(e) => runCommand(tableToolbar.setCellBackground(e.currentTarget.value))} />
          </label>
        </div>
      </Show>

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
