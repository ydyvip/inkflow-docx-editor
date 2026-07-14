import { createSignal, onMount, onCleanup, For, Show } from 'solid-js';
import { EditorState, TextSelection } from 'prosemirror-state';
import { EditorView } from 'prosemirror-view';
import { baseKeymap } from 'prosemirror-commands';
import { keymap } from 'prosemirror-keymap';
import { docSchema } from '../schema';
import { buildEditorPlugins } from './pluginsSetup';
import {
  buildToolbar,
  buildTableToolbar,
  markActive,
  currentBlockAttr,
  currentDocxStyleAttr,
  currentListInfo,
  currentTableNode,
  setDocxStyle,
  setLineSpacing,
  setListStyle,
  setTableAlign,
  isInTable,
} from './commands';
import type { CellBorder } from '../schema';
import {
  TablePropertiesModal,
  type TablePropertiesValues,
} from './TablePropertiesModal';
import { pluginRegistry } from '../plugins/registry';
import { highlightPluginKey, type HighlightMeta } from './highlightPlugin';
import { OutlineTree, type OutlineItem } from '../outline/OutlineTree';
import {
  computeOutline,
  findNodeById,
  scrollPosIntoView,
} from '../outline/computeOutline';
import { CommentsPanel, type CommentItem } from '../comments/CommentsPanel';
import {
  FONT_FAMILIES,
  FONT_SIZES_PT,
  LINE_SPACING_OPTIONS,
  HIGHLIGHT_OPTIONS,
  HEADING_LEVELS,
  LIST_STYLE_OPTIONS,
} from './formatOptions';

export interface EditorApi {
  highlightBlock: (blockId: string) => void;
  highlightCell: (cellId: string) => void;
  clearHighlights: () => void;
  scrollToBlock: (id: string) => void;
  getComments: () => CommentItem[];
}

interface EditorPaneProps {
  initialDoc: any;
  initialComments?: CommentItem[];
  onChange: (json: any) => void;
  onCommentsChange?: (comments: CommentItem[]) => void;
  onReady?: (api: EditorApi) => void;
}

export function EditorPane(props: EditorPaneProps) {
  let hostEl: HTMLDivElement | undefined;
  let view: EditorView | undefined;
  const [version, setVersion] = createSignal(0);
  const [comments, setComments] = createSignal<CommentItem[]>(
    props.initialComments ?? []
  );
  const [showOutline, setShowOutline] = createSignal(true);
  const [showComments, setShowComments] = createSignal(
    (props.initialComments?.length ?? 0) > 0
  );
  const [showTableProps, setShowTableProps] = createSignal(false);

  const dispatchHighlight = (meta: HighlightMeta) => {
    if (!view) return;
    view.dispatch(view.state.tr.setMeta(highlightPluginKey, meta));
  };

  const highlightBlock = (blockId: string) =>
    dispatchHighlight({ ids: [blockId], mode: 'replace' });
  const highlightCell = (cellId: string) =>
    dispatchHighlight({ ids: [cellId], mode: 'replace' });
  const clearHighlights = () => dispatchHighlight({ ids: [], mode: 'clear' });

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
    view.focus();
    view.dispatch(tr);
    scrollPosIntoView(view, hostEl.parentElement ?? hostEl, selPos);
  };

  const jumpToComment = (commentId: number) => {
    if (!view || !hostEl) return;
    let from: number | null = null;
    let to: number | null = null;
    view.state.doc.descendants((node, pos) => {
      if (
        node.isText &&
        node.marks.some(
          (m) => m.type.name === 'comment' && m.attrs.id === commentId
        )
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
    view.focus();
    view.dispatch(tr);
    scrollPosIntoView(view, hostEl.parentElement ?? hostEl, from);
  };

  const addCommentOnSelection = () => {
    if (!view) return;
    const { from, to, empty } = view.state.selection;
    if (empty) {
      window.alert('请先在正文中选中要批注的文字');
      return;
    }
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
      nodeViews: pluginRegistry.nodeViews(docSchema),
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

    const api: EditorApi = {
      highlightBlock,
      highlightCell,
      clearHighlights,
      scrollToBlock,
      getComments: () => comments(),
    };
    props.onReady?.(api);
    props.onCommentsChange?.(comments());
    (window as any).inkflowEditor = api;

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

  const runPluginToolbarItem = (
    run: (v: { state: any; dispatch: any }) => void
  ) => {
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

  const currentAlign = () => {
    version();
    return view ? currentBlockAttr(view.state, 'align') || 'left' : 'left';
  };

  const currentHeadingValue = () => {
    version();
    if (!view) return 'paragraph';
    const node = view.state.selection.$from.parent;
    return node.type.name === 'heading' ? `h${node.attrs.level}` : 'paragraph';
  };

  const currentFontFamily = () => {
    version();
    return view
      ? (currentDocxStyleAttr(
          view.state,
          docSchema.marks.docxStyle,
          'fontFamily'
        ) ?? '')
      : '';
  };

  const currentFontSizePt = () => {
    version();
    if (!view) return '';
    const half = currentDocxStyleAttr(
      view.state,
      docSchema.marks.docxStyle,
      'sizeHalfPt'
    );
    return half ? String(Number(half) / 2) : '';
  };

  const currentColor = () => {
    version();
    return view
      ? currentDocxStyleAttr(view.state, docSchema.marks.docxStyle, 'color') ||
          '#1c2321'
      : '#1c2321';
  };

  const currentHighlight = () => {
    version();
    return view
      ? (currentDocxStyleAttr(
          view.state,
          docSchema.marks.docxStyle,
          'highlight'
        ) ?? '')
      : '';
  };

  const currentLineSpacing = () => {
    version();
    return view
      ? String(currentBlockAttr(view.state, 'lineSpacing') ?? '')
      : '';
  };

  const inTable = () => {
    version();
    return view ? isInTable(view.state) : false;
  };

  const currentListValue = () => {
    version();
    if (!view) return 'none';
    const info = currentListInfo(view.state);
    if (!info) return 'none';
    return `${info.kind}:${info.style}`;
  };

  const isLinkActive = () => {
    version();
    return view ? markActive(view.state, docSchema.marks.link) : false;
  };

  const toolbar = buildToolbar(docSchema);
  const tableToolbar = buildTableToolbar();

  const onHeadingSelect = (value: string) => {
    if (value === 'paragraph') runCommand(toolbar.paragraph);
    else runCommand(toolbar.heading(Number(value.slice(1))));
  };
  const onFontFamilyChange = (value: string) =>
    runCommand(setDocxStyle(docSchema, { fontFamily: value || null }));
  const onFontSizeChange = (ptStr: string) => {
    if (!ptStr) return;
    runCommand(
      setDocxStyle(docSchema, { sizeHalfPt: Math.round(Number(ptStr) * 2) })
    );
  };
  const onColorChange = (hex: string) =>
    runCommand(setDocxStyle(docSchema, { color: hex }));
  const onHighlightChange = (value: string) =>
    runCommand(setDocxStyle(docSchema, { highlight: value || null }));
  const onLineSpacingChange = (value: string) =>
    runCommand(setLineSpacing(value ? Number(value) : null));

  const onListStyleChange = (value: string) => {
    if (value === 'none') {
      runCommand(setListStyle(docSchema, 'none'));
      return;
    }
    const [kind, style] = value.split(':') as ['bullet' | 'ordered', string];
    runCommand(setListStyle(docSchema, kind, style));
  };

  const tablePropsInitial = (): TablePropertiesValues => {
    const table = view ? currentTableNode(view.state) : null;
    return {
      tableAlign: (table?.attrs?.align as any) ?? 'left',
      cellVAlign: 'top',
      textDirection: 'horizontal',
      borderStyle: 'single',
      borderWidth: 1,
      borderColor: '#333333',
      cellBackground: '#ffffff',
    };
  };

  const applyTableProperties = (values: TablePropertiesValues) => {
    if (!view) return;
    runCommand(setTableAlign(values.tableAlign));
    runCommand(
      tableToolbar.setCellVAlign(
        values.cellVAlign === 'top' ? null : values.cellVAlign
      )
    );
    runCommand(
      tableToolbar.setCellTextDirection(
        values.textDirection === 'horizontal' ? null : 'vertical'
      )
    );
    const border: CellBorder | null =
      values.borderStyle === 'none'
        ? { style: 'none', width: 0, color: values.borderColor }
        : {
            style: values.borderStyle,
            width: values.borderWidth,
            color: values.borderColor,
          };
    runCommand(tableToolbar.setCellBorder(border));
    runCommand(tableToolbar.setCellBackground(values.cellBackground));
  };

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
      class={`border border-transparent bg-transparent text-ink-2 font-inherit text-[13px] font-semibold px-2.5 py-1.5 rounded-md cursor-pointer leading-none transition-all whitespace-nowrap hover:bg-surface-2 hover:text-ink-1 focus-visible:outline-2 focus-visible:outline-accent focus-visible:outline-offset-1 ${def.active?.() ? 'bg-accent-wash text-accent-ink border-accent-soft' : ''}`}
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
    <div class="flex flex-col h-full min-h-0">
      {/* Row 1 */}
      <div class="flex items-center gap-1 flex-wrap px-3.5 py-2.5 bg-surface-1 border-b border-line sticky top-0 z-[5]" role="toolbar" aria-label="结构工具栏">
        <select
          class="h-[30px] border border-line-strong bg-paper text-ink-1 text-[13px] rounded-md px-1.5 cursor-pointer max-w-[100px] font-semibold hover:border-accent-soft"
          value={currentHeadingValue()}
          onChange={(e) => onHeadingSelect(e.currentTarget.value)}
          title="段落样式"
        >
          <option value="paragraph">正文</option>
          <For each={HEADING_LEVELS}>
            {(level) => <option value={`h${level}`}>标题 {level}</option>}
          </For>
        </select>
        <span class="w-px h-5 bg-line mx-1" />
        <select
          class="h-[30px] border border-line-strong bg-paper text-ink-1 text-[13px] rounded-md px-1.5 cursor-pointer max-w-[130px] hover:border-accent-soft"
          value={currentListValue()}
          onChange={(e) => onListStyleChange(e.currentTarget.value)}
          title="项目符号与编号"
        >
          <For each={LIST_STYLE_OPTIONS}>
            {(opt) => <option value={opt.value}>{opt.label}</option>}
          </For>
        </select>
        <Btn id="quote" label="引用" onClick={() => runCommand(toolbar.blockquote)} />
        <span class="w-px h-5 bg-line mx-1" />
        <Btn id="link" label="链接" onClick={() => runCommand(toolbar.link)} active={isLinkActive} title="插入/编辑/移除链接" />
        <Btn id="image" label="图片" onClick={() => runCommand(toolbar.image)} />
        <Btn id="table" label="表格" onClick={() => runCommand(toolbar.table)} />
        <For each={pluginButtons}>{(def) => <Btn {...def} />}</For>
        <span class="w-px h-5 bg-line mx-1" />
        <Btn id="add-comment" label="批注" onClick={addCommentOnSelection} title="为选中文字添加批注" />
        <span class="flex-1" />
        <Btn id="toggle-outline" label="目录" onClick={() => setShowOutline((v) => !v)} active={() => showOutline()} title="显示/隐藏文档目录" />
        <Btn id="toggle-comments" label="批注列表" onClick={() => setShowComments((v) => !v)} active={() => showComments()} title="显示/隐藏批注面板" />
      </div>

      {/* Row 2 */}
      <div
        class="flex items-center gap-1 flex-wrap px-3.5 py-2.5 bg-paper border-b border-line sticky top-0 z-[5]"
        role="toolbar"
        aria-label="字体与段落工具栏"
      >
        <select
          class="h-[30px] border border-line-strong bg-paper text-ink-1 text-[13px] rounded-md px-1.5 cursor-pointer max-w-[130px] hover:border-accent-soft"
          value={currentFontFamily()}
          onChange={(e) => onFontFamilyChange(e.currentTarget.value)}
          title="字体"
        >
          <For each={FONT_FAMILIES}>
            {(f) => <option value={f.value}>{f.label}</option>}
          </For>
        </select>
        <select
          class="h-[30px] border border-line-strong bg-paper text-ink-1 text-[13px] rounded-md px-1.5 cursor-pointer max-w-[78px] hover:border-accent-soft"
          value={currentFontSizePt()}
          onChange={(e) => onFontSizeChange(e.currentTarget.value)}
          title="字号"
        >
          <option value="">字号</option>
          <For each={FONT_SIZES_PT}>
            {(pt) => <option value={pt}>{pt}</option>}
          </For>
        </select>
        <Btn id="bold" label="B" onClick={() => runCommand(toolbar.bold)} active={() => isMarkActive(docSchema.marks.strong)} title="加粗" />
        <Btn id="italic" label="I" onClick={() => runCommand(toolbar.italic)} active={() => isMarkActive(docSchema.marks.em)} title="斜体" />
        <Btn id="underline" label="U" onClick={() => runCommand(toolbar.underline)} active={() => isMarkActive(docSchema.marks.underline)} title="下划线" />
        <Btn id="strike" label="S" onClick={() => runCommand(toolbar.strike)} active={() => isMarkActive(docSchema.marks.strike)} title="删除线" />
        <Btn id="code" label="&lt;/&gt;" onClick={() => runCommand(toolbar.code)} active={() => isMarkActive(docSchema.marks.code)} title="行内代码" />
        <label class="flex items-center gap-1 text-xs font-bold text-ink-2 cursor-pointer px-1" title="文字颜色">
          A
          <input
            type="color"
            class="w-[22px] h-[22px] p-0 border border-line-strong rounded cursor-pointer bg-transparent"
            value={currentColor()}
            onInput={(e) => onColorChange(e.currentTarget.value)}
          />
        </label>
        <select
          class="h-[30px] border border-line-strong bg-paper text-ink-1 text-[13px] rounded-md px-1.5 cursor-pointer max-w-[78px] hover:border-accent-soft"
          value={currentHighlight()}
          onChange={(e) => onHighlightChange(e.currentTarget.value)}
          title="高亮"
        >
          <For each={HIGHLIGHT_OPTIONS}>
            {(h) => <option value={h.value}>{h.label}</option>}
          </For>
        </select>
        <span class="w-px h-5 bg-line mx-1" />
        <Btn id="align-left" label="左对齐" onClick={() => runCommand(toolbar.alignLeft)} active={() => currentAlign() === 'left'} title="左对齐" />
        <Btn id="align-center" label="居中" onClick={() => runCommand(toolbar.alignCenter)} active={() => currentAlign() === 'center'} title="居中" />
        <Btn id="align-right" label="右对齐" onClick={() => runCommand(toolbar.alignRight)} active={() => currentAlign() === 'right'} title="右对齐" />
        <Btn id="align-justify" label="两端对齐" onClick={() => runCommand(toolbar.alignJustify)} active={() => currentAlign() === 'justify'} title="两端对齐" />
        <span class="w-px h-5 bg-line mx-1" />
        <Btn id="indent-less" label="减少缩进" onClick={() => runCommand(toolbar.indentLess)} title="减少缩进" />
        <Btn id="indent-more" label="增加缩进" onClick={() => runCommand(toolbar.indentMore)} title="增加缩进" />
        <select
          class="h-[30px] border border-line-strong bg-paper text-ink-1 text-[13px] rounded-md px-1.5 cursor-pointer max-w-[78px] hover:border-accent-soft"
          value={currentLineSpacing()}
          onChange={(e) => onLineSpacingChange(e.currentTarget.value)}
          title="行距"
        >
          <For each={LINE_SPACING_OPTIONS}>
            {(l) => <option value={l.value}>{l.label}</option>}
          </For>
        </select>
        <span class="w-px h-5 bg-line mx-1" />
        <Btn id="clear-format" label="清除格式" onClick={() => runCommand(toolbar.clearFormatting)} title="清除字符与段落格式（不影响批注/链接）" />
      </div>

      {/* Row 3 (table) */}
      <Show when={inTable()}>
        <div
          class="flex items-center gap-1 flex-wrap px-3.5 py-2.5 bg-accent-wash border-b border-accent-soft sticky top-0 z-[5]"
          role="toolbar"
          aria-label="表格工具栏"
        >
          <span class="text-xs text-ink-3 whitespace-nowrap mr-0.5">表格：</span>
          <Btn id="row-before" label="上方插入行" onClick={() => runCommand(tableToolbar.addRowBefore)} />
          <Btn id="row-after" label="下方插入行" onClick={() => runCommand(tableToolbar.addRowAfter)} />
          <Btn id="col-before" label="左侧插入列" onClick={() => runCommand(tableToolbar.addColumnBefore)} />
          <Btn id="col-after" label="右侧插入列" onClick={() => runCommand(tableToolbar.addColumnAfter)} />
          <span class="w-px h-5 bg-line mx-1" />
          <Btn id="del-row" label="删除行" onClick={() => runCommand(tableToolbar.deleteRow)} />
          <Btn id="del-col" label="删除列" onClick={() => runCommand(tableToolbar.deleteColumn)} />
          <Btn id="del-table" label="删除表格" onClick={() => runCommand(tableToolbar.deleteTable)} />
          <span class="w-px h-5 bg-line mx-1" />
          <Btn id="merge" label="合并单元格" onClick={() => runCommand(tableToolbar.mergeCells)} />
          <Btn id="split" label="拆分单元格" onClick={() => runCommand(tableToolbar.splitCell)} />
          <Btn id="header-row" label="表头行" onClick={() => runCommand(tableToolbar.toggleHeaderRow)} />
          <Btn id="header-col" label="表头列" onClick={() => runCommand(tableToolbar.toggleHeaderColumn)} />
          <label class="flex items-center gap-1 text-xs font-bold text-ink-2 cursor-pointer px-1" title="单元格底色">
            底色
            <input
              type="color"
              class="w-[22px] h-[22px] p-0 border border-line-strong rounded cursor-pointer bg-transparent"
              value="#ffffff"
              onInput={(e) =>
                runCommand(
                  tableToolbar.setCellBackground(e.currentTarget.value)
                )
              }
            />
          </label>
          <span class="w-px h-5 bg-line mx-1" />
          <Btn id="table-props" label="表格属性" onClick={() => setShowTableProps(true)} title="表格对齐 / 单元格对齐方向 / 文字方向 / 边框底纹" />
        </div>
      </Show>

      <Show when={showTableProps()}>
        <TablePropertiesModal
          initial={tablePropsInitial()}
          onClose={() => setShowTableProps(false)}
          onApply={applyTableProperties}
        />
      </Show>

      <div class="flex-1 min-h-0 flex overflow-hidden">
        <Show when={showOutline()}>
          <OutlineTree items={outlineItems()} onJump={scrollToBlock} />
        </Show>
        <div class="flex-1 min-w-0 overflow-y-auto px-6 pb-24 pt-10 bg-canvas">
          <div class="max-w-[760px] min-h-[900px] mx-auto bg-paper shadow-[0_1px_2px_rgba(23,26,33,0.06),0_12px_32px_rgba(23,26,33,0.08)] rounded-[3px] px-[76px] py-[72px] border-l-[3px] border-l-accent editor-page" ref={hostEl} />
        </div>
        <Show when={showComments()}>
          <CommentsPanel comments={comments()} onJump={jumpToComment} />
        </Show>
      </div>
    </div>
  );
}
