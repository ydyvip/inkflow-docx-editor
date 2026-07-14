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
  currentListInfo,
  currentTableNode,
  setListStyle,
  setLineSpacing,
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
import { OutlineTree } from '../outline/OutlineTree';
import {
  computeOutline,
  findNodeById,
  scrollPosIntoView,
} from '../outline/computeOutline';
import { CommentsPanel, type CommentItem } from '../comments/CommentsPanel';
import {
  HEADING_LEVELS,
  LIST_STYLE_OPTIONS,
  LINE_SPACING_OPTIONS,
} from './formatOptions';
import { FloatingToolbar } from './FloatingToolbar';

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
          (m) => m.type.name === 'comment' && Number(m.attrs.id) === Number(commentId)
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

  const isLinkActive = () => {
    version();
    return view ? markActive(view.state, docSchema.marks.link) : false;
  };

  const outlineItems = () => {
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

  const toolbar = buildToolbar(docSchema);
  const tableToolbar = buildTableToolbar();

  const onHeadingSelect = (value: string) => {
    if (value === 'paragraph') runCommand(toolbar.paragraph);
    else runCommand(toolbar.heading(Number(value.slice(1))));
  };

  const onListStyleChange = (value: string) => {
    if (value === 'none') {
      runCommand(setListStyle(docSchema, 'none'));
      return;
    }
    const [kind, style] = value.split(':') as ['bullet' | 'ordered', string];
    runCommand(setListStyle(docSchema, kind, style));
  };

  const onLineSpacingChange = (value: string) => {
    runCommand(setLineSpacing(value ? Number(value) : null));
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

  // ---- 清除格式：保留字体(fontFamily)和字号(sizeHalfPt)，其余格式清除 ----
  const handleClearFormat = () => {
    if (!view) return;
    const { from, to } = view.state.selection;
    const tr = view.state.tr;

    // 1) 移除非 docxStyle 的字符格式 marks
    ['strong', 'em', 'underline', 'strike', 'code'].forEach((m) => {
      const mt = docSchema.marks[m];
      if (mt) tr.removeMark(from, to, mt);
    });

    // 2) 对 docxStyle mark：保留字体与字号，移除颜色与高亮
    const docxStyleMark = docSchema.marks.docxStyle;
    if (docxStyleMark) {
      tr.removeMark(from, to, docxStyleMark);
      const preserved: Record<string, any> = {};
      view.state.doc.nodesBetween(from, to, (node) => {
        if (!node.isText) return true;
        const existing = node.marks.find((m) => m.type === docxStyleMark);
        if (existing) {
          if (existing.attrs.fontFamily && !preserved.fontFamily) {
            preserved.fontFamily = existing.attrs.fontFamily;
          }
          if (existing.attrs.sizeHalfPt && !preserved.sizeHalfPt) {
            preserved.sizeHalfPt = existing.attrs.sizeHalfPt;
          }
        }
        return true;
      });

      if (Object.keys(preserved).length > 0) {
        tr.addMark(from, to, docxStyleMark.create(preserved));
      }
    }

    // 3) 重置段落级格式（对齐、缩进、行距）
    view.state.doc.nodesBetween(from, to, (node, pos) => {
      if ((node.type.isTextblock && node.type.name !== 'heading') || node.type.name === 'heading') {
        tr.setNodeMarkup(pos, null, {
          ...node.attrs,
          align: undefined,
          indent: undefined,
          lineSpacing: undefined,
        });
      }
      return true;
    });

    view.dispatch(tr);
    view.focus();
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
      {/* Row 1：结构 —— 标题级别 / 列表 / 引用 / 插入 / 面板开关 */}
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
        <span class="flex-1" />
        <Btn id="toggle-outline" label="目录" onClick={() => setShowOutline((v) => !v)} active={() => showOutline()} title="显示/隐藏文档目录" />
        <Btn id="toggle-comments" label="批注列表" onClick={() => setShowComments((v) => !v)} active={() => showComments()} title="显示/隐藏批注面板" />
      </div>

      {/* Row 2：段落格式 —— 对齐 / 缩进 / 行距 / 清除格式 */}
      <div
        class="flex items-center gap-1 flex-wrap px-3.5 py-2.5 bg-paper border-b border-line sticky top-0 z-[5]"
        role="toolbar"
        aria-label="段落工具栏"
      >
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
        <Btn id="clear-format" label="清除格式" onClick={handleClearFormat} title="清除字符与段落格式（不影响批注/链接）" />
      </div>

      {/* Row 3（条件显示）：表格上下文工具栏 */}
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

      <FloatingToolbar
        view={() => view}
        schema={() => docSchema}
        onAddComment={addCommentOnRange}
        showFontControls={true}
      />

      <div class="flex-1 min-h-0 flex overflow-hidden">
        <Show when={showOutline()}>
          <OutlineTree items={outlineItems()} onJump={scrollToBlock} />
        </Show>
        <div class="flex-1 min-w-0 overflow-y-auto px-6 pb-24 pt-10 bg-canvas">
          <div class="max-w-[760px] min-h-[900px] mx-auto bg-paper shadow-[0_1px_2px_rgba(23,26,33,0.06),0_12px_32px_rgba(23,26,33,0.08)] rounded-[3px] px-[76px] py-[72px] border-l-[3px] border-l-accent editor-page" ref={hostEl} />
        </div>
        <Show when={showComments()}>
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
            onDelete={(id) => {
              if (!view) return;

              // 1) 更新批注列表状态
              const updated = comments().filter((c) => c.id !== id);
              setComments(updated);
              props.onCommentsChange?.(updated);

              // 2) 从正文中移除该 id 对应的所有 comment mark（即取消高亮）
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
                // 同时清除因点击批注而产生的精确选区高亮
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
