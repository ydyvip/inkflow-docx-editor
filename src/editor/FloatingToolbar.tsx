import { Show, For, onMount, onCleanup, createSignal, createMemo } from 'solid-js';
import { EditorView } from 'prosemirror-view';
import type { Schema, MarkType } from 'prosemirror-model';
import {
  FONT_FAMILIES,
  FONT_SIZES_PT,
  HIGHLIGHT_OPTIONS,
} from './formatOptions';

export interface FloatingToolbarProps {
  view: () => EditorView | undefined;
  schema: () => Schema | undefined;
  onAddComment: (from: number, to: number) => void;
  /** 是否显示字体编辑控件（编辑模式=true，预览模式=false） */
  showFontControls?: boolean;
}

/** 当前选中文字内容 */
function getSelectedText(v: EditorView): string {
  const { from, to } = v.state.selection;
  return v.state.doc.textBetween(from, to);
}

/** 判断选区内该 mark 是否全部激活 */
function isMarkActive(v: EditorView, markType: MarkType): boolean {
  const { from, to, empty } = v.state.selection;
  if (empty) return false;
  return v.state.doc.rangeHasMark(from, to, markType);
}

/** 从选区文本节点读取 docxStyle 属性
 *  - 全部一致：返回该值（空字符串也算一致）
 *  - 不一致或没有文本节点：返回 '__mixed__'
 */
function getUniformDocxAttr(
  v: EditorView,
  attrName: string
): string {
  const { from, to, empty } = v.state.selection;
  if (empty) return '';

  let value: string | undefined = undefined;
  let first = true;
  let hasText = false;

  v.state.doc.nodesBetween(from, to, (node) => {
    if (!node.isText) return true;
    hasText = true;
    const mark = node.marks.find((m) => m.type.name === 'docxStyle');
    const nodeValue = mark?.attrs?.[attrName] ?? '';
    if (first) {
      value = nodeValue;
      first = false;
    } else if (value !== nodeValue) {
      value = '__mixed__';
      return false;
    }
    return true;
  });

  if (!hasText) return '';
  if (value === '__mixed__') return '__mixed__';
  return value ?? '';
}

const MIXED_VALUE = '__mixed__';

/** 渲染一个支持“混合”状态的下拉框 */
function StyleSelect(props: {
  value: string;
  onChange: (value: string) => void;
  title: string;
  children: any;
  widthClass?: string;
}) {
  return (
    <select
      class={`h-7 border border-line-strong bg-paper text-ink-1 text-xs rounded px-1 cursor-pointer hover:border-accent-soft ${props.widthClass ?? 'max-w-[90px]'}`}
      value={props.value}
      onChange={(e) => {
        const val = e.currentTarget.value;
        if (val === MIXED_VALUE) return;
        props.onChange(val);
      }}
      title={props.title}
    >
      <option value={MIXED_VALUE} disabled hidden>
        —
      </option>
      {props.children}
    </select>
  );
}

export function FloatingToolbar(props: FloatingToolbarProps) {
  const [visible, setVisible] = createSignal(false);
  const [pos, setPos] = createSignal({ x: 0, y: 0 });
  const [selRange, setSelRange] = createSignal<{ from: number; to: number } | null>(null);
  const [copied, setCopied] = createSignal(false);

  const checkSelection = (force = false) => {
    const v = props.view();
    if (!v) {
      setVisible(false);
      return;
    }
    const { from, to, empty } = v.state.selection;
    if (empty) {
      setVisible(false);
      setSelRange(null);
      return;
    }

    const sel = window.getSelection();
    if (!sel || sel.rangeCount === 0) {
      setVisible(false);
      return;
    }

    const domSel = sel.getRangeAt(0);
    const hostEl = v.dom.closest('.editor-page');
    if (hostEl && !hostEl.contains(domSel.commonAncestorContainer)) {
      setVisible(false);
      return;
    }

    const r = domSel.getBoundingClientRect();
    const toolbarWidth = props.showFontControls ? 360 : 140;
    let x = r.left + r.width / 2 - toolbarWidth / 2;
    let y = r.top - 52;

    x = Math.max(8, Math.min(x, window.innerWidth - toolbarWidth - 8));
    if (y < 8) y = r.bottom + 8;

    setPos({ x, y });
    setSelRange({ from, to });
    if (force) setCopied(false);
    setVisible(true);
  };

  onMount(() => {
    const handleMouseUp = (e: MouseEvent) => {
      const toolbarEl = (e.target as HTMLElement)?.closest('[data-floating-toolbar]');
      if (toolbarEl) return;
      setTimeout(() => checkSelection(true), 10);
    };
    const handleKeyUp = (e: KeyboardEvent) => {
      if (['ArrowUp', 'ArrowDown', 'ArrowLeft', 'ArrowRight', 'Shift'].includes(e.key)) {
        setTimeout(() => checkSelection(true), 10);
      }
    };
    const handleSelectionChange = () => {
      const s = window.getSelection();
      if (!s || s.isCollapsed) {
        setVisible(false);
      }
    };

    document.addEventListener('mouseup', handleMouseUp);
    document.addEventListener('keyup', handleKeyUp);
    document.addEventListener('selectionchange', handleSelectionChange);

    onCleanup(() => {
      document.removeEventListener('mouseup', handleMouseUp);
      document.removeEventListener('keyup', handleKeyUp);
      document.removeEventListener('selectionchange', handleSelectionChange);
    });
  });

  const copySelection = () => {
    const v = props.view();
    const r = selRange();
    if (!v || !r) return;
    const text = getSelectedText(v);
    navigator.clipboard.writeText(text).then(() => {
      setCopied(true);
      setTimeout(() => setCopied(false), 1500);
    });
  };

  const addComment = () => {
    const v = props.view();
    const r = selRange();
    if (!v || !r) return;
    setVisible(false);
    props.onAddComment(r.from, r.to);
  };

  // ===== 字体编辑命令（仅编辑模式） =====
  const canEditFont = () => props.showFontControls && props.schema();

  const getDocStyleAttr = (attrName: string) => {
    const v = props.view();
    if (!v) return '';
    return getUniformDocxAttr(v, attrName);
  };

  const currentFontFamily = createMemo(() => getDocStyleAttr('fontFamily'));
  const currentFontSizePt = createMemo(() => {
    const half = getDocStyleAttr('sizeHalfPt');
    return half ? String(Number(half) / 2) : '';
  });
  const currentHighlight = createMemo(() => getDocStyleAttr('highlight'));
  const currentColor = createMemo(() => {
    const v = props.view();
    if (!v) return '#1c2321';
    const c = getUniformDocxAttr(v, 'color');
    if (!c || c === MIXED_VALUE) return '#1c2321';
    return c;
  });

  const setDocStyleAttr = (attrs: Record<string, any>) => {
    const v = props.view();
    const schema = props.schema();
    if (!v || !schema) return;
    const r = selRange();
    if (!r) return;

    const markType = schema.marks.docxStyle;
    if (!markType) return;

    const tr = v.state.tr;
    // 先从选区中移除旧的 docxStyle mark
    tr.removeMark(r.from, r.to, markType);

    // 以最新输入的属性为准；选区内原有节点如果有 docxStyle，则保留其未覆盖的属性
    const valid: Record<string, any> = {};
    v.state.doc.nodesBetween(r.from, r.to, (node) => {
      if (!node.isText) return true;
      const existing = node.marks.find((m) => m.type === markType);
      if (existing) {
        for (const k of Object.keys(existing.attrs)) {
          if (valid[k] === undefined && existing.attrs[k] != null) {
            valid[k] = existing.attrs[k];
          }
        }
      }
      return true;
    });

    for (const [k, val] of Object.entries(attrs)) {
      if (val != null) valid[k] = val;
    }

    // 过滤掉 null/undefined 的值，如果全空则直接移除 mark
    const final: Record<string, any> = {};
    for (const k of Object.keys(valid)) {
      if (valid[k] !== null && valid[k] !== undefined && valid[k] !== '') {
        final[k] = valid[k];
      }
    }

    if (Object.keys(final).length > 0) {
      tr.addMark(r.from, r.to, markType.create(final));
    }

    v.dispatch(tr);
    // 不调用 v.focus()，避免选区丢失导致看不到选区高亮
    // ProseMirror 的 selection 仍保留在 state 中
  };

  const onColorChange = (hex: string) => setDocStyleAttr({ color: hex });
  const onFontFamilyChange = (value: string) => setDocStyleAttr({ fontFamily: value || '' });
  const onFontSizeChange = (ptStr: string) => {
    if (!ptStr) setDocStyleAttr({ sizeHalfPt: '' });
    else setDocStyleAttr({ sizeHalfPt: Math.round(Number(ptStr) * 2) });
  };
  const HIGHLIGHT_NONE = '__none__';
  const onHighlightChange = (value: string) =>
    setDocStyleAttr({ highlight: value === HIGHLIGHT_NONE ? '' : value || '' });

  const isActive = (markName: string) => {
    const v = props.view();
    const schema = props.schema();
    if (!v || !schema) return false;
    const markType = (schema.marks as any)[markName];
    if (!markType) return false;
    return isMarkActive(v, markType);
  };

  const toggleMark = (markName: string) => {
    const v = props.view();
    const schema = props.schema();
    if (!v || !schema) return;
    const markType = (schema.marks as any)[markName];
    if (!markType) return;

    const r = selRange();
    if (!r) return;

    const active = isMarkActive(v, markType);
    const tr = v.state.tr;
    if (active) {
      tr.removeMark(r.from, r.to, markType);
    } else {
      tr.addMark(r.from, r.to, markType.create());
    }
    v.dispatch(tr);
  };

  return (
    <Show when={visible()}>
      <div
        class="fixed z-50 flex items-center gap-1.5 bg-paper rounded-lg shadow-lg border border-line px-2 py-1.5 select-none"
        data-floating-toolbar
        style={{
          left: `${pos().x}px`,
          top: `${pos().y}px`,
          'max-width': '90vw',
          'flex-wrap': 'wrap',
        }}
        onMouseDown={(e) => {
          const tag = (e.target as HTMLElement)?.tagName;
          if (tag !== 'SELECT' && tag !== 'INPUT' && tag !== 'OPTION' && tag !== 'TEXTAREA') {
            e.preventDefault();
          }
        }}
      >
        {/* 字体编辑控件 - 仅编辑模式 */}
        <Show when={canEditFont()}>
          <StyleSelect
            value={currentFontFamily()}
            onChange={onFontFamilyChange}
            title="字体"
          >
            <For each={FONT_FAMILIES}>
              {(f) => <option value={f.value}>{f.label}</option>}
            </For>
          </StyleSelect>
          <StyleSelect
            value={currentFontSizePt()}
            onChange={onFontSizeChange}
            title="字号"
            widthClass="max-w-[56px]"
          >
            <option value="">字号</option>
            <For each={FONT_SIZES_PT}>
              {(pt) => <option value={pt}>{pt}</option>}
            </For>
          </StyleSelect>
          <button
            type="button"
            class={`h-7 px-2 text-xs font-semibold rounded border border-transparent hover:bg-surface-2 ${isActive('strong') ? 'bg-accent-wash text-accent-ink border-accent-soft' : 'text-ink-2'}`}
            onMouseDown={(e) => { e.preventDefault(); toggleMark('strong'); }}
            title="加粗"
          >
            B
          </button>
          <button
            type="button"
            class={`h-7 px-2 text-xs font-semibold rounded border border-transparent hover:bg-surface-2 ${isActive('em') ? 'bg-accent-wash text-accent-ink border-accent-soft' : 'text-ink-2'}`}
            onMouseDown={(e) => { e.preventDefault(); toggleMark('em'); }}
            title="斜体"
          >
            I
          </button>
          <button
            type="button"
            class={`h-7 px-2 text-xs font-semibold rounded border border-transparent hover:bg-surface-2 ${isActive('underline') ? 'bg-accent-wash text-accent-ink border-accent-soft' : 'text-ink-2'}`}
            onMouseDown={(e) => { e.preventDefault(); toggleMark('underline'); }}
            title="下划线"
          >
            U
          </button>
          <button
            type="button"
            class={`h-7 px-2 text-xs font-semibold rounded border border-transparent hover:bg-surface-2 ${isActive('strike') ? 'bg-accent-wash text-accent-ink border-accent-soft' : 'text-ink-2'}`}
            onMouseDown={(e) => { e.preventDefault(); toggleMark('strike'); }}
            title="删除线"
          >
            S
          </button>
          <button
            type="button"
            class={`h-7 px-2 text-xs font-semibold rounded border border-transparent hover:bg-surface-2 ${isActive('code') ? 'bg-accent-wash text-accent-ink border-accent-soft' : 'text-ink-2'}`}
            onMouseDown={(e) => { e.preventDefault(); toggleMark('code'); }}
            title="行内代码"
          >
            &lt;/&gt;
          </button>
          <label class="flex items-center gap-1 text-xs font-bold text-ink-2 cursor-pointer px-0.5" title="文字颜色">
            A
            <input
              type="color"
              class="w-5 h-5 p-0 border border-line-strong rounded cursor-pointer bg-transparent"
              value={currentColor()}
              onInput={(e) => onColorChange(e.currentTarget.value)}
              onMouseDown={(e) => e.stopPropagation()}
            />
          </label>
          <StyleSelect
            value={currentHighlight()}
            onChange={onHighlightChange}
            title="高亮"
            widthClass="max-w-[56px]"
          >
            <option value="" disabled>
              高亮
            </option>
            <option value={HIGHLIGHT_NONE}>无</option>
            <For each={HIGHLIGHT_OPTIONS.slice(1)}>
              {(h) => <option value={h.value}>{h.label}</option>}
            </For>
          </StyleSelect>
          <span class="w-px h-5 bg-line mx-0.5" />
        </Show>

        {/* 公共操作：复制 + 批注 */}
        <button
          type="button"
          class={`h-7 px-2.5 text-xs font-semibold rounded border transition-all ${copied() ? 'border-accent-soft bg-accent-wash text-accent-ink' : 'border-line-strong bg-paper text-ink-2 hover:bg-accent-wash hover:text-accent-ink hover:border-accent-soft'}`}
          onMouseDown={(e) => { e.preventDefault(); copySelection(); }}
          title="复制选中文字"
        >
          {copied() ? '已复制' : '复制'}
        </button>
        <button
          type="button"
          class="h-7 px-2.5 text-xs font-semibold rounded border border-accent bg-accent text-white hover:bg-accent-ink hover:border-accent-ink transition-all"
          onMouseDown={(e) => { e.preventDefault(); addComment(); }}
          title="添加批注"
        >
          批注
        </button>
      </div>
    </Show>
  );
}
