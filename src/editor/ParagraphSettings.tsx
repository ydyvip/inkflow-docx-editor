import { Show, For, createSignal, createMemo, createEffect, onMount, onCleanup, untrack } from 'solid-js';
import type { EditorView } from 'prosemirror-view';
import {
  buildToolbar,
  currentBlockAttr,
  setLineSpacing,
} from './commands';
import {
  HEADING_LEVELS,
  LINE_SPACING_OPTIONS,
} from './formatOptions';

interface ParagraphSettingsProps {
  view: () => EditorView | undefined;
  version: () => number;
}

interface AnchorInfo {
  left: number;
  top: number;
}

/** 判断当前选区所在的 block 是否受段落设置组件支持 */
function isSupportedBlock(view: EditorView | undefined) {
  if (!view) return false;
  const node = view.state.selection.$from.parent;
  return node.type.name === 'paragraph' || node.type.name === 'heading';
}

export function ParagraphSettings(props: ParagraphSettingsProps) {
  const [open, setOpen] = createSignal(false);
  const [scrollTick, setScrollTick] = createSignal(0);
  let btnEl: HTMLButtonElement | undefined;
  let popEl: HTMLDivElement | undefined;

  const anchor = createMemo<AnchorInfo | null>(() => {
    // 观测 version 与 scrollTick，任何 transaction 或滚动都重新计算位置
    props.version();
    scrollTick();
    const v = props.view();
    if (!v || !isSupportedBlock(v)) return null;

    const $pos = v.state.selection.$from;
    const blockStart = $pos.start();

    try {
      const coords = v.coordsAtPos(blockStart);
      const pageContainer = v.dom.closest('.editor-page') as HTMLElement | null;
      const pageRect = pageContainer?.getBoundingClientRect();

      const gap = 6;
      const buttonSize = 22;
      const left = (pageRect?.left ?? coords.left) - buttonSize - gap;
      const top = coords.top + 2;

      return { left, top };
    } catch {
      return null;
    }
  });

  const blockKind = createMemo(() => {
    props.version();
    const v = props.view();
    if (!v) return 'paragraph' as const;
    const node = v.state.selection.$from.parent;
    if (node.type.name === 'heading') return 'heading' as const;
    return 'paragraph' as const;
  });

  const currentHeadingValue = createMemo(() => {
    props.version();
    const v = props.view();
    if (!v) return 'paragraph';
    const node = v.state.selection.$from.parent;
    if (node.type.name === 'heading') return `h${node.attrs.level}`;
    return 'paragraph';
  });

  const currentAlign = createMemo(() => {
    props.version();
    const v = props.view();
    return v ? currentBlockAttr(v.state, 'align') || 'left' : 'left';
  });

  const currentLineSpacing = createMemo(() => {
    props.version();
    const v = props.view();
    return v ? String(currentBlockAttr(v.state, 'lineSpacing') ?? '') : '';
  });

  const runCommand = (cmd: (state: any, dispatch: any) => boolean) => {
    const v = props.view();
    if (!v) return;
    cmd(v.state, v.dispatch);
    v.focus();
  };

  const getToolbar = () => {
    const v = props.view();
    return v ? buildToolbar(v.state.schema) : null;
  };

  const onHeadingSelect = (value: string) => {
    const t = getToolbar();
    if (!t) return;
    if (value === 'paragraph') runCommand(t.paragraph);
    else runCommand(t.heading(Number(value.slice(1))));
  };

  const onLineSpacingChange = (value: string) => {
    runCommand(setLineSpacing(value ? Number(value) : null));
  };

  const handleAlign = (align: 'left' | 'center' | 'right' | 'justify') => {
    const t = getToolbar();
    if (!t) return;
    const map: Record<string, any> = {
      left: t.alignLeft,
      center: t.alignCenter,
      right: t.alignRight,
      justify: t.alignJustify,
    };
    runCommand(map[align]);
  };

  const handleIndentLess = () => {
    const t = getToolbar();
    if (!t) return;
    runCommand(t.indentLess);
  };

  const handleIndentMore = () => {
    const t = getToolbar();
    if (!t) return;
    runCommand(t.indentMore);
  };

  createEffect(() => {
    // 选区或内容变化时自动关闭 popover（untrack open 避免把 open 自身也当依赖）
    props.version();
    if (untrack(open)) setOpen(false);
  });

  onMount(() => {
    const closeIfOutside = (e: MouseEvent) => {
      const target = e.target as Node;
      if (
        open() &&
        btnEl &&
        popEl &&
        !btnEl.contains(target) &&
        !popEl.contains(target)
      ) {
        setOpen(false);
      }
    };

    document.addEventListener('mousedown', closeIfOutside);
    onCleanup(() => document.removeEventListener('mousedown', closeIfOutside));
  });

  // 监听编辑器滚动容器：滚动时关闭
  createEffect(() => {
    const v = props.view();
    if (!v) return;
    const scroller = v.dom.closest('.overflow-y-auto') as HTMLElement | null;
    if (!scroller) return;

    const handler = () => {
      setScrollTick((n) => n + 1);
      if (open()) setOpen(false);
    };
    scroller.addEventListener('scroll', handler, { passive: true });
    onCleanup(() => scroller.removeEventListener('scroll', handler));
  });

  // 也监听 window 滚动
  onMount(() => {
    const handler = () => {
      setScrollTick((n) => n + 1);
      if (open()) setOpen(false);
    };
    window.addEventListener('scroll', handler, { passive: true });
    onCleanup(() => window.removeEventListener('scroll', handler));
  });

  return (
    <Show when={anchor()}>
      {(pos) => (
        <>
          <button
            ref={btnEl}
            type="button"
            class={`fixed z-40 w-[22px] h-[22px] flex items-center justify-center rounded border text-[11px] font-bold leading-none transition-colors cursor-pointer select-none hover:bg-surface-2 ${
              blockKind() === 'heading'
                ? 'bg-accent-wash text-accent-ink border-accent-soft'
                : 'bg-paper text-ink-2 border-line'
            }`}
            style={{
              left: `${pos().left}px`,
              top: `${pos().top}px`,
            }}
            title={blockKind() === 'heading' ? '段落设置（标题）' : '段落设置'}
            onMouseDown={(e) => {
              e.preventDefault();
              setOpen((v) => !v);
            }}
          >
            {blockKind() === 'heading' ? 'H' : 'T'}
          </button>

          <Show when={open()}>
            <div
              ref={popEl}
              class="fixed z-50 min-w-[220px] max-w-[90vw] bg-paper rounded-lg border border-line shadow-lg px-2.5 py-2 select-none"
              style={{
                left: `${pos().left}px`,
                top: `${pos().top + 28}px`,
              }}
              onMouseDown={(e) => {
                const tag = (e.target as HTMLElement)?.tagName;
                if (tag !== 'SELECT' && tag !== 'INPUT' && tag !== 'OPTION') {
                  e.preventDefault();
                }
              }}
            >
              <div class="flex items-center gap-1.5 mb-2">
                <select
                  class="h-7 flex-1 min-w-0 border border-line-strong bg-paper text-ink-1 text-xs rounded px-1.5 cursor-pointer hover:border-accent-soft font-semibold"
                  value={currentHeadingValue()}
                  onChange={(e) => onHeadingSelect(e.currentTarget.value)}
                  title="段落样式"
                >
                  <option value="paragraph">正文</option>
                  <For each={HEADING_LEVELS}>
                    {(level) => (
                      <option value={`h${level}`}>标题 {level}</option>
                    )}
                  </For>
                </select>
              </div>

              <div class="flex items-center gap-1 mb-2">
                <ParaBtn
                  label="左"
                  active={currentAlign() === 'left'}
                  onClick={() => handleAlign('left')}
                  title="左对齐"
                />
                <ParaBtn
                  label="中"
                  active={currentAlign() === 'center'}
                  onClick={() => handleAlign('center')}
                  title="居中"
                />
                <ParaBtn
                  label="右"
                  active={currentAlign() === 'right'}
                  onClick={() => handleAlign('right')}
                  title="右对齐"
                />
                <ParaBtn
                  label="两端"
                  active={currentAlign() === 'justify'}
                  onClick={() => handleAlign('justify')}
                  title="两端对齐"
                />
              </div>

              <div class="flex items-center gap-1 mb-2">
                <ParaBtn
                  label="减少缩进"
                  onClick={handleIndentLess}
                  title="减少缩进"
                />
                <ParaBtn
                  label="增加缩进"
                  onClick={handleIndentMore}
                  title="增加缩进"
                />
              </div>

              <div class="flex items-center gap-1">
                <select
                  class="h-7 flex-1 min-w-0 border border-line-strong bg-paper text-ink-1 text-xs rounded px-1.5 cursor-pointer hover:border-accent-soft"
                  value={currentLineSpacing()}
                  onChange={(e) => onLineSpacingChange(e.currentTarget.value)}
                  title="行距"
                >
                  <For each={LINE_SPACING_OPTIONS}>
                    {(l) => <option value={l.value}>{l.label}</option>}
                  </For>
                </select>
              </div>
            </div>
          </Show>
        </>
      )}
    </Show>
  );
}

function ParaBtn(props: {
  label: string;
  onClick: () => void;
  active?: boolean;
  title: string;
}) {
  return (
    <button
      type="button"
      class={`flex-1 h-7 px-1.5 text-xs font-semibold rounded border transition-colors cursor-pointer whitespace-nowrap ${
        props.active
          ? 'bg-accent-wash text-accent-ink border-accent-soft'
          : 'bg-paper text-ink-2 border-line-strong hover:bg-surface-2 hover:text-ink-1'
      }`}
      title={props.title}
      onMouseDown={(e) => {
        e.preventDefault();
        props.onClick();
      }}
    >
      {props.label}
    </button>
  );
}
