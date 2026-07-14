import { createSignal, For } from 'solid-js';

export interface TablePropertiesValues {
  tableAlign: 'left' | 'center' | 'right';
  cellVAlign: 'top' | 'middle' | 'bottom';
  textDirection: 'horizontal' | 'vertical';
  borderStyle: 'none' | 'single' | 'dashed' | 'double';
  borderWidth: number;
  borderColor: string;
  cellBackground: string;
}

interface TablePropertiesModalProps {
  initial: TablePropertiesValues;
  onClose: () => void;
  onApply: (values: TablePropertiesValues) => void;
}

/**
 * 表格属性面板
 */
export function TablePropertiesModal(props: TablePropertiesModalProps) {
  const [tableAlign, setTableAlign] = createSignal(props.initial.tableAlign);
  const [cellVAlign, setCellVAlign] = createSignal(props.initial.cellVAlign);
  const [textDirection, setTextDirection] = createSignal(
    props.initial.textDirection
  );
  const [borderStyle, setBorderStyle] = createSignal(props.initial.borderStyle);
  const [borderWidth, setBorderWidth] = createSignal(props.initial.borderWidth);
  const [borderColor, setBorderColor] = createSignal(props.initial.borderColor);
  const [cellBackground, setCellBackground] = createSignal(
    props.initial.cellBackground
  );

  const apply = () => {
    props.onApply({
      tableAlign: tableAlign(),
      cellVAlign: cellVAlign(),
      textDirection: textDirection(),
      borderStyle: borderStyle(),
      borderWidth: borderWidth(),
      borderColor: borderColor(),
      cellBackground: cellBackground(),
    });
    props.onClose();
  };

  return (
    <div class="fixed inset-0 bg-black/35 flex items-center justify-center z-50" onClick={props.onClose}>
      <div class="w-[360px] max-w-[calc(100vw-40px)] max-h-[calc(100vh-80px)] overflow-y-auto bg-paper rounded-xl shadow-2xl p-5 pb-4" onClick={(e) => e.stopPropagation()}>
        <div class="flex items-center justify-between font-display font-semibold text-base text-ink-1 mb-3.5">
          <span>表格属性</span>
          <button
            type="button"
            class="border-0 bg-transparent cursor-pointer text-ink-3 text-sm px-1.5 py-0.5 rounded-md hover:bg-surface-2 hover:text-ink-1"
            onClick={props.onClose}
            title="关闭"
          >
            关闭
          </button>
        </div>

        <div class="mb-3.5">
          <div class="text-[12.5px] font-semibold text-ink-2 mb-1.5">表格对齐</div>
          <div class="flex gap-3.5">
            <For
              each={
                [
                  ['left', '靠左'],
                  ['center', '居中'],
                  ['right', '靠右'],
                ] as const
              }
            >
              {([value, label]) => (
                <label class="flex items-center gap-1.5 text-[13px] text-ink-1 cursor-pointer">
                  <input
                    type="radio"
                    name="tableAlign"
                    value={value}
                    checked={tableAlign() === value}
                    onChange={() => setTableAlign(value)}
                  />
                  {label}
                </label>
              )}
            </For>
          </div>
        </div>

        <div class="mb-3.5">
          <div class="text-[12.5px] font-semibold text-ink-2 mb-1.5">单元格对齐方向（垂直）</div>
          <div class="flex gap-3.5">
            <For
              each={
                [
                  ['top', '顶端'],
                  ['middle', '居中'],
                  ['bottom', '底端'],
                ] as const
              }
            >
              {([value, label]) => (
                <label class="flex items-center gap-1.5 text-[13px] text-ink-1 cursor-pointer">
                  <input
                    type="radio"
                    name="cellVAlign"
                    value={value}
                    checked={cellVAlign() === value}
                    onChange={() => setCellVAlign(value)}
                  />
                  {label}
                </label>
              )}
            </For>
          </div>
        </div>

        <div class="mb-3.5">
          <div class="text-[12.5px] font-semibold text-ink-2 mb-1.5">文字方向</div>
          <div class="flex gap-3.5">
            <For
              each={
                [
                  ['horizontal', '水平'],
                  ['vertical', '垂直'],
                ] as const
              }
            >
              {([value, label]) => (
                <label class="flex items-center gap-1.5 text-[13px] text-ink-1 cursor-pointer">
                  <input
                    type="radio"
                    name="textDirection"
                    value={value}
                    checked={textDirection() === value}
                    onChange={() => setTextDirection(value)}
                  />
                  {label}
                </label>
              )}
            </For>
          </div>
        </div>

        <div class="mb-3.5">
          <div class="text-[12.5px] font-semibold text-ink-2 mb-1.5">边框（应用到选中单元格）</div>
          <div class="flex items-center gap-2">
            <select
              class="h-[30px] border border-line-strong rounded-md px-1.5 text-[13px] bg-paper hover:border-accent-soft"
              value={borderStyle()}
              onChange={(e) =>
                setBorderStyle(
                  e.currentTarget.value as TablePropertiesValues['borderStyle']
                )
              }
            >
              <option value="none">无边框</option>
              <option value="single">单线</option>
              <option value="dashed">虚线</option>
              <option value="double">双线</option>
            </select>
            <input
              type="number"
              min="0.25"
              max="6"
              step="0.25"
              class="w-14 h-[30px] border border-line-strong rounded-md px-1.5 text-[13px]"
              value={borderWidth()}
              onInput={(e) =>
                setBorderWidth(Number(e.currentTarget.value) || 1)
              }
              title="边框粗细（磅）"
            />
            <span class="text-xs text-ink-3">磅</span>
            <input
              type="color"
              class="w-[30px] h-[30px] p-0 border border-line-strong rounded-md cursor-pointer"
              value={borderColor()}
              onInput={(e) => setBorderColor(e.currentTarget.value)}
              title="边框颜色"
            />
          </div>
        </div>

        <div class="mb-3.5">
          <div class="text-[12.5px] font-semibold text-ink-2 mb-1.5">底纹（选中单元格背景色）</div>
          <div class="flex items-center gap-2">
            <input
              type="color"
              class="w-[30px] h-[30px] p-0 border border-line-strong rounded-md cursor-pointer"
              value={cellBackground()}
              onInput={(e) => setCellBackground(e.currentTarget.value)}
            />
          </div>
        </div>

        <div class="flex justify-end gap-2 mt-4 pt-3.5 border-t border-line">
          <button
            type="button"
            class="px-4 py-2 rounded-lg text-[13px] font-semibold cursor-pointer border border-line-strong bg-paper text-ink-2 hover:bg-surface-2"
            onClick={props.onClose}
          >
            取消
          </button>
          <button
            type="button"
            class="px-4 py-2 rounded-lg text-[13px] font-semibold cursor-pointer border border-accent bg-accent text-white hover:bg-accent-ink hover:border-accent-ink"
            onClick={apply}
          >
            应用
          </button>
        </div>
      </div>
    </div>
  );
}
