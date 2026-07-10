import { createSignal, For } from 'solid-js';
import './tableProperties.css';

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
 * 表格属性面板（对应 Word 的"表格属性"对话框，做了简化但覆盖核心项）：
 *   - 表格对齐：整张表格在页面里靠左/居中/靠右
 *   - 单元格对齐方向：选中单元格内容的垂直对齐（顶端/居中/底端）
 *   - 文字方向：选中单元格内文字是水平还是垂直排列
 *   - 边框：线型 + 磅值 + 颜色，应用到选中单元格
 *   - 底纹：选中单元格的背景色
 *
 * 表单状态本地维护，点"应用"时把收集到的值一次性交给调用方（EditorPane），
 * 具体落到 ProseMirror 的哪些 command 由调用方决定——这个组件不感知 ProseMirror。
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
    <div class="table-props-backdrop" onClick={props.onClose}>
      <div class="table-props-modal" onClick={(e) => e.stopPropagation()}>
        <div class="table-props-header">
          <span>表格属性</span>
          <button
            type="button"
            class="table-props-close"
            onClick={props.onClose}
            title="关闭"
          >
            ✕
          </button>
        </div>

        <div class="table-props-section">
          <div class="table-props-label">表格对齐</div>
          <div class="table-props-radio-row">
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
                <label class="table-props-radio">
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

        <div class="table-props-section">
          <div class="table-props-label">单元格对齐方向（垂直）</div>
          <div class="table-props-radio-row">
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
                <label class="table-props-radio">
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

        <div class="table-props-section">
          <div class="table-props-label">文字方向</div>
          <div class="table-props-radio-row">
            <For
              each={
                [
                  ['horizontal', '水平'],
                  ['vertical', '垂直'],
                ] as const
              }
            >
              {([value, label]) => (
                <label class="table-props-radio">
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

        <div class="table-props-section">
          <div class="table-props-label">边框（应用到选中单元格）</div>
          <div class="table-props-row">
            <select
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
              class="table-props-number"
              value={borderWidth()}
              onInput={(e) =>
                setBorderWidth(Number(e.currentTarget.value) || 1)
              }
              title="边框粗细（磅）"
            />
            <span class="table-props-unit">磅</span>
            <input
              type="color"
              value={borderColor()}
              onInput={(e) => setBorderColor(e.currentTarget.value)}
              title="边框颜色"
            />
          </div>
        </div>

        <div class="table-props-section">
          <div class="table-props-label">底纹（选中单元格背景色）</div>
          <div class="table-props-row">
            <input
              type="color"
              value={cellBackground()}
              onInput={(e) => setCellBackground(e.currentTarget.value)}
            />
          </div>
        </div>

        <div class="table-props-footer">
          <button
            type="button"
            class="table-props-btn table-props-btn-ghost"
            onClick={props.onClose}
          >
            取消
          </button>
          <button
            type="button"
            class="table-props-btn table-props-btn-primary"
            onClick={apply}
          >
            应用
          </button>
        </div>
      </div>
    </div>
  );
}
