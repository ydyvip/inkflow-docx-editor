import { TableView } from 'prosemirror-tables';
import type { Node as PMNode } from 'prosemirror-model';

/**
 * prosemirror-tables 的 columnResizing() 插件会给 table 节点安装一个自定义
 * NodeView（TableView），直接用 document.createElement 手搭 DOM 来管理列宽拖拽，
 * 完全绕过了 schema 里 table 节点的 toDOM——这意味着我们在 schema/index.ts 里
 * 给 table 加的 align 属性渲染逻辑实际上从未被调用。
 *
 * 这里继承 TableView，在它自己的列宽计算跑完之后，再补上 align 对应的样式。
 * 同时：只有当表格不是撑满 100% 宽度时，居中/靠右才有视觉意义，所以对齐生效时
 * 顺带让外层 wrapper 收缩到内容宽度（width: fit-content）。
 */
export class AlignAwareTableView extends TableView {
  constructor(node: PMNode, defaultCellMinWidth: number) {
    super(node, defaultCellMinWidth);
    this.applyAlign(node);
  }

  private applyAlign(node: PMNode) {
    const align = node.attrs.align as string | null;
    if (align === 'center' || align === 'right') {
      this.dom.style.width = 'fit-content';
      this.dom.style.marginLeft = 'auto';
      this.dom.style.marginRight = align === 'center' ? 'auto' : '0';
    } else {
      this.dom.style.width = '';
      this.dom.style.marginLeft = '';
      this.dom.style.marginRight = '';
    }
  }

  update(node: PMNode): boolean {
    const result = super.update(node);
    if (result) this.applyAlign(node);
    return result;
  }
}
