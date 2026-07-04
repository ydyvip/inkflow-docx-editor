/**
 * 插件系统 —— 注册中心
 * ------------------------------------------------------------
 * 设计约束（对应方案 §7 插件系统设计）：
 *   - Node plugin  : 向 schema 注入新的块/行内节点
 *   - Mark plugin   : 向 schema 注入新的 mark
 *   - Input rule    : 输入触发（如 ":::" 触发 callout）
 *   - Keymap        : 快捷键扩展
 *
 * ProseMirror 的 schema 在 EditorState 创建后不可变，因此插件的
 * node/mark 注册必须发生在 schema 构建之前（应用启动时一次性完成）。
 * 行为类插件（inputRules / keymap / ProseMirror plugin）在 schema
 * 确定之后统一收集，传入 EditorState.create({ plugins }).
 *
 * 每个插件还可以声明 docx 导出映射（export()），从而让 Export
 * 模块无需硬编码每个自定义节点的处理逻辑（§6.5 / §11 扩展点）。
 */
import type { NodeSpec, MarkSpec, Schema } from "prosemirror-model";
import type { Plugin } from "prosemirror-state";
import type { InputRule } from "prosemirror-inputrules";

export interface ToolbarItem {
  id: string;
  label: string;
  /** 在当前 view 中执行的命令 */
  run: (view: { state: any; dispatch: any }) => void;
  isActive?: (state: any) => boolean;
}

export interface DocxPlugin {
  name: string;
  /** Node plugin */
  nodes?: Record<string, NodeSpec>;
  /** Mark plugin */
  marks?: Record<string, MarkSpec>;
  /** Input rule 工厂（需要 schema 才能创建，故为函数） */
  inputRules?: (schema: Schema) => InputRule[];
  /** Keymap 工厂 */
  keymap?: (schema: Schema) => Record<string, any>;
  /** 依赖 schema 的其它 ProseMirror 插件（历史记录、装饰等） */
  proseMirrorPlugins?: (schema: Schema) => Plugin[];
  /** 工具栏按钮（编辑器模块渲染） */
  toolbar?: (schema: Schema) => ToolbarItem[];
  /**
   * 导出扩展点：告诉 Export 模块如何把该自定义节点转换成 docx.js 元素。
   * 返回 null 表示交给默认规则处理。
   */
  exportNode?: (node: any, toDocxRuns: (content: any[]) => any[]) => any | null;
}

class PluginRegistry {
  private plugins: DocxPlugin[] = [];

  register(plugin: DocxPlugin) {
    if (this.plugins.some((p) => p.name === plugin.name)) return;
    this.plugins.push(plugin);
  }

  all(): DocxPlugin[] {
    return this.plugins;
  }

  nodes(): Record<string, NodeSpec> {
    return this.plugins.reduce((acc, p) => ({ ...acc, ...(p.nodes ?? {}) }), {});
  }

  marks(): Record<string, MarkSpec> {
    return this.plugins.reduce((acc, p) => ({ ...acc, ...(p.marks ?? {}) }), {});
  }
}

export const pluginRegistry = new PluginRegistry();
