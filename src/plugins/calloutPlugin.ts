/**
 * 示例插件：Callout（高亮提示块）
 * 演示插件系统四种扩展点的完整用法（§7.2）。
 *
 * 触发方式：在行首输入 ":::" + 空格
 * 快捷键：Mod-Shift-9
 */
import type { Schema } from "prosemirror-model";
import { wrappingInputRule } from "prosemirror-inputrules";
import { wrapIn, lift } from "prosemirror-commands";
import type { DocxPlugin } from "./registry";

export const calloutPlugin: DocxPlugin = {
  name: "callout",

  nodes: {
    callout: {
      content: "paragraph+",
      group: "block",
      defining: true,
      attrs: { tone: { default: "info" } },
      parseDOM: [
        {
          tag: 'div[data-type="callout"]',
          getAttrs: (dom) => ({
            tone: (dom as HTMLElement).getAttribute("data-tone") || "info",
          }),
        },
      ],
      toDOM(node) {
        return [
          "div",
          { "data-type": "callout", "data-tone": node.attrs.tone, class: `callout callout-${node.attrs.tone}` },
          0,
        ];
      },
    },
  },

  inputRules: (schema: Schema) => [
    wrappingInputRule(/^:::\s$/, schema.nodes.callout),
  ],

  keymap: (schema: Schema) => ({
    "Mod-Shift-9": wrapIn(schema.nodes.callout),
  }),

  toolbar: (schema: Schema) => [
    {
      id: "callout",
      label: "💡 提示块",
      run: (view) => {
        const { state, dispatch } = view;
        const inCallout = (state.selection.$from as any).node(-1)?.type === schema.nodes.callout;
        if (inCallout) {
          lift(state, dispatch);
        } else {
          wrapIn(schema.nodes.callout)(state, dispatch);
        }
      },
      isActive: (state) => {
        const $from = state.selection.$from;
        for (let d = $from.depth; d > 0; d--) {
          if ($from.node(d).type === schema.nodes.callout) return true;
        }
        return false;
      },
    },
  ],

  // 导出映射：告诉 Export 模块如何把 callout 节点转成 docx 元素
  exportNode: (node, _toDocxRuns) => {
    return { __calloutTone: node.attrs?.tone ?? "info", paragraphs: node.content ?? [] };
  },
};
