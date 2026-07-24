/**
 * Math Plugin - LaTeX formula support
 * ------------------------------------------------------------
 * Provides inline ($...$) and block ($$...$$) LaTeX math formulas:
 *   - Rendering: Uses KaTeX to render LaTeX source into clean formulas
 *   - Editing: Double-click formula to enter source editing mode (textarea), Enter/Esc controls
 *   - Insertion: Toolbar buttons + input rules
 *   - Export: Embeds as OMML via docx.js Math/MathRun into DOCX
 *   - Parsing: Extracts text from <m:t> in DOCX <m:oMath> as LaTeX placeholder
 *
 * All features encapsulated as DocxPlugin, registered via pluginRegistry, non-invasive to core Editor.
 */
import katex from 'katex';
import 'katex/dist/katex.min.css';
import type { DocxPlugin } from './registry';
import type { NodeView, EditorView } from 'prosemirror-view';
import type { Node as PMNode, Schema } from 'prosemirror-model';
import { InputRule } from 'prosemirror-inputrules';
import { Math as DocxMath, MathRun } from 'docx';

/* ============================================================ */
/*  Schema                                                      */
/* ============================================================ */

const mathInlineNode = {
  inline: true,
  group: 'inline',
  atom: true,
  attrs: { latex: { default: '' } },
  parseDOM: [
    {
      tag: 'span[data-type="math_inline"]',
      getAttrs: (dom: HTMLElement) => ({
        latex: dom.getAttribute('data-latex') || '',
      }),
    },
  ],
  toDOM(node: any): [string, Record<string, string>] {
    return [
      'span',
      {
        'data-type': 'math_inline',
        'data-latex': node.attrs.latex,
        class: 'math-inline katex-inline',
      },
    ];
  },
};

const mathBlockNode = {
  group: 'block',
  atom: true,
  attrs: { latex: { default: '' } },
  parseDOM: [
    {
      tag: 'div[data-type="math_block"]',
      getAttrs: (dom: HTMLElement) => ({
        latex: dom.getAttribute('data-latex') || '',
      }),
    },
  ],
  toDOM(node: any): [string, Record<string, string>] {
    return [
      'div',
      {
        'data-type': 'math_block',
        'data-latex': node.attrs.latex,
        class: 'math-block katex-block',
      },
    ];
  },
};

/* ============================================================ */
/*  KaTeX Render                                                */
/* ============================================================ */

function renderMath(el: HTMLElement, latex: string, displayMode: boolean) {
  try {
    katex.render(latex, el, {
      displayMode,
      throwOnError: false,
      strict: false,
      trust: true,
    });
  } catch (_e) {
    el.textContent = displayMode ? `$$${latex}$$` : `$${latex}$`;
    el.style.color = 'var(--danger)';
  }
}

/* ============================================================ */
/*  NodeView (render + double-click edit)                       */
/* ============================================================ */

export class MathNodeView implements NodeView {
  dom: HTMLElement;
  node: PMNode;
  view: EditorView;
  getPos: () => number;
  editing = false;
  textarea: HTMLTextAreaElement | null = null;
  private displayMode: boolean;

  constructor(
    node: PMNode,
    view: EditorView,
    getPos: (() => number | undefined) | boolean,
    displayMode: boolean
  ) {
    this.node = node;
    this.view = view;
    this.getPos = getPos as () => number;
    this.displayMode = displayMode;

    this.dom = document.createElement(this.displayMode ? 'div' : 'span');
    this.dom.className = this.displayMode
      ? 'math-block katex-block'
      : 'math-inline katex-inline';
    this.render();

    this.dom.addEventListener('dblclick', (e) => {
      e.stopPropagation();
      this.startEditing();
    });
  }

  private render() {
    const latex = this.node.attrs.latex as string;
    this.dom.innerHTML = '';
    renderMath(this.dom, latex, this.displayMode);
  }

  private startEditing() {
    if (this.editing) return;
    this.editing = true;
    this.dom.innerHTML = '';
    this.dom.style.display = this.displayMode ? 'block' : 'inline-block';

    this.textarea = document.createElement('textarea');
    this.textarea.value = this.node.attrs.latex;
    this.textarea.rows = this.displayMode ? 4 : 1;
    this.textarea.style.cssText = `
      width: 100%;
      font-family: var(--font-mono);
      font-size: 14px;
      padding: 6px 8px;
      border: 2px solid var(--accent);
      border-radius: 4px;
      outline: none;
      resize: vertical;
      background: var(--paper);
      color: var(--ink-1);
      line-height: 1.5;
    `;

    this.textarea.addEventListener('keydown', (e) => {
      if (e.key === 'Escape') {
        e.preventDefault();
        this.stopEditing();
        return;
      }
      const confirmKey = this.displayMode
        ? (e.metaKey || e.ctrlKey) && e.key === 'Enter'
        : e.key === 'Enter' && !e.shiftKey;
      if (confirmKey) {
        e.preventDefault();
        this.save();
      }
    });

    this.textarea.addEventListener('blur', () => {
      // Prevent race with Escape keydown
      setTimeout(() => {
        if (this.editing) this.save();
      }, 50);
    });

    this.dom.appendChild(this.textarea);
    this.textarea.focus();
    this.textarea.select();
  }

  private save() {
    if (!this.editing) return;
    const newLatex = (this.textarea?.value ?? '').trim();
    this.stopEditing();
    if (newLatex !== this.node.attrs.latex) {
      const pos = this.getPos();
      const tr = this.view.state.tr.setNodeMarkup(pos, undefined, {
        ...this.node.attrs,
        latex: newLatex,
      });
      this.view.dispatch(tr);
    }
  }

  private stopEditing() {
    this.editing = false;
    this.textarea = null;
    this.dom.style.display = '';
    this.render();
  }

  update(node: PMNode): boolean {
    if (node.type.name !== this.node.type.name) return false;
    this.node = node;
    if (!this.editing) this.render();
    return true;
  }

  selectNode() {
    this.dom.classList.add('ProseMirror-selectednode');
  }

  deselectNode() {
    this.dom.classList.remove('ProseMirror-selectednode');
  }

  stopEvent(event: Event): boolean {
    // Allow mousedown so PM can select the atom node
    return event.type !== 'mousedown' && event.type !== 'touchstart';
  }

  ignoreMutation() {
    return true;
  }
}

/* ============================================================ */
/*  InputRules                                                  */
/* ============================================================ */

function mathInputRules(schema: Schema) {
  return [
    // 行内公式：$...$
    new InputRule(/\$([^$]*)\$$/, (state, match, start, end) => {
      const latex = match[1]?.trim();
      if (!latex) return null;
      const tr = state.tr.replaceRangeWith(
        start,
        end,
        schema.nodes.math_inline.create({ latex })
      );
      return tr;
    }),
    // 块级公式：$$...$$
    new InputRule(/^\$\$([^$]*)\$\$$/, (state, match, start, end) => {
      const latex = match[1]?.trim();
      if (!latex) return null;
      const tr = state.tr.replaceRangeWith(
        start,
        end,
        schema.nodes.math_block.create({ latex })
      );
      return tr;
    }),
  ];
}

/* ============================================================ */
/*  Toolbar Commands                                            */
/* ============================================================ */

function insertMathInline(schema: Schema) {
  return (view: { state: any; dispatch: any }) => {
    const latex = window.prompt('输入 LaTeX 公式：', 'E = mc^2');
    if (!latex) return;
    view.dispatch(
      view.state.tr.replaceSelectionWith(
        schema.nodes.math_inline.create({ latex: latex.trim() })
      )
    );
  };
}

function insertMathBlock(schema: Schema) {
  return (view: { state: any; dispatch: any }) => {
    const latex = window.prompt('输入块级 LaTeX 公式：', '\\sum_{i=0}^{n} x_i');
    if (!latex) return;
    view.dispatch(
      view.state.tr.replaceSelectionWith(
        schema.nodes.math_block.create({ latex: latex.trim() })
      )
    );
  };
}

/* ============================================================ */
/*  Plugin Definition                                           */
/* ============================================================ */

export const mathPlugin: DocxPlugin = {
  name: 'math',

  nodes: {
    math_inline: mathInlineNode as any,
    math_block: mathBlockNode as any,
  },

  nodeViews: (_schema: Schema) => ({
    math_inline: (node, view, getPos) =>
      new MathNodeView(node, view, getPos, false),
    math_block: (node, view, getPos) =>
      new MathNodeView(node, view, getPos, true),
  }),

  inputRules: mathInputRules,

  toolbar: (schema: Schema) => [
    {
      id: 'math-inline',
      label: '行内公式',
      run: insertMathInline(schema),
      isActive: () => false,
    },
    {
      id: 'math-block',
      label: '块级公式',
      run: insertMathBlock(schema),
      isActive: () => false,
    },
  ],

  exportNode: (node) => {
    const latex = (node.attrs?.latex ?? '') as string;
    if (!latex) return null;
    return new DocxMath({ children: [new MathRun(latex)] });
  },
};
