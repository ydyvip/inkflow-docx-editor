import type { Schema } from 'prosemirror-model';
import { keymap } from 'prosemirror-keymap';
import { history } from 'prosemirror-history';
import { inputRules } from 'prosemirror-inputrules';
import { dropCursor } from 'prosemirror-dropcursor';
import { gapCursor } from 'prosemirror-gapcursor';
import { columnResizing, tableEditing } from 'prosemirror-tables';
import { buildInputRules, buildKeymap } from 'prosemirror-example-setup';
import { pluginRegistry } from '../plugins/registry';
import { highlightPlugin } from './highlightPlugin';
import { ensureIdsPlugin } from './ensureIdsPlugin';

/**
 * Editor 模块的插件装配。
 * 顺序很重要：更具体的 keymap（插件注册的）优先于通用 baseKeymap，
 * 因此放在数组更前面。
 */
export function buildEditorPlugins(schema: Schema) {
  const pluginKeymaps = pluginRegistry
    .all()
    .filter((p) => p.keymap)
    .map((p) => keymap(p.keymap!(schema)));

  const pluginInputRules = pluginRegistry
    .all()
    .flatMap((p) => (p.inputRules ? p.inputRules(schema) : []));

  const pluginExtraPlugins = pluginRegistry
    .all()
    .flatMap((p) => (p.proseMirrorPlugins ? p.proseMirrorPlugins(schema) : []));

  return [
    ...pluginKeymaps,
    keymap(buildKeymap(schema)),
    buildInputRules(schema),
    inputRules({ rules: pluginInputRules }),
    dropCursor(),
    gapCursor(),
    columnResizing(),
    tableEditing(),
    history(),
    highlightPlugin(),
    ensureIdsPlugin(),
    ...pluginExtraPlugins,
  ];
}
