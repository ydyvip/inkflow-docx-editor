import type { Schema, Node as PMNode } from 'prosemirror-model';
import type { EditorState, Transaction } from 'prosemirror-state';

export function createTable(
  schema: Schema,
  rows: number,
  cols: number
): PMNode {
  const cell = () => schema.nodes.table_cell.createAndFill()!;
  const row = () =>
    schema.nodes.table_row.create(null, Array.from({ length: cols }, cell));
  return schema.nodes.table.create(null, Array.from({ length: rows }, row));
}

export function insertTable(rows = 3, cols = 3) {
  return (state: EditorState, dispatch?: (tr: Transaction) => void) => {
    const table = createTable(state.schema, rows, cols);
    if (dispatch) {
      const tr = state.tr.replaceSelectionWith(table);
      dispatch(tr);
    }
    return true;
  };
}
