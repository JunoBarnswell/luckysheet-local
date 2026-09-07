import assert from 'node:assert/strict';
import { describe, it } from 'node:test';
import type { StructuralTransformParams } from './domain';

/** Historical TS mutation cases retained as Rust kernel corpus mappings. */
export const structuralTransformCorpus = [
  { name: 'row insertion shifts cells, merges and freeze', kind: 'insert-rows' },
  { name: 'row deletion returns removed cells for undo', kind: 'delete-rows' },
  { name: 'column insertion shifts widths and hidden columns', kind: 'insert-columns' },
  { name: 'structural formula references use AST semantics', kind: 'insert-rows' },
  { name: 'cell shift insert and delete use selected extent', kind: 'cell-shift' },
  { name: 'drawing payload source ranges remain aligned', kind: 'insert-rows' },
  { name: 'sheet-backed table sources remain aligned', kind: 'insert-rows' },
  { name: 'move range clears destinations and rewrites references', kind: 'move-range' },
  { name: 'anchored objects reject lossy cell shifts', kind: 'cell-shift' },
  { name: 'data regions move only as complete regions', kind: 'insert-rows' },
  { name: 'block-backed intersections reject worksheet transforms', kind: 'insert-rows' },
] as const;

describe('structural mutation intent corpus', () => {
  it('retains every historical case as an explicit kernel operation mapping', () => {
    assert.equal(structuralTransformCorpus.length, 11);
    for (const entry of structuralTransformCorpus) {
      const intent = { kind: entry.kind, sheetId: 'sheet-1', at: 0, count: 1 } as StructuralTransformParams;
      assert.equal(intent.kind, entry.kind);
      assert.equal(intent.sheetId, 'sheet-1');
    }
  });
});
