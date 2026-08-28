import assert from 'node:assert/strict';
import test from 'node:test';
import { createSpreadsheetShortcutRegistry } from './shortcut-registry';

test('Home shortcuts resolve to canonical semantic command ids', () => {
  const registry = createSpreadsheetShortcutRegistry();
  const grid = { scope: 'grid' as const };
  assert.equal(registry.resolve({ key: 'v', ctrlKey: true }, grid)?.id, 'clipboard.paste');
  assert.equal(registry.resolve({ key: 'v', ctrlKey: true, altKey: true }, grid)?.id, 'clipboard.pasteSpecial');
  assert.equal(registry.resolve({ key: 'l', ctrlKey: true, shiftKey: true }, grid)?.id, 'filter.toggle');
  assert.equal(registry.resolve({ key: 'd', ctrlKey: true }, grid)?.id, 'range.fillDown');
  assert.equal(registry.resolve({ key: 'p', ctrlKey: true }, grid)?.id, 'print.preview');
  assert.equal(registry.resolve({ key: 'p', ctrlKey: true, shiftKey: true }, grid)?.id, 'commandPalette.open');
  assert.equal(registry.resolve({ key: 'r', ctrlKey: true }, grid)?.id, 'range.fillRight');
  assert.equal(registry.resolve({ key: 'a', ctrlKey: true }, grid)?.id, 'selection.selectAll');
  assert.equal(registry.resolve({ key: 'F8', shiftKey: true }, grid)?.id, 'selection.addMode');
  assert.equal(registry.resolve({ key: 'Delete' }, grid)?.id, 'range.clearContents');
  assert.equal(registry.resolve({ key: 'F5' }, grid)?.id, 'navigation.goto');
  assert.equal(registry.resolve({ key: 'h', altKey: true }, grid)?.id, 'ribbon.home.keyTips');
  assert.equal(registry.resolve({ key: 'Escape' }, grid)?.id, 'clipboard.cancel');
  assert.deepEqual(registry.resolveSequence({ key: 'e', altKey: true }, grid), { state: { active: true, index: 1 }, preventDefault: true });
  assert.deepEqual(registry.resolveSequence({ key: 's' }, grid, { active: true, index: 1 })?.shortcut, { id: 'clipboard.pasteSpecial', preventDefault: true, scope: 'grid' });
});

test('shortcut bindings remain context-aware', () => {
  const registry = createSpreadsheetShortcutRegistry();
  assert.equal(registry.resolve({ key: 'v', ctrlKey: true }, { scope: 'dialog' }), undefined);
  assert.equal(registry.resolve({ key: 'F4' }, { scope: 'cell-editor', formulaReferenceSelected: false }), undefined);
  assert.equal(registry.resolve({ key: 'F4' }, { scope: 'cell-editor', formulaReferenceSelected: true })?.id, 'formula.toggleAbsolute');
});
