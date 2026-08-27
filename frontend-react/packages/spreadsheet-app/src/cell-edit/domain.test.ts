import assert from 'node:assert/strict';
import test from 'node:test';
import { CellEditDomain } from './domain';
import { rewriteAbsoluteReferenceAtCaret } from './formula-edit';
import { createCellEditorAdapterRegistry } from './builtin-adapters';
import { CellEditError } from './error';
import type { CellEditEntryContext, CellEditIntent } from './contracts';

function entry(overrides: Partial<CellEditEntryContext> = {}): CellEditEntryContext {
  return {
    target: { display: { sheetId: 'sheet-1', row: 1, column: 1 }, canonical: { sheetId: 'sheet-1', row: 1, column: 1 } },
    source: 'f2', surface: 'grid', adapterKind: 'text', editorSurface: { kind: 'text', inputMode: 'text', multiline: true },
    initialDraft: { kind: 'plain', text: 'old' }, caret: { start: 3, end: 3 },
    originalSelection: { unitId: 'unit-1', sheetId: 'sheet-1', ranges: [{ sheetId: 'sheet-1', startRow: 1, endRow: 1, startColumn: 1, endColumn: 1 }], primaryRangeIndex: 0, activeCell: { row: 1, column: 1 }, anchorCell: { row: 1, column: 1 }, selectionKind: 'cells', mode: 'normal' },
    originalCell: { value: 'old' }, baseCellFingerprint: '{"value":"old"}', enterMove: 'down', groupedSheetIds: ['sheet-1'], ...overrides,
  };
}

function begin(domain: CellEditDomain, overrides: Partial<CellEditEntryContext> = {}): void {
  assert.equal(domain.dispatch({ type: 'begin', entry: entry(overrides) }).handled, true);
}

function key(domain: CellEditDomain, value: Partial<Extract<CellEditIntent, { type: 'keyboard' }>['gesture']> & { key: string }): ReturnType<CellEditDomain['dispatch']> {
  return domain.dispatch({ type: 'keyboard', gesture: { key: value.key, code: value.code ?? value.key, alt: value.alt ?? false, ctrl: value.ctrl ?? false, meta: value.meta ?? false, shift: value.shift ?? false, repeat: value.repeat ?? false, composing: value.composing ?? false } });
}

test('direct typing owns Enter state and draft undo/redo without model effects', () => {
  const domain = new CellEditDomain();
  begin(domain, { source: 'direct-typing', initialDraft: { kind: 'plain', text: 'x' }, beforeInitialDraft: { kind: 'plain', text: 'old' }, caret: { start: 1, end: 1 } });
  assert.equal(domain.getSnapshot().status, 'enter');
  assert.equal(domain.getHistoryMetrics().undoCount, 1);
  key(domain, { key: 'y' });
  assert.equal(domain.getSnapshot().session?.draft.text, 'xy');
  const undo = key(domain, { key: 'z', ctrl: true, code: 'KeyZ' });
  assert.equal(undo.effects.some((effect) => effect.type === 'commit'), false);
  assert.equal(domain.getSnapshot().session?.draft.text, 'x');
  key(domain, { key: 'y', ctrl: true, code: 'KeyY' });
  assert.equal(domain.getSnapshot().session?.draft.text, 'xy');
});

test('Alt+Enter inserts a newline while Ctrl+Enter requests one atomic selection commit', () => {
  const domain = new CellEditDomain();
  begin(domain);
  const newline = key(domain, { key: 'Enter', code: 'Enter', alt: true });
  assert.equal(newline.effects.some((effect) => effect.type === 'commit'), false);
  assert.equal(domain.getSnapshot().session?.draft.text, 'old\n');
  const commit = key(domain, { key: 'Enter', code: 'Enter', ctrl: true });
  const effect = commit.effects.find((candidate) => candidate.type === 'commit');
  assert.equal(effect?.type, 'commit');
  if (effect?.type === 'commit') assert.equal(effect.request.toSelection, true);
});

test('formula status toggles Edit and Point and F4 follows the Excel absolute cycle', () => {
  const domain = new CellEditDomain();
  begin(domain, { adapterKind: 'formula', initialDraft: { kind: 'plain', text: '=A1' }, caret: { start: 1, end: 3 } });
  key(domain, { key: 'F2', code: 'F2' });
  assert.equal(domain.getSnapshot().status, 'point');
  key(domain, { key: 'F2', code: 'F2' });
  assert.equal(domain.getSnapshot().status, 'edit');
  let formula = '=A1';
  for (const expected of ['=$A$1', '=A$1', '=$A1', '=A1']) {
    const rewrite = rewriteAbsoluteReferenceAtCaret(formula, { start: 1, end: formula.length });
    assert.ok(rewrite);
    formula = rewrite.text;
    assert.equal(formula, expected);
  }
});

test('IME composition is one draft history action and composing keys keep browser defaults', () => {
  const domain = new CellEditDomain();
  begin(domain, { initialDraft: { kind: 'plain', text: '' }, caret: { start: 0, end: 0 }, originalCell: null, baseCellFingerprint: 'null' });
  domain.dispatch({ type: 'composition.start' });
  assert.equal(key(domain, { key: 'Process', code: 'Process', composing: true }).preventDefault, false);
  domain.dispatch({ type: 'text.replace', text: '中', caret: { start: 1, end: 1 } });
  domain.dispatch({ type: 'text.replace', text: '中文', caret: { start: 2, end: 2 } });
  domain.dispatch({ type: 'composition.end', text: '中文', caret: { start: 2, end: 2 } });
  assert.equal(domain.getHistoryMetrics().undoCount, 1);
  key(domain, { key: 'z', ctrl: true, code: 'KeyZ' });
  assert.equal(domain.getSnapshot().session?.draft.text, '');
});

test('reference pointer gesture coalesces repeated rewrites into one draft undo action', () => {
  const domain = new CellEditDomain();
  begin(domain, { adapterKind: 'formula', initialDraft: { kind: 'plain', text: '=A1' }, caret: { start: 1, end: 3 } });
  domain.dispatch({ type: 'reference.gesture.begin' });
  domain.dispatch({ type: 'reference.insert', referenceText: 'B2', selection: { id: 'r1', sheetId: 'sheet-1', range: { sheetId: 'sheet-1', startRow: 1, endRow: 1, startColumn: 1, endColumn: 1 }, tokenSpan: { start: 1, end: 3 }, colorIndex: 0, operation: 'move' } });
  domain.dispatch({ type: 'caret.set', caret: { start: 1, end: 3 } });
  domain.dispatch({ type: 'reference.insert', referenceText: 'C3', selection: { id: 'r1', sheetId: 'sheet-1', range: { sheetId: 'sheet-1', startRow: 2, endRow: 2, startColumn: 2, endColumn: 2 }, tokenSpan: { start: 1, end: 3 }, colorIndex: 0, operation: 'move' } });
  domain.dispatch({ type: 'reference.gesture.end' });
  assert.equal(domain.getSnapshot().session?.draft.text, '=C3');
  assert.equal(domain.getHistoryMetrics().undoCount, 1);
});

test('validation confirmation keeps the draft active and replays the pending commit', () => {
  const domain = new CellEditDomain();
  begin(domain);
  domain.dispatch({ type: 'commit', moveAfter: 'none' });
  domain.dispatch({ type: 'commit.failed', error: { code: 'CELL_EDIT_CONFIRMATION_REQUIRED', message: 'Confirm', recovery: 'Choose', alertStyle: 'warning' } });
  assert.equal(domain.getSnapshot().session?.validation.kind, 'confirmation-required');
  const effect = domain.dispatch({ type: 'validation.confirm' }).effects.find((candidate) => candidate.type === 'commit');
  assert.equal(effect?.type, 'commit');
  if (effect?.type === 'commit') assert.equal(effect.request.validationConfirmation, true);
});

test('partial character formatting creates canonical rich text and is undoable', () => {
  const domain = new CellEditDomain();
  begin(domain, { caret: { start: 0, end: 2 } });
  domain.dispatch({ type: 'rich-text.format', style: { bold: true, verticalAlignment: 'superscript' } });
  const session = domain.getSnapshot().session;
  assert.equal(session?.draft.kind, 'rich-text');
  if (session?.draft.kind === 'rich-text') assert.equal(session.draft.runs[0]?.style?.verticalAlignment, 'superscript');
  domain.dispatch({ type: 'draft.undo' });
  assert.equal(domain.getSnapshot().session?.draft.kind, 'plain');
});

test('editor registry dispatches number, mask, ComboBox and custom contracts without UI kind branches', () => {
  const registry = createCellEditorAdapterRegistry();
  assert.deepEqual(registry.listKinds(), ['text', 'number', 'datetime', 'validation-list', 'combo-box', 'checkbox', 'mask', 'formula', 'rich-text']);
  const inputContext = { sourceKind: 'direct-entry' as const, cultureId: 'en-US', decimalSeparator: '.', groupSeparator: ',', dateSystem: '1900' as const, referenceDate: { year: 2026, month: 8, day: 28, hour: 0, minute: 0, second: 0, millisecond: 0 } };
  const target = entry().target;
  const number = registry.get('number');
  assert.deepEqual(number.validate({ kind: 'plain', text: '42.5' }, { target, cell: null, inputContext }), { valid: true });
  assert.equal(number.validate({ kind: 'plain', text: 'abc' }, { target, cell: null, inputContext }).valid, false);
  const mask = registry.get('mask');
  assert.equal(mask.validate({ kind: 'plain', text: '12-AB' }, { target, cell: null, inputContext, config: { kind: 'mask', mask: '##-AA' } }).valid, true);
  const combo = registry.get('combo-box');
  const payload = combo.toCommitPayload({ kind: 'plain', text: 'Approved' }, { target, cell: null, inputContext, config: { kind: 'combo-box', items: [{ label: 'Approved', value: 1 }], editable: false } });
  assert.deepEqual(payload, { kind: 'typed-value', value: 1 });
  registry.register({ kind: 'custom:rating', surface: { kind: 'custom', inputMode: 'numeric', multiline: false }, canEnter: () => ({ allowed: true }), createDraft: () => ({ kind: 'plain', text: '3' }), reduce: (_intent, draft) => draft, ownsKey: () => true, validate: () => ({ valid: true }), toCommitPayload: (draft) => ({ kind: 'raw-text', text: draft.text }) });
  assert.equal(registry.resolve({ target, cell: { value: 3, editor: { kind: 'custom', adapterId: 'rating' } }, inputContext }).kind, 'custom:rating');
  assert.throws(() => registry.get('custom:missing'), (error) => error instanceof CellEditError && error.code === 'CELL_EDIT_ADAPTER_NOT_FOUND');
});
