import assert from 'node:assert/strict';
import test from 'node:test';
import {
  FormulaEngine, createFormulaError, createFormulaInputFault, isFormulaInputFault,
  assertFormulaCalculationSnapshot, type ExternalCalculationLink,
} from './index';

const reference = '[Source.xlsx]Sales!B2:B3';
const source = (state: ExternalCalculationLink['state'] = 'connected'): ExternalCalculationLink => ({
  id: 'link', token: 'Source.xlsx', sourceUnitId: 'source', subject: 'reader', sourceRevision: 1,
  accessRevision: state === 'connected' ? 1 : 2, state,
  sheets: state === 'connected' ? [{ id: 'sales', name: 'Sales', rowCount: 100, columnCount: 10 }] : [],
  cells: state === 'connected' ? [
    { address: { sheetId: 'sales', row: 1, column: 1 }, value: 10 },
    { address: { sheetId: 'sales', row: 2, column: 1 }, value: 20 },
  ] : [],
});
const consumers = [
  ...['SUM', 'AVERAGE', 'COUNT', 'COUNTA', 'MIN', 'MAX'].map(name => `=${name}(${reference})`),
  `=IFERROR(COUNT(${reference}),0)`, `=ISERROR(SUM(${reference}))`,
  `=COUNTIF(${reference},">0")`, `=SUMPRODUCT(${reference})`,
  `=AGGREGATE(2,6,${reference})`, `=SUBTOTAL(2,${reference})`,
];

test('all consumed external prerequisites fail closed before aggregate/error handling and recover', async () => {
  for (const state of ['denied', 'unavailable', 'broken'] as const) {
    const engine = new FormulaEngine({ defaultSheetId: 'Target' });
    consumers.forEach((formula, row) => engine.setFormula(`A${row + 1}`, formula));
    engine.setFormula('B1', '=IFERROR(A1,0)');
    engine.applyExternalCalculationLinks([source()]);
    await engine.recalculateAsync();
    const before = consumers.map((_, row) => engine.getCellValue(`A${row + 1}`));
    assert.equal(before[0], 30); assert.equal(before[2], 2); assert.equal(engine.getCellValue('B1'), 30);
    engine.applyExternalCalculationLinks([source(state)]);
    await engine.recalculateAsync();
    for (const cell of [...consumers.map((_, row) => `A${row + 1}`), 'B1']) {
      const value = engine.getCellValue(cell);
      assert.ok(isFormulaInputFault(value), `${state}: ${cell} returned ${JSON.stringify(value)}`);
      assert.equal(value.code, state === 'denied' ? '#BLOCKED!' : state === 'broken' ? '#REF!' : '#N/A');
      assert.equal(value.inputFault.source, 'source');
    }
    const snapshot = structuredClone(engine.exportCalculationSnapshot());
    assertFormulaCalculationSnapshot(snapshot);
    const worker = FormulaEngine.fromCalculationSnapshot(snapshot);
    await worker.recalculateAsync(undefined, undefined, true);
    for (const cell of [...consumers.map((_, row) => `A${row + 1}`), 'B1']) {
      assert.deepEqual(worker.getCellValue(cell), engine.getCellValue(cell), `${state}: Worker ${cell}`);
    }
    engine.applyExternalCalculationLinks([{ ...source(), sourceRevision: 2, accessRevision: 3 }]);
    await engine.recalculateAsync();
    assert.deepEqual(consumers.map((_, row) => engine.getCellValue(`A${row + 1}`)), before);
    assert.equal(engine.getCellValue('B1'), 30);
  }
});

test('ordinary source errors keep Excel behavior while unavailable inputs remain prerequisites', async () => {
  const engine = new FormulaEngine({ defaultSheetId: 'Target' });
  const ordinary = source();
  engine.applyExternalCalculationLinks([{ ...ordinary, cells: ordinary.cells.map((cell, index) =>
    index ? { ...cell, value: createFormulaError('#N/A', 'ordinary source error') } : cell) }]);
  engine.setFormula('A1', `=COUNT(${reference})`);
  engine.setFormula('A2', '=IFERROR([Source.xlsx]Sales!B3,7)');
  await engine.recalculateAsync();
  assert.equal(engine.getCellValue('A1'), 1);
  assert.equal(engine.getCellValue('A2'), 7);
  assert.equal(isFormulaInputFault(engine.getCellValue('A2')), false);
});

test('lazy branches do not consume inaccessible inputs; hidden ranges and audit use the same boundary', async () => {
  const engine = new FormulaEngine({ defaultSheetId: 'Target', blockedRanges: [
    { sheetId: 'Target', startRow: 0, endRow: 1, startColumn: 3, endColumn: 3 },
  ] });
  engine.applyExternalCalculationLinks([source('denied')]);
  engine.setFormula('A1', `=IF(FALSE,COUNT(${reference}),8)`);
  engine.setFormula('A2', '=IFERROR(COUNT(D1:D2),0)');
  engine.setFormula('A3', '=ISERROR(D1)');
  await engine.recalculateAsync();
  assert.equal(engine.getCellValue('A1'), 8);
  assert.ok(isFormulaInputFault(engine.getCellValue('A2')));
  assert.ok(isFormulaInputFault(engine.getCellValue('A3')));
  const trace = engine.evaluateFormulaWithTrace('A2');
  assert.ok(isFormulaInputFault(trace?.value));
  assert.ok(isFormulaInputFault(trace?.steps.at(-1)?.value));
});

test('external input and Worker snapshot reject malformed fault metadata', () => {
  const engine = new FormulaEngine({ defaultSheetId: 'Target' });
  const valid = source();
  const fault = createFormulaInputFault('#BLOCKED!', 'Denied', 'access-denied', 'source');
  engine.applyExternalCalculationLinks([{ ...valid, cells: [{ ...valid.cells[0]!, value: fault }] }]);
  const snapshot = structuredClone(engine.exportCalculationSnapshot());
  assertFormulaCalculationSnapshot(snapshot);
  (snapshot as unknown as { externalLinks: { cells: { value: unknown }[] }[] }).externalLinks[0]!.cells[0]!.value =
    { ...fault, inputFault: { reason: 'unknown', source: '' } };
  assert.throws(() => assertFormulaCalculationSnapshot(snapshot), /EXTERNAL_LINK_CELL_INVALID/);
});
