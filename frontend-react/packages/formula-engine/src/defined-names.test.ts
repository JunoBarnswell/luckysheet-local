import test from 'node:test';
import assert from 'node:assert/strict';
import { FormulaEngine } from './formula-engine';
import { normalizeDefinedNameModels, normalizeDefinedNames, resolveDefinedNameSource } from './defined-names';

test('normalizeDefinedNames uppercases keys', () => {
  assert.deepEqual(normalizeDefinedNames({ TaxRate: '0.15' }), { TAXRATE: '0.15' });
});

test('normalizeDefinedNameModels rejects duplicate scoped identities without case-folding worksheet ids', () => {
  assert.throws(() => normalizeDefinedNameModels([
    { name: 'TaxRate', formula: '0.1', scope: 'workbook' },
    { name: 'taxrate', formula: '0.2', scope: 'workbook' },
  ]), /Defined-name owner identity is duplicated/);
  assert.throws(() => normalizeDefinedNameModels([
    { name: 'LocalRate', formula: '0.1', scope: 'sheet', sheetId: 'Sheet-A' },
    { name: 'localrate', formula: '0.2', scope: 'sheet', sheetId: 'Sheet-A' },
  ]), /Defined-name owner identity is duplicated/);

  const names = normalizeDefinedNameModels([
    { name: 'LocalRate', formula: '0.1', scope: 'sheet', sheetId: 'Sheet-A' },
    { name: 'LocalRate', formula: '0.2', scope: 'sheet', sheetId: 'sheet-a' },
  ]);
  assert.equal(names.length, 2);

  const engine = new FormulaEngine({
    defaultSheetId: 'Sheet-A',
    sheetOrder: [{ id: 'Sheet-A', name: 'Upper' }, { id: 'sheet-a', name: 'Lower' }],
  });
  engine.setDefinedNameModels(names);
  engine.setFormula({ sheetId: 'Sheet-A', row: 0, column: 0 }, '=LocalRate');
  engine.setFormula({ sheetId: 'sheet-a', row: 0, column: 0 }, '=LocalRate');
  assert.equal(engine.getCellValue({ sheetId: 'Sheet-A', row: 0, column: 0 }), 0.1);
  assert.equal(engine.getCellValue({ sheetId: 'sheet-a', row: 0, column: 0 }), 0.2);
});

test('resolveDefinedNameSource supports scalar, range, and formula values', () => {
  const context = {
    currentCell: { sheetId: 'Sheet1', row: 4, column: 2 },
    sheetOrder: [{ id: 'Sheet1', name: 'Sheet1' }],
    readCell: (address: { row: number; column: number }) => (address.row === 0 && address.column === 0 ? 10 : null),
    readRangeMatrix: () => [[1, 2], [3, 4]],
  };
  assert.equal(resolveDefinedNameSource('0.25', context), 0.25);
  assert.deepEqual(resolveDefinedNameSource('A1:B2', context), [[1, 2], [3, 4]]);
  assert.equal(resolveDefinedNameSource('=A1+5', context), 15);
});

test('FormulaEngine resolves defined names in formulas with current cell context', () => {
  const engine = new FormulaEngine({ defaultSheetId: 'Sheet1' });
  engine.setDefinedNames({ TaxRate: '0.1', BaseCell: 'C3' });
  engine.setValue('C3', 100);
  engine.setFormula('D3', '=BaseCell*TaxRate');
  assert.equal(engine.getCellValue('D3'), 10);
});

test('FormulaEngine recalculates formulas when a defined name changes', () => {
  const engine = new FormulaEngine({ defaultSheetId: 'Sheet1' });
  engine.setDefinedNames({ Rate: '2' });
  engine.setValue('A1', 5);
  engine.setFormula('B1', '=A1*Rate');
  assert.equal(engine.getCellValue('B1'), 10);
  engine.setDefinedNames({ Rate: '3' });
  assert.equal(engine.getCellValue('B1'), 15);
});

test('FormulaEngine refreshes formulas through nested names without recalculating unrelated formulas', () => {
  const engine = new FormulaEngine({ defaultSheetId: 'Sheet1' });
  engine.setDefinedNameModels([
    { name: 'Rate', formula: '2', scope: 'workbook' },
    { name: 'TaxedRate', formula: 'Rate*4', scope: 'workbook' },
  ]);
  engine.setFormula('A1', '=TaxedRate');
  engine.setFormula('B1', '=1+1');
  assert.equal(engine.getCellValue('A1'), 8);

  const report = engine.setDefinedNameModels([
    { name: 'Rate', formula: '3', scope: 'workbook' },
    { name: 'TaxedRate', formula: 'Rate*4', scope: 'workbook' },
  ]);

  assert.equal(engine.getCellValue('A1'), 12);
  assert.ok(report.recalculated.some(({ row, column }) => row === 0 && column === 0));
  assert.ok(!report.recalculated.some(({ row, column }) => row === 0 && column === 1));
});

test('FormulaEngine resolves sheet-scoped names before workbook-scoped names', () => {
  const engine = new FormulaEngine({ defaultSheetId: 'Sheet1' });
  engine.setDefinedNameModels([
    { name: 'Rate', formula: '2', scope: 'workbook' },
    { name: 'Rate', formula: '3', scope: 'sheet', sheetId: 'Sheet2' },
  ]);
  engine.setFormula({ sheetId: 'Sheet1', row: 0, column: 0 }, '=Rate');
  engine.setFormula({ sheetId: 'Sheet2', row: 0, column: 0 }, '=Rate');
  assert.equal(engine.getCellValue({ sheetId: 'Sheet1', row: 0, column: 0 }), 2);
  assert.equal(engine.getCellValue({ sheetId: 'Sheet2', row: 0, column: 0 }), 3);
});

test('FormulaEngine indexes defined-name references and anchors by typed owner identity', () => {
  const engine = new FormulaEngine({
    defaultSheetId: 'Sheet1',
    sheetOrder: [{ id: 'Sheet1', name: 'Sheet1' }, { id: 'Sheet2', name: 'Sheet2' }],
  });
  const localName = { scope: 'sheet' as const, sheetId: 'Sheet1', name: 'LocalRange' };
  const qualifiedName = { scope: 'workbook' as const, name: 'QualifiedRange' };
  const beforeLocalName = {
    ...localName,
    formula: '=A1:A3',
    anchor: { sheetId: 'Sheet1', row: 4, column: 2 },
  };
  const afterLocalName = {
    ...localName,
    formula: '=C9',
    anchor: { sheetId: 'Sheet1', row: 6, column: 2 },
  };
  engine.setDefinedNameModels([
    beforeLocalName,
    { ...qualifiedName, formula: "='Sheet2'!B2:B5" },
  ]);

  assert.deepEqual(engine.dependencies.getStructuralDefinedNameDependents('Sheet1', 'row', 1), [localName]);
  assert.deepEqual(engine.dependencies.getRangeDefinedNameDependents('Sheet1', {
    startRow: 1,
    endRow: 1,
    startColumn: 0,
    endColumn: 0,
  }), [localName]);
  assert.deepEqual(engine.dependencies.getDefinedNamesAnchoredAtOrAfter('Sheet1', 'row', 4), [localName]);
  assert.deepEqual(engine.dependencies.getStructuralDefinedNameDependents('Sheet2', 'row', 2), [qualifiedName]);

  assert.throws(() => engine.applyDefinedNameModelDeltas([{
    owner: { scope: 'invalid' as never, name: 'LocalRange' },
    before: beforeLocalName,
    after: afterLocalName,
  }], false), /defined-name owner identity is not unique and stable/);
  assert.deepEqual(engine.dependencies.getStructuralDefinedNameDependents('Sheet1', 'row', 1), [localName]);

  engine.applyDefinedNameModelDeltas([{
    owner: localName,
    before: beforeLocalName,
    after: afterLocalName,
  }], false);
  assert.deepEqual(engine.dependencies.getRangeDefinedNameDependents('Sheet1', {
    startRow: 0,
    endRow: 2,
    startColumn: 0,
    endColumn: 0,
  }), []);
  assert.deepEqual(engine.dependencies.getDefinedNamesAnchoredAtOrAfter('Sheet1', 'row', 6), [localName]);
});

test('FormulaEngine keeps contextless workbook-name references fail-closed in the structural index', () => {
  const engine = new FormulaEngine({ defaultSheetId: 'Sheet1' });
  engine.setDefinedNameModels([{ name: 'Contextual', formula: '=A1', scope: 'workbook' }]);

  assert.deepEqual(engine.dependencies.getDefinedNameReferenceFailures(), [
    { owner: { scope: 'workbook', name: 'Contextual' }, reason: 'unresolved-context' },
  ]);
});

test('scoped names survive the calculation worker snapshot boundary', () => {
  const engine = new FormulaEngine({ defaultSheetId: 'Sheet1', recalculationMode: 'manual' });
  engine.setDefinedNameModels([
    { name: 'Rate', formula: '2', scope: 'workbook' },
    { name: 'Rate', formula: '4', scope: 'sheet', sheetId: 'Sheet2' },
  ]);
  engine.setFormula({ sheetId: 'Sheet2', row: 0, column: 0 }, '=Rate');
  const restored = FormulaEngine.fromCalculationSnapshot(engine.exportCalculationSnapshot());
  restored.executeCalculationTask({
    protocol: 'react-sheets.formula-calculation',
    version: 2,
    taskId: 'scoped-name-snapshot',
    kind: 'recalculate',
    revision: 1,
  });
  assert.equal(restored.getCellValue({ sheetId: 'Sheet2', row: 0, column: 0 }), 4);
});

test('FormulaEngine recalculates volatile formulas when dependencies change', () => {
  const engine = new FormulaEngine({ defaultSheetId: 'Sheet1' });
  engine.setFormula('A1', '=RAND()');
  const first = engine.getCellValue('A1');
  const second = engine.getCellValue('A1');
  assert.equal(first, second);
  engine.setValue('B1', 1);
  const third = engine.getCellValue('A1');
  assert.notEqual(second, third);
});
