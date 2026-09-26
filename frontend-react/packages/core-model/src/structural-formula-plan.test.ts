import assert from 'node:assert/strict';
import { describe, it } from 'node:test';
import { collectFormulaDependencies, collectFormulaReferenceNodes, parseFormula, RangeIndex } from '@react-sheets/formula-engine';
import { StructuralTransform, WorkbookModel, type CellData, type StructuralTransformParams } from './index';

const sheetId = 'sheet-1';
const operations: readonly { name: string; params: StructuralTransformParams; afterFormula: string; row: number; column: number }[] = [
  { name: 'insert rows', params: { kind: 'insert-rows', sheetId, at: 2, count: 1 }, afterFormula: '=Sheet1!$E$6', row: 5, column: 4 },
  { name: 'delete rows', params: { kind: 'delete-rows', sheetId, at: 2, count: 1 }, afterFormula: '=Sheet1!$E$4', row: 3, column: 4 },
  { name: 'insert columns', params: { kind: 'insert-columns', sheetId, at: 2, count: 1 }, afterFormula: '=Sheet1!$F$5', row: 4, column: 5 },
  { name: 'delete columns', params: { kind: 'delete-columns', sheetId, at: 2, count: 1 }, afterFormula: '=Sheet1!$D$5', row: 4, column: 3 },
  ...(['insert', 'delete'] as const).flatMap((operation) => [
    { name: `${operation} cells ${operation === 'insert' ? 'down' : 'up'}`, params: { kind: 'cell-shift', sheetId, axis: 'row', operation,
      sourceRange: { sheetId, startRow: 2, endRow: 2, startColumn: 4, endColumn: 4 } } as const,
    afterFormula: operation === 'insert' ? '=Sheet1!$E$6' : '=Sheet1!$E$4', row: operation === 'insert' ? 5 : 3, column: 4 },
    { name: `${operation} cells ${operation === 'insert' ? 'right' : 'left'}`, params: { kind: 'cell-shift', sheetId, axis: 'column', operation,
      sourceRange: { sheetId, startRow: 4, endRow: 4, startColumn: 2, endColumn: 2 } } as const,
    afterFormula: operation === 'insert' ? '=Sheet1!$F$5' : '=Sheet1!$D$5', row: 4, column: operation === 'insert' ? 5 : 3 },
  ]),
  { name: 'move range', params: { kind: 'move-range', sheetId,
    sourceRange: { sheetId, startRow: 4, endRow: 4, startColumn: 4, endColumn: 4 }, targetOrigin: { row: 6, column: 6 } },
  afterFormula: '=Sheet1!$G$7', row: 6, column: 6 },
];

describe('prepared structural formula writes', () => {
  for (const { name, params, afterFormula, row, column } of operations) {
    it(`${name} rewrites referenced cells without loading other sheets`, () => {
      const workbook = new WorkbookModel(`lazy-formula-${name}`, 'Lazy formulas');
      const target = workbook.getSheet(sheetId);
      target.rowCount = 12;
      target.columnCount = 12;
      target.cells.set(4, 4, { value: 42 });
      const remote = workbook.addSheet('remote', 'Remote', 600_001, 16_384);
      const untouched = workbook.addSheet('untouched', 'Untouched', 600_001, 16_384);
      const formula = '=Sheet1!$E$5';
      const persisted: Record<string, Record<string, CellData>> = {
        '0': { '0': { value: 'unrelated' } },
        '2': { '2': { value: null, formula, style: { fontFamily: ' calibri ' },
          formulaMetadata: { kind: 'normal', sourceFormula: formula },
          presentation: { kind: 'barcode', symbology: 'qr', source: { kind: 'formula', formula },
            parameters: { symbology: 'qr' }, options: { foreground: '#000000', background: '#ffffff',
              showText: true, labelPosition: 'below', quietZone: 2 } } } },
        '600000': { '16383': { value: 'far sparse cell' } },
      };
      const original = structuredClone(persisted);
      remote.cells.deferJSON(persisted);
      untouched.cells.deferJSON({ '600000': { '16383': { value: 'never loaded' } } });
      const remoteAddress = { sheetId: remote.id, row: 2, column: 2 };
      const sheetOrder = workbook.sheetOrder.map((id) => ({ id, name: workbook.getSheet(id).name }));
      const index = new RangeIndex(sheetOrder);
      index.set(remoteAddress, collectFormulaDependencies(parseFormula(formula), remoteAddress, { sheetOrder }));

      const result = StructuralTransform.apply(workbook, params, index);

      assert.equal(remote.cells.isHydrated, false);
      assert.equal(untouched.cells.isHydrated, false);
      const changed = remote.cells.getFormulaOwnerWithoutHydration(2, 2);
      assert.equal(changed?.formula, afterFormula);
      assert.equal(changed?.formulaMetadata?.sourceFormula, afterFormula);
      assert.equal(changed?.style?.fontFamily, 'Calibri');
      assert.equal(changed?.presentation?.kind === 'barcode' && changed.presentation.source.kind === 'formula'
        ? changed.presentation.source.formula : undefined, afterFormula);
      assert.equal(remote.cells.getWithoutHydration(600_000, 16_383)?.value, 'far sparse cell');
      assert.equal(target.cells.get(row, column)?.value, 42);
      assert.deepEqual(persisted, original);
      assert.deepEqual(result.rewrittenFormulaOwners, [remoteAddress]);
      const deltas = result.formulaOwnerDeltas?.filter((delta) => delta.kind === 'formula-cell');
      assert.equal(deltas?.length, 1);
      assert.deepEqual(deltas?.[0]?.beforeAddress, remoteAddress);
      assert.deepEqual(deltas?.[0]?.afterAddress, remoteAddress);
      assert.deepEqual(deltas?.[0]?.before, { formula, sourceFormula: formula, barcodeFormula: formula });
      assert.deepEqual(deltas?.[0]?.after, { formula: afterFormula, sourceFormula: afterFormula, barcodeFormula: afterFormula });
    });
  }

  for (const { name, params, afterFormula, row, column } of operations.filter((operation) => operation.params.kind !== 'move-range')) {
    it(`${name} records the relocated formula address before applying geometry`, () => {
      const workbook = new WorkbookModel(`local-formula-${name}`, 'Local formulas');
      const target = workbook.getSheet(sheetId);
      target.rowCount = 12;
      target.columnCount = 12;
      const formula = '=Sheet1!$E$5';
      target.cells.set(4, 4, { value: null, formula });
      const beforeAddress = { sheetId, row: 4, column: 4 };
      const afterAddress = { sheetId, row, column };
      const sheetOrder = workbook.sheetOrder.map((id) => ({ id, name: workbook.getSheet(id).name }));
      const index = new RangeIndex(sheetOrder);
      index.set(beforeAddress, collectFormulaDependencies(parseFormula(formula), beforeAddress, { sheetOrder }));

      const result = StructuralTransform.apply(workbook, params, index);

      assert.equal(target.cells.get(row, column)?.formula, afterFormula);
      assert.deepEqual(result.rewrittenFormulaOwners, [afterAddress]);
      assert.deepEqual(result.formulaOwnerDeltas, [{ kind: 'formula-cell', beforeAddress, afterAddress,
        before: { formula, sourceFormula: null, barcodeFormula: null },
        after: { formula: afterFormula, sourceFormula: null, barcodeFormula: null } }]);
    });
  }

  it('rejects invalid changed-cell data before structural writes without hydrating its sheet', () => {
    for (const { name, params } of operations) {
      const workbook = new WorkbookModel(`rejected-lazy-formula-${name}`, 'Rejected lazy formula');
      const target = workbook.getSheet(sheetId);
      target.rowCount = 12;
      target.columnCount = 12;
      target.cells.set(4, 4, { value: 42 });
      const remote = workbook.addSheet('remote', 'Remote', 12, 12);
      const formula = '=Sheet1!$E$5';
      remote.cells.deferJSON({ '2': { '2': { value: null, formula, style: { fontFamily: 'invalid\u0000font' } } } });
      const owner = { sheetId: remote.id, row: 2, column: 2 };
      const sheetOrder = workbook.sheetOrder.map((id) => ({ id, name: workbook.getSheet(id).name }));
      const index = new RangeIndex(sheetOrder);
      index.set(owner, collectFormulaDependencies(parseFormula(formula), owner, { sheetOrder }));
      const before = workbook.snapshot();

      assert.throws(() => StructuralTransform.apply(workbook, params, index), /Font family contains control characters/);

      assert.deepEqual(workbook.snapshot(), before);
      assert.equal(remote.cells.isHydrated, false);
    }
  });

  it('plans rule formulas and geometry together without sharing ranges with history facts', () => {
    for (const { name, params, afterFormula, row, column } of operations.filter((operation) => operation.params.kind !== 'move-range')) {
      const workbook = new WorkbookModel(`rule-facts-${name}`, 'Rule facts');
      const sheet = workbook.getSheet(sheetId);
      const formula = '=Sheet1!$E$5';
      const beforeRange = { sheetId, startRow: 4, endRow: 4, startColumn: 4, endColumn: 4 };
      const afterRange = { sheetId, startRow: row, endRow: row, startColumn: column, endColumn: column };
      const anchor = { sheetId, row: 4, column: 4 };
      sheet.conditionalFormats.push({ id: 'cf', sheetId, type: 'highlight', operator: 'formula', value1: formula,
        ranges: [{ ...beforeRange }], formulaAnchor: { ...anchor } });
      sheet.dataValidations.push({ id: 'dv', sheetId, type: 'custom', formula1: formula,
        ranges: [{ ...beforeRange }], formulaAnchor: { ...anchor } });
      const sheetOrder = workbook.sheetOrder.map((id) => ({ id, name: workbook.getSheet(id).name }));
      const index = new RangeIndex(sheetOrder);
      for (const [ruleKind, ruleId, field] of [['conditional-format', 'cf', 'value1'], ['data-validation', 'dv', 'formula1']] as const) {
        index.setFormulaRuleReference({ sheetId, ruleKind, ruleId, field }, collectFormulaReferenceNodes(parseFormula(formula)), anchor);
      }

      const result = StructuralTransform.apply(workbook, params, index);

      assert.equal(sheet.conditionalFormats[0]?.value1, afterFormula);
      assert.equal(sheet.dataValidations[0]?.formula1, afterFormula);
      assert.deepEqual(sheet.conditionalFormats[0]?.ranges, [afterRange]);
      assert.deepEqual(sheet.dataValidations[0]?.ranges, [afterRange]);
      const ruleDeltas = result.formulaOwnerDeltas?.filter((delta) => delta.kind === 'formula-rule');
      assert.equal(ruleDeltas?.length, 2);
      for (const delta of ruleDeltas ?? []) {
        assert.equal(delta.beforeFormula, formula);
        assert.equal(delta.afterFormula, afterFormula);
        assert.deepEqual(delta.beforeRanges, [beforeRange]);
        assert.deepEqual(delta.afterRanges, [afterRange]);
      }
      sheet.dataValidations[0]!.ranges[0]!.startRow = row + 1;
      assert.deepEqual(ruleDeltas?.find((delta) => delta.ruleId === 'dv')?.afterRanges, [afterRange]);
    }
  });

  it('keeps deleted formula cells in removed facts rather than surviving-owner deltas', () => {
    const workbook = new WorkbookModel('removed-formula-facts', 'Removed formula');
    const sheet = workbook.getSheet(sheetId);
    const owner = { sheetId, row: 2, column: 2 };
    const cell = { value: null, formula: '=C3' };
    sheet.cells.set(2, 2, cell);
    const sheetOrder = workbook.sheetOrder.map((id) => ({ id, name: workbook.getSheet(id).name }));
    const index = new RangeIndex(sheetOrder);
    index.set(owner, collectFormulaDependencies(parseFormula(cell.formula), owner, { sheetOrder }));

    const result = StructuralTransform.apply(workbook, { kind: 'delete-rows', sheetId, at: 2, count: 1 }, index);

    assert.deepEqual(result.removedCells, [{ row: 2, column: 2, cell }]);
    assert.deepEqual(result.formulaOwnerDeltas, []);
    assert.deepEqual(result.rewrittenFormulaOwners, []);
  });
});
