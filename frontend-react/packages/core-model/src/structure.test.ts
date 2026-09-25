import assert from 'node:assert/strict';
import { describe, it } from 'node:test';
import { collectFormulaDependencies, parseFormula, RangeIndex } from '@react-sheets/formula-engine';
import { CellMatrix, StructuralTransform as CoreStructuralTransform, WorkbookModel, type StructuralTransformParams } from './index';
import type { ReportSheetDefinition } from './data-model';

const StructuralTransform = {
  apply(workbook: WorkbookModel, params: StructuralTransformParams) {
    const sheetOrder = workbook.sheetOrder.map((id) => ({ id, name: workbook.getSheet(id).name }));
    const index = new RangeIndex(sheetOrder);
    for (const sheet of workbook.getSheets()) {
      sheet.cells.forEach((cell, row, column) => {
        if (cell.formula === undefined) return;
        const owner = { sheetId: sheet.id, row, column };
        try {
          const formula = cell.formula.trimStart().startsWith('=') ? cell.formula : `=${cell.formula}`;
          index.set(owner, collectFormulaDependencies(parseFormula(formula), owner, { sheetOrder }));
        } catch {
          index.set(owner, [], true);
        }
      });
    }
    return CoreStructuralTransform.apply(workbook, params, index);
  },
};

function seedWorkbook(): { workbook: WorkbookModel; sheetId: string } {
  const workbook = new WorkbookModel('unit-test', 'Structural');
  const sheet = workbook.addSheet('s1', 'Sheet1');
  sheet.cells.set(0, 0, { value: 'A0' });
  sheet.cells.set(1, 0, { value: 'A1' });
  sheet.cells.set(2, 0, { value: 'A2' });
  return { workbook, sheetId: sheet.id };
}

function reportDefinition(
  sheetId: string,
  cells: readonly { row: number; column: number }[],
  repeatHeaderRows: number[] = [],
): ReportSheetDefinition {
  return {
    templateSheetId: sheetId,
    bindings: cells.map((cell) => ({ cell: { ...cell }, expression: 'field-id', kind: 'field' as const })),
    pagination: { enabled: true, rowsPerPage: 20, repeatHeaderRows },
    renderMode: 'preview',
    layout: { orientation: 'portrait', marginTopPx: 0, marginRightPx: 0, marginBottomPx: 0, marginLeftPx: 0 },
    dataEntry: [],
  };
}

describe('structural operations', () => {
  it('shiftRows moves cells below the insertion point', () => {
    const matrix = new CellMatrix();
    matrix.set(0, 0, { value: 'top' });
    matrix.set(5, 1, { value: 'bottom' });
    matrix.shiftRows(3, 2, 1);
    assert.equal(matrix.get(0, 0)?.value, 'top');
    assert.equal(matrix.get(7, 1)?.value, 'bottom');
    matrix.shiftRows(3, 2, -1);
    assert.equal(matrix.get(5, 1)?.value, 'bottom');
  });

  it('axis shifts hydrate deferred sparse cells before reading row buckets', () => {
    const matrix = new CellMatrix();
    matrix.deferJSON({ '5': { '1': { value: 'row' }, '4': { value: 'column' } } });
    matrix.shiftRows(3, 2, 1);
    assert.equal(matrix.get(7, 1)?.value, 'row');
    matrix.shiftColumns(3, 2, 1);
    assert.equal(matrix.get(7, 6)?.value, 'column');
  });

  it('StructuralTransform insertRows keeps merges and freeze consistent', () => {
    const { workbook } = seedWorkbook();
    const sheet = workbook.getSheet('s1');
    sheet.merges.push({ range: { sheetId: 's1', startRow: 2, endRow: 3, startColumn: 1, endColumn: 1 }, anchor: { row: 2, column: 1 } });
    sheet.pane = { kind: 'frozen', xSplit: 0, ySplit: 2, startRow: 2, startColumn: 0, state: 'frozen' };
    StructuralTransform.apply(workbook, { kind: 'insert-rows', sheetId: sheet.id, at: 2, count: 3 });
    assert.equal(sheet.cells.get(2, 0)?.value ?? null, null);
    assert.equal(sheet.cells.get(5, 0)?.value, 'A2');
    assert.equal(sheet.merges[0]!.range.startRow, 5);
    assert.equal(sheet.pane.kind === 'frozen' ? sheet.pane.ySplit : 0, 5);
    assert.equal(sheet.rowCount, 1003);
  });

  it('StructuralTransform deleteRows removes region and returns extracted cells for undo', () => {
    const { workbook } = seedWorkbook();
    const sheet = workbook.getSheet('s1');
    sheet.cells.set(1, 0, { value: 'gone' });
    const removed = StructuralTransform.apply(workbook, { kind: 'delete-rows', sheetId: sheet.id, at: 1, count: 1 }).removedCells;
    assert.ok(removed.some((entry) => entry.cell.value === 'gone'));
    assert.equal(sheet.cells.get(1, 0)?.value, 'A2');
    // 恢复
    StructuralTransform.apply(workbook, { kind: 'insert-rows', sheetId: sheet.id, at: 1, count: 1 });
    for (const entry of removed) sheet.cells.set(entry.row, entry.column, entry.cell);
    assert.equal(sheet.cells.get(1, 0)?.value, 'gone');
  });

  it('insertColumns shifts widths and hidden columns', () => {
    const { workbook } = seedWorkbook();
    const sheet = workbook.getSheet('s1');
    sheet.columnWidthsPx[2] = 200;
    sheet.hiddenColumns.add(3);
    StructuralTransform.apply(workbook, { kind: 'insert-columns', sheetId: sheet.id, at: 1, count: 2 });
    assert.equal(sheet.columnWidthsPx[4], 200);
    assert.ok(sheet.hiddenColumns.has(5));
    assert.ok(!sheet.hiddenColumns.has(3));
  });

  it('column insertion before a Sheet Table preserves its column schema', () => {
    const { workbook } = seedWorkbook();
    const sheet = workbook.getSheet('s1');
    sheet.sheetTables.push({
      id: 'table-1', sheetId: sheet.id, name: 'Table1',
      range: { sheetId: sheet.id, startRow: 0, endRow: 2, startColumn: 1, endColumn: 2 },
      hasHeaderRow: true, hasTotalRow: false, showBandedRows: true, showBandedColumns: false,
      showFirstColumn: false, showLastColumn: false, showFilterButton: false, autoExpand: 'none',
      columns: [{ id: 'column-1', name: 'A' }, { id: 'column-2', name: 'B' }],
    });

    StructuralTransform.apply(workbook, { kind: 'insert-columns', sheetId: sheet.id, at: 1, count: 1 });
    StructuralTransform.apply(workbook, { kind: 'insert-columns', sheetId: sheet.id, at: 4, count: 1 });

    assert.equal(sheet.sheetTables[0]!.range.startColumn, 2);
    assert.equal(sheet.sheetTables[0]!.range.endColumn, 3);
    assert.equal(sheet.sheetTables[0]!.columns.length, 2);
  });

  it('rejects column insertion inside a Sheet Table before changing workbook state', () => {
    const { workbook } = seedWorkbook();
    const sheet = workbook.getSheet('s1');
    sheet.cells.set(4, 4, { value: 'tail' });
    sheet.sheetTables.push({
      id: 'table-1', sheetId: sheet.id, name: 'Table1',
      range: { sheetId: sheet.id, startRow: 0, endRow: 2, startColumn: 1, endColumn: 2 },
      hasHeaderRow: true, hasTotalRow: false, showBandedRows: true, showBandedColumns: false,
      showFirstColumn: false, showLastColumn: false, showFilterButton: false, autoExpand: 'none',
      columns: [{ id: 'column-1', name: 'A' }, { id: 'column-2', name: 'B' }],
    });
    const columnCount = sheet.columnCount;

    assert.throws(() => StructuralTransform.apply(workbook, {
      kind: 'insert-columns', sheetId: sheet.id, at: 2, count: 1,
    }), /UNSUPPORTED_FEATURE: inserting a worksheet column inside Sheet Table table-1/);

    assert.equal(sheet.columnCount, columnCount);
    assert.equal(sheet.cells.get(4, 4)?.value, 'tail');
    assert.deepEqual(sheet.sheetTables[0]!.range, {
      sheetId: sheet.id, startRow: 0, endRow: 2, startColumn: 1, endColumn: 2,
    });
    assert.equal(sheet.sheetTables[0]!.columns.length, 2);
  });

  it('structural formula rewrite uses AST references, including absolute and quoted refs', () => {
    const workbook = new WorkbookModel('unit-formula-structure', 'Structural Formula');
    const sheet = workbook.getSheet('sheet-1');
    sheet.name = 'Input Sheet';
    sheet.cells.set(5, 0, { value: null, formula: "=SUM($A$1,'Input Sheet'!$B$1,A1)+\"A1\"" });

    StructuralTransform.apply(workbook, { kind: 'insert-rows', sheetId: sheet.id, at: 0, count: 1 });

    assert.equal(sheet.cells.get(6, 0)?.formula, "=SUM($A$2,'Input Sheet'!$B$2,A2)+\"A1\"");
  });

  it('returns reversible formula-owner state when deleting a referenced row', () => {
    const workbook = new WorkbookModel('unit-formula-owner-delta', 'Formula Owner Delta');
    const sheet = workbook.getSheet('sheet-1');
    sheet.cells.set(1, 0, { value: 7 });
    sheet.cells.set(0, 1, { value: null, formula: '=A2*2' });

    const result = StructuralTransform.apply(workbook, { kind: 'delete-rows', sheetId: sheet.id, at: 1, count: 1 });
    const delta = result.formulaOwnerDeltas?.find((entry) => entry.beforeAddress.row === 0 && entry.beforeAddress.column === 1);

    assert.ok(delta);
    assert.deepEqual(delta.beforeAddress, { sheetId: sheet.id, row: 0, column: 1 });
    assert.deepEqual(delta.afterAddress, { sheetId: sheet.id, row: 0, column: 1 });
    assert.equal(delta.before.formula, '=A2*2');
    assert.equal(delta.after.formula, '=#REF!*2');
    assert.equal(sheet.cells.get(0, 1)?.formula, delta.after.formula);
  });

  it('maps report binding and repeated-header row owners on axis insertion and rejects deleting a binding anchor', () => {
    const { workbook } = seedWorkbook();
    const sheet = workbook.getSheet('s1');
    sheet.reportSheet = reportDefinition(sheet.id, [{ row: 2, column: 0 }], [0, 2]);

    StructuralTransform.apply(workbook, { kind: 'insert-rows', sheetId: sheet.id, at: 2, count: 1 });

    assert.deepEqual(sheet.reportSheet?.bindings[0]?.cell, { row: 3, column: 0 });
    assert.deepEqual(sheet.reportSheet?.pagination.repeatHeaderRows, [0, 3]);

    const before = structuredClone(sheet.reportSheet);
    const value = sheet.cells.get(3, 0);
    assert.throws(() => StructuralTransform.apply(workbook, {
      kind: 'delete-rows', sheetId: sheet.id, at: 3, count: 1,
    }), /removes report binding/);
    assert.deepEqual(sheet.reportSheet, before);
    assert.equal(sheet.cells.get(3, 0), value);
  });

  it('maps report binding anchors before cell-shift mutation and rejects removal', () => {
    const { workbook } = seedWorkbook();
    const sheet = workbook.getSheet('s1');
    sheet.reportSheet = reportDefinition(sheet.id, [{ row: 3, column: 1 }]);

    StructuralTransform.apply(workbook, {
      kind: 'cell-shift', sheetId: sheet.id,
      sourceRange: { sheetId: sheet.id, startRow: 1, endRow: 1, startColumn: 1, endColumn: 1 },
      operation: 'insert', axis: 'row',
    });
    assert.deepEqual(sheet.reportSheet?.bindings[0]?.cell, { row: 4, column: 1 });

    sheet.reportSheet = reportDefinition(sheet.id, [{ row: 1, column: 1 }]);
    sheet.cells.set(1, 1, { value: 'must remain' });
    const before = structuredClone(sheet.reportSheet);
    assert.throws(() => StructuralTransform.apply(workbook, {
      kind: 'cell-shift', sheetId: sheet.id,
      sourceRange: { sheetId: sheet.id, startRow: 1, endRow: 1, startColumn: 1, endColumn: 1 },
      operation: 'delete', axis: 'row',
    }), /removes report binding/);
    assert.deepEqual(sheet.reportSheet, before);
    assert.equal(sheet.cells.get(1, 1)?.value, 'must remain');
  });

  it('moves report bindings with source cells and rejects overwriting a destination binding', () => {
    const { workbook } = seedWorkbook();
    const sheet = workbook.getSheet('s1');
    sheet.cells.set(0, 0, { value: 'source' });
    sheet.reportSheet = reportDefinition(sheet.id, [{ row: 0, column: 0 }]);

    StructuralTransform.apply(workbook, {
      kind: 'move-range', sheetId: sheet.id,
      sourceRange: { sheetId: sheet.id, startRow: 0, endRow: 0, startColumn: 0, endColumn: 0 },
      targetOrigin: { row: 2, column: 2 },
    });
    assert.deepEqual(sheet.reportSheet?.bindings[0]?.cell, { row: 2, column: 2 });

    sheet.cells.set(3, 3, { value: 'another source' });
    sheet.reportSheet = reportDefinition(sheet.id, [{ row: 4, column: 4 }]);
    const before = structuredClone(sheet.reportSheet);
    assert.throws(() => StructuralTransform.apply(workbook, {
      kind: 'move-range', sheetId: sheet.id,
      sourceRange: { sheetId: sheet.id, startRow: 3, endRow: 3, startColumn: 3, endColumn: 3 },
      targetOrigin: { row: 4, column: 4 },
    }), /report binding.*overwritten/);
    assert.deepEqual(sheet.reportSheet, before);
    assert.equal(sheet.cells.get(3, 3)?.value, 'another source');
  });

  it('rewrites persisted non-cell formula owners and their template anchors on axis edits', () => {
    const workbook = new WorkbookModel('unit-persisted-formula-owners', 'Persisted Formula Owners');
    const sheet = workbook.getSheet('sheet-1');
    sheet.name = 'Input Sheet';
    sheet.tableSheet = {
      viewId: 'formula-view',
      columns: [{ fieldId: 'table-calc', caption: 'Table Calc', formula: '=A1' }],
      grouping: [],
    };
    sheet.drawingPayloads.set('formula-shape', {
      kind: 'shape', type: 'rectangle', fill: '#ffffff', stroke: '#000000', propertyFormula: '=A1',
    });
    workbook.dataModel.views.set('formula-view', {
      id: 'formula-view', name: 'Formula View', tableId: 'source-table',
      fields: [{ fieldId: 'view-calc', caption: 'View Calc', formula: "='Input Sheet'!A1" }],
    });
    workbook.setCellStyleTemplate({
      id: 'formula-template', name: 'Formula Template', style: {},
      dataValidation: {
        type: 'custom',
        formulaAnchor: { sheetId: sheet.id, row: 0, column: 0 },
        formula1: '=A1',
        formula2: '=B1',
        listSource: { kind: 'formula', formula: '=C1' },
      },
    });

    StructuralTransform.apply(workbook, { kind: 'insert-rows', sheetId: sheet.id, at: 0, count: 1 });

    assert.equal(sheet.tableSheet?.columns[0]?.formula, '=A2');
    assert.equal((sheet.drawingPayloads.get('formula-shape') as { propertyFormula?: string }).propertyFormula, '=A2');
    assert.equal(workbook.dataModel.views.get('formula-view')?.fields[0]?.formula, "='Input Sheet'!A2");
    const validation = workbook.cellStyleTemplates.get('formula-template')?.dataValidation;
    assert.equal(validation?.formulaAnchor?.row, 1);
    assert.equal(validation?.formula1, '=A2');
    assert.equal(validation?.formula2, '=B2');

    StructuralTransform.apply(workbook, {
      kind: 'cell-shift', sheetId: sheet.id,
      sourceRange: { sheetId: sheet.id, startRow: 1, endRow: 1, startColumn: 0, endColumn: 0 },
      operation: 'insert', axis: 'row',
    });
    assert.equal(sheet.tableSheet?.columns[0]?.formula, '=A3');
    assert.equal((sheet.drawingPayloads.get('formula-shape') as { propertyFormula?: string }).propertyFormula, '=A3');
    assert.equal(workbook.dataModel.views.get('formula-view')?.fields[0]?.formula, "='Input Sheet'!A3");
    assert.equal(validation?.formulaAnchor?.row, 2);
    assert.equal(validation?.formula1, '=A3');

    StructuralTransform.apply(workbook, {
      kind: 'move-range', sheetId: sheet.id,
      sourceRange: { sheetId: sheet.id, startRow: 2, endRow: 2, startColumn: 0, endColumn: 0 },
      targetOrigin: { row: 3, column: 1 },
    });
    assert.equal(sheet.tableSheet?.columns[0]?.formula, '=B4');
    assert.equal((sheet.drawingPayloads.get('formula-shape') as { propertyFormula?: string }).propertyFormula, '=B4');
    assert.equal(workbook.dataModel.views.get('formula-view')?.fields[0]?.formula, "='Input Sheet'!B4");
    assert.deepEqual(validation?.formulaAnchor, { sheetId: sheet.id, row: 3, column: 1 });
    assert.equal(validation?.formula1, '=B4');
    assert.equal(validation?.formula2, '=B2');
    assert.equal(validation?.listSource?.kind === 'formula' ? validation.listSource.formula : undefined, '=C2');
  });

  it('rejects deletion of an anchored style-template formula before changing coordinates', () => {
    const workbook = new WorkbookModel('unit-template-anchor-rejection', 'Template Anchor Rejection');
    const sheet = workbook.getSheet('sheet-1');
    sheet.cells.set(2, 0, { value: 'preserved' });
    workbook.setCellStyleTemplate({
      id: 'formula-template', name: 'Formula Template', style: {},
      dataValidation: {
        type: 'custom',
        formulaAnchor: { sheetId: sheet.id, row: 0, column: 0 },
        formula1: '=A1',
      },
    });
    const rowCount = sheet.rowCount;

    assert.throws(
      () => StructuralTransform.apply(workbook, { kind: 'delete-rows', sheetId: sheet.id, at: 0, count: 1 }),
      /removes cell-style-template:formula-template formula anchor/,
    );
    assert.equal(sheet.cells.get(2, 0)?.value, 'preserved');
    assert.equal(sheet.rowCount, rowCount);
    assert.equal(workbook.cellStyleTemplates.get('formula-template')?.dataValidation?.formulaAnchor?.row, 0);
  });

  it('rejects invalid persisted style-template formula anchors before structural mutation', () => {
    for (const formulaAnchor of [
      null,
      { sheetId: 'missing-sheet', row: 0, column: 0 },
      { sheetId: 'sheet-1', row: 1_048_576, column: 0 },
    ]) {
      const workbook = new WorkbookModel('unit-invalid-template-anchor', 'Invalid Template Anchor');
      const sheet = workbook.getSheet('sheet-1');
      sheet.cells.set(2, 0, { value: 'preserved' });
      workbook.setCellStyleTemplate({
        id: 'formula-template', name: 'Formula Template', style: {},
        dataValidation: { type: 'custom', formulaAnchor: formulaAnchor as { sheetId: string; row: number; column: number }, formula1: '=A1' },
      });
      const rowCount = sheet.rowCount;

      assert.throws(
        () => StructuralTransform.apply(workbook, { kind: 'insert-rows', sheetId: sheet.id, at: 0, count: 1 }),
        /STRUCTURAL_REFERENCE_OWNER_INVALID: cell-style-template:formula-template formula anchor is invalid/,
      );
      assert.equal(sheet.cells.get(2, 0)?.value, 'preserved');
      assert.equal(sheet.rowCount, rowCount);
    }
  });

  it('keeps imported OOXML formula provenance aligned with structural reference rewrites', () => {
    const workbook = new WorkbookModel('unit-formula-provenance-structure', 'Formula Provenance Structure');
    const sheet = workbook.getSheet('sheet-1');
    sheet.cells.set(4, 1, {
      value: null,
      formula: '=SUM(A1:A1)',
      formulaMetadata: { kind: 'normal', sourceFormula: '=_xlfn.SUM(A1:A1)' },
    });

    StructuralTransform.apply(workbook, { kind: 'insert-rows', sheetId: sheet.id, at: 0, count: 1 });

    assert.equal(sheet.cells.get(5, 1)?.formula, '=SUM(A2:A2)');
    assert.equal(sheet.cells.get(5, 1)?.formulaMetadata?.sourceFormula, '=_XLFN.SUM(A2:A2)');
  });

  it('rewrites provenance-only and barcode formula owners at the same cell address', () => {
    const workbook = new WorkbookModel('unit-auxiliary-formula-owners', 'Auxiliary Formula Owners');
    const sheet = workbook.getSheet('sheet-1');
    const owner = { sheetId: sheet.id, row: 4, column: 1 };
    const sheetOrder = workbook.sheetOrder.map((id) => ({ id, name: workbook.getSheet(id).name }));
    sheet.cells.set(owner.row, owner.column, {
      value: null,
      formulaMetadata: { kind: 'normal', sourceFormula: '=A1' },
      presentation: {
        kind: 'barcode',
        symbology: 'qr',
        source: { kind: 'formula', formula: '=A1' },
        parameters: { symbology: 'qr' },
        options: { foreground: '#000000', background: '#ffffff', showText: true, labelPosition: 'below', quietZone: 2 },
      },
    });
    const index = new RangeIndex(sheetOrder);
    for (const [sourceId, formula] of [
      ['structural:formula-provenance', '=A1'],
      ['structural:barcode', '=A1'],
    ] as const) {
      index.setStructuralReference(owner, sourceId, collectFormulaDependencies(parseFormula(formula), owner, { sheetOrder }));
    }

    CoreStructuralTransform.apply(workbook, { kind: 'insert-rows', sheetId: sheet.id, at: 0, count: 1 }, index);

    const moved = sheet.cells.get(owner.row + 1, owner.column);
    assert.equal(moved?.formulaMetadata?.sourceFormula, '=A2');
    assert.equal(moved?.presentation?.kind === 'barcode' && moved.presentation.source.kind === 'formula' ? moved.presentation.source.formula : undefined, '=A2');
  });

  it('rejects preserved-only auxiliary formulas before changing worksheet coordinates', () => {
    const workbook = new WorkbookModel('unit-preserved-only-auxiliary-formula', 'Preserved Formula Owner');
    const sheet = workbook.getSheet('sheet-1');
    const owner = { sheetId: sheet.id, row: 4, column: 1 };
    const sheetOrder = workbook.sheetOrder.map((id) => ({ id, name: workbook.getSheet(id).name }));
    sheet.cells.set(owner.row, owner.column, {
      value: null,
      formulaMetadata: { kind: 'array', preservedOnly: true, sourceFormula: '=A1' },
    });
    const index = new RangeIndex(sheetOrder);
    index.setStructuralReference(owner, 'structural:formula-provenance', collectFormulaDependencies(parseFormula('=A1'), owner, { sheetOrder }));
    const originalRowCount = sheet.rowCount;

    assert.throws(
      () => CoreStructuralTransform.apply(workbook, { kind: 'insert-rows', sheetId: sheet.id, at: 0, count: 1 }, index),
      /formula group.*requires an explicit formula-group transform/,
    );
    assert.equal(sheet.cells.get(owner.row, owner.column)?.formulaMetadata?.sourceFormula, '=A1');
    assert.equal(sheet.cells.get(owner.row + 1, owner.column), undefined);
    assert.equal(sheet.rowCount, originalRowCount);
  });

  it('keeps OOXML formula provenance aligned through cell-shift rewrites', () => {
    const workbook = new WorkbookModel('unit-cell-shift-formula-provenance', 'Cell Shift Formula Provenance');
    const sheet = workbook.getSheet('sheet-1');
    sheet.cells.set(1, 1, {
      value: null,
      formula: '=SUM(A1:A1)',
      formulaMetadata: { kind: 'normal', sourceFormula: '=_xlfn.SUM(A1:A1)' },
    });

    StructuralTransform.apply(workbook, {
      kind: 'cell-shift',
      sheetId: sheet.id,
      sourceRange: { sheetId: sheet.id, startRow: 0, endRow: 0, startColumn: 0, endColumn: 2 },
      operation: 'insert',
      axis: 'row',
    });

    assert.equal(sheet.cells.get(2, 1)?.formula, '=SUM(A2:A2)');
    assert.equal(sheet.cells.get(2, 1)?.formulaMetadata?.sourceFormula, '=_XLFN.SUM(A2:A2)');
  });

  it('rewrites OOXML formula provenance for moved formulas and external formula owners', () => {
    const workbook = new WorkbookModel('unit-move-formula-provenance', 'Move Formula Provenance');
    const sheet = workbook.getSheet('sheet-1');
    sheet.cells.set(0, 0, { value: 7 });
    sheet.cells.set(0, 1, {
      value: null,
      formula: '=SUM(A1:A1)',
      formulaMetadata: { kind: 'normal', sourceFormula: '=_xlfn.SUM(A1:A1)' },
    });
    sheet.cells.set(4, 3, {
      value: null,
      formula: '=SUM(A1:A1)',
      formulaMetadata: { kind: 'normal', sourceFormula: '=_xlfn.SUM(A1:A1)' },
    });

    StructuralTransform.apply(workbook, {
      kind: 'move-range',
      sheetId: sheet.id,
      sourceRange: { sheetId: sheet.id, startRow: 0, endRow: 0, startColumn: 0, endColumn: 1 },
      targetOrigin: { row: 2, column: 2 },
    });

    assert.equal(sheet.cells.get(2, 3)?.formula, '=SUM(C3:C3)');
    assert.equal(sheet.cells.get(2, 3)?.formulaMetadata?.sourceFormula, '=_XLFN.SUM(C3:C3)');
    assert.equal(sheet.cells.get(4, 3)?.formula, '=SUM(C3:C3)');
    assert.equal(sheet.cells.get(4, 3)?.formulaMetadata?.sourceFormula, '=_XLFN.SUM(C3:C3)');
  });

  it('rejects structural edits that would invalidate OOXML formula-group ownership before mutation', () => {
    const workbook = new WorkbookModel('unit-formula-group-structure', 'Formula Group Structure');
    const sheet = workbook.getSheet('sheet-1');
    sheet.cells.set(4, 0, {
      value: null,
      formula: '=1',
      formulaMetadata: { kind: 'shared', sharedIndex: 3, sharedMaster: true, range: 'A5:A6' },
    });

    assert.throws(
      () => StructuralTransform.apply(workbook, { kind: 'insert-rows', sheetId: sheet.id, at: 0, count: 1 }),
      /formula metadata.*requires an explicit formula-group operation/,
    );
    assert.throws(() => StructuralTransform.apply(workbook, {
      kind: 'move-range',
      sheetId: sheet.id,
      sourceRange: { sheetId: sheet.id, startRow: 4, endRow: 4, startColumn: 0, endColumn: 0 },
      targetOrigin: { row: 6, column: 2 },
    }), /formula metadata.*requires an explicit formula-group operation/);
    assert.equal(sheet.cells.get(4, 0)?.formula, '=1');
    assert.equal(sheet.cells.get(4, 0)?.formulaMetadata?.range, 'A5:A6');
  });

  it('cell-shift insert moves the complete affected band by the selection extent', () => {
    const workbook = new WorkbookModel('unit-shift-cells', 'Shift Cells');
    const sheet = workbook.getSheet('sheet-1');
    sheet.cells.set(1, 1, { value: null, formula: '=A1+B1' });
    sheet.review.setNote(1, 1, { id: 'n1', author: 'u', text: 'note', createdAt: 'now', visible: true });
    sheet.review.addThread({ id: 'c1', sheetId: sheet.id, row: 1, column: 1, author: 'u', text: 'comment', createdAt: 'now', replies: [] });

    StructuralTransform.apply(workbook, {
      kind: 'cell-shift',
      sheetId: sheet.id,
      sourceRange: { sheetId: sheet.id, startRow: 0, endRow: 0, startColumn: 0, endColumn: 2 },
      operation: 'insert',
      axis: 'row',
    });

    assert.equal(sheet.cells.get(2, 1)?.formula, '=A2+B2');
    assert.equal(sheet.cells.get(1, 1), undefined);
    assert.ok(sheet.review.hasNoteAt(2, 1));
    assert.equal(sheet.review.getThreadsAt(2, 1)[0]?.row, 2);
  });

  it('cell-shift uses the selected height and width for insert and delete', () => {
    const workbook = new WorkbookModel('unit-shift-extent', 'Shift Extent');
    const sheet = workbook.getSheet('sheet-1');
    sheet.rowCount = 8;
    sheet.columnCount = 8;
    sheet.cells.set(3, 0, { value: 'row-source' });
    sheet.cells.set(0, 3, { value: 'column-source' });

    StructuralTransform.apply(workbook, {
      kind: 'cell-shift', sheetId: sheet.id,
      sourceRange: { sheetId: sheet.id, startRow: 1, endRow: 2, startColumn: 0, endColumn: 1 },
      operation: 'insert', axis: 'row',
    });
    assert.equal(sheet.cells.get(5, 0)?.value, 'row-source');
    assert.equal(sheet.cells.get(3, 0), undefined);

    StructuralTransform.apply(workbook, {
      kind: 'cell-shift', sheetId: sheet.id,
      sourceRange: { sheetId: sheet.id, startRow: 0, endRow: 1, startColumn: 1, endColumn: 2 },
      operation: 'insert', axis: 'column',
    });
    assert.equal(sheet.cells.get(0, 5)?.value, 'column-source');

    StructuralTransform.apply(workbook, {
      kind: 'cell-shift', sheetId: sheet.id,
      sourceRange: { sheetId: sheet.id, startRow: 1, endRow: 2, startColumn: 0, endColumn: 1 },
      operation: 'delete', axis: 'row',
    });
    assert.equal(sheet.cells.get(3, 0)?.value, 'row-source');
    assert.equal(sheet.cells.get(5, 0), undefined);

    StructuralTransform.apply(workbook, {
      kind: 'cell-shift', sheetId: sheet.id,
      sourceRange: { sheetId: sheet.id, startRow: 0, endRow: 1, startColumn: 1, endColumn: 2 },
      operation: 'delete', axis: 'column',
    });
    assert.equal(sheet.cells.get(0, 3)?.value, 'column-source');
  });

  it('structural row shifts update chart source ranges in the canonical drawing payload', () => {
    const workbook = new WorkbookModel('unit-drawing-structure', 'Drawing Structure');
    const sheet = workbook.getSheet('sheet-1');
    sheet.drawings.push({
      id: 'drawing-chart-1',
      sheetId: sheet.id,
      kind: 'chart',
      payloadId: 'chart-1',
      anchor: { kind: 'absolute' },
      transform: { x: 0, y: 0, width: 100, height: 80 },
      zIndex: 1,
    });
    sheet.drawingPayloads.set('chart-1', {
      kind: 'chart',
      chartId: 'chart-1',
      chartType: 'combo',
      subtype: 'custom-combo',
      stacked: 'percent',
      source: { kind: 'worksheet-ranges', ranges: [{ sheetId: sheet.id, startRow: 1, endRow: 3, startColumn: 0, endColumn: 1 }] },
      categoryRange: { sheetId: sheet.id, startRow: 1, endRow: 3, startColumn: 0, endColumn: 0 },
      series: [{ name: 'Sales', range: { sheetId: sheet.id, startRow: 1, endRow: 3, startColumn: 1, endColumn: 1 } }],
      elements: { hiddenData: 'show' },
    });

    StructuralTransform.apply(workbook, { kind: 'insert-rows', sheetId: sheet.id, at: 1, count: 2 });

    const payload = sheet.drawingPayloads.get('chart-1');
    assert.equal(payload?.kind, 'chart');
    if (payload?.kind !== 'chart') throw new Error('Expected chart payload');
    assert.equal(payload.chartType, 'combo');
    assert.equal(payload.stacked, 'percent');
    assert.deepEqual(payload.source.kind === 'worksheet-ranges' ? payload.source.ranges[0] : undefined, { sheetId: sheet.id, startRow: 3, endRow: 5, startColumn: 0, endColumn: 1 });
    assert.deepEqual(payload.categoryRange, { sheetId: sheet.id, startRow: 3, endRow: 5, startColumn: 0, endColumn: 0 });
    assert.deepEqual(payload.series?.[0]?.range, { sheetId: sheet.id, startRow: 3, endRow: 5, startColumn: 1, endColumn: 1 });
    assert.equal(sheet.drawings.filter((drawing) => drawing.kind === 'chart').length, 1);
  });

  it('structural transforms keep sheet-backed workbook table sources aligned', () => {
    const workbook = new WorkbookModel('unit-workbook-table-structure', 'Workbook Table Structure');
    const sheet = workbook.getSheet('sheet-1');
    workbook.addTable({
      id: 'table-1',
      name: 'Sales',
      sourceSheetId: sheet.id,
      sourceRange: { sheetId: sheet.id, startRow: 1, endRow: 4, startColumn: 2, endColumn: 4 },
      rowCount: 3,
      fields: [],
      blockSize: 128,
      blocks: [],
      revision: 0,
    });

    StructuralTransform.apply(workbook, { kind: 'insert-rows', sheetId: sheet.id, at: 1, count: 2 });
    assert.deepEqual(workbook.getTable('table-1').sourceRange, {
      sheetId: sheet.id,
      startRow: 3,
      endRow: 6,
      startColumn: 2,
      endColumn: 4,
    });

    assert.throws(
      () => StructuralTransform.apply(workbook, { kind: 'delete-rows', sheetId: sheet.id, at: 3, count: 1 }),
      /workbook table table-1 requires an explicit table operation/,
    );
  });

  it('move-range clears stale destinations, offsets formulas, and rewrites external references', () => {
    const workbook = new WorkbookModel('unit-move-range', 'Move Range');
    const sheet = workbook.getSheet('sheet-1');
    sheet.cells.set(0, 0, { value: 7 });
    sheet.cells.set(0, 1, { value: null, formula: '=A1' });
    sheet.cells.set(2, 3, { value: 'stale' });
    sheet.cells.set(0, 3, { value: null, formula: '=A1' });

    StructuralTransform.apply(workbook, {
      kind: 'move-range',
      sheetId: sheet.id,
      sourceRange: { sheetId: sheet.id, startRow: 0, endRow: 0, startColumn: 0, endColumn: 1 },
      targetOrigin: { row: 2, column: 2 },
    });

    assert.equal(sheet.cells.get(2, 2)?.value, 7);
    assert.equal(sheet.cells.get(2, 3)?.formula, '=C3');
    assert.equal(sheet.cells.get(0, 0), undefined);
    assert.equal(sheet.cells.get(0, 3)?.formula, '=C3');
  });

  it('rewrites rule formulas and hyperlink addresses on other worksheets when a referenced range moves', () => {
    const workbook = new WorkbookModel('unit-move-metadata-references', 'Move Metadata References');
    const sheet = workbook.getSheet('sheet-1');
    const other = workbook.addSheet('sheet-2', 'Other');
    other.conditionalFormats.push({
      id: 'cf-1', sheetId: other.id,
      ranges: [{ sheetId: other.id, startRow: 0, endRow: 0, startColumn: 0, endColumn: 0 }],
      type: 'highlight', operator: 'formula', value1: '=Sheet1!A1',
    });
    other.dataValidations.push({
      id: 'dv-1', sheetId: other.id,
      ranges: [{ sheetId: other.id, startRow: 0, endRow: 0, startColumn: 0, endColumn: 0 }],
      type: 'custom', formula1: '=Sheet1!A1',
    });
    other.hyperlinks.set('0:0', {
      id: 'link-1', target: { kind: 'sheet', sheetId: sheet.id, address: 'A1' },
    });

    StructuralTransform.apply(workbook, {
      kind: 'move-range', sheetId: sheet.id,
      sourceRange: { sheetId: sheet.id, startRow: 0, endRow: 0, startColumn: 0, endColumn: 0 },
      targetOrigin: { row: 2, column: 2 },
    });

    assert.equal(other.conditionalFormats[0]?.value1, '=Sheet1!C3');
    assert.equal(other.dataValidations[0]?.formula1, '=Sheet1!C3');
    const target = other.hyperlinks.get('0:0')?.target;
    assert.equal(target?.kind === 'sheet' ? target.address : undefined, 'C3');
  });

  it('rejects moving over destination-anchored hyperlinks without mutating either range', () => {
    const workbook = new WorkbookModel('unit-move-hyperlink-reject', 'Move Hyperlink Reject');
    const sheet = workbook.getSheet('sheet-1');
    sheet.cells.set(0, 0, { value: 'source' });
    sheet.hyperlinks.set('1:1', { id: 'destination-link', target: { kind: 'url', url: 'https://example.com' } });

    assert.throws(() => StructuralTransform.apply(workbook, {
      kind: 'move-range',
      sheetId: sheet.id,
      sourceRange: { sheetId: sheet.id, startRow: 0, endRow: 0, startColumn: 0, endColumn: 0 },
      targetOrigin: { row: 1, column: 1 },
    }), /hyperlink 1:1 would be overwritten at the target/);
    assert.equal(sheet.cells.get(0, 0)?.value, 'source');
    assert.equal(sheet.hyperlinks.has('1:1'), true);
  });

  it('rejects a cell shift that would silently drop an anchored object', () => {
    const workbook = new WorkbookModel('unit-shift-reject', 'Shift Reject');
    const sheet = workbook.getSheet('sheet-1');
    sheet.cells.set(1, 0, { value: 'keep' });
    sheet.drawings.push({
      id: 'drawing-edge',
      sheetId: sheet.id,
      kind: 'shape',
      payloadId: 'shape-edge',
      anchor: { kind: 'one-cell', row: 1, column: 0 },
      transform: { x: 0, y: 0, width: 20, height: 20 },
      zIndex: 0,
    });
    sheet.drawingPayloads.set('shape-edge', { kind: 'shape', type: 'rectangle', fill: '#ffffff', stroke: '#000000' });
    sheet.rowCount = 2;
    assert.throws(() => StructuralTransform.apply(workbook, {
      kind: 'cell-shift',
      sheetId: sheet.id,
      sourceRange: { sheetId: sheet.id, startRow: 0, endRow: 1, startColumn: 0, endColumn: 0 },
      operation: 'insert',
      axis: 'row',
    }), /outside worksheet bounds|discard data/);
    assert.equal(sheet.cells.get(1, 0)?.value, 'keep');
    assert.equal(sheet.drawings[0]?.anchor.row, 1);
  });

  it('moves complete data regions when rows are inserted before them and keeps manifest coordinates aligned', () => {
    const workbook = new WorkbookModel('unit-data-region-shift', 'Data Region Shift');
    const sheet = workbook.getSheet('sheet-1');
    const sourceId = 'source-structure';
    workbook.addDataSource({
      schema: 'DataSourceManifest',
      version: 1,
      id: sourceId,
      name: 'Structure source',
      kind: 'worksheet-range',
      sourceSheetId: sheet.id,
      sourceRange: { sheetId: sheet.id, startRow: 5, endRow: 7, startColumn: 2, endColumn: 3 },
      rowCount: 2,
      fields: [
        { id: 'f0', name: 'Code', ordinal: 0, type: 'text' },
        { id: 'f1', name: 'Value', ordinal: 1, type: 'number' },
      ],
      blockRowCount: 65_536,
      blocks: [{
        id: 'structure-block', dataSourceId: sourceId, startRow: 0, rowCount: 2,
        storageKey: 'structure-block', checksum: 'a'.repeat(64), byteLength: 1,
        encoding: 'columnar-v1', revision: 0,
      }],
      revision: 0,
    });
    sheet.addDataRegion({
      id: 'structure-region',
      sourceId,
      range: { sheetId: sheet.id, startRow: 5, endRow: 7, startColumn: 2, endColumn: 3 },
      headerRow: 5,
      revision: 0,
    });
    sheet.cells.set(6, 2, { value: 999, style: { bold: true } });

    StructuralTransform.apply(workbook, { kind: 'insert-rows', sheetId: sheet.id, at: 2, count: 2 });

    assert.deepEqual(sheet.dataRegions[0]?.range, {
      sheetId: sheet.id, startRow: 7, endRow: 9, startColumn: 2, endColumn: 3,
    });
    assert.equal(sheet.dataRegions[0]?.headerRow, 7);
    assert.deepEqual(workbook.getDataSource(sourceId).sourceRange, {
      sheetId: sheet.id, startRow: 7, endRow: 9, startColumn: 2, endColumn: 3,
    });
    assert.equal(sheet.cells.get(8, 2)?.value, 999);
  });

  it('rejects row or column edits that intersect a block-backed region until a block transaction is supplied', () => {
    const workbook = new WorkbookModel('unit-data-region-reject', 'Data Region Reject');
    const sheet = workbook.getSheet('sheet-1');
    sheet.addDataRegion({
      id: 'region-reject',
      sourceId: 'source-reject',
      range: { sheetId: sheet.id, startRow: 5, endRow: 8, startColumn: 2, endColumn: 4 },
      headerRow: 5,
      revision: 0,
    });
    assert.throws(
      () => StructuralTransform.apply(workbook, { kind: 'insert-rows', sheetId: sheet.id, at: 6, count: 1 }),
      /data region region-reject requires a data-block transaction/,
    );
    assert.throws(
      () => StructuralTransform.apply(workbook, { kind: 'delete-columns', sheetId: sheet.id, at: 3, count: 1 }),
      /data region region-reject requires a data-block transaction/,
    );
    assert.deepEqual(sheet.dataRegions[0]?.range, {
      sheetId: sheet.id, startRow: 5, endRow: 8, startColumn: 2, endColumn: 4,
    });
  });

  it('moves a complete block-backed region as metadata while preserving its immutable source rows', () => {
    const workbook = new WorkbookModel('unit-data-region-move', 'Data Region Move');
    const sheet = workbook.getSheet('sheet-1');
    const sourceId = 'source-move';
    workbook.addDataSource({
      schema: 'DataSourceManifest', version: 1, id: sourceId, name: 'Move source', kind: 'worksheet-range',
      sourceSheetId: sheet.id,
      sourceRange: { sheetId: sheet.id, startRow: 2, endRow: 4, startColumn: 0, endColumn: 1 },
      rowCount: 2,
      fields: [{ id: 'f0', name: 'Code', ordinal: 0, type: 'text' }],
      blockRowCount: 65_536,
      blocks: [{ id: 'move-block', dataSourceId: sourceId, startRow: 0, rowCount: 2, storageKey: 'move-block', checksum: 'b'.repeat(64), byteLength: 1, encoding: 'columnar-v1', revision: 0 }],
      revision: 0,
    });
    sheet.addDataRegion({
      id: 'region-move', sourceId,
      range: { sheetId: sheet.id, startRow: 2, endRow: 4, startColumn: 0, endColumn: 1 }, headerRow: 2, revision: 0,
    });
    sheet.cells.set(3, 0, { value: null, style: { italic: true } });

    StructuralTransform.apply(workbook, {
      kind: 'move-range', sheetId: sheet.id,
      sourceRange: { sheetId: sheet.id, startRow: 2, endRow: 4, startColumn: 0, endColumn: 1 },
      targetOrigin: { row: 10, column: 2 },
    });

    assert.deepEqual(sheet.dataRegions[0]?.range, {
      sheetId: sheet.id, startRow: 10, endRow: 12, startColumn: 2, endColumn: 3,
    });
    assert.equal(sheet.dataRegions[0]?.headerRow, 10);
    assert.deepEqual(workbook.getDataSource(sourceId).sourceRange, {
      sheetId: sheet.id, startRow: 10, endRow: 12, startColumn: 2, endColumn: 3,
    });
    assert.equal(sheet.cells.get(11, 2)?.style?.italic, true);
  });
});
