import assert from 'node:assert/strict';
import { describe, it } from 'node:test';
import { WorkbookModel } from '@react-sheets/core-model';
import { createPasteSpecialSpec } from '@react-sheets/sheet-features';
import { createSpillEnvironment } from './formula-spill-sync';
import { hydrateRuntime } from './runtime';
import { createRemoteReadySessionFixture } from './session-test-fixtures';
import { WorkbookSession } from './workbook-session';

function cellValue(app: WorkbookSession, row: number, column: number): string {
  return app.getUiSnapshot().selectedSheet.getCell(row, column)?.value ?? '';
}

describe('WorkbookSession formula integration', () => {
  it('preserves inactive sparse sheets through load, row insertion, undo and redo', async () => {
    const app = createRemoteReadySessionFixture();
    try {
      const runtime = app['runtime'];
      const target = runtime.model.getSheet('sheet-1');
      target.cells.set(4, 4, { value: 42 });
      const other = runtime.model.addSheet('lazy-reference', 'LazyReference', 600_001, 16_384);
      other.cells.set(2, 2, { value: null, formula: '=Sheet1!$E$5' });
      other.cells.set(600_000, 16_383, { value: 'sparse tail' });
      const unused = runtime.model.addSheet('lazy-unused', 'LazyUnused', 600_001, 16_384);
      unused.cells.set(600_000, 16_383, { value: 'unreferenced' });
      hydrateRuntime(runtime, { snapshot: runtime.model.snapshot(), revision: 0 });
      await app.waitForFormulaCalculation();
      const referenced = runtime.model.getSheet(other.id);
      const unreferenced = runtime.model.getSheet(unused.id);
      const owner = { sheetId: other.id, row: 2, column: 2 };
      const engine = runtime.formula;
      assert.equal(referenced.cells.isHydrated, false);
      assert.equal(unreferenced.cells.isHydrated, false);
      assert.equal(engine.getCellResult(owner)?.value, 42);

      app.runCommand('sheet.rows.insert', { sheetId: target.id, at: 2, count: 1 });
      await app.waitForFormulaCalculation();
      assert.equal(runtime.formula, engine);
      assert.equal(referenced.cells.getFormulaOwnerWithoutHydration(2, 2)?.formula, '=Sheet1!$E$6');
      assert.equal(engine.getCellResult(owner)?.value, 42);
      assert.equal(referenced.cells.isHydrated, false);
      assert.equal(unreferenced.cells.isHydrated, false);

      app.undo();
      await app.waitForFormulaCalculation();
      assert.equal(referenced.cells.getFormulaOwnerWithoutHydration(2, 2)?.formula, '=Sheet1!$E$5');
      assert.equal(engine.getCellResult(owner)?.value, 42);
      app.redo();
      await app.waitForFormulaCalculation();
      assert.equal(runtime.formula, engine);
      assert.equal(referenced.cells.getFormulaOwnerWithoutHydration(2, 2)?.formula, '=Sheet1!$E$6');
      assert.equal(engine.getCellResult(owner)?.value, 42);
      assert.equal(referenced.cells.isHydrated, false);
      assert.equal(unreferenced.cells.isHydrated, false);
    } finally {
      app.dispose();
    }
  });

  it('loads every ordinary input for the first formula without materializing its source sheet', async () => {
    const app = new WorkbookSession();
    try {
      const runtime = app['runtime'];
      runtime.model.getSheet('sheet-1').cells.set(0, 0, { value: 20 });
      const inputs = runtime.model.addSheet('lazy-inputs', 'Inputs', 600_001, 16_384);
      inputs.cells.set(0, 0, { value: 10 });
      inputs.cells.set(600_000, 16_383, { value: 7 });
      hydrateRuntime(runtime, { snapshot: runtime.model.snapshot(), revision: 0 });
      await app.waitForFormulaCalculation();
      assert.equal(runtime.model.getSheet(inputs.id).cells.isHydrated, false);
      app.runCommand('sheet.cell.set', { sheetId: 'sheet-1', row: 0, column: 1,
        value: { value: null, formula: '=A1+Inputs!A1+Inputs!XFD600001' } });
      await app.waitForFormulaCalculation();
      assert.equal(runtime.formula.getCellResult({ sheetId: 'sheet-1', row: 0, column: 1 })?.value, 37);
      assert.equal(runtime.model.getSheet(inputs.id).cells.isHydrated, false);
    } finally {
      app.dispose();
    }
  });

  it('reads current spill blockers from deferred storage without hydrating the worksheet', () => {
    const workbook = new WorkbookModel('lazy-spill-occupancy', 'Lazy spill');
    const sheet = workbook.getSheet('sheet-1');
    sheet.cells.deferJSON({ '1': { '0': { value: 'blocker' } } });
    const environment = createSpillEnvironment(sheet);
    assert.equal(environment.isOccupied(1, 0), true);
    assert.equal(environment.isOccupied(1, 1), false);
    assert.equal(sheet.cells.replaceCellWithoutHydration(1, 0, { value: null }), true);
    assert.equal(environment.isOccupied(1, 0), false);
    assert.equal(sheet.cells.replaceCellWithoutHydration(1, 0, { value: null, formula: '=1' }), true);
    assert.equal(environment.isOccupied(1, 0), true);
    assert.equal(sheet.cells.isHydrated, false);
  });

  it('derives the AutoSum current region from resolved formula results', async () => {
    const app = new WorkbookSession();
    const sheetId = app.getActiveSheetId();
    for (const [row, formula] of [[0, '=10'], [1, '=20'], [2, '=30']] as const) {
      app.runCommand('sheet.cell.set', { sheetId, row, column: 0, value: { formula } });
    }
    await app.waitForFormulaCalculation();
    assert.equal(app.selectAddress('A1'), true);
    assert.deepEqual(app.getCurrentRegion(), {
      sheetId,
      startRow: 0,
      endRow: 2,
      startColumn: 0,
      endColumn: 0,
    });
  });

  it('recalculates dependent formulas automatically when source values change', async () => {
    const app = new WorkbookSession();
    const sheetId = app.getActiveSheetId();
    app.runCommand('sheet.cell.set', {
      sheetId,
      row: 0,
      column: 0,
      value: { value: 2 },
    });
    app.runCommand('sheet.cell.set', {
      sheetId,
      row: 0,
      column: 1,
      value: { formula: '=A1*3' },
    });
    await app.waitForFormulaCalculation();
    assert.equal(cellValue(app, 0, 1), '6');

    app.runCommand('sheet.cell.set', {
      sheetId,
      row: 0,
      column: 0,
      value: { value: 5 },
    });
    await app.waitForFormulaCalculation();
    assert.equal(cellValue(app, 0, 1), '15');
  });

  it('loads cached values from preserved-only formula cells before evaluating dependent formulas', async () => {
    const app = new WorkbookSession();
    const runtime = app['runtime'];
    const sheetId = app.getActiveSheetId();
    const sheet = runtime.model.getSheet(sheetId);
    sheet.cells.set(0, 0, {
      value: 11,
      formula: '=INDIRECT("A1")',
      formulaMetadata: { kind: 'normal', preservedOnly: true, reason: 'Formula is preserved without local recalculation' },
    });
    sheet.cells.set(0, 1, { value: null, formula: '=A1+1' });

    hydrateRuntime(runtime, { snapshot: runtime.model.snapshot(), revision: 0 });
    await app.waitForFormulaCalculation();

    assert.equal(cellValue(app, 0, 1), '12');
  });

  it('hydrates preserved-only formula caches when the workbook first gains a calculable formula', async () => {
    const app = new WorkbookSession();
    const runtime = app['runtime'];
    const sheetId = app.getActiveSheetId();
    runtime.model.getSheet(sheetId).cells.set(0, 0, {
      value: 11,
      formula: '=INDIRECT("A1")',
      formulaMetadata: { kind: 'normal', preservedOnly: true, reason: 'Formula is preserved without local recalculation' },
    });

    hydrateRuntime(runtime, { snapshot: runtime.model.snapshot(), revision: 0 });
    app.runCommand('sheet.cell.set', { sheetId, row: 0, column: 1, value: { formula: '=A1+1' } });
    await app.waitForFormulaCalculation();

    assert.equal(cellValue(app, 0, 1), '12');
  });

  it('synchronizes row permutations incrementally without rebuilding the formula engine', async () => {
    const app = createRemoteReadySessionFixture();
    const sheetId = app.getActiveSheetId();
    app.runCommand('sheet.cell.set', { sheetId, row: 0, column: 0, value: { value: 2 } });
    app.runCommand('sheet.cell.set', { sheetId, row: 1, column: 0, value: { value: 1 } });
    app.runCommand('sheet.cell.set', { sheetId, row: 0, column: 2, value: { formula: '=A1', value: null } });
    await app.waitForFormulaCalculation();
    assert.equal(cellValue(app, 0, 2), '2');

    const formulaEngine = app['runtime'].formula;
    app.runCommand('data.sort.rows', {
      sheetId,
      range: { sheetId, startRow: 0, endRow: 1, startColumn: 0, endColumn: 0 },
      criteria: [{ column: 0, ascending: true }],
      hasHeader: false,
    });

    assert.equal(app['runtime'].formula, formulaEngine);
    await app.waitForFormulaCalculation();
    assert.equal(cellValue(app, 0, 0), '1');
    assert.equal(cellValue(app, 1, 0), '2');
    assert.equal(cellValue(app, 0, 2), '1');

    app.undo();
    assert.equal(app['runtime'].formula, formulaEngine);
    await app.waitForFormulaCalculation();
    assert.equal(cellValue(app, 0, 0), '2');
    assert.equal(cellValue(app, 1, 0), '1');
    assert.equal(cellValue(app, 0, 2), '2');
  });

  it('manual recalculation mode defers updates until recalculateFormulas()', async () => {
    const app = new WorkbookSession();
    const sheetId = app.getActiveSheetId();
    app.setRecalculationMode('manual');
    app.runCommand('sheet.cell.set', {
      sheetId,
      row: 0,
      column: 0,
      value: { value: 2 },
    });
    app.runCommand('sheet.cell.set', {
      sheetId,
      row: 0,
      column: 1,
      value: { formula: '=A1*3' },
    });
    assert.equal(cellValue(app, 0, 1), '6');

    app.runCommand('sheet.cell.set', {
      sheetId,
      row: 0,
      column: 0,
      value: { value: 5 },
    });
    assert.equal(cellValue(app, 0, 1), '6');

    await app.recalculateFormulas();
    assert.equal(cellValue(app, 0, 1), '15');
  });

  it('resolves defined names through workbook.name.set', async () => {
    const app = new WorkbookSession();
    const sheetId = app.getActiveSheetId();
    app.runCommand('workbook.name.set', { name: 'TaxRate', value: '0.1' });
    app.runCommand('sheet.cell.set', {
      sheetId,
      row: 2,
      column: 2,
      value: { value: 100 },
    });
    app.runCommand('sheet.cell.set', {
      sheetId,
      row: 3,
      column: 3,
      value: { formula: '=C3*TaxRate' },
    });
    await app.waitForFormulaCalculation();
    assert.equal(cellValue(app, 3, 3), '10');

    app.runCommand('workbook.name.set', { name: 'TaxRate', value: '0.2' });
    await app.waitForFormulaCalculation();
    assert.equal(cellValue(app, 3, 3), '20');
  });

  it('exposes canonical defined-name CRUD and calculation state through Session APIs', async () => {
    const app = new WorkbookSession();
    const sheetId = app.getActiveSheetId();
    const created = app.setDefinedName({ name: 'LocalRate', formula: '0.25', scope: 'sheet', sheetId });
    assert.equal(created.scope, 'sheet');
    assert.equal(app.getDefinedName('LocalRate', sheetId)?.formula, '0.25');
    assert.equal(app.listDefinedNames(sheetId).some((entry) => entry.name === 'LocalRate'), true);

    app.setRecalculationMode('manual');
    assert.equal(app.getRecalculationMode(), 'manual');
    app.removeDefinedName('LocalRate', 'sheet', sheetId);
    assert.equal(app.getDefinedName('LocalRate', sheetId), undefined);
    assert.equal(app.hasPendingFormulaRecalculation(), false);
    await app.waitForFormulaCalculation();
  });

  it('resolves Go To through the active sheet scoped name before workbook scope', () => {
    const app = new WorkbookSession();
    const sheetId = app.getActiveSheetId();
    app.setDefinedName({ name: 'Target', formula: 'C3', scope: 'workbook' });
    app.setDefinedName({ name: 'Target', formula: 'D4', scope: 'sheet', sheetId });
    app.runCommand('navigation.goto', { sheetId, reference: 'Target' });
    assert.deepEqual(app.getSelection().activeCell, { row: 3, column: 3 });
  });

  it('tracks dynamic-array spill ranges and child values', async () => {
    const app = new WorkbookSession();
    const sheetId = app.getActiveSheetId();
    app.runCommand('sheet.cell.set', {
      sheetId,
      row: 0,
      column: 0,
      value: { formula: '=SEQUENCE(2,2,1,1)' },
    });

    await app.waitForFormulaCalculation();
    const sheet = app['runtime'].model.getSheet(sheetId);
    assert.equal(sheet.spillRanges.length, 1);
    assert.equal(sheet.spillRanges[0]?.state, 'ok');
    assert.equal(cellValue(app, 0, 0), '1');
    assert.equal(cellValue(app, 0, 1), '2');
    assert.equal(cellValue(app, 1, 0), '3');
    assert.equal(cellValue(app, 1, 1), '4');
  });

  it('rebuilds persisted dynamic-array spills without self-blocking', async () => {
    const app = new WorkbookSession();
    const runtime = app['runtime'];
    const sheetId = app.getActiveSheetId();
    app.runCommand('sheet.cell.set', {
      sheetId,
      row: 0,
      column: 0,
      value: { formula: '=SEQUENCE(2,2,1,1)' },
    });
    await app.waitForFormulaCalculation();

    hydrateRuntime(runtime, { snapshot: runtime.model.snapshot(), revision: 0 });
    await app.waitForFormulaCalculation();

    const spill = runtime.model.getSheet(sheetId).spillRanges[0];
    assert.equal(spill?.state, 'ok');
    assert.equal(cellValue(app, 1, 1), '4');
  });

  it('blocks spills against merged and table geometry through the canonical environment', () => {
    const app = new WorkbookSession();
    const runtime = app['runtime'];
    const sheetId = app.getActiveSheetId();
    const sheet = runtime.model.getSheet(sheetId);
    sheet.merges.push({
      range: { sheetId, startRow: 0, endRow: 0, startColumn: 2, endColumn: 2 },
      anchor: { row: 0, column: 2 },
    });
    sheet.sheetTables.push({
      id: 'table-blocker',
      sheetId,
      name: 'TableBlocker',
      range: { sheetId, startRow: 3, endRow: 3, startColumn: 3, endColumn: 3 },
      hasHeaderRow: false,
      hasTotalRow: false,
      showBandedRows: false,
      showBandedColumns: false,
      showFirstColumn: false,
      showLastColumn: false,
      showFilterButton: true,
      autoExpand: 'none',
      columns: [],
    });
    runtime.formula.setSpillEnvironment(sheetId, createSpillEnvironment(sheet));

    runtime.formula.setFormula({ sheetId, row: 0, column: 0 }, '=SEQUENCE(1,3)');
    runtime.formula.setFormula({ sheetId, row: 3, column: 0 }, '=SEQUENCE(1,4)');

    const spills = runtime.formula.getSpillsForSheet(sheetId);
    assert.deepEqual(spills.map(({ state, blocker }) => ({ state, blocker })), [
      { state: 'blocked', blocker: { row: 0, column: 2 } },
      { state: 'blocked', blocker: { row: 3, column: 3 } },
    ]);
  });

  it('synchronizes formulas after range paste and cell insert shifts', async () => {
    const app = createRemoteReadySessionFixture();
    const sheetId = app.getActiveSheetId();
    app.runCommand('sheet.cell.set', { sheetId, row: 0, column: 0, value: { value: 2 } });
    app.runCommand('sheet.cell.set', { sheetId, row: 0, column: 1, value: { formula: '=A1*3', value: null } });
    app.runCommand('sheet.range.paste', {
      sheetId,
      targetOrigin: { row: 1, column: 0 },
      clipboard: { schema: 'SparseClipboardPayload', range: { sheetId, startRow: 0, endRow: 0, startColumn: 0, endColumn: 0 }, sourceExtent: { rows: 1, columns: 1 }, occupiedCells: [{ rowOffset: 0, columnOffset: 0, value: { value: 4 } }], transfer: 'copy', rangeMetadata: { columnWidths: [], validations: [], conditionalFormats: [], notes: [], comments: [], hyperlinks: [] } },
      transfer: 'copy',
      spec: createPasteSpecialSpec(),
    });
    assert.equal(app['runtime'].model.getSheet(sheetId).cells.get(1, 0)?.value, 4);

    app.runCommand('sheet.cells.insert', {
      sheetId,
      range: { sheetId, startRow: 0, endRow: 0, startColumn: 0, endColumn: 1 },
      operation: 'insert',
      axis: 'row',
    });
    const moved = app['runtime'].model.getSheet(sheetId).cells.get(1, 1);
    assert.equal(moved?.formula, '=A2*3');
    await app.waitForFormulaCalculation();
    assert.equal(cellValue(app, 1, 1), '6');
  });

  it('synchronizes formula inputs and external references after a canonical range move', async () => {
    const app = createRemoteReadySessionFixture();
    const sheetId = app.getActiveSheetId();
    app.runCommand('sheet.cell.set', { sheetId, row: 0, column: 0, value: { value: 5 } });
    app.runCommand('sheet.cell.set', { sheetId, row: 0, column: 1, value: { formula: '=A1*2', value: null } });
    app.runCommand('sheet.cell.set', { sheetId, row: 0, column: 3, value: { formula: '=A1*3', value: null } });

    app.runCommand('sheet.range.move', {
      sheetId,
      sourceRange: { sheetId, startRow: 0, endRow: 0, startColumn: 0, endColumn: 0 },
      targetOrigin: { row: 2, column: 2 },
    });
    await app.waitForFormulaCalculation();

    const sheet = app['runtime'].model.getSheet(sheetId);
    assert.equal(sheet.cells.get(2, 2)?.value, 5);
    assert.equal(sheet.cells.get(0, 1)?.formula, '=C3*2');
    assert.equal(sheet.cells.get(0, 3)?.formula, '=C3*3');
    assert.equal(cellValue(app, 0, 1), '10');
    assert.equal(cellValue(app, 0, 3), '15');
  });

  it('synchronizes clear and emits #REF! after a deleted-row reference', async () => {
    const app = createRemoteReadySessionFixture();
    const sheetId = app.getActiveSheetId();
    app.runCommand('sheet.cell.set', { sheetId, row: 1, column: 0, value: { value: 7 } });
    app.runCommand('sheet.cell.set', { sheetId, row: 0, column: 1, value: { formula: '=A2*2', value: null } });
    app.runCommand('sheet.range.clear', {
      sheetId,
      range: { sheetId, startRow: 1, endRow: 1, startColumn: 0, endColumn: 0 },
      family: 'contents',
    });
    assert.equal(app['runtime'].model.getSheet(sheetId).cells.get(0, 1)?.formula, '=A2*2');

    app.runCommand('sheet.rows.delete', { sheetId, at: 1, count: 1 });
    const formula = app['runtime'].model.getSheet(sheetId).cells.get(0, 1)?.formula;
    assert.equal(formula, '=#REF!*2');
    await app.waitForFormulaCalculation();
    assert.equal(cellValue(app, 0, 1), '#REF!');
  });

  it('rebuilds 3-D reference ownership after worksheet reorder and undo', () => {
    const app = createRemoteReadySessionFixture();
    const firstSheetId = app.getActiveSheetId();
    app.runCommand('sheet.add', { id: 'three-d-middle', name: 'Middle' });
    app.runCommand('sheet.add', { id: 'three-d-end', name: 'End' });
    app.runCommand('sheet.cell.set', {
      sheetId: firstSheetId,
      row: 0,
      column: 0,
      value: { formula: '=SUM(Sheet1:Middle!A1)', value: null },
    });

    app.moveSheet('three-d-end', 1);
    app.moveSheet('three-d-end', 2);
    app.undo();

    assert.throws(
      () => app.runCommand('sheet.rows.insert', { sheetId: 'three-d-end', at: 0, count: 1 }),
      /UNSUPPORTED_STRUCTURAL_REFERENCE/,
    );
  });

  it('rejects direct writes into a dynamic-array spill child and rolls back the model', async () => {
    const app = new WorkbookSession();
    const sheetId = app.getActiveSheetId();
    app.runCommand('sheet.cell.set', { sheetId, row: 0, column: 0, value: { formula: '=SEQUENCE(2,2,1,1)', value: null } });
    await app.waitForFormulaCalculation();
    assert.throws(() => app.runCommand('sheet.cell.set', { sheetId, row: 0, column: 1, value: { value: 99 } }), /Spill cells are read-only/);
    assert.equal(app['runtime'].model.getSheet(sheetId).cells.get(0, 1), undefined);
    assert.equal(cellValue(app, 0, 1), '2');
  });

  it('allows editing a blocker to recover a blocked dynamic-array spill', async () => {
    const app = new WorkbookSession();
    const sheetId = app.getActiveSheetId();
    app.runCommand('sheet.cell.set', { sheetId, row: 0, column: 1, value: { value: 99 } });
    app.runCommand('sheet.cell.set', { sheetId, row: 0, column: 0, value: { formula: '=SEQUENCE(2,2,1,1)' } });
    await app.waitForFormulaCalculation();
    assert.equal(app['runtime'].model.getSheet(sheetId).spillRanges[0]?.state, 'blocked');

    app.runCommand('sheet.cell.set', { sheetId, row: 0, column: 1, value: { value: null } });
    await app.waitForFormulaCalculation();

    assert.equal(app['runtime'].model.getSheet(sheetId).spillRanges[0]?.state, 'ok');
    assert.equal(cellValue(app, 0, 1), '2');
  });
});
