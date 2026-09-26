import assert from 'node:assert/strict';
import { describe, it } from 'node:test';
import { RangeIndex } from '@react-sheets/formula-engine';
import { applyRowPermutation, createRowPermutationPlan, rowPermutationAffectedColumnEnd, type PivotModel, type RangeRef, WorkbookModel } from './index';
import type { ReportSheetDefinition } from './data-model';

function range(sheetId: string, startRow: number, endRow: number, startColumn: number, endColumn: number): RangeRef {
  return { sheetId, startRow, endRow, startColumn, endColumn };
}

function applyPermutation(workbook: WorkbookModel, selected: RangeRef, sourceRows: readonly number[]): ReturnType<typeof applyRowPermutation> {
  const referenceOwners = buildDefinedNameAnchorIndex(workbook);
  const affectedColumnEnd = rowPermutationAffectedColumnEnd(workbook, selected, referenceOwners);
  return applyRowPermutation(
    workbook,
    createRowPermutationPlan(selected, sourceRows, affectedColumnEnd),
    referenceOwners,
  );
}

function buildDefinedNameAnchorIndex(workbook: WorkbookModel): RangeIndex {
  const index = new RangeIndex();
  for (const entry of workbook.definedNameModels) {
    index.setDefinedNameReference({
      scope: entry.scope,
      name: entry.name,
      ...(entry.sheetId ? { sheetId: entry.sheetId } : {}),
    }, [], undefined, entry.anchor);
  }
  return index;
}

function reportDefinition(sheetId: string): ReportSheetDefinition {
  return {
    templateSheetId: sheetId,
    bindings: [{ cell: { row: 0, column: 0 }, expression: 'field-id', kind: 'field' }],
    pagination: { enabled: true, repeatHeaderRows: [0, 1] },
    renderMode: 'preview',
    layout: { orientation: 'portrait', marginTopPx: 0, marginRightPx: 0, marginBottomPx: 0, marginLeftPx: 0 },
    dataEntry: [],
  };
}

function pivotDefinition(id: string, targetSheetId: string, source: PivotModel['source']): PivotModel {
  return {
    schema: 'PivotDefinition',
    id,
    source,
    target: { sheetId: targetSheetId, anchor: { row: 0, column: 2 } },
    fieldCatalog: { schema: 'PivotFieldCatalog', fields: [] },
    layout: {
      rows: [], columns: [], filters: [], allowMultipleFiltersPerField: true,
      collation: { locale: 'en-US', sensitivity: 'variant', numeric: false, caseFirst: 'false' },
      values: [], subtotalLocation: 'bottom', showRowGrandTotals: true,
      showColumnGrandTotals: true, reportLayout: 'compact',
    },
    refreshPolicy: { mode: 'on-change', preserveFormatting: true, refreshOnLoad: true },
  };
}

describe('canonical row permutation metadata plan', () => {
  it('permutes report binding anchors and repeated header rows with their data rows', () => {
    const workbook = new WorkbookModel('permutation-report-sheet', 'Permutation report sheet');
    const sheet = workbook.getSheet('sheet-1');
    sheet.rowCount = 4;
    sheet.columnCount = 2;
    sheet.reportSheet = reportDefinition(sheet.id);
    sheet.reportSheet!.bindings[0]!.cell.column = 5;

    const changes = applyPermutation(workbook, range(sheet.id, 0, 1, 0, 0), [1, 0]);

    assert.deepEqual(sheet.reportSheet?.bindings[0]?.cell, { row: 1, column: 5 });
    assert.deepEqual(changes.definedNameOwnerDeltas, []);
    assert.equal(
      rowPermutationAffectedColumnEnd(workbook, range(sheet.id, 0, 1, 0, 0), buildDefinedNameAnchorIndex(workbook)),
      5,
    );
    assert.deepEqual(sheet.reportSheet?.pagination.repeatHeaderRows, [1, 0]);
  });

  it('includes indexed defined-name anchors when planning the metadata extent', () => {
    const workbook = new WorkbookModel('permutation-name-extent', 'Permutation name extent');
    const sheet = workbook.getSheet('sheet-1');
    sheet.rowCount = 2;
    sheet.columnCount = 1;
    workbook.setDefinedName({
      name: 'AnchoredOwner',
      scope: 'workbook',
      formula: '=A1',
      anchor: { sheetId: sheet.id, row: 0, column: 12 },
    });
    const indexedNames = buildDefinedNameAnchorIndex(workbook);
    let anchorQueries = 0;
    const referenceOwners = {
      getDefinedNamesAnchoredInRange: (
        sheetId: string,
        anchorRange: Pick<RangeRef, 'startRow' | 'endRow' | 'startColumn' | 'endColumn'>,
      ) => {
        anchorQueries += 1;
        return indexedNames.getDefinedNamesAnchoredInRange(sheetId, anchorRange);
      },
    };
    const selected = range(sheet.id, 0, 1, 0, 0);
    const affectedColumnEnd = rowPermutationAffectedColumnEnd(workbook, selected, referenceOwners);

    const changes = applyRowPermutation(
      workbook,
      createRowPermutationPlan(selected, [1, 0], affectedColumnEnd),
      referenceOwners,
    );

    assert.equal(affectedColumnEnd, 12);
    assert.equal(anchorQueries, 2);
    assert.equal(workbook.getDefinedNameExact('AnchoredOwner', 'workbook')?.anchor?.row, 1);
    assert.equal(changes.definedNameOwnerDeltas[0]?.after.formula, '=A2');
  });

  it('remaps banding and cross-sheet drawing references with a row permutation', () => {
    const workbook = new WorkbookModel('permutation-drawing-references', 'Permutation drawing references');
    const sheet = workbook.getSheet('sheet-1');
    const drawingOwner = workbook.addSheet('sheet-2', 'Drawing owner', 4, 2);
    sheet.rowCount = 4;
    sheet.columnCount = 1;
    sheet.bandedRule = {
      range: range(sheet.id, 0, 0, 0, 0),
      firstColor: '#ffffff',
      secondColor: '#eeeeee',
    };
    drawingOwner.drawingPayloads.set('camera-1', {
      kind: 'camera',
      sourceRange: range(sheet.id, 0, 1, 0, 0),
      refreshPolicy: 'live',
    });

    applyPermutation(workbook, range(sheet.id, 0, 2, 0, 0), [2, 0, 1]);

    assert.deepEqual(sheet.bandedRule?.range, range(sheet.id, 1, 1, 0, 0));
    assert.deepEqual((drawingOwner.drawingPayloads.get('camera-1') as { sourceRange: RangeRef }).sourceRange,
      range(sheet.id, 1, 2, 0, 0));
  });

  it('remaps cross-sheet Sparkline and Pivot sources without moving their owner anchors', () => {
    const workbook = new WorkbookModel('permutation-cross-sheet-owners', 'Permutation cross-sheet owners');
    const source = workbook.getSheet('sheet-1');
    const owner = workbook.addSheet('sheet-2', 'Reference owner', 4, 3);
    source.rowCount = 4;
    source.columnCount = 1;
    owner.sparklines.push({
      id: 'cross-sheet-sparkline', sheetId: owner.id, anchor: { row: 0, column: 1 },
      sourceRange: range(source.id, 0, 0, 0, 0), type: 'line', color: '#000000',
    });
    owner.pivots.push(pivotDefinition('cross-sheet-pivot', owner.id, {
      kind: 'worksheet-ranges',
      ranges: [
        { sourceId: 'source', range: range(source.id, 0, 0, 0, 0) },
        { sourceId: 'owner', range: range(owner.id, 0, 0, 0, 0) },
        { sourceId: 'outside', range: range(source.id, 3, 3, 0, 0) },
      ],
      relationships: [],
    }));
    owner.pivots.push(pivotDefinition('cross-sheet-single-pivot', owner.id, {
      kind: 'worksheet-range', range: range(source.id, 1, 1, 0, 0),
    }));
    const pivot = owner.pivots[0]!;
    assert.equal(pivot.source.kind, 'worksheet-ranges');
    if (pivot.source.kind !== 'worksheet-ranges') throw new Error('Test pivot source must use explicit worksheet ranges');
    const ownerRangeIdentity = pivot.source.ranges[1]!.range;
    const outsideRangeIdentity = pivot.source.ranges[2]!.range;

    applyPermutation(workbook, range(source.id, 0, 2, 0, 0), [2, 0, 1]);

    assert.deepEqual(owner.sparklines[0]?.sourceRange, range(source.id, 1, 1, 0, 0));
    assert.equal(owner.sparklines[0]?.anchor.row, 0);
    assert.deepEqual(pivot.source.ranges[0]?.range, range(source.id, 1, 1, 0, 0));
    assert.strictEqual(pivot.source.ranges[1]?.range, ownerRangeIdentity);
    assert.deepEqual(pivot.source.ranges[1]?.range, range(owner.id, 0, 0, 0, 0));
    assert.strictEqual(pivot.source.ranges[2]?.range, outsideRangeIdentity);
    assert.equal(pivot.target.anchor.row, 0);
    assert.deepEqual(owner.pivots[1]?.source, {
      kind: 'worksheet-range', range: range(source.id, 2, 2, 0, 0),
    });
  });

  it('rejects cross-sheet single-range owners that a permutation would split before changing cells', () => {
    const workbook = new WorkbookModel('permutation-cross-sheet-reject', 'Permutation cross-sheet reject');
    const source = workbook.getSheet('sheet-1');
    const owner = workbook.addSheet('sheet-2', 'Reference owner', 4, 2);
    source.rowCount = 4;
    source.columnCount = 1;
    owner.sparklines.push({
      id: 'cross-sheet-sparkline', sheetId: owner.id, anchor: { row: 0, column: 1 },
      sourceRange: range(source.id, 0, 1, 0, 0), type: 'line', color: '#000000',
    });
    owner.pivots.push(pivotDefinition('cross-sheet-pivot', owner.id, {
      kind: 'worksheet-ranges',
      ranges: [{ sourceId: 'source', range: range(source.id, 0, 1, 0, 0) }],
      relationships: [],
    }));
    source.cells.set(0, 0, { value: 'first' });
    source.cells.set(1, 0, { value: 'second' });
    const selected = range(source.id, 0, 3, 0, 0);
    const beforeSparklineRejection = workbook.snapshot();

    assert.throws(() => applyPermutation(workbook, selected, [2, 0, 3, 1]), /cannot exactly remap sparkline cross-sheet-sparkline/);
    assert.deepEqual(workbook.snapshot(), beforeSparklineRejection);
    owner.sparklines.length = 0;
    const beforePivotRejection = workbook.snapshot();
    assert.throws(() => applyPermutation(workbook, selected, [2, 0, 3, 1]), /cannot exactly remap pivot cross-sheet-pivot source/);
    assert.deepEqual(workbook.snapshot(), beforePivotRejection);
  });

  it('moves workbook table and data-source source ranges with permuted source rows', () => {
    const workbook = new WorkbookModel('permutation-data-model-ranges', 'Permutation data model ranges');
    const sheet = workbook.getSheet('sheet-1');
    sheet.rowCount = 4;
    sheet.columnCount = 1;
    for (const [row, value] of ['first', 'second', 'third', 'fourth'].entries()) {
      sheet.cells.set(row, 0, { value });
    }
    const sourceRange = range(sheet.id, 0, 2, 0, 0);
    workbook.addTable({
      id: 'workbook-table',
      name: 'Source rows',
      sourceSheetId: sheet.id,
      sourceRange: { ...sourceRange },
      rowCount: 2,
      fields: [{ id: 'value', name: 'Value', ordinal: 0, type: 'text' }],
      blockSize: 128,
      blocks: [],
      revision: 0,
    });
    workbook.addDataSource({
      schema: 'DataSourceManifest',
      version: 1,
      id: 'data-source',
      name: 'Source rows',
      kind: 'chunked-table',
      sourceSheetId: sheet.id,
      sourceRange: { ...sourceRange },
      rowCount: 2,
      fields: [{ id: 'value', name: 'Value', ordinal: 0, type: 'text' }],
      blockRowCount: 65_536,
      blocks: [{
        id: 'block-1', dataSourceId: 'data-source', startRow: 0, rowCount: 2,
        storageKey: 'block-1', checksum: 'a'.repeat(64), byteLength: 1,
        encoding: 'columnar-v1', revision: 0,
      }],
      revision: 0,
    });

    applyPermutation(workbook, range(sheet.id, 0, 3, 0, 0), [3, 0, 1, 2]);

    assert.deepEqual(workbook.getTable('workbook-table').sourceRange, range(sheet.id, 1, 3, 0, 0));
    assert.deepEqual(workbook.getDataSource('data-source').sourceRange, range(sheet.id, 1, 3, 0, 0));
    assert.equal(sheet.cells.get(1, 0)?.value, 'first');
    assert.equal(sheet.cells.get(3, 0)?.value, 'third');
  });

  it('rejects split workbook table and data-source source ranges before changing cells', () => {
    const createWorkbook = (id: string): WorkbookModel => {
      const workbook = new WorkbookModel(id, id);
      const sheet = workbook.getSheet('sheet-1');
      sheet.rowCount = 4;
      sheet.columnCount = 1;
      sheet.cells.set(0, 0, { value: 'first' });
      sheet.cells.set(1, 0, { value: 'second' });
      sheet.cells.set(2, 0, { value: 'third' });
      const sourceRange = range(sheet.id, 0, 2, 0, 0);
      workbook.addTable({
        id: 'workbook-table', name: 'Source rows', sourceSheetId: sheet.id,
        sourceRange: { ...sourceRange }, rowCount: 2, fields: [], blockSize: 128, blocks: [], revision: 0,
      });
      workbook.addDataSource({
        schema: 'DataSourceManifest', version: 1, id: 'data-source', name: 'Source rows', kind: 'chunked-table',
        sourceSheetId: sheet.id, sourceRange: { ...sourceRange }, rowCount: 2,
        fields: [{ id: 'value', name: 'Value', ordinal: 0, type: 'text' }], blockRowCount: 65_536,
        blocks: [{ id: 'block-1', dataSourceId: 'data-source', startRow: 0, rowCount: 2, storageKey: 'block-1', checksum: 'a'.repeat(64), byteLength: 1, encoding: 'columnar-v1', revision: 0 }],
        revision: 0,
      });
      return workbook;
    };

    for (const ownerKind of ['workbook-table', 'data-source'] as const) {
      const workbook = createWorkbook(`permutation-split-${ownerKind}`);
      if (ownerKind === 'workbook-table') workbook.removeDataSource('data-source');
      else workbook.removeTable('workbook-table');
      const before = workbook.snapshot();
      const sheet = workbook.getSheet('sheet-1');

      assert.throws(
        () => applyPermutation(workbook, range(sheet.id, 0, 3, 0, 0), [2, 0, 3, 1]),
        /cannot exactly remap/,
      );
      assert.deepEqual(workbook.snapshot(), before);
    }
  });

  it('rejects a row permutation whose metadata scope overstates the canonical owner extent', () => {
    const workbook = new WorkbookModel('permutation-overstated-scope', 'Permutation overstated scope');
    const sheet = workbook.getSheet('sheet-1');
    sheet.rowCount = 4;
    sheet.columnCount = 2;
    sheet.cells.set(0, 0, { value: 'first' });
    sheet.cells.set(1, 0, { value: 'second' });
    const selected = range(sheet.id, 0, 1, 0, 0);
    const before = sheet.cells.toJSON();
    const canonicalEnd = rowPermutationAffectedColumnEnd(
      workbook,
      selected,
      buildDefinedNameAnchorIndex(workbook),
    );

    assert.throws(
      () => applyRowPermutation(
        workbook,
        createRowPermutationPlan(selected, [1, 0], canonicalEnd + 1),
        buildDefinedNameAnchorIndex(workbook),
      ),
      /does not match its canonical owners/,
    );
    assert.deepEqual(sheet.cells.toJSON(), before);
  });

  it('rebases moved formula owners, provenance formulas, and barcode formulas', () => {
    const workbook = new WorkbookModel('permutation-formulas', 'Permutation formulas');
    const sheet = workbook.getSheet('sheet-1');
    sheet.rowCount = 4;
    sheet.columnCount = 8;
    sheet.cells.set(0, 0, {
      value: null,
      formula: '=C2+$D$1',
      formulaMetadata: { kind: 'normal', sourceFormula: '=C2+$D$1' },
      presentation: {
        kind: 'barcode',
        symbology: 'qr',
        source: { kind: 'formula', formula: '=E2' },
        parameters: { symbology: 'qr' },
        options: { foreground: '#000000', background: '#ffffff', showText: false, labelPosition: 'none', quietZone: 0 },
      },
    });
    sheet.cells.set(1, 0, { value: 'second row' });

    applyPermutation(workbook, range(sheet.id, 0, 1, 0, 0), [1, 0]);

    assert.equal(sheet.cells.get(1, 0)?.formula, '=C3+$D$1');
    assert.equal(sheet.cells.get(1, 0)?.formulaMetadata?.sourceFormula, '=C3+$D$1');
    const movedPresentation = sheet.cells.get(1, 0)?.presentation;
    assert.equal(movedPresentation?.kind, 'barcode');
    if (movedPresentation?.kind === 'barcode') {
      assert.equal(movedPresentation.source.kind, 'formula');
      if (movedPresentation.source.kind === 'formula') assert.equal(movedPresentation.source.formula, '=E3');
    }
  });

  it('returns reversible cell and rule formula-owner deltas for a row permutation', () => {
    const workbook = new WorkbookModel('permutation-owner-deltas', 'Permutation owner deltas');
    const sheet = workbook.getSheet('sheet-1');
    sheet.rowCount = 3;
    sheet.columnCount = 1;
    sheet.cells.set(0, 0, { value: null, formula: '=A1' });
    sheet.conditionalFormats.push({
      id: 'cf-permuted',
      sheetId: sheet.id,
      ranges: [range(sheet.id, 0, 0, 4, 4)],
      formulaAnchor: { sheetId: sheet.id, row: 0, column: 4 },
      type: 'highlight',
      operator: 'formula',
      value1: '=A1>0',
    });

    const changes = applyPermutation(workbook, range(sheet.id, 0, 1, 0, 0), [1, 0]);

    assert.deepEqual(changes.formulaOwnerDeltas, [
      {
        kind: 'formula-cell',
        beforeAddress: { sheetId: sheet.id, row: 0, column: 0 },
        afterAddress: { sheetId: sheet.id, row: 1, column: 0 },
        before: { formula: '=A1', sourceFormula: null, barcodeFormula: null },
        after: { formula: '=A2', sourceFormula: null, barcodeFormula: null },
      },
      {
        kind: 'formula-rule',
        sheetId: sheet.id,
        ruleKind: 'conditional-format',
        ruleId: 'cf-permuted',
        field: 'value1',
        beforeFormula: '=A1>0',
        afterFormula: '=A2>0',
        beforeRanges: [range(sheet.id, 0, 0, 4, 4)],
        afterRanges: [range(sheet.id, 1, 1, 4, 4)],
      },
    ]);
  });

  it('rebases rule, defined-name, and reusable-template formulas when their anchors move outside the sorted columns', () => {
    const workbook = new WorkbookModel('permutation-workbook-owners', 'Permutation workbook owners');
    const sheet = workbook.getSheet('sheet-1');
    const otherSheet = workbook.addSheet('sheet-2', 'Other sheet', 4, 2);
    sheet.rowCount = 4;
    sheet.columnCount = 2;
    sheet.cells.set(0, 0, { value: 'first' });
    sheet.cells.set(1, 0, { value: 'second' });
    sheet.conditionalFormats.push({
      id: 'cf-anchored',
      sheetId: sheet.id,
      ranges: [range(sheet.id, 0, 0, 4, 4)],
      formulaAnchor: { sheetId: sheet.id, row: 0, column: 4 },
      type: 'highlight',
      operator: 'formula',
      value1: '=A1>0',
    });
    sheet.conditionalFormats.push({
      id: 'cf-implicit-anchor',
      sheetId: sheet.id,
      ranges: [range(sheet.id, 0, 0, 8, 8)],
      type: 'highlight',
      operator: 'formula',
      value1: '=A1>0',
    });
    sheet.conditionalFormats.push({
      id: 'cf-literal',
      sheetId: sheet.id,
      ranges: [range(sheet.id, 0, 0, 6, 6)],
      type: 'highlight',
      operator: 'greaterThan',
      value1: '0',
    });
    sheet.dataValidations.push({
      id: 'dv-anchored',
      sheetId: sheet.id,
      ranges: [range(sheet.id, 0, 0, 5, 5)],
      formulaAnchor: { sheetId: sheet.id, row: 0, column: 5 },
      type: 'custom',
      formula1: '=A1>0',
      formula2: '=B1',
      listSource: { kind: 'formula', formula: '=C1:C2' },
    });
    workbook.setDefinedName({
      name: 'RelativeOwner',
      scope: 'workbook',
      formula: '=A1',
      anchor: { sheetId: sheet.id, row: 0, column: 6 },
    });
    workbook.setDefinedName({
      name: 'OtherSheetOwner',
      scope: 'workbook',
      formula: '=A1',
      anchor: { sheetId: otherSheet.id, row: 0, column: 9 },
    });
    workbook.setCellStyleTemplate({
      id: 'template-anchored',
      name: 'Anchored validation',
      style: {},
      dataValidation: {
        type: 'custom',
        formula1: '=A1>0',
        formula2: '=B1',
        listSource: { kind: 'formula', formula: '=C1:C2' },
        formulaAnchor: { sheetId: sheet.id, row: 0, column: 7 },
      },
    });
    workbook.setCellStyleTemplate({
      id: 'template-other-sheet',
      name: 'Other sheet validation',
      style: {},
      dataValidation: {
        type: 'custom',
        formula1: '=A1>0',
        formulaAnchor: { sheetId: otherSheet.id, row: 0, column: 10 },
      },
    });

    assert.equal(
      rowPermutationAffectedColumnEnd(workbook, range(sheet.id, 0, 1, 0, 0), buildDefinedNameAnchorIndex(workbook)),
      8,
    );
    const changes = applyPermutation(workbook, range(sheet.id, 0, 1, 0, 0), [1, 0]);

    assert.equal(sheet.conditionalFormats[0]?.formulaAnchor?.row, 1);
    assert.equal(sheet.conditionalFormats[0]?.value1, '=A2>0');
    assert.equal(sheet.conditionalFormats[1]?.formulaAnchor?.row, 1);
    assert.equal(sheet.conditionalFormats[1]?.value1, '=A2>0');
    assert.equal(sheet.conditionalFormats[2]?.formulaAnchor, undefined);
    assert.equal(sheet.dataValidations[0]?.formulaAnchor?.row, 1);
    assert.equal(sheet.dataValidations[0]?.formula1, '=A2>0');
    assert.equal(sheet.dataValidations[0]?.formula2, '=B2');
    assert.equal(sheet.dataValidations[0]?.listSource?.kind, 'formula');
    if (sheet.dataValidations[0]?.listSource?.kind === 'formula') assert.equal(sheet.dataValidations[0].listSource.formula, '=C2:C3');
    assert.equal(workbook.definedNameModels[0]?.formula, '=A2');
    assert.equal(workbook.definedNameModels[0]?.anchor?.row, 1);
    assert.equal(changes.definedNameOwnerDeltas.length, 1);
    assert.deepEqual(changes.definedNameOwnerDeltas[0]?.owner, {
      scope: 'workbook',
      name: 'RelativeOwner',
    });
    assert.equal(changes.definedNameOwnerDeltas[0]?.before.formula, '=A1');
    assert.equal(changes.definedNameOwnerDeltas[0]?.before.anchor?.row, 0);
    assert.equal(changes.definedNameOwnerDeltas[0]?.after.formula, '=A2');
    assert.equal(changes.definedNameOwnerDeltas[0]?.after.anchor?.row, 1);
    assert.equal(workbook.definedNameModels[1]?.formula, '=A1');
    assert.equal(workbook.definedNameModels[1]?.anchor?.row, 0);
    const templateValidation = workbook.cellStyleTemplates.get('template-anchored')?.dataValidation;
    assert.equal(templateValidation?.formulaAnchor?.row, 1);
    assert.equal(templateValidation?.formula1, '=A2>0');
    assert.equal(templateValidation?.formula2, '=B2');
    assert.equal(templateValidation?.listSource?.kind, 'formula');
    if (templateValidation?.listSource?.kind === 'formula') assert.equal(templateValidation.listSource.formula, '=C2:C3');
    const otherTemplateValidation = workbook.cellStyleTemplates.get('template-other-sheet')?.dataValidation;
    assert.equal(otherTemplateValidation?.formulaAnchor?.row, 0);
    assert.equal(otherTemplateValidation?.formula1, '=A1>0');
  });

  it('rejects invalid anchored-name offsets before mutating cells or workbook formula owners', () => {
    const workbook = new WorkbookModel('permutation-owner-reject', 'Permutation owner rejection');
    const sheet = workbook.getSheet('sheet-1');
    sheet.rowCount = 4;
    sheet.columnCount = 2;
    sheet.cells.set(0, 0, { value: 'first' });
    sheet.cells.set(1, 0, { value: 'second' });
    workbook.setDefinedName({
      name: 'OutOfBoundsOwner',
      scope: 'workbook',
      formula: '=A1',
      anchor: { sheetId: sheet.id, row: 1, column: 8 },
    });
    const cellsBefore = sheet.cells.toJSON();

    assert.throws(() => applyPermutation(workbook, range(sheet.id, 0, 1, 0, 0), [1, 0]), /outside worksheet bounds/);
    assert.deepEqual(sheet.cells.toJSON(), cellsBefore);
    assert.equal(workbook.definedNameModels[0]?.formula, '=A1');
    assert.equal(workbook.definedNameModels[0]?.anchor?.row, 1);
  });

  it('rejects a stale defined-name anchor owner before applying a row permutation', () => {
    const workbook = new WorkbookModel('permutation-stale-name-index', 'Permutation stale name index');
    const sheet = workbook.getSheet('sheet-1');
    sheet.rowCount = 2;
    sheet.columnCount = 1;
    sheet.cells.set(0, 0, { value: 'first' });
    sheet.cells.set(1, 0, { value: 'second' });
    const staleIndex = new RangeIndex();
    staleIndex.setDefinedNameReference({ scope: 'workbook', name: 'MissingOwner' }, [], undefined, {
      sheetId: sheet.id,
      row: 0,
      column: 0,
    });
    const before = workbook.snapshot();
    const referenceOwners = buildDefinedNameAnchorIndex(workbook);
    const selected = range(sheet.id, 0, 1, 0, 0);
    const plan = createRowPermutationPlan(
      selected,
      [1, 0],
      rowPermutationAffectedColumnEnd(workbook, selected, referenceOwners),
    );

    assert.throws(
      () => applyRowPermutation(workbook, plan, staleIndex),
      /STRUCTURAL_REFERENCE_INDEX_INVARIANT: defined-name anchor owner workbook:\*:MissingOwner is missing from the workbook/,
    );
    assert.deepEqual(workbook.snapshot(), before);
  });

  it('rejects metadata extents beyond the Excel worksheet column limit', () => {
    const workbook = new WorkbookModel('permutation-column-bound', 'Permutation column bound');
    const sheet = workbook.getSheet('sheet-1');
    assert.throws(
      () => createRowPermutationPlan(range(sheet.id, 0, 1, 0, 0), [1, 0], 16_384),
      /metadata extent is outside worksheet bounds/,
    );
  });

  it('rejects fragmented row outline groups before mutating sorted cells', () => {
    const workbook = new WorkbookModel('permutation-outline-reject', 'Permutation outline rejection');
    const sheet = workbook.getSheet('sheet-1');
    sheet.rowCount = 8;
    sheet.columnCount = 4;
    sheet.cells.set(0, 0, { value: 'first' });
    sheet.cells.set(1, 0, { value: 'second' });
    sheet.outline = { groups: [{ id: 'group-1', axis: 'row', start: 0, end: 1, level: 1, collapsed: false }] };
    const cellsBefore = sheet.cells.toJSON();
    const outlineBefore = structuredClone(sheet.outline);

    assert.throws(
      () => applyPermutation(workbook, range(sheet.id, 0, 3, 0, 1), [2, 0, 3, 1]),
      /cannot exactly remap an outline group/,
    );
    assert.deepEqual(sheet.cells.toJSON(), cellsBefore);
    assert.deepEqual(sheet.outline, outlineBefore);
  });

  it('materializes an implicit rule anchor before transformed range fragments reorder it', () => {
    const workbook = new WorkbookModel('permutation-implicit-rule-anchor', 'Permutation implicit rule anchor');
    const sheet = workbook.getSheet('sheet-1');
    sheet.rowCount = 8;
    sheet.columnCount = 4;
    sheet.conditionalFormats.push({
      id: 'cf-implicit-fixed-anchor',
      sheetId: sheet.id,
      ranges: [range(sheet.id, 2, 5, 3, 3)],
      type: 'highlight',
      operator: 'formula',
      value1: '=A1>0',
    });

    applyPermutation(workbook, range(sheet.id, 0, 4, 0, 1), [3, 0, 2, 1, 4]);

    assert.equal(sheet.conditionalFormats[0]?.formulaAnchor?.row, 2);
    assert.equal(sheet.conditionalFormats[0]?.value1, '=A1>0');
  });

  it('rejects formula groups and references whose row semantics cannot be remapped before changing cells', () => {
    const groupedWorkbook = new WorkbookModel('permutation-group-reject', 'Formula group rejection');
    const groupedSheet = groupedWorkbook.getSheet('sheet-1');
    groupedSheet.rowCount = 4;
    groupedSheet.columnCount = 4;
    groupedSheet.cells.set(0, 0, { value: null, formula: '=A1', formulaMetadata: { kind: 'shared', range: 'A1:A2', sourceFormula: '=A1' } });
    groupedSheet.cells.set(1, 0, { value: 'second row' });
    const groupedBefore = groupedSheet.cells.toJSON();
    assert.throws(
      () => applyPermutation(groupedWorkbook, range(groupedSheet.id, 0, 1, 0, 0), [1, 0]),
      /UNSUPPORTED_STRUCTURAL_REFERENCE: row sort cannot remap formula-group metadata/,
    );
    assert.deepEqual(groupedSheet.cells.toJSON(), groupedBefore);

    for (const formula of ['=[Book]Sheet1!A1', '=SUM(Sheet1!1:3)']) {
      const workbook = new WorkbookModel('permutation-reference-reject', 'Formula reference rejection');
      const sheet = workbook.getSheet('sheet-1');
      sheet.rowCount = 4;
      sheet.columnCount = 4;
      sheet.cells.set(0, 0, { value: null, formula });
      sheet.cells.set(1, 0, { value: 'second row' });
      const before = sheet.cells.toJSON();
      assert.throws(
        () => applyPermutation(workbook, range(sheet.id, 0, 1, 0, 0), [1, 0]),
        /cannot safely offset an external-workbook or whole-row reference/,
      );
      assert.deepEqual(sheet.cells.toJSON(), before);
    }

    const boundsWorkbook = new WorkbookModel('permutation-bounds-reject', 'Formula bounds rejection');
    const boundsSheet = boundsWorkbook.getSheet('sheet-1');
    boundsSheet.rowCount = 4;
    boundsSheet.columnCount = 4;
    boundsSheet.cells.set(0, 0, { value: 'first row' });
    boundsSheet.cells.set(1, 0, { value: null, formula: '=A1' });
    const boundsBefore = boundsSheet.cells.toJSON();
    assert.throws(
      () => applyPermutation(boundsWorkbook, range(boundsSheet.id, 0, 1, 0, 0), [1, 0]),
      /would move a formula reference outside worksheet bounds/,
    );
    assert.deepEqual(boundsSheet.cells.toJSON(), boundsBefore);
  });

  it('moves only metadata whose exact cell is inside the sort rectangle', () => {
    const workbook = new WorkbookModel('permutation-metadata', 'Permutation metadata');
    const sheet = workbook.getSheet('sheet-1');
    sheet.rowCount = 8;
    sheet.columnCount = 30;
    sheet.cells.set(0, 0, { value: 'first' });
    sheet.cells.set(1, 0, { value: 'second' });
    sheet.review.setNote(0, 1, { id: 'inside-note', author: 'u', text: 'inside', createdAt: 'now', visible: true });
    sheet.review.setNote(0, 25, { id: 'outside-note', author: 'u', text: 'outside', createdAt: 'now', visible: true });
    sheet.hyperlinks.set('0:1', { id: 'inside-link', target: { kind: 'url', url: 'https://inside.invalid' } });
    sheet.hyperlinks.set('0:25', { id: 'outside-link', target: { kind: 'url', url: 'https://outside.invalid' } });
    sheet.review.addThread({ id: 'outside-comment', sheetId: sheet.id, row: 0, column: 25, author: 'u', text: 'outside', createdAt: 'now', replies: [] });
    sheet.drawings.push(
      { id: 'inside-drawing', sheetId: sheet.id, kind: 'shape', anchor: { kind: 'one-cell', row: 0, column: 1 }, transform: { x: 0, y: 0, width: 10, height: 10 }, zIndex: 0, payloadId: 'inside' },
      { id: 'outside-drawing', sheetId: sheet.id, kind: 'shape', anchor: { kind: 'one-cell', row: 0, column: 25 }, transform: { x: 0, y: 0, width: 10, height: 10 }, zIndex: 0, payloadId: 'outside' },
    );

    applyPermutation(workbook, range(sheet.id, 0, 1, 0, 1), [1, 0]);

    assert.equal(sheet.cells.get(0, 0)?.value, 'second');
    assert.equal(sheet.review.getNoteAt(1, 1)?.id, 'inside-note');
    assert.equal(sheet.review.getNoteAt(0, 25)?.id, 'outside-note');
    assert.equal(sheet.hyperlinks.get('1:1')?.id, 'inside-link');
    assert.equal(sheet.hyperlinks.get('0:25')?.id, 'outside-link');
    assert.equal(sheet.review.getThreadsAt(0, 25)[0]?.row, 0);
    assert.equal(sheet.review.getThreadsAt(0, 25)[0]?.column, 25);
    assert.equal(sheet.drawings.find((item) => item.id === 'inside-drawing')?.anchor.row, 1);
    assert.equal(sheet.drawings.find((item) => item.id === 'outside-drawing')?.anchor.row, 0);
  });

  it('splits a non-contiguous conditional-format target instead of widening it', () => {
    const workbook = new WorkbookModel('permutation-segments', 'Permutation segments');
    const sheet = workbook.getSheet('sheet-1');
    sheet.rowCount = 8;
    sheet.columnCount = 8;
    sheet.conditionalFormats.push({ id: 'cf-1', sheetId: sheet.id, ranges: [range(sheet.id, 0, 1, 0, 1)], type: 'highlight', style: { bold: true } });

    applyPermutation(workbook, range(sheet.id, 0, 3, 0, 1), [2, 0, 3, 1]);

    assert.deepEqual(sheet.conditionalFormats[0]?.ranges, [range(sheet.id, 1, 1, 0, 1), range(sheet.id, 3, 3, 0, 1)]);
  });

  it('rejects a single-range owner that cannot represent exact target segments atomically', () => {
    const workbook = new WorkbookModel('permutation-reject', 'Permutation reject');
    const sheet = workbook.getSheet('sheet-1');
    sheet.rowCount = 8;
    sheet.columnCount = 8;
    sheet.cells.set(0, 0, { value: 'a' });
    sheet.cells.set(1, 0, { value: 'b' });
    sheet.protectionRules.push({ id: 'protected', scope: 'range', sheetId: sheet.id, range: range(sheet.id, 0, 1, 0, 1), locked: true, allow: {} });

    assert.throws(() => applyPermutation(workbook, range(sheet.id, 0, 3, 0, 1), [2, 0, 3, 1]), /cannot exactly remap protection/);
    assert.equal(sheet.cells.get(0, 0)?.value, 'a');
    assert.equal(sheet.cells.get(1, 0)?.value, 'b');
    assert.deepEqual(sheet.protectionRules[0]?.range, range(sheet.id, 0, 1, 0, 1));
  });
});
