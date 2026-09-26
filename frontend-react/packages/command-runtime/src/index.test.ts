import test from 'node:test';
import assert from 'node:assert/strict';
import { CALCULATION_CONTEXT_EFFECTS, WorkbookModel, type StructuralDefinedNameOwnerDelta, type StructuralFormulaOwnerDelta } from '@react-sheets/core-model';
import { FormulaEngine } from '@react-sheets/formula-engine';
import { CommandRegistry, CommandRuntime, type MutationInfo } from './index';

const cellRange = (params: { row: number; column: number; sheetId?: string }) => [{
  sheetId: params.sheetId ?? 'sheet-1',
  startRow: params.row,
  endRow: params.row,
  startColumn: params.column,
  endColumn: params.column,
}];

const cellSetMetadata = {
  schema: {
    name: 'CellSet',
    validate: (value: unknown) => {
      if (!value || typeof value !== 'object') return false;
      const params = value as Record<string, unknown>;
      return Number.isInteger(params.row) && Number.isInteger(params.column) && 'value' in params;
    },
  },
  permission: { capability: 'test.cell.write' },
  affectedRanges: { resolve: cellRange },
  inversePolicy: { allowedMutationIds: ['cell.restore'], minCount: 1 },
} as const;

const cellRestoreMetadata = {
  schema: {
    name: 'CellRestore',
    validate: (value: unknown) => {
      if (!value || typeof value !== 'object') return false;
      const params = value as Record<string, unknown>;
      return Number.isInteger(params.row) && Number.isInteger(params.column);
    },
  },
  permission: { capability: 'test.cell.write' },
  affectedRanges: { resolve: cellRange },
  inversePolicy: { allowedMutationIds: ['cell.set'], minCount: 1 },
} as const;

test('CommandRuntime keeps formula-rule owners synchronized with a provided FormulaEngine index', () => {
  const workbook = new WorkbookModel('unit-rule-owner-index', 'Rule Owner Index');
  const sheet = workbook.getSheet('sheet-1');
  sheet.conditionalFormats.push({
    id: 'cf-indexed', sheetId: sheet.id,
    ranges: [{ sheetId: sheet.id, startRow: 0, endRow: 0, startColumn: 0, endColumn: 0 }],
    type: 'highlight', operator: 'formula', value1: '=A6',
  });
  const engine = new FormulaEngine({ defaultSheetId: sheet.id, sheetOrder: [{ id: sheet.id, name: sheet.name }] });
  const runtime = new CommandRuntime(workbook);
  runtime.setStructuralReferenceOwnersProvider(() => engine.dependencies);
  let rowFiveOwners: readonly unknown[] = [];
  let rowSixOwners: readonly unknown[] = [];
  runtime.registry.registerCommand({
    id: 'formula-rule.index.inspect',
    execute: (_params, context) => {
      rowFiveOwners = context.structuralReferenceOwners.getStructuralFormulaRuleDependents(sheet.id, 'row', 5);
      rowSixOwners = context.structuralReferenceOwners.getStructuralFormulaRuleDependents(sheet.id, 'row', 6);
      return { operationId: context.operationId, mutationCount: 0, affectedRanges: [] };
    },
  });

  runtime.execute('formula-rule.index.inspect', {});
  assert.equal(rowFiveOwners.length, 1);
  assert.equal(rowSixOwners.length, 0);

  sheet.conditionalFormats[0]!.value1 = '=A7';
  runtime.execute('formula-rule.index.inspect', {});
  assert.equal(rowFiveOwners.length, 0);
  assert.equal(rowSixOwners.length, 1);

  engine.dependencies.clear();
  runtime.execute('formula-rule.index.inspect', {});
  assert.equal(rowSixOwners.length, 1);
});

test('CommandRuntime executes a registered command and tracks history', () => {
  const workbook = new WorkbookModel('unit-1', 'Runtime');
  const runtime = new CommandRuntime(workbook);
  runtime.registry.registerMutation({
    id: 'cell.set',
    handler: (item, context) => {
      const params = item.params as { row: number; column: number; value: string };
      context.workbook.getSheet(item.sheetId).cells.set(params.row, params.column, { value: params.value });
      return { replayed: 'cell.set' };
    },
    metadata: cellSetMetadata,
  });
  runtime.registry.registerMutation({
    id: 'cell.restore',
    handler: (item, context) => {
      const params = item.params as { row: number; column: number; previous?: { value: string } };
      if (params.previous) context.workbook.getSheet(item.sheetId).cells.set(params.row, params.column, params.previous);
      else context.workbook.getSheet(item.sheetId).cells.delete(params.row, params.column);
      return { replayed: 'cell.restore' };
    },
    metadata: cellRestoreMetadata,
  });
  runtime.registry.registerCommand({
    id: 'cell.set',
    execute: (params: { row: number; column: number; value: string }, context) => {
      const sheet = context.workbook.getSheet('sheet-1');
      const previous = sheet.cells.get(params.row, params.column);
      const range = [{ sheetId: 'sheet-1', startRow: params.row, endRow: params.row, startColumn: params.column, endColumn: params.column }];
      context.applyMutation({
        id: 'cell.set',
        unitId: context.workbook.unitId,
        sheetId: 'sheet-1',
        params,
        affectedRanges: range,
        inverse: [{ id: 'cell.restore', unitId: context.workbook.unitId, sheetId: 'sheet-1', params: { row: params.row, column: params.column, previous }, affectedRanges: range }],
        apply: () => {
          sheet.cells.set(params.row, params.column, { value: params.value });
          return { applied: 'cell.set' };
        },
      });
      return { operationId: context.operationId, mutationCount: 1, affectedRanges: range };
    },
  });

  const listenedMutations: MutationInfo[] = [];
  const observedEffects: Array<{ source: string; effect: unknown }> = [];
  const unsubscribe = runtime.onMutation((m, source, effect) => {
    listenedMutations.push(m);
    observedEffects.push({ source, effect });
  });

  const result = runtime.execute('cell.set', { row: 1, column: 1, value: 'A' });
  assert.equal(result.mutationCount, 1);
  assert.equal(listenedMutations.length, 1);
  assert.deepEqual(observedEffects[0], { source: 'command', effect: { applied: 'cell.set' } });
  assert.equal(runtime.getHistoryDepth().undo, 1);
  assert.equal(runtime.undo(), true);
  assert.deepEqual(observedEffects[1], { source: 'undo', effect: { replayed: 'cell.restore' } });
  assert.equal(workbook.getSheet('sheet-1').cells.get(1, 1), undefined);
  assert.equal(runtime.redo(), true);
  assert.deepEqual(observedEffects[2], { source: 'redo', effect: { replayed: 'cell.set' } });
  assert.equal(workbook.getSheet('sheet-1').cells.get(1, 1)?.value, 'A');

  unsubscribe();
});

test('CommandRuntime restores chart linked formulas through local history', () => {
  const workbook = new WorkbookModel('unit-chart-formula-history', 'Chart Formula History');
  const sheet = workbook.getSheet('sheet-1');
  sheet.drawingPayloads.set('chart-1', {
    kind: 'chart',
    chartId: 'chart-1',
    chartType: 'line',
    subtype: 'line',
    source: { kind: 'worksheet-ranges', ranges: [{ sheetId: sheet.id, startRow: 0, endRow: 1, startColumn: 0, endColumn: 0 }] },
    elements: { hiddenData: 'show', titleText: { linkedFormula: '=A1' } },
  });
  const delta: StructuralFormulaOwnerDelta = {
    kind: 'formula-object',
    ownerKind: 'chart-text',
    sheetId: sheet.id,
    payloadId: 'chart-1',
    field: 'titleText.linkedFormula',
    beforeFormula: '=A1',
    afterFormula: '=A2',
  };
  const runtime = new CommandRuntime(workbook);
  const metadata = (name: string, inverseId: string) => ({
    schema: { name, validate: (value: unknown) => !!value && typeof value === 'object' },
    permission: { capability: 'test.chart.write' },
    affectedRanges: { resolve: () => [] },
    inversePolicy: { allowedMutationIds: [inverseId], minCount: 1 },
  });
  runtime.registry.registerMutation({
    id: 'chart.formula.set',
    handler: (item, context) => {
      const payload = context.workbook.getSheet(item.sheetId).drawingPayloads.get('chart-1');
      if (payload?.kind !== 'chart') throw new Error('Expected chart payload during replay');
      payload.elements.titleText!.linkedFormula = (item.params as { formula: string }).formula;
    },
    metadata: metadata('ChartFormulaSet', 'chart.formula.restore'),
  });
  runtime.registry.registerMutation({
    id: 'chart.formula.restore',
    handler: () => undefined,
    metadata: metadata('ChartFormulaRestore', 'chart.formula.set'),
  });
  runtime.registry.registerCommand({
    id: 'chart.formula.set',
    execute: (_params: unknown, context) => {
      context.applyMutation({
        id: 'chart.formula.set',
        unitId: workbook.unitId,
        sheetId: sheet.id,
        params: { formula: '=A2' },
        affectedRanges: [],
        inverse: [{
          id: 'chart.formula.restore',
          unitId: workbook.unitId,
          sheetId: sheet.id,
          params: {},
          affectedRanges: [],
        }],
        apply: () => {
          const payload = sheet.drawingPayloads.get('chart-1');
          if (payload?.kind !== 'chart') throw new Error('Expected chart payload during command');
          payload.elements.titleText!.linkedFormula = '=A2';
          return { formulaOwnerDeltas: [delta] };
        },
      });
      return { operationId: context.operationId, mutationCount: 1, affectedRanges: [] };
    },
  });

  runtime.execute('chart.formula.set', {});
  assert.equal((sheet.drawingPayloads.get('chart-1') as { elements: { titleText: { linkedFormula: string } } }).elements.titleText.linkedFormula, '=A2');
  assert.equal(runtime.undo(), true);
  assert.equal((sheet.drawingPayloads.get('chart-1') as { elements: { titleText: { linkedFormula: string } } }).elements.titleText.linkedFormula, '=A1');
  assert.equal(runtime.redo(), true);
  assert.equal((sheet.drawingPayloads.get('chart-1') as { elements: { titleText: { linkedFormula: string } } }).elements.titleText.linkedFormula, '=A2');
});

test('CommandRuntime replays formula owner history without hydrating deferred cells', () => {
  const workbook = new WorkbookModel('unit-deferred-formula-history', 'Deferred Formula History');
  const sheet = workbook.getSheet('sheet-1');
  const deferredCells = { '0': { '0': { value: null, formula: '=Sales[Amount]' } } };
  sheet.cells.deferJSON(deferredCells);
  const delta: StructuralFormulaOwnerDelta = {
    kind: 'formula-cell',
    beforeAddress: { sheetId: sheet.id, row: 0, column: 0 },
    afterAddress: { sheetId: sheet.id, row: 0, column: 0 },
    before: { formula: '=Sales[Amount]', sourceFormula: null, barcodeFormula: null },
    after: { formula: '=Orders[Amount]', sourceFormula: null, barcodeFormula: null },
  };
  const runtime = new CommandRuntime(workbook);
  const metadata = (name: string, inverseId: string) => ({
    schema: { name, validate: (value: unknown) => !!value && typeof value === 'object' },
    permission: { capability: 'test.formula.write' },
    affectedRanges: { resolve: () => [] },
    inversePolicy: { allowedMutationIds: [inverseId], minCount: 1 },
  });
  runtime.registry.registerMutation({
    id: 'formula.owner.rename',
    handler: () => undefined,
    metadata: metadata('FormulaOwnerRename', 'formula.owner.restore'),
  });
  runtime.registry.registerMutation({
    id: 'formula.owner.restore',
    handler: () => undefined,
    metadata: metadata('FormulaOwnerRestore', 'formula.owner.rename'),
  });
  runtime.registry.registerCommand({
    id: 'formula.owner.rename',
    execute: (_params, context) => {
      context.applyMutation({
        id: 'formula.owner.rename', unitId: workbook.unitId, sheetId: sheet.id, params: {}, affectedRanges: [],
        inverse: [{ id: 'formula.owner.restore', unitId: workbook.unitId, sheetId: sheet.id, params: {}, affectedRanges: [] }],
        apply: () => {
          const current = sheet.cells.getFormulaOwnerWithoutHydration(0, 0)!;
          sheet.cells.replaceFormulaOwnerWithoutHydration(0, 0, { ...current, formula: '=Orders[Amount]' });
          return { formulaOwnerDeltas: [delta] };
        },
      });
      return { operationId: context.operationId, mutationCount: 1, affectedRanges: [] };
    },
  });

  let redoEffect: unknown;
  runtime.onMutation((_mutation, source, effect) => {
    if (source === 'redo') redoEffect = effect;
  });
  runtime.execute('formula.owner.rename', {});
  assert.equal(runtime.getUndoEntries()[0]?.forwardMutations[0]?.structuralFormulaOwnerDeltas?.length, 1);
  assert.equal(sheet.cells.isHydrated, false);
  assert.equal(runtime.undo(), true);
  assert.equal(sheet.cells.toJSON()['0']?.['0']?.formula, '=Sales[Amount]');
  assert.equal(sheet.cells.isHydrated, false);
  assert.equal(runtime.redo(), true);
  assert.equal(sheet.cells.toJSON()['0']?.['0']?.formula, '=Orders[Amount]');
  assert.equal((redoEffect as { kind?: string } | undefined)?.kind, 'structural-transform');
  assert.equal(sheet.cells.isHydrated, false);
  const changedOwner = sheet.cells.getFormulaOwnerWithoutHydration(0, 0)!;
  sheet.cells.replaceFormulaOwnerWithoutHydration(0, 0, { ...changedOwner, formula: '=Broken[Amount]' });
  assert.throws(() => runtime.undo(), /STRUCTURAL_PATCH_PRECONDITION: formula owner/);
  assert.equal(sheet.cells.toJSON()['0']?.['0']?.formula, '=Broken[Amount]');
  assert.equal(sheet.cells.isHydrated, false);
  assert.equal(deferredCells['0']?.['0']?.formula, '=Sales[Amount]');
});

test('CommandRuntime records and guards defined-name owner patches in history', () => {
  const workbook = new WorkbookModel('unit-defined-name-history', 'Defined Name History');
  const sheetId = workbook.primarySheetId;
  const before = { name: 'Rate', formula: '=A1', scope: 'workbook' as const, anchor: { sheetId, row: 0, column: 0 } };
  const after = { ...before, formula: '=A2', anchor: { sheetId, row: 1, column: 0 } };
  workbook.setDefinedName(before);
  const delta: StructuralDefinedNameOwnerDelta = {
    owner: { name: 'Rate', scope: 'workbook' },
    before,
    after,
  };
  const ownerRanges = [{ sheetId, startRow: 0, endRow: 1, startColumn: 0, endColumn: 0 }];
  const runtime = new CommandRuntime(workbook);
  const replayEffects: unknown[] = [];
  runtime.onMutation((_mutation, source, effect) => {
    if (source === 'undo' || source === 'redo') replayEffects.push(effect);
  });
  const metadata = (name: string, allowedMutationIds: string[], ranges = ownerRanges) => ({
    schema: { name, validate: (value: unknown) => !!value && typeof value === 'object' },
    permission: { capability: 'test.defined-name.write' },
    affectedRanges: { resolve: () => ranges, mode: 'exact' as const },
    inversePolicy: { allowedMutationIds, minCount: 1 },
  });

  runtime.registry.registerMutation({
    id: 'defined-name.transform',
    handler: (item, context) => context.workbook.setDefinedName(item.params as StructuralDefinedNameOwnerDelta['after']),
    metadata: metadata('DefinedNameTransform', ['defined-name.restore']),
  });
  runtime.registry.registerMutation({
    id: 'defined-name.restore',
    handler: () => undefined,
    metadata: metadata('DefinedNameRestore', ['defined-name.transform']),
  });
  runtime.registry.registerMutation({
    id: 'name.set',
    handler: (item, context) => context.workbook.setDefinedName((item.params as { model: typeof before }).model),
    metadata: metadata('DefinedNameSet', ['name.remove'], []),
  });
  runtime.registry.registerMutation({
    id: 'name.remove',
    handler: (item, context) => {
      const params = item.params as { name: string; scope?: 'workbook' | 'sheet'; sheetId?: string };
      context.workbook.removeDefinedName(params.name, params.scope, params.sheetId);
    },
    metadata: metadata('DefinedNameRemove', ['name.set'], []),
  });
  runtime.registry.registerCommand({
    id: 'defined-name.transform',
    execute: (_params, context) => {
      context.applyMutation({
        id: 'defined-name.transform',
        unitId: workbook.unitId,
        sheetId,
        params: after,
        affectedRanges: ownerRanges,
        inverse: [{ id: 'defined-name.restore', unitId: workbook.unitId, sheetId, params: {}, affectedRanges: ownerRanges }],
        apply: () => {
          workbook.setDefinedName(after);
          return { definedNameOwnerDeltas: [delta] };
        },
      });
      return { operationId: context.operationId, mutationCount: 1, affectedRanges: ownerRanges };
    },
  });

  const operation = runtime.execute('defined-name.transform', {});
  assert.deepEqual(runtime.getUndoEntries()[0]?.inversePlan[0]?.structuralDefinedNameOwnerDeltas, [delta]);
  runtime.applyCommittedStructuralPatches(operation.operationId, [{
    id: 'defined-name.transform',
    unitId: workbook.unitId,
    sheetId,
    params: after,
    affectedRanges: ownerRanges,
    structuralDefinedNameOwnerDeltas: [delta],
  }], 1);
  assert.equal(runtime.getInvalidHistoryEntries().length, 0);
  workbook.setDefinedName({ ...after, formula: '=A9' });
  assert.throws(() => runtime.undo(), /STRUCTURAL_PATCH_PRECONDITION: defined-name owner/);
  assert.equal(workbook.getDefinedNameExact('Rate', 'workbook')?.formula, '=A9');
  assert.equal(runtime.getHistoryDepth().undo, 1);

  workbook.setDefinedName({ ...after, anchor: { sheetId, row: 4, column: 0 } });
  assert.throws(() => runtime.undo(), /STRUCTURAL_PATCH_PRECONDITION: defined-name owner/);
  assert.equal(workbook.getDefinedNameExact('Rate', 'workbook')?.anchor?.row, 4);

  workbook.setDefinedName(after);
  assert.equal(runtime.undo(), true);
  assert.equal(workbook.getDefinedNameExact('Rate', 'workbook')?.formula, '=A1');
  assert.deepEqual(workbook.getDefinedNameExact('Rate', 'workbook')?.anchor, before.anchor);
  assert.deepEqual((replayEffects[0] as { definedNameOwnerDeltas: StructuralDefinedNameOwnerDelta[] }).definedNameOwnerDeltas, [{
    owner: delta.owner,
    before: delta.after,
    after: delta.before,
  }]);
  assert.equal(runtime.redo(), true);
  assert.equal(workbook.getDefinedNameExact('Rate', 'workbook')?.formula, '=A2');
  assert.deepEqual((replayEffects[1] as { definedNameOwnerDeltas: StructuralDefinedNameOwnerDelta[] }).definedNameOwnerDeltas, [delta]);

  runtime.applyRemoteMutations([{
    id: 'name.set', unitId: workbook.unitId, sheetId, params: { model: { name: 'Other', formula: '=B1', scope: 'workbook' } }, affectedRanges: [],
  }]);
  assert.equal(runtime.getHistoryDepth().undo, 1);
  runtime.applyRemoteMutations([{
    id: 'name.set', unitId: workbook.unitId, sheetId, params: { model: { name: 'Rate', formula: '=A3', scope: 'workbook' } }, affectedRanges: [],
  }]);
  assert.equal(runtime.getHistoryDepth().undo, 0);
  assert.match(runtime.getInvalidHistoryEntries()[0]?.invalidReason ?? '', /defined-name owner patch/);
  assert.equal(workbook.getDefinedNameExact('Rate', 'workbook')?.formula, '=A3');
});

test('CommandRuntime skips full workbook snapshot for empty committed structural owner deltas', () => {
  const workbook = new WorkbookModel('unit-empty-structural-patch', 'Empty structural patch');
  const runtime = new CommandRuntime(workbook);
  const originalSnapshot = workbook.snapshot.bind(workbook);
  workbook.snapshot = () => { throw new Error('empty patch must not snapshot the workbook'); };

  runtime.applyCommittedStructuralPatches('empty-patch', [{
    id: 'rows.inserted',
    unitId: workbook.unitId,
    sheetId: workbook.primarySheetId,
    params: { sheetId: workbook.primarySheetId, at: 0, count: 1 },
    affectedRanges: [],
    structuralFormulaOwnerDeltas: [],
    structuralDefinedNameOwnerDeltas: [],
  }], 1);

  workbook.snapshot = originalSnapshot;
});

test('CommandRuntime preflights committed owner patches without cloning the workbook', () => {
  const workbook = new WorkbookModel('unit-sparse-structural-patch', 'Sparse structural patch');
  const sheet = workbook.getSheet(workbook.primarySheetId);
  sheet.cells.set(0, 0, { value: null, formula: '=A1' });
  sheet.cells.set(1, 0, { value: null, formula: '=A2' });
  const runtime = new CommandRuntime(workbook);
  const delta = (row: number, before: string, after: string): StructuralFormulaOwnerDelta => ({
    kind: 'formula-cell',
    beforeAddress: { sheetId: sheet.id, row, column: 0 },
    afterAddress: { sheetId: sheet.id, row, column: 0 },
    before: { formula: before, sourceFormula: null, barcodeFormula: null },
    after: { formula: after, sourceFormula: null, barcodeFormula: null },
  });
  const originalSnapshot = workbook.snapshot.bind(workbook);
  workbook.snapshot = () => { throw new Error('owner patch preflight must not snapshot the workbook'); };

  runtime.applyCommittedStructuralPatches('sparse-patch', [{
    id: 'rows.inserted',
    unitId: workbook.unitId,
    sheetId: sheet.id,
    params: { sheetId: sheet.id, at: 0, count: 1 },
    affectedRanges: [],
    structuralFormulaOwnerDeltas: [delta(0, '=A1', '=B1'), delta(1, '=A2', '=B2')],
  }], 1);
  workbook.snapshot = originalSnapshot;
  assert.equal(sheet.cells.get(0, 0)?.formula, '=B1');
  assert.equal(sheet.cells.get(1, 0)?.formula, '=B2');

  const rejectedWorkbook = new WorkbookModel('unit-sparse-structural-patch-rejected', 'Rejected sparse patch');
  const rejectedSheet = rejectedWorkbook.getSheet(rejectedWorkbook.primarySheetId);
  rejectedSheet.cells.set(0, 0, { value: null, formula: '=A1' });
  rejectedSheet.cells.set(1, 0, { value: null, formula: '=Wrong' });
  const rejectedRuntime = new CommandRuntime(rejectedWorkbook);
  const before = rejectedWorkbook.snapshot();
  const rejectedDelta = (row: number, expected: string, target: string): StructuralFormulaOwnerDelta => ({
    kind: 'formula-cell',
    beforeAddress: { sheetId: rejectedSheet.id, row, column: 0 },
    afterAddress: { sheetId: rejectedSheet.id, row, column: 0 },
    before: { formula: expected, sourceFormula: null, barcodeFormula: null },
    after: { formula: target, sourceFormula: null, barcodeFormula: null },
  });
  assert.throws(() => rejectedRuntime.applyCommittedStructuralPatches('rejected-patch', [{
    id: 'rows.inserted',
    unitId: rejectedWorkbook.unitId,
    sheetId: rejectedSheet.id,
    params: { sheetId: rejectedSheet.id, at: 0, count: 1 },
    affectedRanges: [],
    structuralFormulaOwnerDeltas: [rejectedDelta(0, '=A1', '=B1'), rejectedDelta(1, '=A2', '=B2')],
  }], 1), /STRUCTURAL_PATCH_PRECONDITION/);
  assert.deepEqual(rejectedWorkbook.snapshot(), before);

  const normalizationWorkbook = new WorkbookModel('unit-sparse-structural-normalization', 'Invalid formula cell style');
  const normalizationSheet = normalizationWorkbook.getSheet(normalizationWorkbook.primarySheetId);
  normalizationSheet.cells.set(0, 0, { value: null, formula: '=A1' });
  normalizationSheet.cells.set(1, 0, { value: null, formula: '=A2' });
  normalizationSheet.cells.getWithoutHydration(1, 0)!.style = { fontFamily: '' };
  const normalizationRuntime = new CommandRuntime(normalizationWorkbook);
  const normalizationBefore = normalizationWorkbook.snapshot();
  const normalizationDelta = (row: number, formula: string): StructuralFormulaOwnerDelta => ({
    kind: 'formula-cell',
    beforeAddress: { sheetId: normalizationSheet.id, row, column: 0 },
    afterAddress: { sheetId: normalizationSheet.id, row, column: 0 },
    before: { formula, sourceFormula: null, barcodeFormula: null },
    after: { formula: `${formula}*2`, sourceFormula: null, barcodeFormula: null },
  });
  assert.throws(() => normalizationRuntime.applyCommittedStructuralPatches('normalization-rejected-patch', [{
    id: 'rows.inserted',
    unitId: normalizationWorkbook.unitId,
    sheetId: normalizationSheet.id,
    params: { sheetId: normalizationSheet.id, at: 0, count: 1 },
    affectedRanges: [],
    structuralFormulaOwnerDeltas: [normalizationDelta(0, '=A1'), normalizationDelta(1, '=A2')],
  }], 1), /Font family must not be empty/);
  assert.deepEqual(normalizationWorkbook.snapshot(), normalizationBefore);
});

test('CommandRuntime emits declared calculation-context effects for command, undo, and redo', () => {
  const workbook = new WorkbookModel('unit-calculation-context', 'Before');
  const runtime = new CommandRuntime(workbook);
  runtime.registry.registerMutation({
    id: 'context.set',
    handler: (item, context) => { context.workbook.name = (item.params as { name: string }).name; },
    metadata: {
      schema: { name: 'ContextSet', validate: (value: unknown) => !!value && typeof value === 'object' && typeof (value as { name?: unknown }).name === 'string' },
      permission: { capability: 'test.context.write' },
      affectedRanges: { resolve: () => [] },
      calculationContextEffect: CALCULATION_CONTEXT_EFFECTS.rebuild,
      inverseIds: ['context.restore'],
    },
  });
  runtime.registry.registerMutation({
    id: 'context.restore',
    handler: (item, context) => { context.workbook.name = (item.params as { name: string }).name; },
    metadata: {
      schema: { name: 'ContextRestore', validate: (value: unknown) => !!value && typeof value === 'object' && typeof (value as { name?: unknown }).name === 'string' },
      permission: { capability: 'test.context.write' },
      affectedRanges: { resolve: () => [] },
      calculationContextEffect: CALCULATION_CONTEXT_EFFECTS.rebuild,
      inverseIds: ['context.set'],
    },
  });
  runtime.registry.registerCommand({
    id: 'context.set',
    execute: (_params: unknown, context) => {
      context.applyMutation({
        id: 'context.set',
        unitId: workbook.unitId,
        sheetId: workbook.primarySheetId,
        params: { name: 'After' },
        affectedRanges: [],
        inverse: [{ id: 'context.restore', unitId: workbook.unitId, sheetId: workbook.primarySheetId, params: { name: 'Before' }, affectedRanges: [] }],
        apply: () => { workbook.name = 'After'; },
      });
      return { operationId: context.operationId, mutationCount: 1, affectedRanges: [] };
    },
  });

  const effects: unknown[] = [];
  runtime.onMutation((_mutation, _source, effect) => effects.push(effect));
  runtime.execute('context.set', {});
  assert.equal(workbook.name, 'After');
  assert.equal(runtime.undo(), true);
  assert.equal(workbook.name, 'Before');
  assert.equal(runtime.redo(), true);
  assert.equal(workbook.name, 'After');
  assert.deepEqual(effects, Array(3).fill(CALCULATION_CONTEXT_EFFECTS.rebuild));
});

test('CommandRegistry rejects malformed calculation-context metadata', () => {
  const runtime = new CommandRuntime(new WorkbookModel('unit-invalid-context-effect', 'Invalid effect'));
  assert.throws(() => runtime.registry.registerMutation({
    id: 'context.invalid',
    handler: () => undefined,
    metadata: {
      schema: { name: 'InvalidContext', validate: () => true },
      permission: { capability: 'test.context.write' },
      affectedRanges: { resolve: () => [] },
      calculationContextEffect: { kind: 'calculation-context', action: 'rebuild-all' } as never,
      inverseIds: ['context.invalid'],
    },
  }), /invalid calculation context effect/);
});

test('CommandRuntime rolls back applied mutations if a command throws mid-execution', () => {
  const workbook = new WorkbookModel('unit-rollback', 'Rollback');
  const runtime = new CommandRuntime(workbook);

  runtime.registry.registerMutation({
    id: 'val.set',
    handler: (item, context) => {
      const params = item.params as { row: number; value: number };
      context.workbook.getSheet(item.sheetId).cells.set(params.row, 0, { value: params.value });
    },
    metadata: {
      schema: { name: 'ValueSet', validate: (value: unknown) => !!value && typeof value === 'object' && Number.isInteger((value as { row?: unknown }).row) },
      permission: { capability: 'test.value.write' },
      affectedRanges: { resolve: () => [] },
      inversePolicy: { allowedMutationIds: ['val.restore'], minCount: 1 },
    },
  });
  runtime.registry.registerMutation({
    id: 'val.restore',
    handler: (item, context) => {
      const params = item.params as { row: number };
      context.workbook.getSheet(item.sheetId).cells.delete(params.row, 0);
    },
    metadata: {
      schema: { name: 'ValueRestore', validate: (value: unknown) => !!value && typeof value === 'object' && Number.isInteger((value as { row?: unknown }).row) },
      permission: { capability: 'test.value.write' },
      affectedRanges: { resolve: () => [] },
      inversePolicy: { allowedMutationIds: ['val.set'], minCount: 1 },
    },
  });

  runtime.registry.registerCommand({
    id: 'failing.transaction',
    execute: (_params: unknown, context) => {
      const sheet = context.workbook.getSheet('sheet-1');
      context.applyMutation({
        id: 'val.set',
        unitId: context.workbook.unitId,
        sheetId: 'sheet-1',
        params: { row: 0, value: 100 },
        affectedRanges: [],
        inverse: [{ id: 'val.restore', unitId: context.workbook.unitId, sheetId: 'sheet-1', params: { row: 0 }, affectedRanges: [] }],
        apply: () => sheet.cells.set(0, 0, { value: 100 }),
      });

      // Now throw an error intentionally
      throw new Error('Simulated failure during multi-mutation command');
    },
  });

  const aborted: Array<{ commandId: string; operationId: string }> = [];
  runtime.onCommandAbort((commandId, _params, operationId) => aborted.push({ commandId, operationId }));
  assert.throws(() => runtime.execute('failing.transaction', {}), /Simulated failure/);
  // The first mutation should have been rolled back
  assert.equal(workbook.getSheet('sheet-1').cells.get(0, 0), undefined);
  assert.equal(runtime.getHistoryDepth().undo, 0);
  assert.equal(aborted.length, 1);
  assert.equal(aborted[0]?.commandId, 'failing.transaction');
  assert.ok(aborted[0]?.operationId);
});

test('CommandRegistry guards against duplicate IDs and unknown lookups', () => {
  const workbook = new WorkbookModel('unit-guard', 'Guards');
  const runtime = new CommandRuntime(workbook);

  runtime.registry.registerCommand({ id: 'cmd.1', execute: () => ({ operationId: '1', mutationCount: 0, affectedRanges: [] }) });
  assert.throws(() => runtime.registry.registerCommand({ id: 'cmd.1', execute: () => ({ operationId: '1', mutationCount: 0, affectedRanges: [] }) }), /Duplicate command/);
  assert.throws(() => runtime.execute('non.existent', {}), /Unknown command/);
  assert.deepEqual(runtime.getHistoryDepth(), { undo: 0, redo: 0 });
});

test('remote mutations reject a different workbook unit', () => {
  const workbook = new WorkbookModel('unit-remote', 'Remote');
  const runtime = new CommandRuntime(workbook);
  runtime.registry.registerMutation({
    id: 'cell.set',
    handler: (item, context) => {
      const params = item.params as { row: number; column: number; value: string };
      context.workbook.getSheet(item.sheetId).cells.set(params.row, params.column, { value: params.value });
    },
    metadata: cellSetMetadata,
  });
  runtime.registry.registerMutation({
    id: 'cell.restore',
    handler: () => undefined,
    metadata: cellRestoreMetadata,
  });

  assert.throws(() => runtime.applyRemoteMutations([{
    id: 'cell.set',
    unitId: 'other-unit',
    sheetId: 'sheet-1',
    params: { row: 0, column: 0, value: 'invalid' },
    affectedRanges: [],
  }]), /Mutation unit mismatch/);
});

test('remote revision validation rejects before applying a mutation', () => {
  const workbook = new WorkbookModel('unit-invalid-revision', 'Invalid revision');
  const runtime = new CommandRuntime(workbook);
  runtime.registry.registerMutation({
    id: 'cell.set',
    handler: (item, context) => {
      const params = item.params as { row: number; column: number; value: string };
      context.workbook.getSheet(item.sheetId).cells.set(params.row, params.column, { value: params.value });
    },
    metadata: cellSetMetadata,
  });
  runtime.registry.registerMutation({ id: 'cell.restore', handler: () => undefined, metadata: cellRestoreMetadata });

  assert.throws(() => runtime.applyRemoteMutations([{
    id: 'cell.set',
    unitId: workbook.unitId,
    sheetId: 'sheet-1',
    params: { row: 3, column: 4, value: 'must-not-apply' },
    affectedRanges: cellRange({ row: 3, column: 4 }),
  }], { revision: 0 }), /Remote revision is invalid/);
  assert.equal(workbook.getSheet('sheet-1').cells.get(3, 4), undefined);
});

test('remote structural history resolves formula sheet names before colliding IDs', () => {
  const workbook = new WorkbookModel('unit-sheet-name-rebase', 'Sheet name rebase');
  workbook.addSheet('End', 'Target');
  workbook.addSheet('end-id', 'End');
  const runtime = new CommandRuntime(workbook);
  const rowMutationMetadata = {
    schema: {
      name: 'RowShift',
      validate: (value: unknown) => !!value && typeof value === 'object'
        && Number.isInteger((value as { at?: unknown }).at)
        && Number.isInteger((value as { count?: unknown }).count),
    },
    permission: { capability: 'test.row.write' },
    affectedRanges: { resolve: () => [] },
    historyRebase: { kind: 'axis', axis: 'row', direction: 1 },
    inversePolicy: { allowedMutationIds: ['rows.deleted'], minCount: 1 },
  } as const;
  runtime.registry.registerMutation({
    id: 'cell.set',
    handler: (item, context) => {
      const params = item.params as { row: number; column: number; value: string };
      context.workbook.getSheet(item.sheetId).cells.set(params.row, params.column, { value: params.value });
    },
    metadata: cellSetMetadata,
  });
  runtime.registry.registerMutation({ id: 'cell.restore', handler: () => undefined, metadata: cellRestoreMetadata });
  runtime.registry.registerMutation({ id: 'rows.inserted', handler: () => undefined, metadata: rowMutationMetadata });
  runtime.registry.registerMutation({
    id: 'rows.deleted',
    handler: () => undefined,
    metadata: {
      ...rowMutationMetadata,
      historyRebase: { kind: 'axis', axis: 'row', direction: -1 },
      inversePolicy: { allowedMutationIds: ['rows.inserted'], minCount: 1 },
    },
  });
  runtime.registry.registerCommand({
    id: 'cell.set',
    execute: (params: { row: number; column: number; value: string }, context) => {
      const sheet = context.workbook.getSheet('sheet-1');
      const previous = sheet.cells.get(params.row, params.column);
      const affectedRanges = cellRange(params);
      context.applyMutation({
        id: 'cell.set',
        unitId: context.workbook.unitId,
        sheetId: 'sheet-1',
        params,
        affectedRanges,
        inverse: [{
          id: 'cell.restore',
          unitId: context.workbook.unitId,
          sheetId: 'sheet-1',
          params: { row: params.row, column: params.column, previous },
          affectedRanges,
        }],
        apply: () => {
          sheet.cells.set(params.row, params.column, { value: params.value });
        },
      });
      return { operationId: context.operationId, mutationCount: 1, affectedRanges };
    },
  });

  const localParams = { row: 0, column: 0, value: 'local', formula: '=End!A1' };
  runtime.execute('cell.set', localParams);
  runtime.applyRemoteMutations([{
    id: 'rows.inserted',
    unitId: workbook.unitId,
    sheetId: 'End',
    params: { at: 0, count: 1 },
    affectedRanges: [],
  }]);

  const forward = runtime.getUndoEntries()[0]?.forwardMutations[0]?.params as { formula?: string } | undefined;
  assert.equal(forward?.formula, '=End!A1');
});

test('CommandRuntime rejects an unregistered mutation before touching the workbook', () => {
  const workbook = new WorkbookModel('unit-unregistered', 'Unregistered');
  const runtime = new CommandRuntime(workbook);
  let applyCalled = false;
  runtime.registry.registerCommand({
    id: 'invalid.mutation',
    execute: (_params: unknown, context) => {
      const affectedRanges = [{ sheetId: 'sheet-1', startRow: 0, endRow: 0, startColumn: 0, endColumn: 0 }];
      context.applyMutation({
        id: 'mutation.not.registered',
        unitId: workbook.unitId,
        sheetId: 'sheet-1',
        params: {},
        affectedRanges,
        inverse: [],
        apply: () => {
          applyCalled = true;
        },
      });
      return { operationId: context.operationId, mutationCount: 1, affectedRanges };
    },
  });

  assert.throws(() => runtime.execute('invalid.mutation', {}), /Unknown mutation: mutation\.not\.registered/);
  assert.equal(applyCalled, false);
  assert.deepEqual(runtime.getHistoryDepth(), { undo: 0, redo: 0 });
});

test('CommandRuntime rejects an inverse that is not registered before applying the mutation', () => {
  const workbook = new WorkbookModel('unit-invalid-inverse', 'Invalid inverse');
  const runtime = new CommandRuntime(workbook);
  let applyCalled = false;
  runtime.registry.registerMutation({
    id: 'primary.set',
    handler: () => undefined,
    metadata: {
      schema: { name: 'PrimarySet', validate: (value: unknown) => !!value && typeof value === 'object' },
      permission: { capability: 'test.write' },
      affectedRanges: { resolve: () => [{ sheetId: 'sheet-1', startRow: 1, endRow: 1, startColumn: 1, endColumn: 1 }] },
      inversePolicy: { allowedMutationIds: ['known.inverse'], minCount: 1 },
    },
  });
  runtime.registry.registerMutation({
    id: 'known.inverse',
    handler: () => undefined,
    metadata: {
      schema: { name: 'KnownInverse', validate: (value: unknown) => !!value && typeof value === 'object' },
      permission: { capability: 'test.write' },
      affectedRanges: { resolve: () => [{ sheetId: 'sheet-1', startRow: 1, endRow: 1, startColumn: 1, endColumn: 1 }] },
      inversePolicy: { allowedMutationIds: ['primary.set'], minCount: 1 },
    },
  });
  runtime.registry.registerCommand({
    id: 'invalid.inverse',
    execute: (_params: unknown, context) => {
      const affectedRanges = [{ sheetId: 'sheet-1', startRow: 1, endRow: 1, startColumn: 1, endColumn: 1 }];
      context.applyMutation({
        id: 'primary.set',
        unitId: workbook.unitId,
        sheetId: 'sheet-1',
        params: {},
        affectedRanges,
        inverse: [{
          id: 'inverse.not.registered',
          unitId: workbook.unitId,
          sheetId: 'sheet-1',
          params: {},
          affectedRanges,
        }],
        apply: () => {
          applyCalled = true;
        },
      });
      return { operationId: context.operationId, mutationCount: 1, affectedRanges };
    },
  });

  assert.throws(() => runtime.execute('invalid.inverse', {}), /unknown inverse inverse\.not\.registered/);
  assert.equal(applyCalled, false);
  assert.deepEqual(runtime.getHistoryDepth(), { undo: 0, redo: 0 });
});

test('CommandRegistry validates schema, permission, affected ranges, and declared inverses', () => {
  const registry = new CommandRegistry();
  const range = { sheetId: 'sheet-1', startRow: 0, endRow: 0, startColumn: 0, endColumn: 0 };
  const metadata = {
    schema: { name: 'EmptyParams', validate: (value: unknown) => value !== null && typeof value === 'object' },
    permission: { capability: 'sheet.write' },
    affectedRanges: { resolve: () => [range] },
    inverseIds: ['cell.restore'],
  } as const;
  registry.registerMutation({
    id: 'cell.set',
    handler: () => undefined,
    metadata,
  });
  registry.registerMutation({
    id: 'cell.restore',
    handler: () => undefined,
    metadata: {
      schema: { name: 'EmptyParams', validate: (value: unknown) => value !== null && typeof value === 'object' },
      permission: { capability: 'sheet.write' },
      affectedRanges: { resolve: () => [range] },
      inverseIds: ['cell.set'],
    },
  });

  const result = registry.validateCompleteness();
  assert.equal(result.ok, true);
  assert.deepEqual(result.issues, []);

  const invalid = registry.validateMutation({
    id: 'cell.set',
    unitId: 'unit-1',
    sheetId: 'sheet-1',
    params: {},
    affectedRanges: [],
    inverse: [{ id: 'cell.restore', unitId: 'unit-1', sheetId: 'sheet-1', params: {}, affectedRanges: [range] }],
    apply: () => undefined,
  });
  assert.equal(invalid.some((entry) => entry.code === 'invalid-affected-ranges'), true);
});

test('CommandRegistry rejects incomplete metadata and declared inverse drift', () => {
  const registry = new CommandRegistry();
  assert.throws(() => registry.registerMutation({
    id: 'missing.contract',
    handler: () => undefined,
    metadata: undefined as never,
  }), /requires canonical metadata/);
  assert.throws(() => registry.registerMutation({
    id: 'broken.registration',
    handler: () => undefined,
    metadata: {} as never,
  }), /must declare a parameter schema/);

  registry.registerMutation({
    id: 'declared.drift',
    handler: () => undefined,
    metadata: {
      schema: { name: 'DeclaredDrift', validate: () => true },
      permission: { capability: 'test.write' },
      affectedRanges: { resolve: () => [] },
      inverseIds: ['missing.inverse'],
    },
  });
  const result = registry.validateCompleteness();
  assert.equal(result.ok, false);
  assert.equal(result.issues.some((entry) => entry.code === 'unknown-inverse' && entry.inverseId === 'missing.inverse'), true);
  assert.throws(() => registry.assertComplete(), /Mutation registry is incomplete/);
});
