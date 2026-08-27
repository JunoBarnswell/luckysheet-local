import test from 'node:test';
import assert from 'node:assert/strict';
import {
  decodeMessage,
  decodeOperationMessage,
  encodeMessage,
  encodeOperationMessage,
  AuthenticationRequiredError,
  WorkbookApiClient,
  validateHistoryRestoreRequest,
  validateOperationEnvelope,
  validateUserPreferences,
  validatePivotDefinition,
  validateWorkbookSnapshot,
} from './index';
import { PIVOT_MAX_MEMBER_COUNT, PIVOT_MEMBER_DISPLAY_LIMIT } from '@react-sheets/core-model';

test('WebSocket presence messages round-trip without becoming a mutation transport', () => {
  const message = { type: 'cursor.updated' as const, unitId: 'unit-1', state: { row: 2, column: 4, sheetId: 'sheet-1' } };
  assert.deepEqual(decodeMessage(encodeMessage(message)), message);
});

test('OperationEnvelope excludes client actor and affected ranges', () => {
  const envelope = {
    schema: 'OperationEnvelope' as const,
    operationId: 'op-1',
    unitId: 'unit-1',
    clientSequence: 1,
    baseRevision: 0,
    mutations: [{ id: 'cell.write', sheetId: 'sheet-1', params: { row: 0, column: 0, value: 'ok' } }],
    createdAt: new Date().toISOString(),
  };
  assert.deepEqual(validateOperationEnvelope(envelope), envelope);
  assert.throws(() => validateOperationEnvelope({ ...envelope, actorId: 'spoofed' }), /server-owned/);
  assert.throws(() => validateOperationEnvelope({
    ...envelope,
    mutations: [{ ...envelope.mutations[0], affectedRanges: [] }],
  }), /server-owned/);
});

test('collaboration messages reject actor-bearing presence and legacy changesets', () => {
  const message = {
    type: 'presence.updated' as const,
    unitId: 'unit-1',
    state: { row: 1, column: 1 },
  };
  assert.deepEqual(decodeOperationMessage(encodeOperationMessage(message)), message);
  assert.throws(() => decodeOperationMessage(JSON.stringify({ ...message, actorId: 'spoofed' })), /server-owned/);
  assert.throws(() => decodeOperationMessage(JSON.stringify({
    type: 'changeset.submit',
    payload: {
      schema: 'CollaborationChangeSet',
      operationId: 'old',
      unitId: 'unit-1',
      actorId: 'spoofed',
      clientSequence: 1,
      baseRevision: 0,
      mutations: [],
      createdAt: new Date().toISOString(),
    },
  })), /Unsupported collaboration message/);
});

test('WorkbookApiClient injects bearer authentication and fails closed without a provider', async () => {
  let request: RequestInit | undefined;
  const api = new WorkbookApiClient({
    authTokenProvider: () => 'token-123',
    fetchImpl: async (_input, init) => {
      request = init;
      return new Response(JSON.stringify({
        snapshot: {
          schema: 'WorkbookSnapshot',
          version: 8,
          unitId: 'unit-1',
          name: 'Workbook',
          dimensionMetrics: { normalFontFamily: 'Calibri', normalFontSizePx: 14.6666666667, maximumDigitWidthPx: 7 },
          calculationSettings: { mode: 'automatic', iterativeCalculation: false, maximumIterations: 100, maximumChange: 0.001, precisionAsDisplayed: false, calculateBeforeSave: true, fullCalculationOnLoad: false },
          dataModel: { sources: [], tables: [], relationships: [], views: [] },
          sheets: [{
            kind: 'worksheet', id: 'sheet-1',
            name: 'Sheet1',
            rowCount: 100,
            columnCount: 26,
            cells: {},
            merges: [],
            pane: { kind: 'none' },
            defaultRowHeightPx: 20,
            defaultColumnWidthPx: 64,
            pivots: [],
            sparklines: [],
            drawings: [],
            drawingPayloads: {},
            review: { notesByCell: {}, notesById: {}, threadIdsByCell: {}, threadsById: {} },
          }],
        },
        revision: 0,
      }), {
        status: 200,
        headers: { 'content-type': 'application/json' },
      });
    },
  });
  await api.getSnapshot('unit-1');
  assert.equal(new Headers(request?.headers).get('authorization'), 'Bearer token-123');
  await assert.rejects(() => new WorkbookApiClient().getSnapshot('unit-1'), AuthenticationRequiredError);
});

test('WorkbookApiClient uses a server-issued guest share token when no bearer exists', async () => {
  let request: RequestInit | undefined;
  const api = new WorkbookApiClient({
    shareTokenProvider: () => 'guest-token',
    fetchImpl: async (_input, init) => {
      request = init;
      return new Response(JSON.stringify({
        snapshot: {
          schema: 'WorkbookSnapshot',
          version: 8,
          unitId: 'unit-guest',
          name: 'Guest workbook',
          dimensionMetrics: { normalFontFamily: 'Calibri', normalFontSizePx: 14.6666666667, maximumDigitWidthPx: 7 },
          calculationSettings: { mode: 'automatic', iterativeCalculation: false, maximumIterations: 100, maximumChange: 0.001, precisionAsDisplayed: false, calculateBeforeSave: true, fullCalculationOnLoad: false },
          dataModel: { sources: [], tables: [], relationships: [], views: [] },
          sheets: [{
            kind: 'worksheet', id: 'sheet-1', name: 'Sheet1', rowCount: 10, columnCount: 10,
            cells: {}, merges: [], pane: { kind: 'none' }, defaultRowHeightPx: 20, defaultColumnWidthPx: 64,
            pivots: [], sparklines: [], drawings: [], drawingPayloads: {},
            review: { notesByCell: {}, notesById: {}, threadIdsByCell: {}, threadsById: {} },
          }],
        },
        revision: 0,
      }), { status: 200, headers: { 'content-type': 'application/json' } });
    },
  });
  await api.getSnapshot('unit-guest');
  const headers = new Headers(request?.headers);
  assert.equal(headers.get('x-workbook-share-token'), 'guest-token');
  assert.equal(headers.has('authorization'), false);
});

test('WorkbookApiClient accepts access roles only from the server projection', async () => {
  const api = new WorkbookApiClient({
    authTokenProvider: () => 'server-token',
    fetchImpl: async () => new Response(JSON.stringify({ unitId: 'unit-access', role: 'editor' }), {
      status: 200,
      headers: { 'content-type': 'application/json' },
    }),
  });
  assert.deepEqual(await api.getAccess('unit-access'), { unitId: 'unit-access', role: 'editor' });

  const malformed = new WorkbookApiClient({
    authTokenProvider: () => 'server-token',
    fetchImpl: async () => new Response(JSON.stringify({ unitId: 'unit-access', role: 'ownerish' }), {
      status: 200,
      headers: { 'content-type': 'application/json' },
    }),
  });
  await assert.rejects(() => malformed.getAccess('unit-access'), /invalid role/);
});

test('WorkbookApiClient validates cursor pages and forwards cursor, limit, and abort signal', async () => {
  let requestedUrl = '';
  let requestedSignal: AbortSignal | undefined;
  const api = new WorkbookApiClient({
    authTokenProvider: () => 'server-token',
    fetchImpl: async (input, init) => {
      requestedUrl = String(input);
      requestedSignal = init?.signal ?? undefined;
      return new Response(JSON.stringify({
        items: [{ unitId: 'unit-page', name: 'Paged', revision: 1, updatedAt: '2026-08-24T00:00:00Z', role: 'owner' }],
        nextCursor: 'next-1',
      }), { status: 200, headers: { 'content-type': 'application/json' } });
    },
  });
  const controller = new AbortController();
  const page = await api.listWorkbookPage({ view: 'all', cursor: 'cursor-1', limit: 25 }, { signal: controller.signal });
  assert.equal(page.nextCursor, 'next-1');
  assert.equal(page.items[0]?.unitId, 'unit-page');
  assert.match(requestedUrl, /cursor=cursor-1/);
  assert.match(requestedUrl, /limit=25/);
  assert.equal(requestedSignal, controller.signal);
  await assert.rejects(() => api.listWorkbookPage({ limit: 51 }), /between 1 and 50/);
});

test('User preferences response validation rejects lossy or malformed preference values', () => {
  assert.deepEqual(validateUserPreferences({
    defaultSpaceId: 'space-1', defaultFolderId: 'folder-1', autoSave: true, autoSync: false,
    offlineCache: true, importCompatibility: 'C', language: 'zh-CN', theme: 'dark', updatedAt: '2026-08-24T00:00:00Z',
  }), {
    defaultSpaceId: 'space-1', defaultFolderId: 'folder-1', autoSave: true, autoSync: false,
    offlineCache: true, importCompatibility: 'C', language: 'zh-CN', theme: 'dark', updatedAt: '2026-08-24T00:00:00Z',
  });
  assert.throws(() => validateUserPreferences({ autoSave: true, autoSync: true, offlineCache: true, importCompatibility: 'strict', theme: 'system' }), /importCompatibility/);
});

test('history restore request is target-revision-only and client API posts no snapshot', async () => {
  assert.deepEqual(validateHistoryRestoreRequest({ targetRevision: 3, reason: 'rollback' }), {
    targetRevision: 3,
    reason: 'rollback',
  });
  assert.throws(() => validateHistoryRestoreRequest({ targetRevision: 3, snapshot: {} }), /unsupported fields/);
  assert.throws(() => validateHistoryRestoreRequest({ targetRevision: 3, actorId: 'spoofed' }), /unsupported fields/);
  assert.throws(() => validateHistoryRestoreRequest({ targetRevision: -1 }), /non-negative/);

  let postedBody = '';
  const api = new WorkbookApiClient({
    authTokenProvider: () => 'token-restore',
    fetchImpl: async (_input, init) => {
      postedBody = String(init?.body ?? '');
      return new Response(JSON.stringify({}), {
        status: 200,
        headers: { 'content-type': 'application/json' },
      });
    },
  });
  await api.restoreToRevision('unit-1', 3, 'rollback');
  assert.deepEqual(JSON.parse(postedBody), { targetRevision: 3, reason: 'rollback' });
  assert.equal(postedBody.includes('snapshot'), false);
  assert.equal('saveSnapshot' in api, false);
});

test('snapshot trust boundary rejects versioned or legacy drawing payloads', () => {
  assert.throws(() => validateWorkbookSnapshot({ schema: 'LegacyWorkbookSnapshot', unitId: 'unit-1' }), /Unsupported workbook snapshot schema/);
  assert.throws(() => validateWorkbookSnapshot({
    schema: 'WorkbookSnapshot',
    version: 8,
    unitId: 'unit-1',
    name: 'Workbook',
    dimensionMetrics: { normalFontFamily: 'Calibri', normalFontSizePx: 14.6666666667, maximumDigitWidthPx: 7 },
    calculationSettings: { mode: 'automatic', iterativeCalculation: false, maximumIterations: 100, maximumChange: 0.001, precisionAsDisplayed: false, calculateBeforeSave: true, fullCalculationOnLoad: false },
    dataModel: { sources: [], tables: [], relationships: [], views: [] },
    sheets: [{
      kind: 'worksheet', id: 'sheet-1',
      name: 'Sheet1',
      rowCount: 10,
      columnCount: 10,
      pane: { kind: 'none' },
      defaultRowHeightPx: 20,
      defaultColumnWidthPx: 64,
      cells: {},
      merges: [],
      pivots: [],
      sparklines: [],
      charts: [],
      drawings: [],
      drawingPayloads: {},
      review: { notesByCell: {}, notesById: {}, threadIdsByCell: {}, threadsById: {} },
    }],
  }), /legacy drawing collections/);
});

test('Pivot subtotal contract rejects malformed custom functions and accepts field-owned modes', () => {
  const legacyValuePlacementKey = ['value', 'Field', 'Id'].join('');
  const base = {
    schema: 'PivotDefinition' as const,
    id: 'pivot-subtotals',
    source: { kind: 'worksheet-range' as const, range: { sheetId: 'sheet-1', startRow: 0, endRow: 2, startColumn: 0, endColumn: 1 } },
    target: { sheetId: 'sheet-1', anchor: { row: 4, column: 0 } },
    fieldCatalog: { schema: 'PivotFieldCatalog' as const, fields: [{ fieldId: 'region', name: 'Region', dataType: 'text' as const, ordinal: 0 }, { fieldId: 'amount', name: 'Amount', dataType: 'number' as const, ordinal: 1 }] },
    layout: { rows: [{ fieldId: 'region', subtotal: { mode: 'none' as const } }], columns: [], filters: [], allowMultipleFiltersPerField: true, collation: { locale: 'en-US', sensitivity: 'variant' as const, numeric: false, caseFirst: 'false' as const }, values: [{ valueId: `value:${'amount'}`, fieldId: 'amount', summarizeBy: 'sum' as const }], subtotalLocation: 'bottom' as const, showRowGrandTotals: true, showColumnGrandTotals: true, reportLayout: 'compact' as const },
    refreshPolicy: { mode: 'on-change' as const, preserveFormatting: true, refreshOnLoad: true },
  };
  validatePivotDefinition(base);
  validatePivotDefinition({ ...base, layout: { ...base.layout, values: [{ ...base.layout.values[0], showAs: { kind: 'difference', baseFieldId: 'region', baseItem: { type: 'text', value: 'East' } } }] } });
  assert.throws(() => validatePivotDefinition({ ...base, layout: { ...base.layout, values: [{ ...base.layout.values[0], baseFieldId: 'region', baseItem: 'East' }] } } as never), /unsupported field/);
  assert.throws(() => validatePivotDefinition({ ...base, layout: { ...base.layout, values: [{ ...base.layout.values[0], showAs: { kind: 'difference', baseFieldId: 'amount', baseItem: { type: 'number', value: 1 } } }] } }), /row or column/);
  assert.throws(() => validatePivotDefinition({ ...base, layout: { ...base.layout, values: [{ ...base.layout.values[0], showAs: { kind: 'percentage-difference', baseFieldId: 'region' } }] } }), /requires|baseItem/);
  const highCardinality = {
    ...base,
    fieldCatalog: {
      ...base.fieldCatalog,
      fields: base.fieldCatalog.fields.map((field, index) => ({
        ...field,
        ...(index === 0 ? { values: Array.from({ length: PIVOT_MEMBER_DISPLAY_LIMIT + 1 }, (_, member) => `Member ${member}`) } : {}),
      })),
    },
  };
  validatePivotDefinition(highCardinality);
  assert.throws(() => validatePivotDefinition({
    ...base,
    fieldCatalog: {
      ...base.fieldCatalog,
      fields: [{ ...base.fieldCatalog.fields[0], values: new Array(PIVOT_MAX_MEMBER_COUNT + 1).fill(null) }, base.fieldCatalog.fields[1]],
    },
  }), /Pivot field values are invalid/);
  validatePivotDefinition({ ...base, source: { kind: 'named-range', name: 'SharedName', sheetId: 'sheet-1' } });
  assert.throws(() => validatePivotDefinition({ ...base, source: { kind: 'named-range', name: 'SharedName', sheetId: '' } }), /Pivot named source is invalid/);
  assert.throws(() => validatePivotDefinition({ ...base, source: { kind: 'named-range', name: 'SharedName', sheetId: 'sheet-1', extra: true } }), /unsupported field: extra/);
  assert.throws(() => validatePivotDefinition({ ...base, layout: { ...base.layout, showGrandTotals: true } }), /unsupported field: showGrandTotals/);
  const { collation: _collation, ...legacyCollationLayout } = base.layout;
  assert.throws(() => validatePivotDefinition({ ...base, layout: legacyCollationLayout }), /collation/);
  assert.throws(() => validatePivotDefinition({ ...base, layout: { ...base.layout, collation: { ...base.layout.collation, locale: '***' } } }), /collation/);
  const { allowMultipleFiltersPerField: _allowMultiple, ...legacyLayout } = base.layout;
  assert.throws(() => validatePivotDefinition({ ...base, layout: legacyLayout }), /Pivot layout is invalid/);
  assert.throws(() => validatePivotDefinition({ ...base, layout: {
    ...base.layout,
    allowMultipleFiltersPerField: false,
    filters: [
      { kind: 'manual', family: 'manual', fieldId: 'region', mode: 'all', memberKeys: [] },
      { kind: 'condition', family: 'label', fieldId: 'region', operator: 'contains', value: 'East' },
    ],
  } }), /multiple filters per field are disabled/);
  assert.throws(() => validatePivotDefinition({ ...base, refreshPolicy: { ...base.refreshPolicy, mode: 'manual', refreshOnLoad: true } }), /contradictory/);
  assert.throws(() => validatePivotDefinition({ ...base, layout: { ...base.layout, rows: [{ fieldId: 'region', subtotal: { mode: 'custom', functions: [] } }] } }), /custom subtotal functions/);
  assert.throws(() => validatePivotDefinition({ ...base, layout: { ...base.layout, reportLayout: 'invalid' as never } }), /Pivot layout is invalid/);
  const { reportLayout: _reportLayout, ...legacyReportLayout } = base.layout;
  assert.throws(() => validatePivotDefinition({ ...base, layout: { ...legacyReportLayout, compact: true, repeatLabels: false } as never }), /unsupported field: compact/);
  validatePivotDefinition({ ...base, layout: {
    ...base.layout,
    filters: [
      { kind: 'manual', family: 'manual', fieldId: 'region', scope: 'field', mode: 'all', memberKeys: [] },
      { kind: 'condition', family: 'label', fieldId: 'region', scope: 'report', operator: 'begins-with', value: 'E' },
    ],
  } });
  assert.throws(() => validatePivotDefinition({ ...base, layout: {
    ...base.layout,
    filters: [{ kind: 'manual', family: 'manual', fieldId: 'amount', scope: 'field', mode: 'all', memberKeys: [] }],
  } }), /row or column field/);
  assert.throws(() => validatePivotDefinition({ ...base, layout: {
    ...base.layout,
    values: [base.layout.values[0]!, { ...base.layout.values[0]!, valueId: base.layout.values[0]!.valueId }],
  } }), /invalid or duplicated/);
  assert.throws(() => validatePivotDefinition({ ...base, layout: {
    ...base.layout,
    values: [{ fieldId: 'amount', summarizeBy: 'sum', [legacyValuePlacementKey]: 'amount' }],
  } as never }), /unsupported field/);
  assert.throws(() => validatePivotDefinition({ ...base, layout: {
    ...base.layout,
    rows: [{ fieldId: 'region', sort: { direction: 'ascending', by: 'value', valueId: 'amount' } }],
  } }), /placement identity is invalid/);
  assert.throws(() => validatePivotDefinition({ ...base, layout: {
    ...base.layout,
    rows: [{ fieldId: 'region', sort: { direction: 'ascending', by: 'value' } }],
  } as never }), /requires valueId/);
  assert.throws(() => validatePivotDefinition({ ...base, layout: {
    ...base.layout,
    rows: [{ fieldId: 'region', sort: { direction: 'ascending', by: 'label', valueId: 'value:amount' } }],
  } as never }), /cannot carry valueId/);
  assert.throws(() => validatePivotDefinition({ ...base, layout: {
    ...base.layout,
    rows: [{ fieldId: 'region', sort: { direction: 'ascending' } }],
  } as never }), /Pivot sort is invalid/);
  const topItems = {
    ...base,
    layout: {
      ...base.layout,
      rows: [{ fieldId: 'region' }],
      filters: [{ kind: 'top-items', family: 'top-items', fieldId: 'region', valueId: 'value:amount', direction: 'top', mode: 'items', threshold: 2 }],
    },
  };
  validatePivotDefinition(topItems);
  assert.throws(() => validatePivotDefinition({ ...topItems, layout: { ...topItems.layout, filters: [{ ...topItems.layout.filters[0], count: 2 }] } } as never), /unsupported field/);
  assert.throws(() => validatePivotDefinition({ ...topItems, layout: { ...topItems.layout, filters: [{ ...topItems.layout.filters[0], valueId: undefined }] } } as never), /top-items filter is invalid|placement identity is invalid/);
  assert.throws(() => validatePivotDefinition({ ...topItems, layout: { ...topItems.layout, filters: [{ ...topItems.layout.filters[0], mode: 'average' }] } } as never), /top-items filter is invalid/);
  assert.throws(() => validatePivotDefinition({ ...topItems, layout: { ...topItems.layout, filters: [{ ...topItems.layout.filters[0], mode: 'percent', threshold: 101 }] } } as never), /top-items filter is invalid/);
});

test('Pivot calculated definitions extend the effective field set without catalog duplication', () => {
  const base = {
    schema: 'PivotDefinition' as const,
    id: 'pivot-calculated-layout',
    source: { kind: 'worksheet-range' as const, range: { sheetId: 'sheet-1', startRow: 0, endRow: 2, startColumn: 0, endColumn: 1 } },
    target: { sheetId: 'sheet-1', anchor: { row: 4, column: 0 } },
    fieldCatalog: { schema: 'PivotFieldCatalog' as const, fields: [{ fieldId: 'region', name: 'Region', dataType: 'text' as const, ordinal: 0 }, { fieldId: 'amount', name: 'Amount', dataType: 'number' as const, ordinal: 1 }] },
    layout: { rows: [], columns: [], filters: [], allowMultipleFiltersPerField: true, collation: { locale: 'en-US', sensitivity: 'variant' as const, numeric: false, caseFirst: 'false' as const }, values: [], subtotalLocation: 'bottom' as const, showRowGrandTotals: true, showColumnGrandTotals: true, reportLayout: 'compact' as const },
    refreshPolicy: { mode: 'on-change' as const, preserveFormatting: true, refreshOnLoad: true },
  };
  const calculated = {
    ...base,
    layout: {
      ...base.layout,
      calculatedFields: [{ fieldId: 'calculated:margin', name: 'Margin', formula: '=amount*1.15' }],
      calculatedItems: [{ fieldId: 'calculated-item:amount:premium', targetFieldId: 'amount', name: 'Premium', formula: '=amount*3' }],
      rows: [{ fieldId: 'calculated:margin' }],
      values: [{ valueId: `value:${'calculated:margin'}`, fieldId: 'calculated:margin', summarizeBy: 'sum' as const }],
    },
  };
  validatePivotDefinition(calculated);
  assert.throws(() => validatePivotDefinition({
    ...calculated,
    layout: { ...calculated.layout, calculatedFields: [{ fieldId: 'amount', name: 'Shadow', formula: '=amount' }] },
  }), /duplicated or collides/);
  assert.throws(() => validatePivotDefinition({
    ...calculated,
    layout: { ...calculated.layout, calculatedItems: [{ fieldId: 'calculated-item:missing', targetFieldId: 'missing', name: 'Missing', formula: '=1' }] },
  }), /targetFieldId is invalid/);
  assert.throws(() => validatePivotDefinition({
    ...calculated,
    layout: { ...calculated.layout, calculatedFields: [{ fieldId: 'calculated:bad', name: 'Bad', formula: '', unsupported: true }] as never },
  }), /unsupported field/);
  assert.throws(() => validatePivotDefinition({
    ...calculated,
    layout: { ...calculated.layout, calculatedItems: [{ fieldId: 'calculated-item:margin:bad', targetFieldId: 'calculated:margin', name: 'Bad', formula: '=amount' }] },
  }), /cannot be a calculated field/);
  assert.throws(() => validatePivotDefinition({
    ...calculated,
    layout: { ...calculated.layout, calculatedItems: [{ ...calculated.layout.calculatedItems![0]!, formula: '=North&1' }] },
  }), /unsupported syntax|operand is missing/);
});

test('Pivot calculated item protocol rejects unknown, ambiguous, and cyclic item formulas', () => {
  const base = {
    schema: 'PivotDefinition' as const,
    id: 'pivot-calculated-item-validation',
    source: { kind: 'worksheet-range' as const, range: { sheetId: 'sheet-1', startRow: 0, endRow: 2, startColumn: 0, endColumn: 2 } },
    target: { sheetId: 'sheet-1', anchor: { row: 4, column: 0 } },
    fieldCatalog: { schema: 'PivotFieldCatalog' as const, fields: [
      { fieldId: 'region', name: 'Region', dataType: 'text' as const, ordinal: 0, values: ['North', 'South'] },
      { fieldId: 'category', name: 'Category', dataType: 'text' as const, ordinal: 1, values: ['North', 'South'] },
      { fieldId: 'amount', name: 'Amount', dataType: 'number' as const, ordinal: 2, values: [100, 50] },
    ] },
    layout: { rows: [{ fieldId: 'region' }], columns: [], filters: [], allowMultipleFiltersPerField: true, collation: { locale: 'en-US', sensitivity: 'variant' as const, numeric: false, caseFirst: 'false' as const }, values: [{ valueId: 'value:amount', fieldId: 'amount', summarizeBy: 'sum' as const }], subtotalLocation: 'bottom' as const, showRowGrandTotals: true, showColumnGrandTotals: true, reportLayout: 'compact' as const },
    refreshPolicy: { mode: 'on-change' as const, preserveFormatting: true, refreshOnLoad: true },
  };
  const item = { fieldId: 'calculated-item:region:commission', targetFieldId: 'region', name: 'Commission', formula: '=Region[North]*0.1' };
  validatePivotDefinition({ ...base, layout: { ...base.layout, calculatedItems: [item] } });
  assert.throws(() => validatePivotDefinition({ ...base, layout: { ...base.layout, calculatedItems: [{ ...item, formula: '=Missing+1' }] } }), /unknown item/);
  assert.throws(() => validatePivotDefinition({ ...base, layout: { ...base.layout, calculatedItems: [{ ...item, formula: '=North+1' }] } }), /ambiguous/);
  assert.throws(() => validatePivotDefinition({ ...base, layout: { ...base.layout, calculatedItems: [
    { fieldId: 'calculated-item:region:a', targetFieldId: 'region', name: 'A', formula: '=B+1' },
    { fieldId: 'calculated-item:region:b', targetFieldId: 'region', name: 'B', formula: '=A+1' },
  ] } }), /dependency cycle/);
});

test('Pivot layout-only update accepts a new calculated field without fieldCatalog', () => {
  const envelope = {
    schema: 'OperationEnvelope' as const,
    operationId: 'pivot-calculated-layout-update',
    unitId: 'unit-1',
    clientSequence: 1,
    baseRevision: 0,
    mutations: [{
      id: 'pivot.update',
      sheetId: 'sheet-1',
      params: {
        sheetId: 'sheet-1',
        pivotId: 'pivot-1',
        calculationProof: { schema: 'PivotCalculationProof', pivotId: 'pivot-1', sourceRevision: 'source-1', layoutRevision: 'layout-2', filterRevision: 'filter-1', occupiedRange: { sheetId: 'sheet-1', startRow: 0, endRow: 5, startColumn: 0, endColumn: 2 } },
        previousCalculationProof: { schema: 'PivotCalculationProof', pivotId: 'pivot-1', sourceRevision: 'source-1', layoutRevision: 'layout-1', filterRevision: 'filter-1', occupiedRange: { sheetId: 'sheet-1', startRow: 0, endRow: 4, startColumn: 0, endColumn: 2 } },
        layout: {
          calculatedFields: [{ fieldId: 'calculated:margin', name: 'Margin', formula: '=amount*1.15' }],
          calculatedItems: [],
        },
      },
    }],
    createdAt: new Date().toISOString(),
  };
  assert.deepEqual(validateOperationEnvelope(envelope), envelope);
  assert.throws(() => validateOperationEnvelope({
    ...envelope,
    mutations: [{
      ...envelope.mutations[0]!,
      params: {
        ...envelope.mutations[0]!.params,
        layout: { calculatedFields: [{ fieldId: 'calculated:margin', name: 'Margin', formula: '=amount' }, { fieldId: 'calculated:margin', name: 'Duplicate', formula: '=amount' }] },
      },
    }],
  }), /duplicated or collides/);
});

test('Pivot protocol preserves typed formula-error members and rejects unknown codes', () => {
  const base = {
    schema: 'PivotDefinition' as const,
    id: 'pivot-error-members',
    source: { kind: 'worksheet-range' as const, range: { sheetId: 'sheet-1', startRow: 0, endRow: 2, startColumn: 0, endColumn: 0 } },
    target: { sheetId: 'sheet-1', anchor: { row: 4, column: 0 } },
    fieldCatalog: { schema: 'PivotFieldCatalog' as const, fields: [{ fieldId: 'member', name: 'Member', dataType: 'error' as const, ordinal: 0, values: [{ kind: 'error' as const, code: '#N/A' as const }] }] },
    layout: { rows: [{ fieldId: 'member' }], columns: [], filters: [{ kind: 'manual' as const, family: 'manual' as const, fieldId: 'member', mode: 'include' as const, memberKeys: [{ type: 'error' as const, value: '#N/A' as const }] }], allowMultipleFiltersPerField: true, collation: { locale: 'en-US', sensitivity: 'variant' as const, numeric: false, caseFirst: 'false' as const }, values: [{ valueId: `value:${'member'}`, fieldId: 'member', summarizeBy: 'count' as const }], subtotalLocation: 'bottom' as const, showRowGrandTotals: true, showColumnGrandTotals: true, reportLayout: 'compact' as const },
    refreshPolicy: { mode: 'on-change' as const, preserveFormatting: true, refreshOnLoad: true },
  };
  validatePivotDefinition(base);
  validatePivotDefinition({ ...base, presentation: {
    styleOptions: { showRowHeaders: true, showColumnHeaders: true, showRowStripes: false, showColumnStripes: false, showLastColumn: false },
    displayOptions: { fillEmptyCells: true, emptyCellText: '—', showErrorValues: true, errorCellText: 'ERR', showFieldHeaders: false, autoFitColumnsOnUpdate: false },
  } });
  validatePivotDefinition({ ...base, layout: { ...base.layout, rows: [{ fieldId: 'member', group: { kind: 'date', unit: 'year', units: ['year', 'quarter', 'month'], startOfWeek: 1 } }] } });
  assert.throws(() => validatePivotDefinition({ ...base, layout: { ...base.layout, rows: [{ fieldId: 'member', group: { kind: 'date', unit: 'year', units: ['year', 'year'] } }] } }), /date group is invalid/);
  validatePivotDefinition({ ...base, layout: { ...base.layout, filters: [{ kind: 'condition', family: 'date', fieldId: 'member', operator: 'between', value: '2024-01-01', value2: '2024-12-31' }] } });
  validatePivotDefinition({ ...base, layout: { ...base.layout, filters: [{ kind: 'condition', family: 'value', fieldId: 'member', valueId: `value:${'member'}`, operator: 'not-between', value: 10, value2: 20 }] } });
  assert.throws(() => validatePivotDefinition({ ...base, layout: { ...base.layout, filters: [{ kind: 'condition', family: 'value', fieldId: 'member', operator: 'greater-than', value: 10 }] } }), /value filter requires valueId/);
  assert.throws(() => validatePivotDefinition({ ...base, layout: { ...base.layout, filters: [{ kind: 'condition', family: 'value', fieldId: 'member', valueId: 'missing-placement', operator: 'greater-than', value: 10 }] } }), /placement identity is invalid/);
  assert.throws(() => validatePivotDefinition({ ...base, layout: { ...base.layout, filters: [{ kind: 'condition', family: 'date', fieldId: 'member', operator: 'between', value: '2024-01-01' }] } }), /range filter requires two bounds/);
  assert.throws(() => validatePivotDefinition({ ...base, presentation: {
    styleOptions: { showRowHeaders: true, showColumnHeaders: true, showRowStripes: false, showColumnStripes: false, showLastColumn: false },
    displayOptions: { fillEmptyCells: true, emptyCellText: '—', showErrorValues: true, errorCellText: 'ERR', showFieldHeaders: false, autoFitColumnsOnUpdate: false, unsupported: true },
  } }), /unsupported field/);
  assert.throws(() => validatePivotDefinition({ ...base, fieldCatalog: { ...base.fieldCatalog, fields: [{ ...base.fieldCatalog.fields[0]!, values: [{ kind: 'error', code: '#NOT-AN-EXCEL-CODE' }] }] } }), /invalid/);
});

test('Pivot worksheet-ranges require stable source nodes and graph endpoints', () => {
  const base = {
    schema: 'PivotDefinition' as const,
    id: 'pivot-relations',
    source: {
      kind: 'worksheet-ranges' as const,
      ranges: [
        { sourceId: 'orders', range: { sheetId: 'sheet-1', startRow: 0, endRow: 2, startColumn: 0, endColumn: 2 } },
        { sourceId: 'customers', range: { sheetId: 'sheet-1', startRow: 0, endRow: 2, startColumn: 4, endColumn: 5 } },
        { sourceId: 'products', range: { sheetId: 'sheet-1', startRow: 0, endRow: 2, startColumn: 7, endColumn: 8 } },
      ],
      relationships: [
        { id: 'orders-customers', left: { sourceId: 'orders', fieldId: 'source:orders:column:0' }, right: { sourceId: 'customers', fieldId: 'source:customers:column:0' }, join: 'left' as const },
        { id: 'orders-products', left: { sourceId: 'orders', fieldId: 'source:orders:column:1' }, right: { sourceId: 'products', fieldId: 'source:products:column:0' }, join: 'left' as const },
      ],
    },
    target: { sheetId: 'sheet-1', anchor: { row: 8, column: 0 } },
    fieldCatalog: { schema: 'PivotFieldCatalog' as const, fields: [] },
    layout: { rows: [], columns: [], filters: [], allowMultipleFiltersPerField: true, collation: { locale: 'en-US', sensitivity: 'variant' as const, numeric: false, caseFirst: 'false' as const }, values: [], subtotalLocation: 'bottom' as const, showRowGrandTotals: true, showColumnGrandTotals: true, reportLayout: 'compact' as const },
    refreshPolicy: { mode: 'on-change' as const, preserveFormatting: true, refreshOnLoad: true },
  };
  validatePivotDefinition(base);
  const inner = {
    ...base,
    source: {
      ...base.source,
      relationships: base.source.relationships.map((relationship) => ({ ...relationship, join: 'inner' as const })),
    },
  };
  validatePivotDefinition(inner);
  assert.throws(() => validatePivotDefinition({ ...base, source: { ...base.source, ranges: [{ sourceId: 'orders', range: base.source.ranges[0]!.range }, { sourceId: 'orders', range: base.source.ranges[1]!.range }, base.source.ranges[2]! ], relationships: [] } }), /sourceId is duplicated/);
  assert.throws(() => validatePivotDefinition({ ...base, source: { ...base.source, relationships: [{ ...base.source.relationships[0]!, left: { sheetId: 'sheet-1', fieldId: 'source:orders:column:0' } }] } }), /Pivot relationship field/);
  assert.throws(() => validatePivotDefinition({ ...base, source: { ...base.source, relationships: [...base.source.relationships, { id: 'products-customers', left: { sourceId: 'products', fieldId: 'source:products:column:0' }, right: { sourceId: 'customers', fieldId: 'source:customers:column:0' }, join: 'inner' as const }] } }), /graph contains a cycle/);
});
