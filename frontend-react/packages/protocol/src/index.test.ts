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
          version: 5,
          unitId: 'unit-1',
          name: 'Workbook',
          dimensionMetrics: { normalFontFamily: 'Calibri', normalFontSizePx: 14.6666666667, maximumDigitWidthPx: 7 },
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
          version: 5,
          unitId: 'unit-guest',
          name: 'Guest workbook',
          dimensionMetrics: { normalFontFamily: 'Calibri', normalFontSizePx: 14.6666666667, maximumDigitWidthPx: 7 },
          dataModel: { sources: [], tables: [], relationships: [], views: [] },
          sheets: [{
            kind: 'worksheet', id: 'sheet-1', name: 'Sheet1', rowCount: 10, columnCount: 10,
            cells: {}, merges: [], pane: { kind: 'none' }, defaultRowHeightPx: 20, defaultColumnWidthPx: 64,
            pivots: [], sparklines: [], drawings: [], drawingPayloads: {},
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
    version: 5,
    unitId: 'unit-1',
    name: 'Workbook',
    dimensionMetrics: { normalFontFamily: 'Calibri', normalFontSizePx: 14.6666666667, maximumDigitWidthPx: 7 },
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
    }],
  }), /legacy drawing collections/);
});

test('Pivot subtotal contract rejects malformed custom functions and accepts field-owned modes', () => {
  const base = {
    schema: 'PivotDefinition' as const,
    id: 'pivot-subtotals',
    source: { kind: 'worksheet-range' as const, range: { sheetId: 'sheet-1', startRow: 0, endRow: 2, startColumn: 0, endColumn: 1 } },
    target: { sheetId: 'sheet-1', anchor: { row: 4, column: 0 } },
    fieldCatalog: { schema: 'PivotFieldCatalog' as const, fields: [{ fieldId: 'region', name: 'Region', dataType: 'text' as const, ordinal: 0 }, { fieldId: 'amount', name: 'Amount', dataType: 'number' as const, ordinal: 1 }] },
    layout: { rows: [{ fieldId: 'region', subtotal: { mode: 'none' as const } }], columns: [], filters: [], allowMultipleFiltersPerField: true, values: [{ fieldId: 'amount', summarizeBy: 'sum' as const }], subtotalLocation: 'bottom' as const, showGrandTotals: true, compact: true, repeatLabels: false },
    refreshPolicy: { mode: 'on-change' as const, preserveFormatting: true, refreshOnLoad: true },
  };
  validatePivotDefinition(base);
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
    layout: { rows: [], columns: [], filters: [], allowMultipleFiltersPerField: true, values: [], subtotalLocation: 'bottom' as const, showGrandTotals: true, compact: true, repeatLabels: false },
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
