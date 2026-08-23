import type { RangeRef, SheetId, UnitId } from './index';

/**
 * Canonical persisted print state.  This type belongs to the workbook
 * domain, rather than the print UI, so a snapshot/replay does not depend on a
 * host-side cache.
 */
export interface PrintDocumentSnapshot {
  schema: 'PrintDocument';
  unitId: UnitId;
  sheetId: SheetId;
  pageSetup: {
    paperSize: 'letter' | 'a4' | 'a3' | 'legal' | 'custom';
    orientation: 'portrait' | 'landscape';
    margins: {
      top: number;
      right: number;
      bottom: number;
      left: number;
      header: number;
      footer: number;
    };
    scale: number;
    fitToWidth?: number;
    fitToHeight?: number;
    printGridlines: boolean;
    printHeadings: boolean;
    centerHorizontally: boolean;
    centerVertically: boolean;
    headerText?: string;
    footerText?: string;
  };
  printAreas: Array<{ sheetId: SheetId; range: RangeRef }>;
  pageBreaks: Array<{ sheetId: SheetId; row?: number; column?: number }>;
  repeatRows?: { start: number; end: number };
  repeatColumns?: { start: number; end: number };
}

export interface QueryStepSnapshot {
  id: string;
  kind: 'source' | 'filter' | 'select-columns' | 'rename-column' | 'sort' | 'group-by' | 'join' | 'pivot' | 'custom';
  name: string;
  config: Record<string, unknown>;
  enabled: boolean;
}

/**
 * Persistence-safe query definition.  Connector credentials are never
 * represented by this contract; secret fields must be redacted markers.
 */
export interface QueryDefinitionSnapshot {
  schema: 'QueryDefinition';
  id: string;
  name: string;
  connectorId: string;
  connectorConfig: Record<string, unknown>;
  steps: QueryStepSnapshot[];
  refreshOnOpen?: boolean;
  refreshPolicy?: { mode: 'manual' | 'on-open' | 'interval'; intervalMs?: number };
  /** Last materialization target; this lets a session refresh after reload. */
  lastTarget?: QueryLoadTargetSnapshot;
  sourceRevision: number;
}

export interface QueryLoadTargetSnapshot {
  kind: 'range' | 'sheet-table' | 'workbook-table' | 'pivot-source';
  sheetId?: string;
  range?: { startRow: number; startColumn: number; endRow?: number; endColumn?: number };
  tableId?: string;
  pivotId?: string;
}

const QUERY_STEP_KINDS = new Set<QueryStepSnapshot['kind']>([
  'source', 'filter', 'select-columns', 'rename-column', 'sort', 'group-by', 'join', 'pivot', 'custom',
]);

const SECRET_KEY = /(?:pass(word)?|secret|token|api[-_]?key|credential|authorization|private[-_]?key|client[-_]?secret)/i;

function assertSafeConfig(value: unknown, key?: string, path = 'connectorConfig'): void {
  if (key && SECRET_KEY.test(key)) {
    if (value !== '[redacted]') throw new Error(`Query connector credential must be redacted at ${path}.${key}`);
    return;
  }
  if (Array.isArray(value)) {
    value.forEach((entry, index) => assertSafeConfig(entry, undefined, `${path}[${index}]`));
    return;
  }
  if (value && typeof value === 'object') {
    for (const [entryKey, entryValue] of Object.entries(value as Record<string, unknown>)) {
      assertSafeConfig(entryValue, entryKey, `${path}.${entryKey}`);
    }
    return;
  }
  if (typeof value === 'function' || typeof value === 'bigint' || typeof value === 'symbol' || value === undefined) {
    throw new Error(`Query connector configuration is not serializable at ${path}`);
  }
}

export function normalizePrintDocumentSnapshot(document: PrintDocumentSnapshot): PrintDocumentSnapshot {
  if (!document || document.schema !== 'PrintDocument' || !document.unitId || !document.sheetId) {
    throw new Error('Invalid print document snapshot');
  }
  if (!document.pageSetup || !document.pageSetup.margins) throw new Error('Print document page setup is required');
  const margins = document.pageSetup.margins;
  const numbers = [margins.top, margins.right, margins.bottom, margins.left, margins.header, margins.footer, document.pageSetup.scale];
  if (!numbers.every((value) => Number.isFinite(value) && value >= 0)) throw new Error('Invalid print page setup');
  if (document.pageSetup.scale <= 0 || document.pageSetup.scale > 400) throw new Error('Print scale must be between 1 and 400');
  if (document.pageSetup.orientation !== 'portrait' && document.pageSetup.orientation !== 'landscape') throw new Error('Invalid print orientation');
  if (!['letter', 'a4', 'a3', 'legal', 'custom'].includes(document.pageSetup.paperSize)) throw new Error('Invalid print paper size');
  const validRange = (range: RangeRef): boolean => Boolean(range)
    && range.sheetId === document.sheetId
    && Number.isInteger(range.startRow) && range.startRow >= 0
    && Number.isInteger(range.endRow) && range.endRow >= range.startRow
    && Number.isInteger(range.startColumn) && range.startColumn >= 0
    && Number.isInteger(range.endColumn) && range.endColumn >= range.startColumn;
  if (!Array.isArray(document.printAreas) || document.printAreas.some((area) => area.sheetId !== document.sheetId || !validRange(area.range))) {
    throw new Error('Print areas must target their document sheet');
  }
  if (!Array.isArray(document.pageBreaks) || document.pageBreaks.some((pageBreak) => {
    const hasRow = pageBreak.row !== undefined;
    const hasColumn = pageBreak.column !== undefined;
    return pageBreak.sheetId !== document.sheetId || hasRow === hasColumn || (hasRow && (!Number.isInteger(pageBreak.row) || pageBreak.row! < 0)) || (hasColumn && (!Number.isInteger(pageBreak.column) || pageBreak.column! < 0));
  })) throw new Error('Invalid print page breaks');
  const normalizeSpan = (span: { start: number; end: number } | undefined) => {
    if (!span) return undefined;
    if (!Number.isInteger(span.start) || !Number.isInteger(span.end) || span.start < 0 || span.end < span.start) throw new Error('Invalid print repeat span');
    return { start: span.start, end: span.end };
  };
  return structuredClone({
    schema: 'PrintDocument' as const,
    unitId: document.unitId,
    sheetId: document.sheetId,
    pageSetup: { ...structuredClone(document.pageSetup), margins: { ...margins } },
    printAreas: document.printAreas.map((area) => ({ sheetId: area.sheetId, range: { ...area.range } })),
    pageBreaks: document.pageBreaks.map((pageBreak) => pageBreak.row !== undefined ? { sheetId: pageBreak.sheetId, row: pageBreak.row } : { sheetId: pageBreak.sheetId, column: pageBreak.column }),
    repeatRows: normalizeSpan(document.repeatRows),
    repeatColumns: normalizeSpan(document.repeatColumns),
  });
}

export function normalizeQueryDefinitionSnapshot(definition: QueryDefinitionSnapshot): QueryDefinitionSnapshot {
  if (!definition || typeof definition !== 'object' || definition.schema !== 'QueryDefinition'
    || typeof definition.id !== 'string' || !definition.id.trim()
    || typeof definition.name !== 'string' || !definition.name.trim()
    || typeof definition.connectorId !== 'string' || !definition.connectorId.trim()) {
    throw new Error('Invalid query definition snapshot');
  }
  if (!Number.isSafeInteger(definition.sourceRevision) || definition.sourceRevision < 0) throw new Error('Invalid query source revision');
  if (!Array.isArray(definition.steps)) throw new Error('Query definition steps must be an array');
  for (const step of definition.steps) {
    if (!step || typeof step.id !== 'string' || !step.id.trim() || typeof step.name !== 'string' || !step.name.trim() || typeof step.kind !== 'string' || !step.kind.trim() || typeof step.enabled !== 'boolean' || !step.config || typeof step.config !== 'object' || Array.isArray(step.config)) {
      throw new Error(`Invalid query step in ${definition.id}`);
    }
    if (!QUERY_STEP_KINDS.has(step.kind as QueryStepSnapshot['kind'])) throw new Error(`Unknown query step kind in ${definition.id}: ${step.kind}`);
  }
  assertSafeConfig(definition.connectorConfig);
  if (definition.refreshPolicy) {
    if (!['manual', 'on-open', 'interval'].includes(definition.refreshPolicy.mode)) throw new Error('Invalid query refresh mode');
    if (definition.refreshPolicy.mode === 'interval' && (!Number.isSafeInteger(definition.refreshPolicy.intervalMs) || definition.refreshPolicy.intervalMs! <= 0)) throw new Error('Invalid query refresh interval');
  }
  if (definition.lastTarget) {
    if (!['range', 'sheet-table', 'workbook-table', 'pivot-source'].includes(definition.lastTarget.kind)) throw new Error('Invalid query load target kind');
    if (definition.lastTarget.sheetId !== undefined && (typeof definition.lastTarget.sheetId !== 'string' || !definition.lastTarget.sheetId.trim())) throw new Error('Invalid query load target sheet');
    if (definition.lastTarget.range) {
      const range = definition.lastTarget.range;
      if (!Number.isSafeInteger(range.startRow) || range.startRow < 0 || !Number.isSafeInteger(range.startColumn) || range.startColumn < 0
        || (range.endRow !== undefined && (!Number.isSafeInteger(range.endRow) || range.endRow < range.startRow))
        || (range.endColumn !== undefined && (!Number.isSafeInteger(range.endColumn) || range.endColumn < range.startColumn))) throw new Error('Invalid query load target range');
    }
  }
  return structuredClone({
    schema: 'QueryDefinition' as const,
    id: definition.id.trim(),
    name: definition.name.trim(),
    connectorId: definition.connectorId.trim(),
    connectorConfig: structuredClone(definition.connectorConfig),
    steps: definition.steps.map((step) => ({ id: step.id.trim(), kind: step.kind as QueryStepSnapshot['kind'], name: step.name.trim(), config: structuredClone(step.config), enabled: step.enabled })),
    ...(definition.refreshOnOpen === undefined ? {} : { refreshOnOpen: definition.refreshOnOpen }),
    ...(definition.refreshPolicy === undefined ? {} : { refreshPolicy: structuredClone(definition.refreshPolicy) }),
    ...(definition.lastTarget === undefined ? {} : { lastTarget: structuredClone(definition.lastTarget) }),
    sourceRevision: definition.sourceRevision,
  });
}
