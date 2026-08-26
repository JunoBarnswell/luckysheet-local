import { createPivotCollator, normalizePivotRefreshPolicy, parsePivotCalculatedItemFormula, PIVOT_MAX_MEMBER_COUNT } from '@react-sheets/core-model';
import type {
  DataSourceManifest,
  PivotDefinition,
  PivotPresentation,
  RangeRef,
  SheetDataRegion,
  TableScalar,
  WorkbookSnapshot,
  WorkbookTableBlock,
} from '@react-sheets/core-model';
import type { AssetRef } from '@react-sheets/core-model';
import { isWorkbookCalculationSettings } from '@react-sheets/formula-engine';
import {
  CONTRACT_ERROR_CODES,
  MAX_WORKBOOK_NAME_LENGTH,
  WORKBOOK_SNAPSHOT_SCHEMA,
  WORKBOOK_SNAPSHOT_VERSION,
  mutationPermission,
  commandPermission,
  type PermissionCapability,
  type PermissionPolicy,
  type ProtectionAction,
  mutationCapability,
  type ContractErrorCode,
} from './generated-contract';

export {
  CONTRACT_ERROR_CODES,
  MAX_WORKBOOK_NAME_LENGTH,
  WORKBOOK_SNAPSHOT_SCHEMA,
  WORKBOOK_SNAPSHOT_VERSION,
  mutationPermission,
  commandPermission,
  mutationCapability,
} from './generated-contract';

export type { PermissionCapability, PermissionPolicy, ProtectionAction } from './generated-contract';

export type ProtocolErrorCode = ContractErrorCode | 'AUTH_CONFIGURATION_ERROR';

export interface ApiError {
  code: ProtocolErrorCode;
  message: string;
  details?: Record<string, unknown>;
}

export interface CollaborationMutation {
  id: string;
  sheetId: string;
  params: unknown;
  affectedRanges: RangeRef[];
}

/**
 * The only client-authored operation wire contract.
 *
 * Actor identity and affected ranges intentionally do not exist on this
 * request type.  The server obtains the actor from the verified bearer token
 * and derives ranges from the authoritative workbook/runtime state.  Keeping
 * those values out of the request makes it impossible for a client to turn a
 * claimed actor or range into an authorization decision by accident.
 */
export const OPERATION_ENVELOPE_SCHEMA = 'OperationEnvelope' as const;

export interface OperationMutation {
  id: string;
  sheetId: string;
  params: unknown;
}

export interface OperationIntent {
  type: 'undo';
  targetOperationId: string;
  targetBaseRevision: number;
}

export interface OperationEnvelope {
  schema: typeof OPERATION_ENVELOPE_SCHEMA;
  operationId: string;
  unitId: string;
  clientSequence: number;
  baseRevision: number;
  mutations: OperationMutation[];
  createdAt: string;
  intent?: OperationIntent;
}

/**
 * Canonical metadata-only mutations for block-backed workbook data. The
 * manifest and region are JSON descriptors; block bytes are transferred only
 * through the dedicated data-block endpoint and never through an operation.
 */
export const DATA_SOURCE_MUTATION_IDS = [
  'dataSource.add',
  'dataSource.update',
  'dataSource.remove',
  'dataRegion.add',
  'dataRegion.remove',
] as const;

export type DataSourceMutationId = typeof DATA_SOURCE_MUTATION_IDS[number];

export type DataSourceMutationParams =
  | { source: DataSourceManifest }
  | { sourceId: string }
  | { region: SheetDataRegion }
  | { regionId: string };

/** Server-authored metadata added after authentication, validation and commit. */
export interface CommittedOperationMutation extends OperationMutation {
  /** Authoritative range calculated by the server; never read from client input. */
  affectedRanges: RangeRef[];
}

export interface CommittedOperationEnvelope extends Omit<OperationEnvelope, 'mutations'> {
  actorId: string;
  origin: 'client' | 'system';
  revision: number;
  committedAt: string;
  mutations: CommittedOperationMutation[];
}

export type WorkbookAclRole = 'owner' | 'editor' | 'commenter' | 'viewer';

export interface WorkbookAclRecord {
  unitId: string;
  subject: string;
  role: WorkbookAclRole;
  createdAt: string;
  updatedAt: string;
}

/** Server-calculated access projection for the authenticated or guest session. */
export interface WorkbookAccessResponse {
  unitId: string;
  role: WorkbookAclRole;
}

export interface OperationCommitResponse {
  operation: CommittedOperationEnvelope;
}

export interface CheckpointResponse {
  created: boolean;
  snapshot: SnapshotResponse;
}

/**
 * Authentication is deliberately supplied by the host application.  The
 * protocol package never reads cookies, localStorage or a client-provided
 * actor id.  A missing/empty token is an authentication failure, not an
 * unauthenticated fallback.
 */
export type AuthTokenProvider = () => string | null | Promise<string | null>;
/** Opaque, server-issued workbook sharing token for a guest session. */
export type ShareTokenProvider = () => string | null | Promise<string | null>;

/** Request controls intentionally expose cancellation but not transport details. */
export interface ApiRequestOptions {
  signal?: AbortSignal;
}

/** Bounded server response used by catalog and revision history endpoints. */
export interface CursorPage<T> {
  items: T[];
  nextCursor: string | null;
}

export const DEFAULT_PAGE_LIMIT = 20 as const;
export const MAX_PAGE_LIMIT = 50 as const;

export function normalizePageLimit(value: number | undefined): number {
  if (value === undefined) return DEFAULT_PAGE_LIMIT;
  if (!Number.isSafeInteger(value) || value < 1 || value > MAX_PAGE_LIMIT) {
    throw new Error(`limit must be an integer between 1 and ${MAX_PAGE_LIMIT}`);
  }
  return value;
}

export interface WorkbookApiClientOptions {
  baseUrl?: string;
  authTokenProvider?: AuthTokenProvider;
  shareTokenProvider?: ShareTokenProvider;
  fetchImpl?: typeof fetch;
}

export class ApiRequestError extends Error {
  readonly status: number;
  readonly code: ProtocolErrorCode;
  readonly details?: Record<string, unknown>;

  constructor(
    message: string,
    status: number,
    code: ProtocolErrorCode = 'INTERNAL_ERROR',
    details?: Record<string, unknown>,
  ) {
    super(message);
    this.name = 'ApiRequestError';
    this.status = status;
    this.code = code;
    this.details = details;
  }
}

export class AuthenticationRequiredError extends Error {
  readonly status = 401;
  readonly code = 'UNAUTHENTICATED' as const;

  constructor(message = 'Authentication or workbook share token is required') {
    super(message);
    this.name = 'AuthenticationRequiredError';
  }
}

const isNonEmptyString = (value: unknown): value is string => typeof value === 'string' && value.trim().length > 0;

function isRangeRef(value: unknown): value is RangeRef {
  if (!value || typeof value !== 'object') return false;
  const range = value as Record<string, unknown>;
  return isNonEmptyString(range.sheetId)
    && Number.isSafeInteger(range.startRow) && Number(range.startRow) >= 0
    && Number.isSafeInteger(range.endRow) && Number(range.endRow) >= Number(range.startRow)
    && Number.isSafeInteger(range.startColumn) && Number(range.startColumn) >= 0
    && Number.isSafeInteger(range.endColumn) && Number(range.endColumn) >= Number(range.startColumn);
}

function isPivotSourceRange(value: unknown): boolean {
  if (!value || typeof value !== 'object' || Array.isArray(value)) return false;
  const sourceRange = value as Record<string, unknown>;
  return isNonEmptyString(sourceRange.sourceId) && isRangeRef(sourceRange.range);
}

function assertMetadataOnly(value: unknown, path: string): void {
  if (Array.isArray(value)) {
    value.forEach((entry, index) => assertMetadataOnly(entry, `${path}[${index}]`));
    return;
  }
  if (!value || typeof value !== 'object') return;
  for (const [key, child] of Object.entries(value as Record<string, unknown>)) {
    if (key === 'bytes' || key === 'buffer' || key === 'arrayBuffer' || key === 'content') {
      throw new Error(`${path}.${key} is not allowed in metadata-only operations`);
    }
    assertMetadataOnly(child, `${path}.${key}`);
  }
}

function validateExactKeys(value: Record<string, unknown>, allowed: readonly string[], label: string): void {
  for (const key of Object.keys(value)) {
    if (!allowed.includes(key)) throw new Error(`${label} contains unsupported field: ${key}`);
  }
}

function validateReviewSnapshot(value: unknown, sheetId: string): void {
  const review = requireRecord(value, 'WorkbookSnapshot review');
  const notesByCell = requireRecord(review.notesByCell, 'WorkbookSnapshot review notesByCell');
  const notesById = requireRecord(review.notesById, 'WorkbookSnapshot review notesById');
  const threadIdsByCell = requireRecord(review.threadIdsByCell, 'WorkbookSnapshot review threadIdsByCell');
  const threadsById = requireRecord(review.threadsById, 'WorkbookSnapshot review threadsById');
  const keyPattern = /^(0|[1-9][0-9]*):(0|[1-9][0-9]*)$/;
  const noteIds = new Set(Object.keys(notesById));
  for (const [id, note] of Object.entries(notesById)) {
    const entry = requireRecord(note, `WorkbookSnapshot review note ${id}`);
    if (!isNonEmptyString(id) || !isNonEmptyString(entry.id) || entry.id !== id) throw new Error(`WorkbookSnapshot review note identity is invalid: ${id}`);
  }
  const indexedNotes = new Set<string>();
  for (const [key, id] of Object.entries(notesByCell)) {
    const [rowText, columnText] = key.split(':');
    const row = Number(rowText);
    const column = Number(columnText);
    if (!keyPattern.test(key) || !Number.isSafeInteger(row) || row < 0 || row > 1_048_575 || !Number.isSafeInteger(column) || column < 0 || column > 16_383
      || !isNonEmptyString(id) || !noteIds.has(id) || indexedNotes.has(id)) throw new Error(`WorkbookSnapshot review note index is invalid: ${key}`);
    indexedNotes.add(id);
  }
  if (indexedNotes.size !== noteIds.size) throw new Error(`WorkbookSnapshot review contains an unindexed note on ${sheetId}`);
  const indexedThreads = new Set<string>();
  for (const [key, ids] of Object.entries(threadIdsByCell)) {
    const [rowText, columnText] = key.split(':');
    const row = Number(rowText);
    const column = Number(columnText);
    if (!keyPattern.test(key) || !Number.isSafeInteger(row) || row < 0 || row > 1_048_575 || !Number.isSafeInteger(column) || column < 0 || column > 16_383
      || !Array.isArray(ids) || new Set(ids).size !== ids.length) throw new Error(`WorkbookSnapshot review thread index is invalid: ${key}`);
    for (const id of ids) {
      if (!isNonEmptyString(id) || !threadsById[id]) throw new Error(`WorkbookSnapshot review thread index references missing id: ${id}`);
      const thread = requireRecord(threadsById[id], `WorkbookSnapshot review thread ${id}`);
      if (thread.id !== id || thread.sheetId !== sheetId || thread.row !== row || thread.column !== column || indexedThreads.has(id)) {
        throw new Error(`WorkbookSnapshot review thread index is incompatible: ${id}`);
      }
      indexedThreads.add(id);
    }
  }
  for (const [id, value] of Object.entries(threadsById)) {
    const thread = requireRecord(value, `WorkbookSnapshot review thread ${id}`);
    if (!isNonEmptyString(id) || thread.id !== id || thread.sheetId !== sheetId || !Number.isSafeInteger(thread.row) || Number(thread.row) < 0 || Number(thread.row) > 1_048_575
      || !Number.isSafeInteger(thread.column) || Number(thread.column) < 0 || Number(thread.column) > 16_383 || !indexedThreads.has(id)) {
      throw new Error(`WorkbookSnapshot review thread identity is invalid: ${id}`);
    }
  }
}

function requireRecord(value: unknown, label: string): Record<string, unknown> {
  if (!value || typeof value !== 'object' || Array.isArray(value)) throw new Error(`${label} must be an object`);
  return value as Record<string, unknown>;
}

const PIVOT_FORMULA_ERROR_CODES = new Set([
  '#NULL!', '#DIV/0!', '#VALUE!', '#REF!', '#NAME?', '#NUM!', '#N/A',
  '#CALC!', '#BLOCKED!', '#SPILL!', '#PARSE!',
]);

function validatePivotScalar(value: unknown, label: string): void {
  if (value === null || typeof value === 'string' || typeof value === 'boolean'
    || (typeof value === 'number' && Number.isFinite(value))) return;
  const error = requireRecord(value, label);
  validateExactKeys(error, ['kind', 'code', 'message'], label);
  if (error.kind !== 'error' || typeof error.code !== 'string' || !PIVOT_FORMULA_ERROR_CODES.has(error.code)
    || (error.message !== undefined && typeof error.message !== 'string')) throw new Error(`${label} is invalid`);
}

function validatePivotMemberKey(value: unknown, label: string): void {
  const member = requireRecord(value, label);
  validateExactKeys(member, ['type', 'value'], label);
  if (!['text', 'number', 'boolean', 'blank', 'error'].includes(String(member.type))) throw new Error(`${label}.type is invalid`);
  if (member.type === 'blank' && member.value !== null) throw new Error(`${label}.value must be null for blank`);
  if (member.type === 'text' && typeof member.value !== 'string') throw new Error(`${label}.value must be text`);
  if (member.type === 'number' && (typeof member.value !== 'number' || !Number.isFinite(member.value))) throw new Error(`${label}.value must be a finite number`);
  if (member.type === 'boolean' && typeof member.value !== 'boolean') throw new Error(`${label}.value must be boolean`);
  if (member.type === 'error' && (typeof member.value !== 'string' || !PIVOT_FORMULA_ERROR_CODES.has(member.value))) throw new Error(`${label}.value must be a formula error code`);
}

function validatePivotShowAs(value: unknown, fieldIds: ReadonlySet<string>, axisFieldIds: ReadonlySet<string>): void {
  if (value === undefined) return;
  const showAs = requireRecord(value, 'Pivot showAs');
  const kind = showAs.kind;
  const totalKinds = new Set(['normal', 'grand-percentage', 'row-percentage', 'column-percentage', 'parent-percentage', 'index']);
  if (typeof kind !== 'string') throw new Error('Pivot showAs kind is invalid');
  if (totalKinds.has(kind)) {
    validateExactKeys(showAs, ['kind'], 'Pivot showAs');
    return;
  }
  if (kind === 'difference' || kind === 'percentage-difference') {
    validateExactKeys(showAs, ['kind', 'baseFieldId', 'baseItem'], 'Pivot difference showAs');
    if (!isNonEmptyString(showAs.baseFieldId) || !fieldIds.has(showAs.baseFieldId) || !axisFieldIds.has(showAs.baseFieldId)) {
      throw new Error('Pivot difference showAs baseFieldId must target a row or column field');
    }
    if (showAs.baseItem === 'previous' || showAs.baseItem === 'next') return;
    validatePivotMemberKey(showAs.baseItem, 'Pivot difference showAs baseItem');
    return;
  }
  if (kind === 'running-total' || kind === 'percentage-running-total') {
    validateExactKeys(showAs, ['kind', 'baseFieldId'], 'Pivot running-total showAs');
    if (!isNonEmptyString(showAs.baseFieldId) || !fieldIds.has(showAs.baseFieldId) || !axisFieldIds.has(showAs.baseFieldId)) {
      throw new Error('Pivot running-total showAs baseFieldId must target a row or column field');
    }
    return;
  }
  if (kind === 'rank') {
    validateExactKeys(showAs, ['kind', 'baseFieldId', 'direction'], 'Pivot rank showAs');
    if (!isNonEmptyString(showAs.baseFieldId) || !fieldIds.has(showAs.baseFieldId) || !axisFieldIds.has(showAs.baseFieldId)
      || !['ascending', 'descending'].includes(String(showAs.direction))) throw new Error('Pivot rank showAs is invalid');
    return;
  }
  throw new Error('Pivot showAs kind is unsupported');
}

function validatePivotSource(value: unknown): void {
  const source = requireRecord(value, 'Pivot source');
  if (source.kind === 'worksheet-range') {
    validateExactKeys(source, ['kind', 'range'], 'Pivot worksheet source');
    if (!isRangeRef(source.range)) throw new Error('Pivot worksheet source range is invalid');
    return;
  }
  if (source.kind === 'worksheet-ranges') {
    validateExactKeys(source, ['kind', 'ranges', 'relationships'], 'Pivot worksheet sources');
    if (!Array.isArray(source.ranges) || !source.ranges.length || !source.ranges.every(isPivotSourceRange) || !Array.isArray(source.relationships)) {
      throw new Error('Pivot worksheet sources are invalid');
    }
    const sourceIds = new Set<string>();
    for (const rawRange of source.ranges) {
      const sourceRange = requireRecord(rawRange, 'Pivot worksheet source range');
      validateExactKeys(sourceRange, ['sourceId', 'range'], 'Pivot worksheet source range');
      const sourceId = String(sourceRange.sourceId);
      if (sourceIds.has(sourceId)) throw new Error('Pivot worksheet sourceId is duplicated');
      sourceIds.add(sourceId);
    }
    const parent = new Map([...sourceIds].map((sourceId) => [sourceId, sourceId]));
    const find = (sourceId: string): string => {
      const current = parent.get(sourceId);
      if (!current || current === sourceId) return sourceId;
      const root = find(current);
      parent.set(sourceId, root);
      return root;
    };
    const incomingLeft = new Set<string>();
    const relationshipIds = new Set<string>();
    for (const relationship of source.relationships) {
      const item = requireRecord(relationship, 'Pivot relationship');
      validateExactKeys(item, ['id', 'left', 'right', 'join'], 'Pivot relationship');
      if (!isNonEmptyString(item.id) || relationshipIds.has(String(item.id)) || (item.join !== 'inner' && item.join !== 'left')) throw new Error('Pivot relationship is invalid');
      relationshipIds.add(String(item.id));
      const left = requireRecord(item.left, 'Pivot relationship left');
      const right = requireRecord(item.right, 'Pivot relationship right');
      for (const side of [item.left, item.right]) {
        const reference = requireRecord(side, 'Pivot relationship field');
        validateExactKeys(reference, ['sourceId', 'fieldId'], 'Pivot relationship field');
        if (!isNonEmptyString(reference.sourceId) || !isNonEmptyString(reference.fieldId)) throw new Error('Pivot relationship field is invalid');
      }
      if (left.sourceId === right.sourceId || !sourceIds.has(String(left.sourceId)) || !sourceIds.has(String(right.sourceId))) throw new Error('Pivot relationship sourceId is invalid');
      if (item.join === 'left') incomingLeft.add(String(right.sourceId));
      const leftRoot = find(String(left.sourceId));
      const rightRoot = find(String(right.sourceId));
      if (leftRoot === rightRoot) throw new Error('Pivot relationship graph contains a cycle');
      parent.set(leftRoot, rightRoot);
    }
    if (sourceIds.size > 1 && !source.relationships.length) throw new Error('Pivot relationship graph is disconnected');
    const roots = source.relationships.some((entry) => (entry as Record<string, unknown>).join === 'left')
      ? [...sourceIds].filter((sourceId) => !incomingLeft.has(sourceId))
      : [[...sourceIds].sort()[0]!];
    if (roots.length !== 1 && sourceIds.size > 1) throw new Error('Pivot relationship graph has an ambiguous root');
    if (sourceIds.size > 1) {
      const graph = new Map([...sourceIds].map((sourceId) => [sourceId, [] as string[]]));
      for (const raw of source.relationships) {
        const item = raw as Record<string, unknown>;
        const left = item.left as Record<string, unknown>;
        const right = item.right as Record<string, unknown>;
        graph.get(String(left.sourceId))!.push(String(right.sourceId));
        graph.get(String(right.sourceId))!.push(String(left.sourceId));
      }
      const visited = new Set<string>();
      const visit = (sourceId: string): void => { if (visited.has(sourceId)) return; visited.add(sourceId); graph.get(sourceId)?.forEach(visit); };
      visit([...sourceIds][0]!);
      if (visited.size !== sourceIds.size) throw new Error('Pivot relationship graph is disconnected');
    }
    return;
  }
  if (source.kind === 'table') {
    validateExactKeys(source, ['kind', 'tableId'], 'Pivot table source');
    if (!isNonEmptyString(source.tableId)) throw new Error('Pivot table source is invalid');
    return;
  }
  if (source.kind === 'named-range') {
    validateExactKeys(source, ['kind', 'name', 'sheetId'], 'Pivot named source');
    if (!isNonEmptyString(source.name) || (source.sheetId !== undefined && !isNonEmptyString(source.sheetId))) throw new Error('Pivot named source is invalid');
    return;
  }
  if (source.kind === 'data-source') {
    validateExactKeys(source, ['kind', 'dataSourceId'], 'Pivot data source');
    if (!isNonEmptyString(source.dataSourceId)) throw new Error('Pivot data source is invalid');
    return;
  }
  throw new Error('Pivot source kind is unsupported');
}

function validatePivotPresentation(value: unknown): asserts value is PivotPresentation {
  const presentation = requireRecord(value, 'Pivot presentation');
  validateExactKeys(presentation, ['styleName', 'styleOptions', 'displayOptions'], 'Pivot presentation');
  if (presentation.styleName !== undefined && typeof presentation.styleName !== 'string') throw new Error('Pivot presentation styleName is invalid');
  const options = requireRecord(presentation.styleOptions, 'Pivot presentation styleOptions');
  validateExactKeys(options, ['showRowHeaders', 'showColumnHeaders', 'showRowStripes', 'showColumnStripes', 'showLastColumn'], 'Pivot presentation styleOptions');
  if (typeof options.showRowHeaders !== 'boolean'
    || typeof options.showColumnHeaders !== 'boolean'
    || typeof options.showRowStripes !== 'boolean'
    || typeof options.showColumnStripes !== 'boolean'
    || typeof options.showLastColumn !== 'boolean') throw new Error('Pivot presentation styleOptions are invalid');
  if (presentation.displayOptions !== undefined) {
    const display = requireRecord(presentation.displayOptions, 'Pivot presentation displayOptions');
    validateExactKeys(display, ['fillEmptyCells', 'emptyCellText', 'showErrorValues', 'errorCellText', 'showFieldHeaders', 'autoFitColumnsOnUpdate'], 'Pivot presentation displayOptions');
    if (typeof display.fillEmptyCells !== 'boolean' || typeof display.emptyCellText !== 'string'
      || typeof display.showErrorValues !== 'boolean' || typeof display.errorCellText !== 'string'
      || typeof display.showFieldHeaders !== 'boolean' || typeof display.autoFitColumnsOnUpdate !== 'boolean') {
      throw new Error('Pivot presentation displayOptions are invalid');
    }
  }
}

function validatePivotCollation(value: unknown): void {
  const collation = requireRecord(value, 'Pivot collation');
  validateExactKeys(collation, ['locale', 'sensitivity', 'numeric', 'caseFirst'], 'Pivot collation');
  if (!isNonEmptyString(collation.locale)
    || !['base', 'accent', 'case', 'variant'].includes(String(collation.sensitivity))
    || typeof collation.numeric !== 'boolean'
    || !['upper', 'lower', 'false'].includes(String(collation.caseFirst))) {
    throw new Error('Pivot collation is invalid');
  }
  createPivotCollator({
    locale: String(collation.locale),
    sensitivity: collation.sensitivity as 'base' | 'accent' | 'case' | 'variant',
    numeric: collation.numeric,
    caseFirst: collation.caseFirst as 'upper' | 'lower' | 'false',
  });
}

function validatePivotGroup(value: unknown): void {
  const group = requireRecord(value, 'Pivot group');
  if (group.kind === 'date') {
    validateExactKeys(group, ['kind', 'unit', 'units', 'startOfWeek', 'start', 'end', 'autoStart', 'autoEnd'], 'Pivot date group');
    const units = group.units === undefined ? [group.unit] : group.units;
    if (!['year', 'quarter', 'month', 'week', 'day'].includes(String(group.unit))
      || !Array.isArray(units) || units.length === 0 || new Set(units).size !== units.length || !units.every((unit) => ['year', 'quarter', 'month', 'week', 'day'].includes(String(unit)))
      || !units.includes(String(group.unit) as typeof group.unit)
      || (group.startOfWeek !== undefined && (!Number.isInteger(group.startOfWeek) || Number(group.startOfWeek) < 0 || Number(group.startOfWeek) > 6))
      || (group.start !== undefined && !['string', 'number'].includes(typeof group.start))
      || (group.end !== undefined && !['string', 'number'].includes(typeof group.end))
      || (group.autoStart !== undefined && typeof group.autoStart !== 'boolean')
      || (group.autoEnd !== undefined && typeof group.autoEnd !== 'boolean')) throw new Error('Pivot date group is invalid');
    return;
  }
  if (group.kind === 'number') {
    validateExactKeys(group, ['kind', 'interval', 'start', 'end', 'autoStart', 'autoEnd'], 'Pivot number group');
    if (typeof group.interval !== 'number' || !Number.isFinite(group.interval) || group.interval <= 0
      || (group.start !== undefined && (typeof group.start !== 'number' || !Number.isFinite(group.start)))
      || (group.end !== undefined && (typeof group.end !== 'number' || !Number.isFinite(group.end)))
      || (group.autoStart !== undefined && typeof group.autoStart !== 'boolean')
      || (group.autoEnd !== undefined && typeof group.autoEnd !== 'boolean')) throw new Error('Pivot number group is invalid');
    return;
  }
  if (group.kind === 'manual') {
    validateExactKeys(group, ['kind', 'groups'], 'Pivot manual group');
    if (!Array.isArray(group.groups)) throw new Error('Pivot manual group is invalid');
    for (const raw of group.groups) {
      const entry = requireRecord(raw, 'Pivot manual group entry');
      validateExactKeys(entry, ['groupId', 'name', 'items'], 'Pivot manual group entry');
      if (!isNonEmptyString(entry.groupId) || typeof entry.name !== 'string' || !Array.isArray(entry.items)) throw new Error('Pivot manual group entry is invalid');
      entry.items.forEach((item, index) => validatePivotMemberKey(item, `Pivot manual group item ${String(index)}`));
    }
    return;
  }
  throw new Error('Pivot group kind is unsupported');
}

function validatePivotSubtotal(value: unknown): void {
  const subtotal = requireRecord(value, 'Pivot subtotal');
  if (subtotal.mode === 'automatic' || subtotal.mode === 'none') {
    validateExactKeys(subtotal, ['mode'], 'Pivot subtotal');
    return;
  }
  if (subtotal.mode === 'custom') {
    validateExactKeys(subtotal, ['mode', 'functions'], 'Pivot custom subtotal');
    const functions = subtotal.functions;
    const supported = ['sum', 'count', 'count-numbers', 'average', 'min', 'max', 'product', 'stdev', 'stdevp', 'var', 'varp', 'distinct-count'];
    if (!Array.isArray(functions) || functions.length === 0 || new Set(functions).size !== functions.length || !functions.every((item) => supported.includes(String(item)))) throw new Error('Pivot custom subtotal functions are invalid');
    return;
  }
  throw new Error('Pivot subtotal mode is unsupported');
}

/**
 * Calculated fields and items are declared by the layout mutation itself.
 * They are not required to be copied into the persisted source field
 * catalogue.  Build the effective field set once so a full definition can
 * validate both the definitions and every layout reference against the same
 * ownership boundary.  A layout-only mutation can validate its own shape and
 * duplicate IDs, while the server validates its targets against the current
 * source catalogue when the patch is merged.
 */
function validatePivotCalculatedItemReferences(
  items: readonly Record<string, unknown>[],
  fields: readonly Record<string, unknown>[],
): void {
  const functions = new Set(['SUM', 'COUNT', 'AVERAGE', 'MIN', 'MAX', 'IF', 'AND', 'OR', 'NOT', 'ROUND', 'ABS', 'CONCAT', 'LEFT', 'RIGHT', 'LEN']);
  const membersByField = new Map<string, unknown[]>();
  const fieldByReference = (reference: string): Record<string, unknown> | undefined => fields.find((field) => field.fieldId === reference || field.name === reference);
  const memberDomainAvailable = fields.every((field) => Array.isArray(field.values));
  for (const field of fields) membersByField.set(String(field.fieldId), Array.isArray(field.values) ? [...field.values] : []);
  for (const item of items) {
    const members = membersByField.get(String(item.targetFieldId));
    if (!members) throw new Error(`Pivot calculated item target field is invalid: ${String(item.targetFieldId)}`);
    if (members.some((member) => String(member) === String(item.name))) throw new Error(`Pivot calculated item member collides with source data: ${String(item.name)}`);
    members.push(item.name);
  }
  const itemByMember = new Map<string, string>();
  for (const item of items) itemByMember.set(`${String(item.targetFieldId)}|${String(item.name).toLocaleUpperCase()}`, String(item.fieldId));
  if (!memberDomainAvailable) {
    const dependencies = new Map<string, string[]>();
    for (const item of items) {
      const refs: string[] = [];
      const targetPrefix = `${String(item.targetFieldId)}|`;
      const token = /[A-Za-z_][A-Za-z0-9_.:-]*/g;
      for (const match of String(item.formula).matchAll(token)) {
        const dependency = itemByMember.get(`${targetPrefix}${match[0]!.toLocaleUpperCase()}`);
        if (dependency && !refs.includes(dependency)) refs.push(dependency);
      }
      dependencies.set(String(item.fieldId), refs);
    }
    const state = new Map<string, 'visiting' | 'visited'>();
    const visit = (id: string, path: string[]): void => {
      if (state.get(id) === 'visited') return;
      if (state.get(id) === 'visiting') throw new Error(`Pivot calculated item dependency cycle: ${[...path, id].join(' -> ')}`);
      state.set(id, 'visiting');
      for (const dependency of dependencies.get(id) ?? []) visit(dependency, [...path, id]);
      state.set(id, 'visited');
    };
    for (const item of items) visit(String(item.fieldId), []);
    return;
  }
  const dependencies = new Map<string, string[]>();
  for (const item of items) {
    // Keep the protocol boundary aligned with the runtime: worksheet
    // functions, cell references, and arbitrary operators are not Pivot-item
    // syntax.  Member resolution below still supplies the typed identity.
    parsePivotCalculatedItemFormula(String(item.formula));
    const formula = String(item.formula).trim().replace(/^=/, '');
    const occupied: Array<{ start: number; end: number }> = [];
    const references: string[] = [];
    const add = (fieldId: string, member: string): void => {
      const key = `${fieldId}|${member.toLocaleUpperCase()}`;
      const itemId = itemByMember.get(key);
      if (itemId && !references.includes(itemId)) references.push(itemId);
    };
    const resolveMember = (field: Record<string, unknown>, text: string): void => {
      const candidates = (membersByField.get(String(field.fieldId)) ?? []).filter((member) => String(member).toLocaleUpperCase() === text.toLocaleUpperCase());
      if (candidates.length !== 1) throw new Error(candidates.length ? `Pivot calculated item reference is ambiguous: ${text}` : `Pivot calculated item references unknown item: ${text}`);
      add(String(field.fieldId), String(candidates[0]));
    };
    const qualified = /([A-Za-z_][A-Za-z0-9_.:-]*)\s*\[([^\]]+)\]/g;
    for (let match = qualified.exec(formula); match; match = qualified.exec(formula)) {
      const field = fieldByReference(match[1]!);
      if (!field) throw new Error(`Pivot calculated item references unknown field: ${match[1]}`);
      resolveMember(field, match[2]!);
      occupied.push({ start: match.index, end: match.index + match[0].length });
    }
    const bare = /[A-Za-z_][A-Za-z0-9_.:-]*/g;
    for (let match = bare.exec(formula); match; match = bare.exec(formula)) {
      if (occupied.some((range) => match!.index >= range.start && match!.index < range.end)) continue;
      const token = match[0]!;
      const after = formula.slice(match.index + token.length).trimStart();
      if (after.startsWith('(')) {
        if (!functions.has(token.toUpperCase())) throw new Error(`Pivot calculated item function is unsupported: ${token}`);
        continue;
      }
      if (functions.has(token.toUpperCase()) || ['TRUE', 'FALSE'].includes(token.toUpperCase())) continue;
      if (/^[A-Z]{1,3}[1-9][0-9]*$/i.test(token)) throw new Error(`Pivot calculated item worksheet reference is unsupported: ${token}`);
      const candidates = fields.flatMap((field) => (membersByField.get(String(field.fieldId)) ?? []).filter((member) => typeof member === 'string' && String(member).toLocaleUpperCase() === token.toLocaleUpperCase()).map(() => field));
      if (candidates.length !== 1) throw new Error(candidates.length ? `Pivot calculated item reference is ambiguous: ${token}` : `Pivot calculated item references unknown item: ${token}`);
      add(String(candidates[0]!.fieldId), token);
    }
    dependencies.set(String(item.fieldId), references);
  }
  const state = new Map<string, 'visiting' | 'visited'>();
  const visit = (id: string, path: string[]): void => {
    if (state.get(id) === 'visited') return;
    if (state.get(id) === 'visiting') throw new Error(`Pivot calculated item dependency cycle: ${[...path, id].join(' -> ')}`);
    state.set(id, 'visiting');
    for (const dependency of dependencies.get(id) ?? []) visit(dependency, [...path, id]);
    state.set(id, 'visited');
  };
  for (const item of items) visit(String(item.fieldId), []);
}

function validatePivotCalculatedDefinitions(
  layout: Record<string, unknown>,
  fieldIds?: ReadonlySet<string>,
  catalogFields?: readonly Record<string, unknown>[],
): Set<string> {
  const effectiveFieldIds = new Set(fieldIds);
  const calculatedItems: Record<string, unknown>[] = [];
  const calculatedFieldIds = new Set<string>();
  const definitionIds = new Set(effectiveFieldIds);
  const validate = (raw: unknown, item: boolean): void => {
    if (raw === undefined || raw === null) return;
    if (!Array.isArray(raw) || raw.length > 10_000) throw new Error('Pivot calculated definitions are invalid');
    for (const rawDefinition of raw) {
      const definition = requireRecord(rawDefinition, item ? 'Pivot calculated item' : 'Pivot calculated field');
      validateExactKeys(
        definition,
        item ? ['fieldId', 'targetFieldId', 'name', 'formula'] : ['fieldId', 'name', 'formula'],
        item ? 'Pivot calculated item' : 'Pivot calculated field',
      );
      if (!isNonEmptyString(definition.fieldId) || !isNonEmptyString(definition.name) || !isNonEmptyString(definition.formula)) {
        throw new Error(item ? 'Pivot calculated item is invalid' : 'Pivot calculated field is invalid');
      }
      if (item) {
        // A calculated item owns a derived member identity, not a source
        // field.  It must never enter the effective field universe used by
        // rows, columns, filters, or Values.  Keep its identity in a
        // separate namespace so it cannot collide with a source/calculated
        // field or another item.
        if (definitionIds.has(definition.fieldId)) throw new Error(`Pivot calculated itemId is duplicated or collides with the field catalogue: ${definition.fieldId}`);
        definitionIds.add(definition.fieldId);
        if (!isNonEmptyString(definition.targetFieldId)) throw new Error('Pivot calculated item targetFieldId is invalid');
        parsePivotCalculatedItemFormula(String(definition.formula));
        calculatedItems.push(definition);
      } else {
        if (definitionIds.has(definition.fieldId)) throw new Error(`Pivot calculated fieldId is duplicated or collides with the field catalogue: ${definition.fieldId}`);
        definitionIds.add(definition.fieldId);
        effectiveFieldIds.add(definition.fieldId);
        calculatedFieldIds.add(String(definition.fieldId));
      }
    }
  };
  validate(layout.calculatedFields, false);
  validate(layout.calculatedItems, true);
  if (fieldIds) {
    for (const item of calculatedItems) {
      if (!effectiveFieldIds.has(String(item.targetFieldId))) throw new Error(`Pivot calculated item targetFieldId is invalid: ${String(item.targetFieldId)}`);
      if (calculatedFieldIds.has(String(item.targetFieldId))) throw new Error(`Pivot calculated item target field cannot be a calculated field: ${String(item.targetFieldId)}`);
    }
  }
  if (catalogFields) validatePivotCalculatedItemReferences(calculatedItems, catalogFields);
  return effectiveFieldIds;
}

/** Rejects every non-canonical Pivot field at the transport boundary. */
export function validatePivotDefinition(value: unknown): asserts value is PivotDefinition {
  const pivot = requireRecord(value, 'Pivot definition');
  validateExactKeys(pivot, ['schema', 'id', 'source', 'target', 'fieldCatalog', 'layout', 'refreshPolicy', 'nativeMetadata', 'presentation'], 'Pivot definition');
  if (pivot.schema !== 'PivotDefinition' || !isNonEmptyString(pivot.id)) throw new Error('Pivot definition identity is invalid');
  validatePivotSource(pivot.source);
  if (pivot.presentation !== undefined) validatePivotPresentation(pivot.presentation);
  const target = requireRecord(pivot.target, 'Pivot target');
  validateExactKeys(target, ['sheetId', 'anchor'], 'Pivot target');
  const anchor = requireRecord(target.anchor, 'Pivot target anchor');
  validateExactKeys(anchor, ['row', 'column'], 'Pivot target anchor');
  if (!isNonEmptyString(target.sheetId) || !Number.isSafeInteger(anchor.row) || Number(anchor.row) < 0 || !Number.isSafeInteger(anchor.column) || Number(anchor.column) < 0) throw new Error('Pivot target is invalid');
  const catalog = requireRecord(pivot.fieldCatalog, 'Pivot field catalog');
  validateExactKeys(catalog, ['schema', 'fields'], 'Pivot field catalog');
  if (catalog.schema !== undefined && catalog.schema !== 'PivotFieldCatalog') throw new Error('Pivot field catalog schema is invalid');
  if (!Array.isArray(catalog.fields)) throw new Error('Pivot field catalog fields are invalid');
  const fieldIds = new Set<string>();
  for (const [ordinal, rawField] of catalog.fields.entries()) {
    const field = requireRecord(rawField, 'Pivot field');
    validateExactKeys(field, ['fieldId', 'name', 'dataType', 'ordinal', 'values'], 'Pivot field');
    if (!isNonEmptyString(field.fieldId) || fieldIds.has(field.fieldId) || !isNonEmptyString(field.name)
      || field.ordinal !== ordinal || !['text', 'number', 'date', 'boolean', 'error', 'mixed'].includes(String(field.dataType))) {
      throw new Error('Pivot field is invalid');
    }
    if (field.values !== undefined && (!Array.isArray(field.values) || field.values.length > PIVOT_MAX_MEMBER_COUNT)) throw new Error('Pivot field values are invalid');
    field.values?.forEach((item, index) => validatePivotScalar(item, `Pivot field value ${String(index)}`));
    fieldIds.add(field.fieldId);
  }
  const layout = requireRecord(pivot.layout, 'Pivot layout');
  validateExactKeys(layout, ['rows', 'columns', 'filters', 'allowMultipleFiltersPerField', 'collation', 'values', 'calculatedFields', 'calculatedItems', 'subtotalLocation', 'showRowGrandTotals', 'showColumnGrandTotals', 'reportLayout', 'expansion'], 'Pivot layout');
  if (!Array.isArray(layout.rows) || !Array.isArray(layout.columns) || !Array.isArray(layout.filters) || !Array.isArray(layout.values)
    || typeof layout.allowMultipleFiltersPerField !== 'boolean'
    || !['top', 'bottom', 'off'].includes(String(layout.subtotalLocation)) || typeof layout.showRowGrandTotals !== 'boolean' || typeof layout.showColumnGrandTotals !== 'boolean' || !['compact', 'outline', 'tabular'].includes(String(layout.reportLayout))) throw new Error('Pivot layout is invalid');
  validatePivotCollation(layout.collation);
  const effectiveFieldIds = validatePivotCalculatedDefinitions(layout, fieldIds, catalog.fields as unknown as readonly Record<string, unknown>[]);
  if (layout.expansion !== undefined) {
    const expansion = requireRecord(layout.expansion, 'Pivot expansion');
    validateExactKeys(expansion, ['expandedNodeIds', 'collapsedNodeIds', 'showButtons'], 'Pivot expansion');
    if (!Array.isArray(expansion.expandedNodeIds) || !expansion.expandedNodeIds.every(isNonEmptyString)
      || !Array.isArray(expansion.collapsedNodeIds) || !expansion.collapsedNodeIds.every(isNonEmptyString)
      || typeof expansion.showButtons !== 'boolean') throw new Error('Pivot expansion is invalid');
  }
  const validatePlacement = (rawPlacement: unknown): void => {
    const placement = requireRecord(rawPlacement, 'Pivot placement');
    validateExactKeys(placement, ['fieldId', 'sort', 'group', 'subtotal'], 'Pivot placement');
    if (!isNonEmptyString(placement.fieldId) || !effectiveFieldIds.has(placement.fieldId)) throw new Error('Pivot placement fieldId is invalid');
    if (placement.sort !== undefined) {
      const sort = requireRecord(placement.sort, 'Pivot sort');
      validateExactKeys(sort, ['direction', 'by', 'valueId'], 'Pivot sort');
      if (!['ascending', 'descending'].includes(String(sort.direction)) || !['label', 'value'].includes(String(sort.by))) throw new Error('Pivot sort is invalid');
      if (sort.by === 'value' && !isNonEmptyString(sort.valueId)) throw new Error('Pivot value sort requires valueId');
      if (sort.by === 'label' && sort.valueId !== undefined) throw new Error('Pivot label sort cannot carry valueId');
    }
    if (placement.group !== undefined) validatePivotGroup(placement.group);
    if (placement.subtotal !== undefined) validatePivotSubtotal(placement.subtotal);
  };
  layout.rows.forEach(validatePlacement);
  layout.columns.forEach(validatePlacement);
  const axisFieldIds = new Set<string>([...layout.rows, ...layout.columns].map((placement) => String((placement as Record<string, unknown>).fieldId)));
  const filterIdentities = new Set<string>();
  const filterFields = new Set<string>();
  for (const rawFilter of layout.filters) {
    const filter = requireRecord(rawFilter, 'Pivot filter');
    if (!isNonEmptyString(filter.fieldId) || !effectiveFieldIds.has(filter.fieldId)) throw new Error('Pivot filter fieldId is invalid');
    if (filter.kind === 'manual') {
      validateExactKeys(filter, ['kind', 'family', 'fieldId', 'scope', 'mode', 'memberKeys'], 'Pivot manual filter');
      if (filter.family !== 'manual') throw new Error('Pivot manual filter family is invalid');
      if (filter.scope !== undefined && !['report', 'field'].includes(String(filter.scope))) throw new Error('Pivot manual filter scope is invalid');
      if (!['all', 'include', 'exclude'].includes(String(filter.mode)) || !Array.isArray(filter.memberKeys)) throw new Error('Pivot manual filter is invalid');
      filter.memberKeys.forEach((item, index) => validatePivotMemberKey(item, `Pivot manual filter member ${String(index)}`));
    } else if (filter.kind === 'condition') {
      validateExactKeys(filter, ['kind', 'family', 'fieldId', 'valueId', 'scope', 'operator', 'value', 'value2', 'dynamic', 'wholeDay'], 'Pivot condition filter');
      if (!['label', 'date', 'value'].includes(String(filter.family))) throw new Error('Pivot condition filter family is invalid');
      if (filter.scope !== undefined && !['report', 'field'].includes(String(filter.scope))) throw new Error('Pivot condition filter scope is invalid');
      if (filter.valueId !== undefined && !isNonEmptyString(filter.valueId)) throw new Error('Pivot condition valueId is invalid');
      const labelOperators = ['equals', 'not-equals', 'begins-with', 'not-begins-with', 'ends-with', 'not-ends-with', 'contains', 'not-contains', 'between', 'not-between', 'greater-than', 'greater-or-equal', 'less-than', 'less-or-equal'];
      const dateOperators = ['equals', 'not-equals', 'before', 'after', 'between', 'not-between'];
      const valueOperators = ['equals', 'not-equals', 'greater-than', 'greater-or-equal', 'less-than', 'less-or-equal', 'between', 'not-between'];
      const operators = filter.family === 'label' ? labelOperators : filter.family === 'date' ? dateOperators : valueOperators;
      if (!operators.includes(String(filter.operator))) throw new Error('Pivot condition operator is invalid for its family');
      if (filter.family === 'value' && !isNonEmptyString(filter.valueId)) throw new Error('Pivot value filter requires valueId');
      if (filter.family !== 'value' && filter.valueId !== undefined) throw new Error('Pivot condition valueId is only valid for value filters');
      validatePivotScalar(filter.value, 'Pivot condition value');
      if (filter.value2 !== undefined) validatePivotScalar(filter.value2, 'Pivot condition upper value');
      const dynamicDates = ['today', 'yesterday', 'tomorrow', 'this-week', 'last-week', 'next-week', 'this-month', 'last-month', 'next-month', 'this-quarter', 'last-quarter', 'next-quarter', 'this-year', 'last-year', 'next-year', 'year-to-date'];
      if (filter.dynamic !== undefined && (filter.family !== 'date' || !dynamicDates.includes(String(filter.dynamic)))) throw new Error('Pivot dynamic date filter is invalid');
      if ((filter.operator === 'between' || filter.operator === 'not-between') && filter.value2 === undefined && filter.dynamic === undefined) throw new Error('Pivot range filter requires two bounds');
      if (filter.family === 'date' && filter.dynamic !== undefined && !['equals', 'between'].includes(String(filter.operator))) throw new Error('Pivot dynamic date operator is invalid');
      if (filter.wholeDay !== undefined && (filter.family !== 'date' || typeof filter.wholeDay !== 'boolean')) throw new Error('Pivot condition wholeDay is invalid');
    } else if (filter.kind === 'top-items') {
      validateExactKeys(filter, ['kind', 'family', 'fieldId', 'scope', 'valueId', 'direction', 'mode', 'threshold'], 'Pivot top-items filter');
      if (filter.family !== 'top-items') throw new Error('Pivot top-items filter family is invalid');
      if (filter.scope !== undefined && !['report', 'field'].includes(String(filter.scope))) throw new Error('Pivot top-items filter scope is invalid');
      if (!isNonEmptyString(filter.valueId) || !['top', 'bottom'].includes(String(filter.direction))
        || !['items', 'percent', 'sum'].includes(String(filter.mode))
        || typeof filter.threshold !== 'number' || !Number.isFinite(filter.threshold) || filter.threshold <= 0
        || (filter.mode === 'items' && (!Number.isSafeInteger(filter.threshold) || filter.threshold < 1))
        || (filter.mode === 'percent' && filter.threshold > 100)) throw new Error('Pivot top-items filter is invalid');
    } else throw new Error('Pivot filter kind is unsupported');
    const scope = filter.scope ?? (axisFieldIds.has(filter.fieldId) ? 'field' : 'report');
    if (scope === 'field' && !axisFieldIds.has(filter.fieldId)) throw new Error('Pivot field filter must target a row or column field');
    const identity = `${filter.fieldId}|${scope}|${String(filter.family)}`;
    if (filterIdentities.has(identity)) throw new Error('Pivot filter family is duplicated');
    filterIdentities.add(identity);
    const fieldScope = `${filter.fieldId}|${scope}`;
    if (!layout.allowMultipleFiltersPerField && filterFields.has(fieldScope)) throw new Error('Pivot multiple filters per field are disabled');
    filterFields.add(fieldScope);
  }
  const valueIds = new Set<string>();
  for (const rawValue of layout.values) {
    const item = requireRecord(rawValue, 'Pivot value field');
    validateExactKeys(item, ['valueId', 'fieldId', 'summarizeBy', 'displayName', 'numberFormat', 'showAs'], 'Pivot value field');
    if (!isNonEmptyString(item.valueId) || valueIds.has(item.valueId)) throw new Error('Pivot value placement identity is invalid or duplicated');
    valueIds.add(item.valueId);
    if (!isNonEmptyString(item.fieldId) || !effectiveFieldIds.has(item.fieldId) || !['sum', 'count', 'count-numbers', 'average', 'min', 'max', 'product', 'stdev', 'stdevp', 'var', 'varp', 'distinct-count'].includes(String(item.summarizeBy))) throw new Error('Pivot value field is invalid');
    validatePivotShowAs(item.showAs, effectiveFieldIds, axisFieldIds);
  }
  for (const rawPlacement of [...layout.rows, ...layout.columns]) {
    const placement = rawPlacement as Record<string, unknown>;
    const sort = placement.sort as Record<string, unknown> | undefined;
    if (sort?.by === 'value' && (!isNonEmptyString(sort.valueId) || !valueIds.has(sort.valueId))) throw new Error('Pivot value sort placement identity is invalid');
  }
  for (const rawFilter of layout.filters) {
    const filter = rawFilter as Record<string, unknown>;
    if (filter.kind === 'top-items' && (!isNonEmptyString(filter.valueId) || !valueIds.has(filter.valueId))) throw new Error('Pivot top-items placement identity is invalid');
    if (filter.kind === 'condition' && filter.valueId !== undefined && (!valueIds.has(String(filter.valueId)) || filter.family !== 'value')) throw new Error('Pivot condition placement identity is invalid');
  }
  const policy = requireRecord(pivot.refreshPolicy, 'Pivot refresh policy');
  validateExactKeys(policy, ['mode', 'preserveFormatting', 'refreshOnLoad'], 'Pivot refresh policy');
  try {
    normalizePivotRefreshPolicy(policy as unknown as PivotDefinition['refreshPolicy']);
  } catch (error) {
    throw new Error(error instanceof Error ? error.message : 'Pivot refresh policy is invalid');
  }
}

export function validateDataSourceManifest(value: unknown): asserts value is DataSourceManifest {
  if (!value || typeof value !== 'object' || Array.isArray(value)) throw new Error('Data source manifest must be an object');
  const source = value as Record<string, unknown>;
  validateExactKeys(source, ['schema', 'version', 'id', 'name', 'kind', 'sourceSheetId', 'sourceRange', 'rowCount', 'fields', 'blockRowCount', 'blocks', 'revision'], 'Data source manifest');
  if (source.schema !== 'DataSourceManifest' || source.version !== 1 || !isNonEmptyString(source.id) || !isNonEmptyString(source.name)) {
    throw new Error('Invalid data source manifest identity');
  }
  if (!['worksheet-range', 'sheet-table', 'chunked-table'].includes(String(source.kind))
    || !Number.isSafeInteger(source.rowCount) || Number(source.rowCount) < 0
    || !Number.isSafeInteger(source.blockRowCount) || Number(source.blockRowCount) <= 0
    || !Number.isSafeInteger(source.revision) || Number(source.revision) < 0
    || !Array.isArray(source.fields) || !Array.isArray(source.blocks)) {
    throw new Error(`Invalid data source manifest shape: ${String(source.id)}`);
  }
  if (source.sourceSheetId !== undefined && !isNonEmptyString(source.sourceSheetId)) {
    throw new Error(`Invalid data source sourceSheetId: ${String(source.id)}`);
  }
  if (source.sourceRange !== undefined && !isRangeRef(source.sourceRange)) throw new Error(`Invalid data source range: ${String(source.id)}`);
  if (source.sourceRange !== undefined && source.sourceSheetId !== source.sourceRange.sheetId) {
    throw new Error(`Data source sourceRange must target sourceSheetId: ${String(source.id)}`);
  }
  if ((source.kind === 'worksheet-range' || source.kind === 'sheet-table')
    && (!isNonEmptyString(source.sourceSheetId) || source.sourceRange === undefined)) {
    throw new Error(`Invalid ${String(source.kind)} data source metadata: ${String(source.id)}`);
  }
  const fieldIds = new Set<string>();
  for (const [index, rawField] of source.fields.entries()) {
    if (!rawField || typeof rawField !== 'object') throw new Error(`Invalid data source field: ${String(source.id)}`);
    const field = rawField as Record<string, unknown>;
    validateExactKeys(field, ['id', 'name', 'ordinal', 'type'], 'Data source field');
    if (!isNonEmptyString(field.id) || fieldIds.has(field.id) || !isNonEmptyString(field.name) || field.ordinal !== index
      || !['text', 'number', 'boolean', 'date', 'mixed'].includes(String(field.type))) {
      throw new Error(`Invalid data source field: ${String(source.id)}`);
    }
    fieldIds.add(field.id);
  }
  const blockIds = new Set<string>();
  for (const rawBlock of source.blocks) {
    if (!rawBlock || typeof rawBlock !== 'object') throw new Error(`Invalid data block: ${String(source.id)}`);
    const block = rawBlock as Record<string, unknown>;
    validateExactKeys(block, ['id', 'dataSourceId', 'startRow', 'rowCount', 'storageKey', 'checksum', 'byteLength', 'encoding', 'revision'], 'Data block');
    if (!isNonEmptyString(block.id) || blockIds.has(block.id) || block.dataSourceId !== source.id
      || !isNonEmptyString(block.storageKey) || !isNonEmptyString(block.checksum)
      || block.encoding !== 'columnar-v1'
      || !Number.isSafeInteger(block.startRow) || Number(block.startRow) < 0
      || !Number.isSafeInteger(block.rowCount) || Number(block.rowCount) <= 0
      || !Number.isSafeInteger(block.byteLength) || Number(block.byteLength) <= 0
      || !Number.isSafeInteger(block.revision) || Number(block.revision) < 0) {
      throw new Error(`Invalid data block: ${String(source.id)}`);
    }
    if (Number(block.startRow) + Number(block.rowCount) > Number(source.rowCount)) {
      throw new Error(`Data block exceeds source rowCount: ${String(block.id)}`);
    }
    if (!/^[A-Fa-f0-9]{64}$/.test(String(block.checksum))) {
      throw new Error(`Invalid data block checksum: ${String(block.id)}`);
    }
    blockIds.add(block.id);
  }
}

export function isDataSourceMutationId(value: string): value is DataSourceMutationId {
  return (DATA_SOURCE_MUTATION_IDS as readonly string[]).includes(value);
}

/** Validate the exact wire shape for data-source mutations. */
export function validateDataSourceMutationParams(
  id: DataSourceMutationId,
  value: unknown,
): asserts value is DataSourceMutationParams {
  assertMetadataOnly(value, `Mutation ${id}.params`);
  if (!value || typeof value !== 'object' || Array.isArray(value)) throw new Error(`Mutation ${id} params must be an object`);
  const params = value as Record<string, unknown>;
  switch (id) {
    case 'dataSource.add':
    case 'dataSource.update':
      validateExactKeys(params, ['source'], `Mutation ${id}`);
      validateDataSourceManifest(params.source);
      return;
    case 'dataSource.remove':
      validateExactKeys(params, ['sourceId'], `Mutation ${id}`);
      if (!isNonEmptyString(params.sourceId)) throw new Error(`Mutation ${id} sourceId is required`);
      return;
    case 'dataRegion.add':
      validateExactKeys(params, ['region'], `Mutation ${id}`);
      if (!params.region || typeof params.region !== 'object' || Array.isArray(params.region)) {
        throw new Error(`Mutation ${id} region is required`);
      }
      const region = params.region as Record<string, unknown>;
      validateExactKeys(region, ['id', 'sourceId', 'range', 'headerRow', 'revision'], 'Sheet data region');
      if (!isNonEmptyString(region.id) || !isNonEmptyString(region.sourceId) || !isRangeRef(region.range)
        || !Number.isSafeInteger(region.headerRow) || Number(region.headerRow) < 0
        || !Number.isSafeInteger(region.revision) || Number(region.revision) < 0) {
        throw new Error(`Mutation ${id} region metadata is invalid`);
      }
      return;
    case 'dataRegion.remove':
      validateExactKeys(params, ['regionId'], `Mutation ${id}`);
      if (!isNonEmptyString(params.regionId)) throw new Error(`Mutation ${id} regionId is required`);
      return;
  }
}

function validatePivotMutationParams(id: string, value: unknown): void {
  if (id === 'pivot.add') {
    validatePivotDefinition(value);
    return;
  }
  if (id !== 'pivot.update') return;
  const params = requireRecord(value, 'Pivot update');
  validateExactKeys(params, ['sheetId', 'pivotId', 'source', 'target', 'fieldCatalog', 'refreshPolicy', 'nativeMetadata', 'presentation', 'layout'], 'Pivot update');
  if (!isNonEmptyString(params.sheetId) || !isNonEmptyString(params.pivotId)) throw new Error('Pivot update identity is invalid');
  if (!['source', 'target', 'fieldCatalog', 'refreshPolicy', 'nativeMetadata', 'presentation', 'layout'].some((key) => params[key] !== undefined)) throw new Error('Pivot update has no canonical change');
  if (params.source !== undefined) validatePivotSource(params.source);
  if (params.target !== undefined) {
    const target = requireRecord(params.target, 'Pivot update target');
    validateExactKeys(target, ['sheetId', 'anchor'], 'Pivot update target');
    const anchor = requireRecord(target.anchor, 'Pivot update target anchor');
    if (!isNonEmptyString(target.sheetId) || !Number.isSafeInteger(anchor.row) || Number(anchor.row) < 0 || !Number.isSafeInteger(anchor.column) || Number(anchor.column) < 0) throw new Error('Pivot update target is invalid');
  }
  if (params.presentation !== undefined) validatePivotPresentation(params.presentation);
  if (params.layout !== undefined) validatePivotCalculatedDefinitions(requireRecord(params.layout, 'Pivot update layout'));
}

/** Validate the only snapshot representation accepted at the wire boundary. */
export function validateWorkbookSnapshot(value: unknown): WorkbookSnapshot {
  if (!value || typeof value !== 'object' || Array.isArray(value)) {
    throw new Error('WorkbookSnapshot must be an object');
  }
  const input = value as Record<string, unknown>;
  if (input.schema !== WORKBOOK_SNAPSHOT_SCHEMA) throw new Error('Unsupported workbook snapshot schema');
  if (input.version !== WORKBOOK_SNAPSHOT_VERSION) throw new Error('Unsupported workbook snapshot version');
  if (!isNonEmptyString(input.unitId) || !isNonEmptyString(input.name)) {
    throw new Error('WorkbookSnapshot requires unitId and name');
  }
  if (input.name.length > MAX_WORKBOOK_NAME_LENGTH) throw new Error('WorkbookSnapshot name is too long');
  const dimensionMetrics = input.dimensionMetrics as Record<string, unknown> | undefined;
  if (!dimensionMetrics || !isNonEmptyString(dimensionMetrics.normalFontFamily)
    || typeof dimensionMetrics.normalFontSizePx !== 'number' || !Number.isFinite(dimensionMetrics.normalFontSizePx) || dimensionMetrics.normalFontSizePx <= 0
    || typeof dimensionMetrics.maximumDigitWidthPx !== 'number' || !Number.isFinite(dimensionMetrics.maximumDigitWidthPx) || dimensionMetrics.maximumDigitWidthPx <= 0) throw new Error('WorkbookSnapshot dimensionMetrics is invalid');
  if (!isWorkbookCalculationSettings(input.calculationSettings)) throw new Error('WorkbookSnapshot calculationSettings is invalid');
  if (input.theme !== undefined) {
    const theme = input.theme as Record<string, unknown>;
    if (!theme || typeof theme !== 'object' || Array.isArray(theme) || !isNonEmptyString(theme.id)
      || !theme.colors || typeof theme.colors !== 'object' || Array.isArray(theme.colors)
      || Object.entries(theme.colors as Record<string, unknown>).some(([key, color]) => !key.trim() || typeof color !== 'string' || !/^#[0-9a-f]{6}$/i.test(color))) {
      throw new Error('WorkbookSnapshot theme is invalid');
    }
  }
  if (!Array.isArray(input.sheets) || input.sheets.length === 0) {
    throw new Error('WorkbookSnapshot requires at least one sheet');
  }
  const dataModel = input.dataModel as Record<string, unknown> | undefined;
  if (!dataModel || !Array.isArray(dataModel.sources) || !Array.isArray(dataModel.tables)
    || !Array.isArray(dataModel.relationships) || !Array.isArray(dataModel.views)) {
    throw new Error('WorkbookSnapshot dataModel is invalid');
  }
  if (input.cellStyleTemplates !== undefined) {
    if (!Array.isArray(input.cellStyleTemplates)) throw new Error('WorkbookSnapshot cellStyleTemplates must be an array');
    const templateIds = new Set<string>();
    for (const template of input.cellStyleTemplates) {
      if (!template || typeof template !== 'object' || Array.isArray(template)) throw new Error('WorkbookSnapshot cell style template is invalid');
      const value = template as Record<string, unknown>;
      if (!isNonEmptyString(value.id) || !isNonEmptyString(value.name) || !value.style || typeof value.style !== 'object' || Array.isArray(value.style) || templateIds.has(value.id)) {
        throw new Error('WorkbookSnapshot cell style template identity is invalid');
      }
      templateIds.add(value.id);
    }
  }
  const sourceIds = new Set<string>();
  for (const source of dataModel.sources) {
    validateDataSourceManifest(source);
    const id = (source as { id: string }).id;
    if (sourceIds.has(id)) throw new Error(`Duplicate data source: ${id}`);
    sourceIds.add(id);
  }
  for (const [index, rawSheet] of input.sheets.entries()) {
    if (!rawSheet || typeof rawSheet !== 'object' || Array.isArray(rawSheet)) {
      throw new Error(`WorkbookSnapshot sheet[${index}] must be an object`);
    }
    const sheet = rawSheet as Record<string, unknown>;
    if (!isNonEmptyString(sheet.id) || !isNonEmptyString(sheet.name)) {
      throw new Error(`WorkbookSnapshot sheet[${index}] requires id and name`);
    }
    if (!['worksheet', 'table-sheet', 'gantt-sheet', 'report-sheet'].includes(String(sheet.kind))) {
      throw new Error(`WorkbookSnapshot sheet[${index}] kind is invalid`);
    }
    if (sheet.kind === 'table-sheet' && (!sheet.tableSheet || typeof sheet.tableSheet !== 'object')) throw new Error(`WorkbookSnapshot sheet[${index}] TableSheet definition is invalid`);
    if (sheet.kind === 'gantt-sheet' && (!sheet.ganttSheet || typeof sheet.ganttSheet !== 'object')) throw new Error(`WorkbookSnapshot sheet[${index}] GanttSheet definition is invalid`);
    if (sheet.kind === 'report-sheet' && (!sheet.reportSheet || typeof sheet.reportSheet !== 'object')) throw new Error(`WorkbookSnapshot sheet[${index}] ReportSheet definition is invalid`);
    if (!Number.isSafeInteger(sheet.rowCount) || Number(sheet.rowCount) <= 0
      || !Number.isSafeInteger(sheet.columnCount) || Number(sheet.columnCount) <= 0
      || !sheet.cells || typeof sheet.cells !== 'object'
      || !Array.isArray(sheet.merges)
      || !Array.isArray(sheet.pivots)
      || !Array.isArray(sheet.sparklines)) {
      throw new Error(`WorkbookSnapshot sheet[${index}] has invalid grid data`);
    }
    validateReviewSnapshot(sheet.review, String(sheet.id));
    if (typeof sheet.defaultRowHeightPx !== 'number' || !Number.isFinite(sheet.defaultRowHeightPx) || sheet.defaultRowHeightPx <= 0
      || typeof sheet.defaultColumnWidthPx !== 'number' || !Number.isFinite(sheet.defaultColumnWidthPx) || sheet.defaultColumnWidthPx <= 0
      || !sheet.pane || typeof sheet.pane !== 'object' || Array.isArray(sheet.pane)
      || !['none', 'frozen', 'split'].includes(String((sheet.pane as Record<string, unknown>).kind))) {
      throw new Error(`WorkbookSnapshot sheet[${index}] has invalid pixel geometry`);
    }
    if (sheet.rowHeightsPx !== undefined && (!sheet.rowHeightsPx || typeof sheet.rowHeightsPx !== 'object' || Array.isArray(sheet.rowHeightsPx))) {
      throw new Error(`WorkbookSnapshot sheet[${index}] rowHeightsPx is invalid`);
    }
    if (sheet.columnWidthsPx !== undefined && (!sheet.columnWidthsPx || typeof sheet.columnWidthsPx !== 'object' || Array.isArray(sheet.columnWidthsPx))) {
      throw new Error(`WorkbookSnapshot sheet[${index}] columnWidthsPx is invalid`);
    }
    if ('charts' in sheet || 'shapes' in sheet || 'images' in sheet) {
      throw new Error(`WorkbookSnapshot sheet[${index}] contains legacy drawing collections`);
    }
    if (!Array.isArray(sheet.drawings) || !sheet.drawingPayloads || typeof sheet.drawingPayloads !== 'object') {
      throw new Error(`WorkbookSnapshot sheet[${index}] requires canonical drawings and payloads`);
    }
    if (sheet.hyperlinks !== undefined) {
      if (!Array.isArray(sheet.hyperlinks) || sheet.hyperlinks.some((entry) => !entry || typeof entry !== 'object')) {
        throw new Error(`WorkbookSnapshot sheet[${index}] hyperlinks are invalid`);
      }
    }
    for (const pivot of sheet.pivots) validatePivotDefinition(pivot);
    if (sheet.dataRegions !== undefined) {
      if (!Array.isArray(sheet.dataRegions)) throw new Error(`WorkbookSnapshot sheet[${index}] dataRegions must be an array`);
      for (const region of sheet.dataRegions) {
        if (!region || typeof region !== 'object') throw new Error(`WorkbookSnapshot sheet[${index}] has invalid data region`);
        const dataRegion = region as Record<string, unknown>;
        if (!isNonEmptyString(dataRegion.id) || !isNonEmptyString(dataRegion.sourceId) || !isRangeRef(dataRegion.range)
          || !Number.isSafeInteger(dataRegion.headerRow) || Number(dataRegion.headerRow) < 0
          || !Number.isSafeInteger(dataRegion.revision) || Number(dataRegion.revision) < 0) {
          throw new Error(`WorkbookSnapshot sheet[${index}] has invalid data region`);
        }
      }
    }
  }
  return value as WorkbookSnapshot;
}

function validateSnapshotResponse(value: unknown, expectedUnitId?: string): SnapshotResponse {
  const input = requireRecord(value, 'Snapshot response');
  validateExactKeys(input, ['unitId', 'snapshot', 'revision', 'checksum'], 'Snapshot response');
  const snapshot = validateWorkbookSnapshot(input.snapshot);
  if (input.unitId !== undefined && !isNonEmptyString(input.unitId)) throw new Error('Snapshot response unitId is invalid');
  if (expectedUnitId && snapshot.unitId !== expectedUnitId) throw new Error('Snapshot response snapshot unitId does not match request');
  if (input.unitId !== undefined && input.unitId !== snapshot.unitId) throw new Error('Snapshot response unitId does not match snapshot');
  if (!Number.isSafeInteger(input.revision) || Number(input.revision) < 0) throw new Error('Snapshot response revision is invalid');
  if (input.checksum !== undefined && !isNonEmptyString(input.checksum)) throw new Error('Snapshot response checksum is invalid');
  return {
    ...(input.unitId === undefined ? {} : { unitId: input.unitId }),
    snapshot,
    revision: Number(input.revision),
    ...(input.checksum === undefined ? {} : { checksum: input.checksum }),
  };
}

function validateIsoTimestamp(value: unknown, label: string): string {
  if (!isNonEmptyString(value) || Number.isNaN(Date.parse(value))) throw new Error(`${label} must be an ISO timestamp`);
  return value;
}

function validateWorkbookSummary(value: unknown): WorkbookSummary {
  const input = requireRecord(value, 'Workbook summary');
  validateExactKeys(input, [
    'unitId', 'name', 'revision', 'updatedAt', 'role', 'ownerSubject', 'spaceId', 'spaceName', 'folderId',
    'locationPath', 'storageLocation', 'syncStatus', 'lifecycle', 'source', 'sourceFileName', 'deletedAt',
    'lastOpenedAt', 'favorite',
  ], 'Workbook summary');
  if (!isNonEmptyString(input.unitId) || !isNonEmptyString(input.name)) throw new Error('Workbook summary identity is invalid');
  if (input.name.length > MAX_WORKBOOK_NAME_LENGTH) throw new Error('Workbook summary name is too long');
  if (!Number.isSafeInteger(input.revision) || Number(input.revision) < 0) throw new Error('Workbook summary revision is invalid');
  validateIsoTimestamp(input.updatedAt, 'Workbook summary updatedAt');
  if (input.role !== undefined && !['owner', 'editor', 'commenter', 'viewer'].includes(String(input.role))) throw new Error('Workbook summary role is invalid');
  if (input.storageLocation !== undefined && !['local', 'remote', 'mirrored'].includes(String(input.storageLocation))) throw new Error('Workbook summary storageLocation is invalid');
  if (input.syncStatus !== undefined && !['synced', 'syncing', 'pending', 'offline', 'conflict', 'error'].includes(String(input.syncStatus))) throw new Error('Workbook summary syncStatus is invalid');
  if (input.lifecycle !== undefined && !['active', 'trashed'].includes(String(input.lifecycle))) throw new Error('Workbook summary lifecycle is invalid');
  if (input.source !== undefined && !['native', 'xlsx-import'].includes(String(input.source))) throw new Error('Workbook summary source is invalid');
  if (input.locationPath !== undefined && (!Array.isArray(input.locationPath) || input.locationPath.some((entry) => typeof entry !== 'string'))) throw new Error('Workbook summary locationPath is invalid');
  for (const key of ['ownerSubject', 'spaceId', 'spaceName', 'folderId', 'sourceFileName', 'deletedAt', 'lastOpenedAt'] as const) {
    if (input[key] !== undefined && input[key] !== null && typeof input[key] !== 'string') throw new Error(`Workbook summary ${key} is invalid`);
  }
  if (input.favorite !== undefined && typeof input.favorite !== 'boolean') throw new Error('Workbook summary favorite is invalid');
  const unitId = input.unitId as string;
  const name = input.name as string;
  const updatedAt = input.updatedAt as string;
  return {
    unitId,
    name,
    revision: Number(input.revision),
    updatedAt,
    ...(input.role === undefined ? {} : { role: input.role as WorkbookAclRole }),
    ...(input.ownerSubject == null ? {} : { ownerSubject: input.ownerSubject as string }),
    ...(input.spaceId == null ? {} : { spaceId: input.spaceId as string }),
    ...(input.spaceName == null ? {} : { spaceName: input.spaceName as string }),
    ...(input.folderId == null ? {} : { folderId: input.folderId as string }),
    ...(input.locationPath === undefined ? {} : { locationPath: input.locationPath as string[] }),
    ...(input.storageLocation === undefined ? {} : { storageLocation: input.storageLocation as WorkbookStorageLocation }),
    ...(input.syncStatus === undefined ? {} : { syncStatus: input.syncStatus as WorkbookSyncStatus }),
    ...(input.lifecycle === undefined ? {} : { lifecycle: input.lifecycle as WorkbookLifecycle }),
    ...(input.source === undefined ? {} : { source: input.source as WorkbookSourceKind }),
    ...(input.sourceFileName == null ? {} : { sourceFileName: input.sourceFileName as string }),
    ...(input.deletedAt == null ? {} : { deletedAt: input.deletedAt as string }),
    ...(input.lastOpenedAt == null ? {} : { lastOpenedAt: input.lastOpenedAt as string }),
    ...(input.favorite === undefined ? {} : { favorite: input.favorite }),
  };
}

function validateCursorPage<T>(value: unknown, itemValidator: (item: unknown) => T, label: string): CursorPage<T> {
  const input = requireRecord(value, label);
  validateExactKeys(input, ['items', 'nextCursor'], label);
  if (!Array.isArray(input.items)) throw new Error(`${label}.items must be an array`);
  if (input.nextCursor !== null && input.nextCursor !== undefined && !isNonEmptyString(input.nextCursor)) {
    throw new Error(`${label}.nextCursor is invalid`);
  }
  return {
    items: input.items.map(itemValidator),
    nextCursor: input.nextCursor == null ? null : input.nextCursor,
  };
}

function validateRevisionRecord(value: unknown): RevisionRecord {
  const input = requireRecord(value, 'Revision record');
  validateExactKeys(input, ['operationId', 'revision', 'createdAt', 'payload'], 'Revision record');
  if (!isNonEmptyString(input.operationId)) throw new Error('Revision record operationId is invalid');
  if (!Number.isSafeInteger(input.revision) || Number(input.revision) < 1) throw new Error('Revision record revision is invalid');
  validateIsoTimestamp(input.createdAt, 'Revision record createdAt');
  const payload = validateCommittedOperationEnvelope(input.payload);
  if (payload.revision !== Number(input.revision)) throw new Error('Revision record payload revision does not match record');
  return { operationId: input.operationId, revision: Number(input.revision), createdAt: input.createdAt as string, payload };
}

export function validateUserPreferences(value: unknown): UserPreferences {
  const input = requireRecord(value, 'User preferences');
  validateExactKeys(input, ['defaultSpaceId', 'defaultFolderId', 'autoSave', 'autoSync', 'offlineCache', 'importCompatibility', 'language', 'theme', 'updatedAt'], 'User preferences');
  for (const key of ['autoSave', 'autoSync', 'offlineCache'] as const) {
    if (typeof input[key] !== 'boolean') throw new Error(`User preferences ${key} is invalid`);
  }
  if (!['A', 'B', 'C'].includes(String(input.importCompatibility))) throw new Error('User preferences importCompatibility is invalid');
  if (!['light', 'dark', 'system'].includes(String(input.theme))) throw new Error('User preferences theme is invalid');
  for (const key of ['defaultSpaceId', 'defaultFolderId', 'language'] as const) {
    if (input[key] !== undefined && input[key] !== null && typeof input[key] !== 'string') throw new Error(`User preferences ${key} is invalid`);
  }
  if (input.updatedAt !== undefined && input.updatedAt !== null) validateIsoTimestamp(input.updatedAt, 'User preferences updatedAt');
  const autoSave = input.autoSave as boolean;
  const autoSync = input.autoSync as boolean;
  const offlineCache = input.offlineCache as boolean;
  const defaultSpaceId = input.defaultSpaceId as string | null | undefined;
  const defaultFolderId = input.defaultFolderId as string | null | undefined;
  const language = input.language as string | null | undefined;
  const updatedAt = input.updatedAt as string | null | undefined;
  return {
    ...(defaultSpaceId == null ? {} : { defaultSpaceId }),
    ...(defaultFolderId == null ? {} : { defaultFolderId }),
    autoSave,
    autoSync,
    offlineCache,
    importCompatibility: input.importCompatibility as UserPreferences['importCompatibility'],
    ...(language == null ? {} : { language }),
    theme: input.theme as UserPreferences['theme'],
    ...(updatedAt == null ? {} : { updatedAt }),
  };
}

export function validateUserPreferencesPatch(value: unknown): UserPreferencesPatch {
  const input = requireRecord(value, 'User preferences patch');
  validateExactKeys(input, ['defaultSpaceId', 'defaultFolderId', 'autoSave', 'autoSync', 'offlineCache', 'importCompatibility', 'language', 'theme'], 'User preferences patch');
  if (Object.keys(input).length === 0) throw new Error('User preferences patch cannot be empty');
  for (const key of ['defaultSpaceId', 'defaultFolderId', 'language'] as const) {
    if (input[key] !== undefined && input[key] !== null && typeof input[key] !== 'string') throw new Error(`User preferences patch ${key} is invalid`);
  }
  for (const key of ['autoSave', 'autoSync', 'offlineCache'] as const) {
    if (input[key] !== undefined && typeof input[key] !== 'boolean') throw new Error(`User preferences patch ${key} is invalid`);
  }
  if (input.importCompatibility !== undefined && !['A', 'B', 'C'].includes(String(input.importCompatibility))) throw new Error('User preferences patch importCompatibility is invalid');
  if (input.theme !== undefined && !['light', 'dark', 'system'].includes(String(input.theme))) throw new Error('User preferences patch theme is invalid');
  return value as UserPreferencesPatch;
}

function validateWorkbookAccessResponse(value: unknown): WorkbookAccessResponse {
  if (!value || typeof value !== 'object' || Array.isArray(value)) {
    throw new Error('workbook access response must be an object');
  }
  const input = value as Record<string, unknown>;
  if (!isNonEmptyString(input.unitId)) throw new Error('workbook access response requires unitId');
  if (input.role !== 'owner' && input.role !== 'editor' && input.role !== 'commenter' && input.role !== 'viewer') {
    throw new Error('workbook access response has an invalid role');
  }
  return { unitId: input.unitId, role: input.role };
}

/** Strict runtime validation used at the REST/WebSocket trust boundary. */
export function validateOperationEnvelope(value: unknown): OperationEnvelope {
  if (!value || typeof value !== 'object') throw new Error('OperationEnvelope must be an object');
  const input = value as Record<string, unknown>;
  if (input.schema !== OPERATION_ENVELOPE_SCHEMA) throw new Error('Unsupported operation schema');
  if (!isNonEmptyString(input.operationId) || !isNonEmptyString(input.unitId)) {
    throw new Error('operationId and unitId are required');
  }
  if (!Number.isSafeInteger(input.clientSequence) || Number(input.clientSequence) < 1) {
    throw new Error('clientSequence must be a positive safe integer');
  }
  if (!Number.isSafeInteger(input.baseRevision) || Number(input.baseRevision) < 0) {
    throw new Error('baseRevision must be a non-negative safe integer');
  }
  if (!isNonEmptyString(input.createdAt) || Number.isNaN(Date.parse(input.createdAt))) {
    throw new Error('createdAt must be an ISO timestamp');
  }
  if (!Array.isArray(input.mutations) || input.mutations.length === 0) {
    throw new Error('mutations must contain at least one mutation');
  }

  // Reject fields that used to be client-controlled security inputs instead
  // of silently ignoring them. This prevents accidental reintroduction of
  // obsolete semantics through an untyped JSON caller.
  if ('actorId' in input || 'affectedRanges' in input) {
    throw new Error('actorId and affectedRanges are server-owned fields');
  }
  let intent: OperationIntent | undefined;
  if (input.intent !== undefined) {
    if (!input.intent || typeof input.intent !== 'object' || Array.isArray(input.intent)) {
      throw new Error('operation intent must be an object');
    }
    const rawIntent = input.intent as Record<string, unknown>;
    const unknownIntentKeys = Object.keys(rawIntent).filter((key) => !['type', 'targetOperationId', 'targetBaseRevision'].includes(key));
    if (unknownIntentKeys.length > 0) throw new Error(`operation intent contains unsupported fields: ${unknownIntentKeys.join(', ')}`);
    if (rawIntent.type !== 'undo' || !isNonEmptyString(rawIntent.targetOperationId)) {
      throw new Error('operation intent must identify an undo target');
    }
    if (!Number.isSafeInteger(rawIntent.targetBaseRevision) || Number(rawIntent.targetBaseRevision) < 0) {
      throw new Error('operation intent targetBaseRevision must be a non-negative safe integer');
    }
    intent = {
      type: 'undo',
      targetOperationId: rawIntent.targetOperationId,
      targetBaseRevision: Number(rawIntent.targetBaseRevision),
    };
  }

  const mutations = input.mutations.map((raw, index) => {
    if (!raw || typeof raw !== 'object') throw new Error(`mutation[${index}] must be an object`);
    const mutation = raw as Record<string, unknown>;
    if (!isNonEmptyString(mutation.id) || !isNonEmptyString(mutation.sheetId)) {
      throw new Error(`mutation[${index}] requires id and sheetId`);
    }
    if (!('params' in mutation)) throw new Error(`mutation[${index}] requires params`);
    if ('affectedRanges' in mutation || 'actorId' in mutation) {
      throw new Error(`mutation[${index}] contains server-owned fields`);
    }
    if (isDataSourceMutationId(mutation.id)) {
      validateDataSourceMutationParams(mutation.id, mutation.params);
    }
    validatePivotMutationParams(mutation.id, mutation.params);
    return {
      id: mutation.id,
      sheetId: mutation.sheetId,
      params: mutation.params,
    } satisfies OperationMutation;
  });

  return {
    schema: OPERATION_ENVELOPE_SCHEMA,
    operationId: input.operationId,
    unitId: input.unitId,
    clientSequence: Number(input.clientSequence),
    baseRevision: Number(input.baseRevision),
    mutations,
    createdAt: input.createdAt,
    ...(intent ? { intent } : {}),
  };
}

export interface SnapshotResponse {
  /** Present on Java responses; optional for local-only test and cache records. */
  unitId?: string;
  snapshot: WorkbookSnapshot;
  revision: number;
  checksum?: string;
}

/**
 * The only client-authored history restore request.  The historical snapshot
 * is deliberately absent: the server resolves targetRevision from its own
 * immutable revision log and creates the restore operation after ACL checks.
 */
export interface HistoryRestoreRequest {
  targetRevision: number;
  reason?: string;
}

/** Server response after committing a server-generated workbook.restore op. */
export interface HistoryRestoreResponse {
  operation: CommittedOperationEnvelope;
  snapshot: SnapshotResponse;
}

export interface HistoryAuditRecord {
  auditId: string;
  unitId: string;
  operationId: string;
  actorId: string;
  action: 'workbook.restore';
  targetRevision: number;
  revision: number;
  reason?: string;
  createdAt: string;
}

/** Strict validation for the REST history restore trust boundary. */
export function validateHistoryRestoreRequest(value: unknown): HistoryRestoreRequest {
  if (!value || typeof value !== 'object' || Array.isArray(value)) {
    throw new Error('History restore request must be an object');
  }
  const input = value as Record<string, unknown>;
  const unknownKeys = Object.keys(input).filter((key) => key !== 'targetRevision' && key !== 'reason');
  if (unknownKeys.length > 0) {
    throw new Error(`History restore request contains unsupported fields: ${unknownKeys.join(', ')}`);
  }
  if (!Number.isSafeInteger(input.targetRevision) || Number(input.targetRevision) < 0) {
    throw new Error('targetRevision must be a non-negative safe integer');
  }
  if (input.reason !== undefined && (typeof input.reason !== 'string' || input.reason.length > 1000)) {
    throw new Error('reason must be a string with at most 1000 characters');
  }
  return {
    targetRevision: Number(input.targetRevision),
    ...(input.reason === undefined ? {} : { reason: input.reason }),
  };
}

export interface CompatibilityReportPayload {
  schema: 'CompatibilityReport';
  fileName: string;
  importLevel: 'A' | 'B' | 'C';
  exportLevel: 'A' | 'B' | 'C';
  dateSystem: '1900' | '1904';
  issues: Array<{
    level: 'A' | 'B' | 'C';
    severity: 'error' | 'warning' | 'info';
    feature: string;
    location?: string;
    message: string;
    preserved: boolean;
  }>;
  summary: {
    editableFeatures: number;
    preservedOnly: number;
    unsupported: number;
  };
}

export interface XlsxImportResponse extends SnapshotResponse {
  report: CompatibilityReportPayload;
}

export interface XlsxExportResponse {
  unitId: string;
  base64: string;
  fileName: string;
  report: CompatibilityReportPayload;
}

/** Sanitized intent for a server-executed database or credentialed REST query. */
export interface ServerQueryRequest {
  queryId: string;
  name: string;
  connectorId: 'sqlite' | 'jdbc' | 'rest';
  sourceRef: string;
  statement: string;
  method?: 'GET' | 'POST';
  body?: unknown;
  parameters?: unknown[];
  steps: Array<{ id: string; kind: string; name: string; config: Record<string, unknown>; enabled: boolean }>;
}

export interface ServerQueryResponse {
  queryId: string;
  connectorId: string;
  sourceRef: string;
  sourceRevision: number;
  columns: string[];
  rows: TableScalar[][];
  rowCount: number;
  executedAt: string;
  durationMs: number;
}

export type GuestShareRole = 'viewer' | 'commenter' | 'editor';

export interface GuestShareRequest {
  role: GuestShareRole;
  expiresAt?: string;
}

export interface GuestShareResponse {
  shareId: string;
  unitId: string;
  role: GuestShareRole;
  expiresAt: string;
  revokedAt?: string;
  createdBy: string;
  createdAt: string;
  /** Returned exactly once when a share is created. */
  token?: string;
}

export interface WorkbookSummary {
  unitId: string;
  name: string;
  revision: number;
  updatedAt: string;
  role?: WorkbookAclRole;
  ownerSubject?: string;
  spaceId?: string;
  spaceName?: string;
  folderId?: string;
  locationPath?: string[];
  storageLocation?: WorkbookStorageLocation;
  syncStatus?: WorkbookSyncStatus;
  lifecycle?: WorkbookLifecycle;
  source?: WorkbookSourceKind;
  sourceFileName?: string;
  deletedAt?: string;
  lastOpenedAt?: string;
  favorite?: boolean;
}

export type WorkbookSourceKind = 'native' | 'xlsx-import';
export type WorkbookStorageLocation = 'local' | 'remote' | 'mirrored';
export type WorkbookSyncStatus = 'synced' | 'syncing' | 'pending' | 'offline' | 'conflict' | 'error';
export type WorkbookLifecycle = 'active' | 'trashed';
export type WorkbookCatalogView = 'all' | 'owned' | 'recent' | 'shared' | 'trash';

export interface WorkbookCatalogQuery {
  folderId?: string;
  query?: string;
  spaceId?: string;
  view?: WorkbookCatalogView;
  cursor?: string;
  limit?: number;
}

export interface WorkbookCreateMetadata {
  folderId?: string;
  source?: WorkbookSourceKind;
  spaceId?: string;
}

export interface WorkbookMetadataPatch {
  folderId?: string | null;
  spaceId?: string | null;
}

export interface WorkbookCopyRequest {
  folderId?: string;
  name?: string;
  spaceId?: string;
}

export interface WorkbookUserState {
  autoSave?: boolean;
  autoSync?: boolean;
  defaultCreateLocation?: 'local' | 'remote';
  favorite?: boolean;
  importCompatibilityLevel?: 'standard' | 'strict';
  language?: string;
  lastOpenedAt?: string;
  offlineCache?: boolean;
  theme?: 'light' | 'system';
  unitId: string;
}

export interface UserPreferences {
  defaultSpaceId?: string;
  defaultFolderId?: string;
  autoSave: boolean;
  autoSync: boolean;
  offlineCache: boolean;
  importCompatibility: 'A' | 'B' | 'C';
  language?: string;
  theme: 'light' | 'dark' | 'system';
  updatedAt?: string;
}

export type UserPreferencesPatch = Partial<Omit<UserPreferences, 'updatedAt'>>;

export type WorkspaceSpaceKind = 'personal' | 'team';

export interface WorkspaceSpace {
  createdAt: string;
  createdBy: string;
  kind: WorkspaceSpaceKind;
  name: string;
  role?: WorkbookAclRole;
  spaceId: string;
  updatedAt: string;
}

export interface WorkspaceFolder {
  folderId: string;
  name: string;
  parentFolderId?: string;
  spaceId: string;
  updatedAt: string;
}

export interface SpaceMember {
  role: WorkbookAclRole;
  spaceId: string;
  subject: string;
  updatedAt: string;
}

export interface WorkbookSourceArtifactMetadata {
  byteLength: number;
  checksum: string;
  createdAt?: string;
  fileName: string;
  mimeType?: string;
  unitId: string;
  updatedAt: string;
  format?: string;
  codecRevision?: number;
}

export interface WorkbookImportRequest extends WorkbookCreateMetadata {
  artifact: Blob;
  artifactFileName: string;
  snapshot: WorkbookSnapshot;
  format: string;
  nativeMetadata: Record<string, unknown>;
}

export interface WorkbookImportResponse {
  artifact: WorkbookSourceArtifactMetadata;
  snapshot: WorkbookSnapshot;
  summary: WorkbookSummary;
}

export interface RevisionRecord {
  operationId: string;
  revision: number;
  createdAt: string;
  payload: CommittedOperationEnvelope;
}

/** Metadata-only remote representation of a content-addressed data block. */
export interface RemoteDataBlockMetadata {
  unitId: string;
  sourceId: string;
  blockId: string;
  checksum: string;
  byteLength: number;
  updatedAt: string;
}

/** Metadata-only remote representation of a workbook-owned image asset. */
export interface RemoteAssetMetadata extends AssetRef {
  unitId: string;
  updatedAt: string;
}

async function sha256(blob: Blob): Promise<string> {
  const digest = await crypto.subtle.digest('SHA-256', await blob.arrayBuffer());
  return Array.from(new Uint8Array(digest), (value) => value.toString(16).padStart(2, '0')).join('');
}

export class WorkbookApiClient {
  private readonly baseUrl: string;
  private readonly authTokenProvider?: AuthTokenProvider;
  private readonly shareTokenProvider?: ShareTokenProvider;
  private readonly fetchImpl: typeof fetch;

  constructor(options: WorkbookApiClientOptions = {}) {
    this.baseUrl = options.baseUrl ?? '';
    this.authTokenProvider = options.authTokenProvider;
    this.shareTokenProvider = options.shareTokenProvider;
    this.fetchImpl = options.fetchImpl ?? fetch;
  }

  private async request(path: string, init: RequestInit = {}): Promise<Response> {
    const token = (await this.authTokenProvider?.())?.trim();
    const shareToken = token ? undefined : (await this.shareTokenProvider?.())?.trim();
    if (!token && !shareToken) throw new AuthenticationRequiredError();
    const headers = new Headers(init.headers);
    if (token) headers.set('authorization', `Bearer ${token}`);
    else headers.set('x-workbook-share-token', shareToken!);
    const response = await this.fetchImpl(`${this.baseUrl}${path}`, { ...init, headers });
    if (response.ok) return response;

    let payload: Partial<ApiError> | undefined;
    try {
      const candidate = await response.clone().json() as unknown;
      if (candidate && typeof candidate === 'object') payload = candidate as Partial<ApiError>;
    } catch {
      // The status remains authoritative when a server returns a non-JSON
      // failure body.  Do not swallow the request failure itself.
    }
    const code = typeof payload?.code === 'string'
      && (CONTRACT_ERROR_CODES as readonly string[]).includes(payload.code)
      ? payload.code as ProtocolErrorCode
      : 'INTERNAL_ERROR';
    throw new ApiRequestError(
      typeof payload?.message === 'string' ? payload.message : `Request failed: ${response.status}`,
      response.status,
      code,
      payload?.details,
    );
  }

  private async json<T>(path: string, init: RequestInit = {}): Promise<T> {
    const response = await this.request(path, init);
    try {
      return await response.json() as T;
    } catch {
      throw new ApiRequestError(`Invalid JSON response from ${path}`, response.status, 'INTERNAL_ERROR');
    }
  }

  async getSnapshot(unitId: string, options: ApiRequestOptions = {}): Promise<SnapshotResponse> {
    return validateSnapshotResponse(await this.json<unknown>(
      `/api/workbooks/${encodeURIComponent(unitId)}/snapshot`,
      options,
    ), unitId);
  }

  async getAccess(unitId: string, options: ApiRequestOptions = {}): Promise<WorkbookAccessResponse> {
    const result = await this.json<unknown>(`/api/workbooks/${encodeURIComponent(unitId)}/access`, options);
    return validateWorkbookAccessResponse(result);
  }

  async listWorkbookAcl(unitId: string, options: ApiRequestOptions = {}): Promise<WorkbookAclRecord[]> {
    return this.json<WorkbookAclRecord[]>(`/api/workbooks/${encodeURIComponent(unitId)}/acl`, options);
  }

  async putWorkbookAcl(unitId: string, subject: string, role: WorkbookAclRole): Promise<WorkbookAclRecord> {
    return this.json<WorkbookAclRecord>(`/api/workbooks/${encodeURIComponent(unitId)}/acl/${encodeURIComponent(subject)}`, {
      method: 'PUT',
      headers: { 'content-type': 'application/json' },
      body: JSON.stringify({ role }),
    });
  }

  async deleteWorkbookAcl(unitId: string, subject: string): Promise<void> {
    await this.request(`/api/workbooks/${encodeURIComponent(unitId)}/acl/${encodeURIComponent(subject)}`, { method: 'DELETE' });
  }

  async createWorkbook(snapshot: WorkbookSnapshot, metadata: WorkbookCreateMetadata = {}): Promise<SnapshotResponse> {
    validateWorkbookSnapshot(snapshot);
    return validateSnapshotResponse(await this.json<unknown>('/api/workbooks', {
      method: 'POST',
      headers: { 'content-type': 'application/json' },
      body: JSON.stringify({ unitId: snapshot.unitId, name: snapshot.name, snapshot, ...metadata }),
    }), snapshot.unitId);
  }

  async listWorkbookPage(query: WorkbookCatalogQuery = {}, options: ApiRequestOptions = {}): Promise<CursorPage<WorkbookSummary>> {
    const search = new URLSearchParams();
    if (query.view) search.set('view', query.view);
    if (query.spaceId) search.set('spaceId', query.spaceId);
    if (query.folderId) search.set('folderId', query.folderId);
    if (query.query) search.set('query', query.query);
    if (query.cursor) search.set('cursor', query.cursor);
    search.set('limit', String(normalizePageLimit(query.limit)));
    const suffix = search.size > 0 ? `?${search.toString()}` : '';
    return validateCursorPage(
      await this.json<unknown>(`/api/workbooks${suffix}`, options),
      validateWorkbookSummary,
      'Workbook catalog page',
    );
  }

  /**
   * Reads the complete catalog through bounded pages for existing callers.
   * New consumers should use listWorkbookPage when they need incremental UI.
   */
  async listWorkbooks(query: WorkbookCatalogQuery = {}, options: ApiRequestOptions = {}): Promise<WorkbookSummary[]> {
    const items: WorkbookSummary[] = [];
    const seen = new Set<string>();
    let cursor = query.cursor;
    do {
      const page = await this.listWorkbookPage({ ...query, cursor }, options);
      items.push(...page.items);
      if (!page.nextCursor) break;
      if (seen.has(page.nextCursor)) throw new ApiRequestError('Workbook catalog pagination cursor did not advance', 200, 'INTERNAL_ERROR');
      seen.add(page.nextCursor);
      cursor = page.nextCursor;
    } while (!options.signal?.aborted);
    if (options.signal?.aborted) throw new DOMException('The operation was aborted.', 'AbortError');
    return items;
  }

  async updateWorkbook(unitId: string, patch: WorkbookMetadataPatch): Promise<WorkbookSummary> {
    return this.json<WorkbookSummary>(`/api/workbooks/${encodeURIComponent(unitId)}`, {
      method: 'PATCH',
      headers: { 'content-type': 'application/json' },
      body: JSON.stringify(patch),
    });
  }

  async copyWorkbook(unitId: string, request: WorkbookCopyRequest = {}): Promise<WorkbookSummary> {
    return this.json<WorkbookSummary>(`/api/workbooks/${encodeURIComponent(unitId)}/copy`, {
      method: 'POST',
      headers: { 'content-type': 'application/json' },
      body: JSON.stringify(request),
    });
  }

  async moveToTrash(unitId: string): Promise<void> {
    await this.request(`/api/workbooks/${encodeURIComponent(unitId)}`, { method: 'DELETE' });
  }

  async restoreFromTrash(unitId: string): Promise<WorkbookSummary> {
    return this.json<WorkbookSummary>(`/api/workbooks/${encodeURIComponent(unitId)}/restore-from-trash`, { method: 'POST' });
  }

  async purgeWorkbook(unitId: string): Promise<void> {
    await this.request(`/api/workbooks/${encodeURIComponent(unitId)}/purge`, { method: 'DELETE' });
  }

  async getWorkbookUserState(unitId: string): Promise<WorkbookUserState> {
    return this.json<WorkbookUserState>(`/api/workbooks/${encodeURIComponent(unitId)}/user-state`);
  }

  async putWorkbookUserState(unitId: string, state: Omit<WorkbookUserState, 'unitId'>): Promise<WorkbookUserState> {
    return this.json<WorkbookUserState>(`/api/workbooks/${encodeURIComponent(unitId)}/user-state`, {
      method: 'PUT',
      headers: { 'content-type': 'application/json' },
      body: JSON.stringify(state),
    });
  }

  async listSpaces(options: ApiRequestOptions = {}): Promise<WorkspaceSpace[]> {
    return this.json<WorkspaceSpace[]>('/api/spaces', options);
  }

  async getUserPreferences(options: ApiRequestOptions = {}): Promise<UserPreferences> {
    return validateUserPreferences(await this.json<unknown>('/api/user-preferences', options));
  }

  async putUserPreferences(preferences: UserPreferencesPatch, options: ApiRequestOptions = {}): Promise<UserPreferences> {
    validateUserPreferencesPatch(preferences);
    return validateUserPreferences(await this.json<unknown>('/api/user-preferences', {
      method: 'PUT',
      headers: { 'content-type': 'application/json' },
      body: JSON.stringify(preferences),
      ...options,
    }));
  }

  async createSpace(input: Pick<WorkspaceSpace, 'kind' | 'name'>): Promise<WorkspaceSpace> {
    return this.json<WorkspaceSpace>('/api/spaces', {
      method: 'POST',
      headers: { 'content-type': 'application/json' },
      body: JSON.stringify(input),
    });
  }

  async listFolders(spaceId: string, options: ApiRequestOptions = {}): Promise<WorkspaceFolder[]> {
    return this.json<WorkspaceFolder[]>(`/api/spaces/${encodeURIComponent(spaceId)}/folders`, options);
  }

  async createFolder(spaceId: string, input: Pick<WorkspaceFolder, 'name' | 'parentFolderId'>): Promise<WorkspaceFolder> {
    return this.json<WorkspaceFolder>(`/api/spaces/${encodeURIComponent(spaceId)}/folders`, {
      method: 'POST',
      headers: { 'content-type': 'application/json' },
      body: JSON.stringify(input),
    });
  }

  async updateFolder(folderId: string, input: Pick<WorkspaceFolder, 'name' | 'parentFolderId'>): Promise<WorkspaceFolder> {
    return this.json<WorkspaceFolder>(`/api/folders/${encodeURIComponent(folderId)}`, {
      method: 'PATCH',
      headers: { 'content-type': 'application/json' },
      body: JSON.stringify(input),
    });
  }

  async deleteFolder(folderId: string): Promise<void> {
    await this.request(`/api/folders/${encodeURIComponent(folderId)}`, { method: 'DELETE' });
  }

  async listSpaceMembers(spaceId: string, options: ApiRequestOptions = {}): Promise<SpaceMember[]> {
    return this.json<SpaceMember[]>(`/api/spaces/${encodeURIComponent(spaceId)}/members`, options);
  }

  async putSpaceMember(spaceId: string, subject: string, role: WorkbookAclRole): Promise<SpaceMember> {
    return this.json<SpaceMember>(`/api/spaces/${encodeURIComponent(spaceId)}/members/${encodeURIComponent(subject)}`, {
      method: 'PUT',
      headers: { 'content-type': 'application/json' },
      body: JSON.stringify({ role }),
    });
  }

  async deleteSpaceMember(spaceId: string, subject: string): Promise<void> {
    await this.request(`/api/spaces/${encodeURIComponent(spaceId)}/members/${encodeURIComponent(subject)}`, { method: 'DELETE' });
  }

  async createWorkbookImport(input: WorkbookImportRequest): Promise<WorkbookImportResponse> {
    validateWorkbookSnapshot(input.snapshot);
    const form = new FormData();
    form.append('file', input.artifact, input.artifactFileName);
    form.append('snapshot', JSON.stringify(input.snapshot));
    form.append('format', input.format);
    form.append('nativeMetadata', JSON.stringify(input.nativeMetadata));
    if (input.snapshot.name) form.append('name', input.snapshot.name);
    if (input.spaceId) form.append('spaceId', input.spaceId);
    if (input.folderId) form.append('folderId', input.folderId);
    const response = await this.json<WorkbookImportResponse>('/api/workbook-imports', { method: 'POST', body: form });
    validateWorkbookSnapshot(response.snapshot);
    validateWorkbookSummary(response.summary);
    if (response.summary.unitId !== input.snapshot.unitId || response.snapshot.unitId !== input.snapshot.unitId) {
      throw new ApiRequestError('Workbook import returned a mismatched identity', 200, 'INTERNAL_ERROR');
    }
    return response;
  }

  async putWorkbookSourceArtifact(unitId: string, artifact: Blob, fileName: string): Promise<WorkbookSourceArtifactMetadata> {
    const checksum = await sha256(artifact);
    return this.json<WorkbookSourceArtifactMetadata>(`/api/workbooks/${encodeURIComponent(unitId)}/native-package-state`, {
      method: 'PUT',
      headers: {
        'content-type': 'application/octet-stream',
        'x-content-sha256': checksum,
        'x-file-name': encodeURIComponent(fileName),
      },
      body: artifact,
    });
  }

  async getWorkbookSourceArtifact(unitId: string): Promise<{ artifact: Blob; metadata: WorkbookSourceArtifactMetadata }> {
    const response = await this.request(`/api/workbooks/${encodeURIComponent(unitId)}/native-package-state`);
    const disposition = response.headers.get('content-disposition') ?? '';
    const fileNameMatch = /filename="?([^";]+)"?/i.exec(disposition);
    const fileName = fileNameMatch?.[1];
    const checksum = response.headers.get('x-content-sha256');
    const byteLength = Number(response.headers.get('content-length') ?? 0);
    const codecRevision = Number(response.headers.get('x-native-codec-revision') ?? 1);
    const format = response.headers.get('x-native-format') ?? undefined;
    if (!fileName || !checksum) throw new ApiRequestError('Workbook source artifact response omitted metadata', response.status, 'INTERNAL_ERROR');
    const artifact = await response.blob();
    return {
      artifact,
      metadata: {
        byteLength: byteLength || artifact.size,
        checksum,
        fileName,
        mimeType: response.headers.get('content-type') ?? undefined,
        unitId,
        updatedAt: response.headers.get('last-modified') ?? new Date().toISOString(),
        format,
        codecRevision: Number.isSafeInteger(codecRevision) && codecRevision > 0 ? codecRevision : 1,
      },
    };
  }

  async commitOperation(unitId: string, operation: OperationEnvelope): Promise<OperationCommitResponse> {
    return this.json<OperationCommitResponse>(`/api/workbooks/${encodeURIComponent(unitId)}/operations`, {
      method: 'POST',
      headers: { 'content-type': 'application/json' },
      body: JSON.stringify(operation),
    });
  }

  async checkpointWorkbook(unitId: string): Promise<CheckpointResponse> {
    return this.json<CheckpointResponse>(`/api/workbooks/${encodeURIComponent(unitId)}/checkpoints`, { method: 'POST' });
  }

  async listRevisionPage(unitId: string, query: { cursor?: string; limit?: number } = {}, options: ApiRequestOptions = {}): Promise<CursorPage<RevisionRecord>> {
    const search = new URLSearchParams();
    if (query.cursor) search.set('cursor', query.cursor);
    search.set('limit', String(normalizePageLimit(query.limit)));
    return validateCursorPage(
      await this.json<unknown>(`/api/workbooks/${encodeURIComponent(unitId)}/revisions?${search.toString()}`, options),
      validateRevisionRecord,
      'Revision page',
    );
  }

  /** Reads all bounded revision pages for the existing history runtime. */
  async listRevisions(unitId: string, options: ApiRequestOptions = {}): Promise<RevisionRecord[]> {
    const items: RevisionRecord[] = [];
    const seen = new Set<string>();
    let cursor: string | undefined;
    do {
      const page = await this.listRevisionPage(unitId, { cursor }, options);
      items.push(...page.items);
      if (!page.nextCursor) break;
      if (seen.has(page.nextCursor)) throw new ApiRequestError('Revision pagination cursor did not advance', 200, 'INTERNAL_ERROR');
      seen.add(page.nextCursor);
      cursor = page.nextCursor;
    } while (!options.signal?.aborted);
    if (options.signal?.aborted) throw new DOMException('The operation was aborted.', 'AbortError');
    return items;
  }

  async getRevisionSnapshot(unitId: string, revision: number, options: ApiRequestOptions = {}): Promise<SnapshotResponse> {
    return validateSnapshotResponse(await this.json<unknown>(
      `/api/workbooks/${encodeURIComponent(unitId)}/revisions/${revision}/snapshot`,
      options,
    ), unitId);
  }

  async restoreToRevision(unitId: string, targetRevision: number, reason?: string): Promise<HistoryRestoreResponse> {
    const body = validateHistoryRestoreRequest({ targetRevision, ...(reason === undefined ? {} : { reason }) });
    return this.json<HistoryRestoreResponse>(
      `/api/workbooks/${encodeURIComponent(unitId)}/restore`,
      {
        method: 'POST',
        headers: { 'content-type': 'application/json' },
        body: JSON.stringify(body),
      },
    );
  }

  async executeServerQuery(unitId: string, request: ServerQueryRequest): Promise<ServerQueryResponse> {
    return this.json<ServerQueryResponse>(`/api/workbooks/${encodeURIComponent(unitId)}/queries/execute`, {
      method: 'POST',
      headers: { 'content-type': 'application/json' },
      body: JSON.stringify(request),
    });
  }

  async cancelServerQuery(unitId: string, queryId: string): Promise<void> {
    await this.request(`/api/workbooks/${encodeURIComponent(unitId)}/queries/${encodeURIComponent(queryId)}/cancel`, {
      method: 'POST',
    });
  }

  async createGuestShare(unitId: string, request: GuestShareRequest): Promise<GuestShareResponse> {
    return this.json<GuestShareResponse>(`/api/workbooks/${encodeURIComponent(unitId)}/shares`, {
      method: 'POST',
      headers: { 'content-type': 'application/json' },
      body: JSON.stringify(request),
    });
  }

  async listGuestShares(unitId: string): Promise<GuestShareResponse[]> {
    return this.json<GuestShareResponse[]>(`/api/workbooks/${encodeURIComponent(unitId)}/shares`);
  }

  async revokeGuestShare(unitId: string, shareId: string): Promise<void> {
    await this.request(`/api/workbooks/${encodeURIComponent(unitId)}/shares/${encodeURIComponent(shareId)}`, {
      method: 'DELETE',
    });
  }

  async putDataBlock(
    unitId: string,
    sourceId: string,
    blockId: string,
    checksum: string,
    bytes: ArrayBuffer,
  ): Promise<RemoteDataBlockMetadata> {
    return this.json<RemoteDataBlockMetadata>(
      `/api/workbooks/${encodeURIComponent(unitId)}/data-sources/${encodeURIComponent(sourceId)}/blocks/${encodeURIComponent(blockId)}`,
      {
        method: 'PUT',
        headers: {
          'content-type': 'application/octet-stream',
          'x-content-sha256': checksum,
        },
        body: bytes,
      },
    );
  }

  async getDataBlock(unitId: string, sourceId: string, blockId: string): Promise<{ bytes: ArrayBuffer; checksum: string }> {
    const response = await this.request(
      `/api/workbooks/${encodeURIComponent(unitId)}/data-sources/${encodeURIComponent(sourceId)}/blocks/${encodeURIComponent(blockId)}`,
    );
    const checksum = response.headers.get('x-content-sha256');
    if (!checksum) throw new ApiRequestError('Data block response omitted checksum', response.status, 'INTERNAL_ERROR');
    return { bytes: await response.arrayBuffer(), checksum };
  }

  async deleteDataBlock(unitId: string, sourceId: string, blockId: string): Promise<void> {
    await this.request(
      `/api/workbooks/${encodeURIComponent(unitId)}/data-sources/${encodeURIComponent(sourceId)}/blocks/${encodeURIComponent(blockId)}`,
      { method: 'DELETE' },
    );
  }

  async putAsset(unitId: string, asset: AssetRef, bytes: ArrayBuffer): Promise<RemoteAssetMetadata> {
    return this.json<RemoteAssetMetadata>(
      `/api/workbooks/${encodeURIComponent(unitId)}/assets/${encodeURIComponent(asset.assetId)}`,
      {
        method: 'PUT',
        headers: {
          'content-type': asset.mimeType,
          'x-content-sha256': asset.contentHash,
          'x-asset-mime-type': asset.mimeType,
          ...(asset.width === undefined ? {} : { 'x-asset-width': String(asset.width) }),
          ...(asset.height === undefined ? {} : { 'x-asset-height': String(asset.height) }),
        },
        body: bytes,
      },
    );
  }

  async getAsset(unitId: string, assetId: string): Promise<{ bytes: ArrayBuffer; contentHash: string; mimeType: string; byteLength: number }> {
    const response = await this.request(
      `/api/workbooks/${encodeURIComponent(unitId)}/assets/${encodeURIComponent(assetId)}`,
    );
    const contentHash = response.headers.get('x-content-sha256');
    const mimeType = response.headers.get('content-type')?.split(';', 1)[0]?.trim();
    const contentLength = response.headers.get('content-length');
    if (!contentHash || !mimeType || !contentLength) throw new ApiRequestError('Asset response omitted canonical metadata', response.status, 'INTERNAL_ERROR');
    const bytes = await response.arrayBuffer();
    const byteLength = Number(contentLength);
    if (!Number.isSafeInteger(byteLength) || byteLength !== bytes.byteLength) throw new ApiRequestError('Asset response byte length is invalid', response.status, 'INTERNAL_ERROR');
    return { bytes, contentHash, mimeType, byteLength };
  }

  async deleteAsset(unitId: string, assetId: string): Promise<void> {
    await this.request(
      `/api/workbooks/${encodeURIComponent(unitId)}/assets/${encodeURIComponent(assetId)}`,
      { method: 'DELETE' },
    );
  }

  async reconcileAssets(unitId: string, assetIds: readonly string[]): Promise<void> {
    await this.request(`/api/workbooks/${encodeURIComponent(unitId)}/assets/reconcile`, {
      method: 'POST',
      headers: { 'content-type': 'application/json' },
      body: JSON.stringify({ assetIds }),
    });
  }

}

/**
 * WebSocket is a committed-event and ephemeral-presence channel. Durable
 * workbook mutations travel only through WorkbookApiClient.commitOperation.
 * Client presence/cursor messages do not carry actorId; the server adds it to
 * broadcasts after token verification.
 */
export type OperationMessage =
  | { type: 'revision.created'; payload: CommittedOperationEnvelope; revision: number }
  | { type: 'presence.updated'; unitId: string; state: unknown }
  | { type: 'cursor.updated'; unitId: string; state: unknown }
  | { type: 'presence.broadcast'; unitId: string; actorId: string; state: unknown }
  | { type: 'cursor.broadcast'; unitId: string; actorId: string; state: unknown };

export type ClientOperationMessage =
  | { type: 'presence.updated'; unitId: string; state: unknown }
  | { type: 'cursor.updated'; unitId: string; state: unknown };

/** The public collaboration message surface is the single operation contract. */
export type CollaborationMessage = OperationMessage;

export function encodeOperationMessage(message: OperationMessage): string {
  return JSON.stringify(message);
}

export function encodeClientOperationMessage(message: ClientOperationMessage): string {
  return JSON.stringify(message);
}

function validateCommittedOperationEnvelope(value: unknown): CommittedOperationEnvelope {
  if (!value || typeof value !== 'object') throw new Error('Committed operation must be an object');
  const input = value as Record<string, unknown>;
  if (input.schema !== OPERATION_ENVELOPE_SCHEMA || !isNonEmptyString(input.operationId) || !isNonEmptyString(input.unitId)) {
    throw new Error('Invalid committed operation schema');
  }
  if (!isNonEmptyString(input.actorId) || !['client', 'system'].includes(String(input.origin)) || !Number.isSafeInteger(input.revision) || Number(input.revision) < 1) {
    throw new Error('Invalid committed operation metadata');
  }
  if (!isNonEmptyString(input.committedAt) || Number.isNaN(Date.parse(input.committedAt))) {
    throw new Error('Invalid committed operation timestamp');
  }
  const operation = validateOperationEnvelope({
    schema: input.schema,
    operationId: input.operationId,
    unitId: input.unitId,
    clientSequence: input.clientSequence,
    baseRevision: input.baseRevision,
    createdAt: input.createdAt,
    ...(input.intent === undefined ? {} : { intent: input.intent }),
    mutations: Array.isArray(input.mutations)
      ? input.mutations.map((mutation) => {
        if (!mutation || typeof mutation !== 'object') return mutation;
        const candidate = mutation as Record<string, unknown>;
        return {
          id: candidate.id,
          sheetId: candidate.sheetId,
          params: candidate.params,
        };
      })
      : input.mutations,
  });
  const mutations = (input.mutations as unknown[]).map((raw, index) => {
    if (!raw || typeof raw !== 'object') throw new Error(`committed mutation[${index}] must be an object`);
    const mutation = raw as Record<string, unknown>;
    if (!Array.isArray(mutation.affectedRanges)) throw new Error(`committed mutation[${index}] requires affectedRanges`);
    if (!mutation.affectedRanges.every(isRangeRef)) throw new Error(`committed mutation[${index}] contains invalid affectedRanges`);
    return { ...operation.mutations[index]!, affectedRanges: mutation.affectedRanges as RangeRef[] };
  });
  return {
    ...operation,
    actorId: input.actorId,
    origin: input.origin as CommittedOperationEnvelope['origin'],
    revision: Number(input.revision),
    committedAt: input.committedAt,
    mutations,
  };
}

export function decodeOperationMessage(input: string): OperationMessage {
  let raw: unknown;
  try {
    raw = JSON.parse(input);
  } catch {
    throw new Error('Invalid collaboration JSON');
  }
  if (!raw || typeof raw !== 'object' || typeof (raw as Record<string, unknown>).type !== 'string') {
    throw new Error('Invalid collaboration message');
  }
  const message = raw as Record<string, unknown>;
  switch (message.type) {
    case 'revision.created':
      if (!Number.isSafeInteger(message.revision) || Number(message.revision) < 1) throw new Error('revision.created requires revision');
      return {
        type: 'revision.created',
        payload: validateCommittedOperationEnvelope(message.payload),
        revision: Number(message.revision),
      };
    case 'presence.updated':
    case 'cursor.updated':
      if (!isNonEmptyString(message.unitId)) throw new Error(`${message.type} requires unitId`);
      if ('actorId' in message) throw new Error('actorId is server-owned');
      return { type: message.type, unitId: message.unitId, state: message.state };
    case 'presence.broadcast':
    case 'cursor.broadcast':
      if (!isNonEmptyString(message.unitId) || !isNonEmptyString(message.actorId)) {
        throw new Error(`${message.type} requires unitId and actorId`);
      }
      return { type: message.type, unitId: message.unitId, actorId: message.actorId, state: message.state };
    default:
      throw new Error(`Unsupported collaboration message: ${String(message.type)}`);
  }
}

export function decodeClientOperationMessage(input: string): ClientOperationMessage {
  const message = decodeOperationMessage(input);
  if (message.type === 'presence.updated' || message.type === 'cursor.updated') return message;
  throw new Error(`Server-only collaboration message: ${message.type}`);
}

export function encodeMessage(message: OperationMessage): string {
  return encodeOperationMessage(message);
}

export function decodeMessage(input: string): OperationMessage {
  return decodeOperationMessage(input);
}

/** 连接状态 */
export type CollabSocketStatus = 'connecting' | 'open' | 'closed';

export interface CollabSocketOptions {
  /** OIDC source for an authenticated browser handshake. */
  authTokenProvider?: AuthTokenProvider;
  /** Server-issued share token source for an anonymous guest handshake. */
  shareTokenProvider?: ShareTokenProvider;
  /** 断线重连基础延迟(毫秒),指数退避 */
  reconnectBaseDelayMs?: number;
  /** 重连延迟上限(毫秒) */
  reconnectMaxDelayMs?: number;
  /** Injectable constructor for deterministic browser/Node tests. */
  webSocketFactory?: (url: string, protocols: string | string[]) => WebSocket;
}

type CollabSocketListener = (message: OperationMessage) => void;
type StatusListener = (status: CollabSocketStatus) => void;
type ProtocolErrorListener = (error: Error) => void;

const BEARER_SUBPROTOCOL_PREFIX = 'bearer.';

function encodeBase64Url(value: string): string {
  const bytes = new TextEncoder().encode(value);
  let binary = '';
  for (const byte of bytes) binary += String.fromCharCode(byte);
  return btoa(binary).replaceAll('+', '-').replaceAll('/', '_').replace(/=+$/g, '');
}

function withShareToken(url: string, shareToken: string): string {
  const target = new URL(url, typeof window === 'undefined' ? 'ws://localhost' : window.location.origin);
  target.searchParams.set('shareToken', shareToken);
  return target.toString();
}

/** Browser-safe bearer token transport for a WebSocket handshake. */
export function createBearerSubprotocol(token: string): string {
  const normalized = token.trim();
  if (!normalized) throw new AuthenticationRequiredError();
  return `${BEARER_SUBPROTOCOL_PREFIX}${encodeBase64Url(normalized)}`;
}

/**
 * 浏览器端协同 WebSocket 客户端。
 * - 断线自动指数退避重连;
 * - 断线期间 send() 的消息进入待发队列,连接恢复后按序冲刷;
 * - 消息格式复用 encode/decodeMessage 协议。
 */
export class CollabSocketClient {
  private socket: WebSocket | null = null;
  private readonly messageListeners = new Set<CollabSocketListener>();
  private readonly statusListeners = new Set<StatusListener>();
  private readonly protocolErrorListeners = new Set<ProtocolErrorListener>();
  private reconnectAttempt = 0;
  private reconnectTimer: ReturnType<typeof setTimeout> | null = null;
  private closedByUser = false;
  private connecting = false;

  constructor(
    private readonly url: string,
    private readonly options: CollabSocketOptions,
  ) {}

  get status(): CollabSocketStatus {
    if (this.socket?.readyState === 1) return 'open';
    if (this.closedByUser) return 'closed';
    return this.socket || this.reconnectTimer || this.connecting ? 'connecting' : 'closed';
  }

  open(): void {
    this.closedByUser = false;
    if (this.socket || this.reconnectTimer) return;
    this.connect();
  }

  close(): void {
    this.closedByUser = true;
    if (this.reconnectTimer) {
      clearTimeout(this.reconnectTimer);
      this.reconnectTimer = null;
    }
    this.socket?.close();
    this.socket = null;
    this.emitStatus('closed');
  }

  /** 发送客户端消息；未连接时返回 false，由 OfflineQueue 保留 operation。 */
  send(message: ClientOperationMessage): boolean {
    const encoded = encodeClientOperationMessage(message);
    if (this.socket?.readyState === 1) {
      this.socket.send(encoded);
      return true;
    }
    return false;
  }

  onMessage(listener: CollabSocketListener): () => void {
    this.messageListeners.add(listener);
    return () => this.messageListeners.delete(listener);
  }

  onStatus(listener: StatusListener): () => void {
    this.statusListeners.add(listener);
    listener(this.status);
    return () => this.statusListeners.delete(listener);
  }

  onProtocolError(listener: ProtocolErrorListener): () => void {
    this.protocolErrorListeners.add(listener);
    return () => this.protocolErrorListeners.delete(listener);
  }

  private async connect(): Promise<void> {
    this.connecting = true;
    this.emitStatus('connecting');
    let token: string | null;
    let shareToken: string | null;
    try {
      token = (await this.options.authTokenProvider?.())?.trim() || null;
      shareToken = token ? null : (await this.options.shareTokenProvider?.())?.trim() || null;
    } catch (cause) {
      this.connecting = false;
      this.failClosed(cause instanceof Error ? cause : new Error('Unable to resolve bearer token'));
      return;
    }
    if (!token && !shareToken) {
      this.connecting = false;
      this.failClosed(new AuthenticationRequiredError());
      return;
    }
    if (this.closedByUser) {
      this.connecting = false;
      return;
    }
    const factory = this.options.webSocketFactory ?? ((target: string, protocols: string | string[]) => new WebSocket(target, protocols));
    const socket = token
      ? factory(this.url, createBearerSubprotocol(token))
      : factory(withShareToken(this.url, shareToken!), []);
    this.socket = socket;
    this.connecting = false;

    socket.onopen = () => {
      this.reconnectAttempt = 0;
      this.emitStatus('open');
    };

    socket.onmessage = (event) => {
      if (typeof event.data !== 'string') {
        this.failClosed(new Error('Collaboration server sent a non-text message'));
        return;
      }
      try {
        const message = decodeOperationMessage(event.data);
        for (const listener of this.messageListeners) listener(message);
      } catch (cause) {
        this.failClosed(cause instanceof Error ? cause : new Error('Invalid collaboration server message'));
      }
    };

    socket.onclose = () => {
      this.socket = null;
      if (this.closedByUser) {
        this.emitStatus('closed');
        return;
      }
      this.scheduleReconnect();
    };

    socket.onerror = () => {
      socket.close();
    };
  }

  private scheduleReconnect(): void {
    if (this.reconnectTimer || this.closedByUser) return;
    const base = this.options.reconnectBaseDelayMs ?? 800;
    const max = this.options.reconnectMaxDelayMs ?? 15000;
    const delay = Math.min(max, base * Math.pow(2, this.reconnectAttempt));
    this.reconnectAttempt += 1;
    this.emitStatus('connecting');
    this.reconnectTimer = setTimeout(() => {
      this.reconnectTimer = null;
      this.connect();
    }, delay);
  }

  private emitStatus(status: CollabSocketStatus): void {
    for (const listener of this.statusListeners) listener(status);
  }

  private failClosed(error: Error): void {
    this.closedByUser = true;
    if (this.reconnectTimer) {
      clearTimeout(this.reconnectTimer);
      this.reconnectTimer = null;
    }
    for (const listener of this.protocolErrorListeners) listener(error);
    this.socket?.close(1002, error.message.slice(0, 120));
    this.socket = null;
    this.connecting = false;
    this.emitStatus('closed');
  }
}
