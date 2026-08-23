import type {
  DataSourceField,
  DataSourceFieldType,
  PivotDefinition,
  PivotFieldDataType,
  PivotScalar,
  PivotSourceRowPath,
  SheetId,
} from '@react-sheets/core-model';
import type { DataSourceContentLoadState, DataSourceContentQuery } from '../data-source/content-query';

export interface PivotBlockSourceField {
  fieldId: string;
  name: string;
  ordinal: number;
  dataType: PivotFieldDataType;
}

export interface PivotBlockSourceRow {
  values: Record<string, PivotScalar>;
  paths: PivotSourceRowPath[];
}

export interface PivotBlockSourceTable {
  fields: PivotBlockSourceField[];
  rows: PivotBlockSourceRow[];
}

export type PivotBlockSourceStatus = 'loading' | 'ready' | 'missing' | 'error';

export interface PivotBlockSourceState {
  status: PivotBlockSourceStatus;
  sourceId: string;
  blockId: string | null;
  error?: string;
}

export interface PivotBlockSourceReadOptions {
  /** Physical worksheet id used by Show Details source row paths. */
  sourceSheetId?: SheetId;
  /** Physical first data row; callers that store a header pass header row + 1. */
  sourceRowStart?: number;
  /** Bounded query size; the default aligns with the block row contract. */
  chunkRowCount?: number;
  onState?: (state: PivotBlockSourceState) => void;
}

export type PivotBlockSourceReadResult =
  | {
    status: 'ready';
    state: PivotBlockSourceState;
    source: PivotBlockSourceTable;
    sourceRevision: number;
  }
  | {
    status: 'loading' | 'missing' | 'error';
    state: PivotBlockSourceState;
    error: string;
  };

function stateFromQuery(next: DataSourceContentLoadState): PivotBlockSourceState {
  return next.error === undefined
    ? { status: next.availability, sourceId: next.sourceId, blockId: next.blockId }
    : { status: next.availability, sourceId: next.sourceId, blockId: next.blockId, error: next.error };
}

function failure(
  status: Exclude<PivotBlockSourceStatus, 'ready'>,
  sourceId: string,
  error: string,
  blockId: string | null = null,
): PivotBlockSourceReadResult {
  const state: PivotBlockSourceState = { status, sourceId, blockId, error };
  return { status, state, error };
}

function fieldType(type: DataSourceFieldType): PivotFieldDataType {
  return type;
}

function validateOptions(options: PivotBlockSourceReadOptions): string | undefined {
  if (options.sourceSheetId !== undefined && !options.sourceSheetId.trim()) return 'sourceSheetId cannot be empty';
  if (options.sourceRowStart !== undefined
    && (!Number.isSafeInteger(options.sourceRowStart) || options.sourceRowStart < 0)) {
    return 'sourceRowStart must be a non-negative safe integer';
  }
  if (options.chunkRowCount !== undefined
    && (!Number.isSafeInteger(options.chunkRowCount) || options.chunkRowCount <= 0)) {
    return 'chunkRowCount must be a positive safe integer';
  }
  return undefined;
}

function canonicalFields(fields: readonly DataSourceField[]): PivotBlockSourceField[] {
  const ids = new Set<string>();
  return fields.map((field, ordinal) => {
    if (!field.id.trim()) throw new Error(`Data source field ${String(ordinal)} has no stable fieldId`);
    if (ids.has(field.id)) throw new Error(`Data source contains duplicate fieldId ${field.id}`);
    if (field.ordinal !== ordinal) throw new Error(`Data source field ${field.id} has a non-contiguous ordinal`);
    ids.add(field.id);
    return {
      fieldId: field.id,
      name: field.name,
      ordinal,
      dataType: fieldType(field.type),
    };
  });
}

function sourceRevision(query: DataSourceContentQuery): number {
  const manifest = query.manifest;
  return manifest.revision;
}

function rowPath(
  sheetId: SheetId,
  rowStart: number,
  rowIndex: number,
): PivotSourceRowPath {
  const row = rowStart + rowIndex;
  if (!Number.isSafeInteger(row) || row < 0) throw new Error(`Source row path exceeds the safe row range: ${String(row)}`);
  return { sheetId, row };
}

/**
 * Read a canonical data-source Pivot source through the block content query.
 * No empty source is returned for load failures: callers receive an explicit
 * loading, missing, or error result and can keep the last valid projection.
 */
export async function readPivotBlockSource(
  pivot: PivotDefinition,
  query: DataSourceContentQuery,
  options: PivotBlockSourceReadOptions = {},
): Promise<PivotBlockSourceReadResult> {
  const queryManifest = query.manifest;
  const sourceId = pivot.source.kind === 'data-source' ? pivot.source.dataSourceId : queryManifest.id;
  const optionError = validateOptions(options);
  if (optionError !== undefined) return failure('error', sourceId, optionError);
  if (pivot.source.kind !== 'data-source') {
    return failure('error', sourceId, 'Pivot source is not a canonical data-source source');
  }
  if (queryManifest.id !== pivot.source.dataSourceId) {
    return failure('error', sourceId, `Data source query ${queryManifest.id} does not match Pivot source ${pivot.source.dataSourceId}`);
  }

  const sourceSheetId = options.sourceSheetId ?? queryManifest.sourceSheetId;
  if (sourceSheetId === undefined || !sourceSheetId.trim()) {
    return failure('error', sourceId, 'Data source has no worksheet identity for source row paths');
  }
  const sourceRowStart = options.sourceRowStart ?? 0;
  const chunkRowCount = options.chunkRowCount ?? 65_536;
  let fields: PivotBlockSourceField[];
  try {
    fields = canonicalFields(queryManifest.fields);
  } catch (error) {
    return failure('error', sourceId, error instanceof Error ? error.message : String(error));
  }

  const states: PivotBlockSourceState[] = [];
  const unsubscribe = options.onState
    ? query.subscribe((next) => {
      const state = stateFromQuery(next);
      states.push(state);
      options.onState!(state);
    })
    : undefined;
  try {
    const rows: PivotBlockSourceRow[] = [];
    for (let startRow = 0; startRow < queryManifest.rowCount; startRow += chunkRowCount) {
      const rowCount = Math.min(chunkRowCount, queryManifest.rowCount - startRow);
      const result = await query.getRows(startRow, rowCount);
      const state = stateFromQuery(result.state);
      if (result.value === undefined) {
        const status = state.status === 'ready' ? 'error' : state.status;
        return failure(status, state.sourceId, state.error ?? `Data source ${sourceId} did not return rows`, state.blockId);
      }
      if (state.status !== 'ready') {
        return failure(state.status, state.sourceId, state.error ?? `Data source ${sourceId} is ${state.status}`, state.blockId);
      }
      for (let localRow = 0; localRow < result.value.length; localRow += 1) {
        const values = result.value[localRow]!;
        if (values.length !== fields.length) {
          return failure('error', sourceId, `Data source row ${String(startRow + localRow)} has ${String(values.length)} fields; expected ${String(fields.length)}`, state.blockId);
        }
        const rowValues: Record<string, PivotScalar> = {};
        fields.forEach((field, ordinal) => {
          rowValues[field.fieldId] = values[ordinal] ?? null;
        });
        rows.push({
          values: rowValues,
          paths: [rowPath(sourceSheetId, sourceRowStart, startRow + localRow)],
        });
      }
    }
    const readyState: PivotBlockSourceState = {
      status: 'ready',
      sourceId,
      blockId: states.at(-1)?.blockId ?? null,
    };
    return {
      status: 'ready',
      state: readyState,
      source: { fields, rows },
      sourceRevision: sourceRevision(query),
    };
  } catch (error) {
    const message = error instanceof Error ? error.message : String(error);
    const current = states.at(-1);
    return failure(current?.status === 'missing' ? 'missing' : 'error', sourceId, message, current?.blockId ?? null);
  } finally {
    unsubscribe?.();
  }
}
