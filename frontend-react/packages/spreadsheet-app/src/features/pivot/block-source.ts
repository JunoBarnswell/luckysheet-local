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
import { createPivotSourceIndex, type PivotSourceIndex } from './source-index';

export interface PivotBlockSourceField {
  fieldId: string;
  name: string;
  ordinal: number;
  dataType: PivotFieldDataType;
}

export type PivotBlockSourceTable = PivotSourceIndex;

export type PivotBlockSourceStatus = 'loading' | 'ready' | 'missing' | 'error';

export interface PivotBlockSourceState {
  status: PivotBlockSourceStatus;
  sourceId: string;
  blockId: string | null;
  error?: string;
}

export interface PivotBlockSourceCellOverlay {
  /** Immutable physical row index in the backing data source. */
  rowIndex: number;
  fieldOrdinal: number;
  value: PivotScalar;
}

export interface PivotBlockSourceReadOptions {
  /** Physical worksheet id used by Show Details source row paths. */
  sourceSheetId?: SheetId;
  /** Physical first data row; callers that store a header pass header row + 1. */
  sourceRowStart?: number;
  /** Sparse canonical CellPatch values, resolved after all source blocks load. */
  resolveCellOverlays?: () => readonly PivotBlockSourceCellOverlay[];
  /** Cancels block acquisition before the source index is published. */
  signal?: AbortSignal;
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
  return undefined;
}

function assertBlockReadActive(signal?: AbortSignal): void {
  if (!signal?.aborted) return;
  const error = new Error('Pivot source block read was cancelled');
  error.name = 'AbortError';
  throw error;
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
  assertBlockReadActive(options.signal);
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
  let fields: PivotBlockSourceField[];
  try {
    fields = canonicalFields(queryManifest.fields);
  } catch (error) {
    return failure('error', sourceId, error instanceof Error ? error.message : String(error));
  }

  const manifestIdentity = JSON.stringify(queryManifest);
  let latestState: PivotBlockSourceState | undefined;
  const unsubscribe = query.subscribe((next) => {
    const state = stateFromQuery(next);
    latestState = state;
    options.onState?.(state);
  });
  try {
    assertBlockReadActive(options.signal);
    const columnValues = fields.map(() => [] as PivotScalar[]);
    let scannedRowCount = 0;
    const scanned = await query.scanRows((values, logicalRow) => {
      assertBlockReadActive(options.signal);
      if (values.length !== fields.length) {
        throw new Error(`Data source row ${String(logicalRow)} has ${String(values.length)} fields; expected ${String(fields.length)}`);
      }
      const physicalRow = query.getPhysicalRow(logicalRow);
      if (physicalRow === undefined) throw new Error(`Data source logical row ${String(logicalRow)} has no physical source row`);
      scannedRowCount += 1;
      for (let ordinal = 0; ordinal < fields.length; ordinal += 1) {
        columnValues[ordinal]!.push(values[ordinal] ?? null);
      }
    });
    assertBlockReadActive(options.signal);
    const scanState = stateFromQuery(scanned.state);
    if (scanned.value === undefined || scanState.status !== 'ready') {
      const status = scanState.status === 'ready' ? 'error' : scanState.status;
      return failure(status, scanState.sourceId, scanState.error ?? `Data source ${sourceId} did not return rows`, scanState.blockId);
    }
    if (JSON.stringify(query.manifest) !== manifestIdentity) {
      return failure('error', sourceId, 'Pivot data source changed while block rows were loading');
    }
    if (scannedRowCount !== queryManifest.rowCount) {
      return failure('error', sourceId, 'Pivot data source row count changed while block rows were loading');
    }
    const overlays = options.resolveCellOverlays?.() ?? [];
    const overlayCells = new Set<string>();
    if (overlays.length > 0) {
      const logicalByPhysicalRow = new Map<number, number>();
      for (let logicalRow = 0; logicalRow < queryManifest.rowCount; logicalRow += 1) {
        const physicalRow = query.getPhysicalRow(logicalRow);
        if (physicalRow === undefined) throw new Error(`Data source logical row ${String(logicalRow)} has no physical source row`);
        if (logicalByPhysicalRow.has(physicalRow)) throw new Error(`Data source row order maps multiple logical rows to physical row ${String(physicalRow)}`);
        logicalByPhysicalRow.set(physicalRow, logicalRow);
      }
      for (const overlay of overlays) {
        assertBlockReadActive(options.signal);
        if (!Number.isSafeInteger(overlay.rowIndex) || overlay.rowIndex < 0 || overlay.rowIndex >= queryManifest.rowCount
          || !Number.isSafeInteger(overlay.fieldOrdinal) || overlay.fieldOrdinal < 0 || overlay.fieldOrdinal >= fields.length) {
          throw new Error('Pivot source CellPatch overlay is outside the canonical data region');
        }
        const logicalRow = logicalByPhysicalRow.get(overlay.rowIndex);
        if (logicalRow === undefined) throw new Error(`Pivot source physical row ${String(overlay.rowIndex)} is absent from the current row order`);
        const cellKey = `${String(overlay.rowIndex)}:${String(overlay.fieldOrdinal)}`;
        if (overlayCells.has(cellKey)) throw new Error('Pivot source has duplicate CellPatch overlays for one cell');
        overlayCells.add(cellKey);
        columnValues[overlay.fieldOrdinal]![logicalRow] = overlay.value;
      }
    }
    assertBlockReadActive(options.signal);
    const readyState: PivotBlockSourceState = {
      status: 'ready',
      sourceId,
      blockId: scanState.blockId ?? latestState?.blockId ?? null,
    };
    return {
      status: 'ready',
      state: readyState,
      source: createPivotSourceIndex({
        columns: fields.map((field, ordinal) => ({ field, values: columnValues[ordinal]! })),
        rowCount: queryManifest.rowCount,
        rowPathAt: (logicalRow) => {
          const physicalRow = query.getPhysicalRow(logicalRow);
          if (physicalRow === undefined) throw new Error(`Data source logical row ${String(logicalRow)} has no physical source row`);
          return [rowPath(sourceSheetId, sourceRowStart, physicalRow)];
        },
      }),
      sourceRevision: sourceRevision(query),
    };
  } catch (error) {
    if (options.signal?.aborted || (error instanceof Error && error.name === 'AbortError')) throw error;
    const message = error instanceof Error ? error.message : String(error);
    const current = latestState;
    return failure(current?.status === 'missing' ? 'missing' : 'error', sourceId, message, current?.blockId ?? null);
  } finally {
    unsubscribe?.();
  }
}
