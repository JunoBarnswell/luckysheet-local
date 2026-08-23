import {
  DEFAULT_DATA_BLOCK_ROW_COUNT,
  normalizeDataSourceManifest,
  type DataBlockRef,
  type DataBlockAvailability,
  type DataSourceField,
  type DataSourceManifest,
  type TableScalar,
} from '@react-sheets/core-model';
import {
  decodeColumnarBlock,
  validateSparseCellOverlay,
  type SparseCellOverlay,
} from './codec';

export type DataSourceFieldRef = string | number;

export interface DataBlockReader {
  get(ref: Pick<DataBlockRef, 'dataSourceId' | 'id' | 'checksum'>): Promise<{
    sourceId: string;
    blockId: string;
    checksum: string;
    bytes: ArrayBuffer;
  } | null>;
}

export interface DataSourceContentQueryOptions {
  /** Each overlay uses block-local row and field coordinates. */
  overlays?: ReadonlyMap<string, SparseCellOverlay>;
}

export interface DataSourceContentLoadState {
  sourceId: string;
  blockId: string | null;
  availability: DataBlockAvailability;
  error?: string;
}

export interface DataSourceContentResult<T> {
  state: DataSourceContentLoadState;
  value?: T;
}

export type DataSourceContentStateListener = (state: DataSourceContentLoadState) => void;

interface LoadedBlock {
  ref: DataBlockRef;
  rows: TableScalar[][];
}

class ContentQueryFailure extends Error {
  constructor(
    readonly availability: Extract<DataBlockAvailability, 'missing' | 'error'>,
    message: string,
  ) {
    super(message);
    this.name = 'DataSourceContentQueryError';
  }
}

function errorMessage(error: unknown): string {
  return error instanceof Error ? error.message : String(error);
}

function cloneOverlay(overlay: SparseCellOverlay): SparseCellOverlay {
  return {
    schema: overlay.schema,
    revision: overlay.revision,
    cells: overlay.cells.map((cell) => ({ ...cell })),
  };
}

function cloneField(field: DataSourceField): DataSourceField {
  return { ...field };
}

function cloneRows(rows: readonly (readonly TableScalar[])[]): TableScalar[][] {
  return rows.map((row) => [...row]);
}

function state(
  sourceId: string,
  blockId: string | null,
  availability: DataBlockAvailability,
  error?: string,
): DataSourceContentLoadState {
  return error === undefined
    ? { sourceId, blockId, availability }
    : { sourceId, blockId, availability, error };
}

function isSafeRowIndex(value: number): boolean {
  return Number.isSafeInteger(value) && value >= 0;
}

/**
 * Asynchronous content access for a block-backed data source. It is the
 * read-side boundary used by renderers, formulas, and Pivot computation:
 * metadata remains in the manifest, bytes remain in LocalDataBlockStore, and
 * sparse edits are applied only to the returned projection.
 */
export class DataSourceContentQuery {
  private readonly source: DataSourceManifest;
  private readonly store: DataBlockReader;
  private readonly overlays: ReadonlyMap<string, SparseCellOverlay>;
  private readonly loadStates = new Map<string, DataSourceContentLoadState>();
  private readonly loadPromises = new Map<string, Promise<LoadedBlock>>();
  private readonly loadedBlocks = new Map<string, LoadedBlock>();
  private readonly listeners = new Set<DataSourceContentStateListener>();

  constructor(
    manifest: DataSourceManifest,
    store: DataBlockReader,
    options: DataSourceContentQueryOptions = {},
  ) {
    const normalized = normalizeDataSourceManifest(structuredClone(manifest));
    if (normalized.blockRowCount !== DEFAULT_DATA_BLOCK_ROW_COUNT) {
      throw new Error(`Data source blockRowCount must be ${String(DEFAULT_DATA_BLOCK_ROW_COUNT)}`);
    }
    for (const block of normalized.blocks) {
      if (block.rowCount > normalized.blockRowCount) {
        throw new Error(`Data block exceeds the configured row block size: ${block.id}`);
      }
      if (block.startRow + block.rowCount > normalized.rowCount) {
        throw new Error(`Data block exceeds the source rowCount: ${block.id}`);
      }
    }
    this.source = normalized;
    this.store = store;
    const overlayMap = new Map<string, SparseCellOverlay>();
    for (const [blockId, inputOverlay] of options.overlays ?? new Map<string, SparseCellOverlay>()) {
      const block = this.source.blocks.find((entry) => entry.id === blockId);
      if (!block) throw new Error(`Sparse overlay targets unknown data block: ${blockId}`);
      validateSparseCellOverlay(inputOverlay, {
        rowCount: block.rowCount,
        columnCount: this.source.fields.length,
      });
      overlayMap.set(blockId, cloneOverlay(inputOverlay));
    }
    this.overlays = overlayMap;
  }

  get manifest(): DataSourceManifest {
    return {
      ...this.source,
      fields: this.source.fields.map(cloneField),
      blocks: this.source.blocks.map((block) => ({ ...block })),
    };
  }

  getField(fieldRef: DataSourceFieldRef): DataSourceField | undefined {
    const field = typeof fieldRef === 'number'
      ? this.source.fields[fieldRef]
      : this.source.fields.find((entry) => entry.id === fieldRef);
    return field === undefined ? undefined : cloneField(field);
  }

  getLoadState(blockId: string): DataSourceContentLoadState | undefined {
    const current = this.loadStates.get(blockId);
    return current === undefined ? undefined : { ...current };
  }

  getLoadStates(): DataSourceContentLoadState[] {
    return [...this.loadStates.values()].map((current) => ({ ...current }));
  }

  subscribe(listener: DataSourceContentStateListener): () => void {
    this.listeners.add(listener);
    return () => this.listeners.delete(listener);
  }

  async getCellValue(
    rowIndex: number,
    fieldRef: DataSourceFieldRef,
  ): Promise<DataSourceContentResult<TableScalar>> {
    const field = this.resolveField(fieldRef);
    if (field === undefined) return this.errorResult(`Unknown data source field: ${String(fieldRef)}`);
    const row = await this.getRowValues(rowIndex);
    if (row.value === undefined) return { state: row.state };
    return { state: row.state, value: row.value[field.ordinal] ?? null };
  }

  async getRowValues(rowIndex: number): Promise<DataSourceContentResult<TableScalar[]>> {
    const result = await this.getRows(rowIndex, 1);
    if (result.value === undefined) return { state: result.state };
    return { state: result.state, value: result.value[0] ?? [] };
  }

  async getFieldValues(
    fieldRef: DataSourceFieldRef,
    startRow: number,
    rowCount: number,
  ): Promise<DataSourceContentResult<TableScalar[]>> {
    const field = this.resolveField(fieldRef);
    if (field === undefined) return this.errorResult(`Unknown data source field: ${String(fieldRef)}`);
    const rows = await this.getRows(startRow, rowCount);
    if (rows.value === undefined) return { state: rows.state };
    return {
      state: rows.state,
      value: rows.value.map((row) => row[field.ordinal] ?? null),
    };
  }

  async getRows(
    startRow: number,
    rowCount: number,
  ): Promise<DataSourceContentResult<TableScalar[][]>> {
    const rangeError = this.validateRange(startRow, rowCount);
    if (rangeError !== undefined) return this.errorResult(rangeError);
    if (rowCount === 0) {
      return {
        state: state(this.source.id, null, 'ready'),
        value: [],
      };
    }

    const refs: DataBlockRef[] = [];
    const seen = new Set<string>();
    for (let row = startRow; row < startRow + rowCount; row += 1) {
      const ref = this.findBlock(row);
      if (ref === undefined) {
        return this.missingResult(`No data block covers source row ${String(row)}`);
      }
      if (!seen.has(ref.id)) {
        seen.add(ref.id);
        refs.push(ref);
      }
    }

    const loaded = new Map<string, LoadedBlock>();
    const outcomes = await Promise.all(refs.map(async (ref) => {
      try {
        return { ref, block: await this.loadBlock(ref) } as const;
      } catch (error) {
        return { ref, error } as const;
      }
    }));
    for (const outcome of outcomes) {
      if ('error' in outcome) {
        const current = this.loadStates.get(outcome.ref.id)
          ?? state(this.source.id, outcome.ref.id, 'error', errorMessage(outcome.error));
        return { state: { ...current } };
      }
      loaded.set(outcome.ref.id, outcome.block);
    }

    const rows: TableScalar[][] = [];
    for (let row = startRow; row < startRow + rowCount; row += 1) {
      const ref = this.findBlock(row)!;
      const block = loaded.get(ref.id)!;
      const localRow = row - ref.startRow;
      const values = block.rows[localRow];
      if (values === undefined) {
        return this.errorResult(`Data block ${ref.id} does not contain source row ${String(row)}`);
      }
      rows.push([...values]);
    }
    return {
      state: state(this.source.id, refs.length === 1 ? refs[0]!.id : null, 'ready'),
      value: rows,
    };
  }

  private resolveField(fieldRef: DataSourceFieldRef): DataSourceField | undefined {
    if (typeof fieldRef === 'number') {
      if (!Number.isSafeInteger(fieldRef) || fieldRef < 0) return undefined;
      return this.source.fields[fieldRef];
    }
    if (fieldRef.trim().length === 0) return undefined;
    return this.source.fields.find((field) => field.id === fieldRef);
  }

  private validateRange(startRow: number, rowCount: number): string | undefined {
    if (!isSafeRowIndex(startRow)) return 'Data source startRow must be a non-negative safe integer';
    if (!Number.isSafeInteger(rowCount) || rowCount < 0) return 'Data source rowCount must be a non-negative safe integer';
    if (startRow + rowCount > this.source.rowCount) return 'Data source query range exceeds rowCount';
    return undefined;
  }

  private findBlock(rowIndex: number): DataBlockRef | undefined {
    let low = 0;
    let high = this.source.blocks.length - 1;
    while (low <= high) {
      const middle = Math.floor((low + high) / 2);
      const block = this.source.blocks[middle]!;
      if (rowIndex < block.startRow) {
        high = middle - 1;
      } else if (rowIndex >= block.startRow + block.rowCount) {
        low = middle + 1;
      } else {
        return block;
      }
    }
    return undefined;
  }

  private async loadBlock(ref: DataBlockRef): Promise<LoadedBlock> {
    const cached = this.loadedBlocks.get(ref.id);
    if (cached !== undefined) return cached;
    const existing = this.loadPromises.get(ref.id);
    if (existing !== undefined) return existing;

    this.publishState(state(this.source.id, ref.id, 'loading'));
    const promise = this.readBlock(ref).then((block) => {
      this.loadedBlocks.set(ref.id, block);
      this.publishState(state(this.source.id, ref.id, 'ready'));
      return block;
    }).catch((error: unknown) => {
      const failure = error instanceof ContentQueryFailure
        ? error
        : new ContentQueryFailure('error', errorMessage(error));
      this.publishState(state(this.source.id, ref.id, failure.availability, failure.message));
      throw failure;
    }).finally(() => {
      if (!this.loadedBlocks.has(ref.id)) this.loadPromises.delete(ref.id);
    });
    this.loadPromises.set(ref.id, promise);
    return promise;
  }

  private async readBlock(ref: DataBlockRef): Promise<LoadedBlock> {
    if (ref.encoding !== 'columnar-v1') {
      throw new ContentQueryFailure('error', `Unsupported data block encoding: ${ref.encoding}`);
    }
    let record: Awaited<ReturnType<DataBlockReader['get']>>;
    try {
      record = await this.store.get(ref);
    } catch (error) {
      throw new ContentQueryFailure('error', `Data block ${ref.id} could not be loaded: ${errorMessage(error)}`);
    }
    if (record === null) {
      throw new ContentQueryFailure('missing', `Data block ${ref.id} is missing from local storage`);
    }
    if (record.sourceId !== this.source.id || record.blockId !== ref.id) {
      throw new ContentQueryFailure('error', `Data block ${ref.id} has an invalid storage identity`);
    }
    if (record.checksum !== ref.checksum) {
      throw new ContentQueryFailure('error', `Data block ${ref.id} checksum does not match the manifest`);
    }
    if (!(record.bytes instanceof ArrayBuffer) || record.bytes.byteLength !== ref.byteLength) {
      throw new ContentQueryFailure('error', `Data block ${ref.id} byteLength does not match the manifest`);
    }
    try {
      const decoded = await decodeColumnarBlock(record.bytes, {
        expectedRowCount: ref.rowCount,
        expectedFields: this.source.fields,
        expectedChecksum: ref.checksum,
      });
      const rows = cloneRows(decoded.rows);
      const overlay = this.overlays.get(ref.id);
      if (overlay !== undefined) {
        for (const cell of overlay.cells) rows[cell.row]![cell.column] = cell.value;
      }
      return { ref, rows };
    } catch (error) {
      throw new ContentQueryFailure('error', `Data block ${ref.id} failed validation: ${errorMessage(error)}`);
    }
  }

  private publishState(next: DataSourceContentLoadState): void {
    if (next.blockId !== null) this.loadStates.set(next.blockId, next);
    for (const listener of this.listeners) {
      try {
        listener({ ...next });
      } catch {
        // A consumer must not break the query state machine.
      }
    }
  }

  private errorResult(message: string): DataSourceContentResult<never> {
    return { state: state(this.source.id, null, 'error', message) };
  }

  private missingResult(message: string): DataSourceContentResult<never> {
    return { state: state(this.source.id, null, 'missing', message) };
  }
}
