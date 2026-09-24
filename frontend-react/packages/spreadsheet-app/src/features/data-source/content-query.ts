import {
  DEFAULT_DATA_BLOCK_ROW_COUNT,
  PIVOT_MEMBER_DISPLAY_LIMIT,
  normalizeDataSourceManifest,
  type DataBlockRef,
  type DataBlockAvailability,
  type DataSourceField,
  type DataSourceManifest,
  type PivotScalar,
  type TableScalar,
} from '@react-sheets/core-model';
import {
  decodeOwnedColumnarBlock,
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

export interface DataSourceScanOptions {
  /** Stop waiting for block reads when the consumer has been superseded. */
  signal?: AbortSignal;
  /** Load the complete immutable block set before visiting the first row. */
  prefetchAllBlocks?: boolean;
}

/** Read-only views over decoded blocks. Row arrays are shared with the cache. */
export interface DataSourceLoadedBlockView {
  readonly ref: DataBlockRef;
  readonly rows: readonly (readonly TableScalar[])[];
}

export type DataSourceContentStateListener = (state: DataSourceContentLoadState) => void;

interface LoadedBlock {
  ref: DataBlockRef;
  rows: TableScalar[][];
}

const MAX_CONCURRENT_BLOCK_READS = 4;

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

function abortError(): Error {
  const error = new Error('Data source scan was cancelled');
  error.name = 'AbortError';
  return error;
}

function assertNotAborted(signal?: AbortSignal): void {
  if (signal?.aborted) throw abortError();
}

function awaitWithoutCancelling<T>(promise: Promise<T>, signal?: AbortSignal): Promise<T> {
  if (signal === undefined) return promise;
  assertNotAborted(signal);
  return new Promise<T>((resolve, reject) => {
    let settled = false;
    let cleanup = (): void => undefined;
    const onAbort = (): void => {
      if (settled) return;
      settled = true;
      cleanup();
      reject(abortError());
    };
    cleanup = (): void => signal.removeEventListener('abort', onAbort);
    signal.addEventListener('abort', onAbort, { once: true });
    promise.then(
      (value) => {
        if (settled) return;
        settled = true;
        cleanup();
        resolve(value);
      },
      (error: unknown) => {
        if (settled) return;
        settled = true;
        cleanup();
        reject(error);
      },
    );
  });
}

/**
 * Asynchronous content access for a block-backed data source. It is the
 * read-side boundary used by renderers, formulas, and Pivot computation:
 * metadata remains in the manifest, bytes remain in LocalDataBlockStore, and
 * sparse edits are applied only to the returned projection.
 */
export class DataSourceContentQuery {
  private source: DataSourceManifest;
  private readonly store: DataBlockReader;
  private readonly overlays: ReadonlyMap<string, SparseCellOverlay>;
  private readonly loadStates = new Map<string, DataSourceContentLoadState>();
  private readonly loadPromises = new Map<string, Promise<LoadedBlock>>();
  private readonly loadedBlocks = new Map<string, LoadedBlock>();
  private readonly blockRefsById = new Map<string, DataBlockRef>();
  private readonly blockIndexesById = new Map<string, number>();
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
      if (block.startRow > normalized.rowCount || block.rowCount > normalized.rowCount - block.startRow) {
        throw new Error(`Data block exceeds the source rowCount: ${block.id}`);
      }
    }
    this.source = normalized;
    this.store = store;
    normalized.blocks.forEach((block, index) => {
      this.blockRefsById.set(block.id, block);
      this.blockIndexesById.set(block.id, index);
    });
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
      ...(this.source.rowOrder === undefined ? {} : { rowOrder: [...this.source.rowOrder] }),
      ...(this.source.sortState === undefined ? {} : {
        sortState: { criteria: this.source.sortState.criteria.map((criterion) => ({ ...criterion })) },
      }),
    };
  }

  /** Preserve decoded immutable blocks when only logical source metadata changed. */
  rebindManifest(manifest: DataSourceManifest): boolean {
    const normalized = normalizeDataSourceManifest(structuredClone(manifest));
    // This identity owns decoded bytes only. Projection metadata, revision,
    // virtual sort order and display naming can change without invalidating a
    // verified immutable block, so retaining the cache is both safe and
    // necessary for lazy multi-sheet navigation.
    const contentIdentity = (source: DataSourceManifest): string => JSON.stringify({
      id: source.id,
      rowCount: source.rowCount,
      fields: source.fields,
      blockRowCount: source.blockRowCount,
      blocks: source.blocks.map(({ id, dataSourceId, startRow, rowCount, checksum, byteLength, encoding }) => ({
        id, dataSourceId, startRow, rowCount, checksum, byteLength, encoding,
      })),
    });
    if (contentIdentity(normalized) !== contentIdentity(this.source)) return false;
    for (const [blockId, loaded] of this.loadedBlocks) {
      const ref = normalized.blocks.find((candidate) => candidate.id === blockId);
      if (ref === undefined) this.loadedBlocks.delete(blockId);
      else this.loadedBlocks.set(blockId, { ...loaded, ref });
    }
    for (const [blockId, current] of this.loadStates) {
      this.loadStates.set(blockId, { ...current, sourceId: normalized.id });
    }
    this.blockRefsById.clear();
    this.blockIndexesById.clear();
    normalized.blocks.forEach((ref, index) => {
      this.blockRefsById.set(ref.id, ref);
      this.blockIndexesById.set(ref.id, index);
    });
    this.source = normalized;
    return true;
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

  /** Resolve the immutable block row behind a logical (possibly sorted) row. */
  getPhysicalRow(logicalRow: number): number | undefined {
    if (!isSafeRowIndex(logicalRow) || logicalRow >= this.source.rowCount) return undefined;
    return this.physicalRow(logicalRow);
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

  /** Synchronous viewport read. It starts a bounded block load when needed. */
  peekCellValue(rowIndex: number, fieldRef: DataSourceFieldRef): DataSourceContentResult<TableScalar> {
    const field = this.resolveField(fieldRef);
    if (field === undefined) return this.errorResult(`Unknown data source field: ${String(fieldRef)}`);
    if (!isSafeRowIndex(rowIndex) || rowIndex >= this.source.rowCount) return this.errorResult(`Data source row is outside range: ${String(rowIndex)}`);
    const physicalRow = this.physicalRow(rowIndex);
    const ref = this.findBlock(physicalRow);
    if (ref === undefined) return this.missingResult(`No data block covers source row ${String(rowIndex)}`);
    const loaded = this.loadedBlocks.get(ref.id);
    if (loaded !== undefined) {
      const row = loaded.rows[physicalRow - ref.startRow];
      if (!row) return this.errorResult(`Data block ${ref.id} does not contain source row ${String(rowIndex)}`);
      return { state: state(this.source.id, ref.id, 'ready'), value: row[field.ordinal] ?? null };
    }
    const current = this.loadStates.get(ref.id);
    // Render reads expose a terminal failure until an explicit asynchronous
    // read retries it. A repaint must not restart a failed network request.
    if (current === undefined && !this.loadPromises.has(ref.id)) void this.loadBlock(ref).catch(() => undefined);
    return { state: current ?? state(this.source.id, ref.id, 'loading') };
  }

  prefetchRows(startRow: number, rowCount: number): void {
    const error = this.validateRange(startRow, rowCount);
    if (error || rowCount === 0) return;
    const refs: DataBlockRef[] = [];
    if (this.source.rowOrder === undefined) {
      const first = this.findBlock(startRow);
      const last = this.findBlock(startRow + rowCount - 1);
      if (!first || !last) return;
      const firstIndex = this.blockIndexesById.get(first.id);
      const lastIndex = this.blockIndexesById.get(last.id);
      if (firstIndex === undefined || lastIndex === undefined) return;
      for (let index = firstIndex; index <= lastIndex; index += 1) {
        const ref = this.source.blocks[index]!;
        if (!this.loadedBlocks.has(ref.id) && !this.loadPromises.has(ref.id)) refs.push(ref);
      }
    } else {
      const scheduled = new Set<string>();
      for (let row = startRow; row < startRow + rowCount; row += 1) {
        const ref = this.findBlock(this.physicalRow(row));
        if (ref && !scheduled.has(ref.id)) {
          scheduled.add(ref.id);
          if (!this.loadedBlocks.has(ref.id) && !this.loadPromises.has(ref.id)) refs.push(ref);
        }
      }
    }
    if (refs.length > 0) void this.loadBlockSet(refs).catch(() => undefined);
  }

  /**
   * Ensure every source block is available without constructing a copied
   * range result. Metadata-only commands such as filtering need the canonical
   * block projection to be readable, but do not need a second full row matrix.
   */
  async ensureAllBlocksLoaded(): Promise<DataSourceContentLoadState> {
    try {
      await this.loadBlockSet(this.source.blocks);
    } catch (error) {
      const failed = this.source.blocks.find((ref) => {
        const availability = this.loadStates.get(ref.id)?.availability;
        return availability === 'missing' || availability === 'error';
      });
      return failed === undefined
        ? state(this.source.id, null, 'error', errorMessage(error))
        : this.loadStates.get(failed.id)!;
    }
    const last = this.source.blocks[this.source.blocks.length - 1];
    return state(this.source.id, last?.id ?? null, 'ready');
  }

  /** Read a physical row after ensureAllBlocksLoaded without allocating a flat row matrix. */
  getLoadedPhysicalRow(physicalRow: number): readonly TableScalar[] | undefined {
    if (!isSafeRowIndex(physicalRow) || physicalRow >= this.source.rowCount) return undefined;
    const ref = this.findBlock(physicalRow);
    return ref === undefined ? undefined : this.loadedBlocks.get(ref.id)?.rows[physicalRow - ref.startRow];
  }

  /**
   * Return decoded block views without rebuilding a copied full-range matrix.
   * Mutation paths that only inspect or reorder rows can use this view while
   * the public getRows copy contract remains unchanged.
   */
  async getAllBlockRows(): Promise<DataSourceContentResult<readonly DataSourceLoadedBlockView[]>> {
    const loaded = await this.ensureAllBlocksLoaded();
    if (loaded.availability !== 'ready') return { state: loaded };
    const blocks: DataSourceLoadedBlockView[] = [];
    for (const ref of this.source.blocks) {
      const block = this.loadedBlocks.get(ref.id);
      if (block === undefined) {
        return this.errorResult(`Data block ${ref.id} was not available after loading`);
      }
      blocks.push(block);
    }
    return { state: loaded, value: blocks };
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

  /**
   * Visit logical rows without constructing a copied row matrix. The visitor
   * may return false to stop before loading later blocks; decoded block rows
   * remain read-only views owned by this query.
   */
  async scanRows(
    visitor: (row: readonly TableScalar[], logicalRow: number) => boolean | void,
    options: DataSourceScanOptions = {},
  ): Promise<DataSourceContentResult<boolean>> {
    const source = this.source;
    const { signal, prefetchAllBlocks = false } = options;
    assertNotAborted(signal);
    let lastState = state(source.id, null, 'ready');
    const visit = (row: readonly TableScalar[], logicalRow: number): DataSourceContentResult<boolean> | undefined => {
      try {
        assertNotAborted(signal);
        return visitor(row, logicalRow) === false ? { state: lastState, value: false } : undefined;
      } catch (error) {
        if (error instanceof Error && error.name === 'AbortError') throw error;
        return this.errorResult(`Data source row ${String(logicalRow)} scan failed: ${errorMessage(error)}`);
      }
    };
    let prefetched: Map<string, LoadedBlock> | undefined;
    if (prefetchAllBlocks) {
      try {
        prefetched = await this.loadBlockSet(source.blocks, signal);
      } catch (error) {
        if (error instanceof Error && error.name === 'AbortError') throw error;
        const failed = source.blocks.find((ref) => {
          const availability = this.loadStates.get(ref.id)?.availability;
          return availability === 'missing' || availability === 'error';
        });
        const current = failed === undefined
          ? state(source.id, null, 'error', errorMessage(error))
          : this.loadStates.get(failed.id)!;
        return { state: { ...current } };
      }
      if (this.source !== source) return this.errorResult('Data source changed while rows were scanning');
    }
    if (source.rowOrder === undefined) {
      let logicalRow = 0;
      for (const ref of source.blocks) {
        let block: LoadedBlock;
        try {
          block = prefetched?.get(ref.id) ?? await this.loadBlock(ref, signal);
        } catch (error) {
          if (error instanceof Error && error.name === 'AbortError') throw error;
          const current = this.loadStates.get(ref.id) ?? state(source.id, ref.id, 'error', errorMessage(error));
          return { state: { ...current } };
        }
        if (this.source !== source) return this.errorResult('Data source changed while rows were scanning');
        lastState = state(source.id, ref.id, 'ready');
        for (const row of block.rows) {
          const result = visit(row, logicalRow++);
          if (result !== undefined) return result;
        }
      }
      return { state: lastState, value: true };
    }
    for (let logicalRow = 0; logicalRow < source.rowCount; logicalRow += 1) {
      assertNotAborted(signal);
      const physicalRow = this.physicalRowFor(source, logicalRow);
      const ref = this.findBlockIn(source, physicalRow);
      if (!ref) return this.missingResult(`No data block covers source row ${String(logicalRow)}`);
      let block: LoadedBlock;
      try {
        block = prefetched?.get(ref.id) ?? await this.loadBlock(ref, signal);
      } catch (error) {
        if (error instanceof Error && error.name === 'AbortError') throw error;
        const current = this.loadStates.get(ref.id) ?? state(source.id, ref.id, 'error', errorMessage(error));
        return { state: { ...current } };
      }
      if (this.source !== source) return this.errorResult('Data source changed while rows were scanning');
      lastState = state(source.id, ref.id, 'ready');
      const row = block.rows[physicalRow - ref.startRow];
      if (!row) return this.errorResult(`Data block ${ref.id} does not contain source row ${String(logicalRow)}`);
      const result = visit(row, logicalRow);
      if (result !== undefined) return result;
    }
    return { state: lastState, value: true };
  }

  /**
   * Load one field's member domain on demand.  The manifest deliberately does
   * not materialize distinct values for large sources; callers should invoke
   * this only when a value picker is opened.
   */
  async getDistinctFieldValues(
    fieldRef: DataSourceFieldRef,
    maxValues = PIVOT_MEMBER_DISPLAY_LIMIT,
    resolveValue?: (logicalRow: number, baseValue: TableScalar) => PivotScalar,
  ): Promise<DataSourceContentResult<PivotScalar[]>> {
    const field = this.resolveField(fieldRef);
    if (field === undefined) return this.errorResult(`Unknown data source field: ${String(fieldRef)}`);
    if (!Number.isSafeInteger(maxValues) || maxValues <= 0) return this.errorResult('Data source distinct-value limit must be a positive safe integer');

    const values: PivotScalar[] = [];
    const seen = new Set<string>();
    let overflow = false;
    const scanned = await this.scanRows((row, logicalRow) => {
      try {
        const baseValue = row[field.ordinal] ?? null;
        const value = resolveValue === undefined ? baseValue : resolveValue(logicalRow, baseValue);
        const key = value === null ? 'null' : `${typeof value}:${JSON.stringify(value)}`;
        if (seen.has(key)) return;
        if (values.length >= maxValues) {
          overflow = true;
          return false;
        }
        seen.add(key);
        values.push(value);
      } catch (error) {
        throw new Error(`Data source field ${field.name} overlay failed: ${errorMessage(error)}`);
      }
    });
    if (scanned.value === undefined) return { state: scanned.state };
    if (overflow) return this.errorResult(`Data source field ${field.name} exceeds the ${String(maxValues)} distinct-value limit`);
    return { state: scanned.state, value: values };
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

    const source = this.source;
    const refs = this.blockRefsForLogicalRange(source, startRow, rowCount);
    if (refs === undefined) {
      return this.missingResult(`No data block covers source rows ${String(startRow)}-${String(startRow + rowCount - 1)}`);
    }

    let loaded: Map<string, LoadedBlock>;
    try {
      loaded = await this.loadBlockSet(refs);
    } catch (error) {
      const failed = refs.find((ref) => {
        const availability = this.loadStates.get(ref.id)?.availability;
        return availability === 'missing' || availability === 'error';
      });
      const current = failed === undefined
        ? state(this.source.id, null, 'error', errorMessage(error))
        : this.loadStates.get(failed.id)!;
      return { state: { ...current } };
    }
    if (this.source !== source) return this.errorResult('Data source changed while rows were loading');

    const rows: TableScalar[][] = [];
    for (let row = startRow; row < startRow + rowCount; row += 1) {
      const physicalRow = this.physicalRowFor(source, row);
      const ref = this.findBlockIn(source, physicalRow);
      if (ref === undefined) {
        return this.missingResult(`No data block covers source row ${String(row)}`);
      }
      const block = loaded.get(ref.id)!;
      const localRow = physicalRow - ref.startRow;
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

  /**
   * Resolve an explicit set of immutable physical rows without converting
   * them through the current virtual sort order. Pivot drill-down provenance
   * is physical, so this preserves the selected source records after sorting.
   */
  async getRowsByPhysicalRow(physicalRows: readonly number[]): Promise<DataSourceContentResult<Map<number, TableScalar[]>>> {
    const source = this.source;
    const selected = [...new Set(physicalRows)];
    for (const physicalRow of selected) {
      if (!isSafeRowIndex(physicalRow) || physicalRow >= source.rowCount) {
        return this.errorResult(`Data source physical row is outside range: ${String(physicalRow)}`);
      }
    }
    const refs: DataBlockRef[] = [];
    const seen = new Set<string>();
    for (const physicalRow of selected) {
      const ref = this.findBlockIn(source, physicalRow);
      if (!ref) return this.missingResult(`No data block covers source physical row ${String(physicalRow)}`);
      if (!seen.has(ref.id)) {
        seen.add(ref.id);
        refs.push(ref);
      }
    }
    let loaded: Map<string, LoadedBlock>;
    try {
      loaded = await this.loadBlockSet(refs);
    } catch (error) {
      const failed = refs.find((ref) => {
        const availability = this.loadStates.get(ref.id)?.availability;
        return availability === 'missing' || availability === 'error';
      });
      const current = failed === undefined
        ? state(source.id, null, 'error', errorMessage(error))
        : this.loadStates.get(failed.id)!;
      return { state: { ...current } };
    }
    if (this.source !== source) return this.errorResult('Data source changed while physical rows were loading');
    const rows = new Map<number, TableScalar[]>();
    for (const physicalRow of selected) {
      const ref = this.findBlockIn(source, physicalRow)!;
      const values = loaded.get(ref.id)?.rows[physicalRow - ref.startRow];
      if (!values) return this.errorResult(`Data block ${ref.id} does not contain source physical row ${String(physicalRow)}`);
      rows.set(physicalRow, [...values]);
    }
    return { state: state(source.id, refs.length === 1 ? refs[0]!.id : null, 'ready'), value: rows };
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
    if (startRow > this.source.rowCount || rowCount > this.source.rowCount - startRow) return 'Data source query range exceeds rowCount';
    return undefined;
  }

  private physicalRow(logicalRow: number): number {
    return this.physicalRowFor(this.source, logicalRow);
  }

  private physicalRowFor(source: DataSourceManifest, logicalRow: number): number {
    return source.rowOrder?.[logicalRow] ?? logicalRow;
  }

  private blockRefsForLogicalRange(source: DataSourceManifest, startRow: number, rowCount: number): DataBlockRef[] | undefined {
    if (source.rowOrder === undefined) {
      const first = this.findBlockIn(source, startRow);
      const last = this.findBlockIn(source, startRow + rowCount - 1);
      if (!first || !last) return undefined;
      const firstIndex = this.blockIndexesById.get(first.id);
      const lastIndex = this.blockIndexesById.get(last.id);
      if (firstIndex === undefined || lastIndex === undefined) return undefined;
      return source.blocks.slice(firstIndex, lastIndex + 1);
    }
    const refs: DataBlockRef[] = [];
    const seen = new Set<string>();
    for (let row = startRow; row < startRow + rowCount; row += 1) {
      const ref = this.findBlockIn(source, this.physicalRowFor(source, row));
      if (!ref) return undefined;
      if (!seen.has(ref.id)) {
        seen.add(ref.id);
        refs.push(ref);
      }
    }
    return refs;
  }

  private findBlock(rowIndex: number): DataBlockRef | undefined {
    return this.findBlockIn(this.source, rowIndex);
  }

  private findBlockIn(source: DataSourceManifest, rowIndex: number): DataBlockRef | undefined {
    let low = 0;
    let high = source.blocks.length - 1;
    while (low <= high) {
      const middle = Math.floor((low + high) / 2);
      const block = source.blocks[middle]!;
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

  private async loadBlockSet(refs: readonly DataBlockRef[], signal?: AbortSignal): Promise<Map<string, LoadedBlock>> {
    assertNotAborted(signal);
    const uniqueRefs: DataBlockRef[] = [];
    const seen = new Set<string>();
    for (const ref of refs) {
      if (seen.has(ref.id)) continue;
      seen.add(ref.id);
      uniqueRefs.push(ref);
    }
    const loaded = new Map<string, LoadedBlock>();
    let nextIndex = 0;
    let stopScheduling = false;
    const failures = new Map<number, unknown>();
    const worker = async (): Promise<void> => {
      while (!stopScheduling) {
        const index = nextIndex;
        nextIndex += 1;
        if (index >= uniqueRefs.length) return;
        const ref = uniqueRefs[index]!;
        try {
          loaded.set(ref.id, await this.loadBlock(ref, signal));
        } catch (error) {
          failures.set(index, error);
          stopScheduling = true;
        }
      }
    };
    await Promise.all(Array.from({ length: Math.min(MAX_CONCURRENT_BLOCK_READS, uniqueRefs.length) }, () => worker()));
    assertNotAborted(signal);
    const firstFailure = [...failures.entries()].sort(([left], [right]) => left - right)[0];
    if (firstFailure !== undefined) throw firstFailure[1];
    return loaded;
  }

  private async loadBlock(ref: DataBlockRef, signal?: AbortSignal): Promise<LoadedBlock> {
    const cached = this.loadedBlocks.get(ref.id);
    if (cached !== undefined) return awaitWithoutCancelling(Promise.resolve(cached), signal);
    const existing = this.loadPromises.get(ref.id);
    if (existing !== undefined) return awaitWithoutCancelling(existing, signal);

    const promise = Promise.resolve().then(() => this.readBlock(ref)).then((block) => {
      const currentRef = this.blockRefsById.get(ref.id);
      if (currentRef === undefined) throw new ContentQueryFailure('error', `Data block ${ref.id} was removed while loading`);
      const currentBlock = block.ref === currentRef ? block : { ...block, ref: currentRef };
      this.loadedBlocks.set(ref.id, currentBlock);
      this.publishState(state(this.source.id, ref.id, 'ready'));
      return currentBlock;
    }).catch((error: unknown) => {
      const failure = error instanceof ContentQueryFailure
        ? error
        : new ContentQueryFailure('error', errorMessage(error));
      this.publishState(state(this.source.id, ref.id, failure.availability, failure.message));
      throw failure;
    }).finally(() => {
      this.loadPromises.delete(ref.id);
    });
    this.loadPromises.set(ref.id, promise);
    // Subscribers may synchronously read again. Register the flight before
    // notifying them so every reader shares this request, including retries.
    this.publishState(state(this.source.id, ref.id, 'loading'));
    return awaitWithoutCancelling(promise, signal);
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
      const decoded = await decodeOwnedColumnarBlock(record.bytes, {
        expectedRowCount: ref.rowCount,
        expectedFields: this.source.fields,
        expectedChecksum: ref.checksum,
      });
      const rows = decoded.rows;
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
