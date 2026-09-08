import { kernelInvoke, KernelInvocationError } from '@react-sheets/kernel-client';
import type { CellData, RangeRef } from './index';

export const KERNEL_PAGE_ROWS = 1024 as const;
export const KERNEL_PAGE_COLUMNS = 32 as const;
export interface KernelReplicaPageDescriptor { readonly sheetId: string; readonly pageRow: number; readonly pageColumn: number; readonly revision: number; readonly checksum: string; readonly byteLength: number; readonly cellCount: number; readonly occupiedRange: RangeRef | null; }
export interface KernelReplicaSheetManifest { readonly sheetId: string; readonly name: string; readonly rowCount: number; readonly columnCount: number; readonly metadata: Record<string, unknown>; }
export interface KernelReplicaManifest { readonly schema: 'WorkbookManifest'; readonly version: 11; readonly unitId: string; readonly name: string; readonly revision: number; readonly sheets: readonly KernelReplicaSheetManifest[]; readonly pages: readonly KernelReplicaPageDescriptor[]; readonly metadata: Record<string, unknown>; }
export interface KernelReplicaPagePayload extends KernelReplicaPageDescriptor { readonly payloadBase64: string; }
export interface KernelReplicaCellAddress { readonly sheetId: string; readonly row: number; readonly column: number; }
export interface KernelPageTransport {
  getPage(
    params: { unitId: string; revision: number; sheetId: string; pageRow: number; pageColumn: number },
    options?: { signal?: AbortSignal },
  ): Promise<KernelReplicaPagePayload>;
}

/** Owns transport coordination only. Page bytes and cell values belong to Rust. */
export class KernelPageReplica {
  private static readonly MAX_PAGE_REQUESTS = 4;
  private manifestValue: KernelReplicaManifest | null = null;
  private readonly requests = new Map<string, Promise<void>>();
  private readonly residentPages = new Set<string>();
  private readonly pageDirectory = new Map<string, KernelReplicaPageDescriptor>();
  private readonly queuedLoads: Array<{
    key: string;
    manifest: KernelReplicaManifest;
    page: KernelReplicaPageDescriptor;
    transport: KernelPageTransport;
    controller: AbortController;
    resolve: () => void;
    reject: (error: unknown) => void;
  }> = [];
  private activeLoads = 0;
  private revisionController: AbortController | null = null;
  constructor(readonly unitId: string) {}
  get manifest(): KernelReplicaManifest {
    if (!this.manifestValue) throw new KernelInvocationError({ code: 'WORKBOOK_NOT_OPEN', message: 'The committed workbook manifest has not been loaded.', object: this.unitId, recovery: 'open-cloud-workbook' });
    return this.manifestValue;
  }
  get revision(): number { return this.manifest.revision; }
  open(manifest: KernelReplicaManifest, pages: readonly KernelReplicaPagePayload[] = []): void {
    if (manifest.unitId !== this.unitId) throw new KernelInvocationError({ code: 'WORKBOOK_ID_MISMATCH', message: 'Manifest belongs to another workbook.', object: manifest.unitId, recovery: 'reload-cloud-workbook' });
    const previousDirectory = this.pageDirectory;
    const previousResidentPages = this.residentPages;
    const nextDirectory = new Map<string, KernelReplicaPageDescriptor>();
    const nextResidentPages = new Set<string>();
    for (const page of manifest.pages) {
      const key = this.pageKey(page.sheetId, page.pageRow, page.pageColumn);
      nextDirectory.set(key, page);
      const previous = previousDirectory.get(key);
      if (previousResidentPages.has(key) && previous && this.samePage(previous, page)) nextResidentPages.add(key);
    }

    // The manifest is the directory boundary. Page bytes cross the kernel
    // boundary independently so a large revision can never form one control
    // frame. Keep the JS projection unchanged until every supplied page has
    // been accepted by the new kernel revision.
    kernelInvoke('open', { manifest });
    const previousController = this.revisionController;
    if (previousController) {
      previousController.abort();
      this.rejectQueuedLoads(previousController);
    }
    const revisionController = new AbortController();
    this.revisionController = revisionController;
    try {
      for (const page of pages) {
        const key = this.pageKey(page.sheetId, page.pageRow, page.pageColumn);
        if (!nextDirectory.has(key)) {
          throw new KernelInvocationError({
            code: 'PAGE_UNKNOWN',
            message: 'Page payload is not present in the committed manifest.',
            object: key,
            recovery: 'reload-cloud-workbook',
          });
        }
        kernelInvoke('page.load', { unitId: this.unitId, revision: manifest.revision, page });
        nextResidentPages.add(key);
      }
    } catch (error) {
      // A same-revision reload reused the prior kernel state, so preserve the
      // old JS projection for callers that reject a corrupt replacement. A
      // failed revision transition is fail-closed; the caller must reload the
      // authoritative workbook before attempting another read.
      if (!this.manifestValue || this.manifestValue.revision !== manifest.revision) {
        revisionController.abort();
        try { kernelInvoke('close', { unitId: this.unitId }); } catch { /* preserve the original typed failure */ }
      }
      throw error;
    }

    this.manifestValue = structuredClone(manifest);
    this.requests.clear();
    this.pageDirectory.clear();
    for (const [key, page] of nextDirectory) this.pageDirectory.set(key, page);
    this.residentPages.clear();
    for (const key of nextResidentPages) this.residentPages.add(key);
  }
  readCell(address: KernelReplicaCellAddress): CellData | undefined {
    return kernelInvoke<{ revision: number; cell: CellData | null }>('cell.get', { unitId: this.unitId, revision: this.revision, address }).cell ?? undefined;
  }
  readRange(range: RangeRef): Array<{ address: KernelReplicaCellAddress; cell: CellData }> {
    return kernelInvoke<{ revision: number; cells: Array<{ address: KernelReplicaCellAddress; cell: CellData }> }>('range.get', { unitId: this.unitId, revision: this.revision, range }).cells;
  }
  readSheetStats(sheetId: string): { sheetId: string; cellCount: number; occupiedRange: RangeRef | null } {
    return kernelInvoke('sheet.stats', { unitId: this.unitId, revision: this.revision, sheetId });
  }
  resolveCurrentRegion(sheetId: string, activeRow: number, activeColumn: number): RangeRef {
    return kernelInvoke<{ revision: number; range: RangeRef }>('dataRegion.resolve', {
      unitId: this.unitId, revision: this.revision, sheetId, activeRow, activeColumn,
    }).range;
  }
  async loadRange(range: RangeRef, transport: KernelPageTransport): Promise<void> {
    const manifest = this.manifest;
    const controller = this.revisionController;
    if (!controller) throw this.staleRevisionError(`${manifest.revision}:${this.pageKey(range.sheetId, 0, 0)}`);
    const pages: KernelReplicaPageDescriptor[] = [];
    const firstPageRow = Math.floor(Math.max(0, range.startRow) / KERNEL_PAGE_ROWS);
    const lastPageRow = Math.floor(Math.max(0, range.endRow) / KERNEL_PAGE_ROWS);
    const firstPageColumn = Math.floor(Math.max(0, range.startColumn) / KERNEL_PAGE_COLUMNS);
    const lastPageColumn = Math.floor(Math.max(0, range.endColumn) / KERNEL_PAGE_COLUMNS);
    for (let pageRow = firstPageRow; pageRow <= lastPageRow; pageRow += 1) {
      for (let pageColumn = firstPageColumn; pageColumn <= lastPageColumn; pageColumn += 1) {
        const key = this.pageKey(range.sheetId, pageRow, pageColumn);
        const page = this.pageDirectory.get(key);
        if (page && !this.residentPages.has(key)) pages.push(page);
      }
    }
    await Promise.all(pages.map(page => this.loadPage(page, manifest, controller, transport)));
  }
  isRangeResident(range: RangeRef): boolean {
    const manifest = this.manifest;
    const firstPageRow = Math.floor(Math.max(0, range.startRow) / KERNEL_PAGE_ROWS);
    const lastPageRow = Math.floor(Math.max(0, range.endRow) / KERNEL_PAGE_ROWS);
    const firstPageColumn = Math.floor(Math.max(0, range.startColumn) / KERNEL_PAGE_COLUMNS);
    const lastPageColumn = Math.floor(Math.max(0, range.endColumn) / KERNEL_PAGE_COLUMNS);
    for (let pageRow = firstPageRow; pageRow <= lastPageRow; pageRow += 1) {
      for (let pageColumn = firstPageColumn; pageColumn <= lastPageColumn; pageColumn += 1) {
        const key = this.pageKey(range.sheetId, pageRow, pageColumn);
        if (this.pageDirectory.has(key) && !this.residentPages.has(key)) return false;
      }
    }
    return true;
  }
  isPageResident(sheetId: string, pageRow: number, pageColumn: number): boolean {
    return this.residentPages.has(this.pageKey(sheetId, pageRow, pageColumn));
  }
  private loadPage(page: KernelReplicaPageDescriptor, manifest: KernelReplicaManifest, controller: AbortController, transport: KernelPageTransport): Promise<void> {
    const pageKey = this.pageKey(page.sheetId, page.pageRow, page.pageColumn);
    if (this.residentPages.has(pageKey)) return Promise.resolve();
    const requestKey = `${manifest.revision}:${pageKey}`;
    const existing = this.requests.get(requestKey);
    if (existing) return existing;
    const request = new Promise<void>((resolve, reject) => {
      this.queuedLoads.push({ key: requestKey, manifest, page, transport, controller, resolve, reject });
      this.drainLoadQueue();
    });
    this.requests.set(requestKey, request);
    void request.finally(() => { if (this.requests.get(requestKey) === request) this.requests.delete(requestKey); }).catch(() => undefined);
    return request;
  }
  private drainLoadQueue(): void {
    while (this.activeLoads < KernelPageReplica.MAX_PAGE_REQUESTS && this.queuedLoads.length > 0) {
      const task = this.queuedLoads.shift()!;
      if (this.isStale(task.controller, task.manifest.revision)) {
        task.reject(this.staleRevisionError(task.key));
        continue;
      }
      this.activeLoads += 1;
      void this.executeLoad(task).then(task.resolve, task.reject).finally(() => {
        this.activeLoads -= 1;
        this.drainLoadQueue();
      });
    }
  }
  private async executeLoad(task: {
    key: string;
    manifest: KernelReplicaManifest;
    page: KernelReplicaPageDescriptor;
    transport: KernelPageTransport;
    controller: AbortController;
  }): Promise<void> {
    try {
      const { manifest, page, transport, controller } = task;
      const payload = await transport.getPage({ unitId: this.unitId, revision: manifest.revision, sheetId: page.sheetId, pageRow: page.pageRow, pageColumn: page.pageColumn }, { signal: controller.signal });
      if (this.isStale(controller, manifest.revision)) throw this.staleRevisionError(task.key);
      kernelInvoke('page.load', { unitId: this.unitId, revision: manifest.revision, page: payload });
      this.residentPages.add(task.key.slice(task.key.indexOf(':') + 1));
    } catch (error) {
      if (this.isStale(task.controller, task.manifest.revision)) throw this.staleRevisionError(task.key);
      throw error;
    }
  }
  private isStale(controller: AbortController, revision: number): boolean {
    return controller.signal.aborted || this.revisionController !== controller || this.manifestValue?.revision !== revision;
  }
  private staleRevisionError(requestKey: string): KernelInvocationError {
    return new KernelInvocationError({ code: 'STALE_REVISION', message: 'Workbook changed while a page was loading.', object: requestKey, recovery: 'reload-visible-pages' });
  }
  private rejectQueuedLoads(controller: AbortController): void {
    for (let index = this.queuedLoads.length - 1; index >= 0; index -= 1) {
      const task = this.queuedLoads[index]!;
      if (task.controller !== controller) continue;
      this.queuedLoads.splice(index, 1);
      task.reject(this.staleRevisionError(task.key));
    }
  }
  private pageKey(sheetId: string, pageRow: number, pageColumn: number): string {
    return `${sheetId}:${pageRow}:${pageColumn}`;
  }
  private samePage(left: KernelReplicaPageDescriptor, right: KernelReplicaPageDescriptor): boolean {
    return left.sheetId === right.sheetId
      && left.pageRow === right.pageRow
      && left.pageColumn === right.pageColumn
      && left.revision === right.revision
      && left.checksum === right.checksum
      && left.byteLength === right.byteLength
      && left.cellCount === right.cellCount
      && JSON.stringify(left.occupiedRange) === JSON.stringify(right.occupiedRange);
  }
}

/** Read-only revision-pinned page projection. Hidden cells retain their canonical values. */
export class WorksheetCells {
  constructor(private readonly replica: KernelPageReplica, readonly sheetId: string) {}
  get revision(): number { return this.replica.revision; }
  get(row: number, column: number): CellData | undefined { return this.replica.readCell({ sheetId: this.sheetId, row, column }); }
  has(row: number, column: number): boolean { return this.get(row, column) !== undefined; }
  *entries(): IterableIterator<{ cell: CellData; row: number; column: number }> {
    const sheet = this.replica.manifest.sheets.find(entry => entry.sheetId === this.sheetId)!;
    yield* this.entriesInRange(0, sheet.rowCount - 1, 0, sheet.columnCount - 1);
  }
  private *entriesInRange(startRow: number, endRow: number, startColumn: number, endColumn: number): IterableIterator<{ cell: CellData; row: number; column: number }> {
    for (const page of this.replica.manifest.pages) {
      if (page.sheetId !== this.sheetId) continue;
      const firstRow = Math.max(startRow, page.pageRow * KERNEL_PAGE_ROWS), lastRow = Math.min(endRow, (page.pageRow + 1) * KERNEL_PAGE_ROWS - 1);
      const firstColumn = Math.max(startColumn, page.pageColumn * KERNEL_PAGE_COLUMNS), lastColumn = Math.min(endColumn, (page.pageColumn + 1) * KERNEL_PAGE_COLUMNS - 1);
      if (firstRow > lastRow || firstColumn > lastColumn) continue;
      for (const entry of this.replica.readRange({ sheetId: this.sheetId, startRow: firstRow, endRow: lastRow, startColumn: firstColumn, endColumn: lastColumn })) yield { cell: entry.cell, row: entry.address.row, column: entry.address.column };
    }
  }
  forEach(callback: (cell: CellData, row: number, column: number) => void): void { for (const entry of this.entries()) callback(entry.cell, entry.row, entry.column); }
  forEachInRange(startRow: number, endRow: number, startColumn: number, endColumn: number, callback: (cell: CellData, row: number, column: number) => void): void { for (const entry of this.entriesInRange(startRow, endRow, startColumn, endColumn)) callback(entry.cell, entry.row, entry.column); }
  forEachInRows(rows: ReadonlySet<number>, callback: (cell: CellData, row: number, column: number) => void): void { const sheet = this.replica.manifest.sheets.find(entry => entry.sheetId === this.sheetId)!; for (const row of rows) this.forEachInRange(row, row, 0, sheet.columnCount - 1, callback); }
  forEachInColumns(columns: ReadonlySet<number>, callback: (cell: CellData, row: number, column: number) => void): void { for (const column of columns) for (const entry of this.entriesInColumn(column)) callback(entry.cell, entry.row, column); }
  *entriesInColumn(column: number): IterableIterator<{ row: number; cell: CellData }> { const sheet = this.replica.manifest.sheets.find(entry => entry.sheetId === this.sheetId)!; for (const entry of this.entriesInRange(0, sheet.rowCount - 1, column, column)) yield { row: entry.row, cell: entry.cell }; }
  count(): number { return this.replica.readSheetStats(this.sheetId).cellCount; }
  occupiedRange(sheetId: string): RangeRef {
    return this.replica.readSheetStats(this.sheetId).occupiedRange ?? { sheetId, startRow: 0, endRow: 0, startColumn: 0, endColumn: 0 };
  }
  currentRegion(activeRow: number, activeColumn: number): RangeRef { return this.replica.resolveCurrentRegion(this.sheetId, activeRow, activeColumn); }
}
