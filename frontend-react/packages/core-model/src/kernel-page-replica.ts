import { kernelInvoke, KernelInvocationError } from '@react-sheets/kernel-client';
import type { CellData, RangeRef } from './index';

export const KERNEL_PAGE_ROWS = 1024 as const;
export const KERNEL_PAGE_COLUMNS = 32 as const;
export interface KernelReplicaPageDescriptor { readonly sheetId: string; readonly pageRow: number; readonly pageColumn: number; readonly revision: number; readonly checksum: string; readonly byteLength: number; readonly cellCount: number; readonly occupiedRange: RangeRef | null; }
export interface KernelReplicaSheetManifest { readonly sheetId: string; readonly name: string; readonly rowCount: number; readonly columnCount: number; readonly metadata: Record<string, unknown>; }
export interface KernelReplicaManifest { readonly schema: 'WorkbookManifest'; readonly version: 11; readonly unitId: string; readonly name: string; readonly revision: number; readonly sheets: readonly KernelReplicaSheetManifest[]; readonly pages: readonly KernelReplicaPageDescriptor[]; readonly metadata: Record<string, unknown>; }
export interface KernelReplicaPagePayload extends KernelReplicaPageDescriptor { readonly payloadBase64: string; }
export interface KernelReplicaCellAddress { readonly sheetId: string; readonly row: number; readonly column: number; }
export interface KernelPageTransport { getPage(params: { unitId: string; revision: number; sheetId: string; pageRow: number; pageColumn: number }): Promise<KernelReplicaPagePayload>; }

/** Owns transport coordination only. Page bytes and cell values belong to Rust. */
export class KernelPageReplica {
  private manifestValue: KernelReplicaManifest | null = null;
  private readonly requests = new Map<string, Promise<void>>();
  constructor(readonly unitId: string) {}
  get manifest(): KernelReplicaManifest {
    if (!this.manifestValue) throw new KernelInvocationError({ code: 'WORKBOOK_NOT_OPEN', message: 'The committed workbook manifest has not been loaded.', object: this.unitId, recovery: 'open-cloud-workbook' });
    return this.manifestValue;
  }
  get revision(): number { return this.manifest.revision; }
  open(manifest: KernelReplicaManifest, pages: readonly KernelReplicaPagePayload[] = []): void {
    if (manifest.unitId !== this.unitId) throw new KernelInvocationError({ code: 'WORKBOOK_ID_MISMATCH', message: 'Manifest belongs to another workbook.', object: manifest.unitId, recovery: 'reload-cloud-workbook' });
    kernelInvoke('open', { manifest, pages });
    this.manifestValue = structuredClone(manifest);
    this.requests.clear();
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
  async loadRange(range: RangeRef, transport: KernelPageTransport): Promise<void> {
    const manifest = this.manifest;
    const pages = manifest.pages.filter(page => page.sheetId === range.sheetId && page.pageRow >= Math.floor(range.startRow / KERNEL_PAGE_ROWS) && page.pageRow <= Math.floor(range.endRow / KERNEL_PAGE_ROWS) && page.pageColumn >= Math.floor(range.startColumn / KERNEL_PAGE_COLUMNS) && page.pageColumn <= Math.floor(range.endColumn / KERNEL_PAGE_COLUMNS));
    for (let offset = 0; offset < pages.length; offset += 4) await Promise.all(pages.slice(offset, offset + 4).map(page => this.loadPage(page, manifest, transport)));
  }
  private loadPage(page: KernelReplicaPageDescriptor, manifest: KernelReplicaManifest, transport: KernelPageTransport): Promise<void> {
    const key = `${manifest.revision}:${page.sheetId}:${page.pageRow}:${page.pageColumn}`;
    const existing = this.requests.get(key);
    if (existing) return existing;
    const request = (async () => {
      const payload = await transport.getPage({ unitId: this.unitId, revision: manifest.revision, sheetId: page.sheetId, pageRow: page.pageRow, pageColumn: page.pageColumn });
      if (this.manifest.revision !== manifest.revision) throw new KernelInvocationError({ code: 'STALE_REVISION', message: 'Workbook changed while a page was loading.', object: key, recovery: 'reload-visible-pages' });
      kernelInvoke('page.load', { unitId: this.unitId, revision: manifest.revision, page: payload });
    })();
    this.requests.set(key, request);
    void request.finally(() => { if (this.requests.get(key) === request) this.requests.delete(key); }).catch(() => undefined);
    return request;
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
}
