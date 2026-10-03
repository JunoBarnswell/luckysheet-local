import { cellAddress } from '@react-sheets/core-model';
import type { CellSnapshot } from './contract';
import { immutableSnapshot } from './value';
import { CommandDispatchError } from '../../../spreadsheet-app/src/workbook-session';
import type { WorkbookObjectPort } from '../../../spreadsheet-app/src/workbook-object-port';
import { SdkError } from '../error';
import { domainFor, registerWorkbookDomain } from './object-domain';
import { WorksheetCollection } from './worksheet-collection';
import { DefinedNameCollection } from './defined-name';
import { WorkbookExternalLinks } from './external-links';
import type { DataActions } from '../data/contract';

export class Workbook {
  readonly id: string;
  readonly worksheets: WorksheetCollection;
  readonly names: DefinedNameCollection;
  readonly externalLinks: WorkbookExternalLinks;
  readonly data: DataActions;
  #closed = false;
  #closeReady: (() => void) | null = null;
  #opening: Promise<Workbook> | null = null;
  #unsubscribeLifetime: () => void = () => {};
  constructor(readonlyPort: WorkbookObjectPort, scope: object, close: () => void, dataFor: (assertCurrent: () => void) => DataActions) {
    this.#port = readonlyPort;
    this.#release = close;
    this.id = readonlyPort.unitId;
    this.worksheets = new WorksheetCollection(this);
    this.externalLinks = new WorkbookExternalLinks(this);
    this.names = new DefinedNameCollection(this);
    this.data = dataFor(() => this.#assertAlive('data'));
    this.#initializeDomain(scope);
    this.#unsubscribeLifetime = readonlyPort.subscribeDisposed(() => this.close());
    Object.freeze(this);
  }
  readonly #port: WorkbookObjectPort;
  readonly #release: () => void;
  get name(): string { this.#assertAlive('workbook.name'); return this.#port.state().name; }
  #assertAlive(operation: string): void {
    if (this.#closed || this.#port.state().disposed) throw this.#error('RUNTIME_DISPOSED', operation, '工作簿对象已关闭。', undefined);
  }
  #error(code: SdkError['code'], operation: string, message: string, cause: unknown, object?: { sheetId?: string; address?: string }): SdkError {
    return new SdkError(code, operation, `${this.id}: ${message}`, '请检查工作簿权限和状态，重新打开后重试。', { cause, object: { workbookId: this.id, ...object } });
  }
  ready(): Promise<Workbook> {
    this.#assertAlive('workbook.open');
    if (this.#opening) return this.#opening;
    this.#opening = new Promise((resolve, reject) => {
      let unsubscribe = () => {};
      const finish = (error?: SdkError) => { clearTimeout(timer); unsubscribe(); this.#closeReady = null; error ? reject(error) : resolve(this); };
      const check = () => {
        if (this.#closed) return finish(this.#error('RUNTIME_DISPOSED', 'workbook.open', '打开过程中对象已关闭。', undefined));
        const state = this.#port.state();
        if (state.phase === 'ready') finish();
        else if (state.phase === 'error') finish(this.#error('REQUEST_REJECTED', 'workbook.open', state.notice, state));
      };
      const timer = setTimeout(() => finish(this.#error('SERVICE_UNAVAILABLE', 'workbook.open', '等待 canonical ready 超时。', this.#port.state())), 30_000);
      this.#closeReady = check;
      unsubscribe = this.#port.subscribe(check);
      check();
    });
    return this.#opening;
  }
  close(): void {
    if (this.#closed) return;
    this.#closed = true;
    this.#unsubscribeLifetime();
    this.#closeReady?.();
    this.#release();
  }
  async undo(): Promise<boolean> { return domainFor(this).history('undo'); }
  async redo(): Promise<boolean> { return domainFor(this).history('redo'); }
  async save(): Promise<void> { await this.#perform('workbook.save', () => this.#port.save()); }
  async flush(): Promise<void> { await this.#perform('workbook.flush', () => this.#port.flush()); }
  async #perform<T>(operation: string, action: () => Promise<T>, object?: { sheetId?: string; address?: string }): Promise<T> {
    try { this.#assertAlive(operation); const result = await action(); this.#assertAlive(operation); return result; }
    catch (cause) {
      if (cause instanceof SdkError) throw cause;
      const code = cause instanceof Error && 'code' in cause && cause.code === 'CIRCULAR_DEPENDENCY' ? 'CIRCULAR_DEPENDENCY' : cause instanceof CommandDispatchError && cause.code === 'PERMISSION_DENIED' ? 'FORBIDDEN' : cause instanceof Error && 'code' in cause && cause.code === 'STALE_OPERATION' ? 'STALE_OPERATION' : 'REQUEST_REJECTED';
      throw this.#error(code, operation, cause instanceof Error ? cause.message : '操作失败。', cause, object);
    }
  }
  #cellSnapshot(resolved: Awaited<ReturnType<WorkbookObjectPort['readCell']>>): CellSnapshot {
    return immutableSnapshot({ value: resolved.cell?.value ?? null,
      ...(resolved.cell?.formula === undefined ? {} : { formula: resolved.cell.formula }),
      calculatedValue: resolved.calculatedValue, formulaHidden: resolved.formulaHidden,
      ...(resolved.cell?.style === undefined ? {} : { style: resolved.cell.style }),
      ...(resolved.cell?.numberFormat === undefined ? {} : { numberFormat: resolved.cell.numberFormat }),
      ...(resolved.cell?.richText === undefined ? {} : { richText: resolved.cell.richText }),
    });
  }
  #initializeDomain(scope: object): void {
    registerWorkbookDomain(this, {
      scope,
      sheet: (sheetId) => {
        this.#assertAlive('worksheet.read');
        try { return this.#port.readWorksheet(sheetId); }
        catch (cause) { throw this.#error('REQUEST_REJECTED', 'worksheet.read', cause instanceof Error ? cause.message : 'Worksheet metadata unavailable.', cause, { sheetId }); }
      },
      names: () => {
        this.#assertAlive('definedNames.read');
        try { return immutableSnapshot(this.#port.readDefinedNames()); }
        catch (cause) { throw this.#error('REQUEST_REJECTED', 'definedNames.read', cause instanceof Error ? cause.message : 'Defined names are unavailable.', cause); }
      },
      sheets: () => { this.#assertAlive('worksheets.read'); return this.#port.sheets(); },
      invalid: (operation, cause, object) => { this.#assertAlive(operation); throw this.#error('INVALID_ARGUMENT', operation, cause instanceof Error ? cause.message : String(cause), cause, object); },
      read: (sheetId, row, column) => this.#perform('cell.read', async () => this.#cellSnapshot(await this.#port.readCell(sheetId, row, column)), { sheetId, address: cellAddress(row, column) }),
      readRange: (range) => this.#perform('range.read', async () => {
        const cells = await this.#port.readCells(range), rows: CellSnapshot[][] = [];
        const width = range.endColumn - range.startColumn + 1;
        for (let index = 0; index < cells.length; index += width) rows.push(cells.slice(index, index + width).map(cell => this.#cellSnapshot(cell)));
        return immutableSnapshot(rows);
      }, { sheetId: range.sheetId }),
      writeRange: (range, inputs) => this.#perform('range.write', async () => {
        const plan = structuredClone(inputs);
        const sheet = this.#port.sheets().find(sheet => sheet.id === range.sheetId);
        if (!sheet || sheet.kind !== 'worksheet' || range.endRow >= sheet.rowCount || range.endColumn >= sheet.columnCount) throw this.#error('INVALID_ARGUMENT', 'range.write', 'Matrix must fit the canonical worksheet extent; use Worksheet.growExtent first.', undefined, { sheetId: range.sheetId });
        const cells = await this.#port.readCells(range);
        if (cells.some(cell => !cell.writable || cell.recordField || cell.sheetId !== range.sheetId)) throw this.#error('UNSUPPORTED_FEATURE', 'range.write', 'Matrix targets must be writable worksheet cells owned by this sheet.', undefined, { sheetId: range.sheetId });
        const width = range.endColumn - range.startColumn + 1;
        const entries = cells.map((cell, index) => ({ row: cell.row, column: cell.column, input: plan[Math.floor(index / width)]![index % width]!, inputContext: cell.inputContext }));
        this.#assertAlive('range.write');
        const result = await this.#port.dispatch({ commandId: 'sheet.cells.commitMatrix', params: { sheetId: range.sheetId, range, entries } });
        if (result.status === 'rejected') throw result.error;
      }, { sheetId: range.sheetId }),
      writeRichText: (range, text, runs) => this.#perform('range.setRichText', async () => {
        const intent = structuredClone({ text, runs });
        const sheet = this.#port.sheets().find(sheet => sheet.id === range.sheetId);
        if (!sheet || sheet.kind !== 'worksheet' || range.endRow >= sheet.rowCount || range.endColumn >= sheet.columnCount) throw this.#error('INVALID_ARGUMENT', 'range.setRichText', 'Rich text must fit the canonical worksheet extent.', undefined, { sheetId: range.sheetId });
        const cells = await this.#port.readCells(range);
        if (cells.some(cell => !cell.writable || cell.recordField || cell.sheetId !== range.sheetId)) throw this.#error('UNSUPPORTED_FEATURE', 'range.setRichText', 'Rich text targets must be writable cells owned by this worksheet.', undefined, { sheetId: range.sheetId });
        this.#assertAlive('range.setRichText');
        const outcome = await this.#port.dispatch({ commandId: 'sheet.cells.commitRichText', params: { ...intent, targets: cells.map(cell => ({ sheetId: cell.sheetId, row: cell.row, column: cell.column })) } });
        if (outcome.status === 'rejected') throw outcome.error;
      }, { sheetId: range.sheetId }),
      command: (operation, descriptor) => this.#perform(operation, async () => {
        const outcome = await this.#port.dispatch(structuredClone(descriptor));
        if (outcome.status === 'rejected') throw outcome.error;
      }),
      history: (direction) => this.#perform(`workbook.${direction}`, async () => this.#port.history(direction)),
      write: (sheetId, row, column, input) => this.#perform(input.kind === 'value' ? 'cell.setValue' : 'cell.setFormula', async () => {
        const resolved = await this.#port.readCell(sheetId, row, column);
        if (!resolved.writable) throw this.#error('UNSUPPORTED_FEATURE', 'cell.write', '此单元格由只读派生 owner 管理。', undefined, { sheetId });
        const canonical = { sheetId: resolved.sheetId, row: resolved.row, column: resolved.column, inputContext: resolved.inputContext };
        if (resolved.recordField && input.kind === 'formula') throw this.#error('UNSUPPORTED_FEATURE', 'cell.setFormula', '记录字段公式必须由 record calculation 定义管理。', undefined, { sheetId });
        const descriptor = resolved.recordField && input.kind === 'value'
          ? { commandId: 'record.set', params: { ...resolved.recordField, value: input.value } }
          : input.kind === 'value'
            ? { commandId: 'sheet.cell.commitTypedValue', params: { ...canonical, value: input.value } }
            : { commandId: 'sheet.cell.commitText', params: { ...canonical, text: input.formula } };
        this.#assertAlive('cell.write');
        const outcome = await this.#port.dispatch(structuredClone(descriptor));
        if (outcome.status === 'rejected') throw outcome.error;
      }, { sheetId, address: cellAddress(row, column) }),
      bind: (source, token) => this.#perform('externalLinks.bind', async () => {
        if (domainFor(source).scope !== scope || source.id === this.id || typeof token !== 'string' || !/^[A-Za-z0-9._:-]{1,200}$/.test(token)) {
          throw this.#error('INVALID_ARGUMENT', 'externalLinks.bind', '来源必须是同一 SDK 中另一工作簿，别名须符合 canonical token 约束。', undefined);
        }
        const sheets = source.worksheets.list().map(sheet => ({ token: sheet.name, sheetId: sheet.id }));
        await source.flush();
        this.#assertAlive('externalLinks.bind');
        await this.#port.bindExternalLink({ id: token, token, sourceUnitId: source.id, sheets });
      }),
      refresh: () => this.#perform('externalLinks.refresh', async () => Object.freeze((await this.#port.refreshExternalLinks()).map(link => Object.freeze({ ...link, ...(link.error ? { error: Object.freeze({ ...link.error }) } : {}) })))),
    });
  }
}
