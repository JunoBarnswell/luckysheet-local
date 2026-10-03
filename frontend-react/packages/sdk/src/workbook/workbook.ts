import { cellAddress } from '@react-sheets/core-model';
import { CommandDispatchError, type WorkbookObjectPort } from '@react-sheets/spreadsheet-app';
import { SdkError } from '../error';
import { domainFor, registerWorkbookDomain } from './object-domain';
import { WorksheetCollection } from './worksheet-collection';
import { WorkbookExternalLinks } from './external-links';

export class Workbook {
  readonly id: string;
  readonly worksheets: WorksheetCollection;
  readonly externalLinks: WorkbookExternalLinks;
  #closed = false;
  #closeReady: (() => void) | null = null;
  #opening: Promise<Workbook> | null = null;
  constructor(readonlyPort: WorkbookObjectPort, scope: object, close: () => void) {
    this.#port = readonlyPort;
    this.#release = close;
    this.id = readonlyPort.unitId;
    this.worksheets = new WorksheetCollection(this);
    this.externalLinks = new WorkbookExternalLinks(this);
    this.#initializeDomain(scope);
    Object.freeze(this);
  }
  readonly #port: WorkbookObjectPort;
  readonly #release: () => void;
  get name(): string { this.#assertAlive('workbook.name'); return this.#port.state().name; }
  #assertAlive(operation: string): void {
    if (this.#closed) throw this.#error('RUNTIME_DISPOSED', operation, '工作簿对象已关闭。', undefined);
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
    this.#closeReady?.();
    this.#release();
  }
  async save(): Promise<void> { await this.#perform('workbook.save', () => this.#port.save()); }
  async flush(): Promise<void> { await this.#perform('workbook.flush', () => this.#port.flush()); }
  async #perform<T>(operation: string, action: () => Promise<T>, object?: { sheetId?: string; address?: string }): Promise<T> {
    try { this.#assertAlive(operation); const result = await action(); this.#assertAlive(operation); return result; }
    catch (cause) {
      if (cause instanceof SdkError) throw cause;
      const code = cause instanceof Error && 'code' in cause && cause.code === 'CIRCULAR_DEPENDENCY' ? 'CIRCULAR_DEPENDENCY' : cause instanceof CommandDispatchError && cause.code === 'PERMISSION_DENIED' ? 'FORBIDDEN' : 'REQUEST_REJECTED';
      throw this.#error(code, operation, cause instanceof Error ? cause.message : '操作失败。', cause, object);
    }
  }
  #initializeDomain(scope: object): void {
    registerWorkbookDomain(this, {
      scope,
      sheets: () => { this.#assertAlive('worksheets.read'); return this.#port.sheets(); },
      invalid: (operation, cause, object) => { this.#assertAlive(operation); throw this.#error('INVALID_ARGUMENT', operation, cause instanceof Error ? cause.message : String(cause), cause, object); },
      read: (sheetId, row, column) => this.#perform('cell.read', async () => {
        const resolved = await this.#port.readCell(sheetId, row, column);
        const result = { value: resolved.cell?.value ?? null,
          ...(resolved.cell?.formula === undefined ? {} : { formula: resolved.cell.formula }),
          calculatedValue: structuredClone(resolved.calculatedValue), formulaHidden: resolved.formulaHidden };
        if (result.calculatedValue && typeof result.calculatedValue === 'object') Object.freeze(result.calculatedValue);
        return Object.freeze(result);
      }, { sheetId, address: cellAddress(row, column) }),
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
        const outcome = await this.#port.dispatch(descriptor);
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
