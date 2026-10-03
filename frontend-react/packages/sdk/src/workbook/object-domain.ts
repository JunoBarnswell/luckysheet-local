import type { CellValue, RangeRef, RichTextRun } from '@react-sheets/core-model';
import type { CommandDescriptor } from '@react-sheets/command-runtime';
import type { WorkbookObjectPort } from '@react-sheets/spreadsheet-app';
import type { CellInput, CellSnapshot, ExternalLinkSnapshot } from './contract';
import type { Workbook } from './workbook';

/** One private identity registry; cell data remains in the canonical runtime. */
export interface WorkbookDomainAccess {
  readonly scope: object;
  sheet(sheetId: string): ReturnType<WorkbookObjectPort['readWorksheet']>;
  names(): ReturnType<WorkbookObjectPort['readDefinedNames']>;
  writeRichText(range: RangeRef, text: string, runs: readonly RichTextRun[]): Promise<void>;
  sheets(): ReturnType<WorkbookObjectPort['sheets']>;
  invalid(operation: string, cause: unknown, object?: { sheetId?: string; address?: string }): never;
  read(sheetId: string, row: number, column: number): Promise<CellSnapshot>;
  readRange(range: RangeRef): Promise<readonly (readonly CellSnapshot[])[]>;
  writeRange(range: RangeRef, inputs: readonly (readonly CellInput[])[]): Promise<void>;
  command(operation: string, descriptor: CommandDescriptor): Promise<void>;
  history(direction: 'undo' | 'redo'): Promise<boolean>;
  write(sheetId: string, row: number, column: number, input: { kind: 'value'; value: CellValue } | { kind: 'formula'; formula: string }): Promise<void>;
  bind(source: Workbook, token: string): Promise<void>;
  refresh(): Promise<readonly ExternalLinkSnapshot[]>;
}
const domains = new WeakMap<Workbook, WorkbookDomainAccess>();
export function registerWorkbookDomain(workbook: Workbook, domain: WorkbookDomainAccess): void {
  if (domains.has(workbook)) throw new Error('Workbook domain already registered');
  domains.set(workbook, domain);
}
export function domainFor(workbook: Workbook): WorkbookDomainAccess {
  const domain = domains.get(workbook);
  if (!domain) throw new Error('Workbook domain is not initialized');
  return domain;
}
