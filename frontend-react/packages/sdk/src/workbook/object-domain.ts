import type { CellValue } from '@react-sheets/core-model';
import type { WorkbookObjectPort } from '@react-sheets/spreadsheet-app';
import type { CellSnapshot, ExternalLinkSnapshot } from './contract';
import type { Workbook } from './workbook';

/** One private identity registry; cell data remains in the canonical runtime. */
export interface WorkbookDomainAccess {
  readonly scope: object;
  sheets(): ReturnType<WorkbookObjectPort['sheets']>;
  invalid(operation: string, cause: unknown, object?: { sheetId?: string; address?: string }): never;
  read(sheetId: string, row: number, column: number): Promise<CellSnapshot>;
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
