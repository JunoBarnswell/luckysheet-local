import type { CellData, ExternalLinkBinding, FormulaValue, RecordFieldAddress, SheetKind, RangeRef, WorksheetPane, MergeSpan, DefinedNameModel, ProtectionRule } from '@react-sheets/core-model';
import type { CellInputInterpretationContext } from '@react-sheets/sheet-features';
import type { CommandDescriptor } from '@react-sheets/command-runtime';
import type { DispatchOutcome } from './workbook-session';

/** Internal access to the existing canonical runtime, never exported by SDK. */
export interface WorkbookObjectPort {
  readonly unitId: string;
  state(): { phase: string; notice: string; name: string; disposed: boolean };
  subscribeDisposed(listener: () => void): () => void;
  sheets(): readonly { id: string; name: string; kind: SheetKind; rowCount: number; columnCount: number; hidden: boolean; pane: WorksheetPane }[];
  readDefinedNames(): readonly DefinedNameModel[];
  readWorksheet(sheetId: string): WorkbookWorksheetRead;
  readCell(sheetId: string, row: number, column: number): Promise<WorkbookCellRead>;
  readCells(range: RangeRef): Promise<readonly WorkbookCellRead[]>;
  history(direction: 'undo' | 'redo'): boolean;
  dispatch(descriptor: CommandDescriptor): Promise<DispatchOutcome>;
  subscribe(listener: () => void): () => void;
  flush(): Promise<void>;
  save(): Promise<void>;
  bindExternalLink(binding: ExternalLinkBinding): Promise<void>;
  refreshExternalLinks(): Promise<readonly { token: string; state: string; sourceRevision: number; error?: { code: string; message: string } }[]>;
}
export interface WorkbookWorksheetRead {
  id: string; name: string; kind: SheetKind; rowCount: number; columnCount: number; hidden: boolean; pane: WorksheetPane;
  defaultRowHeightPx: number; defaultColumnWidthPx: number;
  rowHeightsPx: Record<number, number>; columnWidthsPx: Record<number, number>;
  hiddenRows: number[]; hiddenColumns: number[]; merges: MergeSpan[]; protectionRules: ProtectionRule[];
}
export interface WorkbookCellRead {
    sheetId: string; row: number; column: number; cell?: CellData;
    calculatedValue: FormulaValue; writable: boolean; formulaHidden: boolean;
    recordField?: RecordFieldAddress; inputContext: CellInputInterpretationContext;
}

const ports = new WeakMap<object, WorkbookObjectPort>();
export function registerWorkbookObjectPort(owner: object, port: WorkbookObjectPort): void {
  if (ports.has(owner)) throw new Error('Workbook object port already registered');
  ports.set(owner, port);
}
export function getWorkbookObjectPort(owner: object): WorkbookObjectPort {
  const port = ports.get(owner);
  if (!port) throw new Error('Workbook object runtime is unavailable');
  return port;
}
