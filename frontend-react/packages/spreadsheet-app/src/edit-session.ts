import type { CellData, Row, Column, SheetId } from '@react-sheets/core-model';
import type { SelectionSnapshot } from './selection-service';

export interface EditSessionState {
  sheetId: SheetId;
  row: Row;
  column: Column;
  originalValue: CellData | null;
  originalFormula?: string;
  originalSelection: SelectionSnapshot;
  currentDraft: string;
  referenceMode: boolean;
  isDirty: boolean;
}

export class EditSession {
  private session: EditSessionState | null = null;

  get active(): EditSessionState | null {
    return this.session;
  }

  get editingCell(): { row: number; column: number } | null {
    if (!this.session) return null;
    return { row: this.session.row, column: this.session.column };
  }

  begin(params: {
    sheetId: SheetId;
    row: Row;
    column: Column;
    cell: CellData | undefined;
    selection: SelectionSnapshot;
    initialText?: string;
  }): void {
    const raw = params.initialText ?? params.cell?.formula ?? (params.cell?.value == null ? '' : String(params.cell.value));
    this.session = {
      sheetId: params.sheetId,
      row: params.row,
      column: params.column,
      originalValue: params.cell ? structuredClone(params.cell) : null,
      originalFormula: params.cell?.formula,
      originalSelection: structuredClone(params.selection),
      currentDraft: raw,
      referenceMode: false,
      isDirty: false,
    };
  }

  setDraft(value: string): void {
    if (!this.session) return;
    this.session.currentDraft = value;
    this.session.isDirty = value !== (this.session.originalFormula ?? String(this.session.originalValue?.value ?? ''));
  }

  insertRef(refText: string): void {
    if (!this.session) return;
    this.setDraft(this.session.currentDraft + refText);
    this.session.referenceMode = true;
  }

  toggleAbsoluteReference(): void {
    if (!this.session) return;
    this.setDraft(
      this.session.currentDraft.replace(/(\$?)([A-Za-z]+)(\$?)(\d+)/g, (_match, dCol, col, dRow, row) => {
        const nextDCol = dCol ? '' : '$';
        const nextDRow = dRow ? '' : '$';
        return `${nextDCol}${col}${nextDRow}${row}`;
      }),
    );
  }

  cancel(): string {
    if (!this.session) return '';
    const restore = this.session.originalFormula ?? (this.session.originalValue?.value == null ? '' : String(this.session.originalValue.value));
    this.session = null;
    return restore;
  }

  apply(): { row: number; column: number; draft: string } | null {
    if (!this.session) return null;
    const result = { row: this.session.row, column: this.session.column, draft: this.session.currentDraft };
    this.session = null;
    return result;
  }
}
