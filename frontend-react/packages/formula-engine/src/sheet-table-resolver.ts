export interface SheetTableColumnRef {
  readonly id: string;
  readonly name: string;
}

export interface SheetTableRef {
  readonly id: string;
  readonly sheetId: string;
  readonly name: string;
  readonly range: {
    readonly sheetId: string;
    readonly startRow: number;
    readonly endRow: number;
    readonly startColumn: number;
    readonly endColumn: number;
  };
  readonly hasHeaderRow: boolean;
  readonly hasTotalRow: boolean;
  readonly columns: readonly SheetTableColumnRef[];
}
