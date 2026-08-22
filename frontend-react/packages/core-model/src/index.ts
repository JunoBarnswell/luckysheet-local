export type UnitId = string;
export type SheetId = string;
export type Row = number;
export type Column = number;

export type CellValue = string | number | boolean | null;

export interface CellData {
  value: CellValue;
  formula?: string;
  displayValue?: string;
  styleId?: string;
  numberFormat?: string;
  error?: string;
}

export interface RangeRef {
  sheetId: SheetId;
  startRow: Row;
  endRow: Row;
  startColumn: Column;
  endColumn: Column;
}

export interface MergeSpan {
  range: RangeRef;
  anchor: { row: Row; column: Column };
}

export interface FreezeModel {
  xSplit: number;
  ySplit: number;
  startRow: Row;
  startColumn: Column;
}

export interface SelectionSnapshot {
  unitId: UnitId;
  sheetId: SheetId;
  ranges: RangeRef[];
  primaryRangeIndex: number;
  primaryCell: { row: Row; column: Column };
  phase: 'idle' | 'selected' | 'selecting' | 'editing' | 'preview';
}

export interface ChartModel {
  id: string;
  sheetId: SheetId;
  sourceRanges: RangeRef[];
  type: 'column' | 'line' | 'pie' | 'bar';
  bounds: { x: number; y: number; width: number; height: number };
  title?: string;
}

export interface PivotModel {
  id: string;
  sheetId: SheetId;
  sourceRange: RangeRef;
  rowFields: string[];
  columnFields: string[];
  valueFields: string[];
  filterFields: string[];
}

export interface ShapeModel {
  id: string;
  sheetId: SheetId;
  type: 'rectangle' | 'ellipse' | 'line' | 'arrow';
  bounds: { x: number; y: number; width: number; height: number };
  fill: string;
  stroke: string;
  text?: string;
}

export interface SparklineModel {
  id: string;
  sheetId: SheetId;
  anchor: { row: Row; column: Column };
  sourceRange: RangeRef;
  type: 'line' | 'column' | 'win-loss';
  color: string;
}

export class CellMatrix {
  private readonly rows = new Map<Row, Map<Column, CellData>>();

  get(row: Row, column: Column): CellData | undefined {
    return this.rows.get(row)?.get(column);
  }

  set(row: Row, column: Column, cell: CellData): void {
    let rowMap = this.rows.get(row);
    if (!rowMap) {
      rowMap = new Map<Column, CellData>();
      this.rows.set(row, rowMap);
    }
    rowMap.set(column, cell);
  }

  delete(row: Row, column: Column): void {
    const rowMap = this.rows.get(row);
    rowMap?.delete(column);
    if (rowMap?.size === 0) this.rows.delete(row);
  }

  forEach(callback: (cell: CellData, row: Row, column: Column) => void): void {
    for (const [row, columns] of this.rows) {
      for (const [column, cell] of columns) callback(cell, row, column);
    }
  }

  clone(): CellMatrix {
    const copy = new CellMatrix();
    this.forEach((cell, row, column) => copy.set(row, column, { ...cell }));
    return copy;
  }

  toJSON(): Record<string, Record<string, CellData>> {
    const result: Record<string, Record<string, CellData>> = {};
    this.forEach((cell, row, column) => {
      result[row] ??= {};
      result[row][column] = { ...cell };
    });
    return result;
  }

  static fromJSON(input: Record<string, Record<string, CellData>> | undefined): CellMatrix {
    const matrix = new CellMatrix();
    for (const [row, columns] of Object.entries(input ?? {})) {
      for (const [column, cell] of Object.entries(columns)) {
        matrix.set(Number(row), Number(column), { ...cell });
      }
    }
    return matrix;
  }
}

export class WorksheetModel {
  readonly cells = new CellMatrix();
  readonly merges: MergeSpan[] = [];
  readonly charts: ChartModel[] = [];
  readonly pivots: PivotModel[] = [];
  readonly shapes: ShapeModel[] = [];
  readonly sparklines: SparklineModel[] = [];
  freeze: FreezeModel = { xSplit: 0, ySplit: 0, startRow: 0, startColumn: 0 };

  constructor(
    readonly id: SheetId,
    public name: string,
    public rowCount = 1000,
    public columnCount = 26,
  ) {}
}

export interface WorkbookSnapshotV1 {
  schema: 'WorkbookSnapshotV1';
  unitId: UnitId;
  name: string;
  activeSheetId: SheetId;
  sheets: Array<{
    id: SheetId;
    name: string;
    rowCount: number;
    columnCount: number;
    cells: Record<string, Record<string, CellData>>;
    merges: MergeSpan[];
    freeze: FreezeModel;
    charts: ChartModel[];
    pivots: PivotModel[];
    shapes: ShapeModel[];
    sparklines: SparklineModel[];
  }>;
}

export class WorkbookModel {
  private readonly sheets = new Map<SheetId, WorksheetModel>();
  activeSheetId: SheetId;

  constructor(readonly unitId: UnitId, public name: string) {
    const sheet = new WorksheetModel('sheet-1', 'Sheet1');
    this.sheets.set(sheet.id, sheet);
    this.activeSheetId = sheet.id;
  }

  getSheet(sheetId: SheetId): WorksheetModel {
    const sheet = this.sheets.get(sheetId);
    if (!sheet) throw new Error(`Unknown sheet: ${sheetId}`);
    return sheet;
  }

  getSheets(): WorksheetModel[] {
    return [...this.sheets.values()];
  }

  addSheet(id: SheetId, name: string): WorksheetModel {
    if (this.sheets.has(id)) throw new Error(`Sheet already exists: ${id}`);
    const sheet = new WorksheetModel(id, name);
    this.sheets.set(id, sheet);
    return sheet;
  }

  removeSheet(sheetId: SheetId): WorksheetModel {
    if (this.sheets.size <= 1) throw new Error('A workbook must keep at least one worksheet');
    const sheet = this.getSheet(sheetId);
    this.sheets.delete(sheetId);
    if (this.activeSheetId === sheetId) {
      const firstSheet = this.getSheets()[0];
      if (firstSheet) this.activeSheetId = firstSheet.id;
    }
    return sheet;
  }

  snapshot(): WorkbookSnapshotV1 {
    return {
      schema: 'WorkbookSnapshotV1',
      unitId: this.unitId,
      name: this.name,
      activeSheetId: this.activeSheetId,
      sheets: this.getSheets().map((sheet) => ({
        id: sheet.id,
        name: sheet.name,
        rowCount: sheet.rowCount,
        columnCount: sheet.columnCount,
        cells: sheet.cells.toJSON(),
        merges: structuredClone(sheet.merges),
        freeze: { ...sheet.freeze },
        charts: structuredClone(sheet.charts),
        pivots: structuredClone(sheet.pivots),
        shapes: structuredClone(sheet.shapes),
        sparklines: structuredClone(sheet.sparklines),
      })),
    };
  }

  static fromSnapshot(snapshot: WorkbookSnapshotV1): WorkbookModel {
    if (snapshot.schema !== 'WorkbookSnapshotV1') throw new Error('Unsupported workbook snapshot');
    const workbook = new WorkbookModel(snapshot.unitId, snapshot.name);
    workbook.sheets.clear();
    for (const input of snapshot.sheets) {
      const sheet = new WorksheetModel(input.id, input.name, input.rowCount, input.columnCount);
      const matrix = CellMatrix.fromJSON(input.cells);
      matrix.forEach((cell, row, column) => sheet.cells.set(row, column, cell));
      sheet.merges.push(...structuredClone(input.merges));
      sheet.freeze = { ...input.freeze };
      sheet.charts.push(...structuredClone(input.charts));
      sheet.pivots.push(...structuredClone(input.pivots));
      sheet.shapes.push(...structuredClone(input.shapes));
      sheet.sparklines.push(...structuredClone(input.sparklines));
      workbook.sheets.set(sheet.id, sheet);
    }
    workbook.activeSheetId = snapshot.activeSheetId;
    return workbook;
  }
}
