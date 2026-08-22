export type UnitId = string;
export type SheetId = string;
export type Row = number;
export type Column = number;

export type CellValue = string | number | boolean | null;

export interface CellBorderSide {
  style: 'thin' | 'medium' | 'thick' | 'dashed' | 'double';
  color: string;
}

export interface CellBorders {
  top?: CellBorderSide;
  right?: CellBorderSide;
  bottom?: CellBorderSide;
  left?: CellBorderSide;
}

export interface CellStyle {
  fontFamily?: string;
  fontSize?: number;
  bold?: boolean;
  italic?: boolean;
  underline?: boolean;
  strikethrough?: boolean;
  textColor?: string;
  background?: string;
  horizontalAlignment?: 'left' | 'center' | 'right';
  verticalAlignment?: 'top' | 'middle' | 'bottom';
  wrapText?: boolean;
  numberFormat?: string;
  borders?: CellBorders;
  padding?: number;
}

export interface CellComment {
  id: string;
  author: string;
  text: string;
  createdAt: string;
}

export interface CellData {
  value: CellValue;
  formula?: string;
  displayValue?: string;
  styleId?: string;
  style?: CellStyle;
  numberFormat?: string;
  error?: string;
  comment?: CellComment;
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

export interface ChartSeries {
  name: string;
  range: RangeRef;
  color?: string;
}

export interface ChartModel {
  id: string;
  sheetId: SheetId;
  type: 'column' | 'bar' | 'line' | 'pie' | 'doughnut' | 'area' | 'scatter';
  title?: string;
  sourceRanges: RangeRef[];
  series?: ChartSeries[];
  categoryRange?: RangeRef;
  bounds: { x: number; y: number; width: number; height: number };
  legendPosition?: 'top' | 'bottom' | 'left' | 'right' | 'none';
  showDataLabels?: boolean;
}

export interface PivotValueField {
  field: string;
  summarizeBy: 'sum' | 'count' | 'average' | 'min' | 'max' | 'product';
  displayName?: string;
}

export interface PivotModel {
  id: string;
  sheetId: SheetId;
  sourceRange: RangeRef;
  targetAnchor?: { row: Row; column: Column };
  rowFields: string[];
  columnFields: string[];
  valueFields: PivotValueField[];
  filterFields: string[];
  data?: Record<string, unknown>;
}

export interface ShapeModel {
  id: string;
  sheetId: SheetId;
  type: 'rectangle' | 'rounded-rectangle' | 'ellipse' | 'line' | 'arrow' | 'callout' | 'star';
  bounds: { x: number; y: number; width: number; height: number };
  fill: string;
  stroke: string;
  strokeWidth?: number;
  text?: string;
  textColor?: string;
  fontSize?: number;
  rotation?: number;
}

export interface SparklineModel {
  id: string;
  sheetId: SheetId;
  anchor: { row: Row; column: Column };
  sourceRange: RangeRef;
  type: 'line' | 'column' | 'win-loss';
  color: string;
  negativeColor?: string;
  highlightMax?: boolean;
  highlightMin?: boolean;
}

export type ConditionalFormatType = 'highlight' | 'dataBar' | 'colorScale' | 'iconSet';
export type ConditionalFormatOperator =
  | 'greaterThan'
  | 'lessThan'
  | 'between'
  | 'equal'
  | 'notEqual'
  | 'containsText'
  | 'notContainsText'
  | 'duplicate'
  | 'unique'
  | 'formula';

export interface ConditionalFormatRule {
  id: string;
  sheetId: SheetId;
  ranges: RangeRef[];
  type: ConditionalFormatType;
  operator?: ConditionalFormatOperator;
  value1?: string | number;
  value2?: string | number;
  style?: CellStyle;
  minColor?: string;
  midColor?: string;
  maxColor?: string;
  barColor?: string;
}

export type DataValidationType = 'list' | 'whole' | 'decimal' | 'date' | 'textLength' | 'custom';
export type DataValidationOperator = 'between' | 'notBetween' | 'equal' | 'notEqual' | 'greaterThan' | 'lessThan';

export interface DataValidationRule {
  id: string;
  sheetId: SheetId;
  ranges: RangeRef[];
  type: DataValidationType;
  operator?: DataValidationOperator;
  formula1?: string;
  formula2?: string;
  allowBlank?: boolean;
  showDropdown?: boolean;
  promptTitle?: string;
  promptMessage?: string;
  errorTitle?: string;
  errorMessage?: string;
}

export interface FilterColumnCondition {
  column: Column;
  selectedValues?: string[];
  conditionOperator?: string;
  conditionValue?: string;
}

export interface FilterModel {
  sheetId: SheetId;
  range: RangeRef;
  criteria: Record<Column, FilterColumnCondition>;
}

export interface SortCriterion {
  column: Column;
  ascending: boolean;
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

  has(row: Row, column: Column): boolean {
    return this.rows.get(row)?.has(column) ?? false;
  }

  clear(): void {
    this.rows.clear();
  }

  count(): number {
    let count = 0;
    for (const columns of this.rows.values()) {
      count += columns.size;
    }
    return count;
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
  readonly conditionalFormats: ConditionalFormatRule[] = [];
  readonly dataValidations: DataValidationRule[] = [];
  filter?: FilterModel;
  readonly rowHeights: Record<number, number> = {};
  readonly columnWidths: Record<number, number> = {};
  readonly hiddenRows = new Set<number>();
  readonly hiddenColumns = new Set<number>();
  tabColor?: string;
  freeze: FreezeModel = { xSplit: 0, ySplit: 0, startRow: 0, startColumn: 0 };

  constructor(
    readonly id: SheetId,
    public name: string,
    public rowCount = 1000,
    public columnCount = 26,
  ) {}

  isMerged(row: Row, column: Column): MergeSpan | undefined {
    return this.merges.find(
      (m) =>
        row >= m.range.startRow &&
        row <= m.range.endRow &&
        column >= m.range.startColumn &&
        column <= m.range.endColumn,
    );
  }

  isMergeAnchor(row: Row, column: Column): boolean {
    const merge = this.isMerged(row, column);
    return !merge || (merge.anchor.row === row && merge.anchor.column === column);
  }
}

export interface SheetSnapshotV1 {
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
  conditionalFormats?: ConditionalFormatRule[];
  dataValidations?: DataValidationRule[];
  rowHeights?: Record<number, number>;
  columnWidths?: Record<number, number>;
  hiddenRows?: number[];
  hiddenColumns?: number[];
  tabColor?: string;
}

export interface WorkbookSnapshotV1 {
  schema: 'WorkbookSnapshotV1';
  unitId: UnitId;
  name: string;
  activeSheetId: SheetId;
  definedNames?: Record<string, string>;
  sheets: SheetSnapshotV1[];
}

export class WorkbookModel {
  private readonly sheets = new Map<SheetId, WorksheetModel>();
  activeSheetId: SheetId;
  definedNames: Record<string, string> = {};

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

  getSheetByName(name: string): WorksheetModel | undefined {
    for (const sheet of this.sheets.values()) {
      if (sheet.name.toLowerCase() === name.toLowerCase()) return sheet;
    }
    return undefined;
  }

  getSheets(): WorksheetModel[] {
    return [...this.sheets.values()];
  }

  addSheet(id: SheetId, name: string, rowCount = 1000, columnCount = 26): WorksheetModel {
    if (this.sheets.has(id)) throw new Error(`Sheet already exists: ${id}`);
    const sheet = new WorksheetModel(id, name, rowCount, columnCount);
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
      definedNames: { ...this.definedNames },
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
        conditionalFormats: structuredClone(sheet.conditionalFormats),
        dataValidations: structuredClone(sheet.dataValidations),
        rowHeights: { ...sheet.rowHeights },
        columnWidths: { ...sheet.columnWidths },
        hiddenRows: [...sheet.hiddenRows],
        hiddenColumns: [...sheet.hiddenColumns],
        tabColor: sheet.tabColor,
      })),
    };
  }

  static fromSnapshot(snapshot: WorkbookSnapshotV1): WorkbookModel {
    if (snapshot.schema !== 'WorkbookSnapshotV1') throw new Error('Unsupported workbook snapshot');
    const workbook = new WorkbookModel(snapshot.unitId, snapshot.name);
    workbook.sheets.clear();
    workbook.definedNames = snapshot.definedNames ? { ...snapshot.definedNames } : {};
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
      if (input.conditionalFormats) sheet.conditionalFormats.push(...structuredClone(input.conditionalFormats));
      if (input.dataValidations) sheet.dataValidations.push(...structuredClone(input.dataValidations));
      if (input.rowHeights) Object.assign(sheet.rowHeights, input.rowHeights);
      if (input.columnWidths) Object.assign(sheet.columnWidths, input.columnWidths);
      if (input.hiddenRows) input.hiddenRows.forEach((r) => sheet.hiddenRows.add(r));
      if (input.hiddenColumns) input.hiddenColumns.forEach((c) => sheet.hiddenColumns.add(c));
      sheet.tabColor = input.tabColor;
      workbook.sheets.set(sheet.id, sheet);
    }
    workbook.activeSheetId = snapshot.activeSheetId;
    return workbook;
  }
}
