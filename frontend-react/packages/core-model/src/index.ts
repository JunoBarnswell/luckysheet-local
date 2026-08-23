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
  textRotate?: number;
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
  hyperlink?: string;
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
  pivotId?: string;
  type: 'column' | 'bar' | 'line' | 'pie' | 'doughnut' | 'area' | 'scatter';
  title?: string;
  sourceRanges: RangeRef[];
  series?: ChartSeries[];
  categoryRange?: RangeRef;
  bounds: { x: number; y: number; width: number; height: number };
  legendPosition?: 'top' | 'bottom' | 'left' | 'right' | 'none';
  showDataLabels?: boolean;
}

import type { PivotModel } from './pivot';
export * from './pivot';

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

/** 浮动图片(以内容坐标定位) */
export interface FloatingImage {
  id: string;
  sheetId: SheetId;
  name?: string;
  src: string;
  bounds: { x: number; y: number; width: number; height: number };
}

/** 隔行色带规则 */
export interface BandedRule {
  range: RangeRef;
  firstColor: string;
  secondColor: string;
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

  /** 沿行轴整体平移:dir=+1 下移(插入),dir=-1 上移(删除);越界丢弃 */
  shiftRows(at: Row, count: number, direction: 1 | -1): void {
    const entries: Array<[Row, Column, CellData]> = [];
    const delta = direction * count;
    this.forEach((cell, row, column) => {
      if (row >= at) entries.push([row, column, cell]);
    });
    if (direction === -1) {
      // 从小到大删除,避免覆盖
      entries.sort((a, b) => a[0] - b[0]);
    } else {
      entries.sort((a, b) => b[0] - a[0]);
    }
    for (const [row, column] of entries) this.delete(row, column);
    for (const [row, column, cell] of entries) {
      this.set(row + delta, column, cell);
    }
  }

  /** 沿列轴整体平移:dir=+1 右移(插入),dir=-1 左移(删除) */
  shiftColumns(at: Column, count: number, direction: 1 | -1): void {
    const entries: Array<[Row, Column, CellData]> = [];
    const delta = direction * count;
    this.forEach((cell, row, column) => {
      if (column >= at) entries.push([row, column, cell]);
    });
    if (direction === -1) entries.sort((a, b) => a[1] - b[1]);
    else entries.sort((a, b) => b[1] - a[1]);
    for (const [row, column] of entries) this.delete(row, column);
    for (const [row, column, cell] of entries) {
      this.set(row, column + delta, cell);
    }
  }

  /** 摘除区间内全部单元格并返回(用于删除行的逆操作恢复) */
  extractRegion(startRow: Row, endRow: Row, startColumn: Column, endColumn: Column): Array<{ row: Row; column: Column; cell: CellData }> {
    const extracted: Array<{ row: Row; column: Column; cell: CellData }> = [];
    this.forEach((cell, row, column) => {
      if (row >= startRow && row <= endRow && column >= startColumn && column <= endColumn) {
        extracted.push({ row, column, cell: structuredClone(cell) });
      }
    });
    for (const item of extracted) this.delete(item.row, item.column);
    return extracted;
  }

  placeRegion(items: ReadonlyArray<{ row: Row; column: Column; cell: CellData }>): void {
    for (const item of items) this.set(item.row, item.column, structuredClone(item.cell));
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
  readonly images: FloatingImage[] = [];
  filter?: FilterModel;
  bandedRule?: BandedRule;
  readonly rowHeights: Record<number, number> = {};
  readonly columnWidths: Record<number, number> = {};
  readonly hiddenRows = new Set<number>();
  readonly hiddenColumns = new Set<number>();
  tabColor?: string;
  freeze: FreezeModel = { xSplit: 0, ySplit: 0, startRow: 0, startColumn: 0 };

  /**
   * 在 at 行前插入 count 行:单元格/合并/隐藏集/冻结/筛选/条件格式等全部随动。
   */
  insertRows(at: Row, count: number): void {
    if (count <= 0) return;
    this.cells.shiftRows(at, count, 1);
    for (const merge of this.merges) {
      merge.range.startRow = merge.range.startRow >= at ? merge.range.startRow + count : merge.range.startRow;
      merge.range.endRow = merge.range.endRow >= at ? merge.range.endRow + count : merge.range.endRow;
      if (merge.anchor.row >= at) merge.anchor.row += count;
    }
    for (const rule of [...this.conditionalFormats, ...this.dataValidations]) {
      for (const range of rule.ranges) {
        if (range.startRow >= at) { range.startRow += count; range.endRow += count; }
        else if (range.endRow >= at) range.endRow += count;
      }
    }
    for (const sparkline of this.sparklines) {
      if (sparkline.anchor.row >= at) sparkline.anchor.row += count;
    }
    const shiftedHidden = new Set<number>();
    for (const row of this.hiddenRows) shiftedHidden.add(row >= at ? row + count : row);
    this.hiddenRows.clear();
    for (const row of shiftedHidden) this.hiddenRows.add(row);
    const shiftedHeights: Record<number, number> = {};
    for (const [key, value] of Object.entries(this.rowHeights)) {
      const row = Number(key);
      shiftedHeights[row >= at ? row + count : row] = value;
    }
    for (const key of Object.keys(this.rowHeights)) delete this.rowHeights[Number(key)];
    Object.assign(this.rowHeights, shiftedHeights);
    if (this.freeze.ySplit >= at) this.freeze.ySplit += count;
    if (this.filter) {
      if (this.filter.range.startRow >= at) this.filter.range.startRow += count;
      this.filter.range.endRow += count;
    }
    if (this.bandedRule && this.bandedRule.range.startRow >= at) {
      this.bandedRule.range.startRow += count;
      this.bandedRule.range.endRow += count;
    }
    this.rowCount += count;
  }

  /** 删除 [at, at+count) 行;返回被摘除的单元格以便撤销恢复 */
  deleteRows(at: Row, count: number): Array<{ row: Row; column: Column; cell: CellData }> {
    if (count <= 0) return [];
    const end = at + count - 1;
    const removed = this.cells.extractRegion(at, end, 0, Math.max(this.columnCount - 1, 0));
    this.cells.shiftRows(end + 1, count, -1);
    for (let index = this.merges.length - 1; index >= 0; index--) {
      const merge = this.merges[index]!;
      if (merge.range.startRow > end || merge.range.endRow < at) continue;
      merge.range.endRow = Math.max(at - 1, merge.range.endRow - count);
      merge.range.startRow = Math.min(Math.max(at - 1, merge.range.startRow), merge.range.endRow);
      if (merge.range.endRow < merge.range.startRow) this.merges.splice(index, 1);
    }
    for (const rule of [...this.conditionalFormats, ...this.dataValidations]) {
      for (const range of rule.ranges) {
        if (range.endRow < at) continue;
        if (range.startRow > end) { range.startRow -= count; range.endRow -= count; }
        else range.endRow = Math.max(at - 1, range.endRow - count);
      }
      rule.ranges = rule.ranges.filter((range) => range.endRow >= range.startRow);
    }
    for (let index = this.sparklines.length - 1; index >= 0; index--) {
      const anchorRow = this.sparklines[index]!.anchor.row;
      if (anchorRow >= at && anchorRow <= end) this.sparklines.splice(index, 1);
      else if (anchorRow > end) this.sparklines[index]!.anchor.row -= count;
    }
    const shiftedHidden = new Set<number>();
    for (const row of this.hiddenRows) {
      if (row < at) shiftedHidden.add(row);
      else if (row > end) shiftedHidden.add(row - count);
    }
    this.hiddenRows.clear();
    for (const row of shiftedHidden) this.hiddenRows.add(row);
    const shiftedHeights: Record<number, number> = {};
    for (const [key, value] of Object.entries(this.rowHeights)) {
      const row = Number(key);
      if (row < at) shiftedHeights[row] = value;
      else if (row > end) shiftedHeights[row - count] = value;
    }
    for (const key of Object.keys(this.rowHeights)) delete this.rowHeights[Number(key)];
    Object.assign(this.rowHeights, shiftedHeights);
    if (this.freeze.ySplit > at) this.freeze.ySplit = Math.max(0, this.freeze.ySplit - count);
    if (this.filter) {
      if (this.filter.range.startRow > end) this.filter.range.startRow -= count;
      this.filter.range.endRow = Math.max(at, this.filter.range.endRow - count);
    }
    this.rowCount = Math.max(1, this.rowCount - count);
    return removed;
  }

  insertColumns(at: Column, count: number): void {
    if (count <= 0) return;
    this.cells.shiftColumns(at, count, 1);
    for (const merge of this.merges) {
      merge.range.startColumn = merge.range.startColumn >= at ? merge.range.startColumn + count : merge.range.startColumn;
      merge.range.endColumn = merge.range.endColumn >= at ? merge.range.endColumn + count : merge.range.endColumn;
      if (merge.anchor.column >= at) merge.anchor.column += count;
    }
    for (const rule of [...this.conditionalFormats, ...this.dataValidations]) {
      for (const range of rule.ranges) {
        if (range.startColumn >= at) { range.startColumn += count; range.endColumn += count; }
        else if (range.endColumn >= at) range.endColumn += count;
      }
    }
    const shiftedHidden = new Set<number>();
    for (const column of this.hiddenColumns) shiftedHidden.add(column >= at ? column + count : column);
    this.hiddenColumns.clear();
    for (const column of shiftedHidden) this.hiddenColumns.add(column);
    const shiftedWidths: Record<number, number> = {};
    for (const [key, value] of Object.entries(this.columnWidths)) {
      const column = Number(key);
      shiftedWidths[column >= at ? column + count : column] = value;
    }
    for (const key of Object.keys(this.columnWidths)) delete this.columnWidths[Number(key)];
    Object.assign(this.columnWidths, shiftedWidths);
    if (this.freeze.xSplit >= at) this.freeze.xSplit += count;
    if (this.filter) {
      if (this.filter.range.startColumn >= at) this.filter.range.startColumn += count;
      this.filter.range.endColumn += count;
    }
    this.columnCount += count;
  }

  deleteColumns(at: Column, count: number): Array<{ row: Row; column: Column; cell: CellData }> {
    if (count <= 0) return [];
    const end = at + count - 1;
    const removed = this.cells.extractRegion(0, Math.max(this.rowCount - 1, 0), at, end);
    this.cells.shiftColumns(end + 1, count, -1);
    for (let index = this.merges.length - 1; index >= 0; index--) {
      const merge = this.merges[index]!;
      if (merge.range.startColumn > end || merge.range.endColumn < at) continue;
      merge.range.endColumn = Math.max(at - 1, merge.range.endColumn - count);
      merge.range.startColumn = Math.min(Math.max(at - 1, merge.range.startColumn), merge.range.endColumn);
      if (merge.range.endColumn < merge.range.startColumn) this.merges.splice(index, 1);
    }
    const shiftedHidden = new Set<number>();
    for (const column of this.hiddenColumns) {
      if (column < at) shiftedHidden.add(column);
      else if (column > end) shiftedHidden.add(column - count);
    }
    this.hiddenColumns.clear();
    for (const column of shiftedHidden) this.hiddenColumns.add(column);
    const shiftedWidths: Record<number, number> = {};
    for (const [key, value] of Object.entries(this.columnWidths)) {
      const column = Number(key);
      if (column < at) shiftedWidths[column] = value;
      else if (column > end) shiftedWidths[column - count] = value;
    }
    for (const key of Object.keys(this.columnWidths)) delete this.columnWidths[Number(key)];
    Object.assign(this.columnWidths, shiftedWidths);
    if (this.freeze.xSplit > at) this.freeze.xSplit = Math.max(0, this.freeze.xSplit - count);
    if (this.filter) {
      if (this.filter.range.startColumn > end) this.filter.range.startColumn -= count;
      this.filter.range.endColumn = Math.max(at, this.filter.range.endColumn - count);
    }
    this.columnCount = Math.max(1, this.columnCount - count);
    return removed;
  }

  /** 深拷贝当前工作表(删除工作表撤销恢复用) */
  cloneSheet(): WorksheetModel {
    const copy = new WorksheetModel(this.id, this.name, this.rowCount, this.columnCount);
    this.cells.forEach((cell, row, column) => copy.cells.set(row, column, structuredClone(cell)));
    copy.merges.push(...structuredClone(this.merges));
    copy.charts.push(...structuredClone(this.charts));
    copy.pivots.push(...structuredClone(this.pivots));
    copy.shapes.push(...structuredClone(this.shapes));
    copy.sparklines.push(...structuredClone(this.sparklines));
    copy.conditionalFormats.push(...structuredClone(this.conditionalFormats));
    copy.dataValidations.push(...structuredClone(this.dataValidations));
    copy.images.push(...structuredClone(this.images));
    copy.filter = this.filter ? structuredClone(this.filter) : undefined;
    copy.bandedRule = this.bandedRule ? structuredClone(this.bandedRule) : undefined;
    Object.assign(copy.rowHeights, this.rowHeights);
    Object.assign(copy.columnWidths, this.columnWidths);
    for (const row of this.hiddenRows) copy.hiddenRows.add(row);
    for (const column of this.hiddenColumns) copy.hiddenColumns.add(column);
    copy.tabColor = this.tabColor;
    copy.freeze = { ...this.freeze };
    return copy;
  }

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
  images?: FloatingImage[];
  bandedRule?: BandedRule;
  filter?: FilterModel;
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
  readonly sheets = new Map<SheetId, WorksheetModel>();
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
        images: structuredClone(sheet.images),
        bandedRule: sheet.bandedRule ? structuredClone(sheet.bandedRule) : undefined,
        filter: sheet.filter ? structuredClone(sheet.filter) : undefined,
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
      if (input.images) sheet.images.push(...structuredClone(input.images));
      if (input.bandedRule) sheet.bandedRule = structuredClone(input.bandedRule);
      if (input.filter) sheet.filter = structuredClone(input.filter);
      sheet.tabColor = input.tabColor;
      workbook.sheets.set(sheet.id, sheet);
    }
    workbook.activeSheetId = snapshot.activeSheetId;
    return workbook;
  }
}
