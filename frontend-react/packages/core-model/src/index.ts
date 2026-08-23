export type UnitId = string;
export type SheetId = string;
export type Row = number;
export type Column = number;

import type {
  CellHyperlink,
  CellNote,
  CommentThread,
  DrawingObject,
  DrawingPayload,
  SparklineGroup,
  SheetTableModel,
  OutlineModel,
  SpillRange,
  ProtectionRule,
  DefinedNameModel,
  DefinedNameScope,
} from './domain';
import { normalizeDefinedNameModel } from './domain';
import type { WorkbookSnapshot } from './snapshot';
import {
  normalizePrintDocumentSnapshot,
  normalizeQueryDefinitionSnapshot,
  type PrintDocumentSnapshot,
  type QueryDefinitionSnapshot,
} from './workbook-state';

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
  mentions?: string[];
  replies?: CellCommentReply[];
  resolved?: boolean;
  resolvedAt?: string;
}

export interface CellCommentReply {
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
  /** 公式引擎结果（含错误）。禁止再用 error: string 当真相 */
  formulaValue?: import('./domain').FormulaValue;
  note?: import('./domain').CellNote;
  comment?: CellComment;
  /** @deprecated prefer hyperlinkDetail */
  hyperlink?: string;
  hyperlinkDetail?: CellHyperlink;
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

export type {
  SelectionSnapshot,
  EditSession,
  SheetTableModel,
  OutlineGroup,
  OutlineModel,
  DrawingKind,
  DrawingTransform,
  DrawingObject,
  CellHyperlink,
  HyperlinkTarget,
  DrawingPayload,
  ImageDrawingPayload,
  ShapeDrawingPayload,
  TextBoxDrawingPayload,
  ChartDrawingPayload,
  SparklineGroup,
  CellNote,
  CommentThread,
  CommentReply,
  SpillRange,
  SpillState,
  ProtectionRule,
  DefinedNameModel,
  DefinedNameScope,
  ProtectionScope,
  FormulaErrorCode,
  FormulaValue,
  StructuralOpKind,
  StructuralTransformParams,
} from './domain';
export { createEmptySelection, isFormulaError, createFormulaError, normalizeDefinedNameModel } from './domain';
export { StructuralTransform, type StructuralTransformResult, ensureDrawing } from './structural-transform';
export { applyRowPermutation, validatePermutationMetadata, type RowPermutation } from './data-transform';
export { columnLabel, parseColumnLabel, cellAddress, parseAddress, a1Range } from './address';
export {
  loadWorkbookFromSnapshot,
  createWorkbookSnapshot,
  type WorkbookSnapshot,
} from './snapshot';
export {
  normalizePrintDocumentSnapshot,
  normalizeQueryDefinitionSnapshot,
  type PrintDocumentSnapshot,
  type QueryDefinitionSnapshot,
  type QueryStepSnapshot,
} from './workbook-state';

import type { PivotModel } from './pivot';
export * from './pivot';
import type { WorkbookTableModel } from './data-model';
export * from './data-model';

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
  highlightFirst?: boolean;
  highlightLast?: boolean;
  highlightNegative?: boolean;
  groupId?: string;
  showAxis?: boolean;
  showMarkers?: boolean;
}

/** 隔行色带规则 */
export interface BandedRule {
  range: RangeRef;
  firstColor: string;
  secondColor: string;
}

export type ConditionalFormatType = 'highlight' | 'dataBar' | 'colorScale' | 'iconSet' | 'topBottom';
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
  | 'formula'
  | 'top'
  | 'bottom';

export interface ConditionalFormatTopBottom {
  direction: 'top' | 'bottom';
  /** Number of values, or a percentage when `percent` is true. */
  rank: number;
  percent?: boolean;
}

export interface ConditionalFormatRule {
  id: string;
  sheetId: SheetId;
  ranges: RangeRef[];
  type: ConditionalFormatType;
  /** Lower values are evaluated first. Excel defaults to the insertion order. */
  priority?: number;
  /** Stop evaluating lower-priority rules after this rule matches a cell. */
  stopIfTrue?: boolean;
  operator?: ConditionalFormatOperator;
  value1?: string | number;
  value2?: string | number;
  style?: CellStyle;
  minColor?: string;
  midColor?: string;
  maxColor?: string;
  barColor?: string;
  topBottom?: ConditionalFormatTopBottom;
}

export type DataValidationType = 'list' | 'whole' | 'decimal' | 'date' | 'time' | 'checkbox' | 'textLength' | 'custom';
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
  /** Excel error alert style. Only STOP blocks a write. */
  alertStyle?: 'stop' | 'warning' | 'information';
  showErrorMessage?: boolean;
  showInputMessage?: boolean;
  inputTitle?: string;
  inputMessage?: string;
  showDropdown?: boolean;
  /** Allows comma-separated values for list validation when enabled. */
  multiSelect?: boolean;
  listSource?:
    | { kind: 'values'; values: string[] }
    | { kind: 'range'; range: RangeRef }
    | { kind: 'formula'; formula: string };
  promptTitle?: string;
  promptMessage?: string;
  errorTitle?: string;
  errorMessage?: string;
}

export interface FilterColumnCondition {
  column: Column;
  selectedValues?: string[];
  excludeBlanks?: boolean;
  conditionOperator?: string;
  conditionValue?: string;
  conditionValue2?: string;
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
  readonly pivots: PivotModel[] = [];
  readonly sparklines: SparklineModel[] = [];
  readonly conditionalFormats: ConditionalFormatRule[] = [];
  readonly dataValidations: DataValidationRule[] = [];
  readonly sheetTables: SheetTableModel[] = [];
  readonly drawings: DrawingObject[] = [];
  readonly drawingPayloads = new Map<string, DrawingPayload>();
  readonly notes = new Map<string, CellNote>();
  readonly commentThreads: CommentThread[] = [];
  readonly spillRanges: SpillRange[] = [];
  readonly protectionRules: ProtectionRule[] = [];
  readonly sparklineGroups: SparklineGroup[] = [];
  outline?: OutlineModel;
  showGridlines = true;
  showHeaders = true;
  zoom = 100;
  hidden = false;
  filter?: FilterModel;
  bandedRule?: BandedRule;
  readonly rowHeights: Record<number, number> = {};
  readonly columnWidths: Record<number, number> = {};
  readonly hiddenRows = new Set<number>();
  readonly hiddenColumns = new Set<number>();
  tabColor?: string;
  freeze: FreezeModel = { xSplit: 0, ySplit: 0, startRow: 0, startColumn: 0 };

  /** 深拷贝当前工作表(删除工作表撤销恢复用) */
  cloneSheet(): WorksheetModel {
    return this.cloneWithIdentity(this.id, this.name);
  }

  cloneWithIdentity(id: SheetId, name: string): WorksheetModel {
    const copy = new WorksheetModel(id, name, this.rowCount, this.columnCount);
    this.cells.forEach((cell, row, column) => copy.cells.set(row, column, structuredClone(cell)));
    copy.merges.push(...structuredClone(this.merges));
    copy.pivots.push(...structuredClone(this.pivots));
    copy.sparklines.push(...structuredClone(this.sparklines));
    copy.conditionalFormats.push(...structuredClone(this.conditionalFormats));
    copy.dataValidations.push(...structuredClone(this.dataValidations));
    copy.filter = this.filter ? structuredClone(this.filter) : undefined;
    copy.bandedRule = this.bandedRule ? structuredClone(this.bandedRule) : undefined;
    Object.assign(copy.rowHeights, this.rowHeights);
    Object.assign(copy.columnWidths, this.columnWidths);
    for (const row of this.hiddenRows) copy.hiddenRows.add(row);
    for (const column of this.hiddenColumns) copy.hiddenColumns.add(column);
    copy.sheetTables.push(...structuredClone(this.sheetTables));
    copy.drawings.push(...structuredClone(this.drawings));
    for (const [key, payload] of this.drawingPayloads) copy.drawingPayloads.set(key, structuredClone(payload));
    for (const [key, note] of this.notes) copy.notes.set(key, structuredClone(note));
    copy.commentThreads.push(...structuredClone(this.commentThreads));
    copy.spillRanges.push(...structuredClone(this.spillRanges));
    copy.protectionRules.push(...structuredClone(this.protectionRules));
    copy.sparklineGroups.push(...structuredClone(this.sparklineGroups));
    copy.outline = this.outline ? structuredClone(this.outline) : undefined;
    copy.showGridlines = this.showGridlines;
    copy.showHeaders = this.showHeaders;
    copy.zoom = this.zoom;
    copy.hidden = this.hidden;
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

export function noteCellKey(row: Row, column: Column): string {
  return `${row}:${column}`;
}

export function getDrawingPayload(sheet: WorksheetModel, payloadId: string): DrawingPayload | undefined {
  return sheet.drawingPayloads.get(payloadId);
}

export function getCellNote(sheet: WorksheetModel, row: Row, column: Column): CellNote | undefined {
  return sheet.notes.get(noteCellKey(row, column));
}

export interface SheetSnapshot {
  id: SheetId;
  name: string;
  rowCount: number;
  columnCount: number;
  cells: Record<string, Record<string, CellData>>;
  merges: MergeSpan[];
  freeze: FreezeModel;
  pivots: PivotModel[];
  sparklines: SparklineModel[];
  sparklineGroups?: SparklineGroup[];
  /** Canonical floating-object collection. Legacy per-kind collections are not part of snapshots. */
  drawings: DrawingObject[];
  drawingPayloads: Record<string, DrawingPayload>;
  notes?: Array<{ row: number; column: number; note: CellNote }>;
  commentThreads?: CommentThread[];
  conditionalFormats?: ConditionalFormatRule[];
  dataValidations?: DataValidationRule[];
  rowHeights?: Record<number, number>;
  columnWidths?: Record<number, number>;
  hiddenRows?: number[];
  hiddenColumns?: number[];
  tabColor?: string;
  bandedRule?: BandedRule;
  filter?: FilterModel;
  sheetTables?: SheetTableModel[];
  spillRanges?: SpillRange[];
  protectionRules?: ProtectionRule[];
  showGridlines?: boolean;
  showHeaders?: boolean;
  zoom?: number;
  hidden?: boolean;
  outline?: OutlineModel;
}

export class WorkbookModel {
  readonly sheets = new Map<SheetId, WorksheetModel>();
  readonly tables = new Map<string, WorkbookTableModel>();
  /** Canonical workbook-owned print state; no host-side cache is authoritative. */
  readonly printDocuments = new Map<SheetId, PrintDocumentSnapshot>();
  /** Persistence-safe query definitions; connector credentials are redacted. */
  readonly queryDefinitions = new Map<string, QueryDefinitionSnapshot>();
  /** 工作表 Tab 顺序 */
  sheetOrder: SheetId[] = [];
  activeSheetId: SheetId;
  /** Canonical scoped name definitions. `definedNames` remains the workbook-level formula view used by the formula runtime. */
  readonly definedNameModels: DefinedNameModel[] = [];
  /** @deprecated use definedNameModels and setDefinedName/getDefinedName. */
  definedNames: Record<string, string> = {};

  constructor(readonly unitId: UnitId, public name: string) {
    const sheet = new WorksheetModel('sheet-1', 'Sheet1');
    this.sheets.set(sheet.id, sheet);
    this.sheetOrder = [sheet.id];
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
    return this.sheetOrder
      .map((id) => this.sheets.get(id))
      .filter((sheet): sheet is WorksheetModel => sheet !== undefined);
  }

  getVisibleSheets(): WorksheetModel[] {
    return this.getSheets().filter((sheet) => !sheet.hidden);
  }

  getTable(tableId: string): WorkbookTableModel {
    const table = this.tables.get(tableId);
    if (!table) throw new Error(`Unknown table: ${tableId}`);
    return table;
  }

  getPrintDocument(sheetId: SheetId): PrintDocumentSnapshot | undefined {
    this.getSheet(sheetId);
    const document = this.printDocuments.get(sheetId);
    return document ? structuredClone(document) : undefined;
  }

  setPrintDocument(document: PrintDocumentSnapshot): void {
    if (document.unitId !== this.unitId) throw new Error(`Print document unit mismatch: expected ${this.unitId}, received ${document.unitId}`);
    this.getSheet(document.sheetId);
    this.printDocuments.set(document.sheetId, normalizePrintDocumentSnapshot(document));
  }

  removePrintDocument(sheetId: SheetId): PrintDocumentSnapshot | undefined {
    this.getSheet(sheetId);
    const document = this.printDocuments.get(sheetId);
    this.printDocuments.delete(sheetId);
    return document ? structuredClone(document) : undefined;
  }

  clearPrintDocuments(): void {
    this.printDocuments.clear();
  }

  listPrintDocuments(): PrintDocumentSnapshot[] {
    return [...this.printDocuments.values()].map((document) => structuredClone(document));
  }

  getQueryDefinition(queryId: string): QueryDefinitionSnapshot | undefined {
    const definition = this.queryDefinitions.get(queryId);
    return definition ? structuredClone(definition) : undefined;
  }

  setQueryDefinition(definition: QueryDefinitionSnapshot): void {
    const normalized = normalizeQueryDefinitionSnapshot(definition);
    this.queryDefinitions.set(normalized.id, normalized);
  }

  removeQueryDefinition(queryId: string): QueryDefinitionSnapshot | undefined {
    const definition = this.queryDefinitions.get(queryId);
    this.queryDefinitions.delete(queryId);
    return definition ? structuredClone(definition) : undefined;
  }

  clearQueryDefinitions(): void {
    this.queryDefinitions.clear();
  }

  listQueryDefinitions(): QueryDefinitionSnapshot[] {
    return [...this.queryDefinitions.values()].map((definition) => structuredClone(definition));
  }

  getDefinedName(name: string, sheetId?: SheetId): DefinedNameModel | undefined {
    const normalized = name.trim().toLocaleLowerCase();
    if (sheetId) {
      const local = this.definedNameModels.find((entry) => entry.scope === 'sheet'
        && entry.sheetId === sheetId
        && entry.name.toLocaleLowerCase() === normalized);
      if (local) return structuredClone(local);
    }
    const global = this.definedNameModels.find((entry) => entry.scope === 'workbook'
      && entry.name.toLocaleLowerCase() === normalized);
    return global ? structuredClone(global) : undefined;
  }

  getDefinedNameExact(name: string, scope: DefinedNameScope, sheetId?: SheetId): DefinedNameModel | undefined {
    const normalized = name.trim().toLocaleLowerCase();
    const exact = this.definedNameModels.find((entry) => entry.scope === scope
      && entry.sheetId === sheetId
      && entry.name.toLocaleLowerCase() === normalized);
    return exact ? structuredClone(exact) : undefined;
  }

  listDefinedNames(sheetId?: SheetId): DefinedNameModel[] {
    return this.definedNameModels
      .filter((entry) => entry.scope === 'workbook' || entry.sheetId === sheetId)
      .map((entry) => structuredClone(entry));
  }

  setDefinedName(input: DefinedNameModel): DefinedNameModel {
    const model = normalizeDefinedNameModel(input);
    const index = this.definedNameModels.findIndex((entry) => entry.scope === model.scope
      && entry.sheetId === model.sheetId
      && entry.name.toLocaleLowerCase() === model.name.toLocaleLowerCase());
    if (index >= 0) this.definedNameModels[index] = structuredClone(model);
    else this.definedNameModels.push(structuredClone(model));
    if (model.scope === 'workbook') this.definedNames[model.name] = model.formula;
    return structuredClone(model);
  }

  removeDefinedName(name: string, scope: DefinedNameScope = 'workbook', sheetId?: SheetId): DefinedNameModel | undefined {
    const normalized = name.trim().toLocaleLowerCase();
    const index = this.definedNameModels.findIndex((entry) => entry.scope === scope
      && entry.sheetId === sheetId
      && entry.name.toLocaleLowerCase() === normalized);
    const previous = index >= 0 ? this.definedNameModels[index] : undefined;
    if (index >= 0) this.definedNameModels.splice(index, 1);
    if (scope === 'workbook' && previous) delete this.definedNames[previous.name];
    return previous ? structuredClone(previous) : undefined;
  }

  addTable(table: WorkbookTableModel): void {
    if (this.tables.has(table.id)) throw new Error(`Table already exists: ${table.id}`);
    this.tables.set(table.id, structuredClone(table));
  }

  removeTable(tableId: string): WorkbookTableModel {
    const table = this.getTable(tableId);
    this.tables.delete(tableId);
    return table;
  }

  addSheet(id: SheetId, name: string, rowCount = 1000, columnCount = 26): WorksheetModel {
    if (this.sheets.has(id)) throw new Error(`Sheet already exists: ${id}`);
    const sheet = new WorksheetModel(id, name, rowCount, columnCount);
    this.sheets.set(id, sheet);
    this.sheetOrder.push(id);
    return sheet;
  }

  duplicateSheet(sourceSheetId: SheetId, newId: SheetId, newName: string): WorksheetModel {
    const source = this.getSheet(sourceSheetId);
    const copy = source.cloneWithIdentity(newId, newName);
    this.sheets.set(newId, copy);
    const scopedNames = this.definedNameModels
      .filter((entry) => entry.scope === 'sheet' && entry.sheetId === sourceSheetId)
      .map((entry) => ({ ...entry, sheetId: newId }));
    this.definedNameModels.push(...structuredClone(scopedNames));
    const printDocument = this.printDocuments.get(sourceSheetId);
    if (printDocument) {
      this.printDocuments.set(newId, structuredClone({
        ...printDocument,
        sheetId: newId,
        printAreas: printDocument.printAreas.map((area) => ({ sheetId: newId, range: { ...area.range, sheetId: newId } })),
        pageBreaks: printDocument.pageBreaks.map((pageBreak) => pageBreak.row !== undefined
          ? { sheetId: newId, row: pageBreak.row }
          : { sheetId: newId, column: pageBreak.column }),
      }));
    }
    const sourceIndex = this.sheetOrder.indexOf(sourceSheetId);
    this.sheetOrder.splice(sourceIndex + 1, 0, newId);
    return copy;
  }

  reorderSheet(sheetId: SheetId, toIndex: number): void {
    const fromIndex = this.sheetOrder.indexOf(sheetId);
    if (fromIndex < 0) throw new Error(`Unknown sheet: ${sheetId}`);
    const clamped = Math.max(0, Math.min(toIndex, this.sheetOrder.length - 1));
    this.sheetOrder.splice(fromIndex, 1);
    this.sheetOrder.splice(clamped, 0, sheetId);
  }

  removeSheet(sheetId: SheetId): WorksheetModel {
    if (this.sheets.size <= 1) throw new Error('A workbook must keep at least one worksheet');
    const sheet = this.getSheet(sheetId);
    this.sheets.delete(sheetId);
    this.printDocuments.delete(sheetId);
    for (let index = this.definedNameModels.length - 1; index >= 0; index -= 1) {
      if (this.definedNameModels[index]?.scope === 'sheet' && this.definedNameModels[index]?.sheetId === sheetId) {
        this.definedNameModels.splice(index, 1);
      }
    }
    this.sheetOrder = this.sheetOrder.filter((id) => id !== sheetId);
    if (this.activeSheetId === sheetId) {
      const firstSheet = this.getSheets()[0];
      if (firstSheet) this.activeSheetId = firstSheet.id;
    }
    return sheet;
  }

  snapshot(): WorkbookSnapshot {
    return {
      schema: 'WorkbookSnapshot',
      unitId: this.unitId,
      name: this.name,
      activeSheetId: this.activeSheetId,
      definedNames: { ...this.definedNames },
      definedNameModels: this.definedNameModels.length > 0
        ? structuredClone(this.definedNameModels)
        : Object.entries(this.definedNames).map(([name, formula]) => ({ name, formula, scope: 'workbook' as const })),
      tables: [...this.tables.values()].map((table) => structuredClone(table)),
      printDocuments: this.listPrintDocuments(),
      queryDefinitions: this.listQueryDefinitions(),
      sheets: this.getSheets().map((sheet) => ({
        id: sheet.id,
        name: sheet.name,
        rowCount: sheet.rowCount,
        columnCount: sheet.columnCount,
        cells: sheet.cells.toJSON(),
        merges: structuredClone(sheet.merges),
        freeze: { ...sheet.freeze },
        pivots: structuredClone(sheet.pivots),
        sparklines: structuredClone(sheet.sparklines),
        conditionalFormats: structuredClone(sheet.conditionalFormats),
        dataValidations: structuredClone(sheet.dataValidations),
        rowHeights: { ...sheet.rowHeights },
        columnWidths: { ...sheet.columnWidths },
        hiddenRows: [...sheet.hiddenRows],
        hiddenColumns: [...sheet.hiddenColumns],
        tabColor: sheet.tabColor,
        bandedRule: sheet.bandedRule ? structuredClone(sheet.bandedRule) : undefined,
        filter: sheet.filter ? structuredClone(sheet.filter) : undefined,
        sheetTables: structuredClone(sheet.sheetTables),
        sparklineGroups: structuredClone(sheet.sparklineGroups),
        drawings: structuredClone(sheet.drawings),
        drawingPayloads: Object.fromEntries([...sheet.drawingPayloads.entries()].map(([k, v]) => [k, structuredClone(v)])),
        notes: [...sheet.notes.entries()].map(([key, note]) => {
          const [row, column] = key.split(':').map(Number);
          return { row: row!, column: column!, note: structuredClone(note) };
        }),
        commentThreads: structuredClone(sheet.commentThreads),
        spillRanges: structuredClone(sheet.spillRanges),
        protectionRules: structuredClone(sheet.protectionRules),
        showGridlines: sheet.showGridlines,
        showHeaders: sheet.showHeaders,
        zoom: sheet.zoom,
        hidden: sheet.hidden,
        outline: sheet.outline ? structuredClone(sheet.outline) : undefined,
      })),
    };
  }

  static fromSnapshot(snapshot: WorkbookSnapshot): WorkbookModel {
    if (snapshot.schema !== 'WorkbookSnapshot') throw new Error('Unsupported workbook snapshot schema');
    if (snapshot.sheets.length === 0 || !snapshot.sheets.some((sheet) => sheet.id === snapshot.activeSheetId)) {
      throw new Error(`Workbook snapshot active sheet is not present: ${snapshot.activeSheetId}`);
    }
    const workbook = new WorkbookModel(snapshot.unitId, snapshot.name);
    workbook.sheets.clear();
    workbook.definedNames = snapshot.definedNames ? { ...snapshot.definedNames } : {};
    workbook.definedNameModels.push(...structuredClone(snapshot.definedNameModels ?? Object.entries(workbook.definedNames).map(([name, formula]) => ({ name, formula, scope: 'workbook' as const }))));
    for (const entry of workbook.definedNameModels) {
      if (entry.scope === 'workbook' && workbook.definedNames[entry.name] === undefined) workbook.definedNames[entry.name] = entry.formula;
    }
    for (const table of snapshot.tables ?? []) workbook.tables.set(table.id, structuredClone(table));
    for (const input of snapshot.sheets) {
      const sheet = new WorksheetModel(input.id, input.name, input.rowCount, input.columnCount);
      const matrix = CellMatrix.fromJSON(input.cells);
      matrix.forEach((cell, row, column) => sheet.cells.set(row, column, cell));
      sheet.merges.push(...structuredClone(input.merges));
      sheet.freeze = { ...input.freeze };
      sheet.pivots.push(...structuredClone(input.pivots));
      sheet.sparklines.push(...structuredClone(input.sparklines));
      if (input.sparklineGroups) sheet.sparklineGroups.push(...structuredClone(input.sparklineGroups));
      sheet.drawings.push(...structuredClone(input.drawings));
      for (const [key, payload] of Object.entries(input.drawingPayloads)) {
        sheet.drawingPayloads.set(key, structuredClone(payload));
      }
      if (input.notes) {
        for (const entry of input.notes) sheet.notes.set(noteCellKey(entry.row, entry.column), structuredClone(entry.note));
      }
      if (input.commentThreads) sheet.commentThreads.push(...structuredClone(input.commentThreads));
      if (input.conditionalFormats) sheet.conditionalFormats.push(...structuredClone(input.conditionalFormats));
      if (input.dataValidations) sheet.dataValidations.push(...structuredClone(input.dataValidations));
      if (input.rowHeights) Object.assign(sheet.rowHeights, input.rowHeights);
      if (input.columnWidths) Object.assign(sheet.columnWidths, input.columnWidths);
      if (input.hiddenRows) input.hiddenRows.forEach((r) => sheet.hiddenRows.add(r));
      if (input.hiddenColumns) input.hiddenColumns.forEach((c) => sheet.hiddenColumns.add(c));
      if (input.bandedRule) sheet.bandedRule = structuredClone(input.bandedRule);
      if (input.filter) sheet.filter = structuredClone(input.filter);
      if (input.sheetTables) sheet.sheetTables.push(...structuredClone(input.sheetTables));
      if (input.spillRanges) sheet.spillRanges.push(...structuredClone(input.spillRanges));
      if (input.protectionRules) sheet.protectionRules.push(...structuredClone(input.protectionRules));
      if (input.showGridlines != null) sheet.showGridlines = input.showGridlines;
      if (input.showHeaders != null) sheet.showHeaders = input.showHeaders;
      if (input.zoom != null) sheet.zoom = input.zoom;
      if (input.hidden != null) sheet.hidden = input.hidden;
      if (input.outline) sheet.outline = structuredClone(input.outline);
      sheet.tabColor = input.tabColor;
      workbook.sheets.set(sheet.id, sheet);
    }
    for (const document of snapshot.printDocuments ?? []) workbook.setPrintDocument(document);
    for (const definition of snapshot.queryDefinitions ?? []) workbook.setQueryDefinition(definition);
    workbook.activeSheetId = snapshot.activeSheetId;
    workbook.sheetOrder = snapshot.sheets.map((sheet) => sheet.id);
    return workbook;
  }
}
