import type {
  CellComment,
  CellEditorConfig,
  CellNote,
  CellData,
  FilterCellValue,
  CellStyle,
  CellPresentation,
  ConditionalFormatRule,
  DataValidationRule,
  DrawingObject,
  DrawingPayload,
  WorksheetPane,
  AutoFilterModel,
  FilterCriterion,
  DateGroupItem,
  GanttSheetDefinition,
  ReportSheetDefinition,
  MergeSpan,
  OutlineGroup,
  PivotModel,
  PivotGridProjection,
  PivotResultTree,
  RangeRef,
  TableSheetDefinition,
  SheetTableModel,
  SparklineGroup,
  SparklineModel,
  WorkbookModel,
  WorksheetModel,
} from '@react-sheets/core-model';
import {
  computeBandedCellStyle,
  createConditionalFormatRuntime,
  computeFilterHiddenRows,
  createEffectiveFilterVisualResolver,
  getAutoFilterValueDomain,
  getAutoFilterDateDomain,
  getAutoFilterDomainDescriptor,
  getAutoFilterColorDomain,
  getAutoFilterIconDomain,
  computeOutlineHiddenColumns,
  computeOutlineHiddenRows,
  computeSheetTableCellStyle,
  findSheetTableAt,
  mergePresentationStyles,
  resolveActiveFilterColumns,
  resolveFilterButtonCells,
  resolveFilterButtonStates,
  resolveFilterRangeColumns,
  resolveActiveAutoFilter,
  resolveFilterOwner,
  resolveOutlineControls,
  resolveEffectiveFilterVisual,
  validateDataInput,
  type FilterDateSystem,
  type FilterDateContext,
  type FilterDomainDescriptor,
  type ConditionalOverlay,
  type FilterButtonCell,
  type FilterButtonState,
  type OutlineControl,
} from '@react-sheets/sheet-features';
import { resolveFilterCellValue } from '@react-sheets/core-model';
import { FormulaEngine, isFormulaError, isSpillChild, type FormulaValue } from '@react-sheets/formula-engine';
import { formatValue as formatNumberValue } from '@react-sheets/number-format';
import {
  buildPivotGridProjection,
  getLastValidPivotResult,
  pivotResultMatchesLayoutAndFilter,
  pivotResultMatchesRevision,
  type PivotProjectionSourceState,
} from './features/pivot/engine';
import { cellAddress, columnLabel } from './address';
import type { DataSourceContentQuery } from './features/data-source';
import { createWorkbookCellResolver } from './features/data-source';
import {
  findCommentThreadAt,
  getCellHyperlink,
  resolveHyperlinkDisplay,
  threadToCellComment,
} from './features/review';

/** Canvas-friendly cell snapshot — no row-array projection */
export interface CanvasCellSnapshot {
  address: string;
  displayValue?: string;
  formula?: string;
  style?: CellStyle;
  richText?: import('@react-sheets/core-model').RichTextRun[];
  editor?: CellEditorConfig;
  presentation?: CellPresentation;
  phonetic?: import('@react-sheets/core-model').CellPhoneticMetadata;
  value: string;
  hasComment?: boolean;
  commentText?: string;
  comment?: CellComment;
  note?: CellNote;
  invalid?: boolean;
  hyperlink?: string;
  overlay?: ConditionalOverlay;
}

/** Canvas-friendly sheet snapshot — single getCell path, no SheetView DTO */
export interface CanvasSheetSnapshot {
  kind?: WorksheetModel['kind'];
  id: string;
  name: string;
  columns: string[];
  columnCount: number;
  rowCount: number;
  isEmpty?: boolean;
  occupiedCellCount: number;
  getCell: (row: number, column: number) => CanvasCellSnapshot | undefined;
  /** Sparse model addresses for operations such as AutoFit; never a rectangle scan. */
  forEachOccupiedCell: (
    visitor: (row: number, column: number) => void,
    selection?: { rows?: ReadonlySet<number>; columns?: ReadonlySet<number> },
  ) => void;
  usedRange: RangeRef;
  /** Canonical floating-object aggregate. UI never consumes legacy projections. */
  drawings: DrawingObject[];
  /** Read-only payload projection keyed by DrawingObject.payloadId. */
  drawingPayloads: ReadonlyMap<string, DrawingPayload>;
  drawingGroups?: import('@react-sheets/core-model').DrawingGroup[];
  snapSettings?: import('@react-sheets/core-model').WorksheetSnapSettings;
  pivots: PivotModel[];
  pivotResults: Record<string, PivotResultTree>;
  pivotTaskErrors: Readonly<Record<string, import('./features/pivot/task-protocol').PivotTaskError>>;
  /** Derived worksheet overlay; never materialized in ordinary cells. */
  pivotProjections: Record<string, PivotGridProjection>;
  sparklines: SparklineModel[];
  sparklineGroups?: SparklineGroup[];
  conditionalFormats: ConditionalFormatRule[];
  dataValidations: DataValidationRule[];
  dataRegions: import('@react-sheets/core-model').SheetDataRegion[];
  merges: MergeSpan[];
  pane: WorksheetPane;
  autoFilter?: AutoFilterModel;
  getFilterOwner: (column: number) => { kind: 'worksheet' | 'table'; tableId?: string } | undefined;
  getActiveAutoFilter: (column: number) => AutoFilterModel | undefined;
  defaultRowHeightPx: number;
  defaultColumnWidthPx: number;
  maximumDigitWidthPx: number;
  rowHeightsPx: Record<number, number>;
  columnWidthsPx: Record<number, number>;
  hiddenRows: number[];
  hiddenColumns: number[];
  outlineGroups: OutlineGroup[];
  outlineControls: OutlineControl[];
  filterRangeColumns: number[];
  activeFilterColumns: number[];
  filterButtons: FilterButtonCell[];
  filterButtonStates: FilterButtonState[];
  getFilterValueDomain: (column: number) => string[];
  getFilterDomainDescriptor: (column: number) => FilterDomainDescriptor;
  /** Typed source dates for hierarchical date-group editing; never derived from display text. */
  getFilterDateDomain?: (column: number) => Array<{ value: import('@react-sheets/core-model').FilterScalar; group: DateGroupItem & { hour: number; minute: number; second: number } }>;
  getFilterCriterion: (column: number) => FilterCriterion | undefined;
  getFilterColorDomain: (column: number) => Array<{ target: 'cell' | 'font'; color: string }>;
  getFilterIconDomain: (column: number) => Array<{ iconSet: string; iconId: number }>;
  sheetTables: SheetTableModel[];
  tableSheet?: TableSheetDefinition;
  ganttSheet?: GanttSheetDefinition;
  reportSheet?: ReportSheetDefinition;
  tabColor?: string;
  hidden?: boolean;
}

function toFormulaDisplay(value: FormulaValue): string {
  if (isFormulaError(value)) return value.code;
  if (Array.isArray(value)) {
    return value.length > 0 && Array.isArray(value[0]) ? String(value[0][0]) : String(value[0]);
  }
  return value == null ? '' : String(value);
}

function formatDisplayValue(
  cell: CellData | undefined,
  formula: FormulaEngine,
  sheet: WorksheetModel,
  sheetId: string,
  row: number,
  column: number,
): string {
  if (cell?.formula) {
    return toFormulaDisplay(formula.getCellValue({ sheetId, row, column }));
  }
  const spillValue = formula.getSpillValueAt(sheetId, row, column);
  if (spillValue !== undefined) return toFormulaDisplay(spillValue);
  if (!cell) return '';
  const resolved = resolveFilterCellValue(cell);
  if (resolved.value == null) return '';
  if (typeof resolved.value === 'number') {
    return formatNumberValue(resolved.value, cell.numberFormat ?? cell.style?.numberFormat);
  }
  return resolved.text;
}

function pivotSourceState(
  pivot: PivotModel,
  dataContent: ReadonlyMap<string, DataSourceContentQuery>,
): PivotProjectionSourceState | undefined {
  if (pivot.source.kind !== 'data-source') return undefined;
  const query = dataContent.get(pivot.source.dataSourceId);
  if (!query) return { availability: 'missing', error: `Data source ${pivot.source.dataSourceId} is unavailable` };
  const states = query.getLoadStates();
  if (states.some((state) => state.availability === 'error')) {
    const current = states.find((state) => state.availability === 'error');
    return { availability: 'error', error: current?.error ?? `Data source ${pivot.source.dataSourceId} failed to load` };
  }
  if (states.some((state) => state.availability === 'missing')) {
    const current = states.find((state) => state.availability === 'missing');
    return { availability: 'missing', error: current?.error ?? `Data source ${pivot.source.dataSourceId} is missing a block` };
  }
  if (states.some((state) => state.availability === 'loading') || states.length === 0) return { availability: 'loading' };
  return { availability: 'ready' };
}

export function buildCanvasSheetSnapshot(
  workbook: WorkbookModel,
  sheet: WorksheetModel,
  formula: FormulaEngine,
  showInvalid: boolean,
  cachedPivotResults: Readonly<Record<string, PivotResultTree>> = {},
  dataContent: ReadonlyMap<string, DataSourceContentQuery> = new Map(),
  dateSystem: FilterDateSystem = '1900',
  pivotErrors: Readonly<Record<string, import('./features/pivot/task-protocol').PivotTaskError>> = {},
  dateContext?: FilterDateContext,
): CanvasSheetSnapshot {
  const conditionalRuntime = createConditionalFormatRuntime(sheet);
  const cellResolver = createWorkbookCellResolver(dataContent);
  const resolveFilterCell = (owner: WorksheetModel, row: number, column: number): FilterCellValue => {
    const cell = cellResolver.resolve(owner, row, column)?.cell;
    const spillValue = formula.getSpillValueAt(owner.id, row, column);
    if (spillValue !== undefined) return resolveFilterCellValue(cell, spillValue, dateSystem);
    if (cell?.formula !== undefined) {
      const result = formula.getCellResult({ sheetId: owner.id, row, column });
      // A missing calculation result is not permission to read authored
      // formula text/value.  It is an unresolved filter value until the
      // FormulaEngine publishes the next result.
      const evaluated = result ? result.value : cell.formulaValue !== undefined ? cell.formulaValue : null;
      return resolveFilterCellValue(cell, evaluated, dateSystem);
    }
    return resolveFilterCellValue(cell, undefined, dateSystem);
  };
  const readFilterCell = (row: number, column: number) => resolveFilterCell(sheet, row, column);
  const filterVisual = createEffectiveFilterVisualResolver((row, column) => conditionalRuntime.resolveCell(row, column));
  const filterHidden = computeFilterHiddenRows(sheet, readFilterCell, dateSystem, filterVisual, dateContext);
  const outlineHiddenRows = computeOutlineHiddenRows(sheet);
  const outlineHiddenColumns = computeOutlineHiddenColumns(sheet);
  const hiddenRows = new Set<number>([...sheet.hiddenRows, ...filterHidden, ...outlineHiddenRows]);
  const hiddenColumns = new Set<number>([...sheet.hiddenColumns, ...outlineHiddenColumns]);
  const filterRangeColumns = resolveFilterRangeColumns(sheet);
  const activeFilterColumns = resolveActiveFilterColumns(sheet);
  const filterButtons = resolveFilterButtonCells(sheet);
  const outlineControls = resolveOutlineControls(sheet);
  const viewColumns = Array.from({ length: Math.max(26, sheet.columnCount) }, (_, index) => columnLabel(index));
  const usedRange = sheet.usedRange;
  const advancedTableId = sheet.kind === 'table-sheet' ? sheet.tableSheet?.viewId : sheet.kind === 'gantt-sheet' ? sheet.ganttSheet?.viewId : sheet.kind === 'report-sheet' ? sheet.reportSheet?.tableId : undefined;
  const advancedTable = advancedTableId ? workbook.dataModel.tables.get(advancedTableId) : undefined;

  const resolveModelCell = (row: number, column: number): { cell?: CellData; owner: WorksheetModel; row: number; column: number } => {
    const local = cellResolver.resolve(sheet, row, column)?.cell;
    if (local || row === 0 || !advancedTable?.sourceRange || sheet.kind === 'report-sheet') return { cell: local, owner: sheet, row, column };
    const field = advancedTable.fields[column];
    const sourceSheet = workbook.sheets.get(advancedTable.sourceRange.sheetId);
    const sourceRow = advancedTable.sourceRange.startRow + row;
    if (!field || !sourceSheet || sourceRow > advancedTable.sourceRange.endRow) return { owner: sheet, row, column };
    const sourceColumn = advancedTable.sourceRange.startColumn + field.ordinal;
    return { cell: cellResolver.resolve(sourceSheet, sourceRow, sourceColumn)?.cell, owner: sourceSheet, row: sourceRow, column: sourceColumn };
  };

  const getCell = (row: number, column: number): CanvasCellSnapshot | undefined => {
    if (row < 0 || row >= sheet.rowCount || column < 0 || column >= sheet.columnCount) return undefined;
    const resolved = resolveModelCell(row, column);
    const modelCell = resolved.cell;
    const value = formatDisplayValue(modelCell, formula, resolved.owner, resolved.owner.id, resolved.row, resolved.column);
    const resolvedFilter = resolveFilterCell(resolved.owner, resolved.row, resolved.column);
    const overlay = conditionalRuntime.resolveCell(row, column);
    const table = findSheetTableAt(sheet, row, column);
    const presentation = mergePresentationStyles(
      computeBandedCellStyle(sheet, row, column),
      table ? computeSheetTableCellStyle(table, row, column) : undefined,
    );
    const effectiveStyle = resolveEffectiveFilterVisual(modelCell, overlay, presentation).style;
    const style = Object.keys(effectiveStyle).length > 0 ? effectiveStyle : undefined;
    const validation = validateDataInput(sheet, row, column, resolvedFilter.value);
    const thread = findCommentThreadAt(sheet, row, column);
    const note = sheet.review.getNoteAt(row, column);
    const comment = thread ? threadToCellComment(thread) : undefined;
    const hyperlinkDetail = getCellHyperlink(sheet, row, column) ?? modelCell?.hyperlinkDetail;
    const hyperlink = resolveHyperlinkDisplay(hyperlinkDetail);
    return {
      address: cellAddress(row, column),
      formula: modelCell?.formula,
      style,
      richText: modelCell?.richText ? structuredClone(modelCell.richText) : undefined,
      editor: modelCell?.editor ? structuredClone(modelCell.editor) : undefined,
      presentation: modelCell?.presentation ? structuredClone(modelCell.presentation) : undefined,
      phonetic: modelCell?.phonetic ? structuredClone(modelCell.phonetic) : undefined,
      value,
      displayValue: value,
      hasComment: Boolean(comment || note),
      commentText: comment?.text ?? note?.text,
      comment,
      note: note ? structuredClone(note) : undefined,
      invalid: showInvalid && resolvedFilter.value != null && !validation.valid,
      hyperlink,
      overlay,
    };
  };

  const forEachOccupiedCell: CanvasSheetSnapshot['forEachOccupiedCell'] = (visitor, selection = {}) => {
    const visitMaterialized = (_cell: CellData, row: number, column: number) => {
      if (selection.rows && !selection.rows.has(row)) return;
      if (selection.columns && !selection.columns.has(column)) return;
      visitor(row, column);
    };
    if (selection.rows) sheet.cells.forEachInRows(selection.rows, visitMaterialized);
    else if (selection.columns) sheet.cells.forEachInColumns(selection.columns, visitMaterialized);
    else sheet.cells.forEach(visitMaterialized);

    for (const region of sheet.dataRegions) {
      const startRow = Math.max(region.range.startRow, region.headerRow + 1);
      const rows = selection.rows
        ? [...selection.rows].filter((row) => row >= startRow && row <= region.range.endRow)
        : Array.from({ length: Math.max(0, region.range.endRow - startRow + 1) }, (_, offset) => startRow + offset);
      const columns = selection.columns
        ? [...selection.columns].filter((column) => column >= region.range.startColumn && column <= region.range.endColumn)
        : Array.from({ length: region.range.endColumn - region.range.startColumn + 1 }, (_, offset) => region.range.startColumn + offset);
      for (const row of rows) for (const column of columns) {
        if (!sheet.cells.has(row, column)) visitor(row, column);
      }
    }
  };

  const pivotResults: Record<string, PivotResultTree> = {};
  const pivotProjections: Record<string, PivotGridProjection> = {};
  for (const pivot of sheet.pivots) {
    const sourceState = pivotSourceState(pivot, dataContent);
    const runtimeResult = cachedPivotResults[pivot.id];
    // A snapshot is a projection boundary, never a refresh authority.  A
    // source-stale result is still renderable for every policy as long as its
    // layout/filter contract matches; the coordinator decides if/when it is
    // replaced.
    const reusable = (result: PivotResultTree | undefined) => pivotResultMatchesRevision(workbook, pivot, result, formula)
      || pivotResultMatchesLayoutAndFilter(workbook, pivot, result, formula);
    const retainedResult = reusable(runtimeResult) ? undefined : getLastValidPivotResult(workbook, pivot.id);
    let cachedResult = reusable(runtimeResult) ? runtimeResult : reusable(retainedResult) ? retainedResult : undefined;
    if (cachedResult) pivotResults[pivot.id] = cachedResult;
    try {
      pivotProjections[pivot.id] = buildPivotGridProjection(workbook, pivot, cachedResult, { sourceState, formula, refreshError: pivotErrors[pivot.id]?.message });
      const retained = getLastValidPivotResult(workbook, pivot.id);
      if (!pivotResults[pivot.id] && reusable(retained)) pivotResults[pivot.id] = retained;
    } catch {
      // Invalid target/source is surfaced by command validation. The snapshot
      // remains renderable for the rest of the worksheet.
    }
  }

  return {
    id: sheet.id,
    kind: sheet.kind,
    name: sheet.name,
    columns: viewColumns,
    columnCount: sheet.columnCount,
    rowCount: sheet.rowCount,
    isEmpty: sheet.cells.count() === 0 && sheet.dataRegions.length === 0,
    occupiedCellCount: sheet.cells.count() + sheet.dataRegions.reduce((count, region) => count + (region.range.endRow - region.range.startRow + 1) * (region.range.endColumn - region.range.startColumn + 1), 0),
    getCell,
    forEachOccupiedCell,
    usedRange,
    drawings: structuredClone(sheet.drawings),
    drawingPayloads: new Map(
      [...sheet.drawingPayloads.entries()].map(([payloadId, payload]) => [payloadId, structuredClone(payload)]),
    ),
    drawingGroups: structuredClone(sheet.drawingGroups),
    snapSettings: structuredClone(sheet.snapSettings),
    pivots: [...sheet.pivots],
    pivotResults,
    pivotTaskErrors: pivotErrors,
    pivotProjections,
    sparklines: [...sheet.sparklines],
    sparklineGroups: structuredClone(sheet.sparklineGroups),
    conditionalFormats: [...sheet.conditionalFormats],
    dataValidations: [...sheet.dataValidations],
    dataRegions: sheet.dataRegions.map((region) => structuredClone(region)),
    merges: [...sheet.merges],
    pane: { ...sheet.pane },
    autoFilter: resolveActiveAutoFilter(sheet) ? structuredClone(resolveActiveAutoFilter(sheet)) : undefined,
    defaultRowHeightPx: sheet.defaultRowHeightPx,
    defaultColumnWidthPx: sheet.defaultColumnWidthPx,
    maximumDigitWidthPx: workbook.dimensionMetrics.maximumDigitWidthPx,
    rowHeightsPx: { ...sheet.rowHeightsPx },
    columnWidthsPx: { ...sheet.columnWidthsPx },
    hiddenRows: [...hiddenRows].sort((a, b) => a - b),
    hiddenColumns: [...hiddenColumns].sort((a, b) => a - b),
    outlineGroups: sheet.outline ? structuredClone(sheet.outline.groups) : [],
    outlineControls,
    filterRangeColumns,
    activeFilterColumns,
    filterButtons,
    filterButtonStates: resolveFilterButtonStates(sheet),
    getFilterValueDomain: (column) => getAutoFilterValueDomain(sheet, column, readFilterCell, dateSystem, filterVisual, dateContext),
    getFilterDomainDescriptor: (column) => getAutoFilterDomainDescriptor(sheet, column, readFilterCell, dateSystem, filterVisual, dateContext),
    getFilterDateDomain: (column) => getAutoFilterDateDomain(sheet, column, readFilterCell, dateSystem, filterVisual, dateContext),
    getFilterOwner: (column) => resolveFilterOwner(sheet, column),
    getActiveAutoFilter: (column) => {
      const filter = resolveActiveAutoFilter(sheet, column);
      return filter ? structuredClone(filter) : undefined;
    },
    getFilterCriterion: (column) => resolveActiveAutoFilter(sheet, column)?.columns[column]?.criterion,
    getFilterColorDomain: (column) => getAutoFilterColorDomain(sheet, column, readFilterCell, dateSystem, filterVisual)
      .map(({ target, color }) => ({ target, color })),
    getFilterIconDomain: (column) => getAutoFilterIconDomain(sheet, column, readFilterCell, dateSystem, filterVisual),
    sheetTables: [...sheet.sheetTables],
    tableSheet: sheet.tableSheet ? structuredClone(sheet.tableSheet) : undefined,
    ganttSheet: sheet.ganttSheet ? structuredClone(sheet.ganttSheet) : undefined,
    reportSheet: sheet.reportSheet ? structuredClone(sheet.reportSheet) : undefined,
    tabColor: sheet.tabColor,
    hidden: sheet.hidden,
  };
}

export function buildAllSheetSnapshots(
  workbook: WorkbookModel,
  formula: FormulaEngine,
  pivotResults: Readonly<Record<string, PivotResultTree>>,
  dataContent: ReadonlyMap<string, DataSourceContentQuery> = new Map(),
  pivotErrors: Readonly<Record<string, import('./features/pivot/task-protocol').PivotTaskError>> = {},
): CanvasSheetSnapshot[] {
  return workbook.getSheets().map((sheet) => buildCanvasSheetSnapshot(workbook, sheet, formula, true, pivotResults, dataContent, '1900', pivotErrors));
}
