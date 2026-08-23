import type { CellData, RangeRef, Row, Column, WorksheetModel } from './index';
import type { CellNote, DrawingObject, StructuralTransformParams, CommentThread, SheetTableModel, SpillRange, ProtectionRule } from './domain';
import { WorkbookModel, noteCellKey } from './index';
import { cellAddress, columnLabel, parseColumnLabel } from './address';

export interface StructuralTransformResult {
  removedCells: Array<{ row: Row; column: Column; cell: CellData }>;
}

/** 结构变换唯一入口 — 一次更新 cells/merges/CF/validation/filter/freeze/charts/pivots/tables/names/drawings/notes/comments/protection/公式引用 */
export class StructuralTransform {
  static apply(workbook: WorkbookModel, params: StructuralTransformParams): StructuralTransformResult {
    const sheet = workbook.getSheet(params.sheetId);
    switch (params.kind) {
      case 'insert-rows':
        return applyAxis(workbook, sheet, 'row', params.at ?? 0, params.count ?? 1, 1);
      case 'delete-rows':
        return applyAxis(workbook, sheet, 'row', params.at ?? 0, params.count ?? 1, -1);
      case 'insert-columns':
        return applyAxis(workbook, sheet, 'column', params.at ?? 0, params.count ?? 1, 1);
      case 'delete-columns':
        return applyAxis(workbook, sheet, 'column', params.at ?? 0, params.count ?? 1, -1);
      case 'move-range':
        return applyMoveRange(workbook, sheet, params.sourceRange!, params.targetOrigin!);
      case 'shift-cells-down':
      case 'shift-cells-up':
      case 'shift-cells-right':
      case 'shift-cells-left':
        return applyShiftCells(sheet, params.sourceRange!, params.kind);
      default:
        throw new Error(`Unknown structural op: ${(params as StructuralTransformParams).kind}`);
    }
  }
}

function applyAxis(
  workbook: WorkbookModel,
  sheet: WorksheetModel,
  axis: 'row' | 'column',
  at: number,
  count: number,
  direction: 1 | -1,
): StructuralTransformResult {
  if (count <= 0) return { removedCells: [] };
  const end = at + count - 1;
  let removed: Array<{ row: Row; column: Column; cell: CellData }> = [];

  if (direction === 1) {
    if (axis === 'row') {
      sheet.cells.shiftRows(at, count, 1);
      sheet.rowCount += count;
    } else {
      sheet.cells.shiftColumns(at, count, 1);
      sheet.columnCount += count;
    }
  } else if (axis === 'row') {
    removed = sheet.cells.extractRegion(at, end, 0, Math.max(sheet.columnCount - 1, 0));
    sheet.cells.shiftRows(end + 1, count, -1);
    sheet.rowCount = Math.max(1, sheet.rowCount - count);
  } else {
    removed = sheet.cells.extractRegion(0, Math.max(sheet.rowCount - 1, 0), at, end);
    sheet.cells.shiftColumns(end + 1, count, -1);
    sheet.columnCount = Math.max(1, sheet.columnCount - count);
  }

  shiftMerges(sheet, axis, at, count, direction);
  shiftRuleRanges(sheet.conditionalFormats, axis, at, count, direction);
  shiftRuleRanges(sheet.dataValidations, axis, at, count, direction);
  shiftFilter(sheet, axis, at, count, direction);
  shiftFreeze(sheet, axis, at, count, direction);
  shiftHiddenAndSizes(sheet, axis, at, count, direction);
  shiftSparklines(sheet, axis, at, count, direction);
  shiftPivots(sheet, axis, at, count, direction);
  shiftCharts(sheet, axis, at, count, direction);
  shiftDrawings(sheet, axis, at, count, direction);
  shiftSheetTables(sheet, axis, at, count, direction);
  shiftNotes(sheet, axis, at, count, direction);
  shiftComments(sheet, axis, at, count, direction);
  shiftSpills(sheet, axis, at, count, direction);
  shiftProtection(sheet, axis, at, count, direction);
  shiftBanded(sheet, axis, at, count, direction);
  rewriteFormulas(workbook, sheet.id, axis, at, count, direction);
  rewriteDefinedNames(workbook, sheet, axis, at, count, direction);
  return { removedCells: removed };
}

function applyShiftCells(
  sheet: WorksheetModel,
  range: RangeRef,
  kind: 'shift-cells-down' | 'shift-cells-up' | 'shift-cells-right' | 'shift-cells-left',
): StructuralTransformResult {
  const removedCells: Array<{ row: Row; column: Column; cell: CellData }> = [];
  const startRow = Math.min(range.startRow, range.endRow);
  const endRow = Math.max(range.startRow, range.endRow);
  const startColumn = Math.min(range.startColumn, range.endColumn);
  const endColumn = Math.max(range.startColumn, range.endColumn);

  if (kind === 'shift-cells-down') {
    for (let column = startColumn; column <= endColumn; column++) {
      for (let row = endRow; row > startRow; row--) {
        const previous = sheet.cells.get(row - 1, column);
        if (previous) sheet.cells.set(row, column, structuredClone(previous));
        else sheet.cells.delete(row, column);
      }
      const cleared = sheet.cells.get(startRow, column);
      if (cleared) removedCells.push({ row: startRow, column, cell: structuredClone(cleared) });
      sheet.cells.delete(startRow, column);
    }
  } else if (kind === 'shift-cells-up') {
    for (let column = startColumn; column <= endColumn; column++) {
      for (let row = startRow; row < endRow; row++) {
        const next = sheet.cells.get(row + 1, column);
        if (next) sheet.cells.set(row, column, structuredClone(next));
        else sheet.cells.delete(row, column);
      }
      const cleared = sheet.cells.get(endRow, column);
      if (cleared) removedCells.push({ row: endRow, column, cell: structuredClone(cleared) });
      sheet.cells.delete(endRow, column);
    }
  } else if (kind === 'shift-cells-right') {
    for (let row = startRow; row <= endRow; row++) {
      for (let column = endColumn; column > startColumn; column--) {
        const previous = sheet.cells.get(row, column - 1);
        if (previous) sheet.cells.set(row, column, structuredClone(previous));
        else sheet.cells.delete(row, column);
      }
      const cleared = sheet.cells.get(row, startColumn);
      if (cleared) removedCells.push({ row, column: startColumn, cell: structuredClone(cleared) });
      sheet.cells.delete(row, startColumn);
    }
  } else {
    for (let row = startRow; row <= endRow; row++) {
      for (let column = startColumn; column < endColumn; column++) {
        const next = sheet.cells.get(row, column + 1);
        if (next) sheet.cells.set(row, column, structuredClone(next));
        else sheet.cells.delete(row, column);
      }
      const cleared = sheet.cells.get(row, endColumn);
      if (cleared) removedCells.push({ row, column: endColumn, cell: structuredClone(cleared) });
      sheet.cells.delete(row, endColumn);
    }
  }

  return { removedCells };
}

function shiftRangeRef(range: RangeRef, axis: 'row' | 'column', at: number, count: number, direction: 1 | -1): boolean {
  const startKey = axis === 'row' ? 'startRow' : 'startColumn';
  const endKey = axis === 'row' ? 'endRow' : 'endColumn';
  if (direction === 1) {
    if (range[startKey] >= at) {
      range[startKey] += count;
      range[endKey] += count;
    } else if (range[endKey] >= at) {
      range[endKey] += count;
    }
    return range[endKey] >= range[startKey];
  }
  const end = at + count - 1;
  if (range[endKey] < at) return true;
  if (range[startKey] > end) {
    range[startKey] -= count;
    range[endKey] -= count;
    return true;
  }
  range[endKey] = Math.max(at - 1, range[endKey] - count);
  range[startKey] = Math.min(Math.max(at - 1, range[startKey]), range[endKey]);
  return range[endKey] >= range[startKey];
}

function shiftMerges(sheet: WorksheetModel, axis: 'row' | 'column', at: number, count: number, direction: 1 | -1): void {
  for (let index = sheet.merges.length - 1; index >= 0; index--) {
    const merge = sheet.merges[index]!;
    const keep = shiftRangeRef(merge.range, axis, at, count, direction);
    if (!keep) {
      sheet.merges.splice(index, 1);
      continue;
    }
    if (axis === 'row') {
      if (direction === 1 && merge.anchor.row >= at) merge.anchor.row += count;
      if (direction === -1) {
        if (merge.anchor.row >= at && merge.anchor.row < at + count) merge.anchor.row = merge.range.startRow;
        else if (merge.anchor.row >= at + count) merge.anchor.row -= count;
      }
    } else {
      if (direction === 1 && merge.anchor.column >= at) merge.anchor.column += count;
      if (direction === -1) {
        if (merge.anchor.column >= at && merge.anchor.column < at + count) merge.anchor.column = merge.range.startColumn;
        else if (merge.anchor.column >= at + count) merge.anchor.column -= count;
      }
    }
  }
}

function shiftRuleRanges(rules: Array<{ ranges: RangeRef[] }>, axis: 'row' | 'column', at: number, count: number, direction: 1 | -1): void {
  for (const rule of rules) {
    rule.ranges = rule.ranges.filter((range) => shiftRangeRef(range, axis, at, count, direction));
  }
}

function shiftFilter(sheet: WorksheetModel, axis: 'row' | 'column', at: number, count: number, direction: 1 | -1): void {
  if (!sheet.filter) return;
  if (!shiftRangeRef(sheet.filter.range, axis, at, count, direction)) {
    sheet.filter = undefined;
    return;
  }
  if (axis === 'column') {
    const next: typeof sheet.filter.criteria = {};
    for (const [key, condition] of Object.entries(sheet.filter.criteria)) {
      const column = Number(key);
      const shifted = shiftIndex(column, at, count, direction);
      if (shifted == null) continue;
      next[shifted] = { ...condition, column: shifted };
    }
    sheet.filter.criteria = next;
  }
}

function shiftFreeze(sheet: WorksheetModel, axis: 'row' | 'column', at: number, count: number, direction: 1 | -1): void {
  if (axis === 'row') {
    if (direction === 1 && sheet.freeze.ySplit >= at) sheet.freeze.ySplit += count;
    if (direction === -1 && sheet.freeze.ySplit > at) sheet.freeze.ySplit = Math.max(0, sheet.freeze.ySplit - count);
    if (direction === 1 && sheet.freeze.startRow >= at) sheet.freeze.startRow += count;
    if (direction === -1 && sheet.freeze.startRow > at) sheet.freeze.startRow = Math.max(0, sheet.freeze.startRow - count);
    return;
  }
  if (direction === 1 && sheet.freeze.xSplit >= at) sheet.freeze.xSplit += count;
  if (direction === -1 && sheet.freeze.xSplit > at) sheet.freeze.xSplit = Math.max(0, sheet.freeze.xSplit - count);
  if (direction === 1 && sheet.freeze.startColumn >= at) sheet.freeze.startColumn += count;
  if (direction === -1 && sheet.freeze.startColumn > at) sheet.freeze.startColumn = Math.max(0, sheet.freeze.startColumn - count);
}

function shiftHiddenAndSizes(sheet: WorksheetModel, axis: 'row' | 'column', at: number, count: number, direction: 1 | -1): void {
  if (axis === 'row') {
    remapIndexSet(sheet.hiddenRows, at, count, direction);
    remapSizeMap(sheet.rowHeights, at, count, direction);
    return;
  }
  remapIndexSet(sheet.hiddenColumns, at, count, direction);
  remapSizeMap(sheet.columnWidths, at, count, direction);
}

function remapIndexSet(set: Set<number>, at: number, count: number, direction: 1 | -1): void {
  const next = new Set<number>();
  for (const value of set) {
    const shifted = shiftIndex(value, at, count, direction);
    if (shifted != null) next.add(shifted);
  }
  set.clear();
  for (const value of next) set.add(value);
}

function remapSizeMap(map: Record<number, number>, at: number, count: number, direction: 1 | -1): void {
  const next: Record<number, number> = {};
  for (const [key, value] of Object.entries(map)) {
    const shifted = shiftIndex(Number(key), at, count, direction);
    if (shifted != null) next[shifted] = value;
  }
  for (const key of Object.keys(map)) delete map[Number(key)];
  Object.assign(map, next);
}

function shiftIndex(value: number, at: number, count: number, direction: 1 | -1): number | null {
  if (direction === 1) return value >= at ? value + count : value;
  if (value < at) return value;
  if (value < at + count) return null;
  return value - count;
}

function shiftSparklines(sheet: WorksheetModel, axis: 'row' | 'column', at: number, count: number, direction: 1 | -1): void {
  for (let index = sheet.sparklines.length - 1; index >= 0; index--) {
    const sparkline = sheet.sparklines[index]!;
    shiftRangeRef(sparkline.sourceRange, axis, at, count, direction);
    const position = axis === 'row' ? sparkline.anchor.row : sparkline.anchor.column;
    const shifted = shiftIndex(position, at, count, direction);
    if (shifted == null) {
      sheet.sparklines.splice(index, 1);
      continue;
    }
    if (axis === 'row') sparkline.anchor.row = shifted;
    else sparkline.anchor.column = shifted;
  }
}

function shiftPivots(sheet: WorksheetModel, axis: 'row' | 'column', at: number, count: number, direction: 1 | -1): void {
  for (const pivot of sheet.pivots) {
    shiftRangeRef(pivot.sourceRange, axis, at, count, direction);
    if (pivot.dataSource?.kind === 'worksheet-range') shiftRangeRef(pivot.dataSource.range, axis, at, count, direction);
    if (pivot.dataSource?.kind === 'worksheet-ranges') {
      for (const range of pivot.dataSource.ranges) shiftRangeRef(range, axis, at, count, direction);
    }
    if (pivot.targetAnchor) {
      const position = axis === 'row' ? pivot.targetAnchor.row : pivot.targetAnchor.column;
      const shifted = shiftIndex(position, at, count, direction);
      if (shifted != null) {
        if (axis === 'row') pivot.targetAnchor.row = shifted;
        else pivot.targetAnchor.column = shifted;
      }
    }
  }
}

function shiftCharts(sheet: WorksheetModel, axis: 'row' | 'column', at: number, count: number, direction: 1 | -1): void {
  for (const chart of sheet.charts) {
    for (const range of chart.sourceRanges) shiftRangeRef(range, axis, at, count, direction);
    if (chart.categoryRange) shiftRangeRef(chart.categoryRange, axis, at, count, direction);
    for (const series of chart.series ?? []) shiftRangeRef(series.range, axis, at, count, direction);
  }
}

function shiftDrawings(sheet: WorksheetModel, axis: 'row' | 'column', at: number, count: number, direction: 1 | -1): void {
  for (let index = sheet.drawings.length - 1; index >= 0; index--) {
    const drawing = sheet.drawings[index]!;
    shiftDrawingAnchor(drawing, axis, at, count, direction);
    syncDrawingPayload(sheet, drawing);
  }
}

function shiftDrawingAnchor(drawing: DrawingObject, axis: 'row' | 'column', at: number, count: number, direction: 1 | -1): void {
  if (drawing.anchor.kind === 'absolute') return;
  const start = axis === 'row' ? drawing.anchor.row : drawing.anchor.column;
  if (start != null) {
    const shifted = shiftIndex(start, at, count, direction);
    if (shifted == null) {
      drawing.anchor.row = axis === 'row' ? at : drawing.anchor.row;
      drawing.anchor.column = axis === 'column' ? at : drawing.anchor.column;
    } else if (axis === 'row') drawing.anchor.row = shifted;
    else drawing.anchor.column = shifted;
  }
  const end = axis === 'row' ? drawing.anchor.endRow : drawing.anchor.endColumn;
  if (end != null) {
    const shifted = shiftIndex(end, at, count, direction);
    if (shifted != null) {
      if (axis === 'row') drawing.anchor.endRow = shifted;
      else drawing.anchor.endColumn = shifted;
    }
  }
}

function syncDrawingPayload(sheet: WorksheetModel, drawing: DrawingObject): void {
  const bounds = { x: drawing.transform.x, y: drawing.transform.y, width: drawing.transform.width, height: drawing.transform.height };
  if (drawing.kind === 'chart') {
    const chart = sheet.charts.find((item) => item.id === drawing.payloadId);
    if (chart) chart.bounds = bounds;
  } else if (drawing.kind === 'shape') {
    const shape = sheet.shapes.find((item) => item.id === drawing.payloadId);
    if (shape) shape.bounds = bounds;
  } else if (drawing.kind === 'image') {
    const image = sheet.images.find((item) => item.id === drawing.payloadId);
    if (image) image.bounds = bounds;
  }
}

function shiftSheetTables(sheet: WorksheetModel, axis: 'row' | 'column', at: number, count: number, direction: 1 | -1): void {
  const kept = sheet.sheetTables.filter((table) => shiftRangeRef(table.range, axis, at, count, direction));
  sheet.sheetTables.splice(0, sheet.sheetTables.length, ...kept);
}

function shiftNotes(sheet: WorksheetModel, axis: 'row' | 'column', at: number, count: number, direction: 1 | -1): void {
  const next = new Map<string, CellNote>();
  for (const [key, note] of sheet.notes) {
    const [rowText, columnText] = key.split(':');
    const row = Number(rowText);
    const column = Number(columnText);
    const position = axis === 'row' ? row : column;
    const shifted = shiftIndex(position, at, count, direction);
    if (shifted == null) continue;
    const nextRow = axis === 'row' ? shifted : row;
    const nextColumn = axis === 'column' ? shifted : column;
    next.set(noteCellKey(nextRow, nextColumn), note);
  }
  sheet.notes.clear();
  for (const [key, note] of next) sheet.notes.set(key, note);
}

function shiftComments(sheet: WorksheetModel, axis: 'row' | 'column', at: number, count: number, direction: 1 | -1): void {
  const next: CommentThread[] = [];
  for (const thread of sheet.commentThreads) {
    const position = axis === 'row' ? thread.row : thread.column;
    const shifted = shiftIndex(position, at, count, direction);
    if (shifted == null) continue;
    next.push(axis === 'row' ? { ...thread, row: shifted } : { ...thread, column: shifted });
  }
  sheet.commentThreads.splice(0, sheet.commentThreads.length, ...next);
}

function shiftSpills(sheet: WorksheetModel, axis: 'row' | 'column', at: number, count: number, direction: 1 | -1): void {
  const next: SpillRange[] = [];
  for (const spill of sheet.spillRanges) {
    const keep = shiftRangeRef(spill.range, axis, at, count, direction);
    if (!keep) continue;
    const position = axis === 'row' ? spill.anchor.row : spill.anchor.column;
    const shifted = shiftIndex(position, at, count, direction);
    if (shifted == null) continue;
    if (axis === 'row') spill.anchor.row = shifted;
    else spill.anchor.column = shifted;
    next.push(spill);
  }
  sheet.spillRanges.splice(0, sheet.spillRanges.length, ...next);
}

function shiftProtection(sheet: WorksheetModel, axis: 'row' | 'column', at: number, count: number, direction: 1 | -1): void {
  const kept = sheet.protectionRules.filter((rule) => {
    if (!rule.range) return true;
    return shiftRangeRef(rule.range, axis, at, count, direction);
  });
  sheet.protectionRules.splice(0, sheet.protectionRules.length, ...kept);
}

function shiftBanded(sheet: WorksheetModel, axis: 'row' | 'column', at: number, count: number, direction: 1 | -1): void {
  if (!sheet.bandedRule) return;
  if (!shiftRangeRef(sheet.bandedRule.range, axis, at, count, direction)) sheet.bandedRule = undefined;
}

function rewriteFormulas(workbook: WorkbookModel, sheetId: string, axis: 'row' | 'column', at: number, count: number, direction: 1 | -1): void {
  const sheetName = workbook.getSheet(sheetId).name;
  for (const sheet of workbook.getSheets()) {
    sheet.cells.forEach((cell) => {
      if (!cell.formula) return;
      cell.formula = shiftFormulaText(cell.formula, sheetName, sheet.id === sheetId, axis, at, count, direction);
    });
  }
}

function rewriteDefinedNames(workbook: WorkbookModel, sheet: WorksheetModel, axis: 'row' | 'column', at: number, count: number, direction: 1 | -1): void {
  for (const [name, value] of Object.entries(workbook.definedNames)) {
    workbook.definedNames[name] = shiftFormulaText(value, sheet.name, true, axis, at, count, direction);
  }
}

export function shiftFormulaText(
  formula: string,
  sheetName: string,
  sameSheet: boolean,
  axis: 'row' | 'column',
  at: number,
  count: number,
  direction: 1 | -1,
): string {
  return formula.replace(/(?:((?:'[^']+'|[A-Za-z0-9_]+))!)?(\$?[A-Z]+)(\$?\d+)/g, (match, sheetPart: string | undefined, colPart: string, rowPart: string) => {
    const quoted = sheetPart?.startsWith("'") ? sheetPart.slice(1, -1) : sheetPart;
    const applies = !sheetPart ? sameSheet : quoted === sheetName;
    if (!applies) return match;
    const absCol = colPart.startsWith('$');
    const absRow = rowPart.startsWith('$');
    let colStr = absCol ? colPart.slice(1) : colPart;
    let rowNum = parseInt(absRow ? rowPart.slice(1) : rowPart, 10);
    if (axis === 'column' && !absCol) {
      const shifted = shiftIndex(parseColumnLabel(colStr), at, count, direction);
      if (shifted == null) return '#REF!';
      colStr = columnLabel(shifted);
    }
    if (axis === 'row' && !absRow) {
      const shifted = shiftIndex(rowNum - 1, at, count, direction);
      if (shifted == null) return '#REF!';
      rowNum = shifted + 1;
    }
    return `${sheetPart ? `${sheetPart}!` : ''}${absCol ? '$' : ''}${colStr}${absRow ? '$' : ''}${rowNum}`;
  });
}

function applyMoveRange(
  workbook: WorkbookModel,
  sheet: WorksheetModel,
  source: RangeRef,
  targetOrigin: { row: Row; column: Column },
): StructuralTransformResult {
  const rowDelta = targetOrigin.row - source.startRow;
  const colDelta = targetOrigin.column - source.startColumn;
  const extracted = sheet.cells.extractRegion(source.startRow, source.endRow, source.startColumn, source.endColumn);
  for (const item of extracted) {
    if (item.cell.formula) {
      item.cell.formula = shiftFormulaText(item.cell.formula, sheet.name, true, 'row', 0, rowDelta, rowDelta >= 0 ? 1 : -1);
      item.cell.formula = shiftFormulaText(item.cell.formula, sheet.name, true, 'column', 0, Math.abs(colDelta), colDelta >= 0 ? 1 : -1);
    }
    sheet.cells.set(item.row + rowDelta, item.column + colDelta, item.cell);
  }
  const moved: RangeRef = {
    sheetId: source.sheetId,
    startRow: source.startRow,
    endRow: source.endRow,
    startColumn: source.startColumn,
    endColumn: source.endColumn,
  };
  const relocate = (range: RangeRef): void => {
    if (range.sheetId !== source.sheetId) return;
    if (range.startRow >= source.startRow && range.endRow <= source.endRow && range.startColumn >= source.startColumn && range.endColumn <= source.endColumn) {
      range.startRow += rowDelta;
      range.endRow += rowDelta;
      range.startColumn += colDelta;
      range.endColumn += colDelta;
    }
  };
  for (const merge of sheet.merges) relocate(merge.range);
  for (const rule of [...sheet.conditionalFormats, ...sheet.dataValidations]) {
    for (const range of rule.ranges) relocate(range);
  }
  if (sheet.filter) relocate(sheet.filter.range);
  for (const table of sheet.sheetTables) relocate(table.range);
  void workbook;
  void moved;
  return { removedCells: extracted };
}

export function ensureDrawing(sheet: WorksheetModel, kind: DrawingObject['kind'], payloadId: string, transform: DrawingObject['transform']): DrawingObject {
  const existing = sheet.drawings.find((item) => item.payloadId === payloadId && item.kind === kind);
  if (existing) {
    existing.transform = { ...transform };
    return existing;
  }
  const drawing: DrawingObject = {
    id: `drawing-${payloadId}`,
    sheetId: sheet.id,
    kind,
    anchor: { kind: 'absolute' },
    transform: { ...transform },
    zIndex: sheet.drawings.length,
    payloadId,
  };
  sheet.drawings.push(drawing);
  return drawing;
}

export function usedSheetTables(sheet: WorksheetModel): SheetTableModel[] {
  return sheet.sheetTables;
}

export function usedProtection(sheet: WorksheetModel): ProtectionRule[] {
  return sheet.protectionRules;
}
