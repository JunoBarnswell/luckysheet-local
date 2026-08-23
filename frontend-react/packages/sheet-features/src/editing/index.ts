import type {
  CellData,
  CellStyle,
  FreezeModel,
  RangeRef,
  WorkbookModel,
  WorksheetModel,
} from '@react-sheets/core-model';
import { noteCellKey } from '@react-sheets/core-model';
import { StructuralTransform } from '@react-sheets/core-model';
import { formatValue } from '@react-sheets/number-format';
import type { CommandRuntime, MutationInfo } from '@react-sheets/command-runtime';
import { formatFormula, parseFormula, renameAstSheetReferences } from '@react-sheets/formula-engine';
import {
  copyRangeToClipboardData,
  shiftFormula,
  type ClipboardPayload,
} from '../clipboard';

export type PasteMode = 'all' | 'values' | 'formats' | 'formulas' | 'transpose';
export type FreezePreset = 'none' | 'firstRow' | 'firstColumn' | 'both';
export type GoToSpecialKind =
  | 'blanks'
  | 'constants'
  | 'formulas'
  | 'comments'
  | 'visible'
  | 'errors';

export interface MultiRangeStyleParams {
  sheetId: string;
  ranges: RangeRef[];
  style: Partial<CellStyle>;
}

export interface PasteRangeParams {
  sheetId: string;
  targetOrigin: { row: number; column: number };
  clipboard: ClipboardPayload;
  mode?: PasteMode;
}

export interface FormatCellsParams {
  sheetId: string;
  ranges: RangeRef[];
  numberFormat?: string;
  style?: Partial<CellStyle>;
}

export interface GoToParams {
  sheetId: string;
  reference: string;
}

export interface GoToSpecialParams {
  sheetId: string;
  range: RangeRef;
  kind: GoToSpecialKind;
}

export interface ShiftCellsParams {
  sheetId: string;
  range: RangeRef;
  direction: 'down' | 'up' | 'right' | 'left';
}

export interface SheetViewParams {
  sheetId: string;
  showGridlines?: boolean;
  showHeaders?: boolean;
  zoom?: number;
}

function normalizeRanges(ranges: RangeRef[]): RangeRef[] {
  return ranges.map((range) => ({
    sheetId: range.sheetId,
    startRow: Math.min(range.startRow, range.endRow),
    endRow: Math.max(range.startRow, range.endRow),
    startColumn: Math.min(range.startColumn, range.endColumn),
    endColumn: Math.max(range.startColumn, range.endColumn),
  }));
}

function forEachCell(
  sheet: WorksheetModel,
  range: RangeRef,
  fn: (row: number, column: number, cell: CellData | undefined) => void,
): void {
  for (let row = range.startRow; row <= range.endRow; row++) {
    for (let column = range.startColumn; column <= range.endColumn; column++) {
      fn(row, column, sheet.cells.get(row, column));
    }
  }
}

function parseA1Reference(reference: string): { row: number; column: number } | null {
  const match = reference.trim().match(/^\$?([A-Z]+)\$?(\d+)$/i);
  if (!match) return null;
  const colLetters = match[1]!.toUpperCase();
  let column = 0;
  for (const char of colLetters) column = column * 26 + char.charCodeAt(0) - 64;
  column -= 1;
  const row = Number(match[2]) - 1;
  if (!Number.isFinite(row) || row < 0 || column < 0) return null;
  return { row, column };
}

function shiftKind(direction: ShiftCellsParams['direction']): 'shift-cells-down' | 'shift-cells-up' | 'shift-cells-right' | 'shift-cells-left' {
  return `shift-cells-${direction}` as 'shift-cells-down';
}

export function rewriteFormulasForSheetRename(
  workbook: WorkbookModel,
  _sheetId: string,
  oldName: string,
  newName: string,
): Array<{ sheetId: string; row: number; column: number; previous?: CellData }> {
  const changes: Array<{ sheetId: string; row: number; column: number; previous?: CellData }> = [];
  for (const sheet of workbook.getSheets()) {
    sheet.cells.forEach((cell, row, column) => {
      if (!cell.formula) return;
      try {
        const nextFormula = formatFormula(renameAstSheetReferences(parseFormula(cell.formula), oldName, newName));
        if (nextFormula === cell.formula) return;
        changes.push({ sheetId: sheet.id, row, column, previous: structuredClone(cell) });
        sheet.cells.set(row, column, { ...cell, formula: nextFormula });
      } catch {
        // Unsupported formula syntax is left untouched. A lossy regex rewrite
        // could mutate string literals or structured references.
      }
    });
  }
  return changes;
}

export function resolveGoTo(workbook: WorkbookModel, params: GoToParams): { row: number; column: number } | null {
  const named = workbook.definedNames[params.reference];
  if (named) {
    const ref = named.includes('!') ? named.split('!').pop() ?? named : named;
    const parsed = parseA1Reference(ref.replace(/\$/g, ''));
    if (parsed) return parsed;
  }
  return parseA1Reference(params.reference);
}

export function resolveGoToSpecial(
  workbook: WorkbookModel,
  params: GoToSpecialParams,
): RangeRef[] {
  const sheet = workbook.getSheet(params.sheetId);
  const hits: RangeRef[] = [];
  for (let row = params.range.startRow; row <= params.range.endRow; row++) {
    for (let column = params.range.startColumn; column <= params.range.endColumn; column++) {
      const cell = sheet.cells.get(row, column);
      let match = false;
      switch (params.kind) {
        case 'blanks':
          match = !cell || cell.value == null || cell.value === '';
          break;
        case 'constants':
          match = Boolean(cell && cell.value != null && cell.value !== '' && !cell.formula);
          break;
        case 'formulas':
          match = Boolean(cell?.formula);
          break;
        case 'comments':
          match = Boolean(cell?.comment)
            || sheet.commentThreads.some((thread) => thread.row === row && thread.column === column)
            || sheet.notes.has(noteCellKey(row, column));
          break;
        case 'errors':
          match = Boolean(typeof cell?.value === 'string' && cell.value.startsWith('#'));
          break;
        case 'visible':
          match = !sheet.hiddenRows.has(row) && !sheet.hiddenColumns.has(column);
          break;
      }
      if (match) {
        hits.push({ sheetId: params.sheetId, startRow: row, endRow: row, startColumn: column, endColumn: column });
      }
    }
  }
  return hits;
}

function applyPasteCell(
  mode: PasteMode,
  source: CellData,
  target: CellData | undefined,
  rowOffset: number,
  colOffset: number,
  transposed: boolean,
): CellData {
  const destination = target ? structuredClone(target) : { value: null };
  const sourceFormula = source.formula
    ? shiftFormula(source.formula, transposed ? colOffset : rowOffset, transposed ? rowOffset : colOffset)
    : undefined;

  if (mode === 'values') {
    // Values means values only: no formula, style, number format or cached
    // display metadata may leak into the destination.
    return { value: source.value ?? null };
  }
  if (mode === 'formats') {
    return {
      ...destination,
      style: source.style ? structuredClone(source.style) : undefined,
      numberFormat: source.numberFormat,
    };
  }
  if (mode === 'formulas') {
    if (!sourceFormula) {
      return { ...destination, value: source.value ?? null, formula: undefined };
    }
    return { ...destination, value: null, formula: sourceFormula };
  }
  const next = structuredClone(source);
  if (sourceFormula) next.formula = sourceFormula;
  return next;
}

export function registerEditingCommands(runtime: CommandRuntime): void {
  runtime.registry.registerMutation('range.paste', (item, context) => {
    const params = item.params as PasteRangeParams & { values: CellData[][]; startRow: number; startColumn: number };
    const sheet = context.workbook.getSheet(params.sheetId);
    for (let rowOffset = 0; rowOffset < params.values.length; rowOffset++) {
      const rowValues = params.values[rowOffset] ?? [];
      for (let columnOffset = 0; columnOffset < rowValues.length; columnOffset++) {
        const value = rowValues[columnOffset];
        if (value) sheet.cells.set(params.startRow + rowOffset, params.startColumn + columnOffset, value);
      }
    }
  });

  runtime.registry.registerCommand<PasteRangeParams>({
    id: 'sheet.range.paste',
    execute: (params, context) => {
      const sheet = context.workbook.getSheet(params.sheetId);
      const mode = params.mode ?? 'all';
      let sourceValues = params.clipboard.values;
      if (mode === 'transpose') {
        sourceValues = [];
        const src = params.clipboard.values;
        for (let c = 0; c < (src[0]?.length ?? 0); c++) {
          sourceValues.push(src.map((row) => structuredClone(row[c] ?? { value: null })));
        }
      }
      const previous: Array<{ row: number; column: number; value?: CellData }> = [];
      const affectedRanges: RangeRef[] = [];
      const values: CellData[][] = [];
      for (let rowOffset = 0; rowOffset < sourceValues.length; rowOffset++) {
        const rowValues = sourceValues[rowOffset] ?? [];
        const outRow: CellData[] = [];
        for (let columnOffset = 0; columnOffset < rowValues.length; columnOffset++) {
          const row = params.targetOrigin.row + rowOffset;
          const column = params.targetOrigin.column + columnOffset;
          previous.push({ row, column, value: sheet.cells.get(row, column) });
          affectedRanges.push({ sheetId: params.sheetId, startRow: row, endRow: row, startColumn: column, endColumn: column });
              outRow.push(
                applyPasteCell(
                  mode === 'transpose' ? 'all' : mode,
                  rowValues[columnOffset] ?? { value: null },
                  sheet.cells.get(row, column),
                  rowOffset,
                  columnOffset,
                  mode === 'transpose',
                ),
              );
        }
        values.push(outRow);
      }
      const pasteMode = mode === 'transpose' ? 'all' : mode;
      context.applyMutation({
        id: 'range.paste',
        unitId: context.workbook.unitId,
        sheetId: params.sheetId,
        params: { ...params, values, startRow: params.targetOrigin.row, startColumn: params.targetOrigin.column, mode: pasteMode },
        affectedRanges,
        inverse: previous.map((item) => ({
          id: 'cell.restore' as const,
          unitId: context.workbook.unitId,
          sheetId: params.sheetId,
          params: { row: item.row, column: item.column, previous: item.value },
          affectedRanges: [{ sheetId: params.sheetId, startRow: item.row, endRow: item.row, startColumn: item.column, endColumn: item.column }],
        })),
        apply: () => {
          for (let rowOffset = 0; rowOffset < values.length; rowOffset++) {
            for (let columnOffset = 0; columnOffset < (values[rowOffset]?.length ?? 0); columnOffset++) {
              const cell = values[rowOffset]![columnOffset]!;
              sheet.cells.set(params.targetOrigin.row + rowOffset, params.targetOrigin.column + columnOffset, cell);
            }
          }
        },
      });
      return { operationId: context.operationId, mutationCount: 1, affectedRanges };
    },
  });

  runtime.registry.registerCommand<MultiRangeStyleParams>({
    id: 'sheet.style.setMulti',
    execute: (params, context) => {
      const ranges = normalizeRanges(params.ranges);
      if (ranges.length === 0) return { operationId: context.operationId, mutationCount: 0, affectedRanges: [] };
      let lastResult = { operationId: context.operationId, mutationCount: 0, affectedRanges: [] as RangeRef[] };
      for (const range of ranges) {
        lastResult = runtime.execute('sheet.style.set', { sheetId: params.sheetId, range, style: params.style });
      }
      return lastResult;
    },
  });

  runtime.registry.registerCommand<FormatCellsParams>({
    id: 'sheet.format.set',
    execute: (params, context) => {
      const style: Partial<CellStyle> = { ...params.style };
      if (params.numberFormat !== undefined) style.numberFormat = params.numberFormat;
      return runtime.execute('sheet.style.setMulti', {
        sheetId: params.sheetId,
        ranges: params.ranges,
        style,
      });
    },
  });

  runtime.registry.registerCommand<{ sheetId: string; ranges: RangeRef[]; numberFormat: string }>({
    id: 'sheet.numberFormat.apply',
    execute: (params, context) => {
      const sheet = context.workbook.getSheet(params.sheetId);
      const previous: Array<{ row: number; column: number; value?: CellData }> = [];
      const affectedRanges = normalizeRanges(params.ranges);
      for (const range of affectedRanges) {
        forEachCell(sheet, range, (row, column, cell) => {
          previous.push({ row, column, value: cell ? structuredClone(cell) : undefined });
        });
      }
      context.applyMutation({
        id: 'style.set',
        unitId: context.workbook.unitId,
        sheetId: params.sheetId,
        params,
        affectedRanges,
        inverse: previous.map((item) => ({
          id: 'cell.restore' as const,
          unitId: context.workbook.unitId,
          sheetId: params.sheetId,
          params: { row: item.row, column: item.column, previous: item.value },
          affectedRanges: [{ sheetId: params.sheetId, startRow: item.row, endRow: item.row, startColumn: item.column, endColumn: item.column }],
        })),
        apply: () => {
          for (const range of affectedRanges) {
            forEachCell(sheet, range, (row, column, cell) => {
              let next = cell ? { ...cell } : { value: null as CellData['value'] };
              next.numberFormat = params.numberFormat;
              if (next.style) next.style = { ...next.style, numberFormat: params.numberFormat };
              else next.style = { numberFormat: params.numberFormat };
              next.displayValue = formatValue(next.value, params.numberFormat);
              sheet.cells.set(row, column, next);
            });
          }
        },
      });
      return { operationId: context.operationId, mutationCount: 1, affectedRanges };
    },
  });

  runtime.registry.registerCommand<{ sheetId: string; preset: FreezePreset }>({
    id: 'sheet.freeze.preset',
    execute: (params, context) => {
      const freeze: FreezeModel = { xSplit: 0, ySplit: 0, startRow: 0, startColumn: 0 };
      if (params.preset === 'firstRow' || params.preset === 'both') {
        freeze.ySplit = 1;
        freeze.startRow = 1;
      }
      if (params.preset === 'firstColumn' || params.preset === 'both') {
        freeze.xSplit = 1;
        freeze.startColumn = 1;
      }
      return runtime.execute('sheet.freeze.set', { sheetId: params.sheetId, freeze });
    },
  });

  runtime.registry.registerMutation('view.set', (item, context) => {
    const params = item.params as SheetViewParams;
    const sheet = context.workbook.getSheet(params.sheetId);
    if (params.showGridlines !== undefined) sheet.showGridlines = params.showGridlines;
    if (params.showHeaders !== undefined) sheet.showHeaders = params.showHeaders;
    if (params.zoom !== undefined) sheet.zoom = params.zoom;
  });

  runtime.registry.registerCommand<SheetViewParams>({
    id: 'sheet.view.set',
    execute: (params, context) => {
      const sheet = context.workbook.getSheet(params.sheetId);
      const previous = { showGridlines: sheet.showGridlines, showHeaders: sheet.showHeaders, zoom: sheet.zoom };
      const affectedRanges: RangeRef[] = [];
      context.applyMutation({
        id: 'view.set',
        unitId: context.workbook.unitId,
        sheetId: params.sheetId,
        params,
        affectedRanges,
        inverse: [{
          id: 'view.set',
          unitId: context.workbook.unitId,
          sheetId: params.sheetId,
          params: { sheetId: params.sheetId, ...previous },
          affectedRanges,
        }],
        apply: () => {
          if (params.showGridlines !== undefined) sheet.showGridlines = params.showGridlines;
          if (params.showHeaders !== undefined) sheet.showHeaders = params.showHeaders;
          if (params.zoom !== undefined) sheet.zoom = Math.max(25, Math.min(400, params.zoom));
        },
      });
      return { operationId: context.operationId, mutationCount: 1, affectedRanges };
    },
  });

  runtime.registry.registerCommand<{ sheetId: string; row?: number; column?: number; rows?: number[]; columns?: number[] }>({
    id: 'sheet.autofit',
    execute: (params, context) => {
      const sheet = context.workbook.getSheet(params.sheetId);
      const affectedRanges: RangeRef[] = [];
      const rows = params.rows ?? (params.row !== undefined ? [params.row] : []);
      const columns = params.columns ?? (params.column !== undefined ? [params.column] : []);
      for (const row of rows) {
        let maxHeight = 32;
        for (let column = 0; column < sheet.columnCount; column++) {
          const cell = sheet.cells.get(row, column);
          const text = cell?.displayValue ?? (cell?.value == null ? '' : String(cell.value));
          maxHeight = Math.max(maxHeight, 16 + Math.ceil(text.length / 12) * 14);
        }
        const prev = sheet.rowHeights[row] ?? 32;
        context.applyMutation({
          id: 'row.resize',
          unitId: context.workbook.unitId,
          sheetId: params.sheetId,
          params: { sheetId: params.sheetId, row, height: maxHeight },
          affectedRanges,
          inverse: [{ id: 'row.resize', unitId: context.workbook.unitId, sheetId: params.sheetId, params: { sheetId: params.sheetId, row, height: prev }, affectedRanges }],
          apply: () => { sheet.rowHeights[row] = maxHeight; },
        });
      }
      for (const column of columns) {
        let maxWidth = 128;
        sheet.cells.forEach((cell, row, col) => {
          if (col !== column) return;
          const text = cell.displayValue ?? (cell.value == null ? '' : String(cell.value));
          maxWidth = Math.max(maxWidth, 24 + text.length * 8);
        });
        const prev = sheet.columnWidths[column] ?? 128;
        context.applyMutation({
          id: 'column.resize',
          unitId: context.workbook.unitId,
          sheetId: params.sheetId,
          params: { sheetId: params.sheetId, column, width: maxWidth },
          affectedRanges,
          inverse: [{ id: 'column.resize', unitId: context.workbook.unitId, sheetId: params.sheetId, params: { sheetId: params.sheetId, column, width: prev }, affectedRanges }],
          apply: () => { sheet.columnWidths[column] = maxWidth; },
        });
      }
      return { operationId: context.operationId, mutationCount: rows.length + columns.length, affectedRanges };
    },
  });

  runtime.registry.registerMutation('cells.shifted', (item, context) => {
    const params = item.params as ShiftCellsParams;
    StructuralTransform.apply(context.workbook, {
      kind: shiftKind(params.direction),
      sheetId: params.sheetId,
      at: 0,
      count: 0,
      sourceRange: params.range,
    });
  });
  runtime.registry.registerMutation('cells.shifted.restore', (item, context) => {
    const params = item.params as {
      sheetId: string;
      range: RangeRef;
      cells: Array<{ row: number; column: number; cell: CellData }>;
    };
    const sheet = context.workbook.getSheet(params.sheetId);
    forEachCell(sheet, params.range, (row, column) => sheet.cells.delete(row, column));
    for (const entry of params.cells) sheet.cells.set(entry.row, entry.column, structuredClone(entry.cell));
  });

  runtime.registry.registerCommand<ShiftCellsParams>({
    id: 'sheet.cells.shift',
    execute: (params, context) => {
      const sheet = context.workbook.getSheet(params.sheetId);
      const snapshot: Array<{ row: number; column: number; cell: CellData }> = [];
      forEachCell(sheet, params.range, (row, column, cell) => {
        if (cell) snapshot.push({ row, column, cell: structuredClone(cell) });
      });
      const affectedRanges: RangeRef[] = [params.range];
      context.applyMutation({
        id: 'cells.shifted',
        unitId: context.workbook.unitId,
        sheetId: params.sheetId,
        params,
        affectedRanges,
        inverse: [{
          id: 'cells.shifted.restore' as const,
          unitId: context.workbook.unitId,
          sheetId: params.sheetId,
          params: { sheetId: params.sheetId, range: params.range, cells: snapshot },
          affectedRanges,
        }],
        apply: () => {
          StructuralTransform.apply(context.workbook, {
            kind: shiftKind(params.direction),
            sheetId: params.sheetId,
            at: 0,
            count: 0,
            sourceRange: params.range,
          });
        },
      });
      return { operationId: context.operationId, mutationCount: 1, affectedRanges };
    },
  });

  runtime.registry.registerMutation('sheet.duplicated', (item, context) => {
    const params = item.params as { sourceSheetId: string; newId: string; newName: string };
    context.workbook.duplicateSheet(params.sourceSheetId, params.newId, params.newName);
  });

  runtime.registry.registerCommand<{ sourceSheetId: string; newId: string; newName: string }>({
    id: 'sheet.duplicate',
    execute: (params, context) => {
      const affectedRanges: RangeRef[] = [];
      context.applyMutation({
        id: 'sheet.duplicated',
        unitId: context.workbook.unitId,
        sheetId: params.newId,
        params,
        affectedRanges,
        inverse: [{ id: 'sheet.remove', unitId: context.workbook.unitId, sheetId: params.newId, params: { id: params.newId }, affectedRanges }],
        apply: () => context.workbook.duplicateSheet(params.sourceSheetId, params.newId, params.newName),
      });
      return { operationId: context.operationId, mutationCount: 1, affectedRanges };
    },
  });

  runtime.registry.registerMutation('sheet.hidden', (item, context) => {
    context.workbook.getSheet(item.sheetId).hidden = true;
  });
  runtime.registry.registerMutation('sheet.unhidden', (item, context) => {
    context.workbook.getSheet(item.sheetId).hidden = false;
  });

  runtime.registry.registerCommand<{ sheetId: string }>({
    id: 'sheet.hide',
    execute: (params, context) => {
      const visible = context.workbook.getVisibleSheets();
      if (visible.length <= 1 && !context.workbook.getSheet(params.sheetId).hidden) {
        throw new Error('Cannot hide the only visible worksheet');
      }
      const affectedRanges: RangeRef[] = [];
      context.applyMutation({
        id: 'sheet.hidden',
        unitId: context.workbook.unitId,
        sheetId: params.sheetId,
        params,
        affectedRanges,
        inverse: [{ id: 'sheet.unhidden', unitId: context.workbook.unitId, sheetId: params.sheetId, params, affectedRanges }],
        apply: () => { context.workbook.getSheet(params.sheetId).hidden = true; },
      });
      return { operationId: context.operationId, mutationCount: 1, affectedRanges };
    },
  });

  runtime.registry.registerCommand<{ sheetId: string }>({
    id: 'sheet.unhide',
    execute: (params, context) => {
      const affectedRanges: RangeRef[] = [];
      context.applyMutation({
        id: 'sheet.unhidden',
        unitId: context.workbook.unitId,
        sheetId: params.sheetId,
        params,
        affectedRanges,
        inverse: [{ id: 'sheet.hidden', unitId: context.workbook.unitId, sheetId: params.sheetId, params, affectedRanges }],
        apply: () => { context.workbook.getSheet(params.sheetId).hidden = false; },
      });
      return { operationId: context.operationId, mutationCount: 1, affectedRanges };
    },
  });

  runtime.registry.registerMutation('sheet.reordered', (item, context) => {
    const params = item.params as { sheetId: string; toIndex: number };
    context.workbook.reorderSheet(params.sheetId, params.toIndex);
  });

  runtime.registry.registerCommand<{ sheetId: string; toIndex: number }>({
    id: 'sheet.reorder',
    execute: (params, context) => {
      const previous = [...context.workbook.sheetOrder];
      const fromIndex = previous.indexOf(params.sheetId);
      const affectedRanges: RangeRef[] = [];
      context.applyMutation({
        id: 'sheet.reordered',
        unitId: context.workbook.unitId,
        sheetId: params.sheetId,
        params,
        affectedRanges,
        inverse: [{ id: 'sheet.reordered', unitId: context.workbook.unitId, sheetId: params.sheetId, params: { sheetId: params.sheetId, toIndex: fromIndex }, affectedRanges }],
        apply: () => context.workbook.reorderSheet(params.sheetId, params.toIndex),
      });
      return { operationId: context.operationId, mutationCount: 1, affectedRanges };
    },
  });

  runtime.registry.registerMutation('sheet.tabColor', (item, context) => {
    const params = item.params as { sheetId: string; color?: string };
    context.workbook.getSheet(params.sheetId).tabColor = params.color;
  });

  runtime.registry.registerCommand<{ sheetId: string; color?: string }>({
    id: 'sheet.tabColor.set',
    execute: (params, context) => {
      const sheet = context.workbook.getSheet(params.sheetId);
      const previous = sheet.tabColor;
      const affectedRanges: RangeRef[] = [];
      context.applyMutation({
        id: 'sheet.tabColor',
        unitId: context.workbook.unitId,
        sheetId: params.sheetId,
        params,
        affectedRanges,
        inverse: [{ id: 'sheet.tabColor', unitId: context.workbook.unitId, sheetId: params.sheetId, params: { sheetId: params.sheetId, color: previous }, affectedRanges }],
        apply: () => { sheet.tabColor = params.color; },
      });
      return { operationId: context.operationId, mutationCount: 1, affectedRanges };
    },
  });

  runtime.registry.registerCommand<GoToParams>({
    id: 'navigation.goto',
    execute: (params, context) => {
      const target = resolveGoTo(context.workbook, params);
      if (!target) throw new Error(`Invalid reference: ${params.reference}`);
      return {
        operationId: context.operationId,
        mutationCount: 0,
        affectedRanges: [{
          sheetId: params.sheetId,
          startRow: target.row,
          endRow: target.row,
          startColumn: target.column,
          endColumn: target.column,
        }],
      };
    },
  });

  runtime.registry.registerCommand<GoToSpecialParams>({
    id: 'navigation.gotoSpecial',
    execute: (params, context) => {
      const ranges = resolveGoToSpecial(context.workbook, params);
      return { operationId: context.operationId, mutationCount: 0, affectedRanges: ranges };
    },
  });

}

export function buildClipboardFromRange(workbook: WorkbookModel, range: RangeRef): ClipboardPayload {
  return copyRangeToClipboardData(workbook, range);
}

export function restoreCellFromMutation(
  workbook: WorkbookModel,
  item: MutationInfo<{ row: number; column: number; previous?: CellData }>,
): void {
  const sheet = workbook.getSheet(item.sheetId);
  const { row, column, previous } = item.params;
  if (previous) sheet.cells.set(row, column, previous);
  else sheet.cells.delete(row, column);
}
