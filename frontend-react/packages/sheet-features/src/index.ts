import type {
  BandedRule,
  CellData,
  CellStyle,
  ConditionalFormatRule,
  DataValidationRule,
  AutoFilterModel,
  WorksheetPane,
  MergeSpan,
  RangeRef,
  WorkbookTableModel,
  WorkbookModel,
  WorksheetModel,
  StructuralTransformParams,
  DefinedNameModel,
} from '@react-sheets/core-model';
import { StructuralTransform, normalizeDefinedNameModel } from '@react-sheets/core-model';
import type { CommandRuntime, MutationInfo } from '@react-sheets/command-runtime';
import { isSpillChild } from '@react-sheets/formula-engine';
import { shiftFormula } from './clipboard';
import { buildCellFromText } from './text-input';
import { registerEditingCommands, rewriteFormulasForSheetRename } from './editing';
import { registerDataToolCommands, normalizeConditionalFormatRule, normalizeDataValidationRule, validateDataInput } from './data-features';
import { registerSheetTableCommands } from './sheet-table-commands';
import { validateFilterOwnership } from './sheet-table-features';
import { registerOutlineCommands } from './outline-commands';
import { registerHomeCommands } from './home-commands';

function snapshotCellRegion(
  sheet: WorksheetModel,
  startRow: number,
  endRow: number,
  startColumn: number,
  endColumn: number,
): Array<{ row: number; column: number; cell: CellData }> {
  const extracted: Array<{ row: number; column: number; cell: CellData }> = [];
  sheet.cells.forEach((cell, row, column) => {
    if (row >= startRow && row <= endRow && column >= startColumn && column <= endColumn) {
      extracted.push({ row, column, cell: structuredClone(cell) });
    }
  });
  return extracted;
}

function applyStructuralTransform(workbook: WorkbookModel, params: StructuralTransformParams): void {
  StructuralTransform.apply(workbook, params);
}

export * from './clipboard';
export * from './data-features';
export * from './editing';
export * from './sheet-table-features';
export * from './sheet-table-commands';
export * from './outline-commands';
export * from './outline-features';
export * from './text-input';
export * from './home-commands';


export interface SetCellValueParams {
  sheetId: string;
  row: number;
  column: number;
  value: CellData;
}

/** Host/UI text commit contract. The command owns parsing, validation and
 * content/presentation merging; callers must pass the raw text unchanged. */
export interface CommitTextParams {
  sheetId: string;
  row: number;
  column: number;
  text: string;
  style?: Partial<CellStyle>;
}

export interface AddTableParams extends WorkbookTableModel {}

export interface SetRangeValuesParams {
  sheetId: string;
  startRow: number;
  startColumn: number;
  values: CellData[][];
}

export interface ClearRangeParams {
  sheetId: string;
  range: RangeRef;
  mode?: 'all' | 'contents' | 'formats' | 'notes' | 'hyperlinks';
}

interface ClearRangeRestoreParams {
  sheetId: string;
  range: RangeRef;
  cells: Array<{ row: number; column: number; value?: CellData }>;
  notes: Array<{ row: number; column: number; note: import('@react-sheets/core-model').CellNote }>;
  hyperlinks: Array<{ row: number; column: number; hyperlink: import('@react-sheets/core-model').CellHyperlink }>;
  comments: import('@react-sheets/core-model').CommentThread[];
}

export interface AddSheetParams {
  id: string;
  name: string;
  rowCount?: number;
  columnCount?: number;
}

export interface RenameSheetParams {
  sheetId: string;
  name: string;
}

export interface RenameWorkbookParams { name: string; }

export interface InsertRowParams {
  sheetId: string;
  row: number;
  count: number;
}

export interface DeleteRowParams {
  sheetId: string;
  row: number;
  count: number;
}

export interface ResizeRowParams {
  sheetId: string;
  row: number;
  heightPx: number;
}

export interface InsertColumnParams {
  sheetId: string;
  column: number;
  count: number;
}

export interface DeleteColumnParams {
  sheetId: string;
  column: number;
  count: number;
}

export interface ResizeColumnParams {
  sheetId: string;
  column: number;
  widthPx: number;
}

export interface ColumnsVisibilityParams {
  sheetId: string;
  states: Array<{ column: number; hidden: boolean }>;
}

export interface SetMergeParams {
  sheetId: string;
  range: RangeRef;
}

export interface RemoveMergeParams {
  sheetId: string;
  range: RangeRef;
}

export interface SetFreezeParams {
  sheetId: string;
  pane: WorksheetPane;
}

export interface SetRangeStyleParams {
  sheetId: string;
  range: RangeRef;
  style: Partial<CellStyle>;
}

export interface SortRangeParams {
  sheetId: string;
  range: RangeRef;
  sortColumn: number;
  ascending: boolean;
  hasHeader?: boolean;
}

export interface AutoFillParams {
  sheetId: string;
  sourceRange: RangeRef;
  targetRange: RangeRef;
}

export interface AddConditionalFormatParams {
  sheetId: string;
  rule: ConditionalFormatRule;
}

export interface AddDataValidationParams {
  sheetId: string;
  rule: DataValidationRule;
}

function cellRange(params: SetCellValueParams): RangeRef[] {
  return [
    {
      sheetId: params.sheetId,
      startRow: params.row,
      endRow: params.row,
      startColumn: params.column,
      endColumn: params.column,
    },
  ];
}

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === 'object' && value !== null;
}

function isWorksheetPane(value: unknown): value is WorksheetPane {
  if (!isRecord(value) || !['none', 'frozen', 'split'].includes(String(value.kind))) return false;
  if (value.kind === 'none') return true;
  return typeof value.xSplit === 'number' && Number.isFinite(value.xSplit) && value.xSplit >= 0
    && typeof value.ySplit === 'number' && Number.isFinite(value.ySplit) && value.ySplit >= 0
    && Number.isSafeInteger(value.startRow) && Number(value.startRow) >= 0
    && Number.isSafeInteger(value.startColumn) && Number(value.startColumn) >= 0;
}

function isColumnVisibilityMutation(value: unknown): value is ColumnsVisibilityParams {
  return isRecord(value) && typeof value.sheetId === 'string' && Array.isArray(value.states) && value.states.length > 0
    && value.states.every((state) => isRecord(state) && Number.isSafeInteger(state.column) && Number(state.column) >= 0 && typeof state.hidden === 'boolean');
}

function isRange(value: unknown): value is RangeRef {
  if (!isRecord(value)) return false;
  return typeof value.sheetId === 'string'
    && Number.isInteger(value.startRow) && Number.isInteger(value.endRow)
    && Number.isInteger(value.startColumn) && Number.isInteger(value.endColumn)
    && Number(value.startRow) >= 0 && Number(value.endRow) >= Number(value.startRow)
    && Number(value.startColumn) >= 0 && Number(value.endColumn) >= Number(value.startColumn);
}

function isCellData(value: unknown): value is CellData {
  return isRecord(value)
    && ('value' in value || typeof value.formula === 'string');
}

function isCellSetMutation(value: unknown): value is SetCellValueParams {
  return isRecord(value)
    && typeof value.sheetId === 'string'
    && Number.isInteger(value.row) && Number(value.row) >= 0
    && Number.isInteger(value.column) && Number(value.column) >= 0
    && isCellData(value.value);
}

function isCellRestoreMutation(value: unknown): value is { sheetId: string; row: number; column: number; previous?: CellData } {
  return isRecord(value)
    && typeof value.sheetId === 'string'
    && Number.isInteger(value.row) && Number(value.row) >= 0
    && Number.isInteger(value.column) && Number(value.column) >= 0
    && (value.previous === undefined || isCellData(value.previous));
}

function sheetScopeRange(value: { sheetId: string }): RangeRef[] {
  return [{ sheetId: value.sheetId, startRow: 0, endRow: 0, startColumn: 0, endColumn: 0 }];
}

function isRenameWorkbookMutation(value: unknown): value is RenameWorkbookParams {
  return isRecord(value) && typeof value.name === 'string' && value.name.trim().length > 0;
}

function isRenameSheetMutation(value: unknown): value is RenameSheetParams {
  return isRecord(value) && typeof value.sheetId === 'string' && typeof value.name === 'string' && value.name.trim().length > 0;
}

function isAddSheetMutation(value: unknown): value is AddSheetParams {
  return isRecord(value) && typeof value.id === 'string' && value.id.length > 0 && typeof value.name === 'string'
    && (value.rowCount === undefined || (Number.isInteger(value.rowCount) && Number(value.rowCount) > 0))
    && (value.columnCount === undefined || (Number.isInteger(value.columnCount) && Number(value.columnCount) > 0));
}

function isSheetIdMutation(value: unknown): value is { id: string } {
  return isRecord(value) && typeof value.id === 'string' && value.id.length > 0;
}

function isSheetRestoreMutation(value: unknown): value is { sheet: import('@react-sheets/core-model').SheetSnapshot; index?: number } {
  return isRecord(value) && isRecord(value.sheet) && typeof value.sheet.id === 'string' && typeof value.sheet.name === 'string'
    && (value.index === undefined || (Number.isSafeInteger(value.index) && Number(value.index) >= 0));
}

function isWorkbookTableMutation(value: unknown): value is WorkbookTableModel {
  return isRecord(value) && typeof value.id === 'string' && typeof value.name === 'string'
    && Number.isInteger(value.rowCount) && Number(value.rowCount) >= 0
    && Number.isInteger(value.blockSize) && Number(value.blockSize) > 0
    && Number.isInteger(value.revision) && Number(value.revision) >= 0
    && Array.isArray(value.fields) && Array.isArray(value.blocks);
}

function isTableRemoveMutation(value: unknown): value is { tableId: string; range?: RangeRef } {
  return isRecord(value) && typeof value.tableId === 'string' && (value.range === undefined || isRange(value.range));
}

function workbookTableRanges(value: WorkbookTableModel): RangeRef[] {
  return value.sourceRange ? [structuredClone(value.sourceRange)] : [];
}

function tableRemoveRanges(value: { range?: RangeRef }): RangeRef[] {
  return value.range ? [structuredClone(value.range)] : [];
}

function isSetRangeMutation(value: unknown): value is SetRangeValuesParams {
  if (!isRecord(value) || typeof value.sheetId !== 'string'
    || !Number.isInteger(value.startRow) || Number(value.startRow) < 0
    || !Number.isInteger(value.startColumn) || Number(value.startColumn) < 0
    || !Array.isArray(value.values)) return false;
  return value.values.every((row) => Array.isArray(row) && row.every(isCellData));
}

function setRangeAffectedRanges(value: SetRangeValuesParams): RangeRef[] {
  const rowCount = value.values.length;
  const columnCount = Math.max(1, ...value.values.map((row) => row.length));
  return [{
    sheetId: value.sheetId,
    startRow: value.startRow,
    endRow: value.startRow + Math.max(0, rowCount - 1),
    startColumn: value.startColumn,
    endColumn: value.startColumn + columnCount - 1,
  }];
}

function isClearRangeMutation(value: unknown): value is ClearRangeParams {
  return isRecord(value) && typeof value.sheetId === 'string' && isRange(value.range)
    && (value.mode === undefined || value.mode === 'all' || value.mode === 'contents' || value.mode === 'formats' || value.mode === 'notes' || value.mode === 'hyperlinks');
}

function isClearRangeRestoreMutation(value: unknown): value is ClearRangeRestoreParams {
  return isRecord(value) && typeof value.sheetId === 'string' && isRange(value.range)
    && Array.isArray(value.cells) && value.cells.every((entry) => isRecord(entry) && Number.isInteger(entry.row) && Number.isInteger(entry.column) && (entry.value === undefined || isCellData(entry.value)))
    && Array.isArray(value.notes) && Array.isArray(value.hyperlinks) && Array.isArray(value.comments);
}

function isStyleMutation(value: unknown): value is SetRangeStyleParams | { sheetId: string; ranges: RangeRef[]; numberFormat: string } {
  if (!isRecord(value) || typeof value.sheetId !== 'string') return false;
  if (isRange(value.range)) return isRecord(value.style) || value.style === undefined || typeof value.numberFormat === 'string';
  return Array.isArray(value.ranges) && value.ranges.length > 0 && value.ranges.every(isRange)
    && typeof value.numberFormat === 'string';
}

function styleAffectedRanges(value: SetRangeStyleParams | { sheetId: string; ranges: RangeRef[]; numberFormat: string }): RangeRef[] {
  return 'range' in value ? [structuredClone(value.range)] : value.ranges.map((range) => structuredClone(range));
}

function isSheetAtCountMutation(value: unknown): value is { sheetId: string; at: number; count: number } {
  return isRecord(value) && typeof value.sheetId === 'string'
    && Number.isInteger(value.at) && Number(value.at) >= 0
    && Number.isInteger(value.count) && Number(value.count) > 0;
}

function structuralAffectedRanges(value: { sheetId: string; at: number; count: number }): RangeRef[] {
  return [{ sheetId: value.sheetId, startRow: value.at, endRow: value.at + value.count - 1, startColumn: 0, endColumn: 0 }];
}

function columnStructuralAffectedRanges(value: { sheetId: string; at: number; count: number }): RangeRef[] {
  return [{ sheetId: value.sheetId, startRow: 0, endRow: 0, startColumn: value.at, endColumn: value.at + value.count - 1 }];
}

function isSheetIndexMutation(value: unknown): value is { sheetId: string; index: number } {
  return isRecord(value) && typeof value.sheetId === 'string' && Number.isInteger(value.index) && Number(value.index) >= 0;
}

function isSheetIndicesMutation(value: unknown): value is { sheetId: string; indices: number[] } {
  return isRecord(value) && typeof value.sheetId === 'string' && Array.isArray(value.indices)
    && value.indices.every((index) => Number.isInteger(index) && Number(index) >= 0);
}

function rowAffectedRange(value: { sheetId: string; row: number }): RangeRef[] {
  return [{ sheetId: value.sheetId, startRow: value.row, endRow: value.row, startColumn: 0, endColumn: 0 }];
}

function columnAffectedRange(value: { sheetId: string; column: number }): RangeRef[] {
  return [{ sheetId: value.sheetId, startRow: 0, endRow: 0, startColumn: value.column, endColumn: value.column }];
}

function isSheetViewMutation(value: unknown): value is { sheetId: string; showGridlines?: boolean; showHeaders?: boolean; zoom?: number } {
  return isRecord(value) && typeof value.sheetId === 'string'
    && (value.showGridlines === undefined || typeof value.showGridlines === 'boolean')
    && (value.showHeaders === undefined || typeof value.showHeaders === 'boolean')
    && (value.zoom === undefined || (typeof value.zoom === 'number' && Number.isFinite(value.zoom) && value.zoom > 0));
}

function isFilterMutation(value: unknown): value is { sheetId: string; autoFilter: AutoFilterModel } {
  return isRecord(value) && typeof value.sheetId === 'string' && isRecord(value.autoFilter)
    && typeof value.autoFilter.sheetId === 'string' && isRange(value.autoFilter.range)
    && isRecord(value.autoFilter.columns);
}

function isFilterRemoveMutation(value: unknown): value is { sheetId: string; range?: RangeRef } {
  return isRecord(value) && typeof value.sheetId === 'string' && (value.range === undefined || isRange(value.range));
}

function isConditionalAddMutation(value: unknown): value is AddConditionalFormatParams {
  return isRecord(value) && typeof value.sheetId === 'string' && isRecord(value.rule)
    && typeof value.rule.id === 'string' && typeof value.rule.sheetId === 'string'
    && Array.isArray(value.rule.ranges) && value.rule.ranges.length > 0 && value.rule.ranges.every(isRange);
}

function isRuleRemoveMutation(value: unknown): value is { sheetId: string; ruleId: string } {
  return isRecord(value) && typeof value.sheetId === 'string' && typeof value.ruleId === 'string'
    && (value.ranges === undefined || (Array.isArray(value.ranges) && value.ranges.every(isRange)));
}

function isSheetRangesMutation(value: unknown): value is { sheetId: string; ranges: RangeRef[] } {
  return isRecord(value) && typeof value.sheetId === 'string' && Array.isArray(value.ranges) && value.ranges.every(isRange);
}

function isDataValidationAddMutation(value: unknown): value is AddDataValidationParams {
  return isRecord(value) && typeof value.sheetId === 'string' && isRecord(value.rule)
    && typeof value.rule.id === 'string' && typeof value.rule.sheetId === 'string'
    && Array.isArray(value.rule.ranges) && value.rule.ranges.length > 0 && value.rule.ranges.every(isRange);
}

function isBandedMutation(value: unknown): value is { sheetId: string; rule: BandedRule | null } {
  return isRecord(value) && typeof value.sheetId === 'string'
    && (value.rule === null || (isRecord(value.rule)
      && isRange(value.rule.range)
      && typeof value.rule.firstColor === 'string' && typeof value.rule.secondColor === 'string'));
}

function isNameSetMutation(value: unknown): value is { model: DefinedNameModel } {
  return isRecord(value) && isRecord(value.model)
    && typeof value.model.name === 'string' && value.model.name.trim().length > 0
    && typeof value.model.formula === 'string' && value.model.formula.trim().length > 0
    && (value.model.scope === 'workbook' || value.model.scope === 'sheet')
    && (value.model.scope === 'workbook' || typeof value.model.sheetId === 'string');
}

function isNameRemoveMutation(value: unknown): value is { name: string } {
  return isRecord(value) && typeof value.name === 'string' && value.name.trim().length > 0;
}

function ruleRanges(value: { rule: { ranges: RangeRef[] } }): RangeRef[] {
  return value.rule.ranges.map((range) => structuredClone(range));
}

function removeRuleRanges(value: { ranges?: RangeRef[] }): RangeRef[] {
  return value.ranges?.map((range) => structuredClone(range)) ?? [];
}

function restoreCell(
  workbook: WorkbookModel,
  item: MutationInfo<{ row: number; column: number; previous?: CellData }>,
): void {
  const sheet = workbook.getSheet(item.sheetId);
  const { row, column, previous } = item.params;
  if (previous) sheet.cells.set(row, column, previous);
  else sheet.cells.delete(row, column);
}

export function registerSheetCommands(runtime: CommandRuntime): void {
  registerEditingCommands(runtime);
  registerDataToolCommands(runtime);
  registerSheetTableCommands(runtime);
  registerOutlineCommands(runtime);
  registerHomeCommands(runtime);

  runtime.registry.registerMutation<RenameWorkbookParams>({
    id: 'workbook.renamed',
    handler: (item, context) => {
      if (!isRenameWorkbookMutation(item.params)) throw new Error('Invalid workbook.renamed mutation payload');
      context.workbook.name = item.params.name;
    },
    metadata: {
      schema: { name: 'RenameWorkbook', validate: isRenameWorkbookMutation },
      permission: { capability: 'workbook.rename', roles: ['owner', 'editor'] },
      affectedRanges: { resolve: () => [], mode: 'exact' },
      inverseIds: ['workbook.renamed'],
    },
  });
  runtime.registry.registerCommand<RenameWorkbookParams>({
    id: 'workbook.rename',
    execute: (params, context) => {
      const name = params.name.trim();
      if (!name) throw new Error('Workbook name is required');
      const previous = context.workbook.name;
      const affectedRanges: RangeRef[] = [];
      context.applyMutation({
        id: 'workbook.renamed',
        unitId: context.workbook.unitId,
        sheetId: context.workbook.primarySheetId,
        params: { name },
        affectedRanges,
        inverse: [{ id: 'workbook.renamed', unitId: context.workbook.unitId, sheetId: context.workbook.primarySheetId, params: { name: previous }, affectedRanges }],
        apply: () => { context.workbook.name = name; },
      });
      return { operationId: context.operationId, mutationCount: 1, affectedRanges };
    },
  });

  // 1. Sheet mutations & commands
  runtime.registry.registerMutation<AddSheetParams>({
    id: 'sheet.add',
    handler: (item, context) => {
      if (!isAddSheetMutation(item.params)) throw new Error('Invalid sheet.add mutation payload');
      const params = item.params;
      context.workbook.addSheet(params.id, params.name, params.rowCount, params.columnCount);
    },
    metadata: {
      schema: { name: 'AddSheet', validate: isAddSheetMutation },
      permission: { capability: 'sheet.structure.write', roles: ['owner', 'editor'] },
      affectedRanges: { resolve: () => [], mode: 'exact' },
      inverseIds: ['sheet.remove'],
    },
  });
  runtime.registry.registerMutation<{ id: string }>({
    id: 'sheet.remove',
    handler: (item, context) => {
      if (!isSheetIdMutation(item.params)) throw new Error('Invalid sheet.remove mutation payload');
      context.workbook.removeSheet(item.params.id);
    },
    metadata: {
      schema: { name: 'RemoveSheet', validate: isSheetIdMutation },
      permission: { capability: 'sheet.structure.write', roles: ['owner', 'editor'] },
      affectedRanges: { resolve: () => [], mode: 'exact' },
      inverseIds: ['sheet.restore'],
    },
  });
  runtime.registry.registerMutation<RenameSheetParams>({
    id: 'sheet.rename',
    handler: (item, context) => {
      if (!isRenameSheetMutation(item.params)) throw new Error('Invalid sheet.rename mutation payload');
      const params = item.params;
      context.workbook.getSheet(params.sheetId).name = params.name;
    },
    metadata: {
      schema: { name: 'RenameSheet', validate: isRenameSheetMutation },
      permission: { capability: 'sheet.structure.write', roles: ['owner', 'editor'] },
      affectedRanges: { resolve: () => [], mode: 'exact' },
      inverseIds: ['sheet.rename', 'cell.restore'],
    },
  });
  runtime.registry.registerMutation<{ sheet: import('@react-sheets/core-model').SheetSnapshot; index?: number }>({
    id: 'sheet.restore',
    handler: (item, context) => {
      if (!isSheetRestoreMutation(item.params)) throw new Error('Invalid sheet.restore mutation payload');
      context.workbook.restoreSheetSnapshot(item.params.sheet, item.params.index);
    },
    metadata: {
      schema: { name: 'RestoreSheet', validate: isSheetRestoreMutation },
      permission: { capability: 'sheet.structure.write', roles: ['owner', 'editor'] },
      affectedRanges: { resolve: () => [], mode: 'exact' },
      inverseIds: ['sheet.remove'],
    },
  });

  runtime.registry.registerCommand<{ id: string }>({
    id: 'sheet.remove',
    execute: (paramsInput, context) => {
      const params = paramsInput as { id: string };
      const workbook = context.workbook;
      if (workbook.getSheets().length <= 1) {
        throw new Error('A workbook must keep at least one worksheet');
      }
      const index = workbook.sheetOrder.indexOf(params.id);
      const snapshot = workbook.getSheetSnapshot(params.id);
      const affectedRanges: RangeRef[] = [];
      context.applyMutation({
        id: 'sheet.remove',
        unitId: workbook.unitId,
        sheetId: params.id,
        params,
        affectedRanges,
        inverse: [
          {
            id: 'sheet.restore',
            unitId: workbook.unitId,
            sheetId: params.id,
            params: { sheet: snapshot, index },
            affectedRanges,
          },
        ],
        apply: () => workbook.removeSheet(params.id),
      });
      return { operationId: context.operationId, mutationCount: 1, affectedRanges };
    },
  });

  runtime.registry.registerCommand<AddSheetParams>({
    id: 'sheet.add',
    execute: (params, context) => {
      const affectedRanges: RangeRef[] = [];
      context.applyMutation({
        id: 'sheet.add',
        unitId: context.workbook.unitId,
        sheetId: params.id,
        params,
        affectedRanges,
        inverse: [
          {
            id: 'sheet.remove',
            unitId: context.workbook.unitId,
            sheetId: params.id,
            params: { id: params.id },
            affectedRanges,
          },
        ],
        apply: () =>
          context.workbook.addSheet(params.id, params.name, params.rowCount, params.columnCount),
      });
      return { operationId: context.operationId, mutationCount: 1, affectedRanges };
    },
  });

  runtime.registry.registerCommand<RenameSheetParams>({
    id: 'sheet.rename',
    execute: (params, context) => {
      const sheet = context.workbook.getSheet(params.sheetId);
      const previousName = sheet.name;
      const affectedRanges: RangeRef[] = [];
      const formulaRewrites = previousName !== params.name
        ? rewriteFormulasForSheetRename(context.workbook, params.sheetId, previousName, params.name)
        : [];
      if (formulaRewrites.length > 0) {
        for (const item of formulaRewrites) {
          if (item.previous) context.workbook.getSheet(item.sheetId).cells.set(item.row, item.column, item.previous);
        }
      }
      context.applyMutation({
        id: 'sheet.rename',
        unitId: context.workbook.unitId,
        sheetId: params.sheetId,
        params,
        affectedRanges,
        inverse: [
          {
            id: 'sheet.rename',
            unitId: context.workbook.unitId,
            sheetId: params.sheetId,
            params: { sheetId: params.sheetId, name: previousName },
            affectedRanges,
          },
          ...formulaRewrites.map((item) => ({
            id: 'cell.restore' as const,
            unitId: context.workbook.unitId,
            sheetId: item.sheetId,
            params: { sheetId: item.sheetId, row: item.row, column: item.column, previous: item.previous },
            affectedRanges: cellRange({ sheetId: item.sheetId, row: item.row, column: item.column, value: item.previous ?? { value: null } }),
          })),
        ],
        apply: () => {
          sheet.name = params.name;
          if (previousName !== params.name) {
            rewriteFormulasForSheetRename(context.workbook, params.sheetId, previousName, params.name);
          }
        },
      });
      return { operationId: context.operationId, mutationCount: 1, affectedRanges };
    },
  });

  runtime.registry.registerMutation<WorkbookTableModel>({
    id: 'table.add',
    handler: (item, context) => {
      if (!isWorkbookTableMutation(item.params)) throw new Error('Invalid table.add mutation payload');
      context.workbook.addTable(item.params);
    },
    metadata: {
      schema: { name: 'WorkbookTableModel', validate: isWorkbookTableMutation },
      permission: { capability: 'workbook.table.write', roles: ['owner', 'editor'] },
      affectedRanges: { resolve: workbookTableRanges, mode: 'exact' },
      inverseIds: ['table.remove'],
    },
  });
  runtime.registry.registerMutation<{ tableId: string; range?: RangeRef }>({
    id: 'table.remove',
    handler: (item, context) => {
      if (!isTableRemoveMutation(item.params)) throw new Error('Invalid table.remove mutation payload');
      context.workbook.removeTable(item.params.tableId);
    },
    metadata: {
      schema: { name: 'TableRemove', validate: isTableRemoveMutation },
      permission: { capability: 'workbook.table.write', roles: ['owner', 'editor'] },
      affectedRanges: { resolve: tableRemoveRanges, mode: 'declared' },
      inverseIds: ['table.add'],
    },
  });
  runtime.registry.registerCommand<AddTableParams>({
    id: 'table.add',
    execute: (params, context) => {
      const affectedRanges: RangeRef[] = [];
      context.applyMutation({
        id: 'table.add',
        unitId: context.workbook.unitId,
        sheetId: params.sourceSheetId ?? context.workbook.primarySheetId,
        params: structuredClone(params),
        affectedRanges,
        inverse: [{ id: 'table.remove', unitId: context.workbook.unitId, sheetId: params.sourceSheetId ?? context.workbook.primarySheetId, params: { tableId: params.id, range: params.sourceRange }, affectedRanges }],
        apply: () => context.workbook.addTable(params),
      });
      return { operationId: context.operationId, mutationCount: 1, affectedRanges };
    },
  });
  runtime.registry.registerCommand<{ tableId: string; sheetId: string }>({
    id: 'table.remove',
    execute: (params, context) => {
      const previous = structuredClone(context.workbook.getTable(params.tableId));
      const affectedRanges: RangeRef[] = [];
      context.applyMutation({
        id: 'table.remove',
        unitId: context.workbook.unitId,
        sheetId: params.sheetId,
        params: { tableId: params.tableId, range: previous.sourceRange },
        affectedRanges,
        inverse: [{ id: 'table.add', unitId: context.workbook.unitId, sheetId: params.sheetId, params: previous, affectedRanges }],
        apply: () => context.workbook.removeTable(params.tableId),
      });
      return { operationId: context.operationId, mutationCount: 1, affectedRanges };
    },
  });

  // 2. Cell mutations & commands
  runtime.registry.registerMutation<SetCellValueParams>({
    id: 'cell.set',
    handler: (item, context) => {
      if (!isCellSetMutation(item.params)) throw new Error('Invalid cell.set mutation payload');
      const params = item.params;
      context.workbook.getSheet(params.sheetId).cells.set(params.row, params.column, { ...params.value });
    },
    metadata: {
      schema: { name: 'SetCellValue', validate: isCellSetMutation },
      permission: { capability: 'sheet.cell.write', roles: ['owner', 'editor'] },
      affectedRanges: { resolve: cellRange, mode: 'exact' },
      inverseIds: ['cell.restore'],
    },
  });
  runtime.registry.registerMutation<{ sheetId: string; row: number; column: number; previous?: CellData }>({
    id: 'cell.restore',
    handler: (item, context) => {
      if (!isCellRestoreMutation(item.params)) throw new Error('Invalid cell.restore mutation payload');
      restoreCell(context.workbook, item as MutationInfo<{ row: number; column: number; previous?: CellData }>);
    },
    metadata: {
      schema: { name: 'RestoreCell', validate: isCellRestoreMutation },
      permission: { capability: 'sheet.cell.write', roles: ['owner', 'editor'] },
      affectedRanges: { resolve: (params) => cellRange({ ...params, value: { value: null } }), mode: 'declared' },
      inverseIds: ['cell.set'],
    },
  });

  runtime.registry.registerCommand<SetCellValueParams>({
    id: 'sheet.cell.set',
    execute: (params, context) => {
      const sheet = context.workbook.getSheet(params.sheetId);
      const previous = sheet.cells.get(params.row, params.column);
      const affectedRanges = cellRange(params);
      context.applyMutation({
        id: 'cell.set',
        unitId: context.workbook.unitId,
        sheetId: params.sheetId,
        params,
        affectedRanges,
        inverse: [
          {
            id: 'cell.restore',
            unitId: context.workbook.unitId,
            sheetId: params.sheetId,
            params: { sheetId: params.sheetId, row: params.row, column: params.column, previous },
            affectedRanges,
          },
        ],
        apply: () => sheet.cells.set(params.row, params.column, { ...params.value }),
      });
      return { operationId: context.operationId, mutationCount: 1, affectedRanges };
    },
  });

  runtime.registry.registerCommand<CommitTextParams>({
    id: 'sheet.cell.commitText',
    execute: (params, context) => {
      if (typeof params.text !== 'string') throw new Error('Cell text must be a string');
      if (!Number.isSafeInteger(params.row) || params.row < 0 || !Number.isSafeInteger(params.column) || params.column < 0) {
        throw new Error('Cell row and column must be non-negative integers');
      }
      const sheet = context.workbook.getSheet(params.sheetId);
      for (const spill of sheet.spillRanges) {
        if (isSpillChild(spill, params.row, params.column)) {
          throw new Error('Cannot edit a dynamic-array spill child');
        }
      }

      const previous = sheet.cells.get(params.row, params.column);
      const next = buildCellFromText(params.text, previous, params.style);
      // A formula is validated by the FormulaEngine after commit and does not
      // have a scalar value to validate at this boundary. Scalar input must
      // satisfy the target rule before the cell.set mutation is opened.
      if (!next.formula) {
        const validation = validateDataInput(sheet, params.row, params.column, next.value);
        if (validation.blocking) throw new Error(validation.message ?? 'Cell value failed data validation');
      }
      const affectedRanges = cellRange({
        sheetId: params.sheetId,
        row: params.row,
        column: params.column,
        value: next,
      });
      context.applyMutation({
        id: 'cell.set',
        unitId: context.workbook.unitId,
        sheetId: params.sheetId,
        params: {
          sheetId: params.sheetId,
          row: params.row,
          column: params.column,
          value: structuredClone(next),
        },
        affectedRanges,
        inverse: [{
          id: 'cell.restore',
          unitId: context.workbook.unitId,
          sheetId: params.sheetId,
          params: {
            sheetId: params.sheetId,
            row: params.row,
            column: params.column,
            previous: previous ? structuredClone(previous) : undefined,
          },
          affectedRanges,
        }],
        apply: () => sheet.cells.set(params.row, params.column, structuredClone(next)),
      });
      return { operationId: context.operationId, mutationCount: 1, affectedRanges };
    },
  });

  // 3. Range set
  runtime.registry.registerMutation<SetRangeValuesParams>({
    id: 'range.set',
    handler: (item, context) => {
      if (!isSetRangeMutation(item.params)) throw new Error('Invalid range.set mutation payload');
      const params = item.params;
      const sheet = context.workbook.getSheet(params.sheetId);
      for (let rowOffset = 0; rowOffset < params.values.length; rowOffset += 1) {
        const rowValues = params.values[rowOffset] ?? [];
        for (let columnOffset = 0; columnOffset < rowValues.length; columnOffset += 1) {
          const value = rowValues[columnOffset];
          if (value) sheet.cells.set(params.startRow + rowOffset, params.startColumn + columnOffset, { ...value });
        }
      }
    },
    metadata: {
      schema: { name: 'SetRangeValues', validate: isSetRangeMutation },
      permission: { capability: 'sheet.cell.write', roles: ['owner', 'editor'] },
      affectedRanges: { resolve: setRangeAffectedRanges, mode: 'declared' },
      inverseIds: ['cell.restore'],
    },
  });

  runtime.registry.registerCommand<SetRangeValuesParams>({
    id: 'sheet.range.set',
    execute: (params, context) => {
      const sheet = context.workbook.getSheet(params.sheetId);
      const previous: Array<{ row: number; column: number; value?: CellData }> = [];
      const affectedRanges: RangeRef[] = [];
      for (let rowOffset = 0; rowOffset < params.values.length; rowOffset += 1) {
        const rowValues = params.values[rowOffset] ?? [];
        for (let columnOffset = 0; columnOffset < rowValues.length; columnOffset += 1) {
          const row = params.startRow + rowOffset;
          const column = params.startColumn + columnOffset;
          previous.push({ row, column, value: sheet.cells.get(row, column) });
          affectedRanges.push({
            sheetId: params.sheetId,
            startRow: row,
            endRow: row,
            startColumn: column,
            endColumn: column,
          });
        }
      }
      context.applyMutation({
        id: 'range.set',
        unitId: context.workbook.unitId,
        sheetId: params.sheetId,
        params,
        affectedRanges,
        inverse: previous.map((item) => ({
          id: 'cell.restore',
          unitId: context.workbook.unitId,
          sheetId: params.sheetId,
          params: { sheetId: params.sheetId, row: item.row, column: item.column, previous: item.value },
          affectedRanges: [
            {
              sheetId: params.sheetId,
              startRow: item.row,
              endRow: item.row,
              startColumn: item.column,
              endColumn: item.column,
            },
          ],
        })),
        apply: () => {
          for (let rowOffset = 0; rowOffset < params.values.length; rowOffset += 1) {
            const rowValues = params.values[rowOffset] ?? [];
            for (let columnOffset = 0; columnOffset < rowValues.length; columnOffset += 1) {
              const value = rowValues[columnOffset];
              if (value)
                sheet.cells.set(params.startRow + rowOffset, params.startColumn + columnOffset, {
                  ...value,
                });
            }
          }
        },
      });
      return { operationId: context.operationId, mutationCount: 1, affectedRanges };
    },
  });

  runtime.registry.registerMutation<ClearRangeParams>({
    id: 'range.clear',
    handler: (item, context) => {
    if (!isClearRangeMutation(item.params)) throw new Error('Invalid range.clear mutation payload');
    const params = item.params;
    const sheet = context.workbook.getSheet(params.sheetId);
    for (let row = params.range.startRow; row <= params.range.endRow; row += 1) {
      for (let column = params.range.startColumn; column <= params.range.endColumn; column += 1) {
        const current = sheet.cells.get(row, column);
        if (params.mode === 'notes') {
          if (current?.note || current?.comment) {
            const next = { ...current };
            delete next.note;
            delete next.comment;
            sheet.cells.set(row, column, next);
          }
          sheet.notes.delete(`${row}:${column}`);
          continue;
        }
        if (params.mode === 'hyperlinks') {
          sheet.hyperlinks.delete(`${row}:${column}`);
          if (current?.hyperlink !== undefined || current?.hyperlinkDetail !== undefined) {
            const next = { ...current };
            delete next.hyperlink;
            delete next.hyperlinkDetail;
            sheet.cells.set(row, column, next);
          }
          continue;
        }
        if (params.mode === undefined || params.mode === 'all') sheet.hyperlinks.delete(`${row}:${column}`);
        if (!current) continue;
        if (params.mode === 'formats') {
          const next = { ...current };
          delete next.style;
          delete next.styleId;
          delete next.numberFormat;
          delete next.displayValue;
          sheet.cells.set(row, column, next);
        } else if (params.mode === 'contents') {
          const next = { ...current };
          next.value = null;
          delete next.formula;
          delete next.displayValue;
          sheet.cells.set(row, column, next);
        } else {
          sheet.cells.delete(row, column);
        }
      }
    }
    if (params.mode === undefined || params.mode === 'all') {
      sheet.commentThreads.splice(0, sheet.commentThreads.length, ...sheet.commentThreads.filter((thread) =>
        thread.row < params.range.startRow
        || thread.row > params.range.endRow
        || thread.column < params.range.startColumn
        || thread.column > params.range.endColumn));
      for (let row = params.range.startRow; row <= params.range.endRow; row += 1) {
        for (let column = params.range.startColumn; column <= params.range.endColumn; column += 1) {
          sheet.notes.delete(`${row}:${column}`);
        }
      }
    }
    },
    metadata: {
      schema: { name: 'ClearRange', validate: isClearRangeMutation },
      permission: { capability: 'sheet.cell.write', roles: ['owner', 'editor'] },
      affectedRanges: { resolve: (params) => [structuredClone(params.range)], mode: 'exact' },
      inverseIds: ['range.clear.restore', 'cell.restore'],
    },
  });

  runtime.registry.registerMutation<ClearRangeRestoreParams>({
    id: 'range.clear.restore',
    handler: (item, context) => {
    if (!isClearRangeRestoreMutation(item.params)) throw new Error('Invalid range.clear.restore mutation payload');
    const params = item.params;
    const sheet = context.workbook.getSheet(params.sheetId);
    for (let row = params.range.startRow; row <= params.range.endRow; row += 1) {
      for (let column = params.range.startColumn; column <= params.range.endColumn; column += 1) {
        sheet.cells.delete(row, column);
        sheet.notes.delete(`${row}:${column}`);
        sheet.hyperlinks.delete(`${row}:${column}`);
      }
    }
    sheet.commentThreads.splice(0, sheet.commentThreads.length, ...sheet.commentThreads.filter((thread) =>
      thread.row < params.range.startRow
      || thread.row > params.range.endRow
      || thread.column < params.range.startColumn
      || thread.column > params.range.endColumn));
    for (const item of params.cells) {
      if (item.value) sheet.cells.set(item.row, item.column, structuredClone(item.value));
    }
    for (const item of params.notes) sheet.notes.set(`${item.row}:${item.column}`, structuredClone(item.note));
    for (const item of params.hyperlinks) sheet.hyperlinks.set(`${item.row}:${item.column}`, structuredClone(item.hyperlink));
    sheet.commentThreads.push(...structuredClone(params.comments));
    },
    metadata: {
      schema: { name: 'ClearRangeRestore', validate: isClearRangeRestoreMutation },
      permission: { capability: 'sheet.cell.write', roles: ['owner', 'editor'] },
      affectedRanges: { resolve: (params) => [structuredClone(params.range)], mode: 'exact' },
      inverseIds: ['range.clear'],
    },
  });

  runtime.registry.registerMutation<SetRangeStyleParams | { sheetId: string; ranges: RangeRef[]; numberFormat: string }>({
    id: 'style.set',
    handler: (item, context) => {
    if (!isStyleMutation(item.params)) throw new Error('Invalid style.set mutation payload');
    const params = item.params;
    const sheet = context.workbook.getSheet(params.sheetId);
    const ranges = 'range' in params ? [params.range] : params.ranges;
    const style = 'numberFormat' in params && !('style' in params)
      ? { numberFormat: params.numberFormat }
      : (params as SetRangeStyleParams).style;
    for (const range of ranges) {
      for (let row = range.startRow; row <= range.endRow; row += 1) {
        for (let column = range.startColumn; column <= range.endColumn; column += 1) {
          const current = sheet.cells.get(row, column) ?? { value: null as CellData['value'] };
          const next = { ...current, style: { ...(current.style ?? {}), ...style } };
          if (style.numberFormat !== undefined) next.numberFormat = style.numberFormat;
          sheet.cells.set(row, column, next);
        }
      }
    }
    },
    metadata: {
      schema: { name: 'StyleSet', validate: isStyleMutation },
      permission: { capability: 'sheet.format.write', roles: ['owner', 'editor'] },
      affectedRanges: { resolve: styleAffectedRanges, mode: 'declared' },
      inverseIds: ['cell.restore'],
    },
  });

  // 4. Range Clear
  runtime.registry.registerCommand<ClearRangeParams>({
    id: 'sheet.range.clear',
    execute: (params, context) => {
      const sheet = context.workbook.getSheet(params.sheetId);
      const previous: Array<{ row: number; column: number; value?: CellData }> = [];
      const notes: ClearRangeRestoreParams['notes'] = [];
      const hyperlinks: ClearRangeRestoreParams['hyperlinks'] = [];
      const comments: ClearRangeRestoreParams['comments'] = [];
      const range = {
        ...params.range,
        startRow: Math.min(params.range.startRow, params.range.endRow),
        endRow: Math.max(params.range.startRow, params.range.endRow),
        startColumn: Math.min(params.range.startColumn, params.range.endColumn),
        endColumn: Math.max(params.range.startColumn, params.range.endColumn),
      };
      const affectedRanges: RangeRef[] = [range];

      for (let r = range.startRow; r <= range.endRow; r++) {
        for (let c = range.startColumn; c <= range.endColumn; c++) {
          previous.push({ row: r, column: c, value: structuredClone(sheet.cells.get(r, c)) });
          const note = sheet.notes.get(`${r}:${c}`);
          if (note) notes.push({ row: r, column: c, note: structuredClone(note) });
          const hyperlink = sheet.hyperlinks.get(`${r}:${c}`);
          if (hyperlink) hyperlinks.push({ row: r, column: c, hyperlink: structuredClone(hyperlink) });
        }
      }
      comments.push(...structuredClone(sheet.commentThreads.filter((thread) =>
        thread.row >= range.startRow
        && thread.row <= range.endRow
        && thread.column >= range.startColumn
        && thread.column <= range.endColumn)));

      context.applyMutation({
        id: 'range.clear',
        unitId: context.workbook.unitId,
        sheetId: params.sheetId,
        params: { ...params, range },
        affectedRanges,
        inverse: [{
          id: 'range.clear.restore',
          unitId: context.workbook.unitId,
          sheetId: params.sheetId,
          params: { sheetId: params.sheetId, range, cells: previous, notes, hyperlinks, comments },
          affectedRanges,
        }],
        apply: () => runtime.registry.getMutation('range.clear')({ id: 'range.clear', unitId: context.workbook.unitId, sheetId: params.sheetId, params: { ...params, range }, affectedRanges }, context),
      });
      return { operationId: context.operationId, mutationCount: 1, affectedRanges };
    },
  });

  // 5. Range Style batch application
  runtime.registry.registerCommand<SetRangeStyleParams>({
    id: 'sheet.style.set',
    execute: (params, context) => {
      const sheet = context.workbook.getSheet(params.sheetId);
      const previous: Array<{ row: number; column: number; value?: CellData }> = [];
      const affectedRanges: RangeRef[] = [params.range];

      for (let r = params.range.startRow; r <= params.range.endRow; r++) {
        for (let c = params.range.startColumn; c <= params.range.endColumn; c++) {
          const cell = sheet.cells.get(r, c);
          previous.push({ row: r, column: c, value: cell ? structuredClone(cell) : undefined });
        }
      }

      context.applyMutation({
        id: 'style.set',
        unitId: context.workbook.unitId,
        sheetId: params.sheetId,
        params,
        affectedRanges,
        inverse: previous.map((item) => ({
          id: 'cell.restore',
          unitId: context.workbook.unitId,
          sheetId: params.sheetId,
          params: { sheetId: params.sheetId, row: item.row, column: item.column, previous: item.value },
          affectedRanges: [
            {
              sheetId: params.sheetId,
              startRow: item.row,
              endRow: item.row,
              startColumn: item.column,
              endColumn: item.column,
            },
          ],
        })),
        apply: () => runtime.registry.getMutation('style.set')({ id: 'style.set', unitId: context.workbook.unitId, sheetId: params.sheetId, params, affectedRanges }, context),
      });
      return { operationId: context.operationId, mutationCount: 1, affectedRanges };
    },
  });

  // 6. Merge commands & mutations
  runtime.registry.registerMutation<SetMergeParams>({
    id: 'merge.set',
    handler: (item, context) => {
      if (!isRecord(item.params) || typeof item.params.sheetId !== 'string' || !isRange(item.params.range)) throw new Error('Invalid merge.set mutation payload');
      const params = item.params as SetMergeParams;
      const sheet = context.workbook.getSheet(params.sheetId);
      sheet.merges.push({ range: params.range, anchor: { row: params.range.startRow, column: params.range.startColumn } });
    },
    metadata: {
      schema: { name: 'SetMerge', validate: (value: unknown) => isRecord(value) && typeof value.sheetId === 'string' && isRange(value.range) },
      permission: { capability: 'sheet.merge.write', roles: ['owner', 'editor'] },
      affectedRanges: { resolve: (params) => [structuredClone(params.range)], mode: 'exact' },
      inverseIds: ['merge.remove'],
    },
  });
  runtime.registry.registerMutation<RemoveMergeParams>({
    id: 'merge.remove',
    handler: (item, context) => {
      if (!isRecord(item.params) || typeof item.params.sheetId !== 'string' || !isRange(item.params.range)) throw new Error('Invalid merge.remove mutation payload');
      const params = item.params as RemoveMergeParams;
      const sheet = context.workbook.getSheet(params.sheetId);
      const idx = sheet.merges.findIndex((m) => m.range.startRow === params.range.startRow && m.range.startColumn === params.range.startColumn);
      if (idx >= 0) sheet.merges.splice(idx, 1);
    },
    metadata: {
      schema: { name: 'RemoveMerge', validate: (value: unknown) => isRecord(value) && typeof value.sheetId === 'string' && isRange(value.range) },
      permission: { capability: 'sheet.merge.write', roles: ['owner', 'editor'] },
      affectedRanges: { resolve: (params) => [structuredClone(params.range)], mode: 'exact' },
      inverseIds: ['merge.set'],
    },
  });

  runtime.registry.registerCommand<SetMergeParams>({
    id: 'sheet.merge.set',
    execute: (params, context) => {
      const sheet = context.workbook.getSheet(params.sheetId);
      const span: MergeSpan = {
        range: params.range,
        anchor: { row: params.range.startRow, column: params.range.startColumn },
      };
      const affectedRanges = [params.range];

      context.applyMutation({
        id: 'merge.set',
        unitId: context.workbook.unitId,
        sheetId: params.sheetId,
        params,
        affectedRanges,
        inverse: [
          {
            id: 'merge.remove',
            unitId: context.workbook.unitId,
            sheetId: params.sheetId,
            params: { sheetId: params.sheetId, range: params.range },
            affectedRanges,
          },
        ],
        apply: () => {
          sheet.merges.push(span);
        },
      });
      return { operationId: context.operationId, mutationCount: 1, affectedRanges };
    },
  });

  runtime.registry.registerCommand<RemoveMergeParams>({
    id: 'sheet.merge.remove',
    execute: (params, context) => {
      const sheet = context.workbook.getSheet(params.sheetId);
      const idx = sheet.merges.findIndex(
        (m) =>
          m.range.startRow === params.range.startRow &&
          m.range.startColumn === params.range.startColumn,
      );
      const affectedRanges = [params.range];
      if (idx < 0)
        return { operationId: context.operationId, mutationCount: 0, affectedRanges: [] };
      const previousSpan = sheet.merges[idx]!;

      context.applyMutation({
        id: 'merge.remove',
        unitId: context.workbook.unitId,
        sheetId: params.sheetId,
        params,
        affectedRanges,
        inverse: [
          {
            id: 'merge.set',
            unitId: context.workbook.unitId,
            sheetId: params.sheetId,
            params: { sheetId: params.sheetId, range: previousSpan.range },
            affectedRanges,
          },
        ],
        apply: () => {
          sheet.merges.splice(idx, 1);
        },
      });
      return { operationId: context.operationId, mutationCount: 1, affectedRanges };
    },
  });

  // 7. Freeze commands
  runtime.registry.registerMutation<SetFreezeParams>({
    id: 'freeze.set',
    handler: (item, context) => {
      if (!isRecord(item.params) || typeof item.params.sheetId !== 'string' || !isWorksheetPane(item.params.pane)) throw new Error('Invalid freeze.set mutation payload');
      const params = item.params as SetFreezeParams;
      context.workbook.getSheet(params.sheetId).pane = { ...params.pane };
    },
    metadata: {
      schema: { name: 'SetPane', validate: (value: unknown) => isRecord(value) && typeof value.sheetId === 'string' && isWorksheetPane(value.pane) },
      permission: { capability: 'sheet.view.write', roles: ['owner', 'editor'] },
      affectedRanges: { resolve: () => [], mode: 'exact' },
      inverseIds: ['freeze.set'],
    },
  });

  runtime.registry.registerCommand<SetFreezeParams>({
    id: 'sheet.freeze.set',
    execute: (params, context) => {
      const sheet = context.workbook.getSheet(params.sheetId);
      const previousPane = { ...sheet.pane };
      const affectedRanges: RangeRef[] = [];

      context.applyMutation({
        id: 'freeze.set',
        unitId: context.workbook.unitId,
        sheetId: params.sheetId,
        params,
        affectedRanges,
        inverse: [
          {
            id: 'freeze.set',
            unitId: context.workbook.unitId,
            sheetId: params.sheetId,
            params: { sheetId: params.sheetId, pane: previousPane },
            affectedRanges,
          },
        ],
        apply: () => {
          sheet.pane = { ...params.pane };
        },
      });
      return { operationId: context.operationId, mutationCount: 1, affectedRanges };
    },
  });

  // 8. Row and Column resizing
  runtime.registry.registerMutation<ResizeRowParams>({
    id: 'row.resize',
    handler: (item, context) => {
      if (!isRecord(item.params) || typeof item.params.sheetId !== 'string' || !Number.isInteger(item.params.row) || typeof item.params.heightPx !== 'number' || item.params.heightPx <= 0) throw new Error('Invalid row.resize mutation payload');
      const params = item.params as ResizeRowParams;
      context.workbook.getSheet(params.sheetId).rowHeightsPx[params.row] = params.heightPx;
    },
    metadata: {
      schema: { name: 'ResizeRowPx', validate: (value: unknown) => isRecord(value) && typeof value.sheetId === 'string' && Number.isInteger(value.row) && typeof value.heightPx === 'number' && Number(value.row) >= 0 && Number(value.heightPx) > 0 },
      permission: { capability: 'sheet.dimension.write', roles: ['owner', 'editor'] },
      affectedRanges: { resolve: rowAffectedRange, mode: 'declared' },
      inverseIds: ['row.resize'],
    },
  });
  runtime.registry.registerMutation<ResizeColumnParams>({
    id: 'column.resize',
    handler: (item, context) => {
      if (!isRecord(item.params) || typeof item.params.sheetId !== 'string' || !Number.isInteger(item.params.column) || typeof item.params.widthPx !== 'number' || item.params.widthPx <= 0) throw new Error('Invalid column.resize mutation payload');
      const params = item.params as ResizeColumnParams;
      context.workbook.getSheet(params.sheetId).columnWidthsPx[params.column] = params.widthPx;
    },
    metadata: {
      schema: { name: 'ResizeColumnPx', validate: (value: unknown) => isRecord(value) && typeof value.sheetId === 'string' && Number.isInteger(value.column) && typeof value.widthPx === 'number' && Number(value.column) >= 0 && Number(value.widthPx) > 0 },
      permission: { capability: 'sheet.dimension.write', roles: ['owner', 'editor'] },
      affectedRanges: { resolve: columnAffectedRange, mode: 'declared' },
      inverseIds: ['column.resize'],
    },
  });
  runtime.registry.registerMutation<{ sheetId: string; widthPx: number }>({
    id: 'column.defaultWidth.resize',
    handler: (item, context) => {
      if (!isRecord(item.params) || typeof item.params.sheetId !== 'string' || typeof item.params.widthPx !== 'number' || !Number.isFinite(item.params.widthPx) || item.params.widthPx <= 0) throw new Error('Invalid column.defaultWidth.resize mutation payload');
      context.workbook.getSheet(item.params.sheetId).defaultColumnWidthPx = item.params.widthPx;
    },
    metadata: {
      schema: { name: 'ResizeDefaultColumnWidthPx', validate: (value: unknown) => isRecord(value) && typeof value.sheetId === 'string' && typeof value.widthPx === 'number' && Number.isFinite(value.widthPx) && value.widthPx > 0 },
      permission: { capability: 'sheet.dimension.write', roles: ['owner', 'editor'] },
      affectedRanges: { resolve: sheetScopeRange, mode: 'declared' },
      inverseIds: ['column.defaultWidth.resize'],
    },
  });

  runtime.registry.registerCommand<{ sheetId: string; rows?: Array<Omit<ResizeRowParams, 'sheetId'>>; columns?: Array<Omit<ResizeColumnParams, 'sheetId'>> }>({
    id: 'sheet.dimensions.apply',
    execute: (params, context) => {
      const sheet = context.workbook.getSheet(params.sheetId);
      const affectedRanges: RangeRef[] = [];
      let mutationCount = 0;
      for (const row of params.rows ?? []) {
        if (!Number.isSafeInteger(row.row) || row.row < 0 || !Number.isFinite(row.heightPx) || row.heightPx <= 0) throw new Error('Invalid row pixel size');
        const mutationParams: ResizeRowParams = { sheetId: params.sheetId, ...row };
        const previousHeightPx = sheet.rowHeightsPx[row.row] ?? sheet.defaultRowHeightPx;
        context.applyMutation({
          id: 'row.resize', unitId: context.workbook.unitId, sheetId: params.sheetId, params: mutationParams, affectedRanges,
          inverse: [{ id: 'row.resize', unitId: context.workbook.unitId, sheetId: params.sheetId, params: { sheetId: params.sheetId, row: row.row, heightPx: previousHeightPx }, affectedRanges }],
          apply: () => { sheet.rowHeightsPx[row.row] = row.heightPx; },
        });
        mutationCount += 1;
      }
      for (const column of params.columns ?? []) {
        if (!Number.isSafeInteger(column.column) || column.column < 0 || !Number.isFinite(column.widthPx) || column.widthPx <= 0) throw new Error('Invalid column pixel size');
        const mutationParams: ResizeColumnParams = { sheetId: params.sheetId, ...column };
        const previousWidthPx = sheet.columnWidthsPx[column.column] ?? sheet.defaultColumnWidthPx;
        context.applyMutation({
            id: 'column.resize',
            unitId: context.workbook.unitId,
            sheetId: params.sheetId,
            params: mutationParams,
            affectedRanges,
            inverse: [{ id: 'column.resize', unitId: context.workbook.unitId, sheetId: params.sheetId, params: { sheetId: params.sheetId, column: column.column, widthPx: previousWidthPx }, affectedRanges }],
            apply: () => { sheet.columnWidthsPx[column.column] = column.widthPx; },
        });
        mutationCount += 1;
      }
      return { operationId: context.operationId, mutationCount, affectedRanges };
    },
  });
  runtime.registry.registerCommand<{ sheetId: string; widthPx: number }>({
    id: 'sheet.column.defaultWidth.set',
    execute: (params, context) => {
      if (!Number.isFinite(params.widthPx) || params.widthPx <= 0) throw new Error('Default column width must be positive pixels');
      const sheet = context.workbook.getSheet(params.sheetId);
      const previousWidthPx = sheet.defaultColumnWidthPx;
      const affectedRanges = sheetScopeRange(params);
      context.applyMutation({
        id: 'column.defaultWidth.resize', unitId: context.workbook.unitId, sheetId: params.sheetId, params, affectedRanges,
        inverse: [{ id: 'column.defaultWidth.resize', unitId: context.workbook.unitId, sheetId: params.sheetId, params: { sheetId: params.sheetId, widthPx: previousWidthPx }, affectedRanges }],
        apply: () => { sheet.defaultColumnWidthPx = params.widthPx; },
      });
      return { operationId: context.operationId, mutationCount: 1, affectedRanges };
    },
  });

  // 9. Range Sorting
  runtime.registry.registerCommand<SortRangeParams>({
    id: 'sheet.sort',
    execute: (params, context) => runtime.execute('data.sort.rows', {
      sheetId: params.sheetId,
      range: params.range,
      criteria: [{ column: params.sortColumn, ascending: params.ascending }],
      hasHeader: params.hasHeader,
    }),
  });

  // 10. AutoFill Sequence / Formula shift
  runtime.registry.registerCommand<AutoFillParams>({
    id: 'sheet.autofill',
    execute: (params, context) => {
      const sheet = context.workbook.getSheet(params.sheetId);
      const { sourceRange, targetRange } = params;
      const values: CellData[][] = [];

      const sourceRowCount = sourceRange.endRow - sourceRange.startRow + 1;
      const sourceColCount = sourceRange.endColumn - sourceRange.startColumn + 1;

      for (let r = targetRange.startRow; r <= targetRange.endRow; r++) {
        const rowList: CellData[] = [];
        const sourceR = sourceRange.startRow + ((r - targetRange.startRow) % sourceRowCount);
        const rowOffset = r - sourceR;

        for (let c = targetRange.startColumn; c <= targetRange.endColumn; c++) {
          const sourceC =
            sourceRange.startColumn + ((c - targetRange.startColumn) % sourceColCount);
          const colOffset = c - sourceC;
          const sourceCell = sheet.cells.get(sourceR, sourceC);

          if (!sourceCell) {
            rowList.push({ value: null });
            continue;
          }

          if (sourceCell.formula) {
            const shifted = shiftFormula(sourceCell.formula, rowOffset, colOffset);
            rowList.push({ ...sourceCell, formula: shifted, value: null });
          } else if (typeof sourceCell.value === 'number') {
            // Sequence extension
            const step = rowOffset !== 0 ? Math.floor((r - sourceRange.endRow) / sourceRowCount) + 1 : 0;
            rowList.push({ ...sourceCell, value: sourceCell.value + step });
          } else {
            rowList.push(structuredClone(sourceCell));
          }
        }
        values.push(rowList);
      }

      return runtime.execute('sheet.range.set', {
        sheetId: params.sheetId,
        startRow: targetRange.startRow,
        startColumn: targetRange.startColumn,
        values,
      });
    },
  });

  // 11. 结构操作:行/列插入与删除（统一走 StructuralTransform）
  runtime.registry.registerMutation<{ sheetId: string; at: number; count: number }>({
    id: 'rows.inserted',
    handler: (item, context) => {
      if (!isSheetAtCountMutation(item.params)) throw new Error('Invalid rows.inserted mutation payload');
      const params = item.params;
      applyStructuralTransform(context.workbook, { kind: 'insert-rows', sheetId: params.sheetId, at: params.at, count: params.count });
    },
    metadata: {
      schema: { name: 'RowsInserted', validate: isSheetAtCountMutation },
      permission: { capability: 'sheet.structure.write', roles: ['owner', 'editor'] },
      affectedRanges: { resolve: structuralAffectedRanges, mode: 'declared' },
      inverseIds: ['rows.deleted'],
    },
  });
  runtime.registry.registerMutation<{ sheetId: string; at: number; count: number }>({
    id: 'rows.deleted',
    handler: (item, context) => {
      if (!isSheetAtCountMutation(item.params)) throw new Error('Invalid rows.deleted mutation payload');
      const params = item.params;
      applyStructuralTransform(context.workbook, { kind: 'delete-rows', sheetId: params.sheetId, at: params.at, count: params.count });
    },
    metadata: {
      schema: { name: 'RowsDeleted', validate: isSheetAtCountMutation },
      permission: { capability: 'sheet.structure.write', roles: ['owner', 'editor'] },
      affectedRanges: { resolve: structuralAffectedRanges, mode: 'declared' },
      inverseIds: ['rows.inserted', 'cell.restore'],
    },
  });
  runtime.registry.registerMutation<{ sheetId: string; at: number; count: number }>({
    id: 'columns.inserted',
    handler: (item, context) => {
      if (!isSheetAtCountMutation(item.params)) throw new Error('Invalid columns.inserted mutation payload');
      const params = item.params;
      applyStructuralTransform(context.workbook, { kind: 'insert-columns', sheetId: params.sheetId, at: params.at, count: params.count });
    },
    metadata: {
      schema: { name: 'ColumnsInserted', validate: isSheetAtCountMutation },
      permission: { capability: 'sheet.structure.write', roles: ['owner', 'editor'] },
      affectedRanges: { resolve: columnStructuralAffectedRanges, mode: 'declared' },
      inverseIds: ['columns.deleted'],
    },
  });
  runtime.registry.registerMutation<{ sheetId: string; at: number; count: number }>({
    id: 'columns.deleted',
    handler: (item, context) => {
      if (!isSheetAtCountMutation(item.params)) throw new Error('Invalid columns.deleted mutation payload');
      const params = item.params;
      applyStructuralTransform(context.workbook, { kind: 'delete-columns', sheetId: params.sheetId, at: params.at, count: params.count });
    },
    metadata: {
      schema: { name: 'ColumnsDeleted', validate: isSheetAtCountMutation },
      permission: { capability: 'sheet.structure.write', roles: ['owner', 'editor'] },
      affectedRanges: { resolve: columnStructuralAffectedRanges, mode: 'declared' },
      inverseIds: ['columns.inserted', 'cell.restore'],
    },
  });
  runtime.registry.registerMutation<{ sheetId: string; index: number }>({
    id: 'row.hidden',
    handler: (item, context) => {
      if (!isSheetIndexMutation(item.params)) throw new Error('Invalid row.hidden mutation payload');
      const params = item.params;
      context.workbook.getSheet(params.sheetId).hiddenRows.add(params.index);
    },
    metadata: {
      schema: { name: 'RowHidden', validate: isSheetIndexMutation },
      permission: { capability: 'sheet.visibility.write', roles: ['owner', 'editor'] },
      affectedRanges: { resolve: (params) => rowAffectedRange({ sheetId: params.sheetId, row: params.index }), mode: 'declared' },
      inverseIds: ['row.unhidden'],
    },
  });
  runtime.registry.registerMutation<ColumnsVisibilityParams>({
    id: 'columns.visibility',
    handler: (item, context) => {
      if (!isColumnVisibilityMutation(item.params)) throw new Error('Invalid columns.visibility mutation payload');
      const hiddenColumns = context.workbook.getSheet(item.params.sheetId).hiddenColumns;
      for (const state of item.params.states) {
        if (state.hidden) hiddenColumns.add(state.column);
        else hiddenColumns.delete(state.column);
      }
    },
    metadata: {
      schema: { name: 'ColumnsVisibility', validate: isColumnVisibilityMutation },
      permission: { capability: 'sheet.visibility.write', roles: ['owner', 'editor'] },
      affectedRanges: { resolve: (params) => params.states.map((state) => columnAffectedRange({ sheetId: params.sheetId, column: state.column })[0]!), mode: 'declared' },
      inverseIds: ['columns.visibility'],
    },
  });
  runtime.registry.registerMutation<{ sheetId: string; index: number }>({
    id: 'row.unhidden',
    handler: (item, context) => {
      if (!isSheetIndexMutation(item.params)) throw new Error('Invalid row.unhidden mutation payload');
      const params = item.params;
      context.workbook.getSheet(params.sheetId).hiddenRows.delete(params.index);
    },
    metadata: {
      schema: { name: 'RowUnhidden', validate: isSheetIndexMutation },
      permission: { capability: 'sheet.visibility.write', roles: ['owner', 'editor'] },
      affectedRanges: { resolve: (params) => rowAffectedRange({ sheetId: params.sheetId, row: params.index }), mode: 'declared' },
      inverseIds: ['row.hidden'],
    },
  });
  runtime.registry.registerMutation<{ sheetId: string }>({
    id: 'rows.unhidden.all',
    handler: (item, context) => {
      if (!isRecord(item.params) || typeof item.params.sheetId !== 'string') throw new Error('Invalid rows.unhidden.all mutation payload');
      context.workbook.getSheet(item.params.sheetId).hiddenRows.clear();
    },
    metadata: {
      schema: { name: 'RowsUnhiddenAll', validate: (value: unknown) => isRecord(value) && typeof value.sheetId === 'string' },
      permission: { capability: 'sheet.visibility.write', roles: ['owner', 'editor'] },
      affectedRanges: { resolve: sheetScopeRange, mode: 'declared' },
      inverseIds: ['rows.hidden.restore'],
    },
  });
  runtime.registry.registerMutation<{ sheetId: string; indices: number[] }>({
    id: 'rows.hidden.restore',
    handler: (item, context) => {
      if (!isSheetIndicesMutation(item.params)) throw new Error('Invalid rows.hidden.restore mutation payload');
      const params = item.params;
      const hiddenRows = context.workbook.getSheet(params.sheetId).hiddenRows;
      hiddenRows.clear();
      for (const index of params.indices) hiddenRows.add(index);
    },
    metadata: {
      schema: { name: 'RowsHiddenRestore', validate: isSheetIndicesMutation },
      permission: { capability: 'sheet.visibility.write', roles: ['owner', 'editor'] },
      affectedRanges: { resolve: sheetScopeRange, mode: 'declared' },
      inverseIds: ['rows.unhidden.all'],
    },
  });
  runtime.registry.registerMutation<{ sheetId: string; index: number }>({
    id: 'column.hidden',
    handler: (item, context) => {
      if (!isSheetIndexMutation(item.params)) throw new Error('Invalid column.hidden mutation payload');
      const params = item.params;
      context.workbook.getSheet(params.sheetId).hiddenColumns.add(params.index);
    },
    metadata: {
      schema: { name: 'ColumnHidden', validate: isSheetIndexMutation },
      permission: { capability: 'sheet.visibility.write', roles: ['owner', 'editor'] },
      affectedRanges: { resolve: (params) => columnAffectedRange({ sheetId: params.sheetId, column: params.index }), mode: 'declared' },
      inverseIds: ['column.unhidden'],
    },
  });
  runtime.registry.registerMutation<{ sheetId: string; index: number }>({
    id: 'column.unhidden',
    handler: (item, context) => {
      if (!isSheetIndexMutation(item.params)) throw new Error('Invalid column.unhidden mutation payload');
      const params = item.params;
      context.workbook.getSheet(params.sheetId).hiddenColumns.delete(params.index);
    },
    metadata: {
      schema: { name: 'ColumnUnhidden', validate: isSheetIndexMutation },
      permission: { capability: 'sheet.visibility.write', roles: ['owner', 'editor'] },
      affectedRanges: { resolve: (params) => columnAffectedRange({ sheetId: params.sheetId, column: params.index }), mode: 'declared' },
      inverseIds: ['column.hidden'],
    },
  });
  runtime.registry.registerMutation<{ sheetId: string }>({
    id: 'columns.unhidden.all',
    handler: (item, context) => {
      if (!isRecord(item.params) || typeof item.params.sheetId !== 'string') throw new Error('Invalid columns.unhidden.all mutation payload');
      context.workbook.getSheet(item.params.sheetId).hiddenColumns.clear();
    },
    metadata: {
      schema: { name: 'ColumnsUnhiddenAll', validate: (value: unknown) => isRecord(value) && typeof value.sheetId === 'string' },
      permission: { capability: 'sheet.visibility.write', roles: ['owner', 'editor'] },
      affectedRanges: { resolve: sheetScopeRange, mode: 'declared' },
      inverseIds: ['columns.hidden.restore'],
    },
  });
  runtime.registry.registerMutation<{ sheetId: string; indices: number[] }>({
    id: 'columns.hidden.restore',
    handler: (item, context) => {
      if (!isSheetIndicesMutation(item.params)) throw new Error('Invalid columns.hidden.restore mutation payload');
      const params = item.params;
      const hiddenColumns = context.workbook.getSheet(params.sheetId).hiddenColumns;
      hiddenColumns.clear();
      for (const index of params.indices) hiddenColumns.add(index);
    },
    metadata: {
      schema: { name: 'ColumnsHiddenRestore', validate: isSheetIndicesMutation },
      permission: { capability: 'sheet.visibility.write', roles: ['owner', 'editor'] },
      affectedRanges: { resolve: sheetScopeRange, mode: 'declared' },
      inverseIds: ['columns.unhidden.all'],
    },
  });

  runtime.registry.registerCommand<{ sheetId: string; index: number }>({
    id: 'sheet.row.hide',
    execute: (params, context) => {
      const hiddenRows = context.workbook.getSheet(params.sheetId).hiddenRows;
      if (hiddenRows.has(params.index)) return { operationId: context.operationId, mutationCount: 0, affectedRanges: [] };
      const affectedRanges: RangeRef[] = [{ sheetId: params.sheetId, startRow: params.index, endRow: params.index, startColumn: 0, endColumn: Math.max(0, context.workbook.getSheet(params.sheetId).columnCount - 1) }];
      context.applyMutation({
        id: 'row.hidden',
        unitId: context.workbook.unitId,
        sheetId: params.sheetId,
        params,
        affectedRanges,
        inverse: [{ id: 'row.unhidden', unitId: context.workbook.unitId, sheetId: params.sheetId, params, affectedRanges }],
        apply: () => hiddenRows.add(params.index),
      });
      return { operationId: context.operationId, mutationCount: 1, affectedRanges };
    },
  });
  runtime.registry.registerCommand<{ sheetId: string; columns: number[]; hidden: boolean }>({
    id: 'sheet.columns.visibility.set',
    execute: (params, context) => {
      const sheet = context.workbook.getSheet(params.sheetId);
      const columns = [...new Set(params.columns)].filter((column) => Number.isSafeInteger(column) && column >= 0 && column < sheet.columnCount);
      const changed = columns.filter((column) => sheet.hiddenColumns.has(column) !== params.hidden);
      if (!changed.length) return { operationId: context.operationId, mutationCount: 0, affectedRanges: [] };
      const states = changed.map((column) => ({ column, hidden: params.hidden }));
      const inverseStates = changed.map((column) => ({ column, hidden: !params.hidden }));
      const affectedRanges = states.map((state) => columnAffectedRange({ sheetId: params.sheetId, column: state.column })[0]!);
      context.applyMutation({
        id: 'columns.visibility', unitId: context.workbook.unitId, sheetId: params.sheetId, params: { sheetId: params.sheetId, states }, affectedRanges,
        inverse: [{ id: 'columns.visibility', unitId: context.workbook.unitId, sheetId: params.sheetId, params: { sheetId: params.sheetId, states: inverseStates }, affectedRanges }],
        apply: () => { for (const state of states) { if (state.hidden) sheet.hiddenColumns.add(state.column); else sheet.hiddenColumns.delete(state.column); } },
      });
      return { operationId: context.operationId, mutationCount: 1, affectedRanges };
    },
  });

  runtime.registry.registerCommand<{ sheetId: string; index: number }>({
    id: 'sheet.column.hide',
    execute: (params, context) => {
      const hiddenColumns = context.workbook.getSheet(params.sheetId).hiddenColumns;
      if (hiddenColumns.has(params.index)) return { operationId: context.operationId, mutationCount: 0, affectedRanges: [] };
      const affectedRanges: RangeRef[] = [{ sheetId: params.sheetId, startRow: 0, endRow: Math.max(0, context.workbook.getSheet(params.sheetId).rowCount - 1), startColumn: params.index, endColumn: params.index }];
      context.applyMutation({
        id: 'column.hidden',
        unitId: context.workbook.unitId,
        sheetId: params.sheetId,
        params,
        affectedRanges,
        inverse: [{ id: 'column.unhidden', unitId: context.workbook.unitId, sheetId: params.sheetId, params, affectedRanges }],
        apply: () => hiddenColumns.add(params.index),
      });
      return { operationId: context.operationId, mutationCount: 1, affectedRanges };
    },
  });

  runtime.registry.registerCommand<{ sheetId: string }>({
    id: 'sheet.rows.unhide.all',
    execute: (params, context) => {
      const sheet = context.workbook.getSheet(params.sheetId);
      const previous = [...sheet.hiddenRows].sort((left, right) => left - right);
      if (previous.length === 0) return { operationId: context.operationId, mutationCount: 0, affectedRanges: [] };
      const affectedRanges: RangeRef[] = [{ sheetId: params.sheetId, startRow: 0, endRow: Math.max(0, sheet.rowCount - 1), startColumn: 0, endColumn: Math.max(0, sheet.columnCount - 1) }];
      context.applyMutation({
        id: 'rows.unhidden.all',
        unitId: context.workbook.unitId,
        sheetId: params.sheetId,
        params,
        affectedRanges,
        inverse: [{ id: 'rows.hidden.restore', unitId: context.workbook.unitId, sheetId: params.sheetId, params: { sheetId: params.sheetId, indices: previous }, affectedRanges }],
        apply: () => sheet.hiddenRows.clear(),
      });
      return { operationId: context.operationId, mutationCount: 1, affectedRanges };
    },
  });

  runtime.registry.registerCommand<{ sheetId: string }>({
    id: 'sheet.columns.unhide.all',
    execute: (params, context) => {
      const sheet = context.workbook.getSheet(params.sheetId);
      const previous = [...sheet.hiddenColumns].sort((left, right) => left - right);
      if (previous.length === 0) return { operationId: context.operationId, mutationCount: 0, affectedRanges: [] };
      const affectedRanges: RangeRef[] = [{ sheetId: params.sheetId, startRow: 0, endRow: Math.max(0, sheet.rowCount - 1), startColumn: 0, endColumn: Math.max(0, sheet.columnCount - 1) }];
      context.applyMutation({
        id: 'columns.unhidden.all',
        unitId: context.workbook.unitId,
        sheetId: params.sheetId,
        params,
        affectedRanges,
        inverse: [{ id: 'columns.hidden.restore', unitId: context.workbook.unitId, sheetId: params.sheetId, params: { sheetId: params.sheetId, indices: previous }, affectedRanges }],
        apply: () => sheet.hiddenColumns.clear(),
      });
      return { operationId: context.operationId, mutationCount: 1, affectedRanges };
    },
  });

  runtime.registry.registerCommand<{ sheetId: string; at: number; count: number }>({
    id: 'sheet.rows.insert',
    execute: (params, context) => {
      const sheet = context.workbook.getSheet(params.sheetId);
      const affectedRanges: RangeRef[] = [{
        sheetId: params.sheetId,
        startRow: params.at,
        endRow: params.at + params.count - 1,
        startColumn: 0,
        endColumn: Math.max(0, sheet.columnCount - 1),
      }];
      context.applyMutation({
        id: 'rows.inserted',
        unitId: context.workbook.unitId,
        sheetId: params.sheetId,
        params,
        affectedRanges,
        inverse: [
          {
            id: 'rows.deleted',
            unitId: context.workbook.unitId,
            sheetId: params.sheetId,
            params: { sheetId: params.sheetId, at: params.at, count: params.count },
            affectedRanges,
          },
        ],
        apply: () => applyStructuralTransform(context.workbook, { kind: 'insert-rows', sheetId: params.sheetId, at: params.at, count: params.count }),
      });
      return { operationId: context.operationId, mutationCount: 1, affectedRanges };
    },
  });

  runtime.registry.registerCommand<{ sheetId: string; at: number; count: number }>({
    id: 'sheet.rows.delete',
    execute: (params, context) => {
      const sheet = context.workbook.getSheet(params.sheetId);
      const end = params.at + params.count - 1;
      const removed = snapshotCellRegion(sheet, params.at, end, 0, Math.max(0, sheet.columnCount - 1));
      const affectedRanges: RangeRef[] = [{
        sheetId: params.sheetId,
        startRow: params.at,
        endRow: end,
        startColumn: 0,
        endColumn: Math.max(0, sheet.columnCount - 1),
      }];
      context.applyMutation({
        id: 'rows.deleted',
        unitId: context.workbook.unitId,
        sheetId: params.sheetId,
        params,
        affectedRanges,
        inverse: [
          {
            id: 'rows.inserted',
            unitId: context.workbook.unitId,
            sheetId: params.sheetId,
            params: { sheetId: params.sheetId, at: params.at, count: params.count },
            affectedRanges,
          },
          ...removed.map((entry) => ({
            id: 'cell.restore' as const,
            unitId: context.workbook.unitId,
            sheetId: params.sheetId,
            params: { sheetId: params.sheetId, row: entry.row, column: entry.column, previous: entry.cell },
            affectedRanges: cellRange({ sheetId: params.sheetId, row: entry.row, column: entry.column, value: entry.cell }),
          })),
        ],
        apply: () => applyStructuralTransform(context.workbook, { kind: 'delete-rows', sheetId: params.sheetId, at: params.at, count: params.count }),
      });
      return { operationId: context.operationId, mutationCount: 1, affectedRanges };
    },
  });

  runtime.registry.registerCommand<{ sheetId: string; at: number; count: number }>({
    id: 'sheet.columns.insert',
    execute: (params, context) => {
      const sheet = context.workbook.getSheet(params.sheetId);
      const affectedRanges: RangeRef[] = [{
        sheetId: params.sheetId,
        startRow: 0,
        endRow: Math.max(0, sheet.rowCount - 1),
        startColumn: params.at,
        endColumn: params.at + params.count - 1,
      }];
      context.applyMutation({
        id: 'columns.inserted',
        unitId: context.workbook.unitId,
        sheetId: params.sheetId,
        params,
        affectedRanges,
        inverse: [
          {
            id: 'columns.deleted',
            unitId: context.workbook.unitId,
            sheetId: params.sheetId,
            params: { sheetId: params.sheetId, at: params.at, count: params.count },
            affectedRanges,
          },
        ],
        apply: () => applyStructuralTransform(context.workbook, { kind: 'insert-columns', sheetId: params.sheetId, at: params.at, count: params.count }),
      });
      return { operationId: context.operationId, mutationCount: 1, affectedRanges };
    },
  });

  runtime.registry.registerCommand<{ sheetId: string; at: number; count: number }>({
    id: 'sheet.columns.delete',
    execute: (params, context) => {
      const sheet = context.workbook.getSheet(params.sheetId);
      const end = params.at + params.count - 1;
      const removed = snapshotCellRegion(sheet, 0, Math.max(0, sheet.rowCount - 1), params.at, end);
      const affectedRanges: RangeRef[] = [{
        sheetId: params.sheetId,
        startRow: 0,
        endRow: Math.max(0, sheet.rowCount - 1),
        startColumn: params.at,
        endColumn: end,
      }];
      context.applyMutation({
        id: 'columns.deleted',
        unitId: context.workbook.unitId,
        sheetId: params.sheetId,
        params,
        affectedRanges,
        inverse: [
          {
            id: 'columns.inserted',
            unitId: context.workbook.unitId,
            sheetId: params.sheetId,
            params: { sheetId: params.sheetId, at: params.at, count: params.count },
            affectedRanges,
          },
          ...removed.map((entry) => ({
            id: 'cell.restore' as const,
            unitId: context.workbook.unitId,
            sheetId: params.sheetId,
            params: { sheetId: params.sheetId, row: entry.row, column: entry.column, previous: entry.cell },
            affectedRanges: cellRange({ sheetId: params.sheetId, row: entry.row, column: entry.column, value: entry.cell }),
          })),
        ],
        apply: () => applyStructuralTransform(context.workbook, { kind: 'delete-columns', sheetId: params.sheetId, at: params.at, count: params.count }),
      });
      return { operationId: context.operationId, mutationCount: 1, affectedRanges };
    },
  });

  // 12. 多列排序 / 转置 / 翻转 / 拆分
  runtime.registry.registerCommand<{
    sheetId: string;
    range: RangeRef;
    criteria: Array<{ column: number; ascending: boolean }>;
    hasHeader: boolean;
  }>({
    id: 'sheet.sort.multi',
    execute: (params, context) => runtime.execute('data.sort.rows', params),
  });

  runtime.registry.registerCommand<{ sheetId: string; row: number; column: number; delimiter: string; maxColumns?: number }>({
    id: 'sheet.splitColumn',
    execute: (params, context) => {
      const sheet = context.workbook.getSheet(params.sheetId);
      const cell = sheet.cells.get(params.row, params.column);
      const text = cell?.value == null ? '' : String(cell.value);
      const maxColumns = Math.max(2, params.maxColumns ?? 4);
      const parts = text.split(params.delimiter).slice(0, maxColumns);
      if (parts.length <= 1 && parts[0] === text) {
        return { operationId: context.operationId, mutationCount: 0, affectedRanges: [] };
      }
      const baseStyle = cell?.style ? structuredClone(cell.style) : undefined;
      const values: CellData[][] = [parts.map((part) => ({ value: coerceText(part, cell), style: baseStyle ? structuredClone(baseStyle) : undefined }))];
      while (values[0]!.length < 1) values[0]!.push({ value: null });
      return runtime.execute('sheet.range.set', {
        sheetId: params.sheetId,
        startRow: params.row,
        startColumn: params.column,
        values,
      });
    },
  });

  // 13. 筛选 / 条件格式 / 数据验证 / 色带 / 名称
  runtime.registry.registerMutation<{ sheetId: string; autoFilter: AutoFilterModel }>({
    id: 'autoFilter.set',
    handler: (item, context) => {
      if (!isFilterMutation(item.params)) throw new Error('Invalid autoFilter.set mutation payload');
      const params = item.params;
      const sheet = context.workbook.getSheet(params.sheetId);
      sheet.autoFilter = validateFilterOwnership(sheet, params.autoFilter, { kind: 'worksheet' });
    },
    metadata: {
      schema: { name: 'AutoFilterSet', validate: isFilterMutation },
      permission: { capability: 'sheet.autoFilter.write', roles: ['owner', 'editor'] },
      affectedRanges: { resolve: (params) => [structuredClone(params.autoFilter.range)], mode: 'exact' },
      inverseIds: ['autoFilter.set', 'autoFilter.remove'],
    },
  });
  runtime.registry.registerMutation<{ sheetId: string; range?: RangeRef }>({
    id: 'autoFilter.remove',
    handler: (item, context) => {
      if (!isFilterRemoveMutation(item.params)) throw new Error('Invalid autoFilter.remove mutation payload');
      context.workbook.getSheet(item.params.sheetId).autoFilter = undefined;
    },
    metadata: {
      schema: { name: 'AutoFilterRemove', validate: isFilterRemoveMutation },
      permission: { capability: 'sheet.autoFilter.write', roles: ['owner', 'editor'] },
      affectedRanges: { resolve: (params) => params.range ? [structuredClone(params.range)] : [], mode: 'declared' },
      inverseIds: ['autoFilter.set'],
    },
  });
  runtime.registry.registerCommand<{ sheetId: string; autoFilter: AutoFilterModel }>({
    id: 'sheet.autoFilter.set',
    execute: (params, context) => {
      const sheet = context.workbook.getSheet(params.sheetId);
      const autoFilter = validateFilterOwnership(sheet, params.autoFilter, { kind: 'worksheet' });
      const previous = sheet.autoFilter;
      const affectedRanges: RangeRef[] = [structuredClone(autoFilter.range)];
      context.applyMutation({
        id: 'autoFilter.set',
        unitId: context.workbook.unitId,
        sheetId: params.sheetId,
        params: { ...params, autoFilter },
        affectedRanges,
        inverse: previous
          ? [{
            id: 'autoFilter.set',
            unitId: context.workbook.unitId,
            sheetId: params.sheetId,
            params: { sheetId: params.sheetId, autoFilter: structuredClone(previous) },
            affectedRanges,
          }]
          : [{
            id: 'autoFilter.remove',
            unitId: context.workbook.unitId,
            sheetId: params.sheetId,
            params: { sheetId: params.sheetId, range: autoFilter.range },
            affectedRanges,
          }],
        apply: () => {
          sheet.autoFilter = structuredClone(autoFilter);
        },
      });
      return { operationId: context.operationId, mutationCount: 1, affectedRanges };
    },
  });
  runtime.registry.registerCommand<{ sheetId: string }>({
    id: 'sheet.autoFilter.remove',
    execute: (params, context) => {
      const sheet = context.workbook.getSheet(params.sheetId);
      const previous = sheet.autoFilter;
      if (!previous) return { operationId: context.operationId, mutationCount: 0, affectedRanges: [] };
      const affectedRanges: RangeRef[] = [structuredClone(previous.range)];
      context.applyMutation({
        id: 'autoFilter.remove',
        unitId: context.workbook.unitId,
        sheetId: params.sheetId,
        params: { ...params, range: previous.range },
        affectedRanges,
        inverse: [
          {
            id: 'autoFilter.set',
            unitId: context.workbook.unitId,
            sheetId: params.sheetId,
            params: { sheetId: params.sheetId, autoFilter: structuredClone(previous) },
            affectedRanges,
          },
        ],
        apply: () => {
          context.workbook.getSheet(params.sheetId).autoFilter = undefined;
        },
      });
      return { operationId: context.operationId, mutationCount: 1, affectedRanges };
    },
  });

  runtime.registry.registerMutation<AddConditionalFormatParams>({
    id: 'cf.add',
    handler: (item, context) => {
      if (!isConditionalAddMutation(item.params)) throw new Error('Invalid cf.add mutation payload');
      const params = item.params;
      const sheet = context.workbook.getSheet(params.rule.sheetId);
      const index = sheet.conditionalFormats.findIndex((rule) => rule.id === params.rule.id);
      if (index >= 0) sheet.conditionalFormats[index] = structuredClone(params.rule);
      else sheet.conditionalFormats.push(structuredClone(params.rule));
    },
    metadata: {
      schema: { name: 'ConditionalFormatAdd', validate: isConditionalAddMutation },
      permission: { capability: 'sheet.conditional-format.write', roles: ['owner', 'editor'] },
      affectedRanges: { resolve: ruleRanges, mode: 'exact' },
      inverseIds: ['cf.remove'],
    },
  });
  runtime.registry.registerMutation<{ sheetId: string; ruleId: string; ranges?: RangeRef[] }>({
    id: 'cf.remove',
    handler: (item, context) => {
      if (!isRuleRemoveMutation(item.params)) throw new Error('Invalid cf.remove mutation payload');
      const params = item.params;
      const sheet = context.workbook.getSheet(params.sheetId);
      const index = sheet.conditionalFormats.findIndex((rule) => rule.id === params.ruleId);
      if (index >= 0) sheet.conditionalFormats.splice(index, 1);
    },
    metadata: {
      schema: { name: 'ConditionalFormatRemove', validate: isRuleRemoveMutation },
      permission: { capability: 'sheet.conditional-format.write', roles: ['owner', 'editor'] },
      affectedRanges: { resolve: removeRuleRanges, mode: 'declared' },
      inverseIds: ['cf.add'],
    },
  });
  runtime.registry.registerMutation<{ sheetId: string; ranges: RangeRef[] }>({
    id: 'cf.clear',
    handler: (item, context) => {
      if (!isRecord(item.params) || typeof item.params.sheetId !== 'string') throw new Error('Invalid cf.clear mutation payload');
      context.workbook.getSheet(item.params.sheetId).conditionalFormats.length = 0;
    },
    metadata: {
      schema: { name: 'ConditionalFormatClear', validate: isSheetRangesMutation },
      permission: { capability: 'sheet.conditional-format.write', roles: ['owner', 'editor'] },
      affectedRanges: { resolve: (params) => params.ranges.map((range) => structuredClone(range)), mode: 'declared' },
      inverseIds: ['cf.add'],
    },
  });
  runtime.registry.registerCommand<AddConditionalFormatParams>({
    id: 'sheet.cf.add',
    execute: (params, context) => {
      const sheet = context.workbook.getSheet(params.rule.sheetId);
      const normalizedRule = normalizeConditionalFormatRule(params.rule, sheet.conditionalFormats.length + 1);
      const normalizedParams = { ...params, rule: normalizedRule };
      const affectedRanges: RangeRef[] = structuredClone(normalizedRule.ranges);
      context.applyMutation({
        id: 'cf.add',
        unitId: context.workbook.unitId,
        sheetId: params.rule.sheetId,
        params: normalizedParams,
        affectedRanges,
        inverse: [
          {
            id: 'cf.remove',
            unitId: context.workbook.unitId,
            sheetId: params.rule.sheetId,
            params: { sheetId: normalizedRule.sheetId, ruleId: normalizedRule.id, ranges: normalizedRule.ranges },
            affectedRanges,
          },
        ],
        apply: () => {
          const target = context.workbook.getSheet(normalizedRule.sheetId);
          const index = target.conditionalFormats.findIndex((rule) => rule.id === normalizedRule.id);
          if (index >= 0) target.conditionalFormats[index] = structuredClone(normalizedRule);
          else target.conditionalFormats.push(structuredClone(normalizedRule));
        },
      });
      return { operationId: context.operationId, mutationCount: 1, affectedRanges };
    },
  });
  runtime.registry.registerCommand<{ sheetId: string; ruleId: string }>({
    id: 'sheet.cf.remove',
    execute: (params, context) => {
      const sheet = context.workbook.getSheet(params.sheetId);
      const index = sheet.conditionalFormats.findIndex((rule) => rule.id === params.ruleId);
      if (index < 0) return { operationId: context.operationId, mutationCount: 0, affectedRanges: [] };
      const previous = structuredClone(sheet.conditionalFormats[index]!);
      const affectedRanges: RangeRef[] = structuredClone(previous.ranges);
      context.applyMutation({
        id: 'cf.remove',
        unitId: context.workbook.unitId,
        sheetId: params.sheetId,
        params: { ...params, ranges: previous.ranges },
        affectedRanges,
        inverse: [
          {
            id: 'cf.add',
            unitId: context.workbook.unitId,
            sheetId: params.sheetId,
            params: { sheetId: params.sheetId, rule: previous } satisfies AddConditionalFormatParams,
            affectedRanges,
          },
        ],
        apply: () => {
          const target = context.workbook.getSheet(params.sheetId);
          const idx = target.conditionalFormats.findIndex((rule) => rule.id === params.ruleId);
          if (idx >= 0) target.conditionalFormats.splice(idx, 1);
        },
      });
      return { operationId: context.operationId, mutationCount: 1, affectedRanges };
    },
  });
  runtime.registry.registerCommand<{ sheetId: string }>({
    id: 'sheet.cf.clear',
    execute: (params, context) => {
      const sheet = context.workbook.getSheet(params.sheetId);
      if (sheet.conditionalFormats.length === 0) {
        return { operationId: context.operationId, mutationCount: 0, affectedRanges: [] };
      }
      const previous = structuredClone(sheet.conditionalFormats);
      const affectedRanges: RangeRef[] = previous.flatMap((rule) => structuredClone(rule.ranges));
      context.applyMutation({
        id: 'cf.clear',
        unitId: context.workbook.unitId,
        sheetId: params.sheetId,
        params: { ...params, ranges: affectedRanges },
        affectedRanges,
        inverse: previous.map((rule) => ({
          id: 'cf.add' as const,
          unitId: context.workbook.unitId,
          sheetId: params.sheetId,
          params: { sheetId: params.sheetId, rule } satisfies AddConditionalFormatParams,
          affectedRanges: [] as RangeRef[],
        })),
        apply: () => {
          context.workbook.getSheet(params.sheetId).conditionalFormats.length = 0;
        },
      });
      return { operationId: context.operationId, mutationCount: 1, affectedRanges };
    },
  });

  runtime.registry.registerMutation<AddDataValidationParams>({
    id: 'dv.add',
    handler: (item, context) => {
      if (!isDataValidationAddMutation(item.params)) throw new Error('Invalid dv.add mutation payload');
      const params = item.params;
      const sheet = context.workbook.getSheet(params.rule.sheetId);
      const index = sheet.dataValidations.findIndex((rule) => rule.id === params.rule.id);
      if (index >= 0) sheet.dataValidations[index] = structuredClone(params.rule);
      else sheet.dataValidations.push(structuredClone(params.rule));
    },
    metadata: {
      schema: { name: 'DataValidationAdd', validate: isDataValidationAddMutation },
      permission: { capability: 'sheet.data-validation.write', roles: ['owner', 'editor'] },
      affectedRanges: { resolve: ruleRanges, mode: 'exact' },
      inverseIds: ['dv.remove'],
    },
  });
  runtime.registry.registerMutation<{ sheetId: string; ruleId: string; ranges?: RangeRef[] }>({
    id: 'dv.remove',
    handler: (item, context) => {
      if (!isRuleRemoveMutation(item.params)) throw new Error('Invalid dv.remove mutation payload');
      const params = item.params;
      const sheet = context.workbook.getSheet(params.sheetId);
      const index = sheet.dataValidations.findIndex((rule) => rule.id === params.ruleId);
      if (index >= 0) sheet.dataValidations.splice(index, 1);
    },
    metadata: {
      schema: { name: 'DataValidationRemove', validate: isRuleRemoveMutation },
      permission: { capability: 'sheet.data-validation.write', roles: ['owner', 'editor'] },
      affectedRanges: { resolve: removeRuleRanges, mode: 'declared' },
      inverseIds: ['dv.add'],
    },
  });
  runtime.registry.registerCommand<AddDataValidationParams>({
    id: 'sheet.dv.add',
    execute: (params, context) => {
      const normalizedRule = normalizeDataValidationRule(params.rule);
      const normalizedParams = { ...params, rule: normalizedRule };
      const affectedRanges: RangeRef[] = structuredClone(normalizedRule.ranges);
      context.applyMutation({
        id: 'dv.add',
        unitId: context.workbook.unitId,
        sheetId: params.rule.sheetId,
        params: normalizedParams,
        affectedRanges,
        inverse: [
          {
            id: 'dv.remove',
            unitId: context.workbook.unitId,
            sheetId: params.rule.sheetId,
            params: { sheetId: normalizedRule.sheetId, ruleId: normalizedRule.id, ranges: normalizedRule.ranges },
            affectedRanges,
          },
        ],
        apply: () => {
          const sheet = context.workbook.getSheet(normalizedRule.sheetId);
          const index = sheet.dataValidations.findIndex((rule) => rule.id === normalizedRule.id);
          if (index >= 0) sheet.dataValidations[index] = structuredClone(normalizedRule);
          else sheet.dataValidations.push(structuredClone(normalizedRule));
        },
      });
      return { operationId: context.operationId, mutationCount: 1, affectedRanges };
    },
  });
  runtime.registry.registerCommand<{ sheetId: string; ruleId: string }>({
    id: 'sheet.dv.remove',
    execute: (params, context) => {
      const sheet = context.workbook.getSheet(params.sheetId);
      const index = sheet.dataValidations.findIndex((rule) => rule.id === params.ruleId);
      if (index < 0) return { operationId: context.operationId, mutationCount: 0, affectedRanges: [] };
      const previous = structuredClone(sheet.dataValidations[index]!);
      const affectedRanges: RangeRef[] = structuredClone(previous.ranges);
      context.applyMutation({
        id: 'dv.remove',
        unitId: context.workbook.unitId,
        sheetId: params.sheetId,
        params: { ...params, ranges: previous.ranges },
        affectedRanges,
        inverse: [
          {
            id: 'dv.add',
            unitId: context.workbook.unitId,
            sheetId: params.sheetId,
            params: { sheetId: params.sheetId, rule: previous } satisfies AddDataValidationParams,
            affectedRanges,
          },
        ],
        apply: () => {
          const target = context.workbook.getSheet(params.sheetId);
          const idx = target.dataValidations.findIndex((rule) => rule.id === params.ruleId);
          if (idx >= 0) target.dataValidations.splice(idx, 1);
        },
      });
      return { operationId: context.operationId, mutationCount: 1, affectedRanges };
    },
  });

  runtime.registry.registerMutation<{ sheetId: string; rule: BandedRule | null }>({
    id: 'banded.set',
    handler: (item, context) => {
      if (!isBandedMutation(item.params)) throw new Error('Invalid banded.set mutation payload');
      const params = item.params;
      const sheet = context.workbook.getSheet(params.sheetId);
      sheet.bandedRule = params.rule ? structuredClone(params.rule) : undefined;
    },
    metadata: {
      schema: { name: 'BandedRuleSet', validate: isBandedMutation },
      permission: { capability: 'sheet.format.write', roles: ['owner', 'editor'] },
      affectedRanges: { resolve: (params) => params.rule ? [structuredClone(params.rule.range)] : [], mode: 'declared' },
      inverseIds: ['banded.set'],
    },
  });
  runtime.registry.registerCommand<{ sheetId: string; rule: BandedRule | null }>({
    id: 'sheet.banded.set',
    execute: (params, context) => {
      const previous = context.workbook.getSheet(params.sheetId).bandedRule;
      const affectedRanges: RangeRef[] = params.rule ? [structuredClone(params.rule.range)] : [];
      context.applyMutation({
        id: 'banded.set',
        unitId: context.workbook.unitId,
        sheetId: params.sheetId,
        params,
        affectedRanges,
        inverse: [
          {
            id: 'banded.set',
            unitId: context.workbook.unitId,
            sheetId: params.sheetId,
            params: { sheetId: params.sheetId, rule: previous ? structuredClone(previous) : null },
            affectedRanges,
          },
        ],
        apply: () => {
          const sheet = context.workbook.getSheet(params.sheetId);
          sheet.bandedRule = params.rule ? structuredClone(params.rule) : undefined;
        },
      });
      return { operationId: context.operationId, mutationCount: 1, affectedRanges };
    },
  });

  runtime.registry.registerMutation<{ model: DefinedNameModel }>({
    id: 'name.set',
    handler: (item, context) => {
      if (!isNameSetMutation(item.params)) throw new Error('Invalid name.set mutation payload');
      context.workbook.setDefinedName(item.params.model);
    },
    metadata: {
      schema: { name: 'DefinedNameSet', validate: isNameSetMutation },
      permission: { capability: 'workbook.defined-name.write', roles: ['owner', 'editor'] },
      affectedRanges: { resolve: () => [], mode: 'exact' },
      inverseIds: ['name.set', 'name.remove'],
    },
  });
  runtime.registry.registerMutation<{ name: string; scope?: 'workbook' | 'sheet'; sheetId?: string }>({
    id: 'name.remove',
    handler: (item, context) => {
      if (!isNameRemoveMutation(item.params)) throw new Error('Invalid name.remove mutation payload');
      const params = item.params;
      context.workbook.removeDefinedName(params.name, params.scope ?? 'workbook', params.sheetId);
    },
    metadata: {
      schema: { name: 'DefinedNameRemove', validate: isNameRemoveMutation },
      permission: { capability: 'workbook.defined-name.write', roles: ['owner', 'editor'] },
      affectedRanges: { resolve: () => [], mode: 'exact' },
      inverseIds: ['name.set'],
    },
  });
  runtime.registry.registerCommand<{
    name: string;
    value?: string;
    formula?: string;
    scope?: 'workbook' | 'sheet';
    sheetId?: string;
    hidden?: boolean;
    comment?: string;
  }>({
    id: 'workbook.name.set',
    execute: (params, context) => {
      const model: DefinedNameModel = {
        name: params.name,
        formula: params.formula ?? params.value ?? '',
        scope: params.scope ?? 'workbook',
        ...(params.sheetId ? { sheetId: params.sheetId } : {}),
        ...(params.hidden === undefined ? {} : { hidden: params.hidden }),
        ...(params.comment === undefined ? {} : { comment: params.comment }),
      };
      // Validate before opening a mutation so invalid scope/name input cannot
      // create a history entry or leave the legacy formula view half updated.
      const normalized = normalizeDefinedNameModel(model);
      const previous = context.workbook.getDefinedNameExact(normalized.name, normalized.scope, normalized.sheetId);
      const affectedRanges: RangeRef[] = [];
      context.applyMutation({
        id: 'name.set',
        unitId: context.workbook.unitId,
        sheetId: context.workbook.primarySheetId,
        params: { model: normalized },
        affectedRanges,
        inverse: previous !== undefined
          ? [{ id: 'name.set', unitId: context.workbook.unitId, sheetId: context.workbook.primarySheetId, params: { model: previous }, affectedRanges }]
          : [{ id: 'name.remove', unitId: context.workbook.unitId, sheetId: context.workbook.primarySheetId, params: { name: normalized.name, scope: normalized.scope, sheetId: normalized.sheetId }, affectedRanges }],
        apply: () => {
          context.workbook.setDefinedName(normalized);
        },
      });
      return { operationId: context.operationId, mutationCount: 1, affectedRanges };
    },
  });
  runtime.registry.registerCommand<{ name: string; scope?: 'workbook' | 'sheet'; sheetId?: string }>({
    id: 'workbook.name.remove',
    execute: (params, context) => {
      const previous = context.workbook.getDefinedNameExact(params.name, params.scope ?? 'workbook', params.sheetId);
      if (previous === undefined || previous.scope !== (params.scope ?? 'workbook') || previous.sheetId !== params.sheetId) return { operationId: context.operationId, mutationCount: 0, affectedRanges: [] };
      const affectedRanges: RangeRef[] = [];
      context.applyMutation({
        id: 'name.remove',
        unitId: context.workbook.unitId,
        sheetId: context.workbook.primarySheetId,
        params,
        affectedRanges,
        inverse: [
          { id: 'name.set', unitId: context.workbook.unitId, sheetId: context.workbook.primarySheetId, params: { model: previous }, affectedRanges },
        ],
        apply: () => {
          context.workbook.removeDefinedName(params.name, previous.scope, previous.sheetId);
        },
      });
      return { operationId: context.operationId, mutationCount: 1, affectedRanges };
    },
  });
}

function coerceText(text: string, previousCell: CellData | undefined): CellData['value'] {
  if (typeof previousCell?.value === 'number') {
    const numeric = Number(text.replace(/[$,%]/g, ''));
    if (Number.isFinite(numeric) && text.trim() !== '') return numeric;
  }
  return text;
}
