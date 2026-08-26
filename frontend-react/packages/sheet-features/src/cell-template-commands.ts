import type {
  CellData,
  CellEditorConfig,
  CellStyleTemplate,
  DataValidationRule,
  RangeRef,
} from '@react-sheets/core-model';
import { clearFormulaProvenance } from '@react-sheets/core-model';
import type { CommandContext, CommandRuntime } from '@react-sheets/command-runtime';

export interface SetCellStyleTemplateParams {
  sheetId: string;
  template: CellStyleTemplate;
}

export interface RemoveCellStyleTemplateParams {
  sheetId: string;
  templateId: string;
}

export interface ApplyCellStyleTemplateParams {
  sheetId: string;
  ranges: RangeRef[];
  templateId: string;
}

interface CellEditorSetParams {
  sheetId: string;
  ranges: RangeRef[];
  editor?: CellEditorConfig;
}

export interface CheckboxToggleParams {
  sheetId: string;
  ranges: RangeRef[];
}

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === 'object' && value !== null && !Array.isArray(value);
}

function isRange(value: unknown): value is RangeRef {
  if (!isRecord(value)) return false;
  const startRow = value.startRow;
  const endRow = value.endRow;
  const startColumn = value.startColumn;
  const endColumn = value.endColumn;
  return typeof value.sheetId === 'string'
    && Number.isInteger(startRow) && Number.isInteger(endRow)
    && Number.isInteger(startColumn) && Number.isInteger(endColumn)
    && Number(startRow) >= 0 && Number(startColumn) >= 0
    && Number(endRow) >= Number(startRow) && Number(endColumn) >= Number(startColumn);
}

function isEditor(value: unknown): value is CellEditorConfig {
  if (!isRecord(value) || !['text', 'number', 'date', 'list', 'checkbox'].includes(String(value.kind))) return false;
  return value.values === undefined || (Array.isArray(value.values) && value.values.every((entry) => typeof entry === 'string'));
}

function isTemplate(value: unknown): value is CellStyleTemplate {
  return isRecord(value)
    && typeof value.id === 'string' && value.id.trim().length > 0
    && typeof value.name === 'string' && value.name.trim().length > 0
    && isRecord(value.style)
    && (value.editor === undefined || isEditor(value.editor));
}

function isSetParams(value: unknown): value is SetCellStyleTemplateParams {
  return isRecord(value) && typeof value.sheetId === 'string' && isTemplate(value.template);
}

function isRemoveParams(value: unknown): value is RemoveCellStyleTemplateParams {
  return isRecord(value) && typeof value.sheetId === 'string' && typeof value.templateId === 'string' && value.templateId.trim().length > 0;
}

function isApplyParams(value: unknown): value is ApplyCellStyleTemplateParams {
  return isRecord(value) && typeof value.sheetId === 'string' && typeof value.templateId === 'string'
    && Array.isArray(value.ranges) && value.ranges.length > 0 && value.ranges.every(isRange);
}

function isEditorSetParams(value: unknown): value is CellEditorSetParams {
  return isRecord(value) && typeof value.sheetId === 'string'
    && Array.isArray(value.ranges) && value.ranges.length > 0 && value.ranges.every(isRange)
    && (value.editor === undefined || isEditor(value.editor));
}

function isCheckboxToggleParams(value: unknown): value is CheckboxToggleParams {
  return isRecord(value) && typeof value.sheetId === 'string'
    && Array.isArray(value.ranges) && value.ranges.length > 0 && value.ranges.every(isRange);
}

/** Normalize the only values that can become an in-cell checkbox Boolean. */
export function normalizeCheckboxCellValue(cell: CellData | undefined): boolean {
  if (!cell) return false;
  if (cell.formula !== undefined) throw new Error('Checkbox cannot be applied to a formula cell');
  const value = cell.value;
  if (typeof value === 'boolean') return value;
  if (value === null) return false;
  if (typeof value === 'number' && (value === 0 || value === 1)) return value === 1;
  if (typeof value === 'string') {
    const normalized = value.trim().toUpperCase();
    if (normalized === 'TRUE') return true;
    if (normalized === 'FALSE') return false;
  }
  throw new Error('Checkbox source value must be blank, Boolean, 0/1, TRUE, or FALSE');
}

function assertRangesWithinSheet(sheet: ReturnType<CommandContext['workbook']['getSheet']>, ranges: readonly RangeRef[]): void {
  for (const range of ranges) {
    if (range.sheetId !== sheet.id || range.startRow < 0 || range.endRow >= sheet.rowCount || range.startColumn < 0 || range.endColumn >= sheet.columnCount) {
      throw new Error('Checkbox range is outside worksheet bounds');
    }
  }
}

function normalizeRanges(sheetId: string, ranges: readonly RangeRef[]): RangeRef[] {
  return ranges.map((range) => {
    if (range.sheetId !== sheetId) throw new Error('Template range must target the active worksheet');
    return {
      sheetId,
      startRow: Math.min(range.startRow, range.endRow),
      endRow: Math.max(range.startRow, range.endRow),
      startColumn: Math.min(range.startColumn, range.endColumn),
      endColumn: Math.max(range.startColumn, range.endColumn),
    };
  });
}

function templateValidationId(templateId: string, sheetId: string, range: RangeRef, index: number): string {
  return `template-validation:${templateId}:${sheetId}:${range.startRow}:${range.startColumn}:${range.endRow}:${range.endColumn}:${index}`;
}

function templateValidationRule(template: CellStyleTemplate, sheetId: string, ranges: RangeRef[], index: number): DataValidationRule | undefined {
  const validation = template.dataValidation ?? (template.editor?.kind === 'list'
    ? { type: 'list' as const, listSource: { kind: 'values' as const, values: template.editor.values ?? [] }, allowBlank: true, showDropdown: true }
    : template.editor?.kind === 'checkbox'
      ? { type: 'checkbox' as const, allowBlank: true }
      : template.editor?.kind === 'date'
        ? { type: 'date' as const, allowBlank: true }
        : template.editor?.kind === 'number'
          ? { type: 'decimal' as const, allowBlank: true }
          : undefined);
  if (!validation) return undefined;
  return {
    ...structuredClone(validation),
    id: templateValidationId(template.id, sheetId, ranges[index]!, index),
    sheetId,
    ranges: [structuredClone(ranges[index]!)],
  };
}

function editorAffectedRanges(params: CellEditorSetParams): RangeRef[] {
  return params.ranges.map((range) => structuredClone(range));
}

function applyEditorMutation(params: CellEditorSetParams, context: CommandContext): void {
  const sheet = context.workbook.getSheet(params.sheetId);
  assertRangesWithinSheet(sheet, params.ranges);
  for (const range of params.ranges) {
    for (let row = range.startRow; row <= range.endRow; row += 1) {
      for (let column = range.startColumn; column <= range.endColumn; column += 1) {
        const current = structuredClone(sheet.cells.get(row, column) ?? { value: null as CellData['value'] });
        if (params.editor?.kind === 'checkbox') current.value = normalizeCheckboxCellValue(current);
        if (params.editor) current.editor = structuredClone(params.editor);
        else delete current.editor;
        sheet.cells.set(row, column, current);
      }
    }
  }
}

/** Registers workbook-native template and cell-editor commands. */
export function registerCellTemplateCommands(runtime: CommandRuntime): void {
  runtime.registry.registerMutation<SetCellStyleTemplateParams>({
    id: 'cellTemplate.set',
    handler: (item, context) => context.workbook.setCellStyleTemplate(item.params.template),
    metadata: {
      schema: { name: 'CellStyleTemplateSet', validate: isSetParams },
      permission: { capability: 'sheet.format.write', roles: ['owner', 'editor'] },
      affectedRanges: { resolve: () => [], mode: 'exact' },
      inverseIds: ['cellTemplate.set', 'cellTemplate.remove'],
    },
  });
  runtime.registry.registerMutation<RemoveCellStyleTemplateParams>({
    id: 'cellTemplate.remove',
    handler: (item, context) => { context.workbook.removeCellStyleTemplate(item.params.templateId); },
    metadata: {
      schema: { name: 'CellStyleTemplateRemove', validate: isRemoveParams },
      permission: { capability: 'sheet.format.write', roles: ['owner', 'editor'] },
      affectedRanges: { resolve: () => [], mode: 'exact' },
      inverseIds: ['cellTemplate.set'],
    },
  });
  runtime.registry.registerMutation<CellEditorSetParams>({
    id: 'cell.editor.set',
    handler: (item, context) => applyEditorMutation(item.params, context),
    metadata: {
      schema: { name: 'CellEditorSet', validate: isEditorSetParams },
      permission: { capability: 'sheet.format.write', roles: ['owner', 'editor'] },
      affectedRanges: { resolve: editorAffectedRanges, mode: 'exact' },
      inverseIds: ['cell.restore'],
    },
  });

  runtime.registry.registerCommand<SetCellStyleTemplateParams>({
    id: 'workbook.cellTemplate.set',
    execute: (params, context) => {
      if (!isSetParams(params)) throw new Error('Invalid cell style template');
      const previous = context.workbook.cellStyleTemplates.get(params.template.id);
      context.applyMutation({
        id: 'cellTemplate.set',
        unitId: context.workbook.unitId,
        sheetId: params.sheetId,
        params: { ...params, template: structuredClone(params.template) },
        affectedRanges: [],
        inverse: previous
          ? [{ id: 'cellTemplate.set', unitId: context.workbook.unitId, sheetId: params.sheetId, params: { sheetId: params.sheetId, template: structuredClone(previous) }, affectedRanges: [] }]
          : [{ id: 'cellTemplate.remove', unitId: context.workbook.unitId, sheetId: params.sheetId, params: { sheetId: params.sheetId, templateId: params.template.id }, affectedRanges: [] }],
        apply: () => context.workbook.setCellStyleTemplate(params.template),
      });
      return { operationId: context.operationId, mutationCount: 1, affectedRanges: [] };
    },
  });
  runtime.registry.registerCommand<RemoveCellStyleTemplateParams>({
    id: 'workbook.cellTemplate.remove',
    execute: (params, context) => {
      if (!isRemoveParams(params)) throw new Error('Invalid cell style template removal');
      const previous = context.workbook.cellStyleTemplates.get(params.templateId);
      if (!previous) return { operationId: context.operationId, mutationCount: 0, affectedRanges: [] };
      context.applyMutation({
        id: 'cellTemplate.remove',
        unitId: context.workbook.unitId,
        sheetId: params.sheetId,
        params,
        affectedRanges: [],
        inverse: [{ id: 'cellTemplate.set', unitId: context.workbook.unitId, sheetId: params.sheetId, params: { sheetId: params.sheetId, template: structuredClone(previous) }, affectedRanges: [] }],
        apply: () => { context.workbook.removeCellStyleTemplate(params.templateId); },
      });
      return { operationId: context.operationId, mutationCount: 1, affectedRanges: [] };
    },
  });
  runtime.registry.registerCommand<CellEditorSetParams>({
    id: 'sheet.cellEditor.set',
    execute: (params, context) => {
      if (!isEditorSetParams(params)) throw new Error('Invalid cell editor configuration');
      const ranges = normalizeRanges(params.sheetId, params.ranges);
      const sheet = context.workbook.getSheet(params.sheetId);
      assertRangesWithinSheet(sheet, ranges);
      if (params.editor?.kind === 'checkbox') {
        for (const range of ranges) {
          for (let row = range.startRow; row <= range.endRow; row += 1) {
            for (let column = range.startColumn; column <= range.endColumn; column += 1) normalizeCheckboxCellValue(sheet.cells.get(row, column));
          }
        }
      }
      const previous: Array<{ row: number; column: number; cell?: CellData }> = [];
      for (const range of ranges) {
        for (let row = range.startRow; row <= range.endRow; row += 1) {
          for (let column = range.startColumn; column <= range.endColumn; column += 1) {
            previous.push({ row, column, cell: structuredClone(sheet.cells.get(row, column)) });
          }
        }
      }
      context.applyMutation({
        id: 'cell.editor.set',
        unitId: context.workbook.unitId,
        sheetId: params.sheetId,
        params: { ...params, ranges },
        affectedRanges: ranges,
        inverse: previous.map((entry) => ({
          id: 'cell.restore' as const,
          unitId: context.workbook.unitId,
          sheetId: params.sheetId,
          params: { sheetId: params.sheetId, row: entry.row, column: entry.column, previous: entry.cell },
          affectedRanges: [{ sheetId: params.sheetId, startRow: entry.row, endRow: entry.row, startColumn: entry.column, endColumn: entry.column }],
        })),
        apply: () => applyEditorMutation({ ...params, ranges }, context),
      });
      return { operationId: context.operationId, mutationCount: 1, affectedRanges: ranges };
    },
  });
  runtime.registry.registerCommand<CheckboxToggleParams>({
    id: 'checkbox.toggle',
    execute: (params, context) => {
      if (!isCheckboxToggleParams(params)) throw new Error('Invalid checkbox toggle range');
      const ranges = normalizeRanges(params.sheetId, params.ranges);
      const sheet = context.workbook.getSheet(params.sheetId);
      assertRangesWithinSheet(sheet, ranges);
      const entries: Array<{ row: number; column: number; previous: CellData; next: CellData }> = [];
      const seen = new Set<string>();
      for (const range of ranges) {
        for (let row = range.startRow; row <= range.endRow; row += 1) {
          for (let column = range.startColumn; column <= range.endColumn; column += 1) {
            const key = `${row}:${column}`;
            if (seen.has(key)) continue;
            seen.add(key);
            const previous = structuredClone(sheet.cells.get(row, column));
            if (!previous?.editor || previous.editor.kind !== 'checkbox' || typeof previous.value !== 'boolean') throw new Error(`Cell ${row},${column} is not a canonical Boolean checkbox`);
            entries.push({ row, column, previous, next: { ...previous, value: !previous.value } });
          }
        }
      }
      const affectedRanges = entries.map((entry) => ({ sheetId: params.sheetId, startRow: entry.row, endRow: entry.row, startColumn: entry.column, endColumn: entry.column }));
      for (const entry of entries) {
        const cellRange = { sheetId: params.sheetId, startRow: entry.row, endRow: entry.row, startColumn: entry.column, endColumn: entry.column };
        const next = clearFormulaProvenance(entry.next);
        context.applyMutation({ id: 'cell.set', unitId: context.workbook.unitId, sheetId: params.sheetId, params: { sheetId: params.sheetId, row: entry.row, column: entry.column, value: next }, affectedRanges: [cellRange],
          inverse: [{ id: 'cell.restore', unitId: context.workbook.unitId, sheetId: params.sheetId, params: { sheetId: params.sheetId, row: entry.row, column: entry.column, previous: entry.previous }, affectedRanges: [cellRange] }],
          apply: () => sheet.cells.set(entry.row, entry.column, next),
        });
      }
      return { operationId: context.operationId, mutationCount: entries.length, affectedRanges };
    },
  });
  runtime.registry.registerCommand<ApplyCellStyleTemplateParams>({
    id: 'sheet.cellTemplate.apply',
    execute: (params, context) => {
      if (!isApplyParams(params)) throw new Error('Invalid cell style template application');
      const template = context.workbook.cellStyleTemplates.get(params.templateId);
      if (!template) throw new Error(`Unknown cell style template: ${params.templateId}`);
      const ranges = normalizeRanges(params.sheetId, params.ranges);
      const styleResult = runtime.execute('sheet.style.setMulti', { sheetId: params.sheetId, ranges, style: structuredClone(template.style) });
      let mutationCount = styleResult.mutationCount;
      let affectedRanges = [...styleResult.affectedRanges];
      if (template.editor) {
        const editorResult = runtime.execute('sheet.cellEditor.set', { sheetId: params.sheetId, ranges, editor: structuredClone(template.editor) });
        mutationCount += editorResult.mutationCount;
        affectedRanges = [...affectedRanges, ...editorResult.affectedRanges];
      }
      ranges.forEach((range, index) => {
        const rule = templateValidationRule(template, params.sheetId, ranges, index);
        if (!rule) return;
        const validationResult = runtime.execute('sheet.dv.add', { sheetId: params.sheetId, rule });
        mutationCount += validationResult.mutationCount;
        affectedRanges = [...affectedRanges, ...validationResult.affectedRanges];
      });
      return { operationId: context.operationId, mutationCount, affectedRanges };
    },
  });
}
