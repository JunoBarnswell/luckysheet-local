import type { ReportBinding, ReportSheetDefinition, WorkbookTableModel } from '@react-sheets/core-model';
import type { CanvasSheetSnapshot } from '../../ui-snapshot';

export interface ReportCellProjection {
  row: number;
  column: number;
  value: string;
  sourceRow: number;
  binding: ReportBinding;
}

export interface ReportProjection {
  status: 'ready' | 'error';
  error?: string;
  cells: ReportCellProjection[];
  pageCount: number;
  pageRows: number;
  renderMode: ReportSheetDefinition['renderMode'];
}

function fail(definition: ReportSheetDefinition | undefined, error: string): ReportProjection {
  return { status: 'error', error, cells: [], pageCount: 0, pageRows: definition?.pagination.rowsPerPage ?? 50, renderMode: definition?.renderMode ?? 'design' };
}

function readValue(sheet: CanvasSheetSnapshot, row: number, column: number, table: WorkbookTableModel, sourceSheets: readonly CanvasSheetSnapshot[]): string {
  if (table.sourceRange) {
    const source = sourceSheets.find((candidate) => candidate.id === table.sourceRange!.sheetId);
    if (source) return source.getCell(table.sourceRange.startRow + row, table.sourceRange.startColumn + column)?.value?.trim() ?? '';
  }
  return sheet.getCell(row, column)?.value?.trim() ?? '';
}

export function buildReportProjection(sheet: CanvasSheetSnapshot, tables: readonly WorkbookTableModel[], sourceSheets: readonly CanvasSheetSnapshot[] = []): ReportProjection {
  const definition = sheet.reportSheet;
  if (!definition) return fail(undefined, 'ReportSheet definition is unavailable');
  if (!['design', 'preview', 'paginated'].includes(definition.renderMode)) return fail(definition, 'ReportSheet render mode is invalid');
  if (definition.pagination.enabled && (!Number.isInteger(definition.pagination.rowsPerPage) || (definition.pagination.rowsPerPage ?? 0) <= 0)) return fail(definition, 'ReportSheet rowsPerPage is invalid');
  const table = definition.tableId ? tables.find((candidate) => candidate.id === definition.tableId) : undefined;
  if (definition.tableId && !table) return fail(definition, `ReportSheet binding table ${definition.tableId} is unavailable`);
  if (!table) return { status: 'ready', cells: [], pageCount: 0, pageRows: definition.pagination.rowsPerPage ?? 50, renderMode: definition.renderMode };
  const fields = new Map(table.fields.map((field) => [field.id, field]));
  for (const binding of definition.bindings) {
    if (!Number.isInteger(binding.cell.row) || !Number.isInteger(binding.cell.column) || binding.cell.row < 0 || binding.cell.column < 0 || binding.cell.row >= sheet.rowCount || binding.cell.column >= sheet.columnCount) return fail(definition, 'Report binding cell is outside the template bounds');
    if (!binding.expression.trim()) return fail(definition, 'Report binding expression is empty');
    if ((binding.kind === 'field' || binding.kind === 'group' || binding.kind === 'summary') && !fields.has(binding.expression)) return fail(definition, `Report binding field is unavailable: ${binding.expression}`);
    if (binding.direction && !['vertical', 'horizontal'].includes(binding.direction)) return fail(definition, 'Report binding direction is invalid');
    if (binding.fill && !['none', 'down', 'right'].includes(binding.fill)) return fail(definition, 'Report binding fill is invalid');
  }
  const rowCount = table.sourceRange ? table.rowCount : Math.max(0, sheet.rowCount - 1);
  const pages = definition.pagination.enabled ? Math.ceil(rowCount / (definition.pagination.rowsPerPage ?? 50)) : (rowCount > 0 ? 1 : 0);
  const cells: ReportCellProjection[] = [];
  for (const binding of definition.bindings) {
    const sourceField = fields.get(binding.expression);
    const fieldColumn = sourceField?.ordinal ?? -1;
    const repeated = definition.renderMode !== 'design' && (binding.fill === 'down' || binding.direction === 'vertical' || binding.kind === 'field' || binding.kind === 'group' || binding.kind === 'summary');
    const count = repeated ? rowCount : 1;
    for (let index = 0; index < count; index += 1) {
      const sourceRow = index + 1;
      const row = binding.direction === 'horizontal' || binding.fill === 'right' ? binding.cell.row : binding.cell.row + index;
      const column = binding.direction === 'horizontal' || binding.fill === 'right' ? binding.cell.column + index : binding.cell.column;
      if (row >= sheet.rowCount || column >= sheet.columnCount) return fail(definition, 'Report binding expansion exceeds worksheet bounds');
      let value = binding.expression;
      if (definition.renderMode === 'design' && binding.kind !== 'static') value = `{${binding.expression}}`;
      else if (binding.kind === 'field' || binding.kind === 'group' || binding.kind === 'summary') value = fieldColumn >= 0 ? readValue(sheet, sourceRow, fieldColumn, table, sourceSheets) : '';
      else if (binding.kind === 'formula') value = binding.expression;
      cells.push({ row, column, value, sourceRow, binding: structuredClone(binding) });
    }
  }
  return { status: 'ready', cells, pageCount: pages, pageRows: definition.pagination.rowsPerPage ?? 50, renderMode: definition.renderMode };
}
