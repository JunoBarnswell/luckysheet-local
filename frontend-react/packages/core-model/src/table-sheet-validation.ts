import type { TableSheetDefinition } from './data-model';
function isRecord(value: unknown): value is Record<string, any> { return !!value && typeof value === 'object' && !Array.isArray(value); }
function isTableSheetColumn(value: unknown): value is TableSheetDefinition['columns'][number] {
  if (!isRecord(value) || typeof value.fieldId !== 'string' || value.fieldId.trim().length === 0 || typeof value.caption !== 'string' || value.caption.trim().length === 0) return false;
  if (value.widthPx !== undefined && (typeof value.widthPx !== 'number' || !Number.isFinite(value.widthPx) || value.widthPx <= 0)) return false;
  if (value.type !== undefined && typeof value.type !== 'string') return false;
  return value.formula === undefined || typeof value.formula === 'string';
}

export function isTableSheetDefinition(value: unknown): value is TableSheetDefinition {
  if (!isRecord(value) || typeof value.viewId !== 'string' || value.viewId.trim().length === 0 || !Array.isArray(value.columns) || !Array.isArray(value.grouping)) return false;
  if (!value.columns.every(isTableSheetColumn)) return false;
  const columnIds = new Set(value.columns.map((column) => column.fieldId));
  if (columnIds.size !== value.columns.length) return false;
  if (!value.grouping.every((group) => isRecord(group) && typeof group.fieldId === 'string' && columnIds.has(group.fieldId) && (group.collapsed === undefined || typeof group.collapsed === 'boolean'))) return false;
  if (new Set(value.grouping.map((group) => group.fieldId)).size !== value.grouping.length) return false;
  if (value.sortState !== undefined) {
    if (!Array.isArray(value.sortState) || !value.sortState.every((sort) => isRecord(sort) && typeof sort.fieldId === 'string' && columnIds.has(sort.fieldId) && (sort.direction === 'asc' || sort.direction === 'desc'))) return false;
    if (new Set(value.sortState.map((sort) => sort.fieldId)).size !== value.sortState.length) return false;
  }
  return true;
}
