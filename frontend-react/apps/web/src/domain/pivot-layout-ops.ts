import type {
  PivotAggregateFunction,
  PivotLayout,
  PivotShowAs,
} from '@react-sheets/core-model';
import type { PivotFieldArea, PivotFieldDefinition, PivotSortDirection, PivotValueDefinition } from '../components/pivot/types';

export function clonePivotLayout(layout: PivotLayout): PivotLayout {
  return structuredClone(layout);
}

export function removeFieldFromLayout(layout: PivotLayout, fieldId: string): PivotLayout {
  const next = clonePivotLayout(layout);
  next.filters = next.filters.filter((filter) => filter.field !== fieldId);
  next.rows = next.rows.filter((field) => field.field !== fieldId);
  next.columns = next.columns.filter((field) => field.field !== fieldId);
  next.values = next.values.filter((value) => value.field !== fieldId);
  return next;
}

function summarizeForField(fieldId: string, fields: readonly PivotFieldDefinition[]): PivotAggregateFunction {
  const field = fields.find((entry) => entry.id === fieldId);
  return field?.type === 'number' ? 'sum' : 'count';
}

export function addFieldToLayout(
  layout: PivotLayout,
  fieldId: string,
  area: PivotFieldArea,
  index: number,
  fields: readonly PivotFieldDefinition[],
): PivotLayout {
  const next = removeFieldFromLayout(layout, fieldId);
  if (area === 'values') {
    next.values.splice(Math.max(0, index), 0, { field: fieldId, summarizeBy: summarizeForField(fieldId, fields) });
    return next;
  }
  if (area === 'filters') {
    next.filters.splice(Math.max(0, index), 0, { kind: 'manual', field: fieldId, selected: [] });
    return next;
  }
  next[area].splice(Math.max(0, index), 0, { field: fieldId });
  return next;
}

export function updateValueInLayout(layout: PivotLayout, value: PivotValueDefinition): PivotLayout {
  const next = clonePivotLayout(layout);
  const index = next.values.findIndex((entry, entryIndex) => `${entry.field}-${entryIndex}` === value.id || entry.field === value.fieldId);
  if (index < 0) return next;
  const current = next.values[index]!;
  next.values[index] = {
    field: value.fieldId,
    summarizeBy: value.summary as PivotAggregateFunction,
    displayName: value.displayName,
    numberFormat: value.numberFormat || undefined,
    baseField: value.baseFieldId,
    baseItem: value.baseItem,
    showAs: fromUiShowAs(value.showAs),
  };
  if (!next.values[index]!.displayName) next.values[index]!.displayName = current.displayName;
  return next;
}

export function setFieldSort(layout: PivotLayout, fieldId: string, direction: PivotSortDirection): PivotLayout {
  const next = clonePivotLayout(layout);
  const applySort = (field: PivotLayout['rows'][number]) => {
    if (field.field !== fieldId) return field;
    return {
      ...field,
      sort: direction === 'none' ? undefined : { direction },
    };
  };
  next.rows = next.rows.map(applySort);
  next.columns = next.columns.map(applySort);
  return next;
}

export function setFieldGrouped(layout: PivotLayout, fieldId: string, grouped: boolean, fields: readonly PivotFieldDefinition[]): PivotLayout {
  const next = clonePivotLayout(layout);
  const applyGroup = (field: PivotLayout['rows'][number]) => {
    if (field.field !== fieldId) return field;
    if (!grouped) return { ...field, group: undefined };
    const catalogField = fields.find((entry) => entry.id === fieldId);
    const group =
      catalogField?.type === 'date'
        ? { kind: 'date' as const, unit: 'month' as const }
        : catalogField?.type === 'number'
          ? { kind: 'number' as const, interval: 10 }
          : undefined;
    return { ...field, group };
  };
  next.rows = next.rows.map(applyGroup);
  next.columns = next.columns.map(applyGroup);
  return next;
}

export function setFilterSelection(layout: PivotLayout, fieldId: string, selectedValues: readonly string[]): PivotLayout {
  const next = clonePivotLayout(layout);
  const existing = next.filters.find((filter) => filter.field === fieldId);
  if (existing?.kind === 'manual') {
    existing.selected = [...selectedValues];
    return next;
  }
  next.filters.push({ kind: 'manual', field: fieldId, selected: [...selectedValues] });
  return next;
}

export function setLayoutMode(layout: PivotLayout, mode: 'compact' | 'outline' | 'tabular'): PivotLayout {
  const next = clonePivotLayout(layout);
  next.compact = mode === 'compact';
  next.repeatLabels = mode === 'tabular';
  return next;
}

export function setExpandedField(layout: PivotLayout, fieldId: string, expanded: boolean): PivotLayout {
  const next = clonePivotLayout(layout);
  const expandedFieldIds = new Set(next.expandedFieldIds ?? next.rows.map((field) => field.field));
  if (expanded) expandedFieldIds.add(fieldId);
  else expandedFieldIds.delete(fieldId);
  next.expandedFieldIds = [...expandedFieldIds];
  return next;
}

function fromUiShowAs(showAs: PivotValueDefinition['showAs']): PivotShowAs | undefined {
  if (showAs === 'percent-of-total') return { kind: 'grand-percentage' };
  if (showAs === 'percent-of-row') return { kind: 'row-percentage' };
  if (showAs === 'percent-of-column') return { kind: 'column-percentage' };
  if (showAs === 'percent-of-parent') return { kind: 'parent-percentage' };
  if (showAs === 'difference-from') return { kind: 'difference', base: 'grand' };
  if (showAs === 'percent-difference-from') return { kind: 'percentage-difference', base: 'grand' };
  if (showAs === 'running-total') return { kind: 'running-total', axis: 'row' };
  if (showAs === 'rank') return { kind: 'rank', axis: 'row', direction: 'descending' };
  if (showAs === 'index') return { kind: 'index' };
  return { kind: 'normal' };
}
