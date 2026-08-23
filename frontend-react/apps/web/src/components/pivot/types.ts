import type { ReactNode } from 'react';
import type { PivotResultTree } from '@react-sheets/core-model';

export type PivotFieldType = 'text' | 'number' | 'date' | 'boolean';
export type PivotFieldArea = 'filters' | 'columns' | 'rows' | 'values';
export type PivotSummaryFunction = 'sum' | 'count' | 'count-numbers' | 'average' | 'min' | 'max' | 'product' | 'stdev' | 'stdevp' | 'var' | 'varp' | 'distinct-count';
export type PivotShowAs = 'normal' | 'percent-of-total' | 'percent-of-row' | 'percent-of-column' | 'percent-of-parent' | 'difference-from' | 'percent-difference-from' | 'running-total' | 'rank' | 'index';
export type PivotSortDirection = 'none' | 'ascending' | 'descending';

export interface PivotFieldDefinition {
  id: string;
  label: string;
  type: PivotFieldType;
  description?: string;
  values?: readonly string[];
}

export interface PivotValueDefinition {
  id: string;
  fieldId: string;
  summary: PivotSummaryFunction;
  displayName: string;
  numberFormat: string;
  showAs: PivotShowAs;
  baseFieldId?: string;
  baseItem?: string;
}
export interface PivotCalculatedFieldDefinition { name: string; formula: string; }
export interface PivotCalculatedItemDefinition { fieldId: string; name: string; formula: string; }

export interface PivotDefinition {
  sourceRange?: string;
  filters: string[];
  columns: string[];
  rows: string[];
  values: PivotValueDefinition[];
  calculatedFields: PivotCalculatedFieldDefinition[];
  calculatedItems: PivotCalculatedItemDefinition[];
  filterSelections: Record<string, string[]>;
  sort: Record<string, PivotSortDirection>;
  groupedFields: string[];
  layout: 'compact' | 'outline' | 'tabular';
  showGrandTotals: boolean;
  showSubtotals: boolean;
  expandedFieldIds: string[];
  slicers: string[];
  timelineFieldId?: string;
  timelineStart?: string;
  timelineEnd?: string;
  pivotChart?: { type: 'column' | 'bar' | 'line' | 'pie'; title: string };
}

export interface PivotResult {
  rowCount: number;
  columnCount: number;
  updatedAt?: string;
  summary?: string;
  tree?: PivotResultTree;
}

export interface PivotPanelState {
  loading?: boolean;
  disabled?: boolean;
  error?: string;
  empty?: boolean;
}

export interface PivotPanelCallbacks {
  onCreate?: () => void;
  onPivotSelect?: (pivotId: string) => void;
  onFieldAreaChange: (fieldId: string, area: PivotFieldArea, index: number) => void;
  onRemoveField: (fieldId: string, area: PivotFieldArea, index: number) => void;
  onValueChange: (value: PivotValueDefinition) => void;
  onCalculatedFieldsChange?: (fields: PivotCalculatedFieldDefinition[]) => void;
  onCalculatedItemsChange?: (items: PivotCalculatedItemDefinition[]) => void;
  onFilterChange: (fieldId: string, selectedValues: readonly string[]) => void;
  onSortChange: (fieldId: string, direction: PivotSortDirection) => void;
  onGroupChange: (fieldId: string, grouped: boolean) => void;
  onRefresh: () => void;
  onSourceRangeChange?: (sourceRange: string) => void;
  onLayoutChange: (layout: PivotDefinition['layout']) => void;
  onExpandedChange: (fieldId: string, expanded: boolean) => void;
  onSlicerChange: (fieldId: string, enabled: boolean) => void;
  onTimelineChange: (fieldId: string | undefined) => void;
  onTimelineRangeChange?: (start: string, end: string) => void;
  onPivotChartChange: (chart: PivotDefinition['pivotChart']) => void;
  onClose?: () => void;
}

export interface PivotPanelSlots {
  headerActions?: ReactNode;
  resultSummary?: ReactNode;
}
