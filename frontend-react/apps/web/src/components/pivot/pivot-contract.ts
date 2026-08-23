import type {
  PivotCalculatedField,
  PivotCalculatedItem,
  PivotFieldDefinition,
  PivotModel,
  PivotResultTree,
  PivotSourceRowPath,
  PivotValueField,
} from '@react-sheets/core-model';
import type { ReactNode } from 'react';

export type { PivotCalculatedField, PivotCalculatedItem, PivotFieldDefinition, PivotModel, PivotResultTree, PivotSourceRowPath, PivotValueField };

export type PivotFieldArea = 'filters' | 'columns' | 'rows' | 'values';
export type PivotSortDirection = 'none' | 'ascending' | 'descending';

export interface PivotPanelResult {
  rowCount: number;
  columnCount: number;
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
  onValueChange: (value: PivotValueField) => void;
  onCalculatedFieldsChange?: (fields: PivotCalculatedField[]) => void;
  onCalculatedItemsChange?: (items: PivotCalculatedItem[]) => void;
  onFilterChange: (fieldId: string, selectedValues: readonly string[]) => void;
  onSortChange: (fieldId: string, direction: PivotSortDirection) => void;
  onGroupChange: (fieldId: string, grouped: boolean) => void;
  onRefresh: () => void;
  onSourceRangeChange?: (sourceRange: string) => void;
  onLayoutChange: (layout: 'compact' | 'outline' | 'tabular') => void;
  onExpandedChange: (fieldId: string, expanded: boolean) => void;
  onSlicerChange: (fieldId: string, enabled: boolean) => void;
  onTimelineChange: (fieldId: string | undefined) => void;
  onTimelineRangeChange?: (start: string, end: string) => void;
  onPivotChartChange: (chart: { type: 'column' | 'bar' | 'line' | 'pie'; title: string } | undefined) => void;
}

export interface PivotPanelSlots {
  headerActions?: ReactNode;
  resultSummary?: ReactNode;
}
