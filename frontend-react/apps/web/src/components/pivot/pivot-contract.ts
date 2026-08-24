import type {
  PivotCalculatedField,
  PivotCalculatedItem,
  PivotFieldDefinition,
  PivotFilter,
  PivotGroup,
  PivotMemberKey,
  PivotModel,
  PivotSort,
  PivotValueField,
} from '@react-sheets/core-model';
import type { ReactNode } from 'react';

export type {
  PivotCalculatedField,
  PivotCalculatedItem,
  PivotFieldDefinition,
  PivotFilter,
  PivotGroup,
  PivotMemberKey,
  PivotModel,
  PivotSort,
  PivotValueField,
};

export type PivotFieldArea = 'filters' | 'columns' | 'rows' | 'values';
export type PivotSortDirection = 'none' | 'ascending' | 'descending';

export type PivotFilterMode = 'all' | 'include' | 'exclude';

/**
 * The field-list filter contract is deliberately typed. An empty member list
 * is meaningful only together with `mode: all`; it never means "exclude all".
 */
export interface PivotManualFilterState {
  mode: PivotFilterMode;
  memberKeys: readonly PivotMemberKey[];
}

export interface PivotSlicerControl {
  id: string;
  pivotId: string;
  fieldId: string;
  mode: PivotFilterMode;
  memberKeys: readonly PivotMemberKey[];
  connectedPivotIds?: readonly string[];
}

export interface PivotTimelineControl {
  id: string;
  pivotId: string;
  fieldId: string;
  start?: string;
  end?: string;
  connectedPivotIds?: readonly string[];
}

export interface PivotPanelState {
  loading?: boolean;
  disabled?: boolean;
  error?: string;
  empty?: boolean;
  emptyMessage?: string;
}

export interface PivotPanelCallbacks {
  onCreate?: () => void;
  onPivotSelect?: (pivotId: string) => void;
  onFieldAreaChange: (fieldId: string, area: PivotFieldArea, index: number) => void;
  onRemoveField: (fieldId: string, area: PivotFieldArea, index: number) => void;
  onValueChange: (value: PivotValueField) => void;
  onCalculatedFieldsChange?: (fields: PivotCalculatedField[]) => void;
  onCalculatedItemsChange?: (items: PivotCalculatedItem[]) => void;
  onFilterChange: (fieldId: string, filter: PivotManualFilterState) => void;
  onSortChange: (fieldId: string, sort: PivotSort | undefined) => void;
  onGroupChange: (fieldId: string, group: PivotGroup | undefined) => void;
  onRefresh: () => void;
  onLayoutChange: (layout: 'compact' | 'outline' | 'tabular') => void;
  onLayoutReplace: (layout: import('@react-sheets/core-model').PivotLayout) => void;
  onSlicerChange: (fieldId: string, enabled: boolean) => void;
  onSlicerFilterChange?: (slicerId: string, filter: PivotManualFilterState) => void;
  onTimelineChange: (fieldId: string | undefined) => void;
  onTimelineRangeChange?: (timelineId: string, start: string, end: string) => void;
  onPivotChartChange: (chart: { type: 'column' | 'bar' | 'line' | 'pie'; title: string } | undefined) => void;
}

export interface PivotPanelSlots {
  headerActions?: ReactNode;
  statusSummary?: ReactNode;
}
