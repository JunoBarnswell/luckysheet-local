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
  PivotPresentation,
  PivotDisplayOptions,
  PivotRefreshPolicy,
  PivotSubtotalDefinition,
  PivotSlicerItemProjection,
  PivotSlicerSettings,
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
  PivotPresentation,
  PivotDisplayOptions,
  PivotRefreshPolicy,
  PivotSubtotalDefinition,
  PivotSlicerItemProjection,
  PivotSlicerSettings,
};

export type PivotFieldArea = 'filters' | 'columns' | 'rows' | 'values';
export type PivotExpansionCommand =
  | { kind: 'expand-field'; fieldId: string }
  | { kind: 'collapse-field'; fieldId: string }
  | { kind: 'toggle-buttons'; showButtons: boolean };
export type PivotFieldPaneLayout = 'stacked' | 'side-by-side' | 'areas-2x2' | 'areas-1x4' | 'fields-only' | 'areas-only';
export const PIVOT_FIELD_PANE_LAYOUTS: readonly PivotFieldPaneLayout[] = ['stacked', 'side-by-side', 'areas-2x2', 'areas-1x4', 'fields-only', 'areas-only'];
export function defaultPivotFieldArea(field: Pick<PivotFieldDefinition, 'dataType'>): PivotFieldArea {
  if (field.dataType === 'number') return 'values';
  if (field.dataType === 'date') return 'columns';
  return 'rows';
}
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
  settings: PivotSlicerSettings;
  items: readonly PivotSlicerItemProjection[];
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
  onSubtotalChange?: (fieldId: string, subtotal: PivotSubtotalDefinition) => void;
  onSubtotalLocationChange?: (location: import('@react-sheets/core-model').PivotSubtotalLocation) => void;
  onRefresh: () => void;
  onLayoutChange: (layout: 'compact' | 'outline' | 'tabular') => void;
  onLayoutReplace: (layout: import('@react-sheets/core-model').PivotLayout) => void;
  onPresentationChange?: (presentation: PivotPresentation) => void;
  onDisplayOptionsChange?: (displayOptions: PivotDisplayOptions) => void;
  onRefreshPolicyChange?: (refreshPolicy: PivotRefreshPolicy) => void;
  onExpansionCommand?: (command: PivotExpansionCommand) => void;
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
