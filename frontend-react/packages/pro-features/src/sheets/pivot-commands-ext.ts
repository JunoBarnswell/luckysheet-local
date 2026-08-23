import type {
  PivotAggregateFunction,
  PivotGroup,
  PivotLayout,
  PivotShowAs,
  PivotSlicer,
  PivotTimeline,
} from '@react-sheets/core-model';

/** Public pivot command payload types. Registration lives in spreadsheet-app. */
export interface PivotLayoutCommandParams {
  sheetId: string;
  pivotId: string;
  layout: PivotLayout;
}

export interface PivotAggregateParams {
  sheetId: string;
  pivotId: string;
  field: string;
  summarizeBy: PivotAggregateFunction;
}

export interface PivotShowAsParams {
  sheetId: string;
  pivotId: string;
  field: string;
  showAs: PivotShowAs;
}

export interface PivotGroupParams {
  sheetId: string;
  pivotId: string;
  axis: 'rows' | 'columns';
  field: string;
  group: PivotGroup;
}

export interface PivotDrillDownParams {
  sheetId: string;
  pivotId: string;
  label: string;
  sourceRowPaths: Array<{ sheetId: string; row: number }>;
  targetSheetId: string;
  targetAnchor: { row: number; column: number };
}

export interface PivotSlicerParams {
  sheetId: string;
  pivotId: string;
  slicer: PivotSlicer;
}

export interface PivotTimelineParams {
  sheetId: string;
  pivotId: string;
  timeline: PivotTimeline;
}
