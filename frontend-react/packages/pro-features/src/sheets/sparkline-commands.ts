import type { RangeRef, SparklineGroup, SparklineModel } from '@react-sheets/core-model';

/** Public sparkline command payload types. Registration lives in spreadsheet-app. */
export interface SparklineInsertParams {
  sheetId: string;
  sparkline: SparklineModel;
}

export interface SparklineUpdateParams {
  sheetId: string;
  sparklineId: string;
  patch: Partial<SparklineModel>;
}

export interface SparklineGroupCreateParams {
  sheetId: string;
  group: SparklineGroup;
}

export interface SparklineGroupUpdateParams {
  sheetId: string;
  groupId: string;
  patch: Partial<SparklineGroup>;
}

export interface SparklineInsertDialogParams {
  sheetId: string;
  sparklineId: string;
  dataRange: RangeRef;
  location: { row: number; column: number };
  type: SparklineModel['type'];
  groupId?: string;
}
