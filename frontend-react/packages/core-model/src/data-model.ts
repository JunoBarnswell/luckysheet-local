import type { SheetId } from './index';

export type TableScalar = string | number | boolean | null;

export type TableFieldType = 'text' | 'number' | 'boolean' | 'date' | 'mixed';

export interface WorkbookTableField {
  id: string;
  name: string;
  ordinal: number;
  type: TableFieldType;
}

export interface WorkbookTableBlock {
  id: string;
  tableId: string;
  startRow: number;
  rowCount: number;
  storageKey: string;
  encoding: 'typed-column-v1';
}

export interface WorkbookTableModel {
  id: string;
  name: string;
  sourceSheetId?: SheetId;
  rowCount: number;
  fields: WorkbookTableField[];
  blockSize: number;
  blocks: WorkbookTableBlock[];
  revision: number;
}
