import type { TableScalar, WorkbookTableModel } from '@react-sheets/core-model';

/** Local workbook data-plane page; it is not an HTTP contract. */
export interface TableRowsResponse {
  table: WorkbookTableModel;
  rows: TableScalar[][];
  nextOffset?: number;
}
