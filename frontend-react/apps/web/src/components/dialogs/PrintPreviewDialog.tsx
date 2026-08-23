import { useEffect, useMemo, useState } from 'react';
import { paginateRange, type PrintLayout } from '@react-sheets/pro-features';
import { Box, Button, Dialog, Inline, Stack, Text, TextInput } from '@react-sheets/ui-system';

export interface PrintPreviewRow { rowNumber: number; cells: readonly { value: string }[]; }

export interface PrintPreviewDialogProps {
  open: boolean;
  onClose: () => void;
  sheetId: string;
  rowCount: number;
  columnCount: number;
  columns: readonly string[];
  rows: readonly PrintPreviewRow[];
  getRow?: (row: number) => PrintPreviewRow;
  layout?: PrintLayout;
}

function columnLabel(column: number): string {
  let label = '';
  let remaining = column + 1;
  while (remaining > 0) {
    const modulo = (remaining - 1) % 26;
    label = String.fromCharCode(65 + modulo) + label;
    remaining = Math.floor((remaining - 1) / 26);
  }
  return label;
}

export function PrintPreviewDialog({ columns, getRow, layout, open, onClose, rows, sheetId, rowCount, columnCount }: PrintPreviewDialogProps) {
  const [rowsPerPage, setRowsPerPage] = useState(30);
  const [columnsPerPage, setColumnsPerPage] = useState(6);
  const [pageIndex, setPageIndex] = useState(0);

  const pages = useMemo(() => {
    if (!open || rowCount === 0 || columnCount === 0) return [];
    return paginateRange(
      { sheetId, startRow: 0, endRow: rowCount - 1, startColumn: 0, endColumn: columnCount - 1 },
      Math.max(1, rowsPerPage),
      Math.max(1, columnsPerPage),
    );
  }, [open, rowCount, columnCount, sheetId, rowsPerPage, columnsPerPage]);
  useEffect(() => setPageIndex(0), [rowsPerPage, columnsPerPage, open]);
  const page = pages[pageIndex];

  return (
    <Dialog open={open} onClose={onClose} title="Print preview" maxWidth="xl">
      <Stack gap="md">
          <Inline gap="md" className="items-end">
          <Stack gap="xs"><Text size="xs" weight="medium">Rows / page</Text><TextInput type="number" min={1} max={200} value={rowsPerPage} onChange={(event) => setRowsPerPage(Number(event.target.value) || 1)} className="w-20" /></Stack>
          <Stack gap="xs"><Text size="xs" weight="medium">Columns / page</Text><TextInput type="number" min={1} max={40} value={columnsPerPage} onChange={(event) => setColumnsPerPage(Number(event.target.value) || 1)} className="w-20" /></Stack>
          <Text size="sm" tone="muted">{pages.length} page(s)</Text>
        </Inline>
        <Box className="max-h-[50vh] overflow-auto rounded-xl border border-slate-200 bg-white p-3 dark:border-slate-700 dark:bg-slate-900">
          {page ? <Stack gap="xs" className="min-w-max">
            <Inline gap="none" className="border-b border-slate-300 bg-slate-100 font-semibold">
              <Text size="xs" className="w-12 shrink-0 border-r border-slate-200 px-2 py-1">#</Text>
              {columns.slice(page.range.startColumn, page.range.endColumn + 1).map((column) => <Text key={column} size="xs" className="w-28 shrink-0 border-r border-slate-200 px-2 py-1">{column}</Text>)}
            </Inline>
            {Array.from({ length: page.range.endRow - page.range.startRow + 1 }, (_, offset) => getRow?.(page.range.startRow + offset) ?? rows[page.range.startRow + offset]).filter((row): row is PrintPreviewRow => Boolean(row)).map((row) => <Inline key={row.rowNumber} gap="none" className="border-b border-slate-100"><Text size="xs" className="w-12 shrink-0 border-r border-slate-200 px-2 py-1 text-slate-400">{row.rowNumber}</Text>{row.cells.slice(page.range.startColumn, page.range.endColumn + 1).map((cell, index) => <Text key={`${row.rowNumber}-${index}`} size="xs" className="w-28 shrink-0 truncate border-r border-slate-100 px-2 py-1">{cell.value}</Text>)}</Inline>)}
          </Stack> : <Text size="sm" tone="muted">No printable cells.</Text>}
        </Box>
        <Inline gap="sm" className="items-center justify-between"><Button disabled={pageIndex <= 0} size="xs" variant="ghost" onClick={() => setPageIndex((index) => Math.max(0, index - 1))}>Previous page</Button><Text size="xs" tone="muted">{page ? `Page ${page.page} · Rows ${page.range.startRow + 1}-${page.range.endRow + 1} · ${columnLabel(page.range.startColumn)}-${columnLabel(page.range.endColumn)}` : 'No page'}</Text><Button disabled={pageIndex >= pages.length - 1} size="xs" variant="ghost" onClick={() => setPageIndex((index) => Math.min(pages.length - 1, index + 1))}>Next page</Button></Inline>
        <Inline gap="sm" className="justify-end">
          <Button variant="ghost" onClick={onClose}>Close</Button>
          <Button variant="primary" onClick={() => {
            if (layout) {
              document.documentElement.dataset.printPaper = layout.paper;
              document.documentElement.dataset.printOrientation = layout.orientation;
            }
            window.print();
          }}>Print</Button>
        </Inline>
      </Stack>
    </Dialog>
  );
}
