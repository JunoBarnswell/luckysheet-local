import { useEffect, useMemo, useState } from 'react';
import type { RangeRef } from '@react-sheets/core-model';
import type { PrintLayout } from '@react-sheets/spreadsheet-app';
import { Box, Button, Dialog, Inline, Stack, Text, TextInput } from '@react-sheets/ui-system';

export interface PrintPreviewRow { rowNumber: number; cells: readonly { value: string }[]; }

export interface PrintPreviewPage {
  page: number;
  range: RangeRef;
}

export interface PrintPreviewDialogProps {
  open: boolean;
  onClose: () => void;
  sheetId: string;
  rowCount: number;
  columnCount: number;
  columns: readonly string[];
  rows: readonly PrintPreviewRow[];
  getRow?: (row: number) => PrintPreviewRow | undefined;
  layout?: PrintLayout;
  pages?: readonly PrintPreviewPage[];
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

export function PrintPreviewDialog({
  columns,
  getRow,
  layout,
  open,
  onClose,
  pages: externalPages,
  rows,
  sheetId,
  rowCount,
  columnCount,
}: PrintPreviewDialogProps) {
  const [rowsPerPage, setRowsPerPage] = useState(30);
  const [columnsPerPage, setColumnsPerPage] = useState(6);
  const [pageIndex, setPageIndex] = useState(0);

  const internalPages = useMemo(() => {
    if (!open || rowCount === 0 || columnCount === 0) return [];
    const pages: PrintPreviewPage[] = [];
    for (let startRow = 0; startRow < rowCount; startRow += Math.max(1, rowsPerPage)) {
      for (let startColumn = 0; startColumn < columnCount; startColumn += Math.max(1, columnsPerPage)) {
        pages.push({
          page: pages.length + 1,
          range: {
            sheetId,
            startRow,
            endRow: Math.min(rowCount - 1, startRow + Math.max(1, rowsPerPage) - 1),
            startColumn,
            endColumn: Math.min(columnCount - 1, startColumn + Math.max(1, columnsPerPage) - 1),
          },
        });
      }
    }
    return pages;
  }, [open, rowCount, columnCount, sheetId, rowsPerPage, columnsPerPage]);

  const pages = externalPages && externalPages.length > 0 ? externalPages : internalPages;
  useEffect(() => setPageIndex(0), [rowsPerPage, columnsPerPage, open, externalPages]);
  const page = pages[pageIndex];

  return (
    <Dialog open={open} onClose={onClose} title="Print preview" maxWidth="xl">
      <Stack gap="md">
        {externalPages && externalPages.length > 0 ? (
          <Text size="sm" tone="muted">{pages.length} page(s) from print layout</Text>
        ) : (
          <Inline gap="md" className="items-end">
            <Stack gap="xs"><Text size="xs" weight="medium">Rows / page</Text><TextInput type="number" min={1} max={200} value={rowsPerPage} onChange={(event) => setRowsPerPage(Number(event.target.value) || 1)} className="w-20" /></Stack>
            <Stack gap="xs"><Text size="xs" weight="medium">Columns / page</Text><TextInput type="number" min={1} max={40} value={columnsPerPage} onChange={(event) => setColumnsPerPage(Number(event.target.value) || 1)} className="w-20" /></Stack>
            <Text size="sm" tone="muted">{pages.length} page(s)</Text>
          </Inline>
        )}
        <Box className="max-h-[50vh] overflow-auto rounded-xl border border-slate-200 bg-white p-3 dark:border-slate-700 dark:bg-slate-900">
          {page ? <Stack gap="xs" className="min-w-max">
            <Inline gap="none" className="border-b border-slate-300 bg-slate-100 font-semibold">
              <Text size="xs" className="w-12 shrink-0 border-r border-slate-200 px-2 py-1">#</Text>
              {columns.slice(page.range.startColumn, page.range.endColumn + 1).map((column) => <Text key={column} size="xs" className="w-28 shrink-0 border-r border-slate-200 px-2 py-1">{column}</Text>)}
            </Inline>
            {Array.from({ length: page.range.endRow - page.range.startRow + 1 }, (_, offset) => {
              const row = page.range.startRow + offset;
              return getRow ? getRow(row) : rows[row];
            }).filter((row): row is PrintPreviewRow => Boolean(row)).map((row) => <Inline key={row.rowNumber} gap="none" className="border-b border-slate-100"><Text size="xs" className="w-12 shrink-0 border-r border-slate-200 px-2 py-1 text-slate-400">{row.rowNumber}</Text>{row.cells.slice(page.range.startColumn, page.range.endColumn + 1).map((cell, index) => <Text key={`${row.rowNumber}-${index}`} size="xs" className="w-28 shrink-0 truncate border-r border-slate-100 px-2 py-1">{cell.value}</Text>)}</Inline>)}
          </Stack> : <Text size="sm" tone="muted">No printable cells.</Text>}
        </Box>
        <Inline gap="sm" className="items-center justify-between"><Button disabled={pageIndex <= 0} size="xs" variant="ghost" onClick={() => setPageIndex((index) => Math.max(0, index - 1))}>Previous page</Button><Text size="xs" tone="muted">{page ? `Page ${page.page} · Rows ${page.range.startRow + 1}-${page.range.endRow + 1} · ${columnLabel(page.range.startColumn)}-${columnLabel(page.range.endColumn)}` : 'No page'}</Text><Button disabled={pageIndex >= pages.length - 1} size="xs" variant="ghost" onClick={() => setPageIndex((index) => Math.min(pages.length - 1, index + 1))}>Next page</Button></Inline>
        <Inline gap="sm" className="justify-end">
          <Button variant="ghost" onClick={onClose}>Close</Button>
          <Button variant="primary" onClick={onClose}>Use Export PDF</Button>
        </Inline>
      </Stack>
    </Dialog>
  );
}
