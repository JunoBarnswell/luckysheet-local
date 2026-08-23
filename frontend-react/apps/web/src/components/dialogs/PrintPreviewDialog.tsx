import { useMemo, useState } from 'react';
import { paginateRange } from '@react-sheets/pro-features';
import { Box, Button, Dialog, Inline, Stack, Text } from '@react-sheets/ui-system';

export interface PrintPreviewDialogProps {
  open: boolean;
  onClose: () => void;
  sheetId: string;
  rowCount: number;
  columnCount: number;
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

export function PrintPreviewDialog({ open, onClose, sheetId, rowCount, columnCount }: PrintPreviewDialogProps) {
  const [rowsPerPage, setRowsPerPage] = useState(30);
  const [columnsPerPage, setColumnsPerPage] = useState(6);

  const pages = useMemo(() => {
    if (!open || rowCount === 0 || columnCount === 0) return [];
    return paginateRange(
      { sheetId, startRow: 0, endRow: rowCount - 1, startColumn: 0, endColumn: columnCount - 1 },
      Math.max(1, rowsPerPage),
      Math.max(1, columnsPerPage),
    );
  }, [open, rowCount, columnCount, sheetId, rowsPerPage, columnsPerPage]);

  return (
    <Dialog open={open} onClose={onClose} title="Print preview" maxWidth="xl">
      <Stack gap="md">
        <Inline gap="md">
          <label className="flex items-center gap-2 text-sm">
            Rows / page
            <input type="number" min={1} max={200} value={rowsPerPage} onChange={(event) => setRowsPerPage(Number(event.target.value))} className="w-20 rounded-lg border border-slate-300 px-2 py-1 text-sm dark:border-slate-600 dark:bg-slate-800" />
          </label>
          <label className="flex items-center gap-2 text-sm">
            Columns / page
            <input type="number" min={1} max={40} value={columnsPerPage} onChange={(event) => setColumnsPerPage(Number(event.target.value))} className="w-20 rounded-lg border border-slate-300 px-2 py-1 text-sm dark:border-slate-600 dark:bg-slate-800" />
          </label>
          <Text size="sm" tone="muted">{pages.length} page(s)</Text>
        </Inline>
        <Box className="max-h-[50vh] overflow-y-auto rounded-xl border border-slate-200 bg-white p-3 dark:border-slate-700 dark:bg-slate-900">
          <div className="grid grid-cols-2 gap-3 md:grid-cols-3">
            {pages.map((page) => (
              <Box key={page.page} className="rounded-lg border border-dashed border-slate-300 p-2 dark:border-slate-600">
                <Text size="xs" tone="subtle">Page {page.page}</Text>
                <Text size="xs" tone="default" weight="medium">
                  Rows {page.range.startRow + 1}-{page.range.endRow + 1} · {columnLabel(page.range.startColumn)}-{columnLabel(page.range.endColumn)}
                </Text>
              </Box>
            ))}
          </div>
        </Box>
        <Inline gap="sm" className="justify-end">
          <Button variant="ghost" onClick={onClose}>Close</Button>
          <Button variant="primary" onClick={() => window.print()}>Print</Button>
        </Inline>
      </Stack>
    </Dialog>
  );
}