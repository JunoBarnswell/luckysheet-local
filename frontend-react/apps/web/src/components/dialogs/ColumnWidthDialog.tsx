import { useEffect, useState } from 'react';
import { excelColumnWidthToPixels, MAX_EXCEL_COLUMN_WIDTH, pixelsToExcelColumnWidth } from '@react-sheets/exchange-xlsx';
import { Button, Dialog, Inline, Stack, Text, TextInput } from '@react-sheets/ui-system';

export interface ColumnWidthDialogProps {
  open: boolean;
  columnCount: number;
  initialWidthPx: number;
  defaultMode?: boolean;
  maximumDigitWidthPx: number;
  onClose: () => void;
  onApply: (excelWidth: number) => void;
}

export function ColumnWidthDialog({ open, columnCount, initialWidthPx, defaultMode = false, maximumDigitWidthPx, onClose, onApply }: ColumnWidthDialogProps) {
  const [value, setValue] = useState('8.71');
  useEffect(() => {
    if (open) setValue(pixelsToExcelColumnWidth(initialWidthPx, maximumDigitWidthPx).toFixed(2));
  }, [initialWidthPx, maximumDigitWidthPx, open]);
  const numeric = Number(value);
  const valid = Number.isFinite(numeric) && numeric >= (defaultMode ? 1 / 256 : 0) && numeric <= MAX_EXCEL_COLUMN_WIDTH;
  const pixels = valid ? excelColumnWidthToPixels(numeric, maximumDigitWidthPx) : 0;
  return (
    <Dialog
      open={open}
      onClose={onClose}
      title="Column Width"
      description={defaultMode ? 'Set the default Excel character width for columns without an explicit width.' : `Set the Excel character width for ${columnCount} selected column${columnCount === 1 ? '' : 's'}. Width 0 hides the columns.`}
      maxWidth="sm"
      footer={<><Button size="sm" variant="ghost" onClick={onClose}>Cancel</Button><Button size="sm" variant="primary" disabled={!valid} onClick={() => { onApply(numeric); onClose(); }}>OK</Button></>}
    >
      <Stack gap="sm">
        <Text size="sm" weight="medium">Character width (0–255)</Text>
        <Inline gap="sm" className="items-center">
          <TextInput aria-label="Excel character width" type="number" min={0} max={MAX_EXCEL_COLUMN_WIDTH} step={1 / 256} value={value} onChange={(event) => setValue(event.target.value)} />
          <Text size="sm" tone={valid ? 'muted' : 'danger'}>{valid ? `${pixels}px` : 'Enter 0–255'}</Text>
        </Inline>
      </Stack>
    </Dialog>
  );
}
