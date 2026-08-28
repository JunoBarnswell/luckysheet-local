import { useEffect, useState } from 'react';
import { pixelsToPoints, pointsToPixels } from '@react-sheets/exchange-excel-ooxml';
import { Button, Dialog, Inline, Stack, Text, TextInput } from '@react-sheets/ui-system';
import { MAX_EXCEL_ROW_HEIGHT_POINTS } from '../../editor/column-dimension-controller';

export interface RowHeightDialogProps {
  open: boolean;
  rowCount: number;
  initialHeightPx: number;
  onClose: () => void;
  onApply: (points: number) => void;
}

/** Excel points are a display/exchange unit; the worksheet stores CSS pixels. */
export function RowHeightDialog({ open, rowCount, initialHeightPx, onClose, onApply }: RowHeightDialogProps) {
  const [value, setValue] = useState('15');
  useEffect(() => {
    if (open) setValue(pixelsToPoints(initialHeightPx).toFixed(2));
  }, [initialHeightPx, open]);
  const numeric = Number(value);
  const valid = Number.isFinite(numeric) && numeric >= 0 && numeric <= MAX_EXCEL_ROW_HEIGHT_POINTS;
  const pixels = valid ? Math.round(pointsToPixels(numeric)) : 0;
  return (
    <Dialog
      open={open}
      onClose={onClose}
      title="Row Height"
      description={`Set the height in points for ${rowCount} selected row${rowCount === 1 ? '' : 's'}. A zero value hides the selected rows.`}
      maxWidth="sm"
      footer={<><Button size="sm" variant="ghost" onClick={onClose}>Cancel</Button><Button size="sm" variant="primary" disabled={!valid} onClick={() => { onApply(numeric); onClose(); }}>OK</Button></>}
    >
      <Stack gap="sm">
        <Text size="sm" weight="medium">Point height (0–409)</Text>
        <Inline gap="sm" className="items-center">
          <TextInput aria-label="Row height in points" type="number" min={0} max={MAX_EXCEL_ROW_HEIGHT_POINTS} step={0.01} value={value} onChange={(event) => setValue(event.target.value)} />
          <Text size="sm" tone={valid ? 'muted' : 'danger'}>{valid ? `${pixels}px` : 'Enter 0–409 points'}</Text>
        </Inline>
      </Stack>
    </Dialog>
  );
}
