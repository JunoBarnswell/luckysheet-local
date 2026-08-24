import { useEffect, useState } from 'react';
import { Box, Button, Dialog, DropdownMenu, Inline, RadioOption, Stack, Text, TextInput } from '@react-sheets/ui-system';
import type { RangeRef } from '@react-sheets/core-model';
import type { Locale } from '../../i18n';
import { pivotText } from '../pivot/pivot-localization';

export interface CreatePivotTableRequest { sourceId: string; destination: 'new-sheet' | 'existing-sheet'; targetSheetId?: string; targetReference?: string }
export interface CreatePivotTableDialogProps { open: boolean; sourceRegion: RangeRef; sourceOptions: readonly { id: string; label: string }[]; activeSheetName: string; locale: Locale; onClose: () => void; onCreate: (request: CreatePivotTableRequest) => void }

function columnLabel(column: number): string { let result = ''; for (let value = column + 1; value > 0; value = Math.floor((value - 1) / 26)) result = String.fromCharCode(65 + ((value - 1) % 26)) + result; return result; }
function absoluteRange(sheetName: string, range: RangeRef): string { return `='${sheetName.replaceAll("'", "''")}'!$${columnLabel(range.startColumn)}$${range.startRow + 1}:$${columnLabel(range.endColumn)}$${range.endRow + 1}`; }

export function CreatePivotTableDialog({ open, sourceRegion, sourceOptions, activeSheetName, locale, onClose, onCreate }: CreatePivotTableDialogProps) {
  const [destination, setDestination] = useState<CreatePivotTableRequest['destination']>('new-sheet');
  const [sourceId, setSourceId] = useState('current-region');
  const [targetReference, setTargetReference] = useState('');
  useEffect(() => { if (open) { setDestination('new-sheet'); setSourceId(sourceOptions[0]?.id ?? 'current-region'); setTargetReference(''); } }, [open, sourceOptions]);
  const sourceFormula = sourceId === (sourceOptions[0]?.id ?? 'current-region') ? absoluteRange(activeSheetName, sourceRegion) : sourceOptions.find((option) => option.id === sourceId)?.label ?? sourceId;
  return (
    <Dialog open={open} title={pivotText(locale, 'createTitle')} maxWidth="pivot" bodyClassName="px-3 pb-3 pt-2" onClose={onClose} footer={<Inline gap="sm" className="w-full justify-end"><Button className="min-w-[98px]" size="sm" variant="primary" disabled={destination === 'existing-sheet' && !targetReference.trim()} data-testid="create-pivot-confirm" onClick={() => onCreate({ sourceId, destination, ...(destination === 'existing-sheet' ? { targetReference: targetReference.trim() } : {}) })}>{pivotText(locale, 'confirm')}</Button><Button className="min-w-[98px]" size="sm" variant="outline" onClick={onClose}>{pivotText(locale, 'cancel')}</Button></Inline>}>
      <Stack gap="md" data-testid="create-pivot-dialog">
        <Stack gap="xs"><Text size="sm">{pivotText(locale, 'chooseData')}</Text><Inline gap="none" className="rounded border border-[#cfd3d7] bg-white"><TextInput aria-label={pivotText(locale, 'chooseData')} className="min-w-0 flex-1 !border-0 font-mono text-[12px]" readOnly value={sourceFormula} /><DropdownMenu align="right" trigger={<Button aria-label={pivotText(locale, 'chooseData')} icon="table" iconOnly size="sm" variant="ghost" className="!h-8 !w-8 rounded-none border-l border-[#d4d7da] text-[#2b88c9]" />}><Stack gap="none" className="min-w-[15rem] p-1">{sourceOptions.map((option) => <Button key={option.id} size="sm" variant="ghost" className="justify-start" onClick={() => setSourceId(option.id)}>{option.label}</Button>)}</Stack></DropdownMenu></Inline></Stack>
        <Stack gap="xs"><Text size="sm">{pivotText(locale, 'chooseLocation')}</Text><RadioOption checked={destination === 'new-sheet'} label={pivotText(locale, 'newWorksheet')} name="pivot-destination" onChange={() => setDestination('new-sheet')} /><RadioOption checked={destination === 'existing-sheet'} label={pivotText(locale, 'existingWorksheet')} name="pivot-destination" onChange={() => setDestination('existing-sheet')} /><Box className="pl-8"><Inline gap="none" className="rounded border border-[#d4d7da] bg-white"><TextInput aria-label={pivotText(locale, 'chooseLocation')} data-testid="create-pivot-location" className="min-w-0 flex-1 !border-0" disabled={destination !== 'existing-sheet'} placeholder="A1" value={targetReference} onChange={(event) => setTargetReference(event.target.value)} /><Button aria-label={pivotText(locale, 'chooseLocation')} disabled={destination !== 'existing-sheet'} icon="table" iconOnly size="sm" variant="ghost" className="!h-8 !w-8 rounded-none border-l border-[#d4d7da] text-[#2b88c9]" onClick={() => setTargetReference(`${columnLabel(sourceRegion.startColumn)}${sourceRegion.startRow + 1}`)} /></Inline></Box></Stack>
      </Stack>
    </Dialog>
  );
}
