import React, { useEffect, useState } from 'react';
import type { CellBorders, CellStyle } from '@react-sheets/core-model';
import { Box, Button, ColorPicker, Inline, Select, Stack, Text, TextInput } from '@react-sheets/ui-system';

export interface FormatCellsDraft {
  numberFormat: string;
  style: Partial<CellStyle>;
}

export interface FormatCellsDialogProps {
  open: boolean;
  initial: FormatCellsDraft;
  onClose: () => void;
  onApply: (draft: FormatCellsDraft) => void;
}

type FormatTab = 'number' | 'alignment' | 'font' | 'border' | 'fill' | 'protection';

const TABS: Array<{ id: FormatTab; label: string }> = [
  { id: 'number', label: 'Number' },
  { id: 'alignment', label: 'Alignment' },
  { id: 'font', label: 'Font' },
  { id: 'border', label: 'Border' },
  { id: 'fill', label: 'Fill' },
  { id: 'protection', label: 'Protection' },
];

const NUMBER_PRESETS: Array<{ label: string; value: string }> = [
  { label: 'General', value: 'general' },
  { label: 'Number (0.00)', value: '0.00' },
  { label: 'Comma (#,##0)', value: '#,##0' },
  { label: 'Currency ($)', value: '$#,##0.00' },
  { label: 'Percent (%)', value: '0%' },
  { label: 'Scientific', value: '0.00E+00' },
  { label: 'Date (yyyy-mm-dd)', value: 'yyyy-mm-dd' },
  { label: 'Time (hh:mm)', value: 'hh:mm' },
  { label: 'Text (@)', value: '@' },
];

function borderSide(style: CellBorders['top']): CellBorders {
  return {
    top: style,
    right: style,
    bottom: style,
    left: style,
  };
}

export function FormatCellsDialog({ open, initial, onClose, onApply }: FormatCellsDialogProps): React.ReactElement | null {
  const [tab, setTab] = useState<FormatTab>('number');
  const [draft, setDraft] = useState<FormatCellsDraft>(initial);

  useEffect(() => {
    if (open) {
      setDraft(initial);
      setTab('number');
    }
  }, [open, initial]);

  if (!open) return null;

  const style = draft.style;
  const setStyle = (patch: Partial<CellStyle>) => setDraft((prev) => ({ ...prev, style: { ...prev.style, ...patch } }));

  return (
    <Box className="fixed inset-0 z-50 flex items-start justify-center bg-slate-900/30 pt-16">
      <Box className="w-[36rem] rounded-xl border border-slate-200 bg-white shadow-2xl" data-testid="format-cells-dialog">
        <Stack gap="none">
          <Box className="border-b border-slate-200 px-4 py-3">
            <Text size="sm" weight="semibold">Format Cells</Text>
          </Box>
          <Inline gap="none" className="min-h-[20rem]">
            <Stack gap="xs" className="w-36 shrink-0 border-r border-slate-200 bg-slate-50 p-2">
              {TABS.map((entry) => (
                <Button
                  key={entry.id}
                  size="sm"
                  variant={tab === entry.id ? 'secondary' : 'ghost'}
                  className="justify-start"
                  data-testid={`format-tab-${entry.id}`}
                  onClick={() => setTab(entry.id)}
                >
                  {entry.label}
                </Button>
              ))}
            </Stack>
            <Box className="min-w-0 flex-1 p-4">
              {tab === 'number' ? (
                <Stack gap="md">
                  <Text size="xs" tone="subtle">Category</Text>
                  <Select
                    sizeVariant="sm"
                    value={draft.numberFormat}
                    onChange={(event) => setDraft((prev) => ({ ...prev, numberFormat: event.target.value }))}
                  >
                    {NUMBER_PRESETS.map((preset) => (
                      <option key={preset.value} value={preset.value}>{preset.label}</option>
                    ))}
                  </Select>
                  <Stack gap="xs">
                    <Text size="xs" tone="subtle">Custom format code</Text>
                    <TextInput
                      aria-label="Number format"
                      data-testid="format-number-code"
                      value={draft.numberFormat}
                      onChange={(event) => setDraft((prev) => ({ ...prev, numberFormat: event.target.value }))}
                    />
                  </Stack>
                </Stack>
              ) : null}

              {tab === 'alignment' ? (
                <Stack gap="md">
                  <Stack gap="xs">
                    <Text size="xs" tone="subtle">Horizontal</Text>
                    <Select
                      sizeVariant="sm"
                      value={style.horizontalAlignment ?? 'left'}
                      onChange={(event) => setStyle({ horizontalAlignment: event.target.value as CellStyle['horizontalAlignment'] })}
                    >
                      <option value="left">Left</option>
                      <option value="center">Center</option>
                      <option value="right">Right</option>
                    </Select>
                  </Stack>
                  <Stack gap="xs">
                    <Text size="xs" tone="subtle">Vertical</Text>
                    <Select
                      sizeVariant="sm"
                      value={style.verticalAlignment ?? 'bottom'}
                      onChange={(event) => setStyle({ verticalAlignment: event.target.value as CellStyle['verticalAlignment'] })}
                    >
                      <option value="top">Top</option>
                      <option value="middle">Middle</option>
                      <option value="bottom">Bottom</option>
                    </Select>
                  </Stack>
                  <label className="flex items-center gap-2 text-xs text-slate-700">
                    <input
                      checked={Boolean(style.wrapText)}
                      type="checkbox"
                      onChange={(event) => setStyle({ wrapText: event.target.checked })}
                    />
                    Wrap text
                  </label>
                  <Stack gap="xs">
                    <Text size="xs" tone="subtle">Text rotation (degrees)</Text>
                    <TextInput
                      aria-label="Text rotation"
                      type="number"
                      value={String(style.textRotate ?? 0)}
                      onChange={(event) => setStyle({ textRotate: Number(event.target.value) || 0 })}
                    />
                  </Stack>
                </Stack>
              ) : null}

              {tab === 'font' ? (
                <Stack gap="md">
                  <Stack gap="xs">
                    <Text size="xs" tone="subtle">Font family</Text>
                    <Select
                      sizeVariant="sm"
                      value={style.fontFamily ?? 'Segoe UI'}
                      onChange={(event) => setStyle({ fontFamily: event.target.value })}
                    >
                      <option value="Segoe UI">Segoe UI</option>
                      <option value="Arial">Arial</option>
                      <option value="Calibri">Calibri</option>
                      <option value="Times New Roman">Times New Roman</option>
                      <option value="Courier New">Courier New</option>
                    </Select>
                  </Stack>
                  <Stack gap="xs">
                    <Text size="xs" tone="subtle">Font size</Text>
                    <TextInput
                      aria-label="Font size"
                      type="number"
                      value={String(style.fontSize ?? 11)}
                      onChange={(event) => setStyle({ fontSize: Math.max(8, Number(event.target.value) || 11) })}
                    />
                  </Stack>
                  <Inline gap="sm" className="flex-wrap">
                    <label className="flex items-center gap-1 text-xs">
                      <input checked={Boolean(style.bold)} type="checkbox" onChange={(event) => setStyle({ bold: event.target.checked })} />
                      Bold
                    </label>
                    <label className="flex items-center gap-1 text-xs">
                      <input checked={Boolean(style.italic)} type="checkbox" onChange={(event) => setStyle({ italic: event.target.checked })} />
                      Italic
                    </label>
                    <label className="flex items-center gap-1 text-xs">
                      <input checked={Boolean(style.underline)} type="checkbox" onChange={(event) => setStyle({ underline: event.target.checked })} />
                      Underline
                    </label>
                    <label className="flex items-center gap-1 text-xs">
                      <input checked={Boolean(style.strikethrough)} type="checkbox" onChange={(event) => setStyle({ strikethrough: event.target.checked })} />
                      Strikethrough
                    </label>
                  </Inline>
                  <Stack gap="xs">
                    <Text size="xs" tone="subtle">Font color</Text>
                    <ColorPicker
                      color={style.textColor ?? '#0f172a'}
                      onChange={(color) => setStyle({ textColor: color })}
                    />
                  </Stack>
                </Stack>
              ) : null}

              {tab === 'border' ? (
                <Stack gap="md">
                  <Text size="xs" tone="subtle">Presets</Text>
                  <Inline gap="sm" className="flex-wrap">
                    <Button size="sm" variant="ghost" onClick={() => setStyle({ borders: borderSide({ style: 'thin', color: '#334155' }) })}>
                      All borders
                    </Button>
                    <Button size="sm" variant="ghost" onClick={() => setStyle({ borders: borderSide({ style: 'medium', color: '#0f172a' }) })}>
                      Outline
                    </Button>
                    <Button size="sm" variant="ghost" onClick={() => setStyle({ borders: undefined })}>
                      No border
                    </Button>
                  </Inline>
                  <Stack gap="xs">
                    <Text size="xs" tone="subtle">Border color</Text>
                    <ColorPicker
                      color={style.borders?.top?.color ?? '#334155'}
                      onChange={(color) => {
                        const side = { style: 'thin' as const, color };
                        setStyle({ borders: borderSide(side) });
                      }}
                    />
                  </Stack>
                </Stack>
              ) : null}

              {tab === 'fill' ? (
                <Stack gap="md">
                  <Text size="xs" tone="subtle">Background color</Text>
                  <ColorPicker
                    color={style.background ?? '#ffffff'}
                    onChange={(color) => setStyle({ background: color })}
                  />
                </Stack>
              ) : null}

              {tab === 'protection' ? (
                <Stack gap="md">
                  <Text size="xs" tone="subtle">
                    Cell protection is applied at the sheet level. Use Review → Protect Sheet for locking cells.
                  </Text>
                </Stack>
              ) : null}
            </Box>
          </Inline>
          <Inline gap="sm" className="justify-end border-t border-slate-200 px-4 py-3">
            <Button size="sm" variant="ghost" onClick={onClose}>Cancel</Button>
            <Button
              size="sm"
              variant="primary"
              data-testid="format-cells-apply"
              onClick={() => {
                onApply(draft);
                onClose();
              }}
            >
              OK
            </Button>
          </Inline>
        </Stack>
      </Box>
    </Box>
  );
}
