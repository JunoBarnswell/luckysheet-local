import React, { useEffect, useState } from 'react';
import type { CellBorders, CellStyle } from '@react-sheets/core-model';
import { pixelsToPoints, pointsToPixels } from '@react-sheets/exchange-excel-ooxml';
import { Box, Button, CheckToggle, ColorPicker, Dialog, Inline, Select, Stack, Text, TextInput } from '@react-sheets/ui-system';
import type { Locale } from '../../i18n';
import { homeText, resolveHomeLocale, type HomeUiTextKey } from '../home/home-localization';

export interface FormatCellsDraft {
  numberFormat: string;
  style: Partial<CellStyle>;
}

export interface FormatCellsDialogProps {
  open: boolean;
  initial: FormatCellsDraft;
  locale?: Locale;
  onClose: () => void;
  onApply: (draft: FormatCellsDraft) => void;
}

type FormatTab = 'number' | 'alignment' | 'font' | 'border' | 'fill' | 'protection';

const TABS: Array<{ id: FormatTab; labelKey: HomeUiTextKey }> = [
  { id: 'number', labelKey: 'number' },
  { id: 'alignment', labelKey: 'alignment' },
  { id: 'font', labelKey: 'font' },
  { id: 'border', labelKey: 'border' },
  { id: 'fill', labelKey: 'fill' },
  { id: 'protection', labelKey: 'protection' },
];

const NUMBER_PRESETS: Array<{ labelKey: HomeUiTextKey; value: string }> = [
  { labelKey: 'numberPresetGeneral', value: 'general' },
  { labelKey: 'numberPresetNumber', value: '0.00' },
  { labelKey: 'numberPresetComma', value: '#,##0' },
  { labelKey: 'numberPresetCurrency', value: '$#,##0.00' },
  { labelKey: 'numberPresetPercent', value: '0%' },
  { labelKey: 'numberPresetScientific', value: '0.00E+00' },
  { labelKey: 'numberPresetDate', value: 'yyyy-mm-dd' },
  { labelKey: 'numberPresetTime', value: 'hh:mm' },
  { labelKey: 'numberPresetText', value: '@' },
];

function borderSide(style: CellBorders['top']): CellBorders {
  return {
    top: style,
    right: style,
    bottom: style,
    left: style,
  };
}

export function FormatCellsDialog({ open, initial, locale, onClose, onApply }: FormatCellsDialogProps): React.ReactElement | null {
  const [tab, setTab] = useState<FormatTab>('number');
  const [draft, setDraft] = useState<FormatCellsDraft>(initial);

  useEffect(() => {
    if (open) {
      setDraft(initial);
      setTab('number');
    }
  }, [open, initial]);

  const style = draft.style;
  const setStyle = (patch: Partial<CellStyle>) => setDraft((prev) => ({ ...prev, style: { ...prev.style, ...patch } }));
  const activeLocale = resolveHomeLocale(locale);

  return (
    <Dialog
      open={open}
      onClose={onClose}
      title={homeText(activeLocale, 'formatCells')}
      description={homeText(activeLocale, 'formatCellsDescription')}
      closeLabel={homeText(activeLocale, 'close')}
      testId="format-cells-dialog"
      maxWidth="lg"
      bodyClassName="p-0"
      footer={(
        <>
          <Button size="sm" variant="ghost" onClick={onClose}>{homeText(activeLocale, 'cancel')}</Button>
          <Button
            size="sm"
            variant="primary"
            data-testid="format-cells-apply"
            onClick={() => {
              onApply(draft);
              onClose();
            }}
          >
            {homeText(activeLocale, 'ok')}
          </Button>
        </>
      )}
    >
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
                  {homeText(activeLocale, entry.labelKey)}
                </Button>
              ))}
            </Stack>
            <Box className="min-w-0 flex-1 p-4">
              {tab === 'number' ? (
                <Stack gap="md">
                  <Text size="xs" tone="subtle">{homeText(activeLocale, 'category')}</Text>
                  <Select
                    sizeVariant="sm"
                    value={draft.numberFormat}
                    onChange={(event) => setDraft((prev) => ({ ...prev, numberFormat: event.target.value }))}
                  >
                    {NUMBER_PRESETS.map((preset) => (
                      <option key={preset.value} value={preset.value}>{homeText(activeLocale, preset.labelKey)}</option>
                    ))}
                  </Select>
                  <Stack gap="xs">
                    <Text size="xs" tone="subtle">{homeText(activeLocale, 'customFormat')}</Text>
                    <TextInput
                      aria-label={homeText(activeLocale, 'customFormat')}
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
                    <Text size="xs" tone="subtle">{homeText(activeLocale, 'horizontal')}</Text>
                    <Select
                      sizeVariant="sm"
                      value={style.horizontalAlignment ?? 'left'}
                      onChange={(event) => setStyle({ horizontalAlignment: event.target.value as CellStyle['horizontalAlignment'] })}
                    >
                      <option value="left">{homeText(activeLocale, 'left')}</option>
                      <option value="center">{homeText(activeLocale, 'center')}</option>
                      <option value="right">{homeText(activeLocale, 'right')}</option>
                    </Select>
                  </Stack>
                  <Stack gap="xs">
                    <Text size="xs" tone="subtle">{homeText(activeLocale, 'vertical')}</Text>
                    <Select
                      sizeVariant="sm"
                      value={style.verticalAlignment ?? 'bottom'}
                      onChange={(event) => setStyle({ verticalAlignment: event.target.value as CellStyle['verticalAlignment'] })}
                    >
                      <option value="top">{homeText(activeLocale, 'top')}</option>
                      <option value="middle">{homeText(activeLocale, 'middle')}</option>
                      <option value="bottom">{homeText(activeLocale, 'bottom')}</option>
                    </Select>
                  </Stack>
                  <CheckToggle
                    checked={Boolean(style.wrapText)}
                    label={homeText(activeLocale, 'wrapText')}
                    onChange={(event) => setStyle({ wrapText: event.target.checked })}
                  />
                  <Stack gap="xs">
                    <Text size="xs" tone="subtle">{homeText(activeLocale, 'rotation')}</Text>
                    <TextInput
                      aria-label={homeText(activeLocale, 'rotation')}
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
                    <Text size="xs" tone="subtle">{homeText(activeLocale, 'fontFamily')}</Text>
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
                    <Text size="xs" tone="subtle">{homeText(activeLocale, 'fontSize')}</Text>
                    <TextInput
                      aria-label={homeText(activeLocale, 'fontSize')}
                      type="number"
                      value={String(Number(pixelsToPoints(style.fontSizePx ?? pointsToPixels(11)).toFixed(2)))}
                      onChange={(event) => setStyle({ fontSizePx: pointsToPixels(Math.max(1, Number(event.target.value) || 11)) })}
                    />
                  </Stack>
                  <Inline gap="sm" className="flex-wrap">
                    <CheckToggle checked={Boolean(style.bold)} label={homeText(activeLocale, 'bold')} onChange={(event) => setStyle({ bold: event.target.checked })} />
                    <CheckToggle checked={Boolean(style.italic)} label={homeText(activeLocale, 'italic')} onChange={(event) => setStyle({ italic: event.target.checked })} />
                    <CheckToggle checked={Boolean(style.underline)} label={homeText(activeLocale, 'underline')} onChange={(event) => setStyle({ underline: event.target.checked })} />
                    <CheckToggle checked={Boolean(style.strikethrough)} label={homeText(activeLocale, 'strikethrough')} onChange={(event) => setStyle({ strikethrough: event.target.checked })} />
                  </Inline>
                  <Stack gap="xs">
                    <Text size="xs" tone="subtle">{homeText(activeLocale, 'textColor')}</Text>
                    <ColorPicker
                      color={style.textColor ?? '#0f172a'}
                      onChange={(color) => setStyle({ textColor: color })}
                    />
                  </Stack>
                </Stack>
              ) : null}

              {tab === 'border' ? (
                <Stack gap="md">
                  <Inline gap="sm" className="flex-wrap">
                    <Button size="sm" variant="ghost" onClick={() => setStyle({ borders: borderSide({ style: 'thin', color: '#334155' }) })}>
                      {homeText(activeLocale, 'allBorders')}
                    </Button>
                    <Button size="sm" variant="ghost" onClick={() => setStyle({ borders: borderSide({ style: 'medium', color: '#0f172a' }) })}>
                      {homeText(activeLocale, 'outlineBorders')}
                    </Button>
                    <Button size="sm" variant="ghost" onClick={() => setStyle({ borders: undefined })}>
                      {homeText(activeLocale, 'noBorder')}
                    </Button>
                  </Inline>
                  <Stack gap="xs">
                    <Text size="xs" tone="subtle">{homeText(activeLocale, 'borderColor')}</Text>
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
                  <Text size="xs" tone="subtle">{homeText(activeLocale, 'fillBackground')}</Text>
                  <ColorPicker
                    color={style.background ?? '#ffffff'}
                    onChange={(color) => setStyle({ background: color })}
                  />
                </Stack>
              ) : null}

              {tab === 'protection' ? (
                <Stack gap="md">
                  <Text size="xs" tone="subtle">{homeText(activeLocale, 'protectionHint')}</Text>
                  <CheckToggle
                    checked={style.locked !== false}
                    label={homeText(activeLocale, 'protectionLocked')}
                    onChange={(event) => setStyle({ locked: event.target.checked })}
                  />
                  <CheckToggle
                    checked={Boolean(style.formulaHidden)}
                    label={homeText(activeLocale, 'protectionHidden')}
                    onChange={(event) => setStyle({ formulaHidden: event.target.checked })}
                  />
                </Stack>
              ) : null}
            </Box>
      </Inline>
    </Dialog>
  );
}
