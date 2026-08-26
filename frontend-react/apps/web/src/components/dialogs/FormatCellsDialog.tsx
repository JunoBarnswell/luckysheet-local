import React, { useEffect, useState } from 'react';
import type { BorderLine, BorderPlacement, CellStyle } from '@react-sheets/core-model';
import { pixelsToPoints, pointsToPixels } from '@react-sheets/exchange-excel-ooxml';
import { Box, Button, CheckToggle, ColorPicker, Dialog, Inline, Select, Stack, Text, TextInput } from '@react-sheets/ui-system';
import type { Locale } from '../../i18n';
import { homeText, resolveHomeLocale, type HomeUiTextKey } from '../home/home-localization';
import { FontFamilyControl } from '../FontFamilyControl';

export interface FormatCellsDraft {
  numberFormat: string;
  style: Partial<CellStyle>;
  /** UI-only mixed-state marker; it is not submitted to the workbook model. */
  mixedFontFamily?: boolean;
  border?: { placement: BorderPlacement; line?: BorderLine };
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

const BORDER_ACTIONS: Array<{ placement: BorderPlacement; labelKey: HomeUiTextKey; line?: BorderLine }> = [
  { placement: 'all', labelKey: 'allBorders', line: { style: 'thin', color: '#334155' } },
  { placement: 'outside', labelKey: 'borderOutside', line: { style: 'thin', color: '#334155' } },
  { placement: 'thick-outside', labelKey: 'borderThickOutside', line: { style: 'thick', color: '#334155' } },
  { placement: 'top', labelKey: 'borderTop', line: { style: 'thin', color: '#334155' } },
  { placement: 'bottom', labelKey: 'borderBottom', line: { style: 'thin', color: '#334155' } },
  { placement: 'left', labelKey: 'borderLeft', line: { style: 'thin', color: '#334155' } },
  { placement: 'right', labelKey: 'borderRight', line: { style: 'thin', color: '#334155' } },
  { placement: 'inside-horizontal', labelKey: 'borderInsideHorizontal', line: { style: 'thin', color: '#334155' } },
  { placement: 'inside-vertical', labelKey: 'borderInsideVertical', line: { style: 'thin', color: '#334155' } },
  { placement: 'none', labelKey: 'noBorder' },
];

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
  const setStyle = (patch: Partial<CellStyle>) => setDraft((prev) => ({
    ...prev,
    style: { ...prev.style, ...patch },
    ...(patch.fontFamily === undefined ? {} : { mixedFontFamily: false }),
  }));
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
              const style = { ...draft.style };
              delete style.borders;
              onApply({ ...draft, style });
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
                      value={style.horizontalAlignment ?? 'general'}
                      onChange={(event) => setStyle({ horizontalAlignment: event.target.value as CellStyle['horizontalAlignment'] })}
                    >
                      <option value="general">{homeText(activeLocale, 'general')}</option>
                      <option value="left">{homeText(activeLocale, 'left')}</option>
                      <option value="center">{homeText(activeLocale, 'center')}</option>
                      <option value="right">{homeText(activeLocale, 'right')}</option>
                      <option value="centerContinuous">{homeText(activeLocale, 'centerContinuous')}</option>
                      <option value="justify">{homeText(activeLocale, 'justify')}</option>
                      <option value="distributed">{homeText(activeLocale, 'distributed')}</option>
                      <option value="fill">{homeText(activeLocale, 'fillAlignment')}</option>
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
                      <option value="justify">{homeText(activeLocale, 'justify')}</option>
                      <option value="distributed">{homeText(activeLocale, 'distributed')}</option>
                    </Select>
                  </Stack>
                  <CheckToggle
                    checked={Boolean(style.wrapText)}
                    label={homeText(activeLocale, 'wrapText')}
                    onChange={(event) => setStyle({ wrapText: event.target.checked })}
                  />
                  <CheckToggle
                    checked={Boolean(style.shrinkToFit)}
                    label={homeText(activeLocale, 'shrinkToFit')}
                    onChange={(event) => setStyle({ shrinkToFit: event.target.checked })}
                  />
                  <Stack gap="xs">
                    <Text size="xs" tone="subtle">{homeText(activeLocale, 'indent')}</Text>
                    <TextInput
                      aria-label={homeText(activeLocale, 'indent')}
                      type="number"
                      min={0}
                      max={250}
                      value={String(style.indent ?? 0)}
                      onChange={(event) => setStyle({ indent: Math.min(250, Math.max(0, Number(event.target.value) || 0)) })}
                    />
                  </Stack>
                  <Stack gap="xs">
                    <Text size="xs" tone="subtle">{homeText(activeLocale, 'textOrientation')}</Text>
                    <Select
                      sizeVariant="sm"
                      value={style.textOrientation ?? 'horizontal'}
                      onChange={(event) => setStyle({ textOrientation: event.target.value as CellStyle['textOrientation'], ...(event.target.value === 'horizontal' ? { textRotate: 0 } : {}) })}
                    >
                      <option value="horizontal">Horizontal</option>
                      <option value="rotateUp">Rotate up</option>
                      <option value="rotateDown">Rotate down</option>
                      <option value="stacked">Stacked</option>
                    </Select>
                  </Stack>
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
                    <FontFamilyControl
                      value={style.fontFamily}
                      fallbackValue="Segoe UI"
                      mixed={Boolean(draft.mixedFontFamily)}
                      mixedPlaceholder={homeText(activeLocale, 'mixed')}
                      label={homeText(activeLocale, 'fontFamily')}
                      testId="format-font-family"
                      onCommit={(fontFamily) => setStyle({ fontFamily })}
                    />
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
                    {BORDER_ACTIONS.map((action) => (
                      <Button
                        key={action.placement}
                        size="sm"
                        variant="ghost"
                        data-testid={`format-border-${action.placement}`}
                        onClick={() => setDraft((prev) => ({ ...prev, border: action.line ? { placement: action.placement, line: { ...action.line, color: prev.border?.line?.color ?? action.line.color } } : { placement: action.placement } }))}
                      >
                        {homeText(activeLocale, action.labelKey)}
                      </Button>
                    ))}
                  </Inline>
                  <Stack gap="xs">
                    <Text size="xs" tone="subtle">{homeText(activeLocale, 'borderColor')}</Text>
                    <ColorPicker
                      color={draft.border?.line?.color ?? style.borders?.top?.color ?? '#334155'}
                      onChange={(color) => {
                        setDraft((prev) => ({
                          ...prev,
                          border: {
                            placement: prev.border?.placement ?? 'all',
                            line: { style: prev.border?.line?.style ?? 'thin', color },
                          },
                        }));
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
