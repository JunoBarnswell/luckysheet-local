import { Box, CheckToggle, Inline, Select, Stack, Text, TextInput } from '@react-sheets/ui-system';
import {
  DEFAULT_PIVOT_DISPLAY_OPTIONS,
  DEFAULT_PIVOT_STYLE_OPTIONS,
  normalizePivotDisplayOptions,
  type PivotDisplayOptions,
  type PivotPresentation,
  type PivotRefreshPolicy,
} from '@react-sheets/core-model';
import type { Locale } from '../../i18n';
import { pivotText } from './pivot-localization';

export interface PivotFormatOptionsProps {
  locale: Locale;
  disabled?: boolean;
  presentation?: PivotPresentation;
  refreshPolicy: PivotRefreshPolicy;
  onPresentationChange?: (presentation: PivotPresentation) => void;
  onDisplayOptionsChange?: (displayOptions: PivotDisplayOptions) => void;
  onRefreshPolicyChange?: (refreshPolicy: PivotRefreshPolicy) => void;
}

export function PivotFormatOptions({ disabled = false, locale, onDisplayOptionsChange, onPresentationChange, onRefreshPolicyChange, presentation: input, refreshPolicy }: PivotFormatOptionsProps) {
  const presentation: PivotPresentation = {
    ...(input?.styleName ? { styleName: input.styleName } : {}),
    styleOptions: { ...DEFAULT_PIVOT_STYLE_OPTIONS, ...(input?.styleOptions ?? {}) },
    displayOptions: normalizePivotDisplayOptions(input?.displayOptions),
  };
  const updatePresentation = (patch: Partial<PivotPresentation['styleOptions']> & { styleName?: string }) => {
    const { styleName, ...options } = patch;
    onPresentationChange?.({ ...(styleName ?? presentation.styleName ? { styleName: styleName ?? presentation.styleName } : {}), styleOptions: { ...presentation.styleOptions, ...options }, displayOptions: presentation.displayOptions });
  };
  const updateDisplayOptions = (patch: Partial<PivotDisplayOptions>) => onDisplayOptionsChange?.(normalizePivotDisplayOptions({ ...presentation.displayOptions, ...patch }));

  return <Box className="max-h-[32rem] w-[22rem] overflow-auto p-3"><Stack gap="sm">
    <Stack gap="xs">
      <Text size="sm" weight="medium">{pivotText(locale, 'pivotStyle')}</Text>
      <Select aria-label={pivotText(locale, 'pivotStyle')} sizeVariant="sm" value={presentation.styleName ?? 'PivotStyleLight16'} disabled={disabled} onChange={(event) => updatePresentation({ styleName: event.target.value })}>
        <option value="PivotStyleLight16">{pivotText(locale, 'styleLight')}</option>
        <option value="PivotStyleMedium4">{pivotText(locale, 'styleMedium')}</option>
        <option value="PivotStyleDark2">{pivotText(locale, 'styleDark')}</option>
      </Select>
      <Inline gap="sm" className="flex-wrap">
        {(['showRowHeaders', 'showColumnHeaders', 'showRowStripes', 'showColumnStripes', 'showLastColumn'] as const).map((option) => <CheckToggle key={option} label={pivotText(locale, option === 'showRowHeaders' ? 'rowHeaders' : option === 'showColumnHeaders' ? 'columnHeaders' : option === 'showRowStripes' ? 'bandedRows' : option === 'showColumnStripes' ? 'bandedColumns' : 'lastColumn')} checked={presentation.styleOptions[option]} disabled={disabled} onChange={(event) => updatePresentation({ [option]: event.target.checked })} />)}
      </Inline>
    </Stack>
    <Stack gap="xs" className="border-t border-slate-200 pt-2">
      <Text size="sm" weight="medium">{pivotText(locale, 'pivotOptions')}</Text>
      <Text size="xs" tone="muted">{pivotText(locale, 'layoutFormat')}</Text>
      <CheckToggle label={pivotText(locale, 'fillEmptyCells')} checked={presentation.displayOptions?.fillEmptyCells ?? DEFAULT_PIVOT_DISPLAY_OPTIONS.fillEmptyCells} disabled={disabled} onChange={(event) => updateDisplayOptions({ fillEmptyCells: event.target.checked })} />
      <TextInput aria-label={pivotText(locale, 'emptyCellText')} value={presentation.displayOptions?.emptyCellText ?? ''} disabled={disabled || !(presentation.displayOptions?.fillEmptyCells ?? false)} placeholder={pivotText(locale, 'emptyCellText')} onChange={(event) => updateDisplayOptions({ emptyCellText: event.target.value })} />
      <CheckToggle label={pivotText(locale, 'showErrorValues')} checked={presentation.displayOptions?.showErrorValues ?? DEFAULT_PIVOT_DISPLAY_OPTIONS.showErrorValues} disabled={disabled} onChange={(event) => updateDisplayOptions({ showErrorValues: event.target.checked })} />
      <TextInput aria-label={pivotText(locale, 'errorCellText')} value={presentation.displayOptions?.errorCellText ?? ''} disabled={disabled || !(presentation.displayOptions?.showErrorValues ?? true)} placeholder={pivotText(locale, 'errorCellText')} onChange={(event) => updateDisplayOptions({ errorCellText: event.target.value })} />
      <Text size="xs" tone="muted">{pivotText(locale, 'displayOptions')}</Text>
      <CheckToggle label={pivotText(locale, 'showFieldHeaders')} checked={presentation.displayOptions?.showFieldHeaders ?? true} disabled={disabled} onChange={(event) => updateDisplayOptions({ showFieldHeaders: event.target.checked })} />
      <CheckToggle label={pivotText(locale, 'autoFitColumns')} checked={presentation.displayOptions?.autoFitColumnsOnUpdate ?? true} disabled={disabled} onChange={(event) => updateDisplayOptions({ autoFitColumnsOnUpdate: event.target.checked })} />
      <CheckToggle label={pivotText(locale, 'preserveFormatting')} checked={refreshPolicy.preserveFormatting} disabled={disabled} onChange={(event) => onRefreshPolicyChange?.({ ...refreshPolicy, preserveFormatting: event.target.checked })} />
    </Stack>
  </Stack></Box>;
}
