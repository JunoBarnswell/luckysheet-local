import { useState } from 'react';
import { Box, Button, CheckToggle, Inline, Select, Stack, Text } from '@react-sheets/ui-system';
import type { PivotFieldDefinition } from '@react-sheets/core-model';
import type { PivotConnectionOption, PivotPanelCallbacks } from './pivot-contract';
import type { Locale } from '../../i18n';
import { pivotText } from './pivot-localization';

export interface PivotActionsProps {
  layout: 'compact' | 'outline' | 'tabular';
  showButtons: boolean;
  slicerFieldIds: readonly string[];
  fields: readonly PivotFieldDefinition[];
  callbacks: PivotPanelCallbacks;
  disabled?: boolean;
  locale: Locale;
  connectionControlId?: string;
  connectionOptions?: readonly PivotConnectionOption[];
}

export function PivotActions({ callbacks, connectionControlId, connectionOptions = [], disabled = false, fields, layout, locale, showButtons, slicerFieldIds }: PivotActionsProps) {
  const [controlField, setControlField] = useState(fields[0]?.fieldId ?? '');
  const dateFields = fields.filter((field) => field.dataType === 'date');
  const timelineField = dateFields.find((field) => field.fieldId === controlField)?.fieldId ?? dateFields[0]?.fieldId;
  return (
    <Box as="section" aria-label={pivotText(locale, 'configurationActions')} className="border-t border-line/80 pt-3">
      <Text size="xs" weight="semibold" tone="muted" className="mb-2 block">{pivotText(locale, 'configure')}</Text>
      <Inline gap="xs" className="flex-wrap">
        <Button disabled={disabled} icon="refresh" size="xs" variant="outline" onClick={callbacks.onRefresh}>{pivotText(locale, 'refresh')}</Button>
        <Select aria-label={pivotText(locale, 'view')} disabled={disabled} sizeVariant="sm" value={layout} onChange={(event) => callbacks.onLayoutChange(event.target.value as typeof layout)}><option value="compact">{pivotText(locale, 'compact')}</option><option value="outline">{pivotText(locale, 'outline')}</option><option value="tabular">{pivotText(locale, 'tabular')}</option></Select>
        <Select aria-label={pivotText(locale, 'controlField')} disabled={disabled || fields.length === 0} sizeVariant="sm" value={controlField} onChange={(event) => setControlField(event.target.value)}>{fields.filter((field) => field.fieldId).map((field) => <option key={field.fieldId} value={field.fieldId}>{field.name}</option>)}</Select>
        {callbacks.onExpansionCommand ? <>
          <Button disabled={disabled || !controlField} icon="chevron-down" size="xs" variant="outline" onClick={() => callbacks.onExpansionCommand?.({ kind: 'expand-field', fieldId: controlField })}>{pivotText(locale, 'expandField')}</Button>
          <Button disabled={disabled || !controlField} icon="chevron-up" size="xs" variant="outline" onClick={() => callbacks.onExpansionCommand?.({ kind: 'collapse-field', fieldId: controlField })}>{pivotText(locale, 'collapseField')}</Button>
          <Button disabled={disabled} icon="plus" size="xs" variant="outline" onClick={() => callbacks.onExpansionCommand?.({ kind: 'toggle-buttons', showButtons: !showButtons })}>{showButtons ? pivotText(locale, 'hideButtons') : pivotText(locale, 'showButtons')}</Button>
        </> : null}
        <Button disabled={disabled || !controlField} icon="sliders" size="xs" variant="outline" onClick={() => callbacks.onSlicerChange(controlField, !slicerFieldIds.includes(controlField))}>{pivotText(locale, 'slicer')}</Button>
        {timelineField ? <Button disabled={disabled} icon="history" size="xs" variant="outline" onClick={() => callbacks.onTimelineChange(timelineField)}>{pivotText(locale, 'timeline')}</Button> : null}
        <Button disabled={disabled} icon="chart" size="xs" variant="outline" onClick={() => callbacks.onPivotChartChange({ type: 'column', title: pivotText(locale, 'pivotChart') })}>{pivotText(locale, 'pivotChart')}</Button>
      </Inline>
      {connectionControlId ? <Box className="mt-3 border-t border-line/60 pt-3" aria-label={pivotText(locale, 'reportConnections')}>
        <Text size="xs" weight="semibold" tone="muted" className="mb-2 block">{pivotText(locale, 'reportConnections')}</Text>
        {connectionOptions.length === 0
          ? <Text size="xs" tone="subtle">{pivotText(locale, 'noCompatiblePivots')}</Text>
          : <Stack gap="xs">{connectionOptions.map((option) => <CheckToggle key={option.pivotId} aria-label={`${pivotText(locale, 'reportConnections')}: ${option.label}`} checked={option.selected} disabled={disabled} label={option.label} onChange={(event) => {
            const next = connectionOptions.filter((candidate) => candidate.pivotId === option.pivotId ? event.target.checked : candidate.selected).map((candidate) => ({ pivotId: candidate.pivotId, sourceKey: candidate.sourceKey, fieldId: candidate.fieldId }));
            callbacks.onConnectionsChange?.(connectionControlId, next);
          }} />)}</Stack>}
      </Box> : null}
    </Box>
  );
}
