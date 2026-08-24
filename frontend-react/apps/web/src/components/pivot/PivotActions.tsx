import { useState } from 'react';
import { Box, Button, Inline, Select, Text } from '@react-sheets/ui-system';
import type { PivotFieldDefinition } from '@react-sheets/core-model';
import type { PivotPanelCallbacks } from './pivot-contract';
import type { Locale } from '../../i18n';
import { pivotText } from './pivot-localization';

export interface PivotActionsProps {
  layout: 'compact' | 'outline' | 'tabular';
  slicerFieldIds: readonly string[];
  fields: readonly PivotFieldDefinition[];
  callbacks: PivotPanelCallbacks;
  disabled?: boolean;
  locale: Locale;
}

export function PivotActions({ callbacks, disabled = false, fields, layout, locale, slicerFieldIds }: PivotActionsProps) {
  const [controlField, setControlField] = useState(fields[0]?.fieldId ?? '');
  const dateFields = fields.filter((field) => field.dataType === 'date');
  const timelineField = dateFields.find((field) => field.fieldId === controlField)?.fieldId ?? dateFields[0]?.fieldId;
  return (
    <Box as="section" aria-label="Pivot configuration actions" className="border-t border-line/80 pt-3">
      <Text size="xs" weight="semibold" tone="muted" className="mb-2 block">{pivotText(locale, 'configure')}</Text>
      <Inline gap="xs" className="flex-wrap">
        <Button disabled={disabled} icon="refresh" size="xs" variant="outline" onClick={callbacks.onRefresh}>{pivotText(locale, 'refresh')}</Button>
        <Select aria-label={pivotText(locale, 'view')} disabled={disabled} sizeVariant="sm" value={layout} onChange={(event) => callbacks.onLayoutChange(event.target.value as typeof layout)}><option value="compact">{pivotText(locale, 'compact')}</option><option value="outline">{pivotText(locale, 'outline')}</option><option value="tabular">{pivotText(locale, 'tabular')}</option></Select>
        <Select aria-label={pivotText(locale, 'controlField')} disabled={disabled || fields.length === 0} sizeVariant="sm" value={controlField} onChange={(event) => setControlField(event.target.value)}>{fields.filter((field) => field.fieldId).map((field) => <option key={field.fieldId} value={field.fieldId}>{field.name}</option>)}</Select>
        <Button disabled={disabled || !controlField} icon="sliders" size="xs" variant="outline" onClick={() => callbacks.onSlicerChange(controlField, !slicerFieldIds.includes(controlField))}>{pivotText(locale, 'slicer')}</Button>
        {timelineField ? <Button disabled={disabled} icon="history" size="xs" variant="outline" onClick={() => callbacks.onTimelineChange(timelineField)}>{pivotText(locale, 'timeline')}</Button> : null}
        <Button disabled={disabled} icon="chart" size="xs" variant="outline" onClick={() => callbacks.onPivotChartChange({ type: 'column', title: pivotText(locale, 'pivotChart') })}>{pivotText(locale, 'pivotChart')}</Button>
      </Inline>
    </Box>
  );
}
