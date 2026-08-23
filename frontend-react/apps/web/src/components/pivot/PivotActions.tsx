import { useState } from 'react';
import { Box, Button, Inline, Select, Text } from '@react-sheets/ui-system';
import type { PivotFieldDefinition } from '@react-sheets/core-model';
import type { PivotPanelCallbacks } from './pivot-contract';

export interface PivotActionsProps {
  layout: 'compact' | 'outline' | 'tabular';
  slicerFieldIds: readonly string[];
  fields: readonly PivotFieldDefinition[];
  callbacks: PivotPanelCallbacks;
  disabled?: boolean;
}

export function PivotActions({ callbacks, disabled = false, fields, layout, slicerFieldIds }: PivotActionsProps) {
  const [controlField, setControlField] = useState(fields[0]?.fieldId ?? '');
  const dateFields = fields.filter((field) => field.dataType === 'date');
  const timelineField = dateFields.find((field) => field.fieldId === controlField)?.fieldId ?? dateFields[0]?.fieldId;
  return (
    <Box as="section" aria-label="Pivot configuration actions" className="border-t border-line/80 pt-3">
      <Text size="xs" weight="semibold" tone="muted" className="mb-2 block">CONFIGURE</Text>
      <Inline gap="xs" className="flex-wrap">
        <Button disabled={disabled} icon="refresh" size="xs" variant="outline" onClick={callbacks.onRefresh}>Refresh</Button>
        <Select aria-label="Pivot layout" disabled={disabled} sizeVariant="sm" value={layout} onChange={(event) => callbacks.onLayoutChange(event.target.value as typeof layout)}><option value="compact">Compact</option><option value="outline">Outline</option><option value="tabular">Tabular</option></Select>
        <Select aria-label="Pivot control field" disabled={disabled || fields.length === 0} sizeVariant="sm" value={controlField} onChange={(event) => setControlField(event.target.value)}>{fields.filter((field) => field.fieldId).map((field) => <option key={field.fieldId} value={field.fieldId}>{field.name}</option>)}</Select>
        <Button disabled={disabled || !controlField} icon="sliders" size="xs" variant="outline" onClick={() => callbacks.onSlicerChange(controlField, !slicerFieldIds.includes(controlField))}>Slicer</Button>
        {timelineField ? <Button disabled={disabled} icon="history" size="xs" variant="outline" onClick={() => callbacks.onTimelineChange(timelineField)}>Timeline</Button> : null}
        <Button disabled={disabled} icon="chart" size="xs" variant="outline" onClick={() => callbacks.onPivotChartChange({ type: 'column', title: 'Pivot chart' })}>Pivot chart</Button>
      </Inline>
    </Box>
  );
}
