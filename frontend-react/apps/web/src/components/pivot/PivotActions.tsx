import { useState } from 'react';
import { Box, Button, Inline, Select, Text } from '@react-sheets/ui-system';
import type { PivotFieldDefinition, PivotPanelCallbacks } from './types';

export interface PivotActionsProps {
  layout: 'compact' | 'outline' | 'tabular';
  expandedFieldIds: readonly string[];
  slicers: readonly string[];
  fields: readonly PivotFieldDefinition[];
  callbacks: PivotPanelCallbacks;
  disabled?: boolean;
}

export function PivotActions({ callbacks, disabled = false, expandedFieldIds, fields, layout, slicers }: PivotActionsProps) {
  const [controlField, setControlField] = useState(fields[0]?.id ?? '');
  const dateFields = fields.filter((field) => field.type === 'date');
  const timelineField = dateFields.find((field) => field.id === controlField)?.id ?? dateFields[0]?.id;
  return (
    <Box as="section" aria-label="Pivot configuration actions" className="border-t border-line/80 pt-3">
      <Text size="xs" weight="semibold" tone="muted" className="mb-2 block">CONFIGURE</Text>
      <Inline gap="xs" className="flex-wrap">
        <Button disabled={disabled} icon="refresh" size="xs" variant="outline" onClick={callbacks.onRefresh}>Refresh</Button>
        <Select aria-label="Pivot layout" disabled={disabled} sizeVariant="sm" value={layout} onChange={(event) => callbacks.onLayoutChange(event.target.value as typeof layout)}><option value="compact">Compact</option><option value="outline">Outline</option><option value="tabular">Tabular</option></Select>
        <Button disabled={disabled || !controlField} icon="eye" size="xs" variant="outline" onClick={() => callbacks.onExpandedChange(controlField, !expandedFieldIds.includes(controlField))}>Expand state</Button>
        <Select aria-label="Pivot control field" disabled={disabled || fields.length === 0} sizeVariant="sm" value={controlField} onChange={(event) => setControlField(event.target.value)}>{fields.map((field) => <option key={field.id} value={field.id}>{field.label}</option>)}</Select>
        <Button disabled={disabled || !controlField} icon="sliders" size="xs" variant="outline" onClick={() => callbacks.onSlicerChange(controlField, !slicers.includes(controlField))}>Slicer</Button>
        {timelineField ? <Button disabled={disabled} icon="history" size="xs" variant="outline" onClick={() => callbacks.onTimelineChange(timelineField)}>Timeline</Button> : null}
        <Button disabled={disabled} icon="chart" size="xs" variant="outline" onClick={() => callbacks.onPivotChartChange({ type: 'column', title: 'Pivot chart' })}>Pivot chart</Button>
      </Inline>
    </Box>
  );
}
