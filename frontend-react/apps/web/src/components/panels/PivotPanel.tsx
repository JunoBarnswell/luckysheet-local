import { Button, Inline, Panel, PanelBody, PanelFooter, PanelHeader, PanelTitle, Stack, StatePanel, Text } from '@react-sheets/ui-system';
import { useMemo, type DragEvent } from 'react';
import { PivotActions } from '../pivot/PivotActions';
import { PivotFieldArea } from '../pivot/PivotFieldArea';
import { PivotFieldCatalog } from '../pivot/PivotFieldCatalog';
import { PivotValueEditor } from '../pivot/PivotValueEditor';
import { PivotResultView } from '../pivot/PivotResultView';
import { PivotSlicer } from '../pivot/PivotSlicer';
import { PivotTimeline } from '../pivot/PivotTimeline';
import type { PivotDefinition, PivotFieldArea as Area, PivotFieldDefinition, PivotPanelCallbacks, PivotPanelSlots, PivotPanelState, PivotResult } from '../pivot/types';

export interface PivotPanelProps {
  definition?: PivotDefinition;
  fieldCatalog?: readonly PivotFieldDefinition[];
  result?: PivotResult;
  onShowDetails?: (paths: import('@react-sheets/core-model').PivotSourceRowPath[]) => void;
  state?: PivotPanelState;
  callbacks?: PivotPanelCallbacks;
  slots?: PivotPanelSlots;
  onClose?: () => void;
}

const emptyDefinition: PivotDefinition = { filters: [], columns: [], rows: [], values: [], filterSelections: {}, sort: {}, groupedFields: [], layout: 'compact', showGrandTotals: true, showSubtotals: true, expandedFieldIds: [], slicers: [] };

export function PivotPanel({ callbacks, definition = emptyDefinition, fieldCatalog = [], onClose, onShowDetails, result, slots, state }: PivotPanelProps) {
  const isDisabled = state?.disabled || state?.loading || Boolean(state?.error);
  const selectedFieldIds = useMemo(() => new Set([...definition.filters, ...definition.columns, ...definition.rows, ...definition.values.map((value) => value.fieldId)]), [definition]);

  if (state?.loading) return <Panel className="h-full border-0 bg-transparent shadow-none"><StatePanel kind="loading" description="Preparing pivot field list." /></Panel>;
  if (state?.error) return <Panel className="h-full border-0 bg-transparent shadow-none"><StatePanel kind="error" description={state.error} /></Panel>;
  if (state?.empty || fieldCatalog.length === 0) return <Panel className="h-full border-0 bg-transparent shadow-none"><StatePanel kind={state?.disabled ? 'disabled' : 'empty'} description="No pivot fields are available for this source." /></Panel>;

  const updateArea = (fieldId: string, area: Area, index: number) => callbacks?.onFieldAreaChange(fieldId, area, index);
  const handleDrop = (area: Area) => (event: DragEvent<HTMLElement>, index?: number) => { event.preventDefault(); const fieldId = event.dataTransfer.getData('application/x-pivot-field'); if (fieldId) updateArea(fieldId, area, index ?? definition[area].length); };
  const handleToggle = (fieldId: string, checked: boolean) => {
    if (checked) updateArea(fieldId, 'rows', definition.rows.length);
    else { const currentArea = (['filters', 'columns', 'rows'] as const).find((area) => definition[area].includes(fieldId)); if (currentArea) callbacks?.onRemoveField(fieldId, currentArea, definition[currentArea].indexOf(fieldId)); else { const valueIndex = definition.values.findIndex((value) => value.fieldId === fieldId); if (valueIndex >= 0) callbacks?.onRemoveField(fieldId, 'values', valueIndex); } }
  };
  const moveByKeyboard = (area: Area) => (fieldId: string, index: number, direction: -1 | 1) => updateArea(fieldId, area, index + direction);

  return (
    <Panel className="flex h-full min-h-0 flex-col border-0 bg-transparent shadow-none">
      <PanelHeader className="h-12 shrink-0 border-b border-line/80 px-4"><Inline gap="sm" className="min-w-0 justify-between"><Stack gap="none" className="min-w-0"><PanelTitle size="sm">Pivot field list</PanelTitle><Text size="xs" tone="subtle">Build, filter and configure the pivot view</Text></Stack><Inline gap="xs">{callbacks?.onCreate ? <Button disabled={isDisabled} size="xs" variant="primary" onClick={callbacks.onCreate}>Create</Button> : null}{slots?.headerActions}</Inline></Inline></PanelHeader>
      <PanelBody className="min-h-0 flex-1 overflow-auto p-4">
        <Stack gap="md">
          <PivotFieldCatalog fields={fieldCatalog} selectedFieldIds={selectedFieldIds} disabled={isDisabled} onToggle={handleToggle} onDragField={(event, field) => event.dataTransfer.setData('application/x-pivot-field', field.id)} onKeyboardAssign={(fieldId, area) => updateArea(fieldId, area, definition[area].length)} />
          <Stack gap="sm">
            {(['filters', 'columns', 'rows'] as const).map((area) => <PivotFieldArea key={area} area={area} fields={fieldCatalog} fieldIds={definition[area]} disabled={isDisabled} filterSelections={definition.filterSelections} onDrop={handleDrop(area)} onFilter={(fieldId, values) => callbacks?.onFilterChange(fieldId, values)} onGroup={(fieldId, grouped) => callbacks?.onGroupChange(fieldId, grouped)} onRemove={(fieldId, index) => callbacks?.onRemoveField(fieldId, area, index)} onMoveByKeyboard={moveByKeyboard(area)} onSort={(fieldId, direction) => callbacks?.onSortChange(fieldId, direction)} />)}
            <PivotFieldArea area="values" fields={fieldCatalog} fieldIds={definition.values.map((value) => value.fieldId)} disabled={isDisabled} onDrop={handleDrop('values')} onRemove={(fieldId, index) => callbacks?.onRemoveField(fieldId, 'values', index)} onMoveByKeyboard={moveByKeyboard('values')} />
            <Stack gap="xs">{definition.values.map((value) => <PivotValueEditor key={value.id} fields={fieldCatalog} value={value} disabled={isDisabled} onChange={(next) => callbacks?.onValueChange(next)} />)}</Stack>
          </Stack>
          {callbacks ? <PivotActions callbacks={callbacks} definition={definition} fields={fieldCatalog} disabled={isDisabled} /> : null}
          {definition.slicers.map((fieldId) => { const field = fieldCatalog.find((candidate) => candidate.id === fieldId); return field ? <PivotSlicer key={fieldId} field={field} disabled={isDisabled} selectedValues={definition.filterSelections[fieldId] ?? []} onChange={(values) => callbacks?.onFilterChange(fieldId, values)} /> : null; })}
          {definition.timelineFieldId ? <PivotTimeline fieldLabel={fieldCatalog.find((field) => field.id === definition.timelineFieldId)?.label ?? definition.timelineFieldId} start={definition.timelineStart} end={definition.timelineEnd} disabled={isDisabled} onChange={(start, end) => callbacks?.onTimelineRangeChange?.(start, end)} onClear={() => callbacks?.onTimelineChange(undefined)} /> : null}
          {result?.tree ? <PivotResultView tree={result.tree} disabled={isDisabled} expandedFieldIds={definition.expandedFieldIds} onExpandedChange={callbacks?.onExpandedChange} onShowDetails={onShowDetails} /> : null}
          {result || slots?.resultSummary ? <Text size="xs" tone="subtle">{slots?.resultSummary ?? result?.summary ?? `${result?.rowCount ?? 0} rows × ${result?.columnCount ?? 0} columns`}</Text> : null}
        </Stack>
      </PanelBody>
      {onClose ? <PanelFooter className="shrink-0 border-t border-line/80 px-4 py-2"><Button variant="ghost" size="sm" onClick={onClose}>Close panel</Button></PanelFooter> : null}
    </Panel>
  );
}
