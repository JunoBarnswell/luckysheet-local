import { Button, Inline, Panel, PanelBody, PanelFooter, PanelHeader, PanelTitle, Select, Stack, StatePanel, Text } from '@react-sheets/ui-system';
import { useMemo, type DragEvent } from 'react';
import type {
  PivotFieldDefinition,
  PivotFieldPlacement,
  PivotModel,
  PivotValueField,
} from '@react-sheets/core-model';
import { cellAddress } from '@react-sheets/spreadsheet-app';
import { PivotActions } from '../pivot/PivotActions';
import { PivotFieldArea } from '../pivot/PivotFieldArea';
import { PivotFieldCatalog } from '../pivot/PivotFieldCatalog';
import { PivotValueEditor } from '../pivot/PivotValueEditor';
import { PivotSlicer } from '../pivot/PivotSlicer';
import { PivotTimeline } from '../pivot/PivotTimeline';
import { PivotCalculatedEditor } from '../pivot/PivotCalculatedEditor';
import type {
  PivotFieldArea as Area,
  PivotManualFilterState,
  PivotPanelCallbacks,
  PivotPanelSlots,
  PivotPanelState,
  PivotSlicerControl,
  PivotTimelineControl,
} from '../pivot/pivot-contract';

export interface PivotPanelProps {
  pivotList?: readonly { id: string; label: string }[];
  activePivotId?: string;
  pivot?: PivotModel;
  fieldCatalog?: readonly PivotFieldDefinition[];
  slicerControls?: readonly PivotSlicerControl[];
  timelineControls?: readonly PivotTimelineControl[];
  state?: PivotPanelState;
  callbacks?: PivotPanelCallbacks;
  slots?: PivotPanelSlots;
  onClose?: () => void;
}

function layoutCompactMode(pivot?: PivotModel): 'compact' | 'outline' | 'tabular' {
  if (!pivot) return 'compact';
  if (pivot.layout.compact) return 'compact';
  if (pivot.layout.repeatLabels) return 'tabular';
  return 'outline';
}

function sourceDescription(pivot?: PivotModel): string {
  const source = pivot?.source;
  if (!source) return 'No data source selected';
  if (source.kind === 'worksheet-range') {
    const range = source.range;
    return `Worksheet range · ${cellAddress(range.startRow, range.startColumn)}:${cellAddress(range.endRow, range.endColumn)}`;
  }
  if (source.kind === 'worksheet-ranges') return `Worksheet ranges · ${source.ranges.length} regions`;
  if (source.kind === 'table') return `Table · ${source.tableId}`;
  if (source.kind === 'named-range') return `Named range · ${source.name}`;
  return `Data source · ${source.dataSourceId}`;
}

function filterStates(pivot?: PivotModel): Record<string, PivotManualFilterState> {
  const states: Record<string, PivotManualFilterState> = {};
  for (const filter of pivot?.layout.filters ?? []) {
    if (filter.kind !== 'manual') continue;
    states[filter.fieldId] = {
      mode: filter.mode,
      memberKeys: filter.memberKeys,
    };
  }
  return states;
}

function placementMap(pivot?: PivotModel): ReadonlyMap<string, PivotFieldPlacement> {
  const placements = new Map<string, PivotFieldPlacement>();
  for (const placement of [...(pivot?.layout.rows ?? []), ...(pivot?.layout.columns ?? [])]) {
    placements.set(placement.fieldId, placement);
  }
  return placements;
}

export function PivotPanel({
  activePivotId,
  callbacks,
  fieldCatalog: providedFieldCatalog,
  onClose,
  pivot,
  pivotList = [],
  slicerControls = [],
  slots,
  state,
  timelineControls = [],
}: PivotPanelProps) {
  const fieldCatalog = providedFieldCatalog ?? pivot?.fieldCatalog.fields ?? [];
  const isDisabled = Boolean(state?.disabled || state?.loading || state?.error || !pivot);
  const rows = (pivot?.layout.rows ?? []).map((field) => field.fieldId);
  const columns = (pivot?.layout.columns ?? []).map((field) => field.fieldId);
  const filters = (pivot?.layout.filters ?? []).map((filter) => filter.fieldId);
  const values = pivot?.layout.values ?? [];
  const filterStateMap = useMemo(() => filterStates(pivot), [pivot]);
  const placements = useMemo(() => placementMap(pivot), [pivot]);
  const layoutMode = layoutCompactMode(pivot);
  const calculatedFields = pivot?.layout.calculatedFields ?? [];
  const calculatedItems = pivot?.layout.calculatedItems ?? [];
  const selectedFieldIds = useMemo(
    () => new Set([...filters, ...columns, ...rows, ...values.map((value) => value.fieldId)]),
    [columns, filters, rows, values],
  );

  if (state?.loading) return <Panel className="h-full border-0 bg-transparent shadow-none"><StatePanel kind="loading" description="Preparing the PivotTable field list." /></Panel>;
  if (state?.error) return <Panel className="h-full border-0 bg-transparent shadow-none"><StatePanel kind="error" description={state.error} /></Panel>;
  if (!pivot || state?.empty || fieldCatalog.length === 0) {
    const description = state?.emptyMessage ?? (!pivot ? 'Select a PivotTable to configure its fields.' : 'No fields are available for this PivotTable source.');
    return <Panel className="h-full border-0 bg-transparent shadow-none"><StatePanel kind={state?.disabled ? 'disabled' : 'empty'} description={description} /></Panel>;
  }

  const areaLength = (area: Area) => (area === 'values' ? values.length : area === 'filters' ? filters.length : area === 'columns' ? columns.length : rows.length);
  const areaFieldIds = (area: Area): readonly string[] => area === 'values' ? values.map((value) => value.fieldId) : area === 'filters' ? filters : area === 'columns' ? columns : rows;
  const updateArea = (fieldId: string, area: Area, index: number) => callbacks?.onFieldAreaChange(fieldId, area, index);
  const handleDrop = (area: Area) => (event: DragEvent<HTMLElement>, index?: number) => {
    event.preventDefault();
    const fieldId = event.dataTransfer.getData('application/x-pivot-field');
    if (fieldId) updateArea(fieldId, area, index ?? areaLength(area));
  };
  const handleToggle = (fieldId: string, checked: boolean) => {
    if (checked) {
      updateArea(fieldId, 'rows', rows.length);
      return;
    }
    const currentArea = (['filters', 'columns', 'rows'] as const).find((area) => areaFieldIds(area).includes(fieldId));
    if (currentArea) {
      callbacks?.onRemoveField(fieldId, currentArea, areaFieldIds(currentArea).indexOf(fieldId));
      return;
    }
    const valueIndex = values.findIndex((value) => value.fieldId === fieldId);
    if (valueIndex >= 0) callbacks?.onRemoveField(fieldId, 'values', valueIndex);
  };
  const moveByKeyboard = (area: Area) => (fieldId: string, index: number, direction: -1 | 1) => updateArea(fieldId, area, index + direction);

  return (
    <Panel className="flex h-full min-h-0 flex-col border-0 bg-transparent shadow-none" data-testid="pivot-field-list">
      <PanelHeader className="h-12 shrink-0 border-b border-line/80 px-4">
        <Inline gap="sm" className="min-w-0 justify-between">
          <Stack gap="none" className="min-w-0">
            <PanelTitle size="sm">PivotTable field list</PanelTitle>
            <Text size="xs" tone="subtle">Configure fields, filters, grouping and values. The report is rendered on the worksheet.</Text>
          </Stack>
          <Inline gap="xs">
            {pivotList.length > 1 ? <Select aria-label="Active PivotTable" disabled={isDisabled} sizeVariant="sm" value={activePivotId ?? pivotList[0]?.id} onChange={(event) => callbacks?.onPivotSelect?.(event.target.value)}>{pivotList.map((entry) => <option key={entry.id} value={entry.id}>{entry.label}</option>)}</Select> : null}
            {callbacks?.onCreate ? <Button disabled={isDisabled} size="xs" variant="primary" onClick={callbacks.onCreate}>New PivotTable</Button> : null}
            {slots?.headerActions}
          </Inline>
        </Inline>
      </PanelHeader>
      <PanelBody className="min-h-0 flex-1 overflow-auto p-4">
        <Stack gap="md">
          <Stack gap="xs" className="rounded-lg border border-line/80 bg-white/70 p-2">
            <Text size="xs" weight="semibold" tone="muted">SOURCE</Text>
            <Text size="xs" tone="default">{sourceDescription(pivot)}</Text>
            <Text size="xs" tone="subtle">Target · {pivot.target.sheetId} · row {String(pivot.target.anchor.row + 1)}, column {String(pivot.target.anchor.column + 1)}</Text>
          </Stack>
          <PivotFieldCatalog fields={fieldCatalog} selectedFieldIds={selectedFieldIds} disabled={isDisabled} onToggle={handleToggle} onDragField={(event, field) => { event.dataTransfer.setData('application/x-pivot-field', field.fieldId); }} onKeyboardAssign={(fieldId, area) => updateArea(fieldId, area, areaLength(area))} />
          <Stack gap="sm">
            {(['filters', 'columns', 'rows', 'values'] as const).map((area) => (
              <PivotFieldArea
                key={area}
                area={area}
                fields={fieldCatalog}
                fieldIds={areaFieldIds(area)}
                placements={placements}
                filterStates={filterStateMap}
                disabled={isDisabled}
                onDrop={handleDrop(area)}
                onFilter={(fieldId, filter) => callbacks?.onFilterChange(fieldId, filter)}
                onGroup={(fieldId, group) => callbacks?.onGroupChange(fieldId, group)}
                onRemove={(fieldId, index) => callbacks?.onRemoveField(fieldId, area, index)}
                onMoveByKeyboard={moveByKeyboard(area)}
                onSort={(fieldId, sort) => callbacks?.onSortChange(fieldId, sort)}
              />
            ))}
            <Stack gap="xs">
              {values.map((value, index) => <PivotValueEditor key={`${value.fieldId}-${index}`} fields={fieldCatalog} value={value} disabled={isDisabled} onChange={(next) => callbacks?.onValueChange({ ...next, fieldId: next.fieldId })} />)}
            </Stack>
            <PivotCalculatedEditor fields={fieldCatalog} calculatedFields={calculatedFields} calculatedItems={calculatedItems} disabled={isDisabled} onFieldsChange={(next) => callbacks?.onCalculatedFieldsChange?.(next)} onItemsChange={(next) => callbacks?.onCalculatedItemsChange?.(next)} />
          </Stack>
          {callbacks && pivot ? <PivotActions callbacks={callbacks} layout={layoutMode} slicerFieldIds={slicerControls.map((control) => control.fieldId)} fields={fieldCatalog} disabled={isDisabled} /> : null}
          {slicerControls.map((control) => {
            const field = fieldCatalog.find((candidate) => candidate.fieldId === control.fieldId);
            if (!field) return null;
            return <PivotSlicer key={control.id} field={field} mode={control.mode} memberKeys={control.memberKeys} disabled={isDisabled || !callbacks?.onSlicerFilterChange} onChange={(next) => callbacks?.onSlicerFilterChange?.(control.id, next)} />;
          })}
          {timelineControls.map((control) => {
            const fieldLabel = fieldCatalog.find((field) => field.fieldId === control.fieldId)?.name ?? control.fieldId;
            return <PivotTimeline key={control.id} fieldLabel={fieldLabel} start={control.start} end={control.end} disabled={isDisabled || !callbacks?.onTimelineRangeChange} onChange={(start, end) => callbacks?.onTimelineRangeChange?.(control.id, start, end)} onClear={() => callbacks?.onTimelineChange(undefined)} />;
          })}
          {slots?.statusSummary ? <Text size="xs" tone="subtle">{slots.statusSummary}</Text> : null}
        </Stack>
      </PanelBody>
      {onClose ? <PanelFooter className="shrink-0 border-t border-line/80 px-4 py-2"><Button variant="ghost" size="sm" onClick={onClose}>Close panel</Button></PanelFooter> : null}
    </Panel>
  );
}
