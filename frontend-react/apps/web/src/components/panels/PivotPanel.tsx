import { Button, Inline, Panel, PanelBody, PanelFooter, PanelHeader, PanelTitle, Select, Stack, StatePanel, Text, TextInput } from '@react-sheets/ui-system';
import { useEffect, useMemo, useState, type DragEvent } from 'react';
import type { PivotFieldDefinition, PivotModel, PivotResultTree, PivotValueField } from '@react-sheets/core-model';
import { cellAddress } from '@react-sheets/spreadsheet-app';
import { PivotActions } from '../pivot/PivotActions';
import { PivotFieldArea } from '../pivot/PivotFieldArea';
import { PivotFieldCatalog } from '../pivot/PivotFieldCatalog';
import { PivotValueEditor } from '../pivot/PivotValueEditor';
import { PivotResultView } from '../pivot/PivotResultView';
import { PivotSlicer } from '../pivot/PivotSlicer';
import { PivotTimeline } from '../pivot/PivotTimeline';
import { PivotCalculatedEditor } from '../pivot/PivotCalculatedEditor';
import type { PivotFieldArea as Area, PivotPanelCallbacks, PivotPanelResult, PivotPanelSlots, PivotPanelState } from '../pivot/pivot-contract';

export interface PivotPanelProps {
  pivotList?: readonly { id: string; label: string }[];
  activePivotId?: string;
  pivot?: PivotModel;
  fieldCatalog?: readonly PivotFieldDefinition[];
  result?: PivotPanelResult;
  onShowDetails?: (paths: import('@react-sheets/core-model').PivotSourceRowPath[]) => void;
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

function valueDefinitions(pivot?: PivotModel): PivotValueField[] {
  return pivot?.layout.values.map((value) => structuredClone(value)) ?? [];
}

function filterSelections(pivot?: PivotModel): Record<string, string[]> {
  const selections: Record<string, string[]> = {};
  if (!pivot) return selections;
  for (const filter of pivot.layout.filters) {
    if (filter.kind === 'manual') selections[filter.field] = filter.selected.map(String);
  }
  return selections;
}

export function PivotPanel({ activePivotId, callbacks, pivot, fieldCatalog = [], onClose, onShowDetails, pivotList = [], result, slots, state }: PivotPanelProps) {
  const isDisabled = state?.disabled || state?.loading || Boolean(state?.error);
  const sourceRangeText = pivot
    ? `${cellAddress(pivot.sourceRange.startRow, pivot.sourceRange.startColumn)}:${cellAddress(pivot.sourceRange.endRow, pivot.sourceRange.endColumn)}`
    : '';
  const [sourceRange, setSourceRange] = useState(sourceRangeText);
  useEffect(() => setSourceRange(sourceRangeText), [sourceRangeText]);

  const rows = pivot?.layout.rows.map((field) => field.field) ?? [];
  const columns = pivot?.layout.columns.map((field) => field.field) ?? [];
  const filters = pivot?.layout.filters.map((filter) => filter.field) ?? [];
  const values = valueDefinitions(pivot);
  const filterSelectionMap = filterSelections(pivot);
  const expandedFieldIds = pivot?.layout.expandedFieldIds ?? rows;
  const slicers = pivot?.slicers?.map((slicer) => slicer.field) ?? [];
  const timelineFieldId = pivot?.timelines?.[0]?.field;
  const timelineStart = pivot?.timelines?.[0]?.start;
  const timelineEnd = pivot?.timelines?.[0]?.end;
  const layoutMode = layoutCompactMode(pivot);
  const calculatedFields = pivot?.layout.calculatedFields ?? [];
  const calculatedItems = pivot?.layout.calculatedItems ?? [];

  const selectedFieldIds = useMemo(
    () => new Set([...filters, ...columns, ...rows, ...values.map((value) => value.field)]),
    [filters, columns, rows, values],
  );

  if (state?.loading) return <Panel className="h-full border-0 bg-transparent shadow-none"><StatePanel kind="loading" description="Preparing pivot field list." /></Panel>;
  if (state?.error) return <Panel className="h-full border-0 bg-transparent shadow-none"><StatePanel kind="error" description={state.error} /></Panel>;
  if (state?.empty || fieldCatalog.length === 0) return <Panel className="h-full border-0 bg-transparent shadow-none"><StatePanel kind={state?.disabled ? 'disabled' : 'empty'} description="No pivot fields are available for this source." /></Panel>;

  const areaLength = (area: Area) => (area === 'values' ? values.length : (area === 'filters' ? filters : area === 'columns' ? columns : rows).length);

  const updateArea = (fieldId: string, area: Area, index: number) => callbacks?.onFieldAreaChange(fieldId, area, index);
  const handleDrop = (area: Area) => (event: DragEvent<HTMLElement>, index?: number) => {
    event.preventDefault();
    const fieldId = event.dataTransfer.getData('application/x-pivot-field');
    if (fieldId) updateArea(fieldId, area, index ?? areaLength(area));
  };
  const handleToggle = (fieldId: string, checked: boolean) => {
    if (checked) updateArea(fieldId, 'rows', rows.length);
    else {
      const currentArea = (['filters', 'columns', 'rows'] as const).find((area) => (area === 'filters' ? filters : area === 'columns' ? columns : rows).includes(fieldId));
      if (currentArea) callbacks?.onRemoveField(fieldId, currentArea, (currentArea === 'filters' ? filters : currentArea === 'columns' ? columns : rows).indexOf(fieldId));
      else {
        const valueIndex = values.findIndex((value) => value.field === fieldId);
        if (valueIndex >= 0) callbacks?.onRemoveField(fieldId, 'values', valueIndex);
      }
    }
  };
  const moveByKeyboard = (area: Area) => (fieldId: string, index: number, direction: -1 | 1) => updateArea(fieldId, area, index + direction);

  return (
    <Panel className="flex h-full min-h-0 flex-col border-0 bg-transparent shadow-none">
      <PanelHeader className="h-12 shrink-0 border-b border-line/80 px-4"><Inline gap="sm" className="min-w-0 justify-between"><Stack gap="none" className="min-w-0"><PanelTitle size="sm">Pivot field list</PanelTitle><Text size="xs" tone="subtle">Build, filter and configure the pivot view</Text></Stack><Inline gap="xs">{pivotList.length > 1 ? <Select aria-label="Active pivot table" disabled={isDisabled} sizeVariant="sm" value={activePivotId ?? pivotList[0]?.id} onChange={(event) => callbacks?.onPivotSelect?.(event.target.value)}>{pivotList.map((entry) => <option key={entry.id} value={entry.id}>{entry.label}</option>)}</Select> : null}{callbacks?.onCreate ? <Button disabled={isDisabled} size="xs" variant="primary" onClick={callbacks.onCreate}>New pivot</Button> : null}{slots?.headerActions}</Inline></Inline></PanelHeader>
      <PanelBody className="min-h-0 flex-1 overflow-auto p-4">
        <Stack gap="md">
          <Stack gap="xs">
            <Text size="xs" weight="semibold" tone="muted">Data source range</Text>
            <TextInput
              aria-label="Pivot data source range"
              value={sourceRange}
              placeholder="A1:F100"
              disabled={isDisabled}
              onChange={(event) => setSourceRange(event.target.value)}
              onBlur={() => callbacks?.onSourceRangeChange?.(sourceRange.trim())}
              onKeyDown={(event) => {
                if (event.key === 'Enter') callbacks?.onSourceRangeChange?.(sourceRange.trim());
              }}
            />
          </Stack>
          <PivotFieldCatalog fields={fieldCatalog} selectedFieldIds={selectedFieldIds} disabled={isDisabled} onToggle={handleToggle} onDragField={(event, field) => event.dataTransfer.setData('application/x-pivot-field', field.id)} onKeyboardAssign={(fieldId, area) => updateArea(fieldId, area, areaLength(area))} />
          <Stack gap="sm">
            <PivotFieldArea area="filters" fields={fieldCatalog} fieldIds={filters} disabled={isDisabled} filterSelections={filterSelectionMap} onDrop={handleDrop('filters')} onFilter={(fieldId, selected) => callbacks?.onFilterChange(fieldId, selected)} onGroup={(fieldId, grouped) => callbacks?.onGroupChange(fieldId, grouped)} onRemove={(fieldId, index) => callbacks?.onRemoveField(fieldId, 'filters', index)} onMoveByKeyboard={moveByKeyboard('filters')} onSort={(fieldId, direction) => callbacks?.onSortChange(fieldId, direction)} />
            <PivotFieldArea area="columns" fields={fieldCatalog} fieldIds={columns} disabled={isDisabled} filterSelections={filterSelectionMap} onDrop={handleDrop('columns')} onFilter={(fieldId, selected) => callbacks?.onFilterChange(fieldId, selected)} onGroup={(fieldId, grouped) => callbacks?.onGroupChange(fieldId, grouped)} onRemove={(fieldId, index) => callbacks?.onRemoveField(fieldId, 'columns', index)} onMoveByKeyboard={moveByKeyboard('columns')} onSort={(fieldId, direction) => callbacks?.onSortChange(fieldId, direction)} />
            <PivotFieldArea area="rows" fields={fieldCatalog} fieldIds={rows} disabled={isDisabled} filterSelections={filterSelectionMap} onDrop={handleDrop('rows')} onFilter={(fieldId, selected) => callbacks?.onFilterChange(fieldId, selected)} onGroup={(fieldId, grouped) => callbacks?.onGroupChange(fieldId, grouped)} onRemove={(fieldId, index) => callbacks?.onRemoveField(fieldId, 'rows', index)} onMoveByKeyboard={moveByKeyboard('rows')} onSort={(fieldId, direction) => callbacks?.onSortChange(fieldId, direction)} />
            <PivotFieldArea area="values" fields={fieldCatalog} fieldIds={values.map((value) => value.field)} disabled={isDisabled} onDrop={handleDrop('values')} onRemove={(fieldId, index) => callbacks?.onRemoveField(fieldId, 'values', index)} onMoveByKeyboard={moveByKeyboard('values')} />
            <Stack gap="xs">{values.map((value, index) => <PivotValueEditor key={`${value.field}-${index}`} fields={fieldCatalog} value={value} disabled={isDisabled} onChange={(next) => callbacks?.onValueChange(next)} />)}</Stack>
            <PivotCalculatedEditor fields={fieldCatalog} calculatedFields={calculatedFields} calculatedItems={calculatedItems} disabled={isDisabled} onFieldsChange={(fields) => callbacks?.onCalculatedFieldsChange?.(fields)} onItemsChange={(items) => callbacks?.onCalculatedItemsChange?.(items)} />
          </Stack>
          {callbacks && pivot ? (
            <PivotActions
              callbacks={callbacks}
              layout={layoutMode}
              expandedFieldIds={expandedFieldIds}
              slicers={slicers}
              fields={fieldCatalog}
              disabled={isDisabled}
            />
          ) : null}
          {slicers.map((fieldId) => {
            const field = fieldCatalog.find((candidate) => candidate.id === fieldId);
            return field ? <PivotSlicer key={fieldId} field={field} disabled={isDisabled} selectedValues={filterSelectionMap[fieldId] ?? []} onChange={(selected) => callbacks?.onFilterChange(fieldId, selected)} /> : null;
          })}
          {timelineFieldId ? <PivotTimeline fieldLabel={fieldCatalog.find((field) => field.id === timelineFieldId)?.name ?? timelineFieldId} start={timelineStart} end={timelineEnd} disabled={isDisabled} onChange={(start, end) => callbacks?.onTimelineRangeChange?.(start, end)} onClear={() => callbacks?.onTimelineChange(undefined)} /> : null}
          {result?.tree ? <PivotResultView tree={result.tree as PivotResultTree} disabled={isDisabled} expandedFieldIds={expandedFieldIds} onExpandedChange={callbacks?.onExpandedChange} onShowDetails={onShowDetails} /> : null}
          {result || slots?.resultSummary ? <Text size="xs" tone="subtle">{slots?.resultSummary ?? result?.summary ?? `${result?.rowCount ?? 0} rows × ${result?.columnCount ?? 0} columns`}</Text> : null}
        </Stack>
      </PanelBody>
      {onClose ? <PanelFooter className="shrink-0 border-t border-line/80 px-4 py-2"><Button variant="ghost" size="sm" onClick={onClose}>Close panel</Button></PanelFooter> : null}
    </Panel>
  );
}
