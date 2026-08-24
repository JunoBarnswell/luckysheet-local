import { useEffect, useMemo, useState, type DragEvent } from 'react';
import { Box, Button, CheckToggle, DropdownMenu, Inline, Panel, Select, Stack, StatePanel, Text } from '@react-sheets/ui-system';
import type { PivotFieldDefinition, PivotFieldPlacement, PivotLayout, PivotModel, PivotValueField } from '@react-sheets/core-model';
import type { Locale } from '../../i18n';
import { PivotActions } from '../pivot/PivotActions';
import { PivotCalculatedEditor } from '../pivot/PivotCalculatedEditor';
import { PivotFieldArea } from '../pivot/PivotFieldArea';
import { PivotFieldCatalog } from '../pivot/PivotFieldCatalog';
import { PivotSlicer } from '../pivot/PivotSlicer';
import { PivotTimeline } from '../pivot/PivotTimeline';
import { pivotText } from '../pivot/pivot-localization';
import type { PivotFieldArea as Area, PivotManualFilterState, PivotPanelCallbacks, PivotPanelSlots, PivotPanelState, PivotSlicerControl, PivotTimelineControl } from '../pivot/pivot-contract';

export interface PivotPanelProps {
  locale: Locale;
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

function cloneLayout(layout: PivotLayout): PivotLayout { return structuredClone(layout); }
function layoutMode(layout: PivotLayout): 'compact' | 'outline' | 'tabular' { return layout.compact ? 'compact' : layout.repeatLabels ? 'tabular' : 'outline'; }
function removeField(layout: PivotLayout, fieldId: string): PivotLayout {
  const next = cloneLayout(layout);
  next.filters = next.filters.filter((field) => field.fieldId !== fieldId);
  next.columns = next.columns.filter((field) => field.fieldId !== fieldId);
  next.rows = next.rows.filter((field) => field.fieldId !== fieldId);
  next.values = next.values.filter((field) => field.fieldId !== fieldId);
  return next;
}
function moveField(layout: PivotLayout, field: PivotFieldDefinition, area: Area, index: number): PivotLayout {
  const next = removeField(layout, field.fieldId);
  if (area === 'filters') next.filters.splice(index, 0, { kind: 'manual', fieldId: field.fieldId, mode: 'all', memberKeys: [] });
  else if (area === 'columns') next.columns.splice(index, 0, { fieldId: field.fieldId });
  else if (area === 'rows') next.rows.splice(index, 0, { fieldId: field.fieldId });
  else next.values.splice(index, 0, { fieldId: field.fieldId, summarizeBy: field.dataType === 'number' ? 'sum' : 'count' });
  return next;
}
function defaultArea(field: PivotFieldDefinition): Area {
  if (field.dataType === 'number') return 'values';
  if (field.dataType === 'date') return 'columns';
  return 'rows';
}
function filterStates(layout: PivotLayout): Record<string, PivotManualFilterState> {
  const result: Record<string, PivotManualFilterState> = {};
  for (const filter of layout.filters) if (filter.kind === 'manual') result[filter.fieldId] = { mode: filter.mode, memberKeys: filter.memberKeys };
  return result;
}
function placementMap(layout: PivotLayout): ReadonlyMap<string, PivotFieldPlacement> {
  return new Map([...layout.rows, ...layout.columns].map((placement) => [placement.fieldId, placement]));
}

export function PivotPanel({ activePivotId, callbacks, fieldCatalog: suppliedFields, locale, onClose, pivot, pivotList = [], slicerControls = [], slots, state, timelineControls = [] }: PivotPanelProps) {
  const fields = suppliedFields ?? pivot?.fieldCatalog.fields ?? [];
  const [delayUpdate, setDelayUpdate] = useState(false);
  const [draft, setDraft] = useState<PivotLayout | null>(pivot ? cloneLayout(pivot.layout) : null);
  const [dirty, setDirty] = useState(false);
  useEffect(() => { setDraft(pivot ? cloneLayout(pivot.layout) : null); setDirty(false); }, [pivot?.id]);
  const layout = delayUpdate && draft ? draft : pivot?.layout;
  const disabled = Boolean(state?.disabled || state?.loading || state?.error || !pivot || !layout);

  const filters = layout?.filters.map((field) => field.fieldId) ?? [];
  const columns = layout?.columns.map((field) => field.fieldId) ?? [];
  const rows = layout?.rows.map((field) => field.fieldId) ?? [];
  const values = layout?.values ?? [];
  const selected = useMemo(() => new Set([...filters, ...columns, ...rows, ...values.map((field) => field.fieldId)]), [columns, filters, rows, values]);
  const currentFilterStates = useMemo(() => layout ? filterStates(layout) : {}, [layout]);
  const placements = useMemo(() => layout ? placementMap(layout) : new Map<string, PivotFieldPlacement>(), [layout]);

  if (state?.loading) return <Panel className="h-full border-0 shadow-none"><StatePanel kind="loading" description={pivotText(locale, 'loading')} /></Panel>;
  if (state?.error) return <Panel className="h-full border-0 shadow-none"><StatePanel kind="error" description={state.error || pivotText(locale, 'error')} /></Panel>;
  if (!pivot || !layout || state?.empty || fields.length === 0) return <Panel className="h-full border-0 shadow-none"><StatePanel kind={state?.disabled ? 'disabled' : 'empty'} description={state?.emptyMessage ?? pivotText(locale, 'empty')} /></Panel>;

  const applyLayout = (next: PivotLayout) => {
    if (delayUpdate) { setDraft(cloneLayout(next)); setDirty(true); }
    else callbacks?.onLayoutReplace(next);
  };
  const idsFor = (area: Area): readonly string[] => area === 'filters' ? filters : area === 'columns' ? columns : area === 'rows' ? rows : values.map((field) => field.fieldId);
  const changeArea = (fieldId: string, area: Area, index: number) => {
    const field = fields.find((candidate) => candidate.fieldId === fieldId);
    if (field) applyLayout(moveField(layout, field, area, index));
  };
  const removeFromArea = (fieldId: string) => applyLayout(removeField(layout, fieldId));
  const drop = (area: Area) => (event: DragEvent<HTMLElement>, index?: number) => { event.preventDefault(); const fieldId = event.dataTransfer.getData('application/x-pivot-field'); if (fieldId) changeArea(fieldId, area, index ?? idsFor(area).length); };
  const toggle = (fieldId: string, checked: boolean) => {
    if (!checked) { removeFromArea(fieldId); return; }
    const field = fields.find((candidate) => candidate.fieldId === fieldId);
    if (!field) return;
    const area = defaultArea(field);
    changeArea(fieldId, area, idsFor(area).length);
  };
  const toggleVisible = (fieldIds: readonly string[], checked: boolean) => {
    let next = cloneLayout(layout);
    for (const fieldId of fieldIds) {
      next = removeField(next, fieldId);
      if (!checked) continue;
      const field = fields.find((candidate) => candidate.fieldId === fieldId);
      if (!field) continue;
      const area = defaultArea(field);
      const index = area === 'filters' ? next.filters.length : area === 'columns' ? next.columns.length : area === 'rows' ? next.rows.length : next.values.length;
      next = moveField(next, field, area, index);
    }
    applyLayout(next);
  };
  const filter = (fieldId: string, nextFilter: PivotManualFilterState) => applyLayout({ ...cloneLayout(layout), filters: layout.filters.map((entry) => entry.fieldId === fieldId ? { kind: 'manual' as const, fieldId, mode: nextFilter.mode, memberKeys: [...nextFilter.memberKeys] } : entry) });
  const sort = (fieldId: string, nextSort: Parameters<NonNullable<PivotPanelCallbacks['onSortChange']>>[1]) => applyLayout({ ...cloneLayout(layout), rows: layout.rows.map((entry) => entry.fieldId === fieldId ? { ...entry, sort: nextSort } : entry), columns: layout.columns.map((entry) => entry.fieldId === fieldId ? { ...entry, sort: nextSort } : entry) });
  const group = (fieldId: string, nextGroup: Parameters<NonNullable<PivotPanelCallbacks['onGroupChange']>>[1]) => applyLayout({ ...cloneLayout(layout), rows: layout.rows.map((entry) => entry.fieldId === fieldId ? { ...entry, group: nextGroup } : entry), columns: layout.columns.map((entry) => entry.fieldId === fieldId ? { ...entry, group: nextGroup } : entry) });
  const valueChange = (value: PivotValueField) => applyLayout({ ...cloneLayout(layout), values: layout.values.map((entry) => entry.fieldId === value.fieldId ? value : entry) });

  return (
    <Panel className="flex h-full min-h-0 flex-col rounded-none border-0 bg-white shadow-none" data-testid="pivot-field-list">
      <Inline gap="sm" className="h-14 w-full shrink-0 justify-between px-4">
        <Text size="lg" weight="bold">{pivotText(locale, 'fieldsTitle')}</Text>
        <Inline gap="sm"><Box className="rounded-full border-2 border-[#a529ff] px-3 py-0.5 text-[#8b20e8]"><Text size="sm" weight="bold">AI</Text></Box>{onClose ? <Button aria-label={pivotText(locale, 'close')} icon="x" iconOnly size="sm" variant="ghost" onClick={onClose} /> : null}</Inline>
      </Inline>
      <Stack gap="sm" className="min-h-0 flex-1 px-4 pb-2">
        <Text size="sm">{pivotText(locale, 'addFields')}</Text>
        <PivotFieldCatalog locale={locale} fields={fields} selectedFieldIds={selected} disabled={disabled} onToggle={toggle} onToggleVisible={toggleVisible} onDragField={(event, field) => event.dataTransfer.setData('application/x-pivot-field', field.fieldId)} onKeyboardAssign={(fieldId, area) => changeArea(fieldId, area, idsFor(area).length)} />
        <Text size="sm">{pivotText(locale, 'dragFields')}</Text>
        <Box className="grid min-h-[260px] flex-1 grid-cols-2 grid-rows-2 border border-[#bdbdbd]">
          {(['filters', 'columns', 'rows', 'values'] as const).map((area, index) => <Box key={area} className={`${index % 2 === 0 ? 'border-r border-[#bdbdbd]' : ''}${index < 2 ? ' border-b border-[#bdbdbd]' : ''}`}><PivotFieldArea locale={locale} area={area} fields={fields} fieldIds={idsFor(area)} placements={placements} filterStates={currentFilterStates} valueFields={values} disabled={disabled} onDrop={drop(area)} onFilter={filter} onGroup={group} onRemove={(fieldId) => removeFromArea(fieldId)} onMoveByKeyboard={(fieldId, itemIndex, direction) => changeArea(fieldId, area, itemIndex + direction)} onSort={sort} onValueChange={valueChange} /></Box>)}
        </Box>
        <Inline gap="sm" className="h-9 justify-between">
          <CheckToggle label={pivotText(locale, 'delayUpdate')} checked={delayUpdate} onChange={(event) => { setDelayUpdate(event.target.checked); setDraft(cloneLayout(pivot.layout)); setDirty(false); }} />
          <Button size="sm" variant="outline" disabled={!delayUpdate || !dirty} onClick={() => { if (draft) callbacks?.onLayoutReplace(draft); setDirty(false); }}>{pivotText(locale, 'update')}</Button>
        </Inline>
      </Stack>
      <Inline gap="sm" className="h-12 shrink-0 border-t border-[#d0d0d0] px-4">
        <DropdownMenu align="left" trigger={<Button size="sm" variant="ghost">{pivotText(locale, 'view')}</Button>}><Stack gap="sm" className="w-[19rem] p-2">{callbacks ? <PivotActions locale={locale} callbacks={callbacks} layout={layoutMode(layout)} slicerFieldIds={slicerControls.map((control) => control.fieldId)} fields={fields} disabled={disabled} /> : null}{slicerControls.map((control) => { const field = fields.find((candidate) => candidate.fieldId === control.fieldId); return field ? <PivotSlicer key={control.id} locale={locale} field={field} mode={control.mode} memberKeys={control.memberKeys} disabled={disabled} onChange={(next) => callbacks?.onSlicerFilterChange?.(control.id, next)} /> : null; })}{timelineControls.map((control) => <PivotTimeline key={control.id} locale={locale} fieldLabel={fields.find((field) => field.fieldId === control.fieldId)?.name ?? control.fieldId} start={control.start} end={control.end} disabled={disabled} onChange={(start, end) => callbacks?.onTimelineRangeChange?.(control.id, start, end)} onClear={() => callbacks?.onTimelineChange(undefined)} />)}</Stack></DropdownMenu>
        <Select aria-label={pivotText(locale, 'view')} sizeVariant="sm" value={layoutMode(layout)} onChange={(event) => callbacks?.onLayoutChange(event.target.value as 'compact' | 'outline' | 'tabular')}><option value="compact">{pivotText(locale, 'compact')}</option><option value="outline">{pivotText(locale, 'outline')}</option><option value="tabular">{pivotText(locale, 'tabular')}</option></Select>
        <DropdownMenu align="right" trigger={<Button aria-label={pivotText(locale, 'advancedFields')} icon="plus" iconOnly size="sm" variant="ghost" className="ml-auto" />}><Box className="w-[22rem] p-3"><PivotCalculatedEditor locale={locale} fields={fields} calculatedFields={layout.calculatedFields ?? []} calculatedItems={layout.calculatedItems ?? []} disabled={disabled} onFieldsChange={(next) => applyLayout({ ...cloneLayout(layout), calculatedFields: next })} onItemsChange={(next) => applyLayout({ ...cloneLayout(layout), calculatedItems: next })} /></Box></DropdownMenu>
        {slots?.statusSummary ? <Text size="xs" tone="subtle">{slots.statusSummary}</Text> : null}
      </Inline>
    </Panel>
  );
}
