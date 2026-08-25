import { useEffect, useMemo, useState, type DragEvent, type PointerEvent as ReactPointerEvent } from 'react';
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
import { PIVOT_FIELD_PANE_LAYOUTS, type PivotFieldArea as Area, type PivotFieldPaneLayout, type PivotManualFilterState, type PivotPanelCallbacks, type PivotPanelSlots, type PivotPanelState, type PivotSlicerControl, type PivotTimelineControl } from '../pivot/pivot-contract';
import type { PivotMessageKey } from '../pivot/pivot-localization';

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
  if (area === 'filters') next.filters.splice(index, 0, { kind: 'manual', fieldId: field.fieldId, scope: 'report', mode: 'all', memberKeys: [] });
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

const fieldPaneLayoutLabels: Record<PivotFieldPaneLayout, PivotMessageKey> = {
  stacked: 'stacked',
  'side-by-side': 'sideBySide',
  'areas-2x2': 'areas2x2',
  'areas-1x4': 'areas1x4',
  'fields-only': 'fieldsOnly',
  'areas-only': 'areasOnly',
};

export function PivotPanel({ activePivotId, callbacks, fieldCatalog: suppliedFields, locale, onClose, pivot, pivotList = [], slicerControls = [], slots, state, timelineControls = [] }: PivotPanelProps) {
  const fields = suppliedFields ?? pivot?.fieldCatalog.fields ?? [];
  const [delayUpdate, setDelayUpdate] = useState(false);
  const [draft, setDraft] = useState<PivotLayout | null>(pivot ? cloneLayout(pivot.layout) : null);
  const [dirty, setDirty] = useState(false);
  const [fieldPaneLayout, setFieldPaneLayout] = useState<PivotFieldPaneLayout>(() => typeof window !== 'undefined' && window.innerWidth >= 1280 ? 'side-by-side' : 'stacked');
  const [fieldPaneSplit, setFieldPaneSplit] = useState(42);
  useEffect(() => { setDraft(pivot ? cloneLayout(pivot.layout) : null); setDirty(false); }, [pivot?.id]);
  const layout = delayUpdate && draft ? draft : pivot?.layout;
  const disabled = Boolean(state?.disabled || state?.loading || state?.error || !pivot || !layout);

  const filters = layout?.filters.filter((field) => field.scope !== 'field').map((field) => field.fieldId) ?? [];
  const columns = layout?.columns.map((field) => field.fieldId) ?? [];
  const rows = layout?.rows.map((field) => field.fieldId) ?? [];
  const values = layout?.values ?? [];
  const selected = useMemo(() => new Set([...filters, ...columns, ...rows, ...values.map((field) => field.fieldId)]), [columns, filters, rows, values]);
  const currentFilterStates = useMemo(() => layout ? filterStates(layout) : {}, [layout]);
  const placements = useMemo(() => layout ? placementMap(layout) : new Map<string, PivotFieldPlacement>(), [layout]);
  const showFields = fieldPaneLayout !== 'areas-only';
  const showAreas = fieldPaneLayout !== 'fields-only';
  const sideBySide = fieldPaneLayout === 'side-by-side';

  const beginSplitDrag = (event: ReactPointerEvent<HTMLElement>) => {
    event.preventDefault();
    const startPosition = sideBySide ? event.clientX : event.clientY;
    const startSplit = fieldPaneSplit;
    const parent = event.currentTarget.parentElement;
    const update = (moveEvent: globalThis.PointerEvent) => {
      const bounds = parent?.getBoundingClientRect();
      if (!bounds) return;
      const available = sideBySide ? bounds.width : bounds.height;
      if (available <= 0) return;
      const position = sideBySide ? moveEvent.clientX : moveEvent.clientY;
      setFieldPaneSplit(Math.max(25, Math.min(70, startSplit + ((position - startPosition) / available) * 100)));
    };
    const stop = () => {
      window.removeEventListener('pointermove', update);
      window.removeEventListener('pointerup', stop);
    };
    window.addEventListener('pointermove', update);
    window.addEventListener('pointerup', stop, { once: true });
  };

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
  const filter = (fieldId: string, nextFilter: PivotManualFilterState) => {
    const next = cloneLayout(layout);
    const index = next.filters.findIndex((entry) => entry.fieldId === fieldId);
    const current = index >= 0 ? next.filters[index] : undefined;
    const scope = rows.includes(fieldId) || columns.includes(fieldId) ? 'field' as const : current?.scope ?? 'report' as const;
    const criterion = { kind: 'manual' as const, fieldId, scope, mode: nextFilter.mode, memberKeys: [...nextFilter.memberKeys] };
    if (index >= 0) next.filters[index] = criterion;
    else next.filters.push(criterion);
    applyLayout(next);
  };
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
        <Inline gap="sm" className="h-9 shrink-0 justify-between">
          <Text size="sm" weight="medium">{pivotText(locale, 'fieldPaneLayout')}</Text>
          <Select aria-label={pivotText(locale, 'fieldPaneLayout')} sizeVariant="sm" value={fieldPaneLayout} onChange={(event) => setFieldPaneLayout(event.target.value as PivotFieldPaneLayout)}>
            {PIVOT_FIELD_PANE_LAYOUTS.map((mode) => <option key={mode} value={mode}>{pivotText(locale, fieldPaneLayoutLabels[mode])}</option>)}
          </Select>
        </Inline>
        <Box className={`min-h-0 min-w-0 flex flex-1 ${sideBySide ? 'flex-row' : 'flex-col'} gap-2`}>
          {showFields ? (
            <Box className="min-h-0 min-w-0 flex flex-col" style={sideBySide ? { width: `${showAreas ? fieldPaneSplit : 100}%` } : { flexBasis: `${showAreas ? fieldPaneSplit : 100}%` }}>
              <Text size="sm" className="mb-1 shrink-0">{pivotText(locale, 'addFields')}</Text>
              <PivotFieldCatalog className="flex-1" locale={locale} fields={fields} selectedFieldIds={selected} disabled={disabled} onToggle={toggle} onToggleVisible={toggleVisible} onDragField={(event, field) => event.dataTransfer.setData('application/x-pivot-field', field.fieldId)} onKeyboardAssign={(fieldId, area) => changeArea(fieldId, area, idsFor(area).length)} />
            </Box>
          ) : null}
          {showFields && showAreas ? (
            <Box
              role="separator"
              aria-label={pivotText(locale, 'paneResize')}
              aria-orientation={sideBySide ? 'vertical' : 'horizontal'}
              aria-valuemin={25}
              aria-valuemax={70}
              aria-valuenow={Math.round(fieldPaneSplit)}
              tabIndex={0}
              className={`${sideBySide ? 'h-full w-2 cursor-col-resize' : 'h-2 w-full cursor-row-resize'} shrink-0 touch-none rounded bg-slate-100 hover:bg-blue-100`}
              onPointerDown={beginSplitDrag}
            />
          ) : null}
          {showAreas ? (
            <Box className="min-h-0 min-w-0 flex flex-1 flex-col">
              <Text size="sm" className="mb-1 shrink-0">{pivotText(locale, 'dragFields')}</Text>
              <Box className={`${fieldPaneLayout === 'areas-1x4' ? 'flex flex-col' : 'grid grid-cols-2 grid-rows-2'} min-h-0 flex-1 gap-1`}>
                {(['filters', 'columns', 'rows', 'values'] as const).map((area) => <Box key={area} className="min-h-0 min-w-0 border border-[#bdbdbd]"><PivotFieldArea className="border-0" locale={locale} area={area} fields={fields} fieldIds={idsFor(area)} placements={placements} filterStates={currentFilterStates} valueFields={values} disabled={disabled} onDrop={drop(area)} onFilter={filter} onGroup={group} onRemove={(fieldId) => removeFromArea(fieldId)} onMoveByKeyboard={(fieldId, itemIndex, direction) => changeArea(fieldId, area, itemIndex + direction)} onSort={sort} onValueChange={valueChange} /></Box>)}
              </Box>
            </Box>
          ) : null}
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
