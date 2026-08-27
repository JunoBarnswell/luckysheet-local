import { useEffect, useMemo, useState, type DragEvent, type PointerEvent as ReactPointerEvent } from 'react';
import { Box, Button, CheckToggle, DropdownMenu, Inline, Panel, Select, Stack, StatePanel, Text } from '@react-sheets/ui-system';
import { type PivotFieldDefinition, type PivotFieldPlacement, type PivotLayout, type PivotModel, type PivotValueField } from '@react-sheets/core-model';
import type { Locale } from '../../i18n';
import { buildPivotGroupedFilterMembers } from '@react-sheets/spreadsheet-app';
import { PivotCalculatedEditor } from '../pivot/PivotCalculatedEditor';
import { PivotFieldArea } from '../pivot/PivotFieldArea';
import { PivotFieldCatalog } from '../pivot/PivotFieldCatalog';
import { PivotFormatOptions } from '../pivot/PivotFormatOptions';
import { PivotSlicer } from '../pivot/PivotSlicer';
import { PivotTimeline } from '../pivot/PivotTimeline';
import { pivotTemplate, pivotText } from '../pivot/pivot-localization';
import { DEFAULT_PIVOT_FIELD_PANE_LAYOUT, defaultPivotFieldArea, PIVOT_FIELD_AREAS, PIVOT_FIELD_PANE_LAYOUTS, shouldDeferPivotLayoutUpdates, type PivotFieldArea as Area, type PivotFieldPaneLayout, type PivotManualFilterState, type PivotPanelCallbacks, type PivotPanelSlots, type PivotPanelState, type PivotSlicerControl, type PivotTimelineControl } from '../pivot/pivot-contract';
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
function removeField(layout: PivotLayout, fieldId: string): PivotLayout {
  const next = cloneLayout(layout);
  next.filters = next.filters.filter((field) => field.fieldId !== fieldId);
  next.columns = next.columns.filter((field) => field.fieldId !== fieldId);
  next.rows = next.rows.filter((field) => field.fieldId !== fieldId);
  next.values = next.values.filter((field) => field.fieldId !== fieldId);
  return next;
}
function removeValueField(layout: PivotLayout, valueId: string): PivotLayout {
  const next = cloneLayout(layout);
  next.values = next.values.filter((field) => field.valueId !== valueId);
  return next;
}
function nextValueId(layout: PivotLayout, fieldId: string): string {
  const base = `value:${fieldId}`;
  if (!layout.values.some((field) => field.valueId === base)) return base;
  let index = 2;
  while (layout.values.some((field) => field.valueId === `${base}:${index}`)) index += 1;
  return `${base}:${index}`;
}
function moveField(layout: PivotLayout, field: PivotFieldDefinition, area: Area, index: number): PivotLayout {
  const next = cloneLayout(layout);
  if (area !== 'values') {
    next.filters = next.filters.filter((entry) => entry.fieldId !== field.fieldId);
    next.columns = next.columns.filter((entry) => entry.fieldId !== field.fieldId);
    next.rows = next.rows.filter((entry) => entry.fieldId !== field.fieldId);
  }
  if (area === 'filters') next.filters.splice(index, 0, { kind: 'manual', family: 'manual', fieldId: field.fieldId, scope: 'report', mode: 'all', memberKeys: [] });
  else if (area === 'columns') next.columns.splice(index, 0, { fieldId: field.fieldId });
  else if (area === 'rows') next.rows.splice(index, 0, { fieldId: field.fieldId });
  else next.values.splice(index, 0, { valueId: nextValueId(next, field.fieldId), fieldId: field.fieldId, summarizeBy: field.dataType === 'number' ? 'sum' : 'count' });
  return next;
}
function moveValueField(layout: PivotLayout, valueId: string, index: number): PivotLayout {
  const next = cloneLayout(layout);
  const currentIndex = next.values.findIndex((field) => field.valueId === valueId);
  if (currentIndex < 0) throw new Error(`Unknown Pivot Values placement: ${valueId}`);
  const [value] = next.values.splice(currentIndex, 1);
  next.values.splice(Math.max(0, Math.min(index, next.values.length)), 0, value!);
  return next;
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
  const [delayUpdate, setDelayUpdate] = useState(() => pivot ? shouldDeferPivotLayoutUpdates(pivot) : false);
  const [draft, setDraft] = useState<PivotLayout | null>(pivot ? cloneLayout(pivot.layout) : null);
  const [dirty, setDirty] = useState(false);
  const [fieldPaneLayout, setFieldPaneLayout] = useState<PivotFieldPaneLayout>(DEFAULT_PIVOT_FIELD_PANE_LAYOUT);
  const [fieldPaneSplit, setFieldPaneSplit] = useState(45);
  useEffect(() => { setDraft(pivot ? cloneLayout(pivot.layout) : null); setDelayUpdate(pivot ? shouldDeferPivotLayoutUpdates(pivot) : false); setDirty(false); }, [pivot?.id]);
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
  const workspaceClass = sideBySide
    ? 'flex flex-row'
    : showFields && showAreas ? 'grid grid-rows-[minmax(120px,45%)_8px_minmax(220px,1fr)]' : 'flex flex-col';

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
  const applyDraft = () => {
    if (draft) callbacks?.onLayoutReplace(draft);
    setDirty(false);
  };
  const toggleDeferredUpdate = (enabled: boolean) => {
    if (!enabled && dirty) applyDraft();
    setDelayUpdate(enabled);
    setDraft(enabled ? cloneLayout(pivot.layout) : null);
    if (enabled) setDirty(false);
  };
  const idsFor = (area: Area): readonly string[] => area === 'filters' ? filters : area === 'columns' ? columns : area === 'rows' ? rows : values.map((field) => field.valueId);
  const changeArea = (fieldId: string, area: Area, index: number) => {
    const field = fields.find((candidate) => candidate.fieldId === fieldId);
    if (field) applyLayout(moveField(layout, field, area, index));
  };
  const removeFromArea = (area: Area, placementId: string) => applyLayout(area === 'values' ? removeValueField(layout, placementId) : removeField(layout, placementId));
  const moveWithinArea = (area: Area, placementId: string, index: number) => applyLayout(area === 'values' ? moveValueField(layout, placementId, index) : (() => { const field = fields.find((candidate) => candidate.fieldId === placementId); return field ? moveField(layout, field, area, index) : layout; })());
  const drop = (area: Area) => (event: DragEvent<HTMLElement>, index?: number) => { event.preventDefault(); const valueId = event.dataTransfer.getData('application/x-pivot-value'); if (valueId && area === 'values') { moveWithinArea(area, valueId, index ?? idsFor(area).length); return; } const fieldId = event.dataTransfer.getData('application/x-pivot-field'); if (fieldId) changeArea(fieldId, area, index ?? idsFor(area).length); };
  const toggle = (fieldId: string, checked: boolean) => {
    if (!checked) { applyLayout(removeField(layout, fieldId)); return; }
    const field = fields.find((candidate) => candidate.fieldId === fieldId);
    if (!field) return;
    const area = defaultPivotFieldArea(field);
    changeArea(fieldId, area, idsFor(area).length);
  };
  const toggleVisible = (fieldIds: readonly string[], checked: boolean) => {
    let next = cloneLayout(layout);
    for (const fieldId of fieldIds) {
      next = removeField(next, fieldId);
      if (!checked) continue;
      const field = fields.find((candidate) => candidate.fieldId === fieldId);
      if (!field) continue;
      const area = defaultPivotFieldArea(field);
      const index = area === 'filters' ? next.filters.length : area === 'columns' ? next.columns.length : area === 'rows' ? next.rows.length : next.values.length;
      next = moveField(next, field, area, index);
    }
    applyLayout(next);
  };
  const filter = (fieldId: string, nextFilter: PivotManualFilterState) => {
    const next = cloneLayout(layout);
    const scope = rows.includes(fieldId) || columns.includes(fieldId) ? 'field' as const : next.filters.find((entry) => entry.kind === 'manual' && entry.fieldId === fieldId)?.scope ?? 'report' as const;
    const index = next.filters.findIndex((entry) => entry.kind === 'manual' && entry.fieldId === fieldId && (entry.scope ?? 'report') === scope);
    const criterion = { kind: 'manual' as const, family: 'manual' as const, fieldId, scope, mode: nextFilter.mode, memberKeys: [...nextFilter.memberKeys] };
    if (index >= 0) next.filters[index] = criterion;
    else next.filters.push(criterion);
    applyLayout(next);
  };
  const sort = (fieldId: string, nextSort: Parameters<NonNullable<PivotPanelCallbacks['onSortChange']>>[1]) => applyLayout({ ...cloneLayout(layout), rows: layout.rows.map((entry) => entry.fieldId === fieldId ? { ...entry, sort: nextSort } : entry), columns: layout.columns.map((entry) => entry.fieldId === fieldId ? { ...entry, sort: nextSort } : entry) });
  const group = (fieldId: string, nextGroup: Parameters<NonNullable<PivotPanelCallbacks['onGroupChange']>>[1]) => applyLayout({ ...cloneLayout(layout), rows: layout.rows.map((entry) => entry.fieldId === fieldId ? { ...entry, group: nextGroup } : entry), columns: layout.columns.map((entry) => entry.fieldId === fieldId ? { ...entry, group: nextGroup } : entry) });
  const subtotal = (fieldId: string, nextSubtotal: Parameters<NonNullable<PivotPanelCallbacks['onSubtotalChange']>>[1]) => applyLayout({ ...cloneLayout(layout), rows: layout.rows.map((entry) => entry.fieldId === fieldId ? { ...entry, subtotal: nextSubtotal } : entry), columns: layout.columns.map((entry) => entry.fieldId === fieldId ? { ...entry, subtotal: nextSubtotal } : entry) });
  const valueChange = (value: PivotValueField) => applyLayout({ ...cloneLayout(layout), values: layout.values.map((entry) => entry.valueId === value.valueId ? value : entry) });
  return (
    <Panel className="flex h-full max-h-full min-h-0 flex-col overflow-hidden rounded-none border-0 bg-white shadow-none" data-testid="pivot-field-list">
      <Inline gap="sm" className="h-14 w-full shrink-0 justify-between px-4">
        <Text size="lg" weight="bold">{pivotText(locale, 'fieldsTitle')}</Text>
        <Inline gap="sm"><Box className="rounded-full border-2 border-[#a529ff] px-3 py-0.5 text-[#8b20e8]"><Text size="sm" weight="bold">AI</Text></Box>{onClose ? <Button aria-label={pivotText(locale, 'close')} icon="x" iconOnly size="sm" variant="ghost" onClick={onClose} /> : null}</Inline>
      </Inline>
      <Stack gap="sm" className="min-h-0 flex-1 overflow-hidden px-4 pb-2">
        <Inline gap="sm" className="h-9 shrink-0 justify-between border-b border-slate-100 pb-1">
          <DropdownMenu align="left" trigger={<Button size="sm" variant="ghost" icon="chevron-down" className="justify-start px-1">{pivotText(locale, 'formatAndOptions')}</Button>}><PivotFormatOptions locale={locale} disabled={disabled} presentation={pivot.presentation} refreshPolicy={pivot.refreshPolicy} onPresentationChange={callbacks?.onPresentationChange} onDisplayOptionsChange={callbacks?.onDisplayOptionsChange} onRefreshPolicyChange={callbacks?.onRefreshPolicyChange} /></DropdownMenu>
          <Select aria-label={pivotText(locale, 'fieldPaneLayout')} sizeVariant="sm" value={fieldPaneLayout} onChange={(event) => setFieldPaneLayout(event.target.value as PivotFieldPaneLayout)}>
            {PIVOT_FIELD_PANE_LAYOUTS.map((mode) => <option key={mode} value={mode}>{pivotText(locale, fieldPaneLayoutLabels[mode])}</option>)}
          </Select>
        </Inline>
        <Box className={`${workspaceClass} min-h-0 min-w-0 flex-1 overflow-hidden gap-2`}>
          {showFields ? (
            <Box className="min-h-0 min-w-0 flex flex-col overflow-hidden" style={sideBySide ? { width: `${showAreas ? fieldPaneSplit : 100}%` } : undefined}>
              <Inline gap="xs" className="mb-1 h-7 shrink-0"><Text size="sm" weight="medium">{pivotText(locale, 'addFields')}</Text><Text size="xs" tone="subtle" className="ml-auto">{pivotTemplate(locale, 'selectedFields', { selected: selected.size, total: fields.length })}</Text></Inline>
              <PivotFieldCatalog className="min-h-[120px] flex-1" locale={locale} fields={fields} selectedFieldIds={selected} disabled={disabled} onToggle={toggle} onToggleVisible={toggleVisible} onDragField={(event, field) => event.dataTransfer.setData('application/x-pivot-field', field.fieldId)} onKeyboardAssign={(field) => { const area = defaultPivotFieldArea(field); changeArea(field.fieldId, area, idsFor(area).length); }} onAssignField={(field, area) => changeArea(field.fieldId, area, idsFor(area).length)} />
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
            <Box className="min-h-0 min-w-0 flex flex-1 flex-col overflow-hidden">
              <Inline gap="xs" className="mb-1 h-7 shrink-0"><Text size="sm" weight="medium">{pivotText(locale, 'dragFields')}</Text><Text size="xs" tone="subtle" className="ml-auto">{pivotText(locale, 'dragOrChoose')}</Text></Inline>
              <Box className={`${fieldPaneLayout === 'areas-1x4' ? 'flex flex-col' : 'grid grid-cols-2 grid-rows-[minmax(0,1fr)_minmax(0,1fr)]'} h-full min-h-0 flex-1 gap-1`}>
                {PIVOT_FIELD_AREAS.map((area) => <PivotFieldArea key={area} locale={locale} area={area} fields={fields} baseFields={fields.filter((field) => idsFor('rows').includes(field.fieldId) || idsFor('columns').includes(field.fieldId))} fieldIds={idsFor(area)} placements={placements} filterStates={currentFilterStates} valueFields={values} disabled={disabled} onDrop={drop(area)} onFilter={filter} onGroup={group} onSubtotal={subtotal} onRemove={(placementId) => removeFromArea(area, placementId)} onMoveByKeyboard={(placementId, itemIndex, direction) => moveWithinArea(area, placementId, itemIndex + direction)} onSort={sort} onValueChange={valueChange} />)}
              </Box>
            </Box>
          ) : null}
        </Box>
        <Inline gap="sm" className="h-11 shrink-0 justify-between border-t border-[#d0d0d0] bg-white pt-1">
          <Stack gap="none" className="min-w-0"><CheckToggle label={pivotText(locale, 'delayUpdate')} checked={delayUpdate} onChange={(event) => toggleDeferredUpdate(event.target.checked)} />{delayUpdate && dirty ? <Text size="xs" tone="subtle" className="truncate text-amber-700">{pivotText(locale, 'pendingLayout')}</Text> : null}</Stack>
          <Button size="sm" variant={dirty ? 'primary' : 'outline'} disabled={!delayUpdate || !dirty} onClick={applyDraft}>{pivotText(locale, 'update')}</Button>
        </Inline>
      </Stack>
      <Inline gap="sm" className="h-12 shrink-0 border-t border-[#d0d0d0] px-4">
        {slicerControls.length > 0 || timelineControls.length > 0 ? <DropdownMenu align="left" trigger={<Button aria-label={pivotText(locale, 'pivotControls')} size="sm" variant="ghost">{pivotText(locale, 'pivotControls')}</Button>}><Stack gap="sm" className="w-[19rem] p-2">{slicerControls.map((control) => { const field = fields.find((candidate) => candidate.fieldId === control.fieldId); const placement = placements.get(control.fieldId); const memberOptions = field && placement?.group ? buildPivotGroupedFilterMembers(field.values ?? [], placement.group) : undefined; return field ? <PivotSlicer key={control.id} locale={locale} field={field} memberOptions={memberOptions} itemProjection={control.items} settings={control.settings} mode={control.mode} memberKeys={control.memberKeys} disabled={disabled} onChange={(next) => callbacks?.onSlicerFilterChange?.(control.id, next)} /> : null; })}{timelineControls.map((control) => { const field = fields.find((candidate) => candidate.fieldId === control.fieldId); return <PivotTimeline key={control.id} locale={locale} fieldLabel={field?.name ?? control.fieldId} values={field?.values} level={control.level} bounds={control.bounds} scrollPosition={control.scrollPosition} showHeader={control.showHeader} showSelectionLabel={control.showSelectionLabel} showTimeLevel={control.showTimeLevel} showHorizontalScrollbar={control.showHorizontalScrollbar} caption={control.caption} styleName={control.styleName} start={control.start} end={control.end} disabled={disabled} onChange={(start, end) => callbacks?.onTimelineRangeChange?.(control.id, start, end)} onClear={() => callbacks?.onTimelineRemove?.()} onLevelChange={(level) => callbacks?.onTimelineLevelChange?.(control.id, level)} onWindowChange={(scrollPosition) => callbacks?.onTimelineWindowChange?.(control.id, scrollPosition)} onDisplayChange={(display) => callbacks?.onTimelineDisplayChange?.(control.id, display)} onCaptionChange={(caption) => callbacks?.onTimelineCaptionChange?.(control.id, caption)} onStyleChange={(styleName) => callbacks?.onTimelineStyleChange?.(control.id, styleName)} />; })}</Stack></DropdownMenu> : null}
        <Select aria-label={pivotText(locale, 'subtotalLocation')} sizeVariant="sm" value={layout.subtotalLocation} disabled={disabled} onChange={(event) => { const next = event.target.value as PivotLayout['subtotalLocation']; if (callbacks?.onSubtotalLocationChange) callbacks.onSubtotalLocationChange(next); else applyLayout({ ...cloneLayout(layout), subtotalLocation: next }); }}><option value="top">{pivotText(locale, 'subtotalTop')}</option><option value="bottom">{pivotText(locale, 'subtotalBottom')}</option><option value="off">{pivotText(locale, 'subtotalOff')}</option></Select>
        <DropdownMenu align="right" trigger={<Button aria-label={pivotText(locale, 'advancedFields')} icon="plus" iconOnly size="sm" variant="ghost" className="ml-auto" />}><Box className="w-[22rem] p-3"><PivotCalculatedEditor locale={locale} fields={fields} calculatedFields={layout.calculatedFields ?? []} calculatedItems={layout.calculatedItems ?? []} disabled={disabled} onFieldsChange={(next) => applyLayout({ ...cloneLayout(layout), calculatedFields: next })} onItemsChange={(next) => applyLayout({ ...cloneLayout(layout), calculatedItems: next })} /></Box></DropdownMenu>
        {slots?.statusSummary ? <Text size="xs" tone="subtle">{slots.statusSummary}</Text> : null}
      </Inline>
    </Panel>
  );
}
