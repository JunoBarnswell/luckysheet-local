import { useMemo, useState } from "react";
import type { CommandDescriptor } from "@react-sheets/command-runtime";
import type {
  PivotAggregateFunction,
  PivotFilter,
  PivotFilterFamily,
  PivotFieldDefinition,
  PivotLayout,
  PivotDisplayOptions,
  PivotSource,
  PivotSort,
  DefinedNameModel,
} from "@react-sheets/core-model";
import { DEFAULT_PIVOT_STYLE_OPTIONS } from "@react-sheets/core-model";
import {
  type SelectionState,
  type SidebarPanelId,
  type UiSessionIntent,
  type UiSnapshot,
  type WorkbookSession,
  richTextSelectionHasFlag,
} from "@react-sheets/spreadsheet-app";
import type { RibbonPivotActions } from "@react-sheets/spreadsheet-app";
import { parseRangeInput } from "../domain/range-input";
import type { PivotPanelCallbacks, PivotPanelState, PivotSlicerControl, PivotTimelineControl } from "../components/pivot/pivot-contract";
import type { Locale } from '../i18n';
import { pivotDefinedNameScopeText, pivotText } from '../components/pivot/pivot-localization';

export interface EditorCommandControllerOptions {
  session: WorkbookSession;
  state: UiSnapshot;
  locale: Locale;
  dispatchCommand: (descriptor: CommandDescriptor) => void;
  dispatchSessionIntent: (intent: UiSessionIntent) => void;
}

export interface EditorCommandController {
  selectedRange: SelectionState["ranges"][number] | undefined;
  currentDataRange: ReturnType<WorkbookSession["getCurrentRegion"]>;
  sortColumns: UiSnapshot["selectedSheet"]["columns"];
  pivotSourceRange: ReturnType<WorkbookSession["getCurrentRegion"]>;
  activePivotId: string | undefined;
  setActivePivotId: (pivotId: string | undefined) => void;
  activePivot: UiSnapshot["selectedSheet"]["pivots"][number] | undefined;
  activePivotSheetId: string;
  activePivotSourceRange: ReturnType<WorkbookSession["getCurrentRegion"]> | undefined;
  pivotFields: PivotFieldDefinition[];
  pivotSlicerControls: PivotSlicerControl[];
  pivotTimelineControls: PivotTimelineControl[];
  pivotPanelState: PivotPanelState;
  pivotCallbacks: PivotPanelCallbacks;
  pivotRibbonActions: RibbonPivotActions;
  pivotSourceOptions: PivotSourceOption[];
  selectedDrawing: UiSnapshot["selectedSheet"]["drawings"][number] | undefined;
  buildTotalRowCommand: () => CommandDescriptor | undefined;
  buildSubtotalCommand: () => CommandDescriptor;
  buildRemoveDuplicatesCommand: () => CommandDescriptor;
  buildTextToColumnsCommand: () => CommandDescriptor;
  buildOutlineCommand: (axis: "row" | "column", action: "add" | "remove") => CommandDescriptor | undefined;
  buildFilterSelectionCommand: () => CommandDescriptor;
  buildClearFilterCommand: () => CommandDescriptor;
  buildSortDescriptor: (ascending: boolean) => CommandDescriptor | undefined;
  createPivotFromDialog: (request: { sourceId: string; destination: "new-sheet" | "existing-sheet"; targetReference?: string }) => Promise<void>;
  executeShortcut: (id: string) => boolean;
  selectPanel: (panel: SidebarPanelId) => void;
  applySelection: (selection: SelectionState) => void;
  applyPivotHeaderFilter: (pivotId: string, fieldId: string, filter: PivotFilter | undefined, sort: PivotSort | undefined, scope: 'report' | 'field', family: PivotFilterFamily | 'all') => void;
}

type RangeLike = ReturnType<WorkbookSession["getCurrentRegion"]>;

export interface PivotSourceOption {
  id: string;
  label: string;
  source: PivotSource;
}

export interface PivotSourceOptionInput {
  currentDataRange: RangeLike;
  currentSheetName: string;
  sheetTables: ReadonlyArray<Pick<UiSnapshot["selectedSheet"]["sheetTables"][number], "id" | "name">>;
  definedNameModels: readonly DefinedNameModel[];
  sheetNames: ReadonlyMap<string, string>;
  locale: Locale;
}

function definedNameOptionId(definedName: DefinedNameModel): string {
  if (!definedName.name.trim()) throw new Error('Defined name is required');
  const normalizedName = definedName.name.trim().toLocaleLowerCase();
  if (definedName.scope === 'workbook') {
    if (definedName.sheetId !== undefined) throw new Error(`Workbook-scoped defined name ${definedName.name} cannot specify sheetId`);
    return `name:workbook:${encodeURIComponent(normalizedName)}`;
  }
  if (definedName.scope !== 'sheet') throw new Error(`Defined name ${definedName.name} has an unsupported scope`);
  if (!definedName.sheetId) throw new Error(`Sheet-scoped defined name ${definedName.name} requires a sheetId`);
  return `name:sheet:${encodeURIComponent(definedName.sheetId)}:${encodeURIComponent(normalizedName)}`;
}

function definedNameSource(definedName: DefinedNameModel, sheetNames: ReadonlyMap<string, string>): PivotSource {
  if (!definedName.name.trim()) throw new Error('Defined name is required');
  if (definedName.scope === 'workbook') {
    if (definedName.sheetId !== undefined) throw new Error(`Workbook-scoped defined name ${definedName.name} cannot specify sheetId`);
    return { kind: 'named-range', name: definedName.name };
  }
  if (definedName.scope !== 'sheet') throw new Error(`Defined name ${definedName.name} has an unsupported scope`);
  if (!definedName.sheetId) throw new Error(`Sheet-scoped defined name ${definedName.name} requires a sheetId`);
  if (!sheetNames.has(definedName.sheetId)) throw new Error(`Sheet-scoped defined name ${definedName.name} references unknown sheet ${definedName.sheetId}`);
  return { kind: 'named-range', name: definedName.name, sheetId: definedName.sheetId };
}

export function buildPivotSourceOptions(input: PivotSourceOptionInput): PivotSourceOption[] {
  const options: PivotSourceOption[] = [{
    id: "current-region",
    label: `${pivotText(input.locale, 'currentRegion')} · ${input.currentSheetName}`,
    source: { kind: "worksheet-range", range: input.currentDataRange },
  }];
  for (const table of input.sheetTables) {
    options.push({ id: `sheet-table:${table.id}`, label: `${pivotText(input.locale, 'tableSource')} · ${table.name}`, source: { kind: "table", tableId: table.id } });
  }
  for (const definedName of input.definedNameModels) {
    const source = definedNameSource(definedName, input.sheetNames);
    const scope = pivotDefinedNameScopeText(input.locale, definedName.scope, definedName.sheetId ? input.sheetNames.get(definedName.sheetId) : undefined);
    options.push({
      id: definedNameOptionId(definedName),
      label: `${pivotText(input.locale, 'namedRange')} · ${definedName.name} · ${scope}`,
      source,
    });
  }
  return options;
}

function createWebId(prefix: string): string {
  if (typeof crypto !== "undefined" && typeof crypto.randomUUID === "function") return `${prefix}-${crypto.randomUUID()}`;
  return `${prefix}-${Date.now().toString(36)}-${Math.random().toString(36).slice(2, 8)}`;
}

function cloneLayout(layout: PivotLayout): PivotLayout {
  return structuredClone(layout);
}

function removePivotField(layout: PivotLayout, fieldId: string): PivotLayout {
  const next = cloneLayout(layout);
  next.filters = next.filters.filter((filter) => filter.fieldId !== fieldId);
  next.rows = next.rows.filter((field) => field.fieldId !== fieldId);
  next.columns = next.columns.filter((field) => field.fieldId !== fieldId);
  next.values = next.values.filter((value) => value.fieldId !== fieldId);
  return next;
}

/**
 * Owns selection-derived command construction and Pivot panel actions.
 * The route only wires these typed actions into visual hosts; it never creates
 * chart, shape or pivot drawing domain objects itself.
 */
export function useEditorCommandController({
  session,
  state,
  locale,
  dispatchCommand,
  dispatchSessionIntent,
}: EditorCommandControllerOptions): EditorCommandController {
  const [activePivotId, setActivePivotId] = useState<string>();
  const selectedRange = state.selection.ranges[state.selection.primaryRangeIndex] ?? state.selection.ranges[0];
  const dataRegionContext = session.getDataRegionContext();
  const currentDataRange = dataRegionContext.range;
  const sortColumns = state.selectedSheet.columns.slice(currentDataRange.startColumn, currentDataRange.endColumn + 1);
  const pivotSourceRange = selectedRange && (selectedRange.endRow > selectedRange.startRow || selectedRange.endColumn > selectedRange.startColumn)
    ? selectedRange
    : state.selectedSheet.usedRange;

  const selectedDrawing = state.selectedFloatingId
    ? state.selectedSheet.drawings.find((drawing) => drawing.id === state.selectedFloatingId)
    : undefined;

  const activePivot = state.selectedSheet.pivots.find((pivot) => pivot.id === activePivotId) ?? state.selectedSheet.pivots[0];
  const pivotTree = activePivot ? state.selectedSheet.pivotResults[activePivot.id] : undefined;
  const pivotFields: PivotFieldDefinition[] = pivotTree?.fields.fields ?? session.getPivotFieldCatalog(pivotSourceRange);
  const activePivotSheetId = activePivot ? activePivot.target.sheetId : state.activeSheetId;
  const activePivotSourceRange = activePivot?.source.kind === "worksheet-range" ? activePivot.source.range : undefined;
  const pivotControlRecords = activePivot ? session.listPivotControls(activePivot.id) : [];
  const pivotSlicerControls = pivotControlRecords.flatMap((record) => record.payload.kind === "slicer"
    ? [{ id: record.drawing.id, pivotId: record.payload.pivotId, fieldId: record.payload.fieldId, mode: record.payload.filter.mode, memberKeys: record.payload.filter.memberKeys, settings: record.payload.settings, items: state.selectedSheet.pivotResults[record.payload.pivotId]?.slicerItems?.[record.drawing.id] ?? [], connections: record.payload.connections }]
    : []);
  const pivotTimelineControls = pivotControlRecords.flatMap((record) => record.payload.kind === "timeline"
    ? [{ id: record.drawing.id, pivotId: record.payload.pivotId, fieldId: record.payload.fieldId, start: record.payload.period.start, end: record.payload.period.end, level: record.payload.level, selectionLevel: record.payload.selectionLevel, showHeader: record.payload.showHeader, showSelectionLabel: record.payload.showSelectionLabel, showTimeLevel: record.payload.showTimeLevel, showHorizontalScrollbar: record.payload.showHorizontalScrollbar, scrollPosition: record.payload.scrollPosition, bounds: record.payload.bounds, filterType: record.payload.filterType, caption: record.payload.caption, styleName: record.payload.styleName, connections: record.payload.connections }]
    : []);

  const buildTotalRowCommand = (): CommandDescriptor | undefined => {
    const table = state.selectedSheet.sheetTables.find((entry) =>
      pivotSourceRange.startRow >= entry.range.startRow
      && pivotSourceRange.startRow <= entry.range.endRow
      && pivotSourceRange.startColumn >= entry.range.startColumn
      && pivotSourceRange.startColumn <= entry.range.endColumn);
    return table ? { commandId: "sheetTable.toggleTotalRow", params: { sheetId: state.activeSheetId, tableId: table.id, enabled: !table.hasTotalRow } } : undefined;
  };
  const buildSubtotalCommand = (): CommandDescriptor => ({
    commandId: "data.subtotal",
    params: { sheetId: state.activeSheetId, range: pivotSourceRange, groupColumn: pivotSourceRange.startColumn, valueColumn: pivotSourceRange.startColumn + 1, functionName: "SUM" },
  });
  const buildRemoveDuplicatesCommand = (): CommandDescriptor => ({
    commandId: "data.removeDuplicates",
    params: { sheetId: state.activeSheetId, range: pivotSourceRange, columns: Array.from({ length: pivotSourceRange.endColumn - pivotSourceRange.startColumn + 1 }, (_, index) => pivotSourceRange.startColumn + index), hasHeader: true },
  });
  const buildTextToColumnsCommand = (): CommandDescriptor => ({
    commandId: "data.textToColumns",
    params: { sheetId: state.activeSheetId, range: { ...pivotSourceRange, endColumn: pivotSourceRange.startColumn }, delimiter: ",", maxColumns: 8 },
  });
  const buildOutlineCommand = (axis: "row" | "column", action: "add" | "remove"): CommandDescriptor | undefined => {
    const start = axis === "row" ? pivotSourceRange.startRow : pivotSourceRange.startColumn;
    const end = axis === "row" ? pivotSourceRange.endRow : pivotSourceRange.endColumn;
    if (action === "add") {
      if (end <= start) return undefined;
      return { commandId: "outline.group.add", params: { sheetId: state.activeSheetId, group: { id: createWebId("outline"), axis, start, end, level: 1, collapsed: false } } };
    }
    const group = state.selectedSheet.outlineGroups.find((entry) => entry.axis === axis && entry.start >= start && entry.end <= end);
    return group ? { commandId: "outline.group.remove", params: { sheetId: state.activeSheetId, groupId: group.id } } : undefined;
  };
  const activeFilterOwner = dataRegionContext.owner.kind === 'sheet-table'
    ? { kind: 'table' as const, tableId: dataRegionContext.owner.tableId }
    : { kind: 'worksheet' as const };
  const activeAutoFilter = state.selectedSheet.getActiveAutoFilter(state.selection.activeCell.column);
  const buildFilterSelectionCommand = (): CommandDescriptor => activeFilterOwner?.kind === 'table'
    ? { commandId: 'sheetTable.autoFilter.set', params: { sheetId: state.activeSheetId, tableId: activeFilterOwner.tableId, dataRegionContext } }
    : { commandId: "sheet.autoFilter.toggle", params: { sheetId: state.activeSheetId, range: currentDataRange, dataRegionContext } };
  const filterRange = activeAutoFilter?.range ?? currentDataRange;
  const buildClearFilterCommand = (): CommandDescriptor => activeFilterOwner?.kind === 'table'
    ? { commandId: 'sheetTable.autoFilter.set', params: { sheetId: state.activeSheetId, tableId: activeFilterOwner.tableId, dataRegionContext } }
    : { commandId: "sheet.autoFilter.clearCriteria", params: { sheetId: state.activeSheetId, range: filterRange, dataRegionContext } };
  const buildSortDescriptor = (ascending: boolean): CommandDescriptor | undefined => {
    const range = dataRegionContext.range;
    if (range.endRow <= range.startRow) return undefined;
    return { commandId: "data.sort.quick", params: { sheetId: state.activeSheetId, range, sortColumn: state.selection.activeCell.column, ascending, hasHeader: dataRegionContext.header.kind === 'present', dataRegionContext } };
  };

  const sheetNames = useMemo(() => new Map(state.sheets.map((sheet) => [sheet.id, sheet.name] as const)), [state.sheets]);
  const pivotSourceOptions = useMemo(() => buildPivotSourceOptions({
    currentDataRange,
    currentSheetName: state.selectedSheet.name,
    sheetTables: state.selectedSheet.sheetTables,
    definedNameModels: state.definedNameModels,
    sheetNames,
    locale,
  }), [currentDataRange, locale, sheetNames, state.definedNameModels, state.selectedSheet.name, state.selectedSheet.sheetTables, state.version]);

  const createPivotFromDialog = async (request: { sourceId: string; destination: "new-sheet" | "existing-sheet"; targetReference?: string }) => {
    const source = pivotSourceOptions.find((option) => option.id === request.sourceId)?.source;
    if (!source) {
      session.notify(pivotText(locale, 'invalidSource'));
      return;
    }
    if (request.destination === "new-sheet") {
      const outcome = await session.createPivotTable({ source, destination: { kind: "new-sheet" } });
      if (outcome.status === 'created') session.closeCreatePivotDialog();
      return;
    }
    const target = parseRangeInput(request.targetReference ?? "", state.activeSheetId);
    if (!target) {
      session.notify(pivotText(locale, 'invalidDestination'));
      return;
    }
    const outcome = await session.createPivotTable({ source, destination: { kind: "existing-sheet", sheetId: state.activeSheetId, anchor: { row: target.startRow, column: target.startColumn } } });
    if (outcome.status === 'created') session.closeCreatePivotDialog();
  };

  const updatePivotLayout = (nextLayout: PivotLayout) => {
    if (!activePivot) return Promise.resolve(false);
    return session.updatePivotLayout(activePivot.id, nextLayout).then((outcome) => {
      if (outcome.status === 'updated') session.notify(pivotText(locale, 'layoutUpdated'));
      return outcome.status === 'updated';
    });
  };
  const setPivotReportLayout = (reportLayout: PivotLayout['reportLayout']) => {
    if (activePivot) updatePivotLayout({ ...cloneLayout(activePivot.layout), reportLayout });
  };
  const createPivotSlicer = () => {
    const fieldId = pivotFields[0]?.fieldId;
    if (activePivot && fieldId) session.createPivotSlicerControl(activePivot.id, fieldId);
  };
  const createPivotTimeline = () => {
    const fieldId = pivotFields.find((field) => field.dataType === 'date')?.fieldId;
    if (activePivot && fieldId) session.createPivotTimelineControl(activePivot.id, fieldId);
  };
  const createPivotChart = () => {
    if (!activePivot) return;
    session.createPivotChart(activePivot.id, pivotText(locale, 'pivotChart'));
  };
  const removePivotTimeline = () => {
    for (const control of pivotControlRecords.filter((record) => record.payload.kind === "timeline")) session.removePivotControl(control.drawing.id);
  };
  const pivotCallbacks: PivotPanelCallbacks = {
    onCreate: () => dispatchSessionIntent({ type: "dialog.open", dialog: "create-pivot" }),
    onPivotSelect: setActivePivotId,
    onFieldAreaChange: (fieldId, area, index) => {
      if (!activePivot) return;
      const next = area === "values" ? cloneLayout(activePivot.layout) : removePivotField(activePivot.layout, fieldId);
      if (area === "values") {
        const field = pivotFields.find((entry) => entry.fieldId === fieldId);
        const summarizeBy: PivotAggregateFunction = field?.dataType === "number" ? "sum" : "count";
        const baseValueId = `value:${fieldId}`;
        let valueId = baseValueId;
        let suffix = 2;
        while (next.values.some((value) => value.valueId === valueId)) valueId = `${baseValueId}:${suffix++}`;
        next.values.splice(Math.max(0, index), 0, { valueId, fieldId, summarizeBy });
      } else if (area === "filters") next.filters.splice(Math.max(0, index), 0, { kind: "manual", family: "manual", fieldId, scope: 'report', mode: "all", memberKeys: [] });
      else next[area].splice(Math.max(0, index), 0, { fieldId });
      updatePivotLayout(next);
    },
    onRemoveField: (placementId, area) => {
      if (!activePivot) return;
      const next = cloneLayout(activePivot.layout);
      if (area === "values") next.values = next.values.filter((value) => value.valueId !== placementId);
      else if (area === "filters") next.filters = next.filters.filter((filter) => filter.fieldId !== placementId);
      else next[area] = next[area].filter((field) => field.fieldId !== placementId);
      updatePivotLayout(next);
    },
    onValueChange: (value) => {
      if (!activePivot) return;
      const next = cloneLayout(activePivot.layout);
      const index = next.values.findIndex((entry) => entry.fieldId === value.fieldId);
      if (index >= 0) {
        next.values[index] = structuredClone(value);
        updatePivotLayout(next);
      }
    },
    onCalculatedFieldsChange: (fields) => { if (activePivot) updatePivotLayout({ ...cloneLayout(activePivot.layout), calculatedFields: fields.map((field) => ({ ...field })) }); },
    onCalculatedItemsChange: (items) => { if (activePivot) updatePivotLayout({ ...cloneLayout(activePivot.layout), calculatedItems: items.map((item) => ({ ...item })) }); },
    onFilterChange: (fieldId, filter) => {
      if (!activePivot) return;
      const next = cloneLayout(activePivot.layout);
      const existing = next.filters.find((entry) => entry.fieldId === fieldId);
      if (existing?.kind === "manual") {
        existing.mode = filter.mode;
        existing.memberKeys = [...filter.memberKeys];
      } else next.filters.push({ kind: "manual", family: "manual", fieldId, mode: filter.mode, memberKeys: [...filter.memberKeys] });
      updatePivotLayout(next);
      session.notify(pivotText(locale, 'filterUpdated'));
    },
    onSortChange: (fieldId, sort) => { if (activePivot) updatePivotLayout({ ...cloneLayout(activePivot.layout), rows: activePivot.layout.rows.map((field) => field.fieldId === fieldId ? { ...field, sort } : field), columns: activePivot.layout.columns.map((field) => field.fieldId === fieldId ? { ...field, sort } : field) }); },
    onGroupChange: (fieldId, group) => { if (activePivot) updatePivotLayout({ ...cloneLayout(activePivot.layout), rows: activePivot.layout.rows.map((field) => field.fieldId === fieldId ? { ...field, group } : field), columns: activePivot.layout.columns.map((field) => field.fieldId === fieldId ? { ...field, group } : field) }); },
    onSubtotalChange: (fieldId, subtotal) => { if (activePivot) updatePivotLayout({ ...cloneLayout(activePivot.layout), rows: activePivot.layout.rows.map((field) => field.fieldId === fieldId ? { ...field, subtotal } : field), columns: activePivot.layout.columns.map((field) => field.fieldId === fieldId ? { ...field, subtotal } : field) }); },
    onSubtotalLocationChange: (subtotalLocation) => { if (activePivot) updatePivotLayout({ ...cloneLayout(activePivot.layout), subtotalLocation }); },
    onLayoutReplace: (layout) => activePivot ? updatePivotLayout(cloneLayout(layout)) : false,
    onPresentationChange: (presentation) => { if (activePivot) void session.updatePivotConfiguration(activePivot.id, { presentation: structuredClone(presentation) }); },
    onDisplayOptionsChange: (displayOptions: PivotDisplayOptions) => {
      if (!activePivot) return;
      const current = activePivot.presentation;
      void session.updatePivotConfiguration(activePivot.id, { presentation: {
          ...(current?.styleName ? { styleName: current.styleName } : {}),
          styleOptions: { ...DEFAULT_PIVOT_STYLE_OPTIONS, ...(current?.styleOptions ?? {}) },
          displayOptions: structuredClone(displayOptions),
        } });
    },
    onRefreshPolicyChange: (refreshPolicy) => { if (activePivot) void session.updatePivotConfiguration(activePivot.id, { refreshPolicy: structuredClone(refreshPolicy) }); },
    onTimelineRemove: removePivotTimeline,
    onSlicerFilterChange: (slicerId, filter) => session.setPivotSlicerFilter(slicerId, filter.mode, filter.memberKeys),
    onTimelineRangeChange: (timelineId, start, end) => session.setPivotTimelinePeriod(timelineId, start || undefined, end || undefined),
    onTimelineLevelChange: (timelineId, level) => session.setPivotTimelineLevel(timelineId, level),
    onTimelineWindowChange: (timelineId, scrollPosition) => session.setPivotTimelineWindow(timelineId, scrollPosition),
    onTimelineDisplayChange: (timelineId, display) => session.setPivotTimelineDisplay(timelineId, display),
    onTimelineCaptionChange: (timelineId, caption) => session.setPivotTimelineCaption(timelineId, caption),
    onTimelineStyleChange: (timelineId, styleName) => session.setPivotTimelineStyle(timelineId, styleName),
  };
  const pivotRibbonActions: RibbonPivotActions = {
    onSlicer: createPivotSlicer,
    onTimeline: createPivotTimeline,
    onPivotChart: createPivotChart,
    onLayoutChange: setPivotReportLayout,
  };

  const activePivotTask = activePivot ? state.pivotTaskStates[activePivot.id] : undefined;
  const pivotPanelState: PivotPanelState = {
    disabled: state.phase !== "ready",
    loading: state.phase === "loading" || activePivotTask?.status === 'running',
    error: activePivotTask?.status === 'failed' ? `${activePivotTask.error.code}: ${activePivotTask.error.message}` : state.phase === "error" ? pivotText(locale, 'error') : undefined,
    empty: pivotFields.length === 0,
  };

  const applyPivotHeaderFilter = (pivotId: string, fieldId: string, filter: PivotFilter | undefined, sort: PivotSort | undefined, scope: 'report' | 'field', family: PivotFilterFamily | 'all') => {
    const targetPivot = state.selectedSheet.pivots.find((candidate) => candidate.id === pivotId);
    if (!targetPivot) return;
    const layout = cloneLayout(targetPivot.layout);
    layout.filters = layout.filters.filter((candidate) => candidate.fieldId !== fieldId || (candidate.scope ?? 'report') !== scope || (family !== 'all' && candidate.family !== family));
    if (filter) layout.filters.push(structuredClone(filter));
    layout.rows = layout.rows.map((placement) => placement.fieldId === fieldId ? { ...placement, sort } : placement);
    layout.columns = layout.columns.map((placement) => placement.fieldId === fieldId ? { ...placement, sort } : placement);
    void session.updatePivotLayout(pivotId, layout);
  };

  const executeShortcut = (id: string): boolean => {
    const toggleRichTextFlag = (key: 'bold' | 'italic' | 'underline'): boolean => {
      const edit = session.cellEdit.getSnapshot().session;
      if (!edit || edit.caret.start === edit.caret.end) return false;
      const enabled = !richTextSelectionHasFlag(edit.draft, edit.caret, key);
      const style = key === 'bold' ? { bold: enabled } : key === 'italic' ? { italic: enabled } : { underline: enabled };
      session.cellEdit.dispatch({ type: 'rich-text.format', style });
      return true;
    };
    switch (id) {
      case "history.undo": session.undo(); return true;
      case "history.redo": session.redo(); return true;
      case "history.repeat": session.repeatLastCommand(); return true;
      case "clipboard.copy": session.copy(); return true;
      case "clipboard.cut": session.cut(); return true;
      case "clipboard.paste": session.paste(); return true;
      case "clipboard.pasteSpecial": dispatchSessionIntent({ type: "dialog.open", dialog: "paste-special" }); return true;
      case "clipboard.cancel": session.clearClipboard(); session.cancelFormatPainter(); return true;
      case "workbook.save": void session.saveWorkbook("Keyboard shortcut"); return true;
      case "format.bold": if (!toggleRichTextFlag('bold')) dispatchCommand({ commandId: "sheet.style.set", params: { style: { bold: !state.homeRibbon.style.bold } } }); return true;
      case "format.italic": if (!toggleRichTextFlag('italic')) dispatchCommand({ commandId: "sheet.style.set", params: { style: { italic: !state.homeRibbon.style.italic } } }); return true;
      case "format.underline": if (!toggleRichTextFlag('underline')) dispatchCommand({ commandId: "sheet.style.set", params: { style: { underline: !state.homeRibbon.style.underline } } }); return true;
      case "range.fillDown": session.fillSelection("down"); return true;
      case "range.clearContents": session.clearSelection("contents"); return true;
      case "cells.insert": dispatchSessionIntent({ type: "dialog.open", dialog: "shift-cells", operation: "insert" }); return true;
      case "cells.delete": dispatchSessionIntent({ type: "dialog.open", dialog: "shift-cells", operation: "delete" }); return true;
      case "filter.toggle": session.applyFilterSelection(); return true;
      case "find.open": dispatchSessionIntent({ type: "dialog.open", dialog: "find-replace", findMode: "find" }); return true;
      case "replace.open": dispatchSessionIntent({ type: "dialog.open", dialog: "find-replace", findMode: "replace" }); return true;
      case "commandPalette.open": dispatchSessionIntent({ type: "command-palette.open" }); return true;
      case "name.goto": case "navigation.goto": dispatchSessionIntent({ type: "dialog.open", dialog: "goto" }); return true;
      case "ribbon.home.keyTips": session.setRibbonTab("home"); session.notify("Home shortcuts are active"); return true;
      case "format.cells": dispatchSessionIntent({ type: "dialog.open", dialog: "format-cells" }); return true;
      case "hyperlink.insert": dispatchSessionIntent({ type: "dialog.open", dialog: "hyperlink" }); return true;
      case "row.select": session.selectActiveRow(); return true;
      case "column.select": session.selectActiveColumn(); return true;
      case "sheet.previous": session.selectAdjacentSheet("previous"); return true;
      case "sheet.next": session.selectAdjacentSheet("next"); return true;
      case "formula.autoSum": session.autoSum(); return true;
      case "edit.begin": session.cellEdit.dispatch({ type: 'begin.request', source: 'f2', surface: 'grid' }); return true;
      case "drawing.remove": session.removeSelectedDrawing(); return true;
      case "formula.functionWizard": dispatchSessionIntent({ type: "dialog.open", dialog: "function-wizard" }); return true;
      case "formula.calculate": void session.recalculateFormulas(); return true;
      case "pivot.refresh": {
        const pivotId = state.activeContext.kind === "pivot" ? state.activeContext.pivotId : undefined;
        if (!pivotId) return false;
        session.refreshPivot(pivotId);
        return true;
      }
      default: return false;
    }
  };

  const selectPanel = (panel: SidebarPanelId) => {
    const panelNotice = panel === "inspector" ? "Select a cell and use Review tools for comments." : undefined;
    if (panel === "print") {
      dispatchSessionIntent({ type: "dialog.open", dialog: "print-preview" });
      return;
    }
    if (panelNotice) {
      dispatchSessionIntent({ type: "panel.open", panel, notice: panelNotice });
      return;
    }
    if (panel === "selectionPane") session.setDrawingSelectionMode(true);
    dispatchSessionIntent({ type: "panel.open", panel });
  };

  return {
    selectedRange,
    currentDataRange,
    sortColumns,
    pivotSourceRange,
    activePivotId,
    setActivePivotId,
    activePivot,
    activePivotSheetId,
    activePivotSourceRange,
    pivotFields,
    pivotSlicerControls,
    pivotTimelineControls,
    pivotPanelState,
    pivotCallbacks,
    pivotRibbonActions,
    pivotSourceOptions,
    selectedDrawing,
    buildTotalRowCommand,
    buildSubtotalCommand,
    buildRemoveDuplicatesCommand,
    buildTextToColumnsCommand,
    buildOutlineCommand,
    buildFilterSelectionCommand,
    buildClearFilterCommand,
    buildSortDescriptor,
    createPivotFromDialog,
    executeShortcut,
    selectPanel,
    applySelection: (selection) => session.applyCanvasSelection(selection),
    applyPivotHeaderFilter,
  };
}

export type EditorRange = RangeLike;
