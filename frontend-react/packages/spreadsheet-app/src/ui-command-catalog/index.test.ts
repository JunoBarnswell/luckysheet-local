import assert from 'node:assert/strict';
import { describe, it } from 'node:test';
import type { CommandDescriptor } from '@react-sheets/command-runtime';
import { transformNumberFormatPrecision } from '@react-sheets/number-format';
import {
  buildRibbonCommand,
  getRibbonSurfaces,
  getRibbonCommandDefinition,
  getRibbonGroupDefinition,
  HOME_RIBBON_SURFACES,
  INSERT_RIBBON_SURFACES,
  RIBBON_TAB_SURFACES,
  isRibbonCommandEnabled,
  RIBBON_COMMAND_CATALOG,
  RIBBON_GROUP_CATALOG,
  RIBBON_LAYOUT_SPECS,
  validateRibbonLayoutSpecs,
  type RibbonCommandActions,
  type RibbonCommandContext,
  type RibbonLayoutNode,
} from './index';

function descriptor(commandId: string, params?: unknown): CommandDescriptor {
  return { commandId, ...(params === undefined ? {} : { params }) };
}

function context(overrides: Partial<RibbonCommandContext> = {}): RibbonCommandContext {
  const actions: RibbonCommandActions = {
    onCopy: () => undefined,
    onCut: () => undefined,
    onPaste: () => undefined,
    onUndo: () => undefined,
    onRedo: () => undefined,
    onSave: () => undefined,
    onExportXlsx: () => undefined,
    onImportXlsx: () => undefined,
    onRecalculate: () => undefined,
    onTracePrecedents: () => undefined,
    onTraceDependents: () => undefined,
    onRemoveArrows: () => undefined,
    onToggleShowFormulas: () => undefined,
    onScanFormulaErrors: () => undefined,
    onEvaluateFormula: () => undefined,
    onOpenPrintLayout: () => undefined,
    onSetPrintArea: () => undefined,
    onClearPrintArea: () => undefined,
    onSetPrintTitleRows: () => undefined,
    onSetPrintTitleColumns: () => undefined,
    onSetPrintScale: () => undefined,
    onToggleViewGridlines: () => undefined,
    onTogglePrintGridlines: () => undefined,
    onToggleViewHeadings: () => undefined,
    onTogglePrintHeadings: () => undefined,
    onAutoSum: () => undefined,
    onMerge: () => undefined,
    onFill: () => undefined,
    onFreezeAtPrimary: () => undefined,
    onCreateSheetTable: () => undefined,
    onOpenTableSettings: () => undefined,
    onToggleTableOption: () => undefined,
    onConvertActiveTableToRange: () => undefined,
    onCreateDataSource: () => undefined,
    onToggleSheetTableTotalRow: () => descriptor('sheetTable.update'),
    onApplyFilterSelection: () => descriptor('data.filter.apply'),
    onClearFilter: () => descriptor('data.filter.clear'),
    onGroupRows: () => descriptor('outline.group'),
    onUngroupRows: () => descriptor('outline.ungroup'),
    onGroupColumns: () => descriptor('outline.group'),
    onUngroupColumns: () => descriptor('outline.ungroup'),
    onSubtotal: () => descriptor('data.subtotal'),
    onRemoveDuplicates: () => descriptor('data.removeDuplicates'),
    onTextToColumns: () => descriptor('data.textToColumns'),
    onResolveComment: () => undefined,
    onProtectSelection: () => undefined,
    onUnprotectSelection: () => undefined,
    onShowOutlineLevel: () => undefined,
    onTransposeSelection: () => undefined,
    onFlipSelection: () => undefined,
    onSplitByDelimiter: () => undefined,
    onToggleBandedRows: () => undefined,
    onSetRecalculationMode: () => undefined,
    onOpenDefinedNames: () => undefined,
    onCreateAdvancedSheet: () => undefined,
    onApplyBarcode: () => undefined,
    onCreateDataChart: () => undefined,
    onCreateCamera: () => undefined,
    onCreateFormControl: () => undefined,
    onApplyCheckbox: () => undefined,
    onCreateTextBox: () => undefined,
  };
  return {
    phase: 'ready',
    disabled: false,
    cellStyle: {},
    actions,
    dispatchSessionIntent: () => undefined,
    ...overrides,
  };
}

describe('Ribbon UI command catalog', () => {
  it('keeps command and group identities unique and tab-scoped', () => {
    const commandIds = new Set(RIBBON_COMMAND_CATALOG.map((definition) => definition.id));
    const groupIds = new Set(RIBBON_GROUP_CATALOG.map((definition) => definition.id));
    assert.equal(commandIds.size, RIBBON_COMMAND_CATALOG.length);
    assert.equal(groupIds.size, RIBBON_GROUP_CATALOG.length);
    for (const definition of RIBBON_COMMAND_CATALOG) {
      assert.ok(definition.placements.length > 0);
      for (const placement of definition.placements) {
        assert.ok(groupIds.has(placement.group));
        assert.equal(getRibbonGroupDefinition(placement.group).tab, placement.tab);
      }
    }
  });

  it('keeps every ribbon layout group and command placement resolvable', () => {
    assert.deepEqual(validateRibbonLayoutSpecs(), []);
  });

  it('builds typed commands, callbacks and UI intents from one context', () => {
    let createPivotDialogCalls = 0;
    const current = context({ cellStyle: { bold: false }, openCreatePivotDialog: () => { createPivotDialogCalls += 1; } });
    assert.deepEqual(buildRibbonCommand('bold', current), {
      type: 'command',
      descriptor: { commandId: 'sheet.style.set', params: { style: { bold: true } } },
    });
    assert.deepEqual(buildRibbonCommand('clearContents', current), {
      type: 'command',
      descriptor: { commandId: 'sheet.range.clear', params: { family: 'contents' } },
    });
    assert.deepEqual(buildRibbonCommand('clearCommentsNotes', current), {
      type: 'command',
      descriptor: { commandId: 'sheet.range.clear', params: { family: 'comments-and-notes' } },
    });
    assert.deepEqual(getRibbonSurfaces('home', 'editing', 'wide').filter((surface) => surface.menuId === 'control.clear-menu').map((surface) => surface.commandId), [
      'clearContents', 'clearFormats', 'clearAll', 'clearCommentsNotes', 'clearHyperlinks',
    ]);
    const pivotAction = buildRibbonCommand('pivotTable', current);
    assert.equal(pivotAction?.type, 'callback');
    if (pivotAction?.type === 'callback') pivotAction.invoke();
    assert.equal(createPivotDialogCalls, 1);
    assert.equal(buildRibbonCommand('tableSheet', current)?.type, 'callback');
  });

  it('exposes every canonical cell style preset through the HOME gallery and command palette', () => {
    const presets = [
      ['cellStyleNormal', 'normal'],
      ['cellStyleGood', 'good'],
      ['cellStyleBad', 'bad'],
      ['cellStyleNeutral', 'neutral'],
      ['cellStyleTitle', 'title'],
      ['cellStyleHeading1', 'heading1'],
      ['cellStyleHeading2', 'heading2'],
      ['cellStyleTotal', 'total'],
    ] as const;
    const current = context();
    for (const [commandId, preset] of presets) {
      assert.deepEqual(buildRibbonCommand(commandId, current), {
        type: 'command',
        descriptor: { commandId: 'sheet.style.preset.apply', params: { preset } },
      });
      assert.ok(RIBBON_COMMAND_CATALOG.some((definition) => definition.id === commandId));
    }
    assert.deepEqual(
      getRibbonSurfaces('home', 'styles', 'wide')
        .filter((surface) => surface.menuId === 'control.cell-styles-menu')
        .map((surface) => surface.commandId),
      presets.map(([commandId]) => commandId),
    );
  });

  it('exposes the real PivotTable Analyze and Design actions only for an active Pivot context', () => {
    const calls: string[] = [];
    const activePivot = { sheetId: 'sheet-pivot', pivotId: 'pivot-1' };
    const current = context({
      activePivot,
      pivotActions: {
        onSlicer: () => calls.push('slicer'),
        onTimeline: () => calls.push('timeline'),
        onPivotChart: () => calls.push('chart'),
        onLayoutChange: (layout) => calls.push(`layout:${layout}`),
      },
    });
    assert.equal(getRibbonCommandDefinition('pivotFieldList').placements[0]?.tab, 'pivotAnalyze');
    assert.equal(getRibbonCommandDefinition('pivotLayoutCompact').placements[0]?.tab, 'pivotDesign');
    assert.deepEqual(INSERT_RIBBON_SURFACES.find((surface) => surface.id === 'tables.slicer')?.commandId, 'pivotSlicer');
    assert.deepEqual(getRibbonCommandDefinition('pivotSlicer').placements, [{ tab: 'pivotAnalyze', group: 'pivotAnalyze' }]);
    for (const id of ['pivotSlicer', 'pivotTimeline', 'pivotChart', 'pivotLayoutCompact', 'pivotLayoutOutline', 'pivotLayoutTabular'] as const) {
      assert.equal(isRibbonCommandEnabled(getRibbonCommandDefinition(id), current), true, id);
      const result = buildRibbonCommand(id, current);
      assert.equal(result?.type, 'callback', id);
      if (result?.type === 'callback') result.invoke();
    }
    assert.deepEqual(calls, ['slicer', 'timeline', 'chart', 'layout:compact', 'layout:outline', 'layout:tabular']);
    assert.equal(isRibbonCommandEnabled(getRibbonCommandDefinition('pivotSlicer'), context()), false);
    assert.equal(buildRibbonCommand('pivotLayoutTabular', context()), undefined);
  });

  it('keeps Pivot contextual commands disabled when the canonical permission check rejects them', () => {
    const forbidden = context({
      activePivot: { sheetId: 'sheet-pivot', pivotId: 'pivot-1' },
      pivotActions: { onSlicer: () => undefined, onTimeline: () => undefined, onPivotChart: () => undefined, onLayoutChange: () => undefined },
      canExecute: () => false,
    });
    for (const id of ['pivotRefresh', 'pivotFieldList', 'pivotSlicer', 'pivotTimeline', 'pivotChart', 'pivotLayoutCompact', 'pivotLayoutOutline', 'pivotLayoutTabular'] as const) {
      assert.equal(isRibbonCommandEnabled(getRibbonCommandDefinition(id), forbidden), false, id);
      assert.equal(buildRibbonCommand(id, forbidden), undefined, id);
    }
  });

  it('routes all merge actions through the typed high-level callback', () => {
    const operations: string[] = [];
    const current = context({ actions: { ...context().actions, onMerge: (operation) => { operations.push(operation); } } });
    for (const [command, operation] of [['mergeCenter', 'center'], ['mergeCells', 'cells'], ['mergeAcross', 'across'], ['unmergeCells', 'unmerge']] as const) {
      const action = buildRibbonCommand(command, current);
      assert.equal(action?.type, 'callback');
      if (action?.type === 'callback') action.invoke();
    }
    assert.deepEqual(operations, ['center', 'cells', 'across', 'unmerge']);
    const mergeSurfaces = getRibbonSurfaces('home', 'alignment', 'wide').filter((surface) => surface.menuId === 'control.merge-menu');
    assert.deepEqual(mergeSurfaces.map((surface) => surface.commandId), ['mergeCenter', 'mergeCells', 'mergeAcross', 'unmergeCells']);
  });

  it('exposes TableSheet Designer commands only for the active bound TableSheet', () => {
    const withoutTableSheet = context();
    assert.equal(isRibbonCommandEnabled(getRibbonCommandDefinition('tableSheetFieldList'), withoutTableSheet), false);
    assert.equal(buildRibbonCommand('tableSheetColumnSettings', withoutTableSheet), undefined);

    const withTableSheet = context({ activeTableSheet: { sheetId: 'sheet-table-1', viewId: 'table-1' } });
    assert.equal(isRibbonCommandEnabled(getRibbonCommandDefinition('tableSheetFieldList'), withTableSheet), true);
    assert.deepEqual(buildRibbonCommand('tableSheetFieldList', withTableSheet), {
      type: 'intent',
      intent: { type: 'panel.open', panel: 'data' },
    });
  });

  it('exposes Gantt contextual commands only for the active GanttSheet', () => {
    const withoutGantt = context();
    assert.equal(isRibbonCommandEnabled(getRibbonCommandDefinition('ganttFieldMapping'), withoutGantt), false);
    const withGantt = context({ activeGanttSheet: { sheetId: 'sheet-gantt-1', viewId: 'table-1' } });
    assert.equal(isRibbonCommandEnabled(getRibbonCommandDefinition('ganttFieldMapping'), withGantt), true);
    assert.deepEqual(buildRibbonCommand('ganttTimeline', withGantt), { type: 'intent', intent: { type: 'panel.open', panel: 'data' } });
  });

  it('exposes ReportSheet design commands only for the active ReportSheet', () => {
    assert.equal(isRibbonCommandEnabled(getRibbonCommandDefinition('reportFieldBinding'), context()), false);
    const withReport = context({ activeReportSheet: { sheetId: 'report-1', tableId: 'table-1' } });
    assert.equal(isRibbonCommandEnabled(getRibbonCommandDefinition('reportFieldBinding'), withReport), true);
    assert.deepEqual(buildRibbonCommand('reportPagination', withReport), { type: 'intent', intent: { type: 'panel.open', panel: 'data' } });
  });

  it('exposes Sparkline Design only for the selected sparkline anchor', () => {
    assert.equal(isRibbonCommandEnabled(getRibbonCommandDefinition('sparklineDesign'), context()), false);
    const withSparkline = context({ activeSparkline: { sheetId: 'sheet-1', sparklineId: 'spark-1' } });
    assert.equal(isRibbonCommandEnabled(getRibbonCommandDefinition('sparklineDesign'), withSparkline), true);
    assert.deepEqual(buildRibbonCommand('sparklineDesign', withSparkline), { type: 'intent', intent: { type: 'panel.open', panel: 'sparkline' } });
  });

  it('exposes Shape Format only for selected renderer-backed shapes and reuses drawing commands', () => {
    assert.equal(isRibbonCommandEnabled(getRibbonCommandDefinition('shapeFormatPanel'), context()), false);
    const withOneShape = context({ activeShape: {
      sheetId: 'sheet-1',
      drawingIds: ['shape-1'],
      transforms: [{ drawingId: 'shape-1', transform: { x: 10, y: 20, width: 80, height: 40, rotation: 0 } }],
    } });
    assert.deepEqual(buildRibbonCommand('shapeFormatPanel', withOneShape), { type: 'intent', intent: { type: 'panel.open', panel: 'shape' } });
    const rotate = buildRibbonCommand('shapeRotateClockwise', withOneShape);
    assert.deepEqual(rotate, { type: 'command', descriptor: { commandId: 'drawing.transform.batch', params: { sheetId: 'sheet-1', entries: [{ drawingId: 'shape-1', before: { x: 10, y: 20, width: 80, height: 40, rotation: 0 }, after: { x: 10, y: 20, width: 80, height: 40, rotation: 90 } }] } } });
    assert.equal(buildRibbonCommand('shapeAlignLeft', withOneShape), undefined);
    const withThreeShapes = context({ activeShape: {
      sheetId: 'sheet-1',
      drawingIds: ['shape-1', 'shape-2', 'shape-3'],
      transforms: [
        { drawingId: 'shape-1', transform: { x: 10, y: 20, width: 80, height: 40, rotation: 0 } },
        { drawingId: 'shape-2', transform: { x: 120, y: 50, width: 80, height: 40, rotation: 0 } },
        { drawingId: 'shape-3', transform: { x: 240, y: 80, width: 80, height: 40, rotation: 0 } },
      ],
    } });
    assert.deepEqual(buildRibbonCommand('shapeAlignLeft', withThreeShapes), { type: 'command', descriptor: { commandId: 'drawing.align', params: { sheetId: 'sheet-1', drawingIds: ['shape-1', 'shape-2', 'shape-3'], alignment: 'left' } } });
    assert.deepEqual(buildRibbonCommand('shapeDistributeHorizontal', withThreeShapes), { type: 'command', descriptor: { commandId: 'drawing.distribute', params: { sheetId: 'sheet-1', drawingIds: ['shape-1', 'shape-2', 'shape-3'], axis: 'horizontal' } } });
    assert.deepEqual(buildRibbonCommand('shapeBringForward', withOneShape), { type: 'command', descriptor: { commandId: 'drawing.zorder', params: { sheetId: 'sheet-1', drawingId: 'shape-1', direction: 'forward' } } });
    const forbidden = context({ ...withOneShape, canExecute: () => false });
    assert.equal(buildRibbonCommand('shapeRotateClockwise', forbidden), undefined);
    assert.equal(buildRibbonCommand('shapeBringForward', forbidden), undefined);
    assert.equal(getRibbonSurfaces('shapeFormat', 'shapeFormat', 'wide').length, 16);
  });

  it('honors phase and permission context before building a command', () => {
    const disabled = context({ disabled: true });
    assert.equal(isRibbonCommandEnabled(getRibbonCommandDefinition('bold'), disabled), false);
    assert.equal(buildRibbonCommand('bold', disabled), undefined);

    const forbidden = context({ canExecute: (commandId) => commandId !== 'sheet.style.set' });
    assert.equal(isRibbonCommandEnabled(getRibbonCommandDefinition('bold'), forbidden), false);
    assert.equal(buildRibbonCommand('bold', forbidden), undefined);
    assert.equal(buildRibbonCommand('pivotTable', context()), undefined);
  });

  it('keeps every Home Ribbon surface reachable through one registered command or control', () => {
    const positions = new Set<string>();
    for (const surface of HOME_RIBBON_SURFACES) {
      assert.equal(surface.tab, 'home');
      assert.ok(surface.commandId || surface.controlId);
      assert.equal(positions.has(`${surface.group}:${surface.order}`), false);
      positions.add(`${surface.group}:${surface.order}`);
      if (surface.commandId) assert.ok(getRibbonCommandDefinition(surface.commandId));
    }
    assert.ok(getRibbonSurfaces('home', 'styles', 'compact').some((surface) => surface.commandId === 'cellTemplate'));
  });

  it('keeps the Home layout explicit and renders each visible surface exactly once', () => {
    const collectSurfaceIds = (nodes: readonly RibbonLayoutNode[]): string[] => nodes.flatMap((node) => {
      if (node.kind === 'surface') return [node.surfaceId];
      if ('children' in node) return collectSurfaceIds(node.children);
      return [];
    });
    const layout = RIBBON_LAYOUT_SPECS.home;
    const layoutSurfaceIds = layout.groups.flatMap((group) => collectSurfaceIds(group.children));
    const visibleSurfaceIds = HOME_RIBBON_SURFACES.filter((surface) => !surface.menuId).map((surface) => surface.id);

    assert.equal(new Set(layoutSurfaceIds).size, layoutSurfaceIds.length);
    assert.deepEqual(new Set(layoutSurfaceIds), new Set(visibleSurfaceIds));
    assert.deepEqual(layout.groups.find((group) => group.id === 'history')?.children[0], {
      kind: 'column',
      id: 'history.layout',
      children: [
        { kind: 'surface', id: 'history.undo', surfaceId: 'history.undo' },
        { kind: 'surface', id: 'history.redo', surfaceId: 'history.redo' },
      ],
    });
  });

  it('keeps the complete HOME surface set identical across responsive breakpoints', () => {
    const groups = ['history', 'clipboard', 'font', 'alignment', 'number', 'styles', 'cells', 'editing'] as const;
    const breakpoints = ['wide', 'compact', 'narrow'] as const;
    for (const group of groups) {
      const byBreakpoint = breakpoints.map((breakpoint) => getRibbonSurfaces('home', group, breakpoint).map((surface) => surface.id));
      assert.ok(byBreakpoint.every((ids) => ids.length > 0), `${group} has no surface at one breakpoint`);
      assert.deepEqual(new Set(byBreakpoint[0]), new Set(byBreakpoint[1]), `${group} compact surface drift`);
      assert.deepEqual(new Set(byBreakpoint[0]), new Set(byBreakpoint[2]), `${group} narrow surface drift`);
      for (const surface of getRibbonSurfaces('home', group, 'wide')) {
        assert.equal(Boolean(surface.commandId) !== Boolean(surface.controlId), true, `${surface.id} must be command or control`);
        if (surface.commandId) assert.ok(getRibbonCommandDefinition(surface.commandId), `${surface.id} has no command`);
        if (surface.menuId) assert.ok(getRibbonSurfaces('home', group, 'wide').some((owner) => owner.id === surface.menuId), `${surface.id} has no menu owner`);
      }
    }
  });

  it('keeps Home and Insert surfaces unique and executable', () => {
    const ids = new Set<string>();
    for (const surface of RIBBON_TAB_SURFACES) {
      assert.equal(ids.has(surface.id), false);
      ids.add(surface.id);
      assert.ok(surface.commandId || surface.controlId);
      if (surface.commandId) assert.ok(getRibbonCommandDefinition(surface.commandId));
    }
    assert.equal(INSERT_RIBBON_SURFACES.some((surface) => surface.id.includes('quick')), false);
  });

  it('adjusts decimals through the canonical number-format transformer', () => {
    assert.deepEqual(transformNumberFormatPrecision('$#,##0.00', -1), { ok: true, format: '$#,##0.0', decimalPlaces: 1 });
    assert.deepEqual(transformNumberFormatPrecision('0%', 1), { ok: true, format: '0.0%', decimalPlaces: 1 });
    assert.equal(transformNumberFormatPrecision('general', 1).ok, false);
  });

  it('does not build a style mutation for an unsupported active format', () => {
    const dateContext = context({ cellStyle: { numberFormat: 'yyyy-mm-dd' } });
    assert.equal(buildRibbonCommand('numberFormatDecimalIncrease', dateContext), undefined);
    assert.equal(isRibbonCommandEnabled(getRibbonCommandDefinition('numberFormatDecimalIncrease'), dateContext), false);
  });

  it('builds one canonical style mutation without flattening custom sections', () => {
    const result = buildRibbonCommand('numberFormatDecimalIncrease', context({ cellStyle: { numberFormat: '[Red]#,##0;[Blue]-#,##0' } }));
    assert.deepEqual(result, {
      type: 'command',
      descriptor: { commandId: 'sheet.style.set', params: { style: { numberFormat: '[Red]#,##0.0;[Blue]-#,##0.0' } } },
    });
  });

  it('builds the complete alignment grid as canonical style commands', () => {
    const ready = context({ cellStyle: { indent: 1, verticalAlignment: 'middle' } });
    assert.deepEqual(buildRibbonCommand('alignTop', ready), { type: 'command', descriptor: { commandId: 'sheet.style.set', params: { style: { verticalAlignment: 'top' } } } });
    assert.deepEqual(buildRibbonCommand('alignCenterContinuous', ready), { type: 'command', descriptor: { commandId: 'sheet.style.set', params: { style: { horizontalAlignment: 'centerContinuous' } } } });
    assert.deepEqual(buildRibbonCommand('alignVerticalDistributed', ready), { type: 'command', descriptor: { commandId: 'sheet.style.set', params: { style: { verticalAlignment: 'distributed' } } } });
    assert.deepEqual(buildRibbonCommand('shrinkToFit', context({ cellStyle: { shrinkToFit: false } })), { type: 'command', descriptor: { commandId: 'sheet.style.set', params: { style: { shrinkToFit: true } } } });
    assert.deepEqual(buildRibbonCommand('orientationStacked', ready), { type: 'command', descriptor: { commandId: 'sheet.style.set', params: { style: { textOrientation: 'stacked' } } } });
    assert.deepEqual(buildRibbonCommand('indentIncrease', ready), { type: 'command', descriptor: { commandId: 'sheet.style.set', params: { style: { indent: 2 } } } });
    assert.deepEqual(buildRibbonCommand('indentDecrease', ready), { type: 'command', descriptor: { commandId: 'sheet.style.set', params: { style: { indent: 0 } } } });
    assert.deepEqual(getRibbonSurfaces('home', 'alignment', 'wide').filter((surface) => surface.menuId === 'control.alignment-menu').map((surface) => surface.commandId), [
      'alignGeneral', 'alignCenterContinuous', 'alignJustify', 'alignDistributed', 'alignFill',
    ]);
  });
});
