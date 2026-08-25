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
  type RibbonCommandActions,
  type RibbonCommandContext,
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
    onFill: () => undefined,
    onFreezeAtPrimary: () => undefined,
    onCreateSheetTable: () => undefined,
    onOpenTableSettings: () => undefined,
    onToggleTableOption: () => undefined,
    onConvertActiveTableToRange: () => undefined,
    onCreateDataTable: () => undefined,
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
    sampleAutomationScript: 'return 1;',
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
      assert.ok(groupIds.has(definition.group));
      assert.equal(getRibbonGroupDefinition(definition.group).tab, definition.tab);
    }
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
      descriptor: { commandId: 'sheet.range.clear', params: { mode: 'contents' } },
    });
    const pivotAction = buildRibbonCommand('pivotTable', current);
    assert.equal(pivotAction?.type, 'callback');
    if (pivotAction?.type === 'callback') pivotAction.invoke();
    assert.equal(createPivotDialogCalls, 1);
    assert.equal(buildRibbonCommand('tableSheet', current)?.type, 'callback');
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
    assert.deepEqual(buildRibbonCommand('indentIncrease', ready), { type: 'command', descriptor: { commandId: 'sheet.style.set', params: { style: { indent: 2 } } } });
    assert.deepEqual(buildRibbonCommand('indentDecrease', ready), { type: 'command', descriptor: { commandId: 'sheet.style.set', params: { style: { indent: 0 } } } });
  });
});
