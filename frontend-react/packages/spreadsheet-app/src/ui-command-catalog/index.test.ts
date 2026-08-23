import assert from 'node:assert/strict';
import { describe, it } from 'node:test';
import type { CommandDescriptor } from '@react-sheets/command-runtime';
import {
  buildRibbonCommand,
  getRibbonCommandDefinition,
  getRibbonGroupDefinition,
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
    onFreezeAtPrimary: () => undefined,
    onCreatePivot: () => descriptor('pivot.add', { id: 'pivot-1' }),
    onCreateChart: () => descriptor('chart.insert'),
    onCreateSparkline: () => descriptor('sparkline.add'),
    onCreateShape: () => descriptor('drawing.add'),
    onBringDrawingForward: () => descriptor('drawing.zOrder'),
    onSendDrawingBackward: () => descriptor('drawing.zOrder'),
    onRemoveDrawing: () => descriptor('drawing.remove'),
    onCreateSheetTable: () => undefined,
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
    assert.deepEqual(buildRibbonCommand('quickPivot', current), {
      type: 'command',
      descriptor: { commandId: 'pivot.add', params: { id: 'pivot-1' } },
    });
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
});
