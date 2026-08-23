import assert from 'node:assert/strict';
import { describe, it } from 'node:test';
import { WorkbookModel } from '@react-sheets/core-model';
import { CommandRuntime } from '@react-sheets/command-runtime';
import { registerSheetCommands } from '@react-sheets/sheet-features';
import { CapabilityRegistry } from './index';
import { registerExtendedCommands } from './commands';

describe('M18 deterministic what-if commands', () => {
  it('keeps GROUPBY/PIVOTBY disabled by default', () => {
    const registry = new CapabilityRegistry();
    assert.equal(registry.isEnabled('groupby-pivotby'), false);
    assert.match(registry.get('groupby-pivotby')?.reason ?? '', /disabled|implemented/i);
  });

  it('applies scenario writes as one transaction and supports undo', () => {
    const workbook = new WorkbookModel('what-if-scenario', 'What-if');
    const runtime = new CommandRuntime(workbook);
    registerSheetCommands(runtime);
    registerExtendedCommands(runtime.registry);
    const sheetId = workbook.activeSheetId;
    const command = runtime.execute('extended.whatIf.scenario', {
      sheetId,
      scenario: {
        id: 'growth',
        name: 'Growth',
        changingCells: [{ row: 0, column: 1, value: 25 }],
      },
    });
    assert.equal((command as { plan?: { metadata?: { schema?: string; planHash?: string; sourceRevision?: string; deterministic?: boolean } } }).plan?.metadata?.schema, 'WhatIfPlanV1');
    assert.equal((command as { plan?: { metadata?: { deterministic?: boolean } } }).plan?.metadata?.deterministic, true);
    assert.ok((command as { plan?: { metadata?: { planHash?: string } } }).plan?.metadata?.planHash);
    assert.equal(runtime.getHistoryDepth().undo, 1);
    assert.equal(workbook.getSheet(sheetId).cells.get(0, 1)?.value, 25);
    assert.equal(runtime.undo(), true);
    assert.equal(workbook.getSheet(sheetId).cells.get(0, 1), undefined);
  });

  it('does not partially write an invalid data table', () => {
    const workbook = new WorkbookModel('what-if-invalid', 'What-if');
    const runtime = new CommandRuntime(workbook);
    registerSheetCommands(runtime);
    registerExtendedCommands(runtime.registry);
    const sheetId = workbook.activeSheetId;
    workbook.getSheet(sheetId).cells.set(0, 1, { value: 10 });
    workbook.getSheet(sheetId).cells.set(0, 2, { value: null });
    const result = runtime.execute('extended.whatIf.dataTable', {
      sheetId,
      columnInputCell: { row: 0, column: 1 },
      tableRange: { startRow: 0, startColumn: 0, endRow: 1, endColumn: 2 },
    });
    assert.equal(result.mutationCount, 0);
    assert.equal(workbook.getSheet(sheetId).cells.get(1, 1), undefined);
    assert.equal(workbook.getSheet(sheetId).cells.get(1, 2), undefined);
  });
});
