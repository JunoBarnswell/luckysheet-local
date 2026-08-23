import assert from 'node:assert/strict';
import { describe, it } from 'node:test';
import { WorkbookModel } from '@react-sheets/core-model';
import { CommandRuntime } from '@react-sheets/command-runtime';
import { registerSheetCommands } from '@react-sheets/sheet-features';
import { registerExtendedCommands } from './commands';

describe('M18 deterministic what-if commands', () => {
  it('applies scenario writes as one transaction and supports undo', () => {
    const workbook = new WorkbookModel('what-if-scenario', 'What-if');
    const runtime = new CommandRuntime(workbook);
    registerSheetCommands(runtime);
    registerExtendedCommands(runtime.registry);
    const sheetId = workbook.primarySheetId;
    const command = runtime.execute('extended.whatIf.scenario', {
      sheetId,
      scenario: {
        id: 'growth',
        name: 'Growth',
        changingCells: [{ row: 0, column: 1, value: 25 }],
      },
    });
    assert.equal((command as { plan?: { metadata?: { schema?: string; planHash?: string; sourceRevision?: string; deterministic?: boolean } } }).plan?.metadata?.schema, 'WhatIfPlan');
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
    const sheetId = workbook.primarySheetId;
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

  it('does not mutate on a non-monotonic goal function', () => {
    const workbook = new WorkbookModel('what-if-non-monotonic', 'What-if');
    const runtime = new CommandRuntime(workbook);
    registerSheetCommands(runtime);
    registerExtendedCommands(runtime.registry);
    const sheetId = workbook.primarySheetId;
    runtime.execute('sheet.cell.set', { sheetId, row: 0, column: 0, value: { formula: '=B1*B1' } });
    runtime.execute('sheet.cell.set', { sheetId, row: 0, column: 1, value: { value: 10 } });
    const result = runtime.execute('extended.whatIf.goalSeek', {
      sheetId,
      setCell: { row: 0, column: 0 },
      toValue: 1,
      byChangingCell: { row: 0, column: 1 },
    });
    assert.equal(result.mutationCount, 0);
    assert.match((result as { plan?: { result?: { message?: string } } }).plan?.result?.message ?? '', /non-monotonic/i);
    assert.equal(workbook.getSheet(sheetId).cells.get(0, 1)?.value, 10);
  });

  it('does not mutate a spill range during what-if planning', () => {
    const workbook = new WorkbookModel('what-if-spill', 'What-if');
    const runtime = new CommandRuntime(workbook);
    registerSheetCommands(runtime);
    registerExtendedCommands(runtime.registry);
    const sheetId = workbook.primarySheetId;
    workbook.getSheet(sheetId).spillRanges.push({
      sheetId,
      anchor: { row: 0, column: 1 },
      range: { sheetId, startRow: 0, endRow: 0, startColumn: 1, endColumn: 2 },
      values: [],
      state: 'ok',
    });
    const result = runtime.execute('extended.whatIf.scenario', {
      sheetId,
      scenario: {
        id: 'spill',
        name: 'Spill',
        changingCells: [{ row: 0, column: 1, value: 5 }],
      },
    });
    assert.equal(result.mutationCount, 0);
    assert.equal(workbook.getSheet(sheetId).cells.get(0, 1), undefined);
    assert.match((result as { plan?: { result?: { message?: string } } }).plan?.result?.message ?? '', /spill/i);
  });
});
