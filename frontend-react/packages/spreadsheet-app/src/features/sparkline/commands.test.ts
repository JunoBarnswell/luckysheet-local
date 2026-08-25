import assert from 'node:assert/strict';
import { describe, it } from 'node:test';
import { CommandRuntime } from '@react-sheets/command-runtime';
import { WorkbookModel } from '@react-sheets/core-model';
import { registerSparklineFeature } from './index';

function register(): { workbook: WorkbookModel; runtime: CommandRuntime } {
  const workbook = new WorkbookModel('sparkline-feature-test', 'Sparkline Feature');
  const source = workbook.addSheet('source-2', 'Source 2');
  source.cells.set(0, 0, { value: 1 });
  source.cells.set(0, 1, { value: 3 });
  source.cells.set(0, 2, { value: 2 });
  const runtime = new CommandRuntime(workbook);
  registerSparklineFeature(runtime);
  return { workbook, runtime };
}

describe('sparkline feature contract', () => {
  it('inserts with a cross-sheet source and rejects unknown groups', () => {
    const { workbook, runtime } = register();
    runtime.execute('sparkline.insertDataLocation', {
      sheetId: 'sheet-1',
      sparklineId: 'spark-1',
      dataRange: { sheetId: 'source-2', startRow: 0, endRow: 0, startColumn: 0, endColumn: 2 },
      location: { row: 1, column: 3 },
      type: 'line',
    });
    assert.equal(workbook.getSheet('sheet-1').sparklines[0]?.sourceRange.sheetId, 'source-2');
    assert.throws(() => runtime.execute('sparkline.insertDataLocation', {
      sheetId: 'sheet-1',
      sparklineId: 'spark-2',
      dataRange: { sheetId: 'source-2', startRow: 0, endRow: 0, startColumn: 0, endColumn: 2 },
      location: { row: 2, column: 3 },
      type: 'line',
      groupId: 'missing-group',
    }), /Unknown sparkline group/);
  });

  it('group add/remove/replace is one reversible transaction with member state', () => {
    const { workbook, runtime } = register();
    runtime.execute('sparkline.insertDataLocation', {
      sheetId: 'sheet-1',
      sparklineId: 'spark-1',
      dataRange: { sheetId: 'source-2', startRow: 0, endRow: 0, startColumn: 0, endColumn: 2 },
      location: { row: 1, column: 3 },
      type: 'line',
    });
    runtime.execute('sparkline.group.create', {
      sheetId: 'sheet-1',
      group: { id: 'group-1', sheetId: 'sheet-1', type: 'column', sparklineIds: ['spark-1'], showAxis: true, showMarkers: true },
    });
    const sheet = workbook.getSheet('sheet-1');
    assert.equal(sheet.sparklines[0]?.groupId, 'group-1');
    assert.equal(sheet.sparklines[0]?.showAxis, true);
    assert.equal(runtime.getHistoryDepth().undo, 2);

    runtime.execute('sparkline.group.replace', {
      sheetId: 'sheet-1',
      group: { id: 'group-1', sheetId: 'sheet-1', type: 'win-loss', sparklineIds: ['spark-1'], showAxis: false, showMarkers: false },
    });
    assert.equal(sheet.sparklines[0]?.type, 'win-loss');
    assert.equal(sheet.sparklines[0]?.showAxis, false);
    assert.equal(runtime.undo(), true);
    assert.equal(sheet.sparklines[0]?.type, 'column');
    assert.equal(sheet.sparklines[0]?.showAxis, true);

    runtime.execute('sparkline.group.remove', { sheetId: 'sheet-1', groupId: 'group-1' });
    assert.equal(sheet.sparklineGroups.length, 0);
    assert.equal(sheet.sparklines[0]?.groupId, undefined);
    assert.equal(runtime.undo(), true);
    assert.equal(sheet.sparklineGroups[0]?.id, 'group-1');
    assert.equal(sheet.sparklines[0]?.groupId, 'group-1');
    assert.equal(runtime.redo(), true);
    assert.equal(sheet.sparklineGroups.length, 0);
  });

  it('fails closed for unknown update and remove targets', () => {
    const { runtime } = register();
    assert.throws(() => runtime.execute('sparkline.update', { sheetId: 'sheet-1', sparklineId: 'missing', patch: { color: '#fff' } }), /Unknown sparkline/);
    assert.throws(() => runtime.execute('sparkline.remove', { sheetId: 'sheet-1', sparklineId: 'missing' }), /Unknown sparkline/);
    assert.throws(() => runtime.execute('sparkline.group.remove', { sheetId: 'sheet-1', groupId: 'missing' }), /Unknown sparkline group/);
  });

  it('does not create history or mutation noise for a semantic no-op update', () => {
    const { workbook, runtime } = register();
    runtime.execute('sparkline.insertDataLocation', {
      sheetId: 'sheet-1',
      sparklineId: 'spark-1',
      dataRange: { sheetId: 'source-2', startRow: 0, endRow: 0, startColumn: 0, endColumn: 2 },
      location: { row: 1, column: 3 },
      type: 'line',
      highlightMax: true,
      highlightMin: true,
    });
    const before = workbook.snapshot();
    const beforeDepth = runtime.getHistoryDepth();
    const result = runtime.execute('sparkline.update', {
      sheetId: 'sheet-1',
      sparklineId: 'spark-1',
      patch: { highlightMax: true, highlightMin: true },
    });
    assert.equal(result.mutationCount, 0);
    assert.deepEqual(workbook.snapshot(), before);
    assert.deepEqual(runtime.getHistoryDepth(), beforeDepth);
  });

  it('removing a grouped sparkline updates group membership and restores both on undo', () => {
    const { workbook, runtime } = register();
    runtime.execute('sparkline.insertDataLocation', {
      sheetId: 'sheet-1',
      sparklineId: 'spark-1',
      dataRange: { sheetId: 'source-2', startRow: 0, endRow: 0, startColumn: 0, endColumn: 2 },
      location: { row: 1, column: 3 },
      type: 'line',
    });
    runtime.execute('sparkline.group.create', {
      sheetId: 'sheet-1',
      group: { id: 'group-1', sheetId: 'sheet-1', type: 'line', sparklineIds: ['spark-1'], showAxis: true, showMarkers: false },
    });
    runtime.execute('sparkline.remove', { sheetId: 'sheet-1', sparklineId: 'spark-1' });
    assert.equal(workbook.getSheet('sheet-1').sparklineGroups[0]?.sparklineIds.length, 0);
    runtime.execute('sparkline.group.remove', { sheetId: 'sheet-1', groupId: 'group-1' });
    assert.equal(workbook.getSheet('sheet-1').sparklineGroups.length, 0);
    assert.equal(runtime.undo(), true);
    assert.equal(workbook.getSheet('sheet-1').sparklineGroups[0]?.id, 'group-1');
    assert.equal(runtime.undo(), true);
    assert.equal(workbook.getSheet('sheet-1').sparklines[0]?.groupId, 'group-1');
    assert.deepEqual(workbook.getSheet('sheet-1').sparklineGroups[0]?.sparklineIds, ['spark-1']);
  });
});
