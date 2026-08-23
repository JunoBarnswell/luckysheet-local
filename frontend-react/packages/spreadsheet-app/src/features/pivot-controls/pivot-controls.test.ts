import assert from 'node:assert/strict';
import { describe, it } from 'node:test';
import { CommandRuntime } from '@react-sheets/command-runtime';
import { createPivotMemberKey, WorkbookModel } from '@react-sheets/core-model';
import { registerDrawingFeature } from '../drawing';
import {
  buildPivotSlicerDrawing,
  buildPivotTimelineDrawing,
  listPivotControlRecords,
  listPivotControlsForPivot,
  registerPivotControlFeature,
} from './index';

function createRuntime() {
  const workbook = new WorkbookModel('pivot-controls-test', 'Pivot Controls');
  const runtime = new CommandRuntime(workbook);
  registerDrawingFeature(runtime);
  registerPivotControlFeature(runtime);
  return { workbook, runtime };
}

describe('pivot controls as floating drawing objects', () => {
  it('creates canonical slicer and timeline payloads without using PivotModel arrays', () => {
    const { workbook, runtime } = createRuntime();
    const slicer = buildPivotSlicerDrawing({
      drawingId: 'slicer-drawing',
      payloadId: 'slicer-payload',
      sheetId: 'sheet-1',
      pivotId: 'pivot-sales',
      fieldId: 'field-region',
      transform: { x: 32, y: 16, width: 220, height: 180 },
      zIndex: 1,
    });
    const timeline = buildPivotTimelineDrawing({
      drawingId: 'timeline-drawing',
      payloadId: 'timeline-payload',
      sheetId: 'sheet-1',
      pivotId: 'pivot-sales',
      fieldId: 'field-date',
      period: { start: '2025-01-01', end: '2025-12-31' },
      transform: { x: 280, y: 16, width: 360, height: 72 },
      zIndex: 2,
    });

    runtime.execute('pivot.control.slicer.create', slicer);
    runtime.execute('pivot.control.timeline.create', timeline);

    const sheet = workbook.getSheet('sheet-1');
    assert.deepEqual(sheet.pivots, []);
    assert.deepEqual(listPivotControlRecords(sheet).map((record) => record.payload.kind), ['slicer', 'timeline']);
    assert.equal(listPivotControlsForPivot(sheet, 'pivot-sales').length, 2);
    assert.deepEqual(sheet.drawingPayloads.get('slicer-payload'), slicer.payload);
    assert.deepEqual(sheet.drawingPayloads.get('timeline-payload'), timeline.payload);
  });

  it('updates filter, period, style and links as one undoable payload transaction each', () => {
    const { workbook, runtime } = createRuntime();
    const slicer = buildPivotSlicerDrawing({
      drawingId: 'slicer-drawing',
      payloadId: 'slicer-payload',
      sheetId: 'sheet-1',
      pivotId: 'pivot-sales',
      fieldId: 'field-region',
      transform: { x: 0, y: 0, width: 200, height: 120 },
      zIndex: 1,
    });
    runtime.execute('pivot.control.slicer.create', slicer);
    runtime.execute('pivot.control.slicer.filter.set', {
      sheetId: 'sheet-1',
      drawingId: 'slicer-drawing',
      filter: { mode: 'include', memberKeys: [createPivotMemberKey('East')] },
    });
    const filtered = workbook.getSheet('sheet-1').drawingPayloads.get('slicer-payload');
    assert.equal(filtered?.kind, 'slicer');
    if (filtered?.kind !== 'slicer') throw new Error('Expected slicer payload');
    assert.equal(filtered.filter.mode, 'include');
    assert.equal(filtered.filter.memberKeys[0]?.value, 'East');
    assert.equal(runtime.undo(), true);
    assert.equal((workbook.getSheet('sheet-1').drawingPayloads.get('slicer-payload') as typeof filtered)?.filter.mode, 'all');
    assert.equal(runtime.redo(), true);

    runtime.execute('pivot.control.style.set', {
      sheetId: 'sheet-1',
      drawingId: 'slicer-drawing',
      style: {
        theme: 'dark',
        fill: '#0f172a',
        border: '#334155',
        textColor: '#f8fafc',
        accentColor: '#38bdf8',
      },
    });
    runtime.execute('pivot.control.connections.set', {
      sheetId: 'sheet-1',
      drawingId: 'slicer-drawing',
      connectedPivotIds: ['pivot-detail', 'pivot-detail', ''],
    });
    const updated = workbook.getSheet('sheet-1').drawingPayloads.get('slicer-payload');
    assert.equal(updated?.kind, 'slicer');
    if (updated?.kind !== 'slicer') throw new Error('Expected slicer payload');
    assert.equal(updated.style.theme, 'dark');
    assert.deepEqual(updated.connectedPivotIds, ['pivot-detail']);
  });

  it('updates timeline period while drawing commands own move, resize, delete and undo', () => {
    const { workbook, runtime } = createRuntime();
    const timeline = buildPivotTimelineDrawing({
      drawingId: 'timeline-drawing',
      payloadId: 'timeline-payload',
      sheetId: 'sheet-1',
      pivotId: 'pivot-sales',
      fieldId: 'field-date',
      transform: { x: 10, y: 20, width: 300, height: 80 },
      zIndex: 1,
    });
    runtime.execute('pivot.control.timeline.create', timeline);
    runtime.execute('pivot.control.timeline.period.set', {
      sheetId: 'sheet-1',
      drawingId: 'timeline-drawing',
      period: { start: '2026-01-01', end: '2026-03-31' },
    });
    runtime.execute('drawing.move', {
      sheetId: 'sheet-1',
      drawingId: 'timeline-drawing',
      transform: { x: 40, y: 50, width: 300, height: 80 },
    });
    runtime.execute('drawing.resize', {
      sheetId: 'sheet-1',
      drawingId: 'timeline-drawing',
      transform: { x: 40, y: 50, width: 420, height: 96 },
    });
    assert.deepEqual(workbook.getSheet('sheet-1').drawings[0]?.transform, { x: 40, y: 50, width: 420, height: 96 });

    runtime.execute('drawing.remove', { sheetId: 'sheet-1', drawingId: 'timeline-drawing' });
    assert.equal(workbook.getSheet('sheet-1').drawings.length, 0);
    assert.equal(runtime.undo(), true);
    assert.equal(workbook.getSheet('sheet-1').drawings[0]?.kind, 'timeline');
    assert.deepEqual(workbook.getSheet('sheet-1').drawingPayloads.get('timeline-payload'), {
      ...timeline.payload,
      period: { start: '2026-01-01', end: '2026-03-31' },
    });
    assert.equal(runtime.undo(), true);
    assert.deepEqual(workbook.getSheet('sheet-1').drawings[0]?.transform, { x: 40, y: 50, width: 300, height: 80 });
  });

  it('rejects a control payload without the canonical owned state', () => {
    const { runtime } = createRuntime();
    assert.throws(() => runtime.execute('drawing.add.slicer', {
      sheetId: 'sheet-1',
      drawing: {
        id: 'invalid-slicer',
        sheetId: 'sheet-1',
        kind: 'slicer',
        payloadId: 'invalid-payload',
        anchor: { kind: 'absolute' },
        transform: { x: 0, y: 0, width: 100, height: 100 },
        zIndex: 1,
      },
      payload: {
        kind: 'slicer',
        pivotId: 'pivot-sales',
        fieldId: 'field-region',
        filter: { mode: 'all', memberKeys: [] },
      },
    }));
  });
});
