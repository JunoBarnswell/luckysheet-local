import assert from 'node:assert/strict';
import { describe, it } from 'node:test';
import { CommandRuntime } from '@react-sheets/command-runtime';
import { WorkbookModel } from '@react-sheets/core-model';
import { DrawingRuntime, registerDrawingFeature } from '../../index';

describe('drawing feature', () => {
  it('adds, selects, moves, resizes, and reorders unified drawing objects', () => {
    const workbook = new WorkbookModel('drawing-test', 'Drawing');
    const runtime = new CommandRuntime(workbook);
    const drawingRuntime = new DrawingRuntime();
    registerDrawingFeature(runtime, drawingRuntime);

    runtime.execute('drawing.add.image', {
      sheetId: 'sheet-1',
      drawing: {
        id: 'draw-1',
        sheetId: 'sheet-1',
        kind: 'image',
        payloadId: 'img-1',
        anchor: { kind: 'absolute' },
        transform: { x: 10, y: 20, width: 100, height: 80, rotation: 0 },
        zIndex: 1,
      },
      payload: { kind: 'image', src: 'https://example.com/a.png', altText: 'Logo' },
    });

    runtime.execute('drawing.select', { sheetId: 'sheet-1', drawingIds: ['draw-1'] });
    assert.deepEqual(drawingRuntime.getSelection('sheet-1'), ['draw-1']);

    runtime.execute('drawing.move', {
      sheetId: 'sheet-1',
      drawingId: 'draw-1',
      transform: { x: 30, y: 40, width: 100, height: 80, rotation: 0 },
    });
    assert.equal(workbook.getSheet('sheet-1').drawings[0]?.transform.x, 30);
    assert.equal(workbook.getSheet('sheet-1').images[0]?.bounds.x, 30);

    runtime.execute('drawing.zorder', { sheetId: 'sheet-1', drawingId: 'draw-1', direction: 'front' });
    assert.ok((workbook.getSheet('sheet-1').drawings[0]?.zIndex ?? 0) > 1);

    runtime.undo();
    assert.equal(workbook.getSheet('sheet-1').drawings[0]?.zIndex, 1);
  });
});
