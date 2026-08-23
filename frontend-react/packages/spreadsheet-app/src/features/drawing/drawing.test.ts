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
    assert.equal(workbook.getSheet('sheet-1').drawingPayloads.get('img-1')?.kind, 'image');
    assert.equal(workbook.getSheet('sheet-1').images.length, 0);

    runtime.execute('drawing.zorder', { sheetId: 'sheet-1', drawingId: 'draw-1', direction: 'front' });
    assert.ok((workbook.getSheet('sheet-1').drawings[0]?.zIndex ?? 0) > 1);

    runtime.undo();
    assert.equal(workbook.getSheet('sheet-1').drawings[0]?.zIndex, 1);
  });

  it('removes and restores the complete drawing aggregate through undo/redo and remote replay', () => {
    const workbook = new WorkbookModel('drawing-remove-test', 'Drawing Remove');
    const runtime = new CommandRuntime(workbook);
    registerDrawingFeature(runtime);
    const drawing = {
      id: 'draw-1',
      sheetId: 'sheet-1',
      kind: 'shape' as const,
      payloadId: 'shape-1',
      anchor: { kind: 'absolute' as const },
      transform: { x: 10, y: 20, width: 100, height: 80, rotation: 12 },
      zIndex: 1,
    };
    const payload = { kind: 'shape' as const, type: 'rectangle' as const, fill: '#fff', stroke: '#000', text: 'A' };

    runtime.execute('drawing.add.shape', { sheetId: 'sheet-1', drawing, payload });
    runtime.execute('drawing.remove', { sheetId: 'sheet-1', drawingId: drawing.id });
    assert.equal(workbook.getSheet('sheet-1').drawings.length, 0);
    assert.equal(workbook.getSheet('sheet-1').drawingPayloads.size, 0);

    assert.equal(runtime.undo(), true);
    assert.deepEqual(workbook.getSheet('sheet-1').drawings[0], drawing);
    assert.deepEqual(workbook.getSheet('sheet-1').drawingPayloads.get('shape-1'), payload);
    assert.equal(runtime.redo(), true);
    assert.equal(workbook.getSheet('sheet-1').drawings.length, 0);

    const remoteWorkbook = new WorkbookModel('drawing-remove-test', 'Drawing Remove');
    const remoteRuntime = new CommandRuntime(remoteWorkbook);
    registerDrawingFeature(remoteRuntime);
    remoteRuntime.applyRemoteMutations([
      {
        id: 'drawing.add',
        unitId: remoteWorkbook.unitId,
        sheetId: 'sheet-1',
        params: { sheetId: 'sheet-1', drawing, payload },
        affectedRanges: [],
      },
      {
        id: 'drawing.remove',
        unitId: remoteWorkbook.unitId,
        sheetId: 'sheet-1',
        params: { sheetId: 'sheet-1', drawingId: drawing.id },
        affectedRanges: [],
      },
    ]);
    assert.equal(remoteWorkbook.getSheet('sheet-1').drawings.length, 0);
    assert.equal(remoteWorkbook.getSheet('sheet-1').drawingPayloads.size, 0);
  });

  it('restores all z-indices when undoing a forward/backward reorder', () => {
    const workbook = new WorkbookModel('drawing-zorder-test', 'Drawing Z');
    const runtime = new CommandRuntime(workbook);
    registerDrawingFeature(runtime);
    const create = (id: string, zIndex: number) => ({
      sheetId: 'sheet-1',
      drawing: {
        id,
        sheetId: 'sheet-1',
        kind: 'shape' as const,
        payloadId: `${id}-payload`,
        anchor: { kind: 'absolute' as const },
        transform: { x: 0, y: 0, width: 10, height: 10 },
        zIndex,
      },
      payload: { kind: 'shape' as const, type: 'rectangle' as const, fill: '#fff', stroke: '#000' },
    });
    runtime.execute('drawing.add.shape', create('a', 1));
    runtime.execute('drawing.add.shape', create('b', 2));
    runtime.execute('drawing.zorder', { sheetId: 'sheet-1', drawingId: 'a', direction: 'forward' });
    assert.deepEqual(workbook.getSheet('sheet-1').drawings.map((entry) => entry.zIndex), [2, 1]);
    assert.equal(runtime.undo(), true);
    assert.deepEqual(workbook.getSheet('sheet-1').drawings.map((entry) => entry.zIndex), [1, 2]);
  });
});
