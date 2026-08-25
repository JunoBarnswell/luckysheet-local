import assert from 'node:assert/strict';
import { describe, it } from 'node:test';
import { CommandRuntime } from '@react-sheets/command-runtime';
import { StructuralTransform, WorkbookModel, type DrawingObject, type DrawingPayload } from '@react-sheets/core-model';
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
    assert.equal(workbook.getSheet('sheet-1').drawings.filter((entry) => entry.kind === 'image').length, 1);

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

  it('keeps pointer previews transient and commits one transform operation', () => {
    const workbook = new WorkbookModel('drawing-pointer-test', 'Drawing Pointer');
    const runtime = new CommandRuntime(workbook);
    const drawingRuntime = new DrawingRuntime();
    registerDrawingFeature(runtime, drawingRuntime);
    runtime.execute('drawing.add', {
      sheetId: 'sheet-1',
      drawing: {
        id: 'draw-pointer',
        sheetId: 'sheet-1',
        kind: 'shape',
        payloadId: 'shape-pointer',
        anchor: { kind: 'absolute' },
        transform: { x: 10, y: 10, width: 40, height: 30, rotation: 0 },
        zIndex: 1,
      },
      payload: { kind: 'shape', type: 'rectangle', fill: '#fff', stroke: '#000' },
    });
    runtime.clearHistory();
    const transaction = drawingRuntime.beginPointerTransform(workbook.getSheet('sheet-1'), 'draw-pointer');
    drawingRuntime.previewPointerTransform(transaction.id, { x: 21, y: 29, width: 44, height: 36, rotation: 13 }, 8);
    assert.deepEqual(workbook.getSheet('sheet-1').drawings[0]?.transform, { x: 10, y: 10, width: 40, height: 30, rotation: 0 });
    const commit = drawingRuntime.finishPointerTransform(transaction.id);
    runtime.execute('drawing.transform.commit', { sheetId: 'sheet-1', ...commit });
    assert.deepEqual(workbook.getSheet('sheet-1').drawings[0]?.transform, { x: 24, y: 32, width: 48, height: 40, rotation: 13 });
    assert.equal(runtime.getHistoryDepth().undo, 1);
    assert.equal(runtime.undo(), true);
    assert.deepEqual(workbook.getSheet('sheet-1').drawings[0]?.transform, { x: 10, y: 10, width: 40, height: 30, rotation: 0 });
    assert.equal(runtime.redo(), true);
  });

  it('reconciles selections and cancels pointer gestures for removed drawings', () => {
    const workbook = new WorkbookModel('drawing-reconcile-test', 'Drawing Reconcile');
    const runtime = new CommandRuntime(workbook);
    const drawingRuntime = new DrawingRuntime();
    registerDrawingFeature(runtime, drawingRuntime);
    runtime.execute('drawing.add', {
      sheetId: 'sheet-1',
      drawing: {
        id: 'draw-reconcile',
        sheetId: 'sheet-1',
        kind: 'shape',
        payloadId: 'shape-reconcile',
        anchor: { kind: 'absolute' },
        transform: { x: 10, y: 10, width: 40, height: 30, rotation: 0 },
        zIndex: 1,
      },
      payload: { kind: 'shape', type: 'rectangle', fill: '#fff', stroke: '#000' },
    });
    drawingRuntime.select('sheet-1', ['draw-reconcile']);
    const transaction = drawingRuntime.beginPointerTransform(workbook.getSheet('sheet-1'), 'draw-reconcile');

    const result = drawingRuntime.reconcile('sheet-1', []);

    assert.deepEqual(result.selection, []);
    assert.deepEqual(result.cancelledPointerTransactionIds, [transaction.id]);
    assert.deepEqual(drawingRuntime.getSelection('sheet-1'), []);
    assert.throws(() => drawingRuntime.finishPointerTransform(transaction.id), /Unknown drawing pointer transaction/);
  });

  it('applies the same existence reconciliation to every canonical drawing kind', () => {
    const workbook = new WorkbookModel('drawing-kinds-reconcile-test', 'Drawing Kinds Reconcile');
    const runtime = new CommandRuntime(workbook);
    const drawingRuntime = new DrawingRuntime();
    registerDrawingFeature(runtime, drawingRuntime);
    const sheetId = 'sheet-1';
    const style = { theme: 'light' as const, fill: '#fff', border: '#000', textColor: '#000', accentColor: '#2563eb' };
    const payloads: Array<{ kind: DrawingObject['kind']; payload: DrawingPayload }> = [
      { kind: 'image', payload: { kind: 'image', src: 'data:image/png;base64,AA==', altText: 'Image' } },
      { kind: 'shape', payload: { kind: 'shape', type: 'rectangle', fill: '#fff', stroke: '#000' } },
      { kind: 'textbox', payload: { kind: 'textbox', text: 'Text' } },
      { kind: 'chart', payload: { kind: 'chart', chartId: 'payload-kind-3', chartType: 'column', sourceRanges: [], elements: { hiddenData: 'show' } } },
      { kind: 'data-chart', payload: { kind: 'data-chart', tableId: 'table-kind', plots: [{ type: 'column', valueFieldId: 'value', aggregate: 'sum' }], config: {} } },
      { kind: 'camera', payload: { kind: 'camera', sourceRange: { sheetId, startRow: 0, endRow: 1, startColumn: 0, endColumn: 1 }, refreshPolicy: 'live' } },
      { kind: 'form-control', payload: { kind: 'form-control', controlType: 'button', value: false, enabled: true, style: { fill: '#fff', border: '#000', textColor: '#000' } } },
      { kind: 'slicer', payload: { kind: 'slicer', pivotId: 'pivot-kind', fieldId: 'field-kind', filter: { mode: 'all', memberKeys: [] }, style } },
      { kind: 'timeline', payload: { kind: 'timeline', pivotId: 'pivot-kind', fieldId: 'field-kind', period: {}, style } },
    ];
    const drawings = payloads.map(({ kind, payload }, index) => ({
      id: `drawing-kind-${index}`,
      sheetId,
      kind,
      payloadId: `payload-kind-${index}`,
      anchor: { kind: 'absolute' as const },
      transform: { x: index * 20, y: 0, width: 20, height: 20, rotation: 0 },
      zIndex: index,
      payload,
    }));
    for (const entry of drawings) {
      runtime.execute('drawing.add', { sheetId, drawing: entry, payload: entry.payload });
    }
    drawingRuntime.select(sheetId, drawings.map((drawing) => drawing.id));

    for (const drawing of drawings) {
      runtime.execute('drawing.remove', { sheetId, drawingId: drawing.id });
      const result = drawingRuntime.reconcile(sheetId, workbook.getSheet(sheetId).drawings.map((entry) => entry.id));
      assert.equal(result.selection.includes(drawing.id), false);
    }
    assert.deepEqual(drawingRuntime.getSelection(sheetId), []);
  });

  it('supports crop and alt text as payload mutations with inverse and remote replay', () => {
    const workbook = new WorkbookModel('drawing-image-test', 'Drawing Image');
    const runtime = new CommandRuntime(workbook);
    registerDrawingFeature(runtime);
    runtime.execute('drawing.add.image', {
      sheetId: 'sheet-1',
      drawing: {
        id: 'draw-image',
        sheetId: 'sheet-1',
        kind: 'image',
        payloadId: 'image-payload',
        anchor: { kind: 'two-cell', row: 1, column: 1, endRow: 4, endColumn: 5 },
        transform: { x: 10, y: 10, width: 100, height: 80, rotation: 0 },
        zIndex: 1,
      },
      payload: { kind: 'image', src: 'https://example.com/image.png', altText: 'Before' },
    });
    runtime.execute('drawing.image.altText', { sheetId: 'sheet-1', drawingId: 'draw-image', altText: 'Accessible image' });
    runtime.execute('drawing.image.crop', { sheetId: 'sheet-1', drawingId: 'draw-image', crop: { left: 0.1, top: 0.2, right: 0.1, bottom: 0 } });
    const payload = workbook.getSheet('sheet-1').drawingPayloads.get('image-payload') as { kind: 'image'; altText?: string; crop?: unknown };
    assert.equal(payload.altText, 'Accessible image');
    assert.deepEqual(payload.crop, { left: 0.1, top: 0.2, right: 0.1, bottom: 0 });
    assert.equal(runtime.undo(), true);
    assert.equal((workbook.getSheet('sheet-1').drawingPayloads.get('image-payload') as { crop?: unknown }).crop, undefined);
    assert.equal(runtime.redo(), true);

    const remoteWorkbook = new WorkbookModel('drawing-image-test', 'Drawing Image');
    const remoteRuntime = new CommandRuntime(remoteWorkbook);
    registerDrawingFeature(remoteRuntime);
    remoteRuntime.applyRemoteMutations(runtime.getUndoEntries().flatMap((entry) => entry.redo));
    const remotePayload = remoteWorkbook.getSheet('sheet-1').drawingPayloads.get('image-payload') as { kind: 'image'; altText?: string; crop?: unknown };
    assert.equal(remotePayload.kind, 'image');
    assert.equal(remotePayload.altText, 'Accessible image');
    assert.deepEqual(remotePayload.crop, { left: 0.1, top: 0.2, right: 0.1, bottom: 0 });
  });

  it('preserves anchored drawings when rows are inserted', () => {
    const workbook = new WorkbookModel('drawing-anchor-test', 'Drawing Anchor');
    const runtime = new CommandRuntime(workbook);
    registerDrawingFeature(runtime);
    runtime.execute('drawing.add', {
      sheetId: 'sheet-1',
      drawing: {
        id: 'draw-anchor',
        sheetId: 'sheet-1',
        kind: 'shape',
        payloadId: 'shape-anchor',
        anchor: { kind: 'two-cell', row: 2, column: 2, endRow: 5, endColumn: 4 },
        transform: { x: 0, y: 0, width: 20, height: 20 },
        zIndex: 1,
      },
      payload: { kind: 'shape', type: 'ellipse', fill: '#fff', stroke: '#000' },
    });
    StructuralTransform.apply(workbook, { kind: 'insert-rows', sheetId: 'sheet-1', at: 1, count: 2 });
    assert.deepEqual(workbook.getSheet('sheet-1').drawings[0]?.anchor, { kind: 'two-cell', row: 4, column: 2, endRow: 7, endColumn: 4 });
  });

  it('aligns, distributes and copies selected drawings as canonical transactions', () => {
    const workbook = new WorkbookModel('drawing-layout-test', 'Drawing Layout');
    const runtime = new CommandRuntime(workbook);
    const drawingRuntime = new DrawingRuntime();
    registerDrawingFeature(runtime, drawingRuntime);
    const add = (id: string, x: number, y: number) => runtime.execute('drawing.add', {
      sheetId: 'sheet-1',
      drawing: {
        id: `drawing-${id}`,
        sheetId: 'sheet-1',
        kind: 'shape' as const,
        payloadId: `payload-${id}`,
        anchor: { kind: 'absolute' as const },
        transform: { x, y, width: 20, height: 20 },
        zIndex: 1,
      },
      payload: { kind: 'shape' as const, type: 'rectangle' as const, fill: '#fff', stroke: '#000' },
    });
    add('a', 10, 10);
    add('b', 80, 50);
    add('c', 150, 90);
    runtime.clearHistory();
    runtime.execute('drawing.select', { sheetId: 'sheet-1', drawingIds: ['drawing-a', 'drawing-b'], mode: 'replace' });
    assert.deepEqual(drawingRuntime.getSelection('sheet-1'), ['drawing-a', 'drawing-b']);
    runtime.execute('drawing.align', { sheetId: 'sheet-1', drawingIds: ['drawing-a', 'drawing-b'], alignment: 'top' });
    assert.equal(workbook.getSheet('sheet-1').drawings.find((entry) => entry.id === 'drawing-b')?.transform.y, 10);
    assert.equal(runtime.getHistoryDepth().undo, 1);
    runtime.execute('drawing.copy', { sheetId: 'sheet-1', sourceDrawingId: 'drawing-a', drawingId: 'drawing-copy', payloadId: 'payload-copy', offset: { x: 12, y: 14 } });
    assert.equal(workbook.getSheet('sheet-1').drawings.length, 4);
    assert.equal(workbook.getSheet('sheet-1').drawings.find((entry) => entry.id === 'drawing-copy')?.transform.x, 22);
    assert.equal(runtime.undo(), true);
    assert.equal(workbook.getSheet('sheet-1').drawings.length, 3);
  });
});
