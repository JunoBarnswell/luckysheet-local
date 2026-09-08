import assert from 'node:assert/strict';
import { describe, it } from 'node:test';
import { DrawingRuntime, registerDrawingFeature } from '../../index';
import { openCanonicalTestRuntime } from '../../../../core-model/src/canonical-test-runtime.test';

const imageAsset = (assetId: string) => ({
  schema: 'AssetRef' as const,
  assetId,
  contentHash: 'a'.repeat(64),
  mimeType: 'image/png',
  byteLength: 1,
});

async function setup(unitId: string) {
  const fixture = await openCanonicalTestRuntime(unitId, 'Drawing');
  registerDrawingFeature(fixture.runtime, new DrawingRuntime());
  return fixture;
}

describe('drawing feature', () => {
  it('plans and commits the canonical drawing aggregate, including history replay', async () => {
    const { workbook, runtime, close } = await setup('drawing-canonical-aggregate');
    try {
      const planned: string[] = [];
      runtime.onMutation((mutation) => planned.push(mutation.id));
      await runtime.execute('drawing.add.shape', {
        sheetId: 'sheet-1',
        drawing: {
          id: 'draw-1', sheetId: 'sheet-1', kind: 'shape', payloadId: 'shape-1',
          anchor: { kind: 'absolute' }, transform: { x: 10, y: 20, width: 100, height: 80, rotation: 0 }, zIndex: 1,
        },
        payload: { kind: 'shape', type: 'rectangle', fill: '#fff', stroke: '#000' },
      });
      await runtime.execute('drawing.move', {
        sheetId: 'sheet-1', drawingId: 'draw-1',
        transform: { x: 30, y: 40, width: 100, height: 80, rotation: 0 },
      });
      await runtime.execute('drawing.zorder', { sheetId: 'sheet-1', drawingId: 'draw-1', direction: 'front' });

      const sheet = workbook.getSheet('sheet-1');
      assert.deepEqual(planned, ['drawing.add', 'drawing.transform', 'drawing.zorder']);
      assert.equal(sheet.drawings[0]?.transform.x, 30);
      assert.equal(sheet.drawingPayloads.get('shape-1')?.kind, 'shape');
      assert.ok((sheet.drawings[0]?.zIndex ?? 0) > 1);

      assert.equal(await runtime.undo(), true);
      assert.equal(sheet.drawings[0]?.zIndex, 1);
      assert.equal(await runtime.redo(), true);
      assert.ok((sheet.drawings[0]?.zIndex ?? 0) > 1);
    } finally {
      close();
    }
  });

  it('commits one pointer transform with canonical geometry', async () => {
    const { workbook, runtime, close } = await setup('drawing-canonical-pointer');
    try {
      await runtime.execute('drawing.add', {
        sheetId: 'sheet-1',
        drawing: {
          id: 'draw-pointer', sheetId: 'sheet-1', kind: 'shape', payloadId: 'shape-pointer',
          anchor: { kind: 'absolute' }, transform: { x: 10, y: 10, width: 40, height: 30, rotation: 0 }, zIndex: 1,
        },
        payload: { kind: 'shape', type: 'rectangle', fill: '#fff', stroke: '#000' },
      });
      runtime.clearHistory();
      const drawingRuntime = new DrawingRuntime();
      const transaction = drawingRuntime.beginPointerTransform(workbook.getSheet('sheet-1'), 'draw-pointer');
      drawingRuntime.previewPointerTransform(workbook.getSheet('sheet-1'), transaction.id, { x: 21, y: 29, width: 44, height: 36, rotation: 13 });
      const commit = drawingRuntime.finishPointerTransform(transaction.id);
      await runtime.execute('drawing.transform.commit', { sheetId: 'sheet-1', ...commit });
      assert.deepEqual(workbook.getSheet('sheet-1').drawings[0]?.transform, { x: 24, y: 32, width: 48, height: 40, rotation: 13 });
      assert.equal(await runtime.undo(), true);
      assert.deepEqual(workbook.getSheet('sheet-1').drawings[0]?.transform, { x: 10, y: 10, width: 40, height: 30, rotation: 0 });
    } finally {
      close();
    }
  });

  it('commits typed image payload updates and fails closed before a kernel commit', async () => {
    const { workbook, runtime, close } = await setup('drawing-canonical-image');
    try {
      await runtime.execute('drawing.add.image', {
        sheetId: 'sheet-1',
        drawing: {
          id: 'draw-image', sheetId: 'sheet-1', kind: 'image', payloadId: 'image-payload',
          anchor: { kind: 'two-cell', row: 1, column: 1, endRow: 4, endColumn: 5 },
          transform: { x: 10, y: 10, width: 100, height: 80, rotation: 0 }, zIndex: 1,
        },
        payload: { kind: 'image', asset: imageAsset('asset-image'), altText: 'Before' },
      });
      await runtime.execute('drawing.image.altText', { sheetId: 'sheet-1', drawingId: 'draw-image', altText: 'Accessible image' });
      await runtime.execute('drawing.image.crop', { sheetId: 'sheet-1', drawingId: 'draw-image', crop: { left: 0.1, top: 0.2, right: 0.1, bottom: 0 } });
      const payload = workbook.getSheet('sheet-1').drawingPayloads.get('image-payload');
      assert.equal(payload?.kind, 'image');
      assert.equal(payload?.kind === 'image' ? payload.altText : undefined, 'Accessible image');
      assert.deepEqual(payload?.kind === 'image' ? payload.crop : undefined, { left: 0.1, top: 0.2, right: 0.1, bottom: 0 });

      const revision = workbook.revision;
      await assert.rejects(() => runtime.execute('drawing.image.effects', {
        sheetId: 'sheet-1', drawingId: 'draw-image', effects: { brightness: 2 },
      }), /Image effects are invalid/);
      assert.equal(workbook.revision, revision);
    } finally {
      close();
    }
  });
});
