import assert from 'node:assert/strict';
import test from 'node:test';
import { kernelInvoke } from '@react-sheets/kernel-client';
import { WorkbookModel, type KernelReplicaManifest } from './index';
import { assertCanonicalWorkbookSnapshot, migrateStoredWorkbookSnapshot } from './snapshot';
import { openCanonicalTestRuntime } from './canonical-test-runtime.test';
import { registerSheetCommands } from '../../sheet-features/src/index';

test('WorkbookModel is opened from the canonical manifest and exposes a read-only page replica', async () => {
  const { workbook, close } = await openCanonicalTestRuntime('core-model-manifest');
  try {
    const manifest = workbook.manifest();
    assert.equal(manifest.schema, 'WorkbookManifest');
    assert.equal(manifest.version, 11);
    assert.equal(manifest.revision, 0);
    assert.equal(workbook.primarySheetId, 'sheet-1');
    assert.equal(workbook.getSheet('sheet-1').cells.get(0, 0), undefined);
    assert.equal('set' in workbook.getSheet('sheet-1').cells, false);
  } finally {
    close();
  }
});

test('typed cell command commits a new manifest revision and page value', async () => {
  const { workbook, runtime, close } = await openCanonicalTestRuntime('core-model-command');
  try {
    registerSheetCommands(runtime);
    const result = await runtime.execute('sheet.cell.set', {
      sheetId: 'sheet-1', row: 2, column: 3, value: { value: 'canonical' },
    });
    assert.equal(result.mutationCount, 1);
    assert.equal(workbook.revision, 1);
    assert.equal(workbook.manifest().revision, 1);
    assert.equal(workbook.getSheet('sheet-1').cells.get(2, 3)?.value, 'canonical');
    assert.equal(runtime.getHistoryDepth().undo, 1);
  } finally {
    close();
  }
});

test('command parameter rejection leaves the committed manifest and page replica unchanged', async () => {
  const { workbook, runtime, close } = await openCanonicalTestRuntime('core-model-rejection');
  try {
    registerSheetCommands(runtime);
    await assert.rejects(
      runtime.execute('sheet.cell.set', { sheetId: 'sheet-1', row: -1, column: 0, value: { value: 1 } }),
      /Invalid cell set parameters|invalid/i,
    );
    assert.equal(workbook.revision, 0);
    assert.equal(runtime.getHistoryDepth().undo, 0);
    assert.equal(workbook.getSheet('sheet-1').cells.get(0, 0), undefined);
  } finally {
    close();
  }
});

test('manifest identity mismatch fails before replacing the active replica', async () => {
  const { workbook, close } = await openCanonicalTestRuntime('core-model-identity');
  try {
    const manifest = structuredClone(workbook.manifest()) as KernelReplicaManifest;
    (manifest as { unitId: string }).unitId = 'different-workbook';
    assert.throws(() => workbook.applyCommittedManifest(manifest), /KERNEL_MANIFEST_IDENTITY_MISMATCH/);
    assert.equal(workbook.manifest().unitId, 'core-model-identity');
  } finally {
    close();
  }
});

test('manifest metadata remains the sole workbook state boundary for restored definitions', async () => {
  const { workbook, close } = await openCanonicalTestRuntime('core-model-metadata');
  try {
    const committed = workbook.manifest();
    assert.equal(committed.revision, 0);
    assert.equal(committed.sheets.length, 1);
    assert.equal(committed.sheets[0]?.sheetId, 'sheet-1');
    assert.deepEqual(committed.metadata, {});
  } finally {
    close();
  }
});

test('stored snapshot migration accepts only an explicit canonical snapshot revision', () => {
  const canonical = {
    schema: 'WorkbookSnapshot', version: 10, unitId: 'snapshot-test', name: 'Snapshot',
    dimensionMetrics: { normalFontFamily: 'Calibri', normalFontSizePx: 14, maximumDigitWidthPx: 7 },
    calculationSettings: { mode: 'automatic', iterativeCalculation: false, maximumIterations: 100, maximumChange: 0.001, precisionAsDisplayed: false, calculateBeforeSave: true, fullCalculationOnLoad: false },
    editingOptions: { allowEditDirectly: true, moveAfterEnter: true, enterDirection: 'down', formulaAutoComplete: true, valueAutoComplete: true, fixedDecimalPlaces: null }, definedNameModels: [],
    dataModel: { sources: [], tables: [], relationships: [], views: [] }, sheets: [],
  } as any;
  assert.equal(assertCanonicalWorkbookSnapshot(canonical).version, 10);
  assert.equal(migrateStoredWorkbookSnapshot(canonical).schema, 'WorkbookSnapshot');
  assert.throws(() => migrateStoredWorkbookSnapshot({ ...canonical, version: 99 }), /Unsupported workbook snapshot version/);
});

test('the kernel manifest operation returns the same canonical identity used by the page replica', async () => {
  const { workbook, close } = await openCanonicalTestRuntime('core-model-kernel-manifest');
  try {
    const manifest = kernelInvoke<KernelReplicaManifest>('manifest', { unitId: workbook.unitId, revision: workbook.revision });
    assert.deepEqual(manifest, workbook.manifest());
  } finally {
    close();
  }
});
