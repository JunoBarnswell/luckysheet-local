import assert from 'node:assert/strict';
import { test } from 'node:test';
import { zipSync, strToU8 } from 'fflate';
import { WorkbookModel } from '@react-sheets/core-model';
import { assertNativeInputSize, resolveNativeDocumentResourceLimits } from './native-resource-budget';
import { DEFAULT_NATIVE_DOCUMENT_RESOURCE_LIMITS } from './types';
import { importNativeDocumentWithWorker, type NativeDocumentWorkerPort } from './worker-port';
import { importOoxmlDocument } from './import';
import { exportOoxmlDocument } from './export';
import { loadOpcPackageGraph, parseLoadedOoxml } from './ooxml';
import { NativeDocumentError } from './native-document-error';

const isBudget = (error: unknown) => error instanceof NativeDocumentError && error.code === 'NATIVE_DOCUMENT_RESOURCE_LIMIT';
test('input admission precedes slice and worker creation/submission', async () => {
  class UncopyableBytes extends Uint8Array { override slice(): Uint8Array<ArrayBuffer> { throw new Error('copy happened before admission'); } }
  const bytes = new UncopyableBytes(9);
  const options = { compatibilityTarget: 'B' as const, limits: { maxArchiveBytes: 8 } };
  const port: NativeDocumentWorkerPort = { submit: () => { throw new Error('worker called before admission'); }, cancel() {}, dispose() {} };
  await assert.rejects(importNativeDocumentWithWorker({ fileName: 'hostile.xlsx', buffer: bytes, options }, port), isBudget);
  await assert.rejects(importOoxmlDocument({ fileName: 'hostile.xlsx', buffer: bytes, options }), isBudget);
  assert.throws(() => resolveNativeDocumentResourceLimits({ maxArchiveBytes: DEFAULT_NATIVE_DOCUMENT_RESOURCE_LIMITS.maxArchiveBytes + 1 }), isBudget);
  assert.throws(() => resolveNativeDocumentResourceLimits({ maxCells: NaN }), isBudget);
  assert.throws(() => assertNativeInputSize(DEFAULT_NATIVE_DOCUMENT_RESOURCE_LIMITS.maxArchiveBytes + 1), isBudget);
});
test('same-row disjoint merges consume one workbook-wide comparison budget', async () => {
  const workbook = new WorkbookModel('merge-budget', 'Merge budget');
  const exported = await exportOoxmlDocument({ snapshot: workbook.snapshot(), fileName: 'merges.xlsx', options: { compatibilityTarget: 'B' } });
  const loaded = loadOpcPackageGraph(exported.buffer);
  const sheetPart = Object.keys(loaded.files).find(p => /worksheets\/sheet\d+\.xml$/.test(p))!;
  loaded.files[sheetPart] = strToU8('<worksheet><sheetData/><mergeCells>' + ['A1:B1', 'C1:D1', 'E1:F1', 'G1:H1'].map(ref => `<mergeCell ref="${ref}"/>`).join('') + '</mergeCells></worksheet>');
  assert.throws(() => parseLoadedOoxml(loaded, { limits: { maxMergeComparisons: 5 } }), isBudget);
  const parsed = parseLoadedOoxml(loaded, { limits: { maxMerges: 4, maxMergeComparisons: 6 } });
  assert.equal(parsed.snapshot.sheets[0]!.merges.length, 4);
  assert.throws(() => parseLoadedOoxml(loaded, { limits: { maxMerges: 3 } }), isBudget);
  const buffer = zipSync(loaded.files).buffer as ArrayBuffer;
  const imported = await importOoxmlDocument({ fileName: 'valid.xlsx', buffer, options: { compatibilityTarget: 'B', limits: { maxArchiveBytes: buffer.byteLength } } });
  assert.equal(imported.snapshot.sheets[0]!.merges.length, 4);
});
