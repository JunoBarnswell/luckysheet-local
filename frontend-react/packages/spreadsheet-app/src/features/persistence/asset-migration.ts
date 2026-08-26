import { migrateStoredWorkbookSnapshot, type WorkbookSnapshot } from '@react-sheets/core-model';
import type { AssetStore } from './asset-store';
import { computeChecksum } from './checksum';
import type { WorkspaceRecord } from './storage';

/** One-time boundary migration for v6 snapshots that embedded image bytes. */
export async function migrateLegacyImageAssets(value: unknown, assetStore: AssetStore): Promise<WorkbookSnapshot> {
  const input = structuredClone(value) as Record<string, unknown>;
  await assetizeImageRecords(input, assetStore);
  return migrateStoredWorkbookSnapshot(input);
}

export async function normalizeWorkspaceRecordWithAssets(record: WorkspaceRecord, assetStore: AssetStore): Promise<WorkspaceRecord> {
  const normalizedRecord = structuredClone(record);
  const snapshot = await migrateLegacyImageAssets(normalizedRecord.snapshot, assetStore);
  await assetizeImageRecords(normalizedRecord.pending, assetStore);
  const pendingPayload = {
    schema: normalizedRecord.pending.schema,
    unitId: normalizedRecord.pending.unitId,
    nextClientSequence: normalizedRecord.pending.nextClientSequence,
    operations: normalizedRecord.pending.operations,
  };
  return {
    ...normalizedRecord,
    snapshot,
    checksum: computeChecksum(JSON.stringify(snapshot)),
    pending: { ...normalizedRecord.pending, checksum: computeChecksum(JSON.stringify(pendingPayload)) },
  };
}

async function assetizeImageRecords(value: unknown, assetStore: AssetStore): Promise<void> {
  if (!value || typeof value !== 'object') return;
  if (Array.isArray(value)) {
    for (const entry of value) await assetizeImageRecords(entry, assetStore);
    return;
  }
  const record = value as Record<string, unknown>;
  if (record.kind === 'image' && typeof record.src === 'string') {
    const asset = await assetStore.put({ content: dataUrlToBlob(record.src) });
    delete record.src;
    record.asset = asset;
  }
  for (const entry of Object.values(record)) await assetizeImageRecords(entry, assetStore);
}

function dataUrlToBlob(value: string): Blob {
  const match = /^data:(image\/[A-Za-z0-9.+-]{1,80})(;base64)?,(.*)$/s.exec(value);
  if (!match) throw new Error('ASSET_MIGRATION_INVALID_DATA_URL: image source is not a supported data URL');
  const [, mimeType, encoding, payload] = match;
  if (encoding) {
    const binary = atob(payload);
    const bytes = Uint8Array.from(binary, (character) => character.charCodeAt(0));
    return new Blob([bytes], { type: mimeType });
  }
  return new Blob([decodeURIComponent(payload)], { type: mimeType });
}
