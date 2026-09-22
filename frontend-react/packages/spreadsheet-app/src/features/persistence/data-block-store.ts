import type { DataBlockRef } from '@react-sheets/core-model';
import { computeBinaryChecksum } from './checksum';
import { memoryKey, WorkspaceStorageError, type WorkspaceMemoryCoordinator } from './memory';

export interface DataBlockRecord {
  schema: 'DataBlockRecord';
  sourceId: string;
  blockId: string;
  checksum: string;
  bytes: ArrayBuffer;
  updatedAt: string;
}

function cloneBytes(bytes: ArrayBuffer): ArrayBuffer { return bytes.slice(0); }

function cloneRecord(record: DataBlockRecord): DataBlockRecord {
  return { ...record, bytes: cloneBytes(record.bytes) };
}

async function assertRecord(record: DataBlockRecord): Promise<void> {
  if (record.schema !== 'DataBlockRecord' || !record.sourceId || !record.blockId || !record.checksum) {
    throw new Error('Invalid data block record');
  }
  if (!(record.bytes instanceof ArrayBuffer)) throw new Error('Data block bytes must be an ArrayBuffer');
  if (await computeBinaryChecksum(record.bytes) !== record.checksum) {
    throw new Error(`Data block checksum mismatch: ${record.blockId}`);
  }
}

function storageSourceId(unitId: string | (() => string) | undefined, sourceId: string): string {
  const resolved = typeof unitId === 'function' ? unitId() : unitId;
  return resolved?.trim() ? `unit:${resolved.trim()}:source:${sourceId}` : sourceId;
}

/** Session-memory bytes for block-backed sources. */
export class LocalDataBlockStore {
  constructor(
    private readonly coordinator: WorkspaceMemoryCoordinator,
    private readonly unitId?: string | (() => string),
  ) {}

  async put(ref: DataBlockRef, bytes: ArrayBuffer): Promise<DataBlockRecord> {
    if (bytes.byteLength !== ref.byteLength) {
      throw new WorkspaceStorageError({
        code: 'STORAGE_SCHEMA_INVALID', operation: 'data-block-put',
        message: `Data block length does not match manifest: ${ref.id}`,
        recovery: 'Read the bytes matching the canonical block descriptor before writing.',
      });
    }
    const checksum = await computeBinaryChecksum(bytes);
    if (checksum !== ref.checksum) throw new Error(`Data block checksum does not match manifest: ${ref.id}`);
    const record: DataBlockRecord = {
      schema: 'DataBlockRecord',
      sourceId: storageSourceId(this.unitId, ref.dataSourceId),
      blockId: ref.id,
      checksum,
      bytes: cloneBytes(bytes),
      updatedAt: new Date().toISOString(),
    };
    await assertRecord(record);
    return this.coordinator.transaction(async (transaction) => {
      const key = memoryKey(record.sourceId, record.blockId);
      const existing = transaction.get<DataBlockRecord>('dataBlocks', key);
      if (existing) {
        await assertRecord(existing);
        if (existing.checksum !== checksum || existing.bytes.byteLength !== bytes.byteLength) {
          throw new WorkspaceStorageError({
            code: 'STORAGE_REVISION_CONFLICT', operation: 'data-block-put',
            message: `Data block is immutable: ${ref.dataSourceId}/${ref.id}`,
            recovery: 'Upload changed content using a new block id.',
          });
        }
        return cloneRecord(existing);
      }
      transaction.set('dataBlocks', key, cloneRecord(record));
      return cloneRecord(record);
    });
  }

  async get(ref: Pick<DataBlockRef, 'dataSourceId' | 'id' | 'checksum'> & Partial<Pick<DataBlockRef, 'byteLength'>>): Promise<DataBlockRecord | null> {
    return this.coordinator.read(async (transaction) => {
      const sourceId = storageSourceId(this.unitId, ref.dataSourceId);
      const record = transaction.get<DataBlockRecord>('dataBlocks', memoryKey(sourceId, ref.id));
      if (!record) return null;
      await assertRecord(record);
      if (record.checksum !== ref.checksum) throw new Error(`Data block manifest checksum mismatch: ${ref.id}`);
      if (ref.byteLength !== undefined && record.bytes.byteLength !== ref.byteLength) throw new Error(`Data block manifest byteLength mismatch: ${ref.id}`);
      return cloneRecord(record);
    });
  }

  async remove(sourceId: string, blockId: string): Promise<void> {
    await this.coordinator.transaction((transaction) => {
      transaction.delete('dataBlocks', memoryKey(storageSourceId(this.unitId, sourceId), blockId));
    });
  }

  async removeSource(sourceId: string): Promise<void> {
    const storageId = storageSourceId(this.unitId, sourceId);
    await this.coordinator.transaction((transaction) => {
      for (const row of transaction.getAll<DataBlockRecord>('dataBlocks')) {
        if (row.sourceId === storageId) transaction.delete('dataBlocks', memoryKey(row.sourceId, row.blockId));
      }
    });
  }
}
