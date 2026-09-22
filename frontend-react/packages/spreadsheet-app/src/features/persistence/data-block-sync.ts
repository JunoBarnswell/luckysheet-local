import type { DataBlockRef } from '@react-sheets/core-model';
import type { WorkbookApiClient } from '@react-sheets/protocol';
import { LocalDataBlockStore } from './data-block-store';

export interface DataBlockSyncOptions {
  unitId: () => string;
  isRemoteAvailable: () => boolean;
}

/**
 * The only sync boundary for large source bytes. Snapshot and command layers
 * exchange DataBlockRef metadata; this class transfers the matching bytes
 * only after local checksum validation has succeeded.
 */
export class DataBlockSynchronizer {
  constructor(
    private readonly local: LocalDataBlockStore,
    private readonly api: WorkbookApiClient,
    private readonly options: DataBlockSyncOptions,
  ) {}

  async put(ref: DataBlockRef, bytes: ArrayBuffer): Promise<void> {
    await this.local.put(ref, bytes);
    if (!this.options.isRemoteAvailable()) return;
    const metadata = await this.api.putDataBlock(this.options.unitId(), ref.dataSourceId, ref.id, ref.checksum, bytes);
    if (metadata.checksum !== ref.checksum || metadata.byteLength !== ref.byteLength) {
      throw new Error(`Remote data block acknowledgement mismatched manifest: ${ref.id}`);
    }
  }

  async get(ref: DataBlockRef): Promise<ArrayBuffer> {
    const local = await this.local.get(ref);
    if (local) return local.bytes;
    if (!this.options.isRemoteAvailable()) throw new Error(`Data block is unavailable offline: ${ref.id}`);
    const remote = await this.api.getDataBlock(this.options.unitId(), ref.dataSourceId, ref.id);
    if (remote.checksum !== ref.checksum) throw new Error(`Remote data block checksum mismatched manifest: ${ref.id}`);
    await this.local.put(ref, remote.bytes);
    return remote.bytes;
  }

  async remove(ref: DataBlockRef): Promise<void> {
    await this.local.remove(ref.dataSourceId, ref.id);
    if (this.options.isRemoteAvailable()) await this.api.deleteDataBlock(this.options.unitId(), ref.dataSourceId, ref.id);
  }
}
