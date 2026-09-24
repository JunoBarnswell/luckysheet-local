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
    if (metadata.sourceId !== ref.dataSourceId || metadata.blockId !== ref.id || metadata.checksum !== ref.checksum || metadata.byteLength !== ref.byteLength) {
      throw new Error(`Remote data block acknowledgement mismatched manifest: ${ref.id}`);
    }
  }

  async get(ref: DataBlockRef): Promise<ArrayBuffer> {
    const local = await this.local.get(ref);
    if (local) return local.bytes;
    if (!this.options.isRemoteAvailable()) throw new Error(`Data block is unavailable offline: ${ref.id}`);
    const remote = await this.api.getDataBlock(this.options.unitId(), ref.dataSourceId, ref.id);
    if (remote.checksum !== ref.checksum || remote.byteLength !== ref.byteLength || remote.bytes.byteLength !== ref.byteLength) throw new Error(`Remote data block descriptor mismatched manifest: ${ref.id}`);
    await this.local.put(ref, remote.bytes);
    return remote.bytes;
  }

  async remove(ref: DataBlockRef, options: { remoteRequired?: boolean } = {}): Promise<void> {
    // Server query execution creates the block remotely before the browser
    // receives its descriptor. A stale connection flag must not turn cleanup
    // into a local-only delete and strand that remote object.
    let remoteFailure: unknown;
    if (this.options.isRemoteAvailable() || options.remoteRequired === true) {
      try {
        await this.api.deleteDataBlock(this.options.unitId(), ref.dataSourceId, ref.id);
      } catch (error) {
        remoteFailure = error;
      }
    }
    try {
      await this.local.remove(ref.dataSourceId, ref.id);
    } catch (localFailure) {
      if (remoteFailure !== undefined) throw new AggregateError([remoteFailure, localFailure], `Data block cleanup failed: ${ref.id}`);
      throw localFailure;
    }
    if (remoteFailure !== undefined) throw remoteFailure;
  }
}
