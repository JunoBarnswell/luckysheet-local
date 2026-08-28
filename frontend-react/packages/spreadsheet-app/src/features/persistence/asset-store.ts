import { assertAssetRef, isSupportedAssetMime, type AssetRef } from '@react-sheets/core-model';
import { WorkbookApiClient } from '@react-sheets/protocol';
import { memoryKey, type WorkspaceMemoryCoordinator } from './memory';

export interface AssetPutInput {
  content: Blob;
  mimeType?: string;
  width?: number;
  height?: number;
}

export interface AssetStore {
  put(input: AssetPutInput): Promise<AssetRef>;
  get(ref: AssetRef): Promise<Blob>;
  release(ref: AssetRef): Promise<void>;
  reconcile(references: readonly AssetRef[]): Promise<void>;
}

export interface LocalAssetRecord extends AssetRef {
  unitId: string;
  bytes: ArrayBuffer;
}

function clone<T>(value: T): T {
  return structuredClone(value);
}

async function sha256(content: Blob): Promise<string> {
  if (!globalThis.crypto?.subtle) throw new Error('ASSET_HASH_UNAVAILABLE: Web Crypto SHA-256 is required');
  const digest = await globalThis.crypto.subtle.digest('SHA-256', await content.arrayBuffer());
  return Array.from(new Uint8Array(digest), (value) => value.toString(16).padStart(2, '0')).join('');
}

function refFromContent(content: Blob, contentHash: string, input: AssetPutInput): AssetRef {
  const ref: AssetRef = {
    schema: 'AssetRef',
    assetId: `asset-${contentHash}`,
    contentHash,
    mimeType: input.mimeType ?? content.type,
    byteLength: content.size,
    ...(input.width === undefined ? {} : { width: input.width }),
    ...(input.height === undefined ? {} : { height: input.height }),
  };
  assertAssetRef(ref);
  return ref;
}

function assetRef(record: LocalAssetRecord): AssetRef {
  const { unitId: _unitId, bytes: _bytes, ...ref } = record;
  return clone(ref);
}

function validateContent(ref: AssetRef, bytes: ArrayBuffer): void {
  if (bytes.byteLength !== ref.byteLength) throw new Error(`ASSET_LENGTH_MISMATCH: ${ref.assetId}`);
}

/** Browser-local, content-addressed store. It never serializes bytes into WorkbookSnapshot. */
export class LocalAssetStore implements AssetStore {
  private readonly unitId: string;

  constructor(unitId: string, private readonly coordinator: WorkspaceMemoryCoordinator) {
    if (!unitId.trim()) throw new Error('AssetStore unitId is required');
    this.unitId = unitId;
  }

  async put(input: AssetPutInput): Promise<AssetRef> {
    if (!input.content || input.content.size <= 0) throw new Error('ASSET_EMPTY: image content is empty');
    const mimeType = input.mimeType ?? input.content.type;
    if (!isSupportedAssetMime(mimeType)) throw new Error(`ASSET_MIME_UNSUPPORTED: ${mimeType}`);
    const contentHash = await sha256(input.content);
    const ref = refFromContent(input.content, contentHash, { ...input, mimeType });
    const bytes = await input.content.arrayBuffer();
    const record: LocalAssetRecord = { ...ref, unitId: this.unitId, bytes };
    return this.coordinator.transaction((transaction) => {
      const key = memoryKey(this.unitId, ref.assetId);
      const existing = transaction.get<LocalAssetRecord>('assets', key);
      if (existing) validateContent(ref, existing.bytes);
      else transaction.set('assets', key, clone(record));
      return clone(ref);
    });
  }

  async get(ref: AssetRef): Promise<Blob> {
    assertAssetRef(ref);
    const record = await this.coordinator.read((transaction) => transaction.get<LocalAssetRecord>('assets', memoryKey(this.unitId, ref.assetId)));
    if (!record) throw new Error(`ASSET_MISSING: ${ref.assetId}`);
    if (record.contentHash !== ref.contentHash || record.mimeType !== ref.mimeType) throw new Error(`ASSET_METADATA_MISMATCH: ${ref.assetId}`);
    validateContent(ref, record.bytes);
    const blob = new Blob([record.bytes], { type: record.mimeType });
    if (await sha256(blob) !== ref.contentHash) throw new Error(`ASSET_HASH_MISMATCH: ${ref.assetId}`);
    return blob;
  }

  async release(ref: AssetRef): Promise<void> {
    assertAssetRef(ref);
    await this.coordinator.transaction((transaction) => transaction.delete('assets', memoryKey(this.unitId, ref.assetId)));
  }

  async reconcile(references: readonly AssetRef[]): Promise<void> {
    const referenced = new Set(references.map((ref) => ref.assetId));
    await this.coordinator.transaction((transaction) => {
      for (const record of transaction.getAll<LocalAssetRecord>('assets')) {
        if (record.unitId === this.unitId && !referenced.has(record.assetId)) {
          transaction.delete('assets', memoryKey(record.unitId, record.assetId));
        }
      }
    });
  }
}

/** Remote authoritative store. WorkbookSnapshot contains references only; bytes travel through this API. */
export class RemoteAssetStore implements AssetStore {
  constructor(private readonly unitId: string, private readonly api: WorkbookApiClient) {
    if (!unitId.trim()) throw new Error('AssetStore unitId is required');
  }

  async put(input: AssetPutInput): Promise<AssetRef> {
    if (!input.content || input.content.size <= 0) throw new Error('ASSET_EMPTY: image content is empty');
    const mimeType = input.mimeType ?? input.content.type;
    if (!isSupportedAssetMime(mimeType)) throw new Error(`ASSET_MIME_UNSUPPORTED: ${mimeType}`);
    const contentHash = await sha256(input.content);
    const ref = refFromContent(input.content, contentHash, { ...input, mimeType });
    const metadata = await this.api.putAsset(this.unitId, ref, await input.content.arrayBuffer());
    assertAssetRef(metadata);
    if (metadata.assetId !== ref.assetId || metadata.contentHash !== ref.contentHash || metadata.byteLength !== ref.byteLength || metadata.mimeType !== ref.mimeType) {
      throw new Error(`ASSET_METADATA_MISMATCH: ${ref.assetId}`);
    }
    return ref;
  }

  async get(ref: AssetRef): Promise<Blob> {
    assertAssetRef(ref);
    const result = await this.api.getAsset(this.unitId, ref.assetId);
    if (result.contentHash !== ref.contentHash || result.mimeType !== ref.mimeType || result.byteLength !== ref.byteLength) throw new Error(`ASSET_METADATA_MISMATCH: ${ref.assetId}`);
    const blob = new Blob([result.bytes], { type: result.mimeType });
    if (await sha256(blob) !== ref.contentHash) throw new Error(`ASSET_HASH_MISMATCH: ${ref.assetId}`);
    return blob;
  }

  async release(ref: AssetRef): Promise<void> {
    assertAssetRef(ref);
    await this.api.deleteAsset(this.unitId, ref.assetId);
  }

  async reconcile(references: readonly AssetRef[]): Promise<void> {
    for (const ref of references) assertAssetRef(ref);
    await this.api.reconcileAssets(this.unitId, references.map((ref) => ref.assetId));
  }
}
