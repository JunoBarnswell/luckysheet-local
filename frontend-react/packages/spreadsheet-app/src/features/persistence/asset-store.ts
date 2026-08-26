import { assertAssetRef, type AssetRef } from '@react-sheets/core-model';
import { WorkbookApiClient } from '@react-sheets/protocol';
import {
  ASSET_STORE_NAME,
  openWorkspaceDatabase,
  requestResult,
  transactionComplete,
  type IndexedDbStoreOptions,
} from './indexed-db';

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

interface LocalAssetRecord extends AssetRef {
  unitId: string;
  bytes: ArrayBuffer;
}

const memoryAssets = new Map<string, Map<string, LocalAssetRecord>>();

function clone<T>(value: T): T {
  return structuredClone(value);
}

function memoryRecords(databaseName: string, unitId: string): Map<string, LocalAssetRecord> {
  const key = `${databaseName}:${unitId}`;
  let records = memoryAssets.get(key);
  if (!records) {
    records = new Map<string, LocalAssetRecord>();
    memoryAssets.set(key, records);
  }
  return records;
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
  private readonly options: IndexedDbStoreOptions;
  private readonly unitId: string;
  private databasePromise: Promise<IDBDatabase | null> | null = null;

  constructor(unitId: string, options: IndexedDbStoreOptions = {}) {
    if (!unitId.trim()) throw new Error('AssetStore unitId is required');
    this.unitId = unitId;
    this.options = options;
  }

  private database(): Promise<IDBDatabase | null> {
    if (!this.databasePromise) {
      const pending = openWorkspaceDatabase(this.options);
      this.databasePromise = pending.catch((error) => {
        this.databasePromise = null;
        throw error;
      });
    }
    return this.databasePromise;
  }

  async put(input: AssetPutInput): Promise<AssetRef> {
    if (!input.content || input.content.size <= 0) throw new Error('ASSET_EMPTY: image content is empty');
    const mimeType = input.mimeType ?? input.content.type;
    if (!mimeType.startsWith('image/')) throw new Error(`ASSET_MIME_UNSUPPORTED: ${mimeType}`);
    const contentHash = await sha256(input.content);
    const ref = refFromContent(input.content, contentHash, { ...input, mimeType });
    const database = await this.database();
    const bytes = await input.content.arrayBuffer();
    const record: LocalAssetRecord = { ...ref, unitId: this.unitId, bytes };
    if (!database) {
      const records = memoryRecords(this.options.databaseName ?? 'react-sheets-workspaces', this.unitId);
      const existing = records.get(ref.assetId);
      if (existing) validateContent(ref, existing.bytes);
      else records.set(ref.assetId, clone(record));
      return clone(ref);
    }
    const transaction = database.transaction(ASSET_STORE_NAME, 'readwrite');
    const store = transaction.objectStore(ASSET_STORE_NAME);
    const existing = await requestResult(store.get([this.unitId, ref.assetId])) as LocalAssetRecord | undefined;
    if (existing) validateContent(ref, existing.bytes);
    else store.put(clone(record));
    await transactionComplete(transaction);
    return clone(ref);
  }

  async get(ref: AssetRef): Promise<Blob> {
    assertAssetRef(ref);
    const database = await this.database();
    let record: LocalAssetRecord | undefined;
    if (!database) record = memoryRecords(this.options.databaseName ?? 'react-sheets-workspaces', this.unitId).get(ref.assetId);
    else {
      const transaction = database.transaction(ASSET_STORE_NAME, 'readonly');
      const complete = transactionComplete(transaction);
      record = await requestResult(transaction.objectStore(ASSET_STORE_NAME).get([this.unitId, ref.assetId])) as LocalAssetRecord | undefined;
      await complete;
    }
    if (!record) throw new Error(`ASSET_MISSING: ${ref.assetId}`);
    if (record.contentHash !== ref.contentHash || record.mimeType !== ref.mimeType) throw new Error(`ASSET_METADATA_MISMATCH: ${ref.assetId}`);
    validateContent(ref, record.bytes);
    const blob = new Blob([record.bytes], { type: record.mimeType });
    if (await sha256(blob) !== ref.contentHash) throw new Error(`ASSET_HASH_MISMATCH: ${ref.assetId}`);
    return blob;
  }

  async release(ref: AssetRef): Promise<void> {
    assertAssetRef(ref);
    const database = await this.database();
    if (!database) {
      memoryRecords(this.options.databaseName ?? 'react-sheets-workspaces', this.unitId).delete(ref.assetId);
      return;
    }
    const transaction = database.transaction(ASSET_STORE_NAME, 'readwrite');
    transaction.objectStore(ASSET_STORE_NAME).delete([this.unitId, ref.assetId]);
    await transactionComplete(transaction);
  }

  async reconcile(references: readonly AssetRef[]): Promise<void> {
    const referenced = new Set(references.map((ref) => ref.assetId));
    const database = await this.database();
    if (!database) {
      const records = memoryRecords(this.options.databaseName ?? 'react-sheets-workspaces', this.unitId);
      for (const assetId of records.keys()) if (!referenced.has(assetId)) records.delete(assetId);
      return;
    }
    const transaction = database.transaction(ASSET_STORE_NAME, 'readwrite');
    const store = transaction.objectStore(ASSET_STORE_NAME);
    const records = await requestResult(store.getAll()) as LocalAssetRecord[];
    for (const record of records) if (record.unitId === this.unitId && !referenced.has(record.assetId)) store.delete([record.unitId, record.assetId]);
    await transactionComplete(transaction);
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
    if (!mimeType.startsWith('image/')) throw new Error(`ASSET_MIME_UNSUPPORTED: ${mimeType}`);
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
