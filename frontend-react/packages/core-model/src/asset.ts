export const ASSET_REF_SCHEMA = 'AssetRef' as const;

export interface AssetRef {
  schema: typeof ASSET_REF_SCHEMA;
  assetId: string;
  contentHash: string;
  mimeType: string;
  byteLength: number;
  width?: number;
  height?: number;
}

export function isAssetRef(value: unknown): value is AssetRef {
  if (!value || typeof value !== 'object' || Array.isArray(value)) return false;
  const candidate = value as Partial<AssetRef>;
  return candidate.schema === ASSET_REF_SCHEMA
    && typeof candidate.assetId === 'string'
    && candidate.assetId.trim().length > 0
    && typeof candidate.contentHash === 'string'
    && /^[a-f0-9]{64}$/.test(candidate.contentHash)
    && typeof candidate.mimeType === 'string'
    && candidate.mimeType.startsWith('image/')
    && typeof candidate.byteLength === 'number'
    && Number.isSafeInteger(candidate.byteLength)
    && candidate.byteLength >= 0
    && (candidate.width === undefined || (Number.isSafeInteger(candidate.width) && candidate.width > 0))
    && (candidate.height === undefined || (Number.isSafeInteger(candidate.height) && candidate.height > 0));
}

export function assertAssetRef(value: unknown, label = 'AssetRef'): asserts value is AssetRef {
  if (!isAssetRef(value)) throw new Error(`${label} is invalid`);
}
