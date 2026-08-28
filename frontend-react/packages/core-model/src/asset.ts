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

/** AssetRef 覆盖本地绘图对象的文件内容；具体对象仍通过各自 payload 表达语义。 */
export function isSupportedAssetMime(mimeType: string): boolean {
  return /^(?:image|model|text|application|audio|video)\//i.test(mimeType.trim());
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
    && isSupportedAssetMime(candidate.mimeType)
    && typeof candidate.byteLength === 'number'
    && Number.isSafeInteger(candidate.byteLength)
    && candidate.byteLength >= 0
    && (candidate.width === undefined || (Number.isSafeInteger(candidate.width) && candidate.width > 0))
    && (candidate.height === undefined || (Number.isSafeInteger(candidate.height) && candidate.height > 0));
}

export function assertAssetRef(value: unknown, label = 'AssetRef'): asserts value is AssetRef {
  if (!isAssetRef(value)) throw new Error(`${label} is invalid`);
}
