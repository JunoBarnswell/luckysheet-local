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

/**
 * Compute the canonical SHA-256 identity for an asset without depending on a
 * host crypto implementation.  Native document import is intentionally
 * synchronous at the XML-tree boundary; using a small, deterministic digest
 * here lets an imported media part become a real AssetRef before the
 * snapshot is exposed to the rest of the runtime.  The persistence stores
 * still verify the same identity with Web Crypto when bytes are committed.
 */
export function sha256Hex(bytes: Uint8Array): string {
  const words = new Uint32Array(64);
  const state = new Uint32Array([
    0x6a09e667, 0xbb67ae85, 0x3c6ef372, 0xa54ff53a,
    0x510e527f, 0x9b05688c, 0x1f83d9ab, 0x5be0cd19,
  ]);
  const bitLength = bytes.length * 8;
  const paddedLength = (((bytes.length + 9) + 63) >> 6) << 6;
  const padded = new Uint8Array(paddedLength);
  padded.set(bytes);
  padded[bytes.length] = 0x80;
  const high = Math.floor(bitLength / 0x100000000);
  const low = bitLength >>> 0;
  padded[padded.length - 8] = (high >>> 24) & 0xff;
  padded[padded.length - 7] = (high >>> 16) & 0xff;
  padded[padded.length - 6] = (high >>> 8) & 0xff;
  padded[padded.length - 5] = high & 0xff;
  padded[padded.length - 4] = (low >>> 24) & 0xff;
  padded[padded.length - 3] = (low >>> 16) & 0xff;
  padded[padded.length - 2] = (low >>> 8) & 0xff;
  padded[padded.length - 1] = low & 0xff;

  const constants = [
    0x428a2f98, 0x71374491, 0xb5c0fbcf, 0xe9b5dba5, 0x3956c25b,
    0x59f111f1, 0x923f82a4, 0xab1c5ed5, 0xd807aa98, 0x12835b01,
    0x243185be, 0x550c7dc3, 0x72be5d74, 0x80deb1fe, 0x9bdc06a7,
    0xc19bf174, 0xe49b69c1, 0xefbe4786, 0x0fc19dc6, 0x240ca1cc,
    0x2de92c6f, 0x4a7484aa, 0x5cb0a9dc, 0x76f988da, 0x983e5152,
    0xa831c66d, 0xb00327c8, 0xbf597fc7, 0xc6e00bf3, 0xd5a79147,
    0x06ca6351, 0x14292967, 0x27b70a85, 0x2e1b2138, 0x4d2c6dfc,
    0x53380d13, 0x650a7354, 0x766a0abb, 0x81c2c92e, 0x92722c85,
    0xa2bfe8a1, 0xa81a664b, 0xc24b8b70, 0xc76c51a3, 0xd192e819,
    0xd6990624, 0xf40e3585, 0x106aa070, 0x19a4c116, 0x1e376c08,
    0x2748774c, 0x34b0bcb5, 0x391c0cb3, 0x4ed8aa4a, 0x5b9cca4f,
    0x682e6ff3, 0x748f82ee, 0x78a5636f, 0x84c87814, 0x8cc70208,
    0x90befffa, 0xa4506ceb, 0xbef9a3f7, 0xc67178f2,
  ];
  const rotr = (value: number, amount: number): number => (value >>> amount) | (value << (32 - amount));
  const add = (...values: number[]): number => values.reduce((sum, value) => (sum + value) >>> 0, 0);
  for (let offset = 0; offset < padded.length; offset += 64) {
    for (let index = 0; index < 16; index += 1) {
      const position = offset + index * 4;
      words[index] = ((padded[position]! << 24) | (padded[position + 1]! << 16) | (padded[position + 2]! << 8) | padded[position + 3]!) >>> 0;
    }
    for (let index = 16; index < 64; index += 1) {
      const value = words[index - 15]!;
      const other = words[index - 2]!;
      const sigma0 = rotr(value, 7) ^ rotr(value, 18) ^ (value >>> 3);
      const sigma1 = rotr(other, 17) ^ rotr(other, 19) ^ (other >>> 10);
      words[index] = add(words[index - 16]!, sigma0, words[index - 7]!, sigma1);
    }
    let a = state[0]!;
    let b = state[1]!;
    let c = state[2]!;
    let d = state[3]!;
    let e = state[4]!;
    let f = state[5]!;
    let g = state[6]!;
    let h = state[7]!;
    for (let index = 0; index < 64; index += 1) {
      const sum1 = rotr(e, 6) ^ rotr(e, 11) ^ rotr(e, 25);
      const choose = (e & f) ^ (~e & g);
      const temp1 = add(h, sum1, choose, constants[index]!, words[index]!);
      const sum0 = rotr(a, 2) ^ rotr(a, 13) ^ rotr(a, 22);
      const majority = (a & b) ^ (a & c) ^ (b & c);
      const temp2 = add(sum0, majority);
      h = g; g = f; f = e; e = add(d, temp1); d = c; c = b; b = a; a = add(temp1, temp2);
    }
    state[0] = add(state[0]!, a); state[1] = add(state[1]!, b); state[2] = add(state[2]!, c); state[3] = add(state[3]!, d);
    state[4] = add(state[4]!, e); state[5] = add(state[5]!, f); state[6] = add(state[6]!, g); state[7] = add(state[7]!, h);
  }
  return [...state].map((value) => value.toString(16).padStart(8, '0')).join('');
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
