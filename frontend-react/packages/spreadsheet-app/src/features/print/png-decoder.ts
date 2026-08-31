import { unzlibSync } from 'fflate';

/**
 * Resource limits are part of the synchronous print boundary.  A PNG may be
 * small on disk while expanding to a very large scanline buffer, so both the
 * encoded input and the decoded working set are bounded before allocation.
 */
export const DEFAULT_PNG_RESOURCE_LIMITS = Object.freeze({
  maxWidth: 8192,
  maxHeight: 8192,
  maxPixels: 16_777_216,
  maxInputBytes: 64 * 1024 * 1024,
  maxDecodedBytes: 128 * 1024 * 1024,
  maxChunks: 100_000,
});

export type PngDecodeErrorCode =
  | 'PNG_INVALID_SIGNATURE'
  | 'PNG_TRUNCATED'
  | 'PNG_CHUNK_INVALID'
  | 'PNG_CRC_MISMATCH'
  | 'PNG_STRUCTURE_INVALID'
  | 'PNG_UNSUPPORTED_FORMAT'
  | 'PNG_RESOURCE_LIMIT'
  | 'PNG_DEFLATE_INVALID'
  | 'PNG_SCANLINE_INVALID'
  | 'PNG_DIMENSIONS_INVALID'
  | 'PNG_PALETTE_INVALID';

export class PngDecodeError extends Error {
  readonly code: PngDecodeErrorCode;

  constructor(code: PngDecodeErrorCode, message: string, cause?: unknown) {
    super(`${code}: ${message}`, cause === undefined ? undefined : { cause });
    this.name = 'PngDecodeError';
    this.code = code;
  }
}

export interface PngDecodeOptions {
  maxWidth?: number;
  maxHeight?: number;
  maxPixels?: number;
  maxInputBytes?: number;
  maxDecodedBytes?: number;
  maxChunks?: number;
}

export interface DecodedPngRgb {
  width: number;
  height: number;
  /** Row-major RGB bytes with every source alpha composited onto white. */
  rgb: Uint8Array;
}

const PNG_SIGNATURE = new Uint8Array([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a]);
const CRC_TABLE = createCrcTable();

function createCrcTable(): Uint32Array {
  const table = new Uint32Array(256);
  for (let index = 0; index < table.length; index += 1) {
    let value = index;
    for (let bit = 0; bit < 8; bit += 1) value = (value & 1) === 0 ? value >>> 1 : 0xedb88320 ^ (value >>> 1);
    table[index] = value >>> 0;
  }
  return table;
}

function crcUpdate(crc: number, bytes: Uint8Array): number {
  let value = crc >>> 0;
  for (const byte of bytes) value = CRC_TABLE[(value ^ byte) & 0xff]! ^ (value >>> 8);
  return value >>> 0;
}

function crc32(type: Uint8Array, data: Uint8Array): number {
  return (crcUpdate(crcUpdate(0xffffffff, type), data) ^ 0xffffffff) >>> 0;
}

function readUint32(bytes: Uint8Array, offset: number): number {
  return ((bytes[offset]! << 24) | (bytes[offset + 1]! << 16) | (bytes[offset + 2]! << 8) | bytes[offset + 3]!) >>> 0;
}

function readUint16(bytes: Uint8Array, offset: number): number {
  return (bytes[offset]! << 8) | bytes[offset + 1]!;
}

function fail(code: PngDecodeErrorCode, message: string, cause?: unknown): never {
  throw new PngDecodeError(code, message, cause);
}

function concat(parts: readonly Uint8Array[]): Uint8Array {
  const size = parts.reduce((total, part) => total + part.byteLength, 0);
  const result = new Uint8Array(size);
  let offset = 0;
  for (const part of parts) {
    result.set(part, offset);
    offset += part.byteLength;
  }
  return result;
}

function paeth(left: number, above: number, upperLeft: number): number {
  const estimate = left + above - upperLeft;
  const leftDistance = Math.abs(estimate - left);
  const aboveDistance = Math.abs(estimate - above);
  const upperLeftDistance = Math.abs(estimate - upperLeft);
  if (leftDistance <= aboveDistance && leftDistance <= upperLeftDistance) return left;
  if (aboveDistance <= upperLeftDistance) return above;
  return upperLeft;
}

function compositeOnWhite(channel: number, alpha: number): number {
  // Integer arithmetic makes the white-background result stable across hosts.
  return Math.floor((channel * alpha + 255 * (255 - alpha) + 127) / 255);
}

function isCriticalChunk(type: string): boolean {
  return (type.charCodeAt(0) & 0x20) === 0;
}

function chunkType(bytes: Uint8Array, offset: number): string {
  return String.fromCharCode(bytes[offset]!, bytes[offset + 1]!, bytes[offset + 2]!, bytes[offset + 3]!);
}

function validateLimits(options: PngDecodeOptions): Required<PngDecodeOptions> {
  const limits = {
    maxWidth: options.maxWidth ?? DEFAULT_PNG_RESOURCE_LIMITS.maxWidth,
    maxHeight: options.maxHeight ?? DEFAULT_PNG_RESOURCE_LIMITS.maxHeight,
    maxPixels: options.maxPixels ?? DEFAULT_PNG_RESOURCE_LIMITS.maxPixels,
    maxInputBytes: options.maxInputBytes ?? DEFAULT_PNG_RESOURCE_LIMITS.maxInputBytes,
    maxDecodedBytes: options.maxDecodedBytes ?? DEFAULT_PNG_RESOURCE_LIMITS.maxDecodedBytes,
    maxChunks: options.maxChunks ?? DEFAULT_PNG_RESOURCE_LIMITS.maxChunks,
  };
  for (const [name, value] of Object.entries(limits)) {
    if (!Number.isSafeInteger(value) || value <= 0) fail('PNG_RESOURCE_LIMIT', `${name} must be a positive safe integer`);
  }
  return limits;
}

/**
 * Decode the PNG subset used by the print resource boundary.
 *
 * Supported: non-interlaced 8-bit grayscale, RGB, indexed, grayscale-alpha,
 * and RGBA PNGs; all five PNG row filters; PLTE and tRNS transparency.  The
 * result is always RGB and alpha is composited against an opaque white page.
 */
export function decodePngToRgb(bytes: Uint8Array, options: PngDecodeOptions = {}): DecodedPngRgb {
  const limits = validateLimits(options);
  if (!(bytes instanceof Uint8Array)) fail('PNG_CHUNK_INVALID', 'PNG input must be a Uint8Array');
  if (bytes.byteLength > limits.maxInputBytes) fail('PNG_RESOURCE_LIMIT', `encoded PNG exceeds ${limits.maxInputBytes} bytes`);
  if (bytes.byteLength < PNG_SIGNATURE.byteLength) fail('PNG_TRUNCATED', 'PNG signature is incomplete');
  for (let index = 0; index < PNG_SIGNATURE.length; index += 1) if (bytes[index] !== PNG_SIGNATURE[index]) fail('PNG_INVALID_SIGNATURE', 'PNG signature is invalid');

  let offset = PNG_SIGNATURE.length;
  let chunkCount = 0;
  let seenHeader = false;
  let seenPalette = false;
  let seenTransparency = false;
  let seenImageData = false;
  let endedImageData = false;
  let seenEnd = false;
  let width = 0;
  let height = 0;
  let bitDepth = 0;
  let colorType = 0;
  let compressedLength = 0;
  let palette: Uint8Array | undefined;
  let transparency: Uint8Array | undefined;
  const imageData: Uint8Array[] = [];

  while (offset < bytes.length) {
    chunkCount += 1;
    if (chunkCount > limits.maxChunks) fail('PNG_RESOURCE_LIMIT', `PNG contains more than ${limits.maxChunks} chunks`);
    if (offset + 12 > bytes.length) fail('PNG_TRUNCATED', 'PNG chunk header is incomplete');
    const length = readUint32(bytes, offset);
    const typeOffset = offset + 4;
    const dataOffset = offset + 8;
    const end = dataOffset + length;
    const crcOffset = end;
    if (!Number.isSafeInteger(end) || end < dataOffset || crcOffset + 4 > bytes.length) fail('PNG_TRUNCATED', 'PNG chunk data is incomplete');
    const typeBytes = bytes.subarray(typeOffset, dataOffset);
    const type = chunkType(bytes, typeOffset);
    if (!/^[A-Za-z]{4}$/.test(type)) fail('PNG_CHUNK_INVALID', `invalid PNG chunk type ${type}`);
    const data = bytes.subarray(dataOffset, end);
    const expectedCrc = readUint32(bytes, crcOffset);
    if (crc32(typeBytes, data) !== expectedCrc) fail('PNG_CRC_MISMATCH', `CRC mismatch in ${type}`);
    offset = crcOffset + 4;

    if (type === 'IHDR') {
      if (seenHeader || data.length !== 13 || chunkCount !== 1) fail('PNG_STRUCTURE_INVALID', 'IHDR must be the first and only header chunk');
      seenHeader = true;
      width = readUint32(data, 0);
      height = readUint32(data, 4);
      bitDepth = data[8]!;
      colorType = data[9]!;
      const compressionMethod = data[10]!;
      const filterMethod = data[11]!;
      const interlaceMethod = data[12]!;
      if (width < 1 || height < 1 || width > limits.maxWidth || height > limits.maxHeight || width * height > limits.maxPixels) fail('PNG_DIMENSIONS_INVALID', `PNG dimensions ${width}x${height} exceed the resource budget`);
      if (bitDepth !== 8 || ![0, 2, 3, 4, 6].includes(colorType)) fail('PNG_UNSUPPORTED_FORMAT', `only 8-bit PNG color types 0, 2, 3, 4 and 6 are supported (got depth ${bitDepth}, type ${colorType})`);
      if (compressionMethod !== 0 || filterMethod !== 0) fail('PNG_UNSUPPORTED_FORMAT', 'PNG compression and filter methods must be zero');
      if (interlaceMethod !== 0) fail('PNG_UNSUPPORTED_FORMAT', 'interlaced PNGs are not supported by the synchronous print decoder');
      continue;
    }
    if (!seenHeader) fail('PNG_STRUCTURE_INVALID', `${type} appears before IHDR`);
    if (type === 'PLTE') {
      if (colorType === 0 || colorType === 4) fail('PNG_STRUCTURE_INVALID', `PLTE is not valid for color type ${colorType}`);
      if (seenPalette || seenImageData || data.length === 0 || data.length % 3 !== 0 || data.length > 768) fail('PNG_PALETTE_INVALID', 'PLTE must appear once before IDAT and contain 1-256 RGB entries');
      seenPalette = true;
      palette = data.slice();
      continue;
    }
    if (type === 'tRNS') {
      if (seenTransparency || seenImageData) fail('PNG_STRUCTURE_INVALID', 'tRNS must appear once before IDAT');
      if (colorType === 3 && !seenPalette) fail('PNG_STRUCTURE_INVALID', 'indexed tRNS requires PLTE first');
      if (colorType === 4 || colorType === 6) fail('PNG_UNSUPPORTED_FORMAT', 'tRNS is not valid for explicit-alpha PNG color types');
      const validLength = colorType === 0 ? data.length === 2 : colorType === 2 ? data.length === 6 : colorType === 3 ? data.length <= (palette?.length ?? 0) / 3 : false;
      if (!validLength) fail('PNG_PALETTE_INVALID', `invalid tRNS length ${data.length} for color type ${colorType}`);
      seenTransparency = true;
      transparency = data.slice();
      continue;
    }
    if (type === 'IDAT') {
      if (endedImageData) fail('PNG_STRUCTURE_INVALID', 'IDAT chunks must be consecutive');
      if (colorType === 3 && !seenPalette) fail('PNG_PALETTE_INVALID', 'indexed PNG requires PLTE before IDAT');
      seenImageData = true;
      compressedLength += data.length;
      if (compressedLength > limits.maxInputBytes) fail('PNG_RESOURCE_LIMIT', `compressed IDAT data exceeds ${limits.maxInputBytes} bytes`);
      imageData.push(data.slice());
      continue;
    }
    if (type === 'IEND') {
      if (data.length !== 0 || !seenImageData || seenEnd) fail('PNG_STRUCTURE_INVALID', 'IEND must be empty and follow IDAT');
      seenEnd = true;
      break;
    }
    if (seenImageData) endedImageData = true;
    if (isCriticalChunk(type)) fail('PNG_UNSUPPORTED_FORMAT', `unsupported critical PNG chunk ${type}`);
  }

  if (!seenEnd || offset !== bytes.length) fail('PNG_STRUCTURE_INVALID', 'PNG must end immediately after IEND');
  if (!seenHeader || !seenImageData || imageData.length === 0) fail('PNG_STRUCTURE_INVALID', 'PNG must contain IHDR, IDAT and IEND');
  if (colorType === 3 && !palette) fail('PNG_PALETTE_INVALID', 'indexed PNG is missing PLTE');
  if (colorType === 0 && transparency && readUint16(transparency, 0) > 0xff) fail('PNG_PALETTE_INVALID', '8-bit grayscale tRNS sample is out of range');
  if (colorType === 2 && transparency && (readUint16(transparency, 0) > 0xff || readUint16(transparency, 2) > 0xff || readUint16(transparency, 4) > 0xff)) fail('PNG_PALETTE_INVALID', '8-bit RGB tRNS sample is out of range');

  const channels = colorType === 0 || colorType === 3 ? 1 : colorType === 2 ? 3 : colorType === 4 ? 2 : 4;
  const rowBytes = width * channels;
  const inflatedLength = height * (rowBytes + 1);
  const rgbLength = width * height * 3;
  if (inflatedLength > limits.maxDecodedBytes || rgbLength > limits.maxDecodedBytes) fail('PNG_RESOURCE_LIMIT', `decoded PNG exceeds ${limits.maxDecodedBytes} bytes`);
  let inflated: Uint8Array;
  try {
    inflated = unzlibSync(concat(imageData));
  } catch (error) {
    fail('PNG_DEFLATE_INVALID', 'IDAT zlib stream could not be decoded', error);
  }
  if (inflated.length !== inflatedLength) fail('PNG_SCANLINE_INVALID', `decoded scanline length ${inflated.length} does not match ${inflatedLength}`);

  const previous = new Uint8Array(rowBytes);
  const current = new Uint8Array(rowBytes);
  const rgb = new Uint8Array(rgbLength);
  const transparentGray = colorType === 0 && transparency ? readUint16(transparency, 0) : -1;
  const transparentRgb = colorType === 2 && transparency ? [readUint16(transparency, 0), readUint16(transparency, 2), readUint16(transparency, 4)] : undefined;
  const paletteEntries = (palette?.length ?? 0) / 3;
  let inputOffset = 0;
  let outputOffset = 0;
  for (let rowIndex = 0; rowIndex < height; rowIndex += 1) {
    const filter = inflated[inputOffset++]!;
    if (filter > 4) fail('PNG_SCANLINE_INVALID', `unsupported PNG row filter ${filter}`);
    for (let index = 0; index < rowBytes; index += 1) {
      const raw = inflated[inputOffset + index]!;
      const left = index >= channels ? current[index - channels]! : 0;
      const above = previous[index]!;
      const upperLeft = index >= channels ? previous[index - channels]! : 0;
      const predictor = filter === 0 ? 0 : filter === 1 ? left : filter === 2 ? above : filter === 3 ? Math.floor((left + above) / 2) : paeth(left, above, upperLeft);
      current[index] = (raw + predictor) & 0xff;
    }
    inputOffset += rowBytes;

    for (let pixel = 0; pixel < width; pixel += 1) {
      let red = 0;
      let green = 0;
      let blue = 0;
      let alpha = 255;
      if (colorType === 0) {
        const gray = current[pixel]!;
        red = gray; green = gray; blue = gray;
        if (gray === transparentGray) alpha = 0;
      } else if (colorType === 2) {
        const source = pixel * 3;
        red = current[source]!; green = current[source + 1]!; blue = current[source + 2]!;
        if (transparentRgb && red === transparentRgb[0] && green === transparentRgb[1] && blue === transparentRgb[2]) alpha = 0;
      } else if (colorType === 3) {
        const paletteIndex = current[pixel]!;
        if (paletteIndex >= paletteEntries || !palette) fail('PNG_PALETTE_INVALID', `palette index ${paletteIndex} is out of range`);
        const source = paletteIndex * 3;
        red = palette[source]!; green = palette[source + 1]!; blue = palette[source + 2]!;
        alpha = transparency?.[paletteIndex] ?? 255;
      } else if (colorType === 4) {
        const source = pixel * 2;
        red = current[source]!; green = red; blue = red; alpha = current[source + 1]!;
      } else {
        const source = pixel * 4;
        red = current[source]!; green = current[source + 1]!; blue = current[source + 2]!; alpha = current[source + 3]!;
      }
      rgb[outputOffset++] = compositeOnWhite(red, alpha);
      rgb[outputOffset++] = compositeOnWhite(green, alpha);
      rgb[outputOffset++] = compositeOnWhite(blue, alpha);
    }
    previous.set(current);
  }
  return { width, height, rgb };
}
