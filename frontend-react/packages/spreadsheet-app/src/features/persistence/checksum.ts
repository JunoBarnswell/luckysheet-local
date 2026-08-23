const ROUND_CONSTANTS = new Uint32Array([
  0x428a2f98, 0x71374491, 0xb5c0fbcf, 0xe9b5dba5, 0x3956c25b, 0x59f111f1, 0x923f82a4, 0xab1c5ed5,
  0xd807aa98, 0x12835b01, 0x243185be, 0x550c7dc3, 0x72be5d74, 0x80deb1fe, 0x9bdc06a7, 0xc19bf174,
  0xe49b69c1, 0xefbe4786, 0x0fc19dc6, 0x240ca1cc, 0x2de92c6f, 0x4a7484aa, 0x5cb0a9dc, 0x76f988da,
  0x983e5152, 0xa831c66d, 0xb00327c8, 0xbf597fc7, 0xc6e00bf3, 0xd5a79147, 0x06ca6351, 0x14292967,
  0x27b70a85, 0x2e1b2138, 0x4d2c6dfc, 0x53380d13, 0x650a7354, 0x766a0abb, 0x81c2c92e, 0x92722c85,
  0xa2bfe8a1,
  0xa81a664b, 0xc24b8b70, 0xc76c51a3, 0xd192e819, 0xd6990624, 0xf40e3585, 0x106aa070, 0x19a4c116,
  0x1e376c08, 0x2748774c, 0x34b0bcb5, 0x391c0cb3, 0x4ed8aa4a, 0x5b9cca4f, 0x682e6ff3, 0x748f82ee,
  0x78a5636f, 0x84c87814, 0x8cc70208, 0x90befffa, 0xa4506ceb, 0xbef9a3f7, 0xc67178f2,
]);

function rotateRight(value: number, shift: number): number {
  return (value >>> shift) | (value << (32 - shift));
}

function word(words: Uint32Array, index: number): number {
  return words[index] ?? 0;
}

function hex(value: number): string {
  return (value >>> 0).toString(16).padStart(8, '0');
}

export function computeChecksum(message: string): string {
  const bytes = new TextEncoder().encode(message);
  const bitLength = bytes.length * 8;
  const paddedLength = ((bytes.length + 9 + 63) >> 6) << 6;
  const padded = new Uint8Array(paddedLength);
  padded.set(bytes);
  padded[bytes.length] = 0x80;
  const view = new DataView(padded.buffer);
  view.setUint32(paddedLength - 4, bitLength, false);

  let a0 = 0x6a09e667;
  let a1 = 0xbb67ae85;
  let a2 = 0x3c6ef372;
  let a3 = 0xa54ff53a;
  let a4 = 0x510e527f;
  let a5 = 0x9b05688c;
  let a6 = 0x1f83d9ab;
  let a7 = 0x5be0cd19;
  const schedule = new Uint32Array(64);

  for (let offset = 0; offset < padded.length; offset += 64) {
    for (let index = 0; index < 16; index += 1) schedule[index] = view.getUint32(offset + index * 4, false);
    for (let index = 16; index < 64; index += 1) {
      const s0 = rotateRight(word(schedule, index - 15), 7) ^ rotateRight(word(schedule, index - 15), 18) ^ (word(schedule, index - 15) >>> 3);
      const s1 = rotateRight(word(schedule, index - 2), 17) ^ rotateRight(word(schedule, index - 2), 19) ^ (word(schedule, index - 2) >>> 10);
      schedule[index] = (word(schedule, index - 16) + s0 + word(schedule, index - 7) + s1) >>> 0;
    }

    let a = a0;
    let b = a1;
    let c = a2;
    let d = a3;
    let e = a4;
    let f = a5;
    let g = a6;
    let h = a7;
    for (let index = 0; index < 64; index += 1) {
      const s1 = rotateRight(e, 6) ^ rotateRight(e, 11) ^ rotateRight(e, 25);
      const choice = (e & f) ^ (~e & g);
      const temp1 = (h + s1 + choice + word(ROUND_CONSTANTS, index) + word(schedule, index)) >>> 0;
      const s0 = rotateRight(a, 2) ^ rotateRight(a, 13) ^ rotateRight(a, 22);
      const majority = (a & b) ^ (a & c) ^ (b & c);
      const temp2 = (s0 + majority) >>> 0;
      h = g;
      g = f;
      f = e;
      e = (d + temp1) >>> 0;
      d = c;
      c = b;
      b = a;
      a = (temp1 + temp2) >>> 0;
    }
    a0 = (a0 + a) >>> 0;
    a1 = (a1 + b) >>> 0;
    a2 = (a2 + c) >>> 0;
    a3 = (a3 + d) >>> 0;
    a4 = (a4 + e) >>> 0;
    a5 = (a5 + f) >>> 0;
    a6 = (a6 + g) >>> 0;
    a7 = (a7 + h) >>> 0;
  }
  return `${hex(a0)}${hex(a1)}${hex(a2)}${hex(a3)}${hex(a4)}${hex(a5)}${hex(a6)}${hex(a7)}`;
}

/**
 * Binary payloads must never be coerced into a JavaScript string just to be
 * checksummed. WebCrypto keeps the byte representation intact and is used for
 * block and OOXML artifacts that can be tens of megabytes.
 */
export async function computeBinaryChecksum(bytes: ArrayBuffer): Promise<string> {
  const subtle = globalThis.crypto?.subtle;
  if (!subtle) throw new Error('WebCrypto is required to checksum binary workspace artifacts');
  const digest = await subtle.digest('SHA-256', bytes);
  return [...new Uint8Array(digest)].map((value) => value.toString(16).padStart(2, '0')).join('');
}

export function verifyChecksum(message: string, expected: string): boolean {
  return computeChecksum(message) === expected;
}
