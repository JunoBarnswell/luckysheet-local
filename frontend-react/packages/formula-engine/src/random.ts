import type { CellAddress } from './ast';

export interface CalculationEntropyContext {
  readonly cycleId: number;
  readonly entropySeed: string;
  readonly passIndex: number;
  /** One host clock sample shared by every volatile date formula in this cycle. */
  readonly calculationTimeUtcMs: number;
  /** Host-local UTC offset at the sampled instant; Excel date serials have no timezone. */
  readonly calculationTimeZoneOffsetMinutes: number;
}

export function isValidCalculationTimeUtcMs(value: unknown): value is number {
  return typeof value === 'number'
    && Number.isSafeInteger(value)
    && value >= 0
    && Number.isFinite(new Date(value).getTime());
}

export function createCalculationEntropyContext(
  seed: string,
  cycleId: number,
  passIndex = 0,
  calculationTimeUtcMs = Date.now(),
  calculationTimeZoneOffsetMinutes = new Date(calculationTimeUtcMs).getTimezoneOffset(),
): CalculationEntropyContext {
  if (!seed.trim()) throw new Error('Calculation entropy requires a non-empty seed');
  if (!Number.isSafeInteger(cycleId) || cycleId < 0) throw new Error('Calculation entropy cycleId must be a non-negative integer');
  if (!Number.isSafeInteger(passIndex) || passIndex < 0) throw new Error('Calculation entropy passIndex must be a non-negative integer');
  if (!isValidCalculationTimeUtcMs(calculationTimeUtcMs)) {
    throw new Error('Calculation entropy time must be a non-negative UTC millisecond timestamp');
  }
  if (!Number.isSafeInteger(calculationTimeZoneOffsetMinutes)
    || Math.abs(calculationTimeZoneOffsetMinutes) > 14 * 60) {
    throw new Error('Calculation entropy timezone offset must be within fourteen hours');
  }
  return { cycleId, entropySeed: seed, passIndex, calculationTimeUtcMs, calculationTimeZoneOffsetMinutes };
}

/** Stable, order-independent random value for one volatile formula occurrence. */
export function formulaRandom(
  entropy: CalculationEntropyContext,
  address: CellAddress,
  functionName: string,
  occurrence: string,
  elementIndex = 0,
): number {
  const key = [
    entropy.entropySeed,
    entropy.cycleId,
    entropy.passIndex,
    address.sheetId,
    address.row,
    address.column,
    functionName.toUpperCase(),
    occurrence,
    elementIndex,
  ].join('|');
  let state = fnv1a(key);
  state = mix32(state ^ 0x9e3779b9);
  return (state >>> 0) / 0x1_0000_0000;
}

function fnv1a(value: string): number {
  let hash = 0x811c9dc5;
  for (let index = 0; index < value.length; index += 1) {
    hash ^= value.charCodeAt(index);
    hash = Math.imul(hash, 0x01000193);
  }
  return hash >>> 0;
}

function mix32(value: number): number {
  let state = value >>> 0;
  state ^= state >>> 16;
  state = Math.imul(state, 0x21f0aaad);
  state ^= state >>> 15;
  state = Math.imul(state, 0x735a2d97);
  state ^= state >>> 15;
  return state >>> 0;
}
