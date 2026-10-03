import { DEFAULT_NATIVE_DOCUMENT_RESOURCE_LIMITS, type NativeDocumentResourceLimits } from './types';
import { NativeDocumentError } from './native-document-error';

/** Caller limits can reduce the immutable host ceilings, never raise them. */
export function resolveNativeDocumentResourceLimits(overrides: Partial<NativeDocumentResourceLimits> = {}): NativeDocumentResourceLimits {
  const limits = { ...DEFAULT_NATIVE_DOCUMENT_RESOURCE_LIMITS };
  for (const key of Object.keys(limits) as (keyof NativeDocumentResourceLimits)[]) {
    const value = overrides[key];
    if (value === undefined) continue;
    if (!Number.isSafeInteger(value) || value <= 0 || value > limits[key]) resourceLimit(`Invalid native document limit: ${key}`);
    limits[key] = value;
  }
  return limits;
}
export function resourceLimit(message: string): never {
  throw new NativeDocumentError({ code: 'NATIVE_DOCUMENT_RESOURCE_LIMIT', message, recovery: 'Reduce the document size or complexity before importing.' });
}
export function requireNativeInputBudget(input: ArrayBuffer | Uint8Array, overrides?: Partial<NativeDocumentResourceLimits>): void {
  assertNativeInputSize(input.byteLength, overrides);
}
export function assertNativeInputSize(byteLength: number, overrides?: Partial<NativeDocumentResourceLimits>): void {
  if (!Number.isSafeInteger(byteLength) || byteLength < 0 || byteLength > resolveNativeDocumentResourceLimits(overrides).maxArchiveBytes) resourceLimit('Native document input exceeds its byte budget');
}
