import type { NativeDocumentCapability, NativeDocumentFormat } from './types';

/** Capability declaration is descriptive; execution belongs to NativeDocumentTransport. */
export const NATIVE_DOCUMENT_CAPABILITY_MANIFEST: readonly NativeDocumentCapability[] = [
  { family: 'ooxml', variants: ['xlsx', 'xlsm', 'xltx', 'xltm', 'xlam'], import: 'server', export: 'server', preserveUnknownParts: true },
];

export function capabilityForFormat(format: NativeDocumentFormat): NativeDocumentCapability {
  const capability = NATIVE_DOCUMENT_CAPABILITY_MANIFEST.find((entry) => entry.family === format.family && entry.variants.includes(format.variant));
  if (!capability) throw new Error(`NATIVE_DOCUMENT_FORMAT_UNSUPPORTED: ${format.family}/${format.variant}`);
  return capability;
}

export function listNativeDocumentCapabilities(): readonly NativeDocumentCapability[] { return NATIVE_DOCUMENT_CAPABILITY_MANIFEST; }
