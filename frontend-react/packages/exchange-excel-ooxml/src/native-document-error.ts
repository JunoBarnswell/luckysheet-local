import type { NativeDocumentFormat } from './types';

export type NativeDocumentErrorCode =
  | 'NATIVE_DOCUMENT_TRANSPORT_UNAVAILABLE'
  | 'NATIVE_DOCUMENT_TRANSPORT_FAILED'
  | 'NATIVE_DOCUMENT_INVALID'
  | 'NATIVE_DOCUMENT_UNSUPPORTED'
  | 'NATIVE_DOCUMENT_REVISION_CONFLICT';

export class NativeDocumentError extends Error {
  readonly code: NativeDocumentErrorCode | string;
  readonly format?: NativeDocumentFormat;
  readonly location?: string;
  readonly recovery?: string;
  constructor(input: { code: NativeDocumentErrorCode | string; message: string; format?: NativeDocumentFormat; location?: string; recovery?: string; cause?: unknown }) {
    super(`${input.code}: ${input.message}`, input.cause === undefined ? undefined : { cause: input.cause });
    this.name = 'NativeDocumentError'; this.code = input.code; this.format = input.format; this.location = input.location; this.recovery = input.recovery;
  }
}

export function asNativeDocumentError(error: unknown, context: { format?: NativeDocumentFormat; location?: string } = {}): NativeDocumentError {
  if (error instanceof NativeDocumentError) return error;
  return new NativeDocumentError({ code: 'NATIVE_DOCUMENT_TRANSPORT_FAILED', message: error instanceof Error ? error.message : String(error), format: context.format, location: context.location, cause: error });
}
