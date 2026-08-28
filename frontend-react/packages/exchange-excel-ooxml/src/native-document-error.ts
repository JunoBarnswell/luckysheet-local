import type { NativeDocumentFormat } from './types';

export type NativeDocumentErrorCode =
  | 'NATIVE_DOCUMENT_DETECTION_FAILED'
  | 'NATIVE_DOCUMENT_CODEC_FAILED'
  | 'NATIVE_DOCUMENT_RESOURCE_LIMIT'
  | 'NATIVE_DOCUMENT_INVALID'
  | 'NATIVE_DOCUMENT_UNSUPPORTED'
  | 'NATIVE_DOCUMENT_UNCHANGED_SAVE_REQUIRED';

/** Observable failure at the native document boundary. */
export class NativeDocumentError extends Error {
  readonly code: NativeDocumentErrorCode | string;
  readonly format?: NativeDocumentFormat;
  readonly location?: string;
  readonly recovery?: string;

  constructor(input: {
    code: NativeDocumentErrorCode | string;
    message: string;
    format?: NativeDocumentFormat;
    location?: string;
    recovery?: string;
    cause?: unknown;
  }) {
    super(`${input.code}: ${input.message}`, input.cause === undefined ? undefined : { cause: input.cause });
    this.name = 'NativeDocumentError';
    this.code = input.code;
    this.format = input.format;
    this.location = input.location;
    this.recovery = input.recovery;
  }
}

export function asNativeDocumentError(error: unknown, context: { fileName?: string; format?: NativeDocumentFormat } = {}): NativeDocumentError {
  if (error instanceof NativeDocumentError) return error;
  const message = error instanceof Error ? error.message : String(error);
  return new NativeDocumentError({
    code: 'NATIVE_DOCUMENT_CODEC_FAILED',
    message: context.fileName ? `${context.fileName}: ${message}` : message,
    format: context.format,
    cause: error,
  });
}
