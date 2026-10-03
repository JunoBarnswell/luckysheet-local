export type SdkErrorCode = 'STALE_OPERATION' | 'INVALID_ARGUMENT' | 'UNAUTHENTICATED' | 'FORBIDDEN' | 'SERVICE_UNAVAILABLE' | 'CONTRACT_INVALID' | 'AUTH_CONFIGURATION_ERROR' | 'REQUEST_REJECTED' | 'RUNTIME_DISPOSED' | 'UNSUPPORTED_FEATURE';

export class SdkError extends Error {
  readonly code: SdkErrorCode;
  readonly operation: string;
  readonly recovery: string;
  readonly status?: number;
  constructor(code: SdkErrorCode, operation: string, message: string, recovery: string, options: { status?: number; cause?: unknown } = {}) {
    super(message, { cause: options.cause });
    this.name = 'SdkError';
    this.code = code;
    this.operation = operation;
    this.recovery = recovery;
    this.status = options.status;
  }
}
