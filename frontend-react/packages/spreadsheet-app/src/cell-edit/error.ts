import type { CellEditAddress, CellEditErrorCode, CellEditFailure } from './contracts';

export class CellEditError extends Error implements CellEditFailure {
  readonly code: CellEditErrorCode;
  readonly target?: CellEditAddress;
  readonly recovery: string;
  readonly alertStyle?: 'stop' | 'warning' | 'information';
  readonly title?: string;

  constructor(failure: CellEditFailure, options?: ErrorOptions) {
    super(failure.message, options);
    this.name = 'CellEditError';
    this.code = failure.code;
    this.target = failure.target;
    this.recovery = failure.recovery;
    this.alertStyle = failure.alertStyle;
    this.title = failure.title;
  }

  toFailure(): CellEditFailure {
    return {
      code: this.code,
      message: this.message,
      recovery: this.recovery,
      ...(this.target ? { target: this.target } : {}),
      ...(this.alertStyle ? { alertStyle: this.alertStyle } : {}),
      ...(this.title ? { title: this.title } : {}),
    };
  }
}
export function isCellEditError(value: unknown): value is CellEditError {
  return value instanceof CellEditError;
}
