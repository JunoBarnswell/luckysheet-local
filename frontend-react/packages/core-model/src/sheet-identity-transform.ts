type SheetId = string;

export type SheetIdentityTransformKind = 'rename' | 'duplicate' | 'delete';

/** UI intent for the worksheet identity mutation. */
export interface SheetIdentityTransformSpec {
  readonly kind: SheetIdentityTransformKind;
  readonly sourceSheetId: SheetId;
  readonly sourceName: string;
  readonly targetSheetId?: SheetId;
  readonly targetName?: string;
}

export interface SheetReferenceInvalidation {
  readonly participant: string;
  readonly ownerSheetId?: SheetId;
  readonly reference: string;
  readonly reason: 'deleted-sheet-reference' | 'unsupported-formula';
}

export class SheetIdentityTransformError extends Error {
  readonly code = 'SHEET_IDENTITY_TRANSFORM_REJECTED';

  constructor(message: string, readonly invalidations: readonly SheetReferenceInvalidation[] = []) {
    super(message);
    this.name = 'SheetIdentityTransformError';
  }
}

/** Validate and freeze the command payload; there is intentionally no apply callback. */
export function createSheetIdentityIntent(input: SheetIdentityTransformSpec): SheetIdentityTransformSpec {
  const sourceSheetId = input.sourceSheetId.trim();
  const sourceName = input.sourceName.trim();
  if (!sourceSheetId || !sourceName) throw new SheetIdentityTransformError('Sheet identity requires sourceSheetId and sourceName');
  if (input.kind === 'rename' && !input.targetName?.trim()) throw new SheetIdentityTransformError('Sheet rename requires a non-empty targetName');
  if (input.kind === 'duplicate' && (!input.targetSheetId?.trim() || !input.targetName?.trim())) {
    throw new SheetIdentityTransformError('Sheet duplicate requires targetSheetId and targetName');
  }
  return Object.freeze({
    ...input,
    sourceSheetId,
    sourceName,
    targetSheetId: input.targetSheetId?.trim(),
    targetName: input.targetName?.trim(),
  });
}
