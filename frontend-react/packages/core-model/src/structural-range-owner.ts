import type { RangeRef, SheetId } from './index';

/** Exact geometry facts; block contents and formula state have different owners. */
export type StructuralRangeOwnerDelta =
  | {
    readonly ownerKind: 'data-region';
    readonly sheetId: SheetId;
    readonly regionId: string;
    readonly before: { readonly range: Readonly<RangeRef>; readonly headerRow: number };
    readonly after: { readonly range: Readonly<RangeRef>; readonly headerRow: number };
  }
  | {
    readonly ownerKind: 'workbook-table' | 'data-source';
    readonly ownerId: string;
    readonly before: Readonly<RangeRef>;
    readonly after: Readonly<RangeRef>;
  }
  | {
    readonly ownerKind: 'sheet-table';
    readonly sheetId: SheetId;
    readonly ownerId: string;
    readonly before: Readonly<RangeRef>;
    readonly after: Readonly<RangeRef>;
  }
  | {
    readonly ownerKind: 'validation-list-source';
    /** Worksheet that owns the validation rule; before/after may reference another worksheet. */
    readonly sheetId: SheetId;
    readonly ownerId: string;
    readonly before: Readonly<RangeRef>;
    readonly after: Readonly<RangeRef>;
    readonly beforeOwnerRanges: readonly Readonly<RangeRef>[];
    readonly afterOwnerRanges: readonly Readonly<RangeRef>[];
  }
  | {
    readonly ownerKind: 'conditional-format' | 'data-validation';
    readonly sheetId: SheetId;
    readonly ownerId: string;
    readonly before: readonly Readonly<RangeRef>[];
    readonly after: readonly Readonly<RangeRef>[];
  };
