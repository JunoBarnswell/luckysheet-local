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
  };
