import type { CellShiftSpec, StructuralTransformParams } from './domain';

/** UI intent only; the kernel owns every workbook mutation and reference rewrite. */
export type StructuralTransformIntent = Readonly<StructuralTransformParams>;

/** Serializable cell-shift intent. Validation and execution belong to the kernel. */
export type CellShiftIntent = Readonly<CellShiftSpec>;
