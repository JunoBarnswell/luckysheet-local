import type { CellAddress, FormulaReferenceNode } from './ast';

export interface CellDependency {
  readonly kind: 'cell';
  readonly address: CellAddress;
}

export interface RangeDependency {
  readonly kind: 'range';
  readonly start: CellAddress;
  readonly end: CellAddress;
}

export interface StructuralReferenceDependency {
  readonly kind: 'reference';
  readonly reference: FormulaReferenceNode;
}

export interface NameDependency {
  readonly kind: 'name';
  /** Display-independent name token; scope binding happens in FormulaEngine. */
  readonly name: string;
}

export type FormulaDependency = CellDependency | RangeDependency | StructuralReferenceDependency | NameDependency;
