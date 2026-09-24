import type { CellAddress, FormulaReferenceNode } from './ast';
import { assertCellAddress, cellAddressKey, compareCellAddresses } from './address';
import { FormulaReferenceError } from './errors';
import { ReferenceIndex, type IndexedReferenceOwnerSource } from './reference-index';
import type { FormulaSheetIdentity } from './sheet-reference';

export type { FormulaSheetIdentity } from './sheet-reference';

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

interface IndexEntry {
  readonly dependencies: readonly FormulaDependency[];
}

export type StructuralReferenceOwnerSource = IndexedReferenceOwnerSource;

export class RangeIndex {
  private readonly entries = new Map<string, IndexEntry>();
  private readonly invalidFormulaOwners = new Map<string, CellAddress>();
  private readonly invalidStructuralFormulaOwners = new Map<string, CellAddress>();
  private readonly referenceIndex: ReferenceIndex;

  constructor(sheetOrder: readonly FormulaSheetIdentity[] = []) {
    this.referenceIndex = new ReferenceIndex(sheetOrder);
  }

  setSheetOrder(sheetOrder: readonly FormulaSheetIdentity[]): void {
    this.referenceIndex.setSheetOrder(sheetOrder);
  }

  set(owner: CellAddress, dependencies: readonly FormulaDependency[], invalidFormula = false): void {
    assertCellAddress(owner);
    const normalizedDependencies = deduplicateDependencies(dependencies);
    const ownerKey = cellAddressKey(owner);
    this.referenceIndex.set(owner, normalizedDependencies);
    this.entries.set(ownerKey, { dependencies: normalizedDependencies });
    if (invalidFormula) this.invalidFormulaOwners.set(ownerKey, copyAddress(owner));
    else this.invalidFormulaOwners.delete(ownerKey);
  }

  setStructuralReference(
    owner: CellAddress,
    sourceId: string,
    dependencies: readonly FormulaDependency[],
    invalidFormula = false,
  ): void {
    assertCellAddress(owner);
    assertStructuralSourceId(sourceId);
    const normalizedDependencies = deduplicateDependencies(dependencies);
    const ownerKey = cellAddressKey(owner);
    const sourceKey = `${ownerKey}\u0000${sourceId}`;
    this.referenceIndex.set(owner, normalizedDependencies, sourceId);
    if (invalidFormula) this.invalidStructuralFormulaOwners.set(sourceKey, copyAddress(owner));
    else this.invalidStructuralFormulaOwners.delete(sourceKey);
  }

  removeStructuralReference(owner: CellAddress, sourceId: string): boolean {
    assertCellAddress(owner);
    assertStructuralSourceId(sourceId);
    const sourceKey = `${cellAddressKey(owner)}\u0000${sourceId}`;
    const removed = this.referenceIndex.remove(owner, sourceId);
    const invalidRemoved = this.invalidStructuralFormulaOwners.delete(sourceKey);
    if (invalidRemoved && !removed) {
      throw new Error('REFERENCE_INDEX_INVARIANT: invalid structural formula owner is missing from reference postings');
    }
    return removed;
  }

  getStructuralReferenceOwnersInRange(
    sheetId: string,
    range: { readonly startRow: number; readonly endRow: number; readonly startColumn: number; readonly endColumn: number },
  ): readonly StructuralReferenceOwnerSource[] {
    return this.referenceIndex.getOwnersInRange(sheetId, range)
      .filter(({ sourceId }) => sourceId.startsWith('structural:'));
  }

  add(owner: CellAddress, dependencies: readonly FormulaDependency[]): void {
    this.set(owner, dependencies);
  }

  remove(owner: CellAddress): boolean {
    const ownerKey = cellAddressKey(owner);
    const entry = this.entries.get(ownerKey);
    if (!entry) return false;
    if (!this.referenceIndex.remove(owner)) {
      throw new Error('REFERENCE_INDEX_INVARIANT: formula owner is missing from reference postings');
    }
    this.entries.delete(ownerKey);
    this.invalidFormulaOwners.delete(ownerKey);
    return true;
  }

  getDependencies(owner: CellAddress): readonly FormulaDependency[] {
    const entry = this.entries.get(cellAddressKey(owner));
    return entry?.dependencies.map(copyDependency) ?? [];
  }

  getDependents(address: CellAddress): readonly CellAddress[] {
    return this.referenceIndex.getDependents(address);
  }

  clear(): void {
    this.entries.clear();
    this.invalidFormulaOwners.clear();
    this.invalidStructuralFormulaOwners.clear();
    this.referenceIndex.clear();
  }

  get size(): number {
    return this.entries.size;
  }

  getStructuralDependents(sheetId: string, axis: 'row' | 'column', at: number): readonly CellAddress[] {
    return this.referenceIndex.getStructuralDependents(sheetId, axis, at);
  }

  getRangeDependents(
    sheetId: string,
    range: { readonly startRow: number; readonly endRow: number; readonly startColumn: number; readonly endColumn: number },
  ): readonly CellAddress[] {
    return this.referenceIndex.getRangeDependents(sheetId, range);
  }

  getInvalidFormulaOwners(): readonly CellAddress[] {
    const owners = new Map<string, CellAddress>();
    for (const owner of this.invalidFormulaOwners.values()) owners.set(cellAddressKey(owner), owner);
    for (const owner of this.invalidStructuralFormulaOwners.values()) owners.set(cellAddressKey(owner), owner);
    return [...owners.values()].map(copyAddress).sort(compareCellAddresses);
  }
}

function assertStructuralSourceId(sourceId: string): void {
  if (typeof sourceId !== 'string' || !sourceId.startsWith('structural:') || sourceId.length === 'structural:'.length) {
    throw new FormulaReferenceError('Structural reference owner source identity is invalid');
  }
}

export function normalizeRange(start: CellAddress, end: CellAddress): RangeDependency {
  assertCellAddress(start);
  assertCellAddress(end);
  if (start.sheetId !== end.sheetId) {
    throw new FormulaReferenceError('A range cannot cross worksheets');
  }
  return {
    kind: 'range',
    start: {
      sheetId: start.sheetId,
      row: Math.min(start.row, end.row),
      column: Math.min(start.column, end.column),
    },
    end: {
      sheetId: start.sheetId,
      row: Math.max(start.row, end.row),
      column: Math.max(start.column, end.column),
    },
  };
}

function deduplicateDependencies(dependencies: readonly FormulaDependency[]): readonly FormulaDependency[] {
  const seen = new Set<string>();
  const result: FormulaDependency[] = [];
  for (const dependency of dependencies) {
    const normalized = dependency.kind === 'cell'
      ? { kind: 'cell' as const, address: copyAddress(assertAndReturn(dependency.address)) }
      : dependency.kind === 'range'
        ? normalizeRange(dependency.start, dependency.end)
        : dependency.kind === 'reference'
        ? { kind: 'reference' as const, reference: structuredClone(dependency.reference) }
          : { kind: 'name' as const, name: dependency.name };
    const key = dependencyKey(normalized);
    if (seen.has(key)) continue;
    seen.add(key);
    result.push(normalized);
  }
  return result;
}

function dependencyKey(dependency: FormulaDependency): string {
  return dependency.kind === 'cell'
    ? `cell:${cellAddressKey(dependency.address)}`
    : dependency.kind === 'range'
      ? `range:${cellAddressKey(dependency.start)}:${cellAddressKey(dependency.end)}`
      : dependency.kind === 'reference'
        ? `reference:${JSON.stringify(dependency.reference)}`
        : `name:${dependency.name}`;
}

function assertAndReturn(address: CellAddress): CellAddress {
  assertCellAddress(address);
  return address;
}

function copyAddress(address: CellAddress): CellAddress {
  return { sheetId: address.sheetId, row: address.row, column: address.column };
}

function copyDependency(dependency: FormulaDependency): FormulaDependency {
  return dependency.kind === 'cell'
    ? { kind: 'cell', address: copyAddress(dependency.address) }
    : dependency.kind === 'range'
      ? { kind: 'range', start: copyAddress(dependency.start), end: copyAddress(dependency.end) }
      : dependency.kind === 'reference'
        ? { kind: 'reference', reference: structuredClone(dependency.reference) }
        : { kind: 'name', name: dependency.name };
}
