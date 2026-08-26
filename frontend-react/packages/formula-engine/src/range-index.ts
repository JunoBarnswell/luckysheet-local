import type { CellAddress, FormulaReferenceNode } from './ast';
import { assertCellAddress, cellAddressKey, compareCellAddresses } from './address';
import { FormulaReferenceError } from './errors';

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

export type FormulaDependency = CellDependency | RangeDependency | StructuralReferenceDependency;

interface IndexEntry {
  readonly owner: CellAddress;
  readonly dependencies: readonly FormulaDependency[];
}

interface RangeEntry { readonly ownerKey: string; readonly range: RangeDependency; }
interface RangeTreeNode {
  readonly center: number;
  readonly crossing: readonly RangeEntry[];
  readonly left?: RangeTreeNode;
  readonly right?: RangeTreeNode;
}

export class RangeIndex {
  private readonly entries = new Map<string, IndexEntry>();
  private readonly cellDependents = new Map<string, Set<string>>();
  private readonly rangeDependencies = new Map<string, readonly RangeDependency[]>();
  private readonly rangeTrees = new Map<string, RangeTreeNode | undefined>();

  set(owner: CellAddress, dependencies: readonly FormulaDependency[]): void {
    assertCellAddress(owner);
    this.remove(owner);

    const normalizedDependencies = deduplicateDependencies(dependencies);
    const ownerKey = cellAddressKey(owner);
    this.entries.set(ownerKey, { owner: copyAddress(owner), dependencies: normalizedDependencies });

    const ranges: RangeDependency[] = [];
    for (const dependency of normalizedDependencies) {
      if (dependency.kind === 'cell') {
        const dependencyKey = cellAddressKey(dependency.address);
        let dependents = this.cellDependents.get(dependencyKey);
        if (!dependents) {
          dependents = new Set<string>();
          this.cellDependents.set(dependencyKey, dependents);
        }
        dependents.add(ownerKey);
      } else if (dependency.kind === 'range') {
        ranges.push(dependency);
      }
    }
    if (ranges.length > 0) this.rangeDependencies.set(ownerKey, ranges);
    this.rebuildRangeTrees();
  }

  add(owner: CellAddress, dependencies: readonly FormulaDependency[]): void {
    this.set(owner, dependencies);
  }

  remove(owner: CellAddress): boolean {
    const ownerKey = cellAddressKey(owner);
    const entry = this.entries.get(ownerKey);
    if (!entry) return false;

    for (const dependency of entry.dependencies) {
      if (dependency.kind !== 'cell') continue;
      const dependencyKey = cellAddressKey(dependency.address);
      const dependents = this.cellDependents.get(dependencyKey);
      dependents?.delete(ownerKey);
      if (dependents?.size === 0) this.cellDependents.delete(dependencyKey);
    }
    this.entries.delete(ownerKey);
    this.rangeDependencies.delete(ownerKey);
    this.rebuildRangeTrees();
    return true;
  }

  getDependencies(owner: CellAddress): readonly FormulaDependency[] {
    const entry = this.entries.get(cellAddressKey(owner));
    return entry?.dependencies.map(copyDependency) ?? [];
  }

  getDependents(address: CellAddress): readonly CellAddress[] {
    const addressKey = cellAddressKey(address);
    const dependentKeys = new Set<string>(this.cellDependents.get(addressKey) ?? []);

    const tree = this.rangeTrees.get(address.sheetId);
    for (const ownerKey of queryRangeTree(tree, address)) dependentKeys.add(ownerKey);
    for (const entry of this.entries.values()) {
      if (entry.dependencies.some((dependency) => dependency.kind === 'reference' && referenceContainsAddress(dependency.reference, address))) {
        dependentKeys.add(cellAddressKey(entry.owner));
      }
    }

    const dependents: CellAddress[] = [];
    for (const dependentKey of dependentKeys) {
      const entry = this.entries.get(dependentKey);
      if (entry) dependents.push(copyAddress(entry.owner));
    }
    dependents.sort(compareCellAddresses);
    return dependents;
  }

  clear(): void {
    this.entries.clear();
    this.cellDependents.clear();
    this.rangeDependencies.clear();
    this.rangeTrees.clear();
  }

  get size(): number {
    return this.entries.size;
  }

  private rebuildRangeTrees(): void {
    const bySheet = new Map<string, RangeEntry[]>();
    for (const [ownerKey, ranges] of this.rangeDependencies) {
      for (const range of ranges) {
        const entries = bySheet.get(range.start.sheetId) ?? [];
        entries.push({ ownerKey, range });
        bySheet.set(range.start.sheetId, entries);
      }
    }
    this.rangeTrees.clear();
    for (const [sheetId, entries] of bySheet) this.rangeTrees.set(sheetId, buildRangeTree(entries));
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

function containsAddress(range: RangeDependency, address: CellAddress): boolean {
  return (
    range.start.sheetId === address.sheetId &&
    address.row >= range.start.row &&
    address.row <= range.end.row &&
    address.column >= range.start.column &&
    address.column <= range.end.column
  );
}

function buildRangeTree(entries: readonly RangeEntry[]): RangeTreeNode | undefined {
  if (entries.length === 0) return undefined;
  const coordinates = entries.flatMap((entry) => [entry.range.start.row, entry.range.end.row]).sort((left, right) => left - right);
  const center = coordinates[Math.floor(coordinates.length / 2)]!;
  const left: RangeEntry[] = [];
  const right: RangeEntry[] = [];
  const crossing: RangeEntry[] = [];
  for (const entry of entries) {
    if (entry.range.end.row < center) left.push(entry);
    else if (entry.range.start.row > center) right.push(entry);
    else crossing.push(entry);
  }
  return {
    center,
    crossing: crossing.sort((a, b) => a.range.start.row - b.range.start.row),
    left: buildRangeTree(left),
    right: buildRangeTree(right),
  };
}

function queryRangeTree(node: RangeTreeNode | undefined, address: CellAddress): Set<string> {
  const result = new Set<string>();
  if (!node) return result;
  if (address.row < node.center) {
    for (const entry of node.crossing) {
      if (entry.range.start.row > address.row) break;
      if (containsAddress(entry.range, address)) result.add(entry.ownerKey);
    }
    for (const ownerKey of queryRangeTree(node.left, address)) result.add(ownerKey);
    return result;
  }
  if (address.row > node.center) {
    for (const entry of node.crossing) {
      if (entry.range.end.row < address.row) continue;
      if (containsAddress(entry.range, address)) result.add(entry.ownerKey);
    }
    for (const ownerKey of queryRangeTree(node.right, address)) result.add(ownerKey);
    return result;
  }
  for (const entry of node.crossing) if (containsAddress(entry.range, address)) result.add(entry.ownerKey);
  return result;
}

function deduplicateDependencies(dependencies: readonly FormulaDependency[]): readonly FormulaDependency[] {
  const seen = new Set<string>();
  const result: FormulaDependency[] = [];
  for (const dependency of dependencies) {
    const normalized = dependency.kind === 'cell'
      ? { kind: 'cell' as const, address: copyAddress(assertAndReturn(dependency.address)) }
      : dependency.kind === 'range'
        ? normalizeRange(dependency.start, dependency.end)
        : { kind: 'reference' as const, reference: dependency.reference };
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
      : `reference:${JSON.stringify(dependency.reference)}`;
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
      : { kind: 'reference', reference: structuredClone(dependency.reference) };
}

function referenceContainsAddress(reference: FormulaReferenceNode, address: CellAddress): boolean {
  switch (reference.type) {
    case 'cell-reference':
      return (reference.reference.sheetId ?? address.sheetId) === address.sheetId
        && reference.reference.row === address.row
        && reference.reference.column === address.column;
    case 'range-reference':
      return referenceContainsAddress(reference.start, address) || referenceContainsAddress(reference.end, address)
        || (reference.start.reference.sheetId ?? address.sheetId) === address.sheetId
          && address.row >= Math.min(reference.start.reference.row, reference.end.reference.row)
          && address.row <= Math.max(reference.start.reference.row, reference.end.reference.row)
          && address.column >= Math.min(reference.start.reference.column, reference.end.reference.column)
          && address.column <= Math.max(reference.start.reference.column, reference.end.reference.column);
    case 'whole-column-reference':
      return (reference.sheetId ?? address.sheetId) === address.sheetId
        && address.column >= reference.startColumn
        && address.column <= reference.endColumn;
    case 'whole-row-reference':
      return (reference.sheetId ?? address.sheetId) === address.sheetId
        && address.row >= reference.startRow
        && address.row <= reference.endRow;
    case 'reference-union':
      return reference.references.some((entry) => referenceContainsAddress(entry, address));
    case 'reference-intersection':
      return referenceContainsAddress(reference.left, address) && referenceContainsAddress(reference.right, address);
    case 'sheet-range-reference':
      return reference.qualifier.startSheetId === address.sheetId || reference.qualifier.endSheetId === address.sheetId
        ? referenceContainsAddress(reference.reference, address)
        : false;
    case 'external-reference':
      return false;
    case 'spill-reference':
      return referenceContainsAddress(reference.operand as FormulaReferenceNode, address);
    case 'table-reference':
      return false;
    case 'invalid-reference':
      return false;
  }
}
