import type { CellAddress } from './ast';
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

export type FormulaDependency = CellDependency | RangeDependency;

interface IndexEntry {
  readonly owner: CellAddress;
  readonly dependencies: readonly FormulaDependency[];
}

export class RangeIndex {
  private readonly entries = new Map<string, IndexEntry>();
  private readonly cellDependents = new Map<string, Set<string>>();
  private readonly rangeDependencies = new Map<string, readonly RangeDependency[]>();

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
      } else {
        ranges.push(dependency);
      }
    }
    if (ranges.length > 0) this.rangeDependencies.set(ownerKey, ranges);
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
    return true;
  }

  getDependencies(owner: CellAddress): readonly FormulaDependency[] {
    const entry = this.entries.get(cellAddressKey(owner));
    return entry?.dependencies.map(copyDependency) ?? [];
  }

  getDependents(address: CellAddress): readonly CellAddress[] {
    const addressKey = cellAddressKey(address);
    const dependentKeys = new Set<string>(this.cellDependents.get(addressKey) ?? []);

    for (const [ownerKey, ranges] of this.rangeDependencies) {
      if (ranges.some((range) => containsAddress(range, address))) dependentKeys.add(ownerKey);
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
  }

  get size(): number {
    return this.entries.size;
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

function deduplicateDependencies(dependencies: readonly FormulaDependency[]): readonly FormulaDependency[] {
  const seen = new Set<string>();
  const result: FormulaDependency[] = [];
  for (const dependency of dependencies) {
    const normalized = dependency.kind === 'cell'
      ? { kind: 'cell' as const, address: copyAddress(assertAndReturn(dependency.address)) }
      : normalizeRange(dependency.start, dependency.end);
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
    : `range:${cellAddressKey(dependency.start)}:${cellAddressKey(dependency.end)}`;
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
    : { kind: 'range', start: copyAddress(dependency.start), end: copyAddress(dependency.end) };
}
