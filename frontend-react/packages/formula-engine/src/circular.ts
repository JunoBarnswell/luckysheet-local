import type { CellAddress, FormulaReferenceNode } from './ast';
import { cellAddressKey, compareCellAddresses } from './address';
import { FormulaReferenceError } from './errors';
import type { FormulaDependency, RangeDependency } from './range-index';
import { resolveFormulaSheetId, type FormulaSheetIdentity } from './sheet-reference';

export interface FormulaGraphNode {
  readonly address: CellAddress;
  readonly dependencies: readonly FormulaDependency[];
}

export interface CircularComponent {
  readonly members: readonly CellAddress[];
  readonly cyclic: boolean;
}

export interface FormulaGraphAnalysis {
  readonly components: readonly CircularComponent[];
  readonly calculationOrder: readonly string[];
  readonly prerequisites: ReadonlyMap<string, readonly string[]>;
}

/**
 * Decompose formula dependencies without expanding blank cells or recursing
 * through the JavaScript call stack. Range edges query a row-sorted formula
 * address index, so a sparse range visits formulas inside the range rather
 * than every formula in the workbook.
 */
export function findFormulaComponents(
  nodes: readonly FormulaGraphNode[],
  sheetOrder: readonly FormulaSheetIdentity[] = [],
): readonly CircularComponent[] {
  return analyzeFormulaGraph(nodes, sheetOrder).components;
}

export function analyzeFormulaGraph(
  nodes: readonly FormulaGraphNode[],
  sheetOrder: readonly FormulaSheetIdentity[] = [],
): FormulaGraphAnalysis {
  const ordered = [...nodes].sort((left, right) => compareCellAddresses(left.address, right.address));
  const byKey = new Map(ordered.map((node) => [cellAddressKey(node.address), node]));
  const formulasBySheet = new Map<string, CellAddress[]>();
  for (const node of ordered) {
    const addresses = formulasBySheet.get(node.address.sheetId) ?? [];
    addresses.push(node.address);
    formulasBySheet.set(node.address.sheetId, addresses);
  }

  const adjacency = new Map<string, readonly string[]>();
  const reverse = new Map<string, Set<string>>();
  for (const node of ordered) reverse.set(cellAddressKey(node.address), new Set());
  for (const node of ordered) {
    const ownerKey = cellAddressKey(node.address);
    const targets = new Set<string>();
    for (const dependency of node.dependencies) {
      if (dependency.kind === 'cell') {
        const key = cellAddressKey(dependency.address);
        if (byKey.has(key)) targets.add(key);
      } else if (dependency.kind === 'range') {
        addRangeTargets(dependency, formulasBySheet, targets);
      } else if (dependency.kind === 'reference') {
        for (const range of referenceRanges(dependency.reference, node.address, sheetOrder)) {
          addRangeTargets(range, formulasBySheet, targets);
        }
      }
    }
    const sortedTargets = [...targets].sort();
    adjacency.set(ownerKey, sortedTargets);
    for (const target of sortedTargets) reverse.get(target)?.add(ownerKey);
  }

  const finished: string[] = [];
  const visited = new Set<string>();
  for (const node of ordered) {
    const root = cellAddressKey(node.address);
    if (visited.has(root)) continue;
    visited.add(root);
    const stack: Array<{ key: string; next: number }> = [{ key: root, next: 0 }];
    while (stack.length > 0) {
      const frame = stack[stack.length - 1]!;
      const neighbors = adjacency.get(frame.key) ?? [];
      if (frame.next < neighbors.length) {
        const target = neighbors[frame.next++]!;
        if (!visited.has(target)) {
          visited.add(target);
          stack.push({ key: target, next: 0 });
        }
      } else {
        finished.push(frame.key);
        stack.pop();
      }
    }
  }

  const assigned = new Set<string>();
  const components: CircularComponent[] = [];
  for (let index = finished.length - 1; index >= 0; index -= 1) {
    const root = finished[index]!;
    if (assigned.has(root)) continue;
    assigned.add(root);
    const members: CellAddress[] = [];
    const stack = [root];
    while (stack.length > 0) {
      const key = stack.pop()!;
      const node = byKey.get(key);
      if (node) members.push({ ...node.address });
      for (const predecessor of reverse.get(key) ?? []) {
        if (!assigned.has(predecessor)) {
          assigned.add(predecessor);
          stack.push(predecessor);
        }
      }
    }
    members.sort(compareCellAddresses);
    const onlyKey = cellAddressKey(members[0]!);
    const cyclic = members.length > 1 || (adjacency.get(onlyKey) ?? []).includes(onlyKey);
    components.push({ members, cyclic });
  }

  return {
    components: components.sort((left, right) => compareCellAddresses(left.members[0]!, right.members[0]!)),
    calculationOrder: finished,
    prerequisites: adjacency,
  };
}

function addRangeTargets(
  range: RangeDependency,
  formulasBySheet: ReadonlyMap<string, readonly CellAddress[]>,
  targets: Set<string>,
): void {
  const addresses = formulasBySheet.get(range.start.sheetId);
  if (!addresses) return;
  let low = 0;
  let high = addresses.length;
  while (low < high) {
    const middle = (low + high) >>> 1;
    if (addresses[middle]!.row < range.start.row) low = middle + 1;
    else high = middle;
  }
  for (let index = low; index < addresses.length; index += 1) {
    const address = addresses[index]!;
    if (address.row > range.end.row) break;
    if (address.column >= range.start.column && address.column <= range.end.column) {
      targets.add(cellAddressKey(address));
    }
  }
}

function referenceRanges(
  reference: FormulaReferenceNode,
  owner: CellAddress,
  sheetOrder: readonly FormulaSheetIdentity[],
): RangeDependency[] {
  switch (reference.type) {
    case 'cell-reference': {
      const sheetId = resolveGraphSheetId(reference.reference.sheetId, owner, sheetOrder);
      if (!sheetId) return [];
      const address = { ...reference.reference, sheetId };
      return [{ kind: 'range', start: address, end: { ...address } }];
    }
    case 'range-reference': {
      const startSheetId = resolveGraphSheetId(reference.start.reference.sheetId, owner, sheetOrder);
      if (!startSheetId) return [];
      const endReference = reference.end.reference.sheetId;
      const endSheetId = endReference === undefined
        ? startSheetId
        : resolveGraphSheetId(endReference, owner, sheetOrder);
      if (!endSheetId || startSheetId !== endSheetId) return [];
      return [{
        kind: 'range',
        start: { sheetId: startSheetId, row: Math.min(reference.start.reference.row, reference.end.reference.row), column: Math.min(reference.start.reference.column, reference.end.reference.column) },
        end: { sheetId: startSheetId, row: Math.max(reference.start.reference.row, reference.end.reference.row), column: Math.max(reference.start.reference.column, reference.end.reference.column) },
      }];
    }
    case 'whole-row-reference': {
      const sheetId = resolveGraphSheetId(reference.sheetId, owner, sheetOrder);
      return sheetId
        ? [{ kind: 'range', start: { sheetId, row: reference.startRow, column: 0 }, end: { sheetId, row: reference.endRow, column: Number.MAX_SAFE_INTEGER } }]
        : [];
    }
    case 'whole-column-reference': {
      const sheetId = resolveGraphSheetId(reference.sheetId, owner, sheetOrder);
      return sheetId
        ? [{ kind: 'range', start: { sheetId, row: 0, column: reference.startColumn }, end: { sheetId, row: Number.MAX_SAFE_INTEGER, column: reference.endColumn } }]
        : [];
    }
    case 'sheet-range-reference': {
      const startSheetId = resolveGraphSheetId(reference.qualifier.startSheetId, owner, sheetOrder);
      const endSheetId = resolveGraphSheetId(reference.qualifier.endSheetId, owner, sheetOrder);
      if (!startSheetId || !endSheetId) return [];
      const startIndex = sheetOrder.findIndex((sheet) => sheet.id === startSheetId);
      const endIndex = sheetOrder.findIndex((sheet) => sheet.id === endSheetId);
      if (startIndex < 0 || endIndex < 0) return [];
      const ranges: RangeDependency[] = [];
      for (const sheet of sheetOrder.slice(Math.min(startIndex, endIndex), Math.max(startIndex, endIndex) + 1)) {
        ranges.push(...referenceRanges(reference.reference, { ...owner, sheetId: sheet.id }, sheetOrder));
      }
      return ranges;
    }
    case 'reference-union':
      return reference.references.flatMap((item) => referenceRanges(item, owner, sheetOrder));
    case 'reference-intersection': {
      const intersections: RangeDependency[] = [];
      for (const left of referenceRanges(reference.left, owner, sheetOrder)) {
        for (const right of referenceRanges(reference.right, owner, sheetOrder)) {
          if (left.start.sheetId !== right.start.sheetId) continue;
          const startRow = Math.max(left.start.row, right.start.row);
          const endRow = Math.min(left.end.row, right.end.row);
          const startColumn = Math.max(left.start.column, right.start.column);
          const endColumn = Math.min(left.end.column, right.end.column);
          if (startRow <= endRow && startColumn <= endColumn) intersections.push({
            kind: 'range',
            start: { sheetId: left.start.sheetId, row: startRow, column: startColumn },
            end: { sheetId: left.start.sheetId, row: endRow, column: endColumn },
          });
        }
      }
      return intersections;
    }
    case 'spill-reference':
      return reference.operand.type === 'cell-reference'
        ? referenceRanges(reference.operand, owner, sheetOrder)
        : [];
    case 'external-reference':
    case 'table-reference':
    case 'invalid-reference':
      return [];
  }
}

function resolveGraphSheetId(
  reference: string | undefined,
  owner: CellAddress,
  sheetOrder: readonly FormulaSheetIdentity[],
): string | undefined {
  try {
    return resolveFormulaSheetId(reference, owner.sheetId, sheetOrder);
  } catch (error) {
    if (error instanceof FormulaReferenceError) return undefined;
    throw error;
  }
}
