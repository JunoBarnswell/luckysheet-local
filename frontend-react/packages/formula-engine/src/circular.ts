import type { CellAddress } from './ast';
import { cellAddressKey, compareCellAddresses } from './address';
import type { FormulaDependency } from './range-index';

export interface FormulaGraphNode {
  readonly address: CellAddress;
  readonly dependencies: readonly FormulaDependency[];
}

export interface CircularComponent {
  readonly members: readonly CellAddress[];
  readonly cyclic: boolean;
}

/**
 * Deterministically decompose a formula dependency graph into SCCs.
 * Range dependencies are kept as structural edges and are only connected to
 * formula cells whose addresses are inside the range; the graph never
 * materializes blank cells.
 */
export function findFormulaComponents(nodes: readonly FormulaGraphNode[]): readonly CircularComponent[] {
  const ordered = [...nodes].sort((left, right) => compareCellAddresses(left.address, right.address));
  const byKey = new Map(ordered.map((node) => [cellAddressKey(node.address), node]));
  const adjacency = new Map<string, readonly string[]>();
  for (const node of ordered) {
    const targets = new Set<string>();
    for (const dependency of node.dependencies) {
      if (dependency.kind === 'cell') {
        const key = cellAddressKey(dependency.address);
        if (byKey.has(key)) targets.add(key);
      } else if (dependency.kind === 'range') {
        for (const candidate of ordered) {
          if (candidate.address.sheetId !== dependency.start.sheetId) continue;
          if (candidate.address.row < dependency.start.row || candidate.address.row > dependency.end.row) continue;
          if (candidate.address.column < dependency.start.column || candidate.address.column > dependency.end.column) continue;
          targets.add(cellAddressKey(candidate.address));
        }
      }
    }
    adjacency.set(cellAddressKey(node.address), [...targets].sort());
  }

  let sequence = 0;
  const index = new Map<string, number>();
  const lowLink = new Map<string, number>();
  const stack: string[] = [];
  const onStack = new Set<string>();
  const components: CircularComponent[] = [];

  const visit = (key: string): void => {
    index.set(key, sequence);
    lowLink.set(key, sequence);
    sequence += 1;
    stack.push(key);
    onStack.add(key);

    for (const target of adjacency.get(key) ?? []) {
      if (!index.has(target)) {
        visit(target);
        lowLink.set(key, Math.min(lowLink.get(key)!, lowLink.get(target)!));
      } else if (onStack.has(target)) {
        lowLink.set(key, Math.min(lowLink.get(key)!, index.get(target)!));
      }
    }

    if (lowLink.get(key) !== index.get(key)) return;
    const memberKeys: string[] = [];
    let member: string | undefined;
    do {
      member = stack.pop();
      if (member === undefined) break;
      onStack.delete(member);
      memberKeys.push(member);
    } while (member !== key);
    const members = memberKeys
      .map((memberKey) => byKey.get(memberKey)!.address)
      .sort(compareCellAddresses)
      .map((address) => ({ ...address }));
    const only = cellAddressKey(members[0]!);
    const cyclic = members.length > 1 || (adjacency.get(only) ?? []).includes(only);
    components.push({ members, cyclic });
  };

  for (const node of ordered) {
    const key = cellAddressKey(node.address);
    if (!index.has(key)) visit(key);
  }

  return components.sort((left, right) => compareCellAddresses(left.members[0]!, right.members[0]!));
}
