import type { MergeSpan } from '@react-sheets/core-model';

interface MergeNode {
  startRow: number;
  endRow: number;
  startColumn: number;
  endColumn: number;
  firstIndex: number;
  entries?: Array<{ merge: MergeSpan; index: number }>;
  left?: MergeNode;
  right?: MergeNode;
}

const LEAF_SIZE = 16;

/** Static point index that preserves the canonical first-match merge ordering. */
export function createMergeSpatialIndex(merges: readonly MergeSpan[]): (row: number, column: number) => MergeSpan | undefined {
  const root = buildNode(merges.map((merge, index) => ({ merge, index })));
  return (row, column) => findMerge(root, row, column)?.merge;
}

function buildNode(entries: Array<{ merge: MergeSpan; index: number }>): MergeNode | undefined {
  if (entries.length === 0) return undefined;
  const bounds = entries.reduce((result, { merge, index }) => ({
    startRow: Math.min(result.startRow, merge.range.startRow),
    endRow: Math.max(result.endRow, merge.range.endRow),
    startColumn: Math.min(result.startColumn, merge.range.startColumn),
    endColumn: Math.max(result.endColumn, merge.range.endColumn),
    firstIndex: Math.min(result.firstIndex, index),
  }), { startRow: Infinity, endRow: -Infinity, startColumn: Infinity, endColumn: -Infinity, firstIndex: Infinity });
  if (entries.length <= LEAF_SIZE) return { ...bounds, entries };

  const splitRows = bounds.endRow - bounds.startRow >= bounds.endColumn - bounds.startColumn;
  entries.sort((first, second) => {
    const firstRange = first.merge.range;
    const secondRange = second.merge.range;
    return splitRows
      ? firstRange.startRow + firstRange.endRow - secondRange.startRow - secondRange.endRow
      : firstRange.startColumn + firstRange.endColumn - secondRange.startColumn - secondRange.endColumn;
  });
  const middle = Math.floor(entries.length / 2);
  return {
    ...bounds,
    left: buildNode(entries.slice(0, middle)),
    right: buildNode(entries.slice(middle)),
  };
}

function findMerge(node: MergeNode | undefined, row: number, column: number, bestIndex = Infinity): { merge: MergeSpan; index: number } | undefined {
  if (!node || node.firstIndex >= bestIndex
    || row < node.startRow || row > node.endRow
    || column < node.startColumn || column > node.endColumn) return undefined;
  if (node.entries) {
    return node.entries.reduce<{ merge: MergeSpan; index: number } | undefined>((best, entry) => {
      const range = entry.merge.range;
      return entry.index < (best?.index ?? bestIndex)
        && row >= range.startRow && row <= range.endRow
        && column >= range.startColumn && column <= range.endColumn
        ? entry
        : best;
    }, undefined);
  }

  const children = [node.left, node.right].filter((child): child is MergeNode => Boolean(child));
  children.sort((first, second) => first.firstIndex - second.firstIndex);
  let best: { merge: MergeSpan; index: number } | undefined;
  for (const child of children) {
    const candidate = findMerge(child, row, column, best?.index ?? bestIndex);
    if (candidate && candidate.index < (best?.index ?? bestIndex)) best = candidate;
  }
  return best;
}
