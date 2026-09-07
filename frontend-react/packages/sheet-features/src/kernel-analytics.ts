import type { RangeRef, WorksheetModel } from '@react-sheets/core-model';
import { kernelInvoke } from '@react-sheets/kernel-client';

export interface ResolvedVisibility {
  readonly revision: number;
  readonly rows: ReadonlyMap<number, { manualHidden: boolean; filterHidden: boolean; outlineHidden: boolean }>;
  readonly columns: ReadonlyMap<number, { manualHidden: boolean }>;
  isRowHidden(row: number): boolean;
  isColumnHidden(column: number): boolean;
}

export interface KernelFilterOwner { id: string; range: RangeRef; columns: Array<{ column: number; predicate: unknown }> }
export interface KernelFilterRequest {
  revision: number;
  range: RangeRef;
  owners?: KernelFilterOwner[];
  conditions?: Array<{ column: number; predicate: unknown }>;
  sort?: Array<{ column: number; descending?: boolean }>;
  offset?: number;
  limit?: number;
}
export interface KernelFilterResult {
  revision: number;
  ownerId?: string;
  visibleRows: number[];
  totalRows: number;
  visibility: { rows: number; hidden: Array<number | string>; reasons: Record<string, Array<number | string>> };
  domain: Record<string, unknown>;
}

export function executeKernelFilter(unitId: string, request: KernelFilterRequest): KernelFilterResult {
  const response = kernelInvoke<KernelFilterResult & { kind: 'filter' }>('analytics.execute', {
    unitId,
    revision: request.revision,
    request: { kind: 'filter', params: request },
  });
  if (response.kind !== 'filter' || response.revision !== request.revision || !response.visibility) {
    throw new Error('KERNEL_ANALYTICS_RESPONSE_INVALID: filter response revision or kind is invalid');
  }
  return response;
}

function reasonSet(bits: readonly (number | string)[] | undefined): Set<number> {
  const rows = new Set<number>();
  for (let wordIndex = 0; wordIndex < (bits?.length ?? 0); wordIndex += 1) {
    let word = BigInt(bits![wordIndex] ?? 0);
    while (word !== 0n) {
      const bit = word & -word;
      let offset = 0;
      for (let probe = bit; probe > 1n; probe >>= 1n) offset += 1;
      rows.add(wordIndex * 64 + offset);
      word ^= bit;
    }
  }
  return rows;
}

/** Convert the kernel bitmap into the single projection consumed by formula, chart and print. */
export function resolvedVisibilityFromKernel(
  sheet: WorksheetModel,
  result: KernelFilterResult,
  rangeStartRow = 0,
  outlineRows: ReadonlySet<number> = new Set(),
): ResolvedVisibility {
  const rows = new Map<number, { manualHidden: boolean; filterHidden: boolean; outlineHidden: boolean }>();
  const filterRows = reasonSet(result.visibility.reasons.Filter ?? result.visibility.reasons.filter);
  const manualRows = reasonSet(result.visibility.reasons.ManualHidden ?? result.visibility.reasons.manualHidden);
  const outline = reasonSet(result.visibility.reasons.OutlineHidden ?? result.visibility.reasons.outlineHidden);
  const hiddenRows = reasonSet(result.visibility.hidden);
  const candidates = new Set([...filterRows, ...manualRows, ...outline, ...hiddenRows, ...[...sheet.hiddenRows].map((row) => row - rangeStartRow)]);
  for (const row of candidates) {
    const absoluteRow = row + rangeStartRow;
    const manualHidden = sheet.hiddenRows.has(absoluteRow) || manualRows.has(row);
    const filterHidden = filterRows.has(row) || hiddenRows.has(row) && !manualHidden && !outline.has(row);
    const outlineHidden = outlineRows.has(absoluteRow) || outline.has(row);
    if (manualHidden || filterHidden || outlineHidden) rows.set(absoluteRow, { manualHidden, filterHidden, outlineHidden });
  }
  const columns = new Map<number, { manualHidden: boolean }>();
  for (const column of sheet.hiddenColumns) columns.set(column, { manualHidden: true });
  return {
    revision: result.revision,
    rows,
    columns,
    isRowHidden: (row) => rows.get(row) !== undefined,
    isColumnHidden: (column) => columns.get(column) !== undefined,
  };
}
