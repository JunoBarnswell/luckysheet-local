import { MAX_COLUMN_INDEX, MAX_ROW_INDEX } from '@react-sheets/formula-engine';
import type { ReportSheetDefinition } from './data-model';

type ReportBindingCell = ReportSheetDefinition['bindings'][number]['cell'];

/** Map persisted report binding anchors and optional repeated-header rows as structural owners. */
export function mapReportSheetCoordinates(
  definition: ReportSheetDefinition,
  mapCell: (cell: Readonly<ReportBindingCell>) => ReportBindingCell | null,
  mapRepeatedHeaderRow: ((row: number) => number | null) | undefined,
  operation: string,
): ReportSheetDefinition {
  if (!definition || typeof definition !== 'object' || Array.isArray(definition) || !Array.isArray(definition.bindings)) {
    throw new Error(`STRUCTURAL_REFERENCE_OWNER_INVALID: ${operation} report bindings are invalid`);
  }
  const pagination: unknown = definition.pagination;
  if (!pagination || typeof pagination !== 'object' || Array.isArray(pagination)) {
    throw new Error(`STRUCTURAL_REFERENCE_OWNER_INVALID: ${operation} report pagination is invalid`);
  }
  let bindingsChanged = false;
  const bindings = definition.bindings.map((binding, index) => {
    const cell: unknown = binding?.cell;
    if (!cell || typeof cell !== 'object'
      || !Number.isSafeInteger(binding.cell.row) || binding.cell.row < 0 || binding.cell.row > MAX_ROW_INDEX
      || !Number.isSafeInteger(binding.cell.column) || binding.cell.column < 0 || binding.cell.column > MAX_COLUMN_INDEX) {
      throw new Error(`STRUCTURAL_REFERENCE_OWNER_INVALID: ${operation} report binding ${index} has an invalid cell`);
    }
    const mapped = mapCell(binding.cell);
    if (!mapped) {
      throw new Error(`UNSUPPORTED_STRUCTURAL_REFERENCE: ${operation} removes report binding ${index} at ${binding.cell.row}:${binding.cell.column}`);
    }
    if (!Number.isSafeInteger(mapped.row) || mapped.row < 0 || mapped.row > MAX_ROW_INDEX
      || !Number.isSafeInteger(mapped.column) || mapped.column < 0 || mapped.column > MAX_COLUMN_INDEX) {
      throw new Error(`UNSUPPORTED_STRUCTURAL_REFERENCE: ${operation} moves report binding ${index} outside worksheet bounds`);
    }
    if (mapped.row === binding.cell.row && mapped.column === binding.cell.column) return binding;
    bindingsChanged = true;
    return { ...binding, cell: mapped };
  });

  let headersChanged = false;
  const repeatHeaderRowsValue = (pagination as ReportSheetDefinition['pagination']).repeatHeaderRows;
  if (repeatHeaderRowsValue !== undefined) {
    if (!Array.isArray(repeatHeaderRowsValue)) {
      throw new Error(`STRUCTURAL_REFERENCE_OWNER_INVALID: ${operation} repeated header rows are invalid`);
    }
    const repeatHeaderRows: number[] = [];
    for (const row of repeatHeaderRowsValue) {
      if (!Number.isSafeInteger(row) || row < 0 || row > MAX_ROW_INDEX) {
        throw new Error(`STRUCTURAL_REFERENCE_OWNER_INVALID: ${operation} repeated header row is invalid`);
      }
      const mapped = mapRepeatedHeaderRow ? mapRepeatedHeaderRow(row) : row;
      if (mapped === null) {
        headersChanged = true;
        continue;
      }
      if (!Number.isSafeInteger(mapped) || mapped < 0 || mapped > MAX_ROW_INDEX) {
        throw new Error(`UNSUPPORTED_STRUCTURAL_REFERENCE: ${operation} moves a repeated header outside worksheet bounds`);
      }
      repeatHeaderRows.push(mapped);
      headersChanged ||= mapped !== row;
    }
    if (headersChanged) {
      return {
        ...definition,
        bindings: bindingsChanged ? bindings : definition.bindings,
        pagination: { ...definition.pagination, repeatHeaderRows },
      };
    }
  }

  return bindingsChanged
    ? { ...definition, bindings }
    : definition;
}
