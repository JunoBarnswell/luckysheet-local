import type { CellAddress, CellReferenceNode, FormulaAst, RangeReferenceNode, ParsedCellReference } from './ast';
import { assertCellAddress, cellAddressKey } from './address';
import { FormulaReferenceError } from './errors';
import { normalizeRange, type CellDependency, type FormulaDependency, type RangeDependency } from './range-index';
import { resolveSheetTableReference, type SheetTableRef } from './sheet-table-resolver';
import { isFormulaError } from './values';

export interface CollectFormulaDependenciesOptions {
  readonly sheetTables?: ReadonlyMap<string, SheetTableRef>;
}

export function collectFormulaDependencies(
  ast: FormulaAst,
  owner: CellAddress,
  options: CollectFormulaDependenciesOptions = {},
): readonly FormulaDependency[] {
  assertCellAddress(owner);
  const dependencies: FormulaDependency[] = [];
  const seen = new Set<string>();
  visit(ast, owner, dependencies, seen, options.sheetTables);
  return dependencies;
}

export function resolveCellReference(reference: ParsedCellReference, currentCell: CellAddress): CellAddress {
  const sheetId = reference.sheetId ?? currentCell.sheetId;
  if (!sheetId) throw new FormulaReferenceError('Cell reference is missing a worksheet');
  return { sheetId, row: reference.row, column: reference.column };
}

export function resolveRangeReference(node: RangeReferenceNode, currentCell: CellAddress): RangeDependency {
  const start = resolveCellReference(node.start.reference, currentCell);
  const end = resolveCellReference(
    node.end.reference,
    node.start.reference.sheetId === undefined ? currentCell : start,
  );
  return normalizeRange(start, end);
}

function visit(
  node: FormulaAst,
  owner: CellAddress,
  dependencies: FormulaDependency[],
  seen: Set<string>,
  sheetTables?: ReadonlyMap<string, SheetTableRef>,
): void {
  switch (node.type) {
    case 'cell-reference': {
      const dependency: CellDependency = { kind: 'cell', address: resolveCellReference(node.reference, owner) };
      addDependency(dependency, dependencies, seen);
      return;
    }
    case 'range-reference': {
      addDependency(resolveRangeReference(node, owner), dependencies, seen);
      return;
    }
    case 'unary-expression':
      visit(node.operand, owner, dependencies, seen, sheetTables);
      return;
    case 'binary-expression':
      visit(node.left, owner, dependencies, seen, sheetTables);
      visit(node.right, owner, dependencies, seen, sheetTables);
      return;
    case 'function-call':
      for (const argument of node.arguments) visit(argument, owner, dependencies, seen, sheetTables);
      return;
    case 'name-reference':
      return;
    case 'invalid-reference':
      // An invalidated reference has no dependency target.  Keeping it out
      // of the index prevents a later write to the deleted coordinate from
      // resurrecting a formula which must remain #REF!.
      return;
    case 'table-reference': {
      if (!sheetTables) return;
      const resolved = resolveSheetTableReference(
        node.tableName,
        {
          specifier: node.specifier,
          columnName: node.columnName,
          thisRow: node.thisRow,
        },
        owner,
        sheetTables,
      );
      if (isFormulaError(resolved)) return;
      if ('start' in resolved && 'end' in resolved) {
        addDependency(resolved, dependencies, seen);
        return;
      }
      const dependency: CellDependency = { kind: 'cell', address: resolved as CellAddress };
      addDependency(dependency, dependencies, seen);
      return;
    }
    case 'number-literal':
    case 'string-literal':
    case 'boolean-literal':
      return;
  }
}

function addDependency(dependency: FormulaDependency, dependencies: FormulaDependency[], seen: Set<string>): void {
  const key = dependency.kind === 'cell'
    ? `cell:${cellAddressKey(dependency.address)}`
    : `range:${cellAddressKey(dependency.start)}:${cellAddressKey(dependency.end)}`;
  if (seen.has(key)) return;
  seen.add(key);
  dependencies.push(dependency);
}
