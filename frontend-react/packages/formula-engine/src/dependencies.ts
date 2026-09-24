import type { CellAddress, CellReferenceNode, FormulaAst, FormulaReferenceNode, RangeReferenceNode, ParsedCellReference } from './ast';
import { assertCellAddress, cellAddressKey } from './address';
import { FormulaReferenceError } from './errors';
import { normalizeRange, type CellDependency, type FormulaDependency, type NameDependency, type RangeDependency } from './range-index';
import { resolveSheetTableReference, type SheetTableRef } from './sheet-table-resolver';
import { resolveFormulaSheetId, type FormulaSheetIdentity } from './sheet-reference';
import { isFormulaError } from './values';

export interface CollectFormulaDependenciesOptions {
  readonly sheetTables?: ReadonlyMap<string, SheetTableRef>;
  readonly sheetOrder?: readonly FormulaSheetIdentity[];
}

export function collectFormulaDependencies(
  ast: FormulaAst,
  owner: CellAddress,
  options: CollectFormulaDependenciesOptions = {},
): readonly FormulaDependency[] {
  assertCellAddress(owner);
  const dependencies: FormulaDependency[] = [];
  const seen = new Set<string>();
  visit(ast, owner, dependencies, seen, options.sheetTables, options.sheetOrder);
  return dependencies;
}

export function resolveCellReference(
  reference: ParsedCellReference,
  currentCell: CellAddress,
  sheetOrder?: readonly FormulaSheetIdentity[],
): CellAddress {
  const sheetId = reference.sheetId ?? currentCell.sheetId;
  if (!sheetId) throw new FormulaReferenceError('Cell reference is missing a worksheet');
  return { sheetId: resolveFormulaSheetId(reference.sheetId, currentCell.sheetId, sheetOrder ?? []), row: reference.row, column: reference.column };
}

export function resolveRangeReference(
  node: RangeReferenceNode,
  currentCell: CellAddress,
  sheetOrder?: readonly FormulaSheetIdentity[],
): RangeDependency {
  const start = resolveCellReference(node.start.reference, currentCell, sheetOrder);
  const end = resolveCellReference(
    node.end.reference,
    node.start.reference.sheetId === undefined ? currentCell : start,
    sheetOrder,
  );
  return normalizeRange(start, end);
}

function visit(
  node: FormulaAst,
  owner: CellAddress,
  dependencies: FormulaDependency[],
  seen: Set<string>,
  sheetTables?: ReadonlyMap<string, SheetTableRef>,
  sheetOrder?: readonly FormulaSheetIdentity[],
): void {
  switch (node.type) {
    case 'cell-reference': {
      const dependency: CellDependency = { kind: 'cell', address: resolveCellReference(node.reference, owner, sheetOrder) };
      addDependency(dependency, dependencies, seen);
      return;
    }
    case 'range-reference': {
      addDependency(resolveRangeReference(node, owner, sheetOrder), dependencies, seen);
      return;
    }
    case 'whole-column-reference':
    case 'whole-row-reference':
    case 'sheet-range-reference':
    case 'external-reference': {
      addDependency({ kind: 'reference', reference: node }, dependencies, seen);
      return;
    }
    case 'reference-union':
    case 'reference-intersection':
      addDependency({ kind: 'reference', reference: node }, dependencies, seen);
      collectNestedTableDependencies(node, owner, dependencies, seen, sheetTables, sheetOrder);
      return;
    case 'unary-expression':
      visit(node.operand, owner, dependencies, seen, sheetTables, sheetOrder);
      return;
    case 'spill-reference':
      visit(node.operand, owner, dependencies, seen, sheetTables, sheetOrder);
      return;
    case 'binary-expression':
      visit(node.left, owner, dependencies, seen, sheetTables, sheetOrder);
      visit(node.right, owner, dependencies, seen, sheetTables, sheetOrder);
      return;
    case 'function-call':
      for (const argument of node.arguments) visit(argument, owner, dependencies, seen, sheetTables, sheetOrder);
      return;
    case 'name-reference': {
      const dependency: NameDependency = { kind: 'name', name: node.name.trim().toUpperCase() };
      addDependency(dependency, dependencies, seen);
      return;
    }
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
          columnEndName: node.columnEndName,
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

function collectNestedTableDependencies(
  reference: FormulaReferenceNode,
  owner: CellAddress,
  dependencies: FormulaDependency[],
  seen: Set<string>,
  sheetTables?: ReadonlyMap<string, SheetTableRef>,
  sheetOrder?: readonly FormulaSheetIdentity[],
): void {
  switch (reference.type) {
    case 'table-reference':
      visit(reference, owner, dependencies, seen, sheetTables, sheetOrder);
      return;
    case 'reference-union':
      for (const item of reference.references) collectNestedTableDependencies(item, owner, dependencies, seen, sheetTables, sheetOrder);
      return;
    case 'reference-intersection':
      collectNestedTableDependencies(reference.left, owner, dependencies, seen, sheetTables, sheetOrder);
      collectNestedTableDependencies(reference.right, owner, dependencies, seen, sheetTables, sheetOrder);
      return;
    case 'spill-reference':
      if (isFormulaReference(reference.operand)) {
        collectNestedTableDependencies(reference.operand, owner, dependencies, seen, sheetTables, sheetOrder);
      }
      return;
    default:
      return;
  }
}

function isFormulaReference(node: FormulaAst): node is FormulaReferenceNode {
  switch (node.type) {
    case 'cell-reference':
    case 'invalid-reference':
    case 'range-reference':
    case 'whole-column-reference':
    case 'whole-row-reference':
    case 'spill-reference':
    case 'table-reference':
    case 'reference-union':
    case 'reference-intersection':
    case 'sheet-range-reference':
    case 'external-reference':
      return true;
    default:
      return false;
  }
}

function addDependency(dependency: FormulaDependency, dependencies: FormulaDependency[], seen: Set<string>): void {
  const key = dependency.kind === 'cell'
    ? `cell:${cellAddressKey(dependency.address)}`
    : dependency.kind === 'range'
      ? `range:${cellAddressKey(dependency.start)}:${cellAddressKey(dependency.end)}`
      : dependency.kind === 'reference'
        ? `reference:${JSON.stringify(dependency.reference)}`
        : `name:${dependency.name}`;
  if (seen.has(key)) return;
  seen.add(key);
  dependencies.push(dependency);
}
