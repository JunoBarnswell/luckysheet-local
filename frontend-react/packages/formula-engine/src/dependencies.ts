import type { CellAddress, CellReferenceNode, FormulaAst, RangeReferenceNode, ParsedCellReference } from './ast';
import { assertCellAddress, cellAddressKey } from './address';
import { FormulaReferenceError } from './errors';
import { normalizeRange, type CellDependency, type FormulaDependency, type RangeDependency } from './range-index';

export function collectFormulaDependencies(ast: FormulaAst, owner: CellAddress): readonly FormulaDependency[] {
  assertCellAddress(owner);
  const dependencies: FormulaDependency[] = [];
  const seen = new Set<string>();
  visit(ast, owner, dependencies, seen);
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
      visit(node.operand, owner, dependencies, seen);
      return;
    case 'binary-expression':
      visit(node.left, owner, dependencies, seen);
      visit(node.right, owner, dependencies, seen);
      return;
    case 'function-call':
      for (const argument of node.arguments) visit(argument, owner, dependencies, seen);
      return;
    case 'number-literal':
    case 'string-literal':
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
