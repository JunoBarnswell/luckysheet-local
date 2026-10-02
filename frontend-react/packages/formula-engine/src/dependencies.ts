import type { CellAddress, CellReferenceNode, FormulaAst, FormulaReferenceNode, RangeReferenceNode, ParsedCellReference } from './ast';
import { assertCellAddress, cellAddressKey } from './address';
import { FormulaReferenceError } from './errors';
import { normalizeRange, type CellDependency, type FormulaDependency, type NameDependency, type RangeDependency } from './range-index';
import { resolveSheetTableReference, type SheetTableRef } from './sheet-table-resolver';
import { resolveFormulaSheetId, type FormulaSheetIdentity } from './sheet-reference';
import { isFormulaError } from './values';
import { getFunctionDescriptor } from './functions';
import { visitLexicalArguments } from './lexical-scope';

export interface CollectFormulaDependenciesOptions {
  readonly sheetTables?: ReadonlyMap<string, SheetTableRef>;
  readonly sheetOrder: readonly FormulaSheetIdentity[];
  readonly resolveNameAst?: (name: string, owner: CellAddress) => FormulaAst | undefined;
  /** Geometry owners remain indexed for structure, but do not create value cycles. */
  readonly valueDependencies?: boolean;
}

export function collectFormulaDependencies(
  ast: FormulaAst,
  owner: CellAddress,
  options: CollectFormulaDependenciesOptions,
): readonly FormulaDependency[] {
  assertCellAddress(owner);
  const dependencies: FormulaDependency[] = [];
  const seen = new Set<string>();
  visit(ast, owner, dependencies, seen, options.sheetTables, options.sheetOrder, new Set(), options.valueDependencies ?? false, options.resolveNameAst);
  addProjectedConsumerDependencies(ast, owner, options, dependencies, seen);
  return dependencies;
}

/** SUMIF projects the target using criteria dimensions, including beyond its authored endpoint. */
function addProjectedConsumerDependencies(ast: FormulaAst, owner: CellAddress, options: CollectFormulaDependenciesOptions, dependencies: FormulaDependency[], seen: Set<string>): void {
  function envelope(node: FormulaAst, bindings: ReadonlyMap<string, FormulaAst>, names = new Set<string>()): RangeDependency | undefined {
    if (node.type === 'cell-reference') { const address = resolveCellReference(node.reference, owner, options.sheetOrder); return { kind: 'range', start: address, end: address }; }
    if (node.type === 'range-reference') return resolveRangeReference(node, owner, options.sheetOrder);
    if (node.type === 'whole-column-reference') { const sheetId = resolveFormulaSheetId(node.sheetId, owner.sheetId, options.sheetOrder); return { kind: 'range', start: { sheetId, row: 0, column: node.startColumn }, end: { sheetId, row: 1_048_575, column: node.endColumn } }; }
    if (node.type === 'whole-row-reference') { const sheetId = resolveFormulaSheetId(node.sheetId, owner.sheetId, options.sheetOrder); return { kind: 'range', start: { sheetId, row: node.startRow, column: 0 }, end: { sheetId, row: node.endRow, column: 16_383 } }; }
    if (node.type === 'table-reference' && options.sheetTables) { const resolved = resolveSheetTableReference(node.tableName, node, owner, options.sheetTables); return isFormulaError(resolved) ? undefined : 'start' in resolved ? resolved : { kind: 'range', start: resolved, end: resolved }; }
    if (node.type === 'name-reference') {
      const id = node.name.toUpperCase();
      if (names.has(id)) return undefined;
      names.add(id);
      const definition = bindings.get(id) ?? options.resolveNameAst?.(node.name, owner);
      return definition ? envelope(definition, bindings, names) : undefined;
    }
    if (node.type === 'function-call' && node.name.toUpperCase() === 'INDEX' && node.arguments[0]) return envelope(node.arguments[0], bindings, names);
    return undefined;
  }
  function walk(node: FormulaAst, bindings: ReadonlyMap<string, FormulaAst> = new Map()): void {
    if (node.type === 'function-call') {
      const id = node.name.toUpperCase();
      if (id === 'LET') {
        const local = new Map(bindings);
        for (let index = 0; index < node.arguments.length - 1; index += 2) {
          const variable = node.arguments[index]!;
          const value = node.arguments[index + 1]!;
          walk(value, local);
          if (variable.type === 'name-reference') local.set(variable.name.toUpperCase(), value);
        }
        if (node.arguments.length) walk(node.arguments[node.arguments.length - 1]!, local);
        return;
      }
      if ((id === 'SUMIF' || id === 'AVERAGEIF') && node.arguments.length === 3) {
        const criteria = envelope(node.arguments[0]!, bindings);
        const target = envelope(node.arguments[2]!, bindings);
        if (criteria && target) addDependency({ kind: 'range', start: target.start, end: { sheetId: target.start.sheetId, row: Math.min(1_048_575, target.end.row + criteria.end.row - criteria.start.row), column: Math.min(16_383, target.end.column + criteria.end.column - criteria.start.column) } }, dependencies, seen);
      }
      if (node.callee) walk(node.callee, bindings);
      for (const argument of node.arguments) walk(argument, bindings);
    } else if (node.type === 'binary-expression') { walk(node.left, bindings); walk(node.right, bindings); }
    else if (node.type === 'unary-expression') walk(node.operand, bindings);
  }
  walk(ast);
}

export function collectFormulaReferenceNodes(ast: FormulaAst): readonly FormulaReferenceNode[] {
  const references: FormulaReferenceNode[] = [];
  const visitNode = (node: FormulaAst): void => {
    switch (node.type) {
      case 'cell-reference':
      case 'range-reference':
      case 'whole-column-reference':
      case 'whole-row-reference':
      case 'reference-union':
      case 'reference-intersection':
      case 'sheet-range-reference':
      case 'external-reference':
      case 'spill-reference':
        references.push(node);
        return;
      case 'unary-expression':
        visitNode(node.operand);
        return;
      case 'binary-expression':
        visitNode(node.left);
        visitNode(node.right);
        return;
      case 'function-call':
        if (node.callee) visitNode(node.callee);
        for (const argument of node.arguments) visitNode(argument);
        return;
      case 'error-literal':
      case 'number-literal':
      case 'string-literal':
      case 'boolean-literal':
      case 'name-reference':
      case 'invalid-reference':
        return;
      case 'table-reference':
        return;
    }
  };
  visitNode(ast);
  return references.map((reference) => structuredClone(reference));
}

export function resolveCellReference(
  reference: ParsedCellReference,
  currentCell: CellAddress,
  sheetOrder: readonly FormulaSheetIdentity[],
): CellAddress {
  const sheetId = reference.sheetId ?? currentCell.sheetId;
  if (!sheetId) throw new FormulaReferenceError('Cell reference is missing a worksheet');
  return { sheetId: resolveFormulaSheetId(reference.sheetId, currentCell.sheetId, sheetOrder), row: reference.row, column: reference.column };
}

export function resolveRangeReference(
  node: RangeReferenceNode,
  currentCell: CellAddress,
  sheetOrder: readonly FormulaSheetIdentity[],
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
  sheetTables: ReadonlyMap<string, SheetTableRef> | undefined,
  sheetOrder: readonly FormulaSheetIdentity[],
  bound: ReadonlySet<string> = new Set(),
  valueDependencies = false,
  resolveNameAst?: CollectFormulaDependenciesOptions['resolveNameAst'],
): void {
  function visitGeometryInputs(argument: FormulaAst, names = new Set<string>()): void {
    if (['cell-reference', 'range-reference', 'whole-row-reference', 'whole-column-reference', 'sheet-range-reference', 'external-reference', 'table-reference'].includes(argument.type)) return;
    if (argument.type === 'reference-union') { for (const reference of argument.references) visitGeometryInputs(reference, names); return; }
    if (argument.type === 'reference-intersection') { visitGeometryInputs(argument.left, names); visitGeometryInputs(argument.right, names); return; }
    if (argument.type === 'name-reference' && resolveNameAst) {
      const id = argument.name.toUpperCase();
      if (names.has(id)) return;
      const definition = resolveNameAst(argument.name, owner);
      if (definition) { visitGeometryInputs(definition, new Set([...names, id])); return; }
    }
    if (argument.type === 'function-call' && ['INDEX', 'OFFSET'].includes(argument.name.toUpperCase())) {
      if (argument.arguments[0]) visitGeometryInputs(argument.arguments[0], names);
      for (const control of argument.arguments.slice(1)) visit(control, owner, dependencies, seen, sheetTables, sheetOrder, bound, true, resolveNameAst);
      return;
    }
    visit(argument, owner, dependencies, seen, sheetTables, sheetOrder, bound, true, resolveNameAst);
  }
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
      visit(node.operand, owner, dependencies, seen, sheetTables, sheetOrder, bound, valueDependencies, resolveNameAst);
      return;
    case 'spill-reference':
      visit(node.operand, owner, dependencies, seen, sheetTables, sheetOrder, bound, valueDependencies, resolveNameAst);
      return;
    case 'binary-expression':
      visit(node.left, owner, dependencies, seen, sheetTables, sheetOrder, bound, valueDependencies, resolveNameAst);
      visit(node.right, owner, dependencies, seen, sheetTables, sheetOrder, bound, valueDependencies, resolveNameAst);
      return;
    case 'function-call':
      if (valueDependencies && ['ROW', 'COLUMN', 'ROWS', 'COLUMNS'].includes(node.name.toUpperCase())) {
        for (const argument of node.arguments) visitGeometryInputs(argument);
        return;
      }
      if (visitLexicalArguments(node, bound, (child, local) => visit(child, owner, dependencies, seen, sheetTables, sheetOrder, local, valueDependencies, resolveNameAst))) return;
      if (node.callee) visit(node.callee, owner, dependencies, seen, sheetTables, sheetOrder, bound, valueDependencies, resolveNameAst);
      if (node.name && !bound.has(node.name.toUpperCase()) && !getFunctionDescriptor(node.name)) addDependency({ kind: 'name', name: node.name.toUpperCase() }, dependencies, seen);
      for (const argument of node.arguments) visit(argument, owner, dependencies, seen, sheetTables, sheetOrder, bound, valueDependencies, resolveNameAst);
      return;
    case 'name-reference': {
      if (bound.has(node.name.toUpperCase())) return;
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
    case 'error-literal':
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
  sheetTables: ReadonlyMap<string, SheetTableRef> | undefined,
  sheetOrder: readonly FormulaSheetIdentity[],
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
