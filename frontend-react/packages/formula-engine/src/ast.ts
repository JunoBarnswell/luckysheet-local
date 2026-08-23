export interface SourceSpan {
  readonly start: number;
  readonly end: number;
}

export interface CellAddress {
  readonly sheetId: string;
  readonly row: number;
  readonly column: number;
}

export interface ParsedCellReference {
  readonly sheetId?: string;
  readonly row: number;
  readonly column: number;
  readonly absoluteRow: boolean;
  readonly absoluteColumn: boolean;
}

export interface NumberLiteralNode {
  readonly type: 'number-literal';
  readonly value: number;
  readonly span: SourceSpan;
  /** True when the source contained explicit grouping parentheses. */
  readonly parenthesized?: boolean;
}

export interface StringLiteralNode {
  readonly type: 'string-literal';
  readonly value: string;
  readonly span: SourceSpan;
  readonly parenthesized?: boolean;
}

export interface BooleanLiteralNode {
  readonly type: 'boolean-literal';
  readonly value: boolean;
  readonly span: SourceSpan;
  readonly parenthesized?: boolean;
}

export interface CellReferenceNode {
  readonly type: 'cell-reference';
  readonly reference: ParsedCellReference;
  readonly span: SourceSpan;
  readonly parenthesized?: boolean;
}

/**
 * A reference which was invalidated by a structural mutation.
 *
 * Keeping this as an AST node (instead of clamping the coordinate to a
 * surviving cell) is important: the formula text remains round-trippable as
 * `#REF!`, dependency collection stays fail-closed, and a subsequent
 * recalculation cannot silently read a different cell.
 */
export interface InvalidReferenceNode {
  readonly type: 'invalid-reference';
  readonly code: '#REF!';
  readonly span: SourceSpan;
  readonly parenthesized?: boolean;
}

export interface RangeReferenceNode {
  readonly type: 'range-reference';
  readonly start: CellReferenceNode;
  readonly end: CellReferenceNode;
  readonly span: SourceSpan;
  readonly parenthesized?: boolean;
}

export interface UnaryExpressionNode {
  readonly type: 'unary-expression';
  readonly operator: '+' | '-' | '%';
  readonly operand: FormulaAst;
  readonly span: SourceSpan;
  readonly parenthesized?: boolean;
}

export type BinaryOperator =
  | '+'
  | '-'
  | '*'
  | '/'
  | '^'
  | '&'
  | '='
  | '<>'
  | '<'
  | '<='
  | '>'
  | '>=';

export interface BinaryExpressionNode {
  readonly type: 'binary-expression';
  readonly operator: BinaryOperator;
  readonly left: FormulaAst;
  readonly right: FormulaAst;
  readonly span: SourceSpan;
  readonly parenthesized?: boolean;
}

export interface FunctionCallNode {
  readonly type: 'function-call';
  readonly name: string;
  readonly arguments: readonly FormulaAst[];
  readonly span: SourceSpan;
  readonly parenthesized?: boolean;
}

export interface NameReferenceNode {
  readonly type: 'name-reference';
  readonly name: string;
  readonly span: SourceSpan;
  readonly parenthesized?: boolean;
}

export type TableReferenceSpecifier = 'all' | 'headers' | 'data' | 'totals';

export interface TableReferenceNode {
  readonly type: 'table-reference';
  readonly tableName: string;
  readonly specifier?: TableReferenceSpecifier;
  readonly columnName?: string;
  readonly thisRow: boolean;
  readonly span: SourceSpan;
  readonly parenthesized?: boolean;
}

export type FormulaAst =
  | NumberLiteralNode
  | StringLiteralNode
  | BooleanLiteralNode
  | CellReferenceNode
  | InvalidReferenceNode
  | RangeReferenceNode
  | NameReferenceNode
  | TableReferenceNode
  | UnaryExpressionNode
  | BinaryExpressionNode
  | FunctionCallNode;
