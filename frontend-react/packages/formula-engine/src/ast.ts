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
}

export interface StringLiteralNode {
  readonly type: 'string-literal';
  readonly value: string;
  readonly span: SourceSpan;
}

export interface BooleanLiteralNode {
  readonly type: 'boolean-literal';
  readonly value: boolean;
  readonly span: SourceSpan;
}

export interface CellReferenceNode {
  readonly type: 'cell-reference';
  readonly reference: ParsedCellReference;
  readonly span: SourceSpan;
}

export interface RangeReferenceNode {
  readonly type: 'range-reference';
  readonly start: CellReferenceNode;
  readonly end: CellReferenceNode;
  readonly span: SourceSpan;
}

export interface UnaryExpressionNode {
  readonly type: 'unary-expression';
  readonly operator: '+' | '-' | '%';
  readonly operand: FormulaAst;
  readonly span: SourceSpan;
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
}

export interface FunctionCallNode {
  readonly type: 'function-call';
  readonly name: string;
  readonly arguments: readonly FormulaAst[];
  readonly span: SourceSpan;
}

export interface NameReferenceNode {
  readonly type: 'name-reference';
  readonly name: string;
  readonly span: SourceSpan;
}

export type TableReferenceSpecifier = 'all' | 'headers' | 'data' | 'totals';

export interface TableReferenceNode {
  readonly type: 'table-reference';
  readonly tableName: string;
  readonly specifier?: TableReferenceSpecifier;
  readonly columnName?: string;
  readonly thisRow: boolean;
  readonly span: SourceSpan;
}

export type FormulaAst =
  | NumberLiteralNode
  | StringLiteralNode
  | BooleanLiteralNode
  | CellReferenceNode
  | RangeReferenceNode
  | NameReferenceNode
  | TableReferenceNode
  | UnaryExpressionNode
  | BinaryExpressionNode
  | FunctionCallNode;
