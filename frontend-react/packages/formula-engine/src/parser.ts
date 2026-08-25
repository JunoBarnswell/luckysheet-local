import type {
  BinaryExpressionNode,
  BinaryOperator,
  BooleanLiteralNode,
  CellReferenceNode,
  FormulaAst,
  FunctionCallNode,
  NumberLiteralNode,
  ParsedCellReference,
  RangeReferenceNode,
  SpillReferenceNode,
  SourceSpan,
  StringLiteralNode,
  TableReferenceNode,
  TableReferenceSpecifier,
  UnaryExpressionNode,
} from './ast';
import { tryParseCellReferenceText } from './address';
import { FormulaSyntaxError } from './errors';
import { lexFormula, type Token, type TokenKind } from './lexer';

export function parseFormula(source: string): FormulaAst {
  return new Parser(lexFormula(source)).parse();
}

class Parser {
  private index = 0;

  constructor(private readonly tokens: readonly Token[]) {}

  parse(): FormulaAst {
    const expression = this.parseComparison();
    if (!this.check('eof')) {
      const token = this.peek();
      throw new FormulaSyntaxError(`Unexpected token: ${token.lexeme || token.kind}`, token.span.start);
    }
    return expression;
  }

  private parseComparison(): FormulaAst {
    let expression = this.parseConcat();
    while (
      this.check('equal') ||
      this.check('not-equal') ||
      this.check('less-than') ||
      this.check('less-than-equal') ||
      this.check('greater-than') ||
      this.check('greater-than-equal')
    ) {
      const opToken = this.advance();
      let operator: BinaryOperator;
      switch (opToken.kind) {
        case 'equal': operator = '='; break;
        case 'not-equal': operator = '<>'; break;
        case 'less-than': operator = '<'; break;
        case 'less-than-equal': operator = '<='; break;
        case 'greater-than': operator = '>'; break;
        case 'greater-than-equal': operator = '>='; break;
        default: operator = '=';
      }
      const right = this.parseConcat();
      const node: BinaryExpressionNode = {
        type: 'binary-expression',
        operator,
        left: expression,
        right,
        span: spanFrom(expression.span, right.span),
      };
      expression = node;
    }
    return expression;
  }

  private parseConcat(): FormulaAst {
    let expression = this.parseAdditive();
    while (this.check('ampersand')) {
      this.advance();
      const right = this.parseAdditive();
      const node: BinaryExpressionNode = {
        type: 'binary-expression',
        operator: '&',
        left: expression,
        right,
        span: spanFrom(expression.span, right.span),
      };
      expression = node;
    }
    return expression;
  }

  private parseAdditive(): FormulaAst {
    let expression = this.parseMultiplicative();
    while (this.check('plus') || this.check('minus')) {
      const operatorToken = this.advance();
      const right = this.parseMultiplicative();
      const node: BinaryExpressionNode = {
        type: 'binary-expression',
        operator: operatorToken.kind === 'plus' ? '+' : '-',
        left: expression,
        right,
        span: spanFrom(expression.span, right.span),
      };
      expression = node;
    }
    return expression;
  }

  private parseMultiplicative(): FormulaAst {
    let expression = this.parseExponentiation();
    while (this.check('star') || this.check('slash')) {
      const operatorToken = this.advance();
      const right = this.parseExponentiation();
      const node: BinaryExpressionNode = {
        type: 'binary-expression',
        operator: operatorToken.kind === 'star' ? '*' : '/',
        left: expression,
        right,
        span: spanFrom(expression.span, right.span),
      };
      expression = node;
    }
    return expression;
  }

  private parseExponentiation(): FormulaAst {
    let expression = this.parseUnary();
    while (this.check('caret')) {
      this.advance();
      const right = this.parseUnary();
      const node: BinaryExpressionNode = {
        type: 'binary-expression',
        operator: '^',
        left: expression,
        right,
        span: spanFrom(expression.span, right.span),
      };
      expression = node;
    }
    return expression;
  }

  private parseUnary(): FormulaAst {
    if (this.check('plus') || this.check('minus') || this.check('at-sign')) {
      const operatorToken = this.advance();
      const operand = this.parseUnary();
      const node: UnaryExpressionNode = {
        type: 'unary-expression',
        operator: operatorToken.kind === 'plus' ? '+' : operatorToken.kind === 'minus' ? '-' : '@',
        operand,
        span: spanFrom(operatorToken.span, operand.span),
      };
      return node;
    }
    let expr = this.parsePrimary();
    while (this.check('percent') || this.check('spill-operator')) {
      const operatorToken = this.advance();
      if (operatorToken.kind === 'spill-operator') {
        const node: SpillReferenceNode = {
          type: 'spill-reference',
          operand: expr,
          span: spanFrom(expr.span, operatorToken.span),
        };
        expr = node;
        continue;
      }
      const node: UnaryExpressionNode = {
        type: 'unary-expression',
        operator: '%',
        operand: expr,
        span: spanFrom(expr.span, operatorToken.span),
      };
      expr = node;
    }
    return expr;
  }

  private parsePrimary(): FormulaAst {
    const token = this.peek();
    if (token.kind === 'number') {
      this.advance();
      const node: NumberLiteralNode = { type: 'number-literal', value: Number(token.lexeme), span: token.span };
      return node;
    }

    if (token.kind === 'error-reference') {
      this.advance();
      return { type: 'invalid-reference', code: '#REF!', span: token.span };
    }

    if (token.kind === 'string' && !this.checkNext('bang')) {
      this.advance();
      const node: StringLiteralNode = { type: 'string-literal', value: token.value ?? '', span: token.span };
      return node;
    }

    if (token.kind === 'identifier') {
      const upper = token.lexeme.toUpperCase();
      if (upper === 'TRUE' || upper === 'FALSE') {
        if (!this.checkNext('left-paren')) {
          this.advance();
          const node: BooleanLiteralNode = {
            type: 'boolean-literal',
            value: upper === 'TRUE',
            span: token.span,
          };
          return node;
        }
      }
      if (this.checkNext('left-paren')) {
        return this.parseFunctionCall();
      }
      if (this.checkNext('left-bracket')) {
        this.advance();
        return this.parseTableReference(token);
      }
      const savedIndex = this.index;
      try {
        return this.parseReference();
      } catch {
        // parseReference consumes the leading identifier before failing.
        if (this.index === savedIndex) this.advance();
        return { type: 'name-reference', name: token.lexeme, span: token.span };
      }
    }

    if (token.kind === 'string') {
      return this.parseReference();
    }

    if (this.match('left-paren')) {
      const expression = this.parseComparison();
      this.expect('right-paren', 'Expected closing parenthesis');
      // Keep explicit grouping in the AST.  A formatter cannot otherwise
      // distinguish `=(A1+B1)*C1` from `=A1+B1*C1` after parsing.
      return { ...expression, parenthesized: true };
    }

    throw new FormulaSyntaxError(`Expected expression, found ${token.lexeme || token.kind}`, token.span.start);
  }

  private parseFunctionCall(): FunctionCallNode {
    const nameToken = this.expect('identifier', 'Expected function name');
    this.expect('left-paren', 'Expected opening parenthesis');
    const argumentsList: FormulaAst[] = [];
    if (!this.check('right-paren')) {
      do {
        if (this.check('right-paren')) break;
        argumentsList.push(this.parseComparison());
      } while (this.match('comma'));
    }
    const closingToken = this.expect('right-paren', 'Expected closing parenthesis');
    return {
      type: 'function-call',
      name: nameToken.lexeme,
      arguments: argumentsList,
      span: { start: nameToken.span.start, end: closingToken.span.end },
    };
  }

  private parseTableReference(nameToken: Token): TableReferenceNode {
    this.expect('left-bracket', 'Expected [ after table name');

    if (this.match('left-bracket')) {
      const specifier = this.parseTableSpecifierToken();
      this.expect('right-bracket', 'Expected ] after table specifier');
      this.expect('comma', 'Expected , between structured table references');
      this.expect('left-bracket', 'Expected [ before table column');
      const columnName = this.parseTableColumnName();
      const innerClose = this.expect('right-bracket', 'Expected ] after table column');
      const outerClose = this.expect('right-bracket', 'Expected ] after structured table reference');
      return {
        type: 'table-reference',
        tableName: nameToken.lexeme,
        specifier,
        columnName,
        thisRow: false,
        span: { start: nameToken.span.start, end: outerClose.span.end },
      };
    }

    const thisRow = this.match('at-sign');
    if (this.check('table-specifier')) {
      const specifier = this.parseTableSpecifierToken();
      const closingToken = this.expect('right-bracket', 'Expected ] after table specifier');
      return {
        type: 'table-reference',
        tableName: nameToken.lexeme,
        specifier,
        thisRow: false,
        span: { start: nameToken.span.start, end: closingToken.span.end },
      };
    }

    const columnName = this.parseTableColumnName();
    const closingToken = this.expect('right-bracket', 'Expected ] after table column');
    return {
      type: 'table-reference',
      tableName: nameToken.lexeme,
      columnName,
      thisRow,
      span: { start: nameToken.span.start, end: closingToken.span.end },
    };
  }

  private parseTableColumnName(): string {
    const columnToken = this.peek();
    if (columnToken.kind === 'string') {
      this.advance();
      return columnToken.value ?? '';
    }
    return this.expect('identifier', 'Expected table column name').lexeme;
  }

  private parseTableSpecifierToken(): TableReferenceSpecifier {
    const token = this.expect('table-specifier', 'Expected table specifier such as #All or #Data');
    const normalized = token.value?.trim().toLowerCase() ?? '';
    switch (normalized) {
      case 'all':
        return 'all';
      case 'headers':
        return 'headers';
      case 'data':
        return 'data';
      case 'totals':
        return 'totals';
      default:
        throw new FormulaSyntaxError(`Unknown table specifier: ${token.lexeme}`, token.span.start);
    }
  }

  private parseReference(): CellReferenceNode | RangeReferenceNode {
    const firstToken = this.advance();
    let sheetId: string | undefined;
    let cellToken = firstToken;
    if (this.match('bang')) {
      sheetId = firstToken.kind === 'string' ? firstToken.value ?? '' : firstToken.lexeme;
      cellToken = this.expect('identifier', 'Expected cell reference after sheet name');
    }

    const start = this.createCellReference(cellToken, sheetId);
    if (!this.match('colon')) return start;

    let endSheetId: string | undefined;
    let endToken = this.peek();
    if ((endToken.kind === 'identifier' || endToken.kind === 'string') && this.checkNext('bang')) {
      this.advance();
      endSheetId = endToken.kind === 'string' ? endToken.value ?? '' : endToken.lexeme;
      this.expect('bang', 'Expected separator after sheet name');
      endToken = this.peek();
    }
    endToken = this.expect('identifier', 'Expected ending cell reference');
    const end = this.createCellReference(endToken, endSheetId);
    return {
      type: 'range-reference',
      start,
      end,
      span: { start: start.span.start, end: end.span.end },
    };
  }

  private createCellReference(token: Token, sheetId: string | undefined): CellReferenceNode {
    const reference = token.kind === 'identifier' ? tryParseCellReferenceText(token.lexeme) : undefined;
    if (!reference) throw new FormulaSyntaxError(`Invalid cell reference: ${token.lexeme}`, token.span.start);
    const parsedReference: ParsedCellReference = sheetId === undefined ? reference : { ...reference, sheetId };
    return { type: 'cell-reference', reference: parsedReference, span: token.span };
  }

  private expect(kind: TokenKind, message: string): Token {
    if (this.check(kind)) return this.advance();
    const token = this.peek();
    throw new FormulaSyntaxError(message, token.span.start);
  }

  private match(kind: TokenKind): boolean {
    if (!this.check(kind)) return false;
    this.advance();
    return true;
  }

  private check(kind: TokenKind): boolean {
    return this.peek().kind === kind;
  }

  private checkNext(kind: TokenKind): boolean {
    return (this.tokens[this.index + 1] ?? this.tokens[this.index])?.kind === kind;
  }

  private advance(): Token {
    const token = this.peek();
    if (this.index < this.tokens.length - 1) this.index += 1;
    return token;
  }

  private peek(): Token {
    return this.tokens[this.index] ?? this.tokens[this.tokens.length - 1]!;
  }
}

function spanFrom(start: SourceSpan, end: SourceSpan): SourceSpan {
  return { start: start.start, end: end.end };
}
