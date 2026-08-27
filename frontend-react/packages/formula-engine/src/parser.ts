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
  FormulaReferenceNode,
  ReferenceIntersectionNode,
  ReferenceUnionNode,
  SheetRangeReferenceNode,
  ExternalReferenceNode,
  WholeColumnReferenceNode,
  WholeRowReferenceNode,
  SpillReferenceNode,
  SourceSpan,
  StringLiteralNode,
  TableReferenceNode,
  TableReferenceSpecifier,
  UnaryExpressionNode,
} from './ast';
import { columnNameToIndex, tryParseCellReferenceText } from './address';
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

    if (token.kind === 'left-bracket') {
      return this.parseExternalReference();
    }

    if (this.match('left-paren')) {
      const expression = this.parseComparison();
      if (this.match('comma')) {
        const references: FormulaReferenceNode[] = [this.expectReference(expression)];
        do {
          references.push(this.expectReference(this.parseComparison()));
        } while (this.match('comma'));
        const closingToken = this.expect('right-paren', 'Expected closing parenthesis');
        const node: ReferenceUnionNode = {
          type: 'reference-union',
          references,
          span: { start: expression.span.start, end: closingToken.span.end },
        };
        return node;
      }
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
      if (!this.check('table-specifier')) {
        const columnName = this.parseTableColumnName();
        this.expect('right-bracket', 'Expected ] after structured table column');
        this.expect('colon', 'Expected : between structured table columns');
        this.expect('left-bracket', 'Expected [ before ending table column');
        const columnEndName = this.parseTableColumnName();
        const innerClose = this.expect('right-bracket', 'Expected ] after ending table column');
        const outerClose = this.expect('right-bracket', 'Expected ] after structured table reference');
        return {
          type: 'table-reference',
          tableName: nameToken.lexeme,
          columnName,
          columnEndName,
          thisRow: false,
          span: { start: nameToken.span.start, end: outerClose.span.end },
        };
      }
      const specifier = this.parseTableSpecifierToken();
      this.expect('right-bracket', 'Expected ] after table specifier');
      this.expect('comma', 'Expected , between structured table references');
      this.expect('left-bracket', 'Expected [ before table column');
      const columnName = this.parseTableColumnName();
      let columnEndName: string | undefined;
      if (this.match('right-bracket') && this.match('colon')) {
        this.expect('left-bracket', 'Expected [ before ending table column');
        columnEndName = this.parseTableColumnName();
      } else {
        // The common single-column form has not consumed its inner close yet.
        if (this.tokens[this.index - 1]?.kind !== 'right-bracket') this.expect('right-bracket', 'Expected ] after table column');
      }
      const innerClose = this.tokens[this.index - 1]!.kind === 'right-bracket' ? this.tokens[this.index - 1]! : this.expect('right-bracket', 'Expected ] after table column');
      const outerClose = this.expect('right-bracket', 'Expected ] after structured table reference');
      return {
        type: 'table-reference',
        tableName: nameToken.lexeme,
        specifier,
        columnName,
        columnEndName,
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

  private parseReference(): FormulaReferenceNode {
    const firstToken = this.peek();
    if (firstToken.kind === 'left-bracket') {
      return this.parseExternalReference();
    }

    if (this.isSheetRangeQualifier(firstToken)) {
      this.advance();
      this.expect('colon', 'Expected worksheet range separator');
      const endSheetToken = this.expectAny(['identifier', 'string'], 'Expected ending worksheet name');
      this.expect('bang', 'Expected separator after worksheet range');
      const startReference = this.parseReferenceEndpoint(this.peek(), undefined, this.check('colon'));
      const reference = this.match('colon')
        ? this.combineReferenceEndpoints(startReference, this.parseReferenceEndpoint(this.peek(), undefined))
        : startReference;
      const node: SheetRangeReferenceNode = {
        type: 'sheet-range-reference',
        qualifier: {
          startSheetId: firstToken.kind === 'string' ? firstToken.value ?? '' : firstToken.lexeme,
          endSheetId: endSheetToken.kind === 'string' ? endSheetToken.value ?? '' : endSheetToken.lexeme,
        },
        reference: reference as SheetRangeReferenceNode['reference'],
        span: { start: firstToken.span.start, end: reference.span.end },
      };
      return this.parseReferenceOperators(node);
    }

    let sheetId: string | undefined;
    let cellToken = firstToken;
    if (this.checkNext('bang')) {
      this.advance();
      this.expect('bang', 'Expected separator after sheet name');
      sheetId = firstToken.kind === 'string' ? firstToken.value ?? '' : firstToken.lexeme;
      cellToken = this.peek();
    }

    const start = this.parseReferenceEndpoint(cellToken, sheetId, this.check('colon'));
    if (!this.match('colon')) return this.parseReferenceOperators(start);

    let endSheetId: string | undefined;
    let endToken = this.peek();
    if ((endToken.kind === 'identifier' || endToken.kind === 'string') && this.checkNext('bang')) {
      this.advance();
      endSheetId = endToken.kind === 'string' ? endToken.value ?? '' : endToken.lexeme;
      this.expect('bang', 'Expected separator after sheet name');
      endToken = this.peek();
    }
    endToken = this.peek();
    const end = this.parseReferenceEndpoint(endToken, endSheetId);
    const reference = this.combineReferenceEndpoints(start, end);
    return this.parseReferenceOperators(reference);
  }

  private parseReferenceOperators(initial: FormulaReferenceNode): FormulaReferenceNode {
    let expression = initial;
    while (this.match('reference-intersection')) {
      const right = this.parseReferenceOperand();
      const node: ReferenceIntersectionNode = {
        type: 'reference-intersection',
        left: expression,
        right,
        span: { start: expression.span.start, end: right.span.end },
      };
      expression = node;
    }
    return expression;
  }

  private parseReferenceOperand(): FormulaReferenceNode {
    const token = this.peek();
    if (token.kind !== 'identifier' && token.kind !== 'string' && token.kind !== 'number') {
      throw new FormulaSyntaxError('Expected reference after intersection operator', token.span.start);
    }
    return this.parseReference();
  }

  private parseReferenceEndpoint(token: Token, sheetId: string | undefined, allowWhole = true): CellReferenceNode | WholeColumnReferenceNode | WholeRowReferenceNode {
    if (token.kind === 'identifier') {
      const cell = tryParseCellReferenceText(token.lexeme);
      if (cell) {
        this.advance();
        return { type: 'cell-reference', reference: sheetId === undefined ? cell : { ...cell, sheetId }, span: token.span };
      }
      const column = allowWhole ? columnNameToIndex(token.lexeme.replace(/^\$/, '')) : undefined;
      if (column !== undefined) {
        this.advance();
        return { type: 'whole-column-reference', sheetId, startColumn: column, endColumn: column, span: token.span };
      }
    }
    if (allowWhole && token.kind === 'number' && /^\d+$/.test(token.lexeme)) {
      const row = Number(token.lexeme);
      if (row > 0 && Number.isSafeInteger(row)) {
        this.advance();
        return { type: 'whole-row-reference', sheetId, startRow: row - 1, endRow: row - 1, span: token.span };
      }
    }
    throw new FormulaSyntaxError(`Invalid reference endpoint: ${token.lexeme}`, token.span.start);
  }

  private combineReferenceEndpoints(
    start: CellReferenceNode | WholeColumnReferenceNode | WholeRowReferenceNode,
    end: CellReferenceNode | WholeColumnReferenceNode | WholeRowReferenceNode,
  ): FormulaReferenceNode {
    if (start.type === 'cell-reference' && end.type === 'cell-reference') {
      return { type: 'range-reference', start, end, span: { start: start.span.start, end: end.span.end } };
    }
    if (start.type === 'whole-column-reference' && end.type === 'whole-column-reference') {
      return {
        type: 'whole-column-reference',
        sheetId: start.sheetId ?? end.sheetId,
        startColumn: Math.min(start.startColumn, end.startColumn),
        endColumn: Math.max(start.endColumn, end.endColumn),
        span: { start: start.span.start, end: end.span.end },
      };
    }
    if (start.type === 'whole-row-reference' && end.type === 'whole-row-reference') {
      return {
        type: 'whole-row-reference',
        sheetId: start.sheetId ?? end.sheetId,
        startRow: Math.min(start.startRow, end.startRow),
        endRow: Math.max(start.endRow, end.endRow),
        span: { start: start.span.start, end: end.span.end },
      };
    }
    throw new FormulaSyntaxError('Reference range endpoints must use the same reference domain', end.span.start);
  }

  private parseExternalReference(): ExternalReferenceNode {
    this.expect('left-bracket', 'Expected [ before external workbook');
    const workbook = this.expect('identifier', 'Expected external workbook identity');
    this.expect('right-bracket', 'Expected ] after external workbook identity');
    const sheet = this.expectAny(['identifier', 'string'], 'Expected worksheet after external workbook');
    this.expect('bang', 'Expected separator after external worksheet');
    const startReference = this.parseReferenceEndpoint(this.peek(), undefined, this.check('colon'));
    const reference = this.match('colon')
      ? this.combineReferenceEndpoints(startReference, this.parseReferenceEndpoint(this.peek(), undefined))
      : startReference;
    return {
      type: 'external-reference',
      qualifier: {
        workbookId: workbook.lexeme,
        sheetId: sheet.kind === 'string' ? sheet.value ?? '' : sheet.lexeme,
      },
      reference: reference as ExternalReferenceNode['reference'],
      span: { start: workbook.span.start - 1, end: reference.span.end },
    };
  }

  private isSheetRangeQualifier(token: Token): boolean {
    return (token.kind === 'identifier' || token.kind === 'string')
      && this.tokens[this.index + 1]?.kind === 'colon'
      && (this.tokens[this.index + 2]?.kind === 'identifier' || this.tokens[this.index + 2]?.kind === 'string')
      && this.tokens[this.index + 3]?.kind === 'bang';
  }

  private expectAny(kinds: readonly TokenKind[], message: string): Token {
    if (kinds.includes(this.peek().kind)) return this.advance();
    throw new FormulaSyntaxError(message, this.peek().span.start);
  }

  private expectReference(node: FormulaAst): FormulaReferenceNode {
    if (node.type === 'cell-reference' || node.type === 'range-reference' || node.type === 'whole-column-reference' || node.type === 'whole-row-reference' || node.type === 'spill-reference' || node.type === 'table-reference' || node.type === 'reference-union' || node.type === 'reference-intersection' || node.type === 'sheet-range-reference' || node.type === 'external-reference' || node.type === 'invalid-reference') return node;
    throw new FormulaSyntaxError('Reference union requires reference operands', node.span.start);
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
