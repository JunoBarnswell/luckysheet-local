import type {
  BinaryExpressionNode,
  CellReferenceNode,
  FormulaAst,
  FunctionCallNode,
  NumberLiteralNode,
  ParsedCellReference,
  RangeReferenceNode,
  SourceSpan,
  StringLiteralNode,
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
    const expression = this.parseAdditive();
    if (!this.check('eof')) {
      const token = this.peek();
      throw new FormulaSyntaxError(`Unexpected token: ${token.lexeme || token.kind}`, token.span.start);
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
    let expression = this.parseUnary();
    while (this.check('star') || this.check('slash')) {
      const operatorToken = this.advance();
      const right = this.parseUnary();
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

  private parseUnary(): FormulaAst {
    if (this.check('plus') || this.check('minus')) {
      const operatorToken = this.advance();
      const operand = this.parseUnary();
      const node: UnaryExpressionNode = {
        type: 'unary-expression',
        operator: operatorToken.kind === 'plus' ? '+' : '-',
        operand,
        span: spanFrom(operatorToken.span, operand.span),
      };
      return node;
    }
    return this.parsePrimary();
  }

  private parsePrimary(): FormulaAst {
    const token = this.peek();
    if (token.kind === 'number') {
      this.advance();
      const node: NumberLiteralNode = { type: 'number-literal', value: Number(token.lexeme), span: token.span };
      return node;
    }

    if (token.kind === 'string' && !this.checkNext('bang')) {
      this.advance();
      const node: StringLiteralNode = { type: 'string-literal', value: token.value ?? '', span: token.span };
      return node;
    }

    if (token.kind === 'identifier' || token.kind === 'string') {
      if (token.kind === 'identifier' && this.checkNext('left-paren')) return this.parseFunctionCall();
      return this.parseReference();
    }

    if (this.match('left-paren')) {
      const expression = this.parseAdditive();
      this.expect('right-paren', 'Expected closing parenthesis');
      return expression;
    }

    throw new FormulaSyntaxError(`Expected expression, found ${token.lexeme || token.kind}`, token.span.start);
  }

  private parseFunctionCall(): FunctionCallNode {
    const nameToken = this.expect('identifier', 'Expected function name');
    this.expect('left-paren', 'Expected opening parenthesis');
    const argumentsList: FormulaAst[] = [];
    if (!this.check('right-paren')) {
      do {
        argumentsList.push(this.parseAdditive());
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
