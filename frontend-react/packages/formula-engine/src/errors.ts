export class FormulaLexError extends Error {
  readonly position: number;

  constructor(message: string, position: number) {
    super(message);
    this.name = 'FormulaLexError';
    this.position = position;
  }
}

export class FormulaSyntaxError extends Error {
  readonly position: number;

  constructor(message: string, position: number) {
    super(message);
    this.name = 'FormulaSyntaxError';
    this.position = position;
  }
}

export class FormulaReferenceError extends Error {
  constructor(message: string) {
    super(message);
    this.name = 'FormulaReferenceError';
  }
}
