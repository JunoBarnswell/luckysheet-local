import type { FormulaAst, FunctionCallNode } from './ast';

/** Shared lexical ownership for evaluation dependency/name traversal. */
export function visitLexicalArguments(
  node: FunctionCallNode,
  bound: ReadonlySet<string>,
  visit: (node: FormulaAst, bound: ReadonlySet<string>) => void,
): boolean {
  const id = node.name.toUpperCase();
  if (id === 'LET' && node.arguments.length >= 3 && node.arguments.length % 2 === 1) {
    const local = new Set(bound);
    for (let index = 0; index < node.arguments.length - 1; index += 2) {
      visit(node.arguments[index + 1]!, local);
      const name = node.arguments[index]!;
      if (name.type === 'name-reference') local.add(name.name.toUpperCase());
    }
    visit(node.arguments[node.arguments.length - 1]!, local);
    return true;
  }
  if (id === 'LAMBDA' && node.arguments.length) {
    const local = new Set(bound);
    for (const parameter of node.arguments.slice(0, -1)) if (parameter.type === 'name-reference') local.add(parameter.name.toUpperCase());
    visit(node.arguments[node.arguments.length - 1]!, local);
    return true;
  }
  return false;
}
