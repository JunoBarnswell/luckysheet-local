import type { DefinedNameModel, WorkbookTableModel } from '@react-sheets/core-model';
import { listFunctionLibrary } from '@react-sheets/formula-engine';
import type { CellEditCaret, FormulaAutocompleteCandidate } from './contracts';

interface TrieNode {
  children: Map<string, TrieNode>;
  candidateIndexes: number[];
}

export interface FormulaAutocompleteSource {
  revision: string;
  definedNames: readonly DefinedNameModel[];
  tables: readonly WorkbookTableModel[];
}

export interface FormulaAutocompleteQueryResult {
  candidates: readonly FormulaAutocompleteCandidate[];
  replacementSpan: CellEditCaret;
  hint?: { functionName: string; argumentIndex: number };
}

function createNode(): TrieNode {
  return { children: new Map(), candidateIndexes: [] };
}

function candidateKey(candidate: FormulaAutocompleteCandidate): string {
  return candidate.label.toLocaleUpperCase('en-US');
}

function formulaCandidates(source: FormulaAutocompleteSource): FormulaAutocompleteCandidate[] {
  const candidates: FormulaAutocompleteCandidate[] = [];
  for (const entry of listFunctionLibrary()) {
    candidates.push({ id: `function:${entry.id}`, kind: 'function', label: entry.id, insertionText: `${entry.id}(`, detail: entry.category });
  }
  for (const name of source.definedNames) {
    candidates.push({ id: `name:${name.scope}:${name.sheetId ?? ''}:${name.name}`, kind: 'defined-name', label: name.name, insertionText: name.name, detail: name.formula });
  }
  for (const table of source.tables) {
    candidates.push({ id: `table:${table.id}`, kind: 'table', label: table.name, insertionText: table.name, detail: `${table.rowCount} rows` });
    for (const field of table.fields) {
      const reference = `${table.name}[${field.name}]`;
      candidates.push({ id: `table-column:${table.id}:${field.id}`, kind: 'table-column', label: reference, insertionText: reference, detail: field.type });
    }
  }
  return candidates.sort((left, right) => left.label.localeCompare(right.label));
}

function queryToken(draft: string, caret: number): { prefix: string; replacementSpan: CellEditCaret; hint?: { functionName: string; argumentIndex: number } } | null {
  if (!draft.startsWith('=')) return null;
  const before = draft.slice(0, caret);
  const structured = /([A-Za-z_\\][A-Za-z0-9_.\\]*)\[[^\]]*$/.exec(before);
  if (structured?.[0]) {
    return { prefix: structured[0], replacementSpan: { start: caret - structured[0].length, end: caret } };
  }
  const token = /([A-Za-z_\\][A-Za-z0-9_.\\]*)$/.exec(before)?.[1];
  if (!token) {
    const call = /([A-Za-z_][A-Za-z0-9_.]*)\(([^()]*)$/.exec(before);
    if (!call?.[1]) return null;
    return { prefix: '', replacementSpan: { start: caret, end: caret }, hint: { functionName: call[1].toUpperCase(), argumentIndex: (call[2]?.match(/,/g)?.length ?? 0) + 1 } };
  }
  return { prefix: token, replacementSpan: { start: caret - token.length, end: caret } };
}

export class FormulaAutocompleteIndex {
  private root: TrieNode = createNode();
  private candidates: readonly FormulaAutocompleteCandidate[] = [];
  private revision = '';

  get sourceRevision(): string {
    return this.revision;
  }

  rebuild(source: FormulaAutocompleteSource): void {
    const candidates = formulaCandidates(source);
    const root = createNode();
    for (let index = 0; index < candidates.length; index += 1) {
      let node = root;
      for (const character of candidateKey(candidates[index]!)) {
        const next = node.children.get(character) ?? createNode();
        node.children.set(character, next);
        node = next;
      }
      node.candidateIndexes.push(index);
    }
    this.root = root;
    this.candidates = candidates;
    this.revision = source.revision;
  }

  async rebuildAsync(source: FormulaAutocompleteSource, signal: AbortSignal): Promise<void> {
    const candidates = formulaCandidates(source);
    const root = createNode();
    for (let index = 0; index < candidates.length; index += 1) {
      if (signal.aborted) throw new DOMException('Formula autocomplete rebuild was aborted', 'AbortError');
      let node = root;
      for (const character of candidateKey(candidates[index]!)) {
        const next = node.children.get(character) ?? createNode();
        node.children.set(character, next);
        node = next;
      }
      node.candidateIndexes.push(index);
      if (index > 0 && index % 256 === 0) await Promise.resolve();
    }
    if (signal.aborted) throw new DOMException('Formula autocomplete rebuild was aborted', 'AbortError');
    this.root = root;
    this.candidates = candidates;
    this.revision = source.revision;
  }

  query(draft: string, caret: CellEditCaret, limit = 50): FormulaAutocompleteQueryResult | null {
    const token = queryToken(draft, Math.min(caret.start, caret.end));
    if (!token) return null;
    if (token.hint) return { candidates: [], replacementSpan: token.replacementSpan, hint: token.hint };
    let node = this.root;
    for (const character of token.prefix.toLocaleUpperCase('en-US')) {
      const next = node.children.get(character);
      if (!next) return { candidates: [], replacementSpan: token.replacementSpan, ...(token.hint ? { hint: token.hint } : {}) };
      node = next;
    }
    const indexes: number[] = [];
    const stack = [node];
    while (stack.length > 0 && indexes.length < limit) {
      const current = stack.pop()!;
      for (const index of current.candidateIndexes) {
        indexes.push(index);
        if (indexes.length >= limit) break;
      }
      if (indexes.length >= limit) break;
      const children = [...current.children.entries()].sort(([left], [right]) => right.localeCompare(left));
      for (const [, child] of children) stack.push(child);
    }
    return { candidates: indexes.map((index) => this.candidates[index]!), replacementSpan: token.replacementSpan, ...(token.hint ? { hint: token.hint } : {}) };
  }
}
