/**
 * Converts Excel/WPS namespace spellings into the Formula Engine's canonical
 * dialect at the OOXML boundary.  The scanner deliberately skips string
 * literals, so text such as `"_xlfn.FILTER"` remains user data.
 *
 * `@` is already a first-class unary operator in the canonical AST and is
 * therefore retained rather than erased.  `_xlfn.SINGLE(...)` is its older
 * OOXML spelling and becomes `@(...)`.
 */
export function normalizeExternalFormula(source: string): string {
  let normalized = '';
  let quoted = false;
  for (let index = 0; index < source.length;) {
    const character = source[index]!;
    if (character === '"') {
      normalized += character;
      if (quoted && source[index + 1] === '"') {
        normalized += source[index + 1];
        index += 2;
        continue;
      }
      quoted = !quoted;
      index += 1;
      continue;
    }
    if (quoted) {
      normalized += character;
      index += 1;
      continue;
    }
    const remaining = source.slice(index);
    const single = /^_xlfn\.SINGLE\b/i.exec(remaining);
    if (single) {
      normalized += '@';
      index += single[0].length;
      continue;
    }
    const prefix = /^_(?:xlfn|xlws)\./i.exec(remaining);
    if (prefix) {
      index += prefix[0].length;
      continue;
    }
    normalized += character;
    index += 1;
  }
  return normalized;
}
