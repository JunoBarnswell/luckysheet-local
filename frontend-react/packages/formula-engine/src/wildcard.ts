/** Excel wildcard matching without a regular expression or exponential backtracking. */
export function compileExcelWildcard(pattern: string): (value: string) => boolean {
  const tokens: Array<{ kind: 'literal' | 'one' | 'many'; value?: string }> = [];
  if (pattern.length > 32767) throw new Error('UNSUPPORTED_FEATURE: Wildcard pattern exceeds the text budget');
  const chars = Array.from(pattern.toLocaleLowerCase('en-US'));
  for (let index = 0; index < chars.length; index++) {
    const char = chars[index]!;
    if (char === '~' && index + 1 < chars.length) tokens.push({ kind: 'literal', value: chars[++index]! });
    else if (char === '*') { if (tokens.at(-1)?.kind !== 'many') tokens.push({ kind: 'many' }); }
    else if (char === '?') tokens.push({ kind: 'one' });
    else tokens.push({ kind: 'literal', value: char });
  }
  return (value) => {
    if (value.length > 32767) throw new Error('UNSUPPORTED_FEATURE: Wildcard value exceeds the text budget');
    const input = Array.from(value.toLocaleLowerCase('en-US'));
    let at = 0; let token = 0; let star = -1; let retry = 0; let work = 0;
    while (at < input.length) {
      if (++work > 1000000) throw new Error('UNSUPPORTED_FEATURE: Wildcard matching budget exceeded');
      const next = tokens[token];
      if (next?.kind === 'one' || (next?.kind === 'literal' && next.value === input[at])) { at++; token++; }
      else if (next?.kind === 'many') { star = token++; retry = at; }
      else if (star >= 0) { token = star + 1; at = ++retry; }
      else return false;
    }
    while (tokens[token]?.kind === 'many') token++;
    return token === tokens.length;
  };
}
