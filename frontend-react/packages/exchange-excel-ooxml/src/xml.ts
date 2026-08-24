/**
 * A small namespace-tolerant XML reader used by the OOXML package layer.
 *
 * It intentionally parses XML as a tree instead of using regular expressions:
 * OOXML permits namespace prefixes, reordered attributes, escaped text and
 * rich shared strings.  The parser is not a general-purpose HTML parser and
 * rejects malformed XML rather than guessing at a workbook's contents.
 */

export interface XmlNode {
  /** Original qualified name, e.g. `x:sheet` or `sheet`. */
  name: string;
  attrs: Record<string, string>;
  children: XmlNode[];
  text: string;
}

export function localName(name: string): string {
  const separator = name.indexOf(':');
  return separator >= 0 ? name.slice(separator + 1) : name;
}

export function parseXml(source: string): XmlNode {
  const root: XmlNode = { name: '#document', attrs: {}, children: [], text: '' };
  const stack: XmlNode[] = [root];
  let index = 0;

  while (index < source.length) {
    const open = source.indexOf('<', index);
    if (open < 0) {
      appendText(stack[stack.length - 1]!, source.slice(index));
      break;
    }
    if (open > index) appendText(stack[stack.length - 1]!, source.slice(index, open));
    index = open;

    if (source.startsWith('<!--', index)) {
      const end = source.indexOf('-->', index + 4);
      if (end < 0) throw new Error('Malformed XML: unterminated comment');
      index = end + 3;
      continue;
    }
    if (source.startsWith('<![CDATA[', index)) {
      const end = source.indexOf(']]>', index + 9);
      if (end < 0) throw new Error('Malformed XML: unterminated CDATA');
      appendText(stack[stack.length - 1]!, source.slice(index + 9, end));
      index = end + 3;
      continue;
    }
    if (source.startsWith('<?', index)) {
      const end = source.indexOf('?>', index + 2);
      if (end < 0) throw new Error('Malformed XML: unterminated processing instruction');
      index = end + 2;
      continue;
    }
    if (source.startsWith('<!', index)) {
      const end = findTagEnd(source, index + 2);
      if (end < 0) throw new Error('Malformed XML: unterminated declaration');
      index = end + 1;
      continue;
    }

    const tagEnd = findTagEnd(source, index + 1);
    if (tagEnd < 0) throw new Error('Malformed XML: unterminated tag');
    let body = source.slice(index + 1, tagEnd).trim();
    if (body.startsWith('/')) {
      const name = body.slice(1).trim();
      const current = stack.pop();
      if (!current || current.name !== name) {
        throw new Error(`Malformed XML: closing ${name} does not match ${current?.name ?? 'root'}`);
      }
    } else {
      const selfClosing = body.endsWith('/');
      if (selfClosing) body = body.slice(0, -1).trim();
      const { name, attrs } = parseStartTag(body);
      const node: XmlNode = { name, attrs, children: [], text: '' };
      stack[stack.length - 1]!.children.push(node);
      if (!selfClosing) stack.push(node);
    }
    index = tagEnd + 1;
  }

  if (stack.length !== 1) throw new Error(`Malformed XML: unclosed ${stack[stack.length - 1]!.name}`);
  return root;
}

function findTagEnd(source: string, start: number): number {
  let quote = '';
  for (let index = start; index < source.length; index += 1) {
    const char = source[index]!;
    if (quote) {
      if (char === quote) quote = '';
    } else if (char === '"' || char === "'") {
      quote = char;
    } else if (char === '>') {
      return index;
    }
  }
  return -1;
}

function parseStartTag(body: string): { name: string; attrs: Record<string, string> } {
  let index = 0;
  while (index < body.length && !isWhitespace(body[index]!)) index += 1;
  const name = body.slice(0, index);
  if (!name) throw new Error('Malformed XML: missing element name');
  const attrs: Record<string, string> = {};
  while (index < body.length) {
    while (index < body.length && isWhitespace(body[index]!)) index += 1;
    if (index >= body.length) break;
    const start = index;
    while (index < body.length && !isWhitespace(body[index]!) && body[index] !== '=') index += 1;
    const attrName = body.slice(start, index);
    while (index < body.length && isWhitespace(body[index]!)) index += 1;
    if (body[index] !== '=') throw new Error(`Malformed XML: attribute ${attrName} has no value`);
    index += 1;
    while (index < body.length && isWhitespace(body[index]!)) index += 1;
    const quote = body[index];
    if (quote !== '"' && quote !== "'") throw new Error(`Malformed XML: attribute ${attrName} is not quoted`);
    index += 1;
    const valueStart = index;
    while (index < body.length && body[index] !== quote) index += 1;
    if (index >= body.length) throw new Error(`Malformed XML: attribute ${attrName} is unterminated`);
    attrs[attrName] = decodeXmlEntities(body.slice(valueStart, index));
    index += 1;
  }
  return { name, attrs };
}

function appendText(node: XmlNode, raw: string): void {
  if (raw) node.text += decodeXmlEntities(raw);
}

function isWhitespace(value: string): boolean {
  return value === ' ' || value === '\t' || value === '\r' || value === '\n';
}

export function child(node: XmlNode | undefined, name: string): XmlNode | undefined {
  return node?.children.find((candidate) => localName(candidate.name) === name);
}

export function children(node: XmlNode | undefined, name: string): XmlNode[] {
  return node?.children.filter((candidate) => localName(candidate.name) === name) ?? [];
}

export function descendants(node: XmlNode | undefined, name: string): XmlNode[] {
  if (!node) return [];
  const result: XmlNode[] = [];
  for (const candidate of node.children) {
    if (localName(candidate.name) === name) result.push(candidate);
    result.push(...descendants(candidate, name));
  }
  return result;
}

export function textContent(node: XmlNode | undefined): string {
  if (!node) return '';
  return node.text + node.children.map(textContent).join('');
}

export function encodeXml(value: string): string {
  return value
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&apos;');
}

export function decodeXmlEntities(value: string): string {
  return value.replace(/&(#x[0-9a-f]+|#\d+|amp|lt|gt|quot|apos);/gi, (entity, code: string) => {
    if (code === 'amp') return '&';
    if (code === 'lt') return '<';
    if (code === 'gt') return '>';
    if (code === 'quot') return '"';
    if (code === 'apos') return "'";
    const number = code.toLowerCase().startsWith('#x')
      ? Number.parseInt(code.slice(2), 16)
      : Number.parseInt(code.slice(1), 10);
    return Number.isFinite(number) ? String.fromCodePoint(number) : entity;
  });
}

/** Serialize a parsed node when preserving an unknown XML child. */
export function serializeXml(node: XmlNode): string {
  if (node.name === '#document') return node.children.map(serializeXml).join('');
  const attrs = Object.entries(node.attrs).map(([key, value]) => ` ${key}="${encodeXml(value)}"`).join('');
  const content = `${encodeXml(node.text)}${node.children.map(serializeXml).join('')}`;
  return content ? `<${node.name}${attrs}>${content}</${node.name}>` : `<${node.name}${attrs}/>`;
}
