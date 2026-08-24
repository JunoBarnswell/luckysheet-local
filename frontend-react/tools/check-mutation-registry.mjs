import { readFile, readdir } from 'node:fs/promises';
import { join, relative, resolve } from 'node:path';

/**
 * Static mutation-contract gate.
 *
 * Every production mutation registration must carry one explicit contract:
 * schema, permission, affected-range resolver, and inverse policy. Runtime
 * validation remains authoritative for dynamically constructed payloads; this
 * gate refuses dynamic registration/reference syntax so a feature cannot hide
 * a missing contract behind a helper or computed id.
 */

const root = resolve(process.argv[2] ?? process.cwd());
const sourceRoot = join(root, 'packages');

async function walk(directory) {
  const entries = await readdir(directory, { withFileTypes: true });
  const files = [];
  for (const entry of entries) {
    if (entry.name === 'node_modules' || entry.name === 'dist') continue;
    const target = join(directory, entry.name);
    if (entry.isDirectory()) files.push(...(await walk(target)));
    else if (/\.(?:ts|tsx|js|mjs)$/.test(entry.name)) files.push(target);
  }
  return files;
}

function isTestFile(file) {
  return /(?:^|[\\/])(?:[^\\/]+\.)?(?:test|spec)\.[^\\/]+$/i.test(file)
    || /(?:^|[\\/])(?:test|tests|__tests__)(?:[\\/]|$)/i.test(file);
}

function skipString(source, index) {
  const quote = source[index];
  let cursor = index + 1;
  while (cursor < source.length) {
    if (source[cursor] === '\\') {
      cursor += 2;
      continue;
    }
    if (source[cursor] === quote) return cursor + 1;
    cursor += 1;
  }
  return source.length;
}

function skipComment(source, index) {
  if (source[index + 1] === '/') {
    const end = source.indexOf('\n', index + 2);
    return end < 0 ? source.length : end + 1;
  }
  const end = source.indexOf('*/', index + 2);
  return end < 0 ? source.length : end + 2;
}

function findMatching(source, start, opening, closing) {
  let depth = 0;
  for (let index = start; index < source.length; index += 1) {
    const character = source[index];
    if (character === '"' || character === "'" || character === '`') {
      index = skipString(source, index) - 1;
      continue;
    }
    if (character === '/' && (source[index + 1] === '/' || source[index + 1] === '*')) {
      index = skipComment(source, index) - 1;
      continue;
    }
    if (character === opening) depth += 1;
    else if (character === closing) {
      depth -= 1;
      if (depth === 0) return index;
    }
  }
  return -1;
}

function lineNumber(source, index) {
  return source.slice(0, index).split('\n').length;
}

function firstLiteralId(object) {
  const match = object.match(/\bid\s*:\s*(['"])([^'"]+)\1/);
  return match ? match[2] : undefined;
}

function splitTopLevelArguments(source) {
  const args = [];
  let start = 0;
  let braces = 0;
  let brackets = 0;
  let parentheses = 0;
  for (let index = 0; index < source.length; index += 1) {
    const character = source[index];
    if (character === '"' || character === "'" || character === '`') {
      index = skipString(source, index) - 1;
      continue;
    }
    if (character === '{') braces += 1;
    else if (character === '}') braces -= 1;
    else if (character === '[') brackets += 1;
    else if (character === ']') brackets -= 1;
    else if (character === '(') parentheses += 1;
    else if (character === ')') parentheses -= 1;
    else if (character === ',' && braces === 0 && brackets === 0 && parentheses === 0) {
      args.push(source.slice(start, index));
      start = index + 1;
    }
  }
  args.push(source.slice(start));
  return args;
}

function metadataObject(call) {
  const metadataMatch = /\bmetadata\s*:\s*\{/g.exec(call.arguments);
  if (metadataMatch) {
    const objectStart = metadataMatch.index + metadataMatch[0].lastIndexOf('{');
    const objectEnd = findMatching(call.arguments, objectStart, '{', '}');
    return objectEnd < 0 ? undefined : call.arguments.slice(objectStart, objectEnd + 1);
  }
  const args = splitTopLevelArguments(call.arguments);
  const candidate = args[2]?.trimStart();
  if (!candidate || candidate[0] !== '{') return undefined;
  const objectEnd = findMatching(candidate, 0, '{', '}');
  return objectEnd < 0 ? undefined : candidate.slice(0, objectEnd + 1);
}

function hasMetadataField(metadata, field) {
  return metadata !== undefined && new RegExp(`\\b${field}\\s*:`).test(metadata);
}

function parseRegistrationCalls(source) {
  const calls = [];
  const marker = /\b(registerMutationHandler|registerMutation)\s*(?:<[^>]*>)?\s*\(/g;
  for (const match of source.matchAll(marker)) {
    const openIndex = (match.index ?? 0) + match[0].lastIndexOf('(');
    const closeIndex = findMatching(source, openIndex, '(', ')');
    if (closeIndex < 0) continue;
    const args = source.slice(openIndex + 1, closeIndex);
    calls.push({
      kind: match[1],
      arguments: args,
      index: match.index ?? 0,
      line: lineNumber(source, match.index ?? 0),
    });
  }
  return calls;
}

function parseMutationCalls(source) {
  const calls = [];
  const marker = /context\.applyMutation\s*\(\s*\{/g;
  for (const match of source.matchAll(marker)) {
    const objectStart = (match.index ?? 0) + match[0].lastIndexOf('{');
    const objectEnd = findMatching(source, objectStart, '{', '}');
    if (objectEnd < 0) continue;
    const object = source.slice(objectStart, objectEnd + 1);
    const inverseProperty = /\binverse\s*:\s*/.exec(object);
    const primaryRegion = inverseProperty ? object.slice(0, inverseProperty.index) : object;
    const primary = primaryRegion.match(/\bid\s*:\s*(['"])([^'"]+)\1/);
    const dynamicPrimary = primary === null && /\bid\s*:\s*(?!['"])[^,}\n]+/.test(primaryRegion);
    const inverseIds = [];
    const dynamicInverse = [];
    if (inverseProperty) {
      const valueStart = inverseProperty.index + inverseProperty[0].length;
      let arrayStart = valueStart;
      while (arrayStart < object.length && /\s/.test(object[arrayStart])) arrayStart += 1;
      if (object[arrayStart] === '[') {
        const arrayEnd = findMatching(object, arrayStart, '[', ']');
        if (arrayEnd >= 0) {
          const inverseValue = object.slice(arrayStart, arrayEnd + 1);
          for (let index = 1; index < inverseValue.length - 1; index += 1) {
            while (index < inverseValue.length - 1 && /[\s,]/.test(inverseValue[index])) index += 1;
            if (inverseValue[index] !== '{') {
              break;
            }
            const entryEnd = findMatching(inverseValue, index, '{', '}');
            if (entryEnd < 0) {
              dynamicInverse.push(inverseValue);
              break;
            }
            const entry = inverseValue.slice(index, entryEnd + 1);
            const literal = entry.match(/^\{\s*id\s*:\s*(['"])([^'"]+)\1/);
            if (literal) inverseIds.push(literal[2]);
            else dynamicInverse.push(entry);
            index = entryEnd;
          }
        }
      }
    }
    calls.push({
      primaryId: primary?.[2],
      dynamicPrimary,
      inverseIds,
      dynamicInverse,
      index: match.index ?? 0,
      line: lineNumber(source, match.index ?? 0),
    });
  }
  return calls;
}

const files = await walk(sourceRoot);
const registered = new Map();
const references = [];
const violations = [];

for (const file of files) {
  const relPath = relative(root, file).replaceAll('\\', '/');
  const source = await readFile(file, 'utf8');
  if (isTestFile(relPath)) {
    for (const call of parseRegistrationCalls(source)) {
      if (call.kind !== 'registerMutation' || call.arguments.trimStart()[0] === '{') continue;
      violations.push(`${relPath}:${call.line}: registerMutation must use the canonical object contract`);
    }
    continue;
  }
  if (relPath === 'packages/command-runtime/src/index.ts') continue;

  for (const call of parseRegistrationCalls(source)) {
    const location = `${relPath}:${call.line}`;
    if (call.kind === 'registerMutationHandler') {
      violations.push(`${location}: registerMutationHandler is not a canonical mutation registration`);
      continue;
    }
    const first = call.arguments.trimStart();
    const firstArgument = first[0];
    const id = firstArgument === '{' ? firstLiteralId(first) : first.match(/^(['"])([^'"]+)\1/)?.[2];
    if (!id) {
      violations.push(`${location}: mutation registration id must be a literal canonical id`);
      continue;
    }
    const metadata = metadataObject(call);
    for (const field of ['schema', 'permission', 'affectedRanges']) {
      if (!hasMetadataField(metadata, field)) violations.push(`${location}: mutation ${id} missing metadata.${field}`);
    }
    if (!hasMetadataField(metadata, 'inversePolicy') && !hasMetadataField(metadata, 'inverseIds')) {
      violations.push(`${location}: mutation ${id} missing metadata.inversePolicy`);
    }
    const owners = registered.get(id) ?? [];
    owners.push(location);
    registered.set(id, owners);
  }

  for (const call of parseMutationCalls(source)) {
    if (call.dynamicPrimary) violations.push(`${relPath}:${call.line}: applyMutation id must be a literal registered mutation`);
    else references.push({ id: call.primaryId, kind: 'applyMutation', file: relPath, line: call.line });
    for (const id of call.inverseIds) references.push({ id, kind: 'inverse', file: relPath, line: call.line });
    for (const _dynamic of call.dynamicInverse) violations.push(`${relPath}:${call.line}: inverse mutation id must be a literal registered mutation`);
  }
}

for (const [id, owners] of registered) {
  if (owners.length > 1) violations.push(`mutation ${id} has duplicate registrations: ${owners.join(', ')}`);
}
for (const reference of references) {
  if (!registered.has(reference.id)) {
    violations.push(`${reference.file}:${reference.line}: ${reference.kind} references unregistered mutation "${reference.id}"`);
  }
}

if (violations.length > 0) {
  console.error(violations.join('\n'));
  process.exitCode = 1;
} else {
  console.log(`Mutation registry contract gate passed: ${registered.size} registrations, ${references.length} mutation references.`);
}
