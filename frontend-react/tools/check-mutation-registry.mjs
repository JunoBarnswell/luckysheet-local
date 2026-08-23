import { readFile, readdir } from 'node:fs/promises';
import { join, relative, resolve } from 'node:path';

/**
 * Static half of the mutation registry gate.
 *
 * Runtime validation in command-runtime is authoritative for dynamic helper
 * calls. This gate covers every literal mutation id in source so an inverse or
 * direct `context.applyMutation({ id: ... })` cannot silently drift away from
 * the registered handler set. Dynamic ids are intentionally reported as
 * informational because generic feature helpers resolve them at runtime.
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

function literalIds(source, pattern) {
  const result = [];
  for (const match of source.matchAll(pattern)) {
    result.push({ id: match[2], index: match.index ?? 0 });
  }
  return result;
}

function mutationCalls(source) {
  const calls = [];
  const marker = /context\.applyMutation\s*\(\s*\{/g;
  for (const match of source.matchAll(marker)) {
    const objectStart = (match.index ?? 0) + match[0].lastIndexOf('{');
    const objectEnd = findMatching(source, objectStart, '{', '}');
    if (objectEnd < 0) continue;
    const object = source.slice(objectStart, objectEnd + 1);
    const primary = object.match(/\bid\s*:\s*(['"])([^'"]+)\1/);
    if (!primary) continue;
    const inverseProperty = /\binverse\s*:\s*/.exec(object);
    const inverseIds = [];
    if (inverseProperty) {
      const valueStart = inverseProperty.index + inverseProperty[0].length;
      let arrayStart = valueStart;
      while (arrayStart < object.length && /\s/.test(object[arrayStart])) arrayStart += 1;
      if (object[arrayStart] === '[') {
        const arrayEnd = findMatching(object, arrayStart, '[', ']');
        if (arrayEnd >= 0) {
          const inverseValue = object.slice(arrayStart, arrayEnd + 1);
          for (const inverse of inverseValue.matchAll(/\bid\s*:\s*(['"])([^'"]+)\1/g)) {
            inverseIds.push(inverse[2]);
          }
        }
      }
    }
    calls.push({ primaryId: primary[2], inverseIds, index: match.index ?? 0 });
  }
  return calls;
}

function isTestFile(file) {
  return /(?:^|[\\/])(?:[^\\/]+\.)?(?:test|spec)\.[^\\/]+$/i.test(file) || /(?:^|[\\/])(?:test|tests|__tests__)(?:[\\/]|$)/i.test(file);
}

const files = await walk(sourceRoot);
const registered = new Map();
const references = [];
const dynamic = [];

for (const file of files) {
  const source = await readFile(file, 'utf8');
  const relPath = relative(root, file).replaceAll('\\', '/');
  if (isTestFile(relPath)) continue;

  for (const match of literalIds(source, /(?:registerMutationHandler|registerMutation)\s*(?:<[^;()]*>)?\s*\(\s*(['"])([^'"]+)\1/g)) {
    const owners = registered.get(match.id) ?? [];
    owners.push(relPath);
    registered.set(match.id, owners);
  }

  if (/(?:registerMutationHandler|registerMutation)\s*(?:<[^;()]*>)?\s*\(\s*[A-Za-z_$]/.test(source)) {
    dynamic.push(`${relPath}: dynamic mutation registration requires runtime validation`);
  }

  for (const call of mutationCalls(source)) {
    references.push({ id: call.primaryId, kind: 'applyMutation', file: relPath });
    for (const inverseId of call.inverseIds) {
      references.push({ id: inverseId, kind: 'inverse', file: relPath });
    }
  }
}

const missing = references.filter((reference) => !registered.has(reference.id));
if (missing.length > 0) {
  const unique = new Map();
  for (const entry of missing) {
    const key = `${entry.kind}:${entry.id}`;
    if (!unique.has(key)) unique.set(key, entry);
  }
  for (const entry of unique.values()) {
    console.error(`${entry.file}: ${entry.kind} references unregistered mutation "${entry.id}"`);
  }
  process.exitCode = 1;
} else {
  console.log(`Mutation registry literal gate passed: ${registered.size} registrations, ${references.length} mutation references.`);
}

if (dynamic.length > 0 && process.env.MUTATION_REGISTRY_STRICT_DYNAMIC === '1') {
  console.error(dynamic.join('\n'));
  process.exitCode = 1;
}
