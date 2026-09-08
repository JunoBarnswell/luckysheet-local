import { readFile } from 'node:fs/promises';
import { dirname, resolve } from 'node:path';
import { fileURLToPath } from 'node:url';

const root = resolve(dirname(fileURLToPath(import.meta.url)), '..');
const corpusPath = resolve(root, 'frontend-react/packages/formula-engine/migration-corpus.json');
const rustTestPath = resolve(root, 'kernel/formula/tests/migration_corpus.rs');
const bindingPath = resolve(root, 'frontend-react/packages/formula-engine/src/formula-engine.ts');
const corpus = JSON.parse(await readFile(corpusPath, 'utf8'));
const rustTests = await readFile(rustTestPath, 'utf8');
const binding = await readFile(bindingPath, 'utf8');

if (corpus.status !== 'rust-parity') {
  throw new Error('Formula migration corpus must be accepted only after canonical Rust parity');
}
if (!Array.isArray(corpus.files) || corpus.files.length !== 11) {
  throw new Error('Formula migration corpus must retain all 11 archived source groups');
}
const cases = corpus.files.flatMap(file => {
  if (typeof file.sourceText !== 'string' || file.sourceText.length === 0) {
    throw new Error('Formula migration evidence is missing for ' + file.source);
  }
  return file.cases;
});
if (cases.length !== 56) {
  throw new Error('Formula migration corpus must retain 56 semantic cases, found ' + cases.length);
}
const declaredTests = new Set(
  [...rustTests.matchAll(/^fn\s+(migration_[a-z0-9_]+)\s*\(/gm)].map(match => match[1]),
);
for (const testCase of cases) {
  if (testCase.status !== 'rust-parity' || typeof testCase.rustTest !== 'string') {
    throw new Error('Formula case is not mapped to Rust parity: ' + testCase.name);
  }
  if (!declaredTests.has(testCase.rustTest)) {
    throw new Error('Formula case maps to a missing Rust test: ' + testCase.rustTest);
  }
}
if (!binding.includes('kernelInvoke') || !binding.includes("'formula.evaluate'")) {
  throw new Error('FormulaEngine must invoke the canonical kernel');
}
for (const removed of [
  'setValue(',
  'setFormula(',
  'setDefinedNames(',
  'setSheetTables(',
  'createCalculationSessionPort(',
  'consumeCalculationSession',
  'installBrowserCalculationWorkerEntry',
]) {
  if (binding.includes(removed)) {
    throw new Error('FormulaEngine still exposes removed TypeScript runtime behavior: ' + removed);
  }
}

console.log('Formula parity gate passed: ' + cases.length + ' cases mapped to ' + declaredTests.size + ' Rust integration tests.');
