import { readFile } from 'node:fs/promises';

const corpus = JSON.parse(await readFile('frontend-react/packages/formula-engine/migration-corpus.json', 'utf8'));
if (corpus.status !== 'requires-rust-parity') throw new Error('Formula migration corpus must remain explicitly gated until Rust parity is proven');
console.error('BLOCKED: formula migration corpus requires Rust parity coverage; TypeScript-only tests cannot pass this gate.');
process.exitCode = 2;
