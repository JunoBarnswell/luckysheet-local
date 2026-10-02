import { performance } from 'node:perf_hooks';
import { FormulaEngine, isFormulaError, type CalculationInputUpdate } from '../packages/formula-engine/src/index';
import { findLookupIndex } from '../packages/formula-engine/src/functions/lookup-engine';

const size = Number(process.env.FORMULA_BENCHMARK_ROWS ?? 100_000);
const samples = Number(process.env.FORMULA_BENCHMARK_SAMPLES ?? 30);
if (!Number.isSafeInteger(size) || size < 100 || !Number.isSafeInteger(samples) || samples < 5) throw new Error('Invalid benchmark dimensions');
const percentile = (values: number[], quantile: number) => [...values].sort((a, b) => a - b)[Math.ceil(values.length * quantile) - 1];
const engine = new FormulaEngine();
engine.setSpillEnvironment('Sheet1', { rowCount: size, columnCount: 20, isOccupied: () => false });
const inputs: CalculationInputUpdate[] = [];
for (let row = 0; row < 100; row++) {
  inputs.push({ address: { sheetId: 'Sheet1', row: row * Math.floor(size / 100), column: 0 }, input: { kind: 'value', value: 'group' } });
  inputs.push({ address: { sheetId: 'Sheet1', row: row * Math.floor(size / 100), column: 1 }, input: { kind: 'value', value: row + 1 } });
}
engine.synchronizeInputs(inputs);
const timings: Record<string, unknown> = {};
for (const [name, formula, expected] of [
  ['conditional', `=SUMIFS(B1:B${size},A1:A${size},"group")`, 5050],
  ['index', `=INDEX(B1:B${size},1)`, 1],
  ['subtotal', `=SUBTOTAL(9,B1:B${size})`, 5050],
] as const) {
  engine.setFormula('D1', formula);
  const before = process.memoryUsage();
  const elapsed: number[] = [];
  for (let pass = 0; pass < samples; pass++) {
    const start = performance.now();
    engine.recalculateCell('D1');
    elapsed.push(performance.now() - start);
    const value = engine.getCellValue('D1');
    if (value !== expected) throw new Error(`${name}: expected ${expected}; received ${JSON.stringify(value)}`);
  }
  const after = process.memoryUsage();
  timings[name] = { p50Ms: percentile(elapsed, 0.5), p95Ms: percentile(elapsed, 0.95), heapDeltaBytes: after.heapUsed - before.heapUsed, rssDeltaBytes: after.rss - before.rss };
}
const vector = Array.from({ length: size }, (_, index) => index + 1);
let reads = 0;
const observed = new Proxy(vector, { get(target, key, receiver) { if (typeof key === 'string' && /^\d+$/.test(key)) reads++; return Reflect.get(target, key, receiver); } });
const lookupStart = performance.now();
if (findLookupIndex(size - 1, observed, 0, 2) !== size - 2) throw new Error('Binary lookup produced the wrong position');
timings.binaryLookup = { elapsedMs: performance.now() - lookupStart, elementReads: reads };
const chain = new FormulaEngine();
const chainLength = 6000;
chain.synchronizeInputs(Array.from({ length: chainLength }, (_, row) => ({ address: { sheetId: 'Sheet1', row, column: 0 }, input: row === chainLength - 1 ? { kind: 'value' as const, value: 1 } : { kind: 'formula' as const, formula: `=A${row + 2}+1` } })));
const chainStart = performance.now();
chain.recalculate();
const chainValue = chain.getCellValue('A1');
timings.deepChain = { elapsedMs: performance.now() - chainStart, result: isFormulaError(chainValue) ? chainValue.code : chainValue, expected: chainLength };
process.stdout.write(`${JSON.stringify({ node: process.version, rows: size, occupiedRows: 100, samples, timings }, null, 2)}\n`);
