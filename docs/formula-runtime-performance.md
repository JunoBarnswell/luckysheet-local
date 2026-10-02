# Formula runtime: catalog coverage and scaling

## Scope and ownership

The executable registry contains 136 functions. `FUNCTION_DESCRIPTORS`, the UI
function library, argument contracts and the acceptance corpus must have the
same IDs. Each ID has an independently specified successful example, inline and
structured-clone Worker execution, and excessive-argument rejection. This is
coverage of the current registry, not certification of every Excel function or
every optional argument. Unimplemented functions return `#NAME?`; the former
ROMAN placeholder has been removed.

WorkbookModel remains the authored owner. Command → Java commit → ACK replay →
FormulaEngine → Worker result → Canvas remains the calculation path. The newer
linked-calculation migration and protocol are described in
[linked-calculation-architecture.md](linked-calculation-architecture.md); this
formula optimization adds no independent persisted value store or migration.

## Reference consumption and calculation

- `FormulaRangeView` supplies dimensions, coordinate reads and sparse occupied
  entries from canonical inputs, explicit scenario overrides and spill children.
  It does not copy a range into another value store.
- Conditional aggregates, COUNTBLANK, order statistics and SUMPRODUCT consume
  sparse entries. Counts account for implicit blanks analytically. SUMIF and
  AVERAGEIF project their target from its top-left cell and index the actual
  projected read geometry. Shape mismatches and permission errors are observable.
- INDEX keeps selected reference geometry; ROWS/COLUMNS consume dimensions.
  Lookup functions read a virtual vector and only the selected return coordinates.
  Binary search modes perform logarithmic reads and require a correctly sorted
  input; they do not first allocate or validate the entire vector.
- SUBTOTAL and AGGREGATE share the canonical reference/visibility cursor and
  streaming aggregates. Duplicate conditional and lookup implementations were
  removed. AGGREGATE's function/options/k arguments and array padding are explicit.
- Dependency graph traversal and evaluation order are iterative. Clean
  prerequisites retain their current results during affected-subgraph calculation.
  A body edit reuses topology only when both reference ownership and value
  dependencies are unchanged. Shape queries retain calculated control dependencies
  and can recover from a formerly circular value expression.
- IF, IFERROR, IFNA and scalar CHOOSE evaluate selected branches. Scalar functions
  broadcast compatible arrays. Error values, incompatible shapes and nonfinite
  scalar builtin results remain typed errors.
- Volatile roots include their dependents; clearing a manual input retains its
  dirty address. SJS.TABLE scenario evaluation does not overwrite canonical
  formula results or authored input values.

Dense functions still consume dense arrays when the operation needs every
coordinate. Sparse reference consumption does not make FILTER, SORT, UNIQUE,
GROUPBY, PIVOTBY or every array expression constant-time. Generated/output arrays
retain the existing 100,000-cell limit; IFERROR may materialize its first operand.

## Reproducible benchmark

Run from `frontend-react`:

```sh
npm run benchmark:formulas
FORMULA_BENCHMARK_ROWS=1000000 FORMULA_BENCHMARK_SAMPLES=30 npm run benchmark:formulas
```

Node v24.19.0, this shared host. Baseline `b3f8cc65` and optimized `192f300a`
execute the same script, with 100 occupied rows spread over a 100,000-row range,
30 recalculations per case after initial calculation, and checked expected values.
Numbers below are milliseconds; p95 is an observed percentile of these samples.

| Consumer | Baseline p50 / p95 | Optimized p50 / p95 |
| --- | ---: | ---: |
| SUMIFS | 310.932 / 566.204 | 0.860 / 2.104 |
| INDEX | 96.243 / 119.023 | 0.456 / 0.927 |
| SUBTOTAL | 355.178 / 472.867 | 0.550 / 1.555 |

With a 1,000,000-row range and the same 100 occupied rows:

| Consumer | Optimized p50 / p95 |
| --- | ---: |
| SUMIFS | 1.082 / 3.334 |
| INDEX | 0.367 / 1.115 |
| SUBTOTAL | 0.850 / 2.495 |

The 100,000-element binary lookup reads 99,999 elements at baseline and 18 after
optimization; the million-element case reads 21. A reverse-address chain of
6,000 cells returns `#VALUE!` at baseline and the correct result 6,000 after the
change. Its baseline timing is not a valid successful-calculation comparison.

Heap/RSS deltas are emitted for investigation; garbage collection can make them
negative and they are not an allocation budget. These are sparse-range local
measurements, not dense-million-record or production-hardware certification.

## Acceptance and remaining limits

`npm run test:formula-engine` covers every engine test, and the CI calculation
gate includes them plus application and structural/collaboration integration.
Catalog additions cannot silently lack a contract, UI entry or acceptance vector.
Scaling tests check that sparse consumers do not invoke matrix materialization,
that binary lookup stays within logarithmic reads, and that deep chains and
circular rejection/recovery preserve correct values.

Real browser acceptance uses Chromium, Spring Boot Java 21 and file-backed H2.
Authored inputs are entered through the real formula bar and committed to Java.
Read-only observers collect actual Worker messages and Canvas text; calculation
results are never injected. The corpus includes volatile invariants, array shapes,
five typed rejection cases, million-coordinate sparse ranges, source edits and
reloading the persisted workbook. Raw evidence lives outside the repository in
`/workspace/formula-pr346-*`.

Native desktop Excel acceptance is **Blocked** because it is unavailable. Many
Excel functions, locale-specific formatting/date syntax and complete optional
argument behavior remain outside demonstrated acceptance. The broader frontend
unit suite remains red; its failures must be recorded separately from the green
focused calculation gate. Linked sources and Record field limitations are owned
by the linked-calculation document, not certified by the worksheet corpus.

## Rollback

Revert the formula optimization commits without downgrading the separately
merged linked-calculation snapshot/Worker contracts. Reload clients to discard
derived Worker state. Rolling back linked-data Flyway V13 requires the matched
database backup and binaries described in its migration document.
