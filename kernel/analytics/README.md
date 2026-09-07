# Canonical analytics execution

`AnalyticsRuntime` is owned by one workbook host instance. Call
`runtime.execute(request, &reader, &cancel_flag)` to retain completed columnar
source indexes across layout/filter changes. The reader revision is pinned both
before input scanning and before result publication. Cancellation leaves valid
source caches intact. Never share a runtime across workbook identities.

The stateless `execute` and `execute_with_cancel` entry points are available for
one-shot callers. On WASM, provide `env.analytics_now_ms: () => number` using the
host monotonic clock. Full-sheet tasks belong to the server task scheduler;
synchronous WASM invocation is not a browser cancellation mechanism.

## Data flow

Page reader → numeric/boolean vectors or dictionary codes → filter masks →
row/column member ids → one sparse hash aggregate → merged ancestor states →
ordered axis indexes → viewport cells / paginated drilldown.

`PivotRequest` uses body-range absolute column addresses and independent
`valueId` identities. Its row/column fields accept grouping, sorting and custom
subtotal definitions. `includeRowTotals` and `includeColumnTotals` are independent.
Percentages are fractional numeric results; formatting belongs to the document.
`PivotResult.rows[].rowId`, `columns[].columnId` and sparse cell coordinates are
global result indexes, including when a viewport has a nonzero offset. Drilldown
returns only the requested source-row page.

`QueryRequest` currently accepts worksheet ranges with explicit result limits,
filter criteria, inner/left joins, grouping, aggregate output columns, stable
sorting and projection. Join output column identities must not overlap existing
columns. Input row paths retain numeric row ids only; cell values are materialized
only for a returned query page. Group projections reject non-grouped fields.

## Verification and remaining integration

The Rust tests include a one-million-row / three-column scan, sparse result and
viewport cardinality, paginated drilldown, typed aggregation, repeated Values
placements, Top-N total recomputation, filter domains, join multiplicity,
grouped projection rejection, cache reuse and cancellation at 256-cell scan
checkpoints. These tests do not constitute desktop Excel or browser acceptance.

The following remain integration work and must not be reported as completed:

- The frontend legacy `engine.ts` result tree / full projection contract and
  session source-registration path are not yet removed. The current frontend
  sparse-to-tree translator still expands result column cells.
- Full browser/server task transport, native concurrent cancellation plumbing,
  task proof publication, and external connector batch ingestion are owned by
  their host layers and are not implemented by this crate's range-query API.
- Query external spill sort and stream ingestion are not implemented. Working
  memory is budget checked; it does not pretend to spill.
- Calculated-field expression parsing is implemented and tested in its module,
  but calculated fields/items are not wired into the pivot aggregate planner.
- Per-page incremental source reuse, persisted collation behavior, all subtotal
  layout modes, conditional-format-derived filter visuals, independent slicer
  domain pagination and native Excel interoperability require closure.
