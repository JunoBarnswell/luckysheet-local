# Formula kernel integration contract

`editor` owns the lexer, Pratt parser, source AST, UTF-16 spans and absolute reference flags. `parser` lowers that source AST to revision-independent resolved expressions. `references` applies structural changes through the same source AST. There is no TypeScript evaluator or alternate reference scanner in this crate.

`FormulaRuntime` stores parsed authored formulas, indexed dependencies, dirty roots and computed results. Register every worksheet name/id before registering formulas. Page commits call `invalidate_range`; changed formula definitions call `set_formula`/`remove_formula`. Production page loading must not copy ordinary cells into `set_value`: that method provides explicit values for embedders and isolated evaluation, while page-backed workbooks use `CellReader`.

Point and rectangle invalidation query ordered point references and dyadic row intervals. Recalculation visits the dirty closure, merges dynamic reference/spill dependencies and publishes cached results only after all reads and scheduler checks succeed. Formula generation increments on successful nonempty recalculation. Spill child reads take a scalar directly from the cached array instead of cloning the entire array per child.

Excel errors are `Ok(FormulaValue::Scalar(Scalar::Error(...)))`. Missing pages, cancellation, stale revisions, unsupported capabilities and resource exhaustion remain `Err(KernelError)`; `IFERROR` cannot intercept these host failures. `CalculationServices` provides cancellation checkpoints, canonical visibility and external function execution. The default service rejects unavailable visibility/external contracts. Hosts must use the service-aware entrypoints for production calculation and inspection.

## Host entrypoints

- `register_sheet(name, id)`, scoped `define_name`, typed `define_table`, and `set_context(CalculationContext)`.
- `evaluate_with_services(formula, address, reader, services)`.
- `recalculate_with_services(reader, services)` returns changed native results; the framed host response should contain a summary, not millions of serialized values.
- `inspect_query_with_services(reader, query, services)`: projections `entries`, `spills`, `status`; default 512 entries, limit 1–4096, opaque keyset cursor, optional sheet and exact address.
- `trace_with_services(address, reader, services)` executes the root expression and records actual expression results.
- `spill_value_with_services(address, reader, services)` uses the same indexed spill owner and propagates failed reads.
- `evaluate_with_overrides_and_services(formula, address, reader, overrides, services)` invalidates the affected calculation closure in an isolated session without modifying committed cached results.
- `function_capabilities()` is derived from executable Rust module registries.

`editor::parse_source` supports the UI before a workbook is open. Hosts must call `editor::validate_editor` on deserialized ASTs before format/offset/remap/F4/reference operations. `format_editor` excludes the leading `=`. Explicit external-workbook and 3D references are currently rejected with `UNSUPPORTED_FEATURE`; their source bytes must remain preserved by the document owner.

## Evidence and remaining acceptance

Native verification includes the library/runtime suite and `tests/migration_corpus.rs`. The latter maps all 56 archived TypeScript semantic cases to 11 executable Rust integration groups. It covers persistent delta state, workbook/sheet names, external namespace normalization, trace, dynamic arrays and budgets, GROUPBY/PIVOTBY, SJS.TABLE, dependency/index behavior, partial recalculation, structured table selectors, and spill ownership. These are kernel tests, not an Excel producer interoperability claim.

The complete F01–F10 acceptance still requires the native/WASM differential run, million-row performance corpus, and real Excel producer/reopen verification. Shared-formula templates, persistent aggregate indexes, full number-format/locale behavior, and paged transport for very large spill results remain explicit performance/interoperability work. The current `TEXT` formatter covers common numeric/percentage/quoted-prefix/date masks, not the complete Excel number-format grammar. Structural transforms that would split a reference into unsupported noncontiguous regions reject rather than silently approximate.
