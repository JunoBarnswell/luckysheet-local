# Structural Editing & Reference Integrity Runtime

## Scope and invariants

This redesign covers axis and cell insert/delete, move/copy/cut/paste, fill, drag, sort, table resize, worksheet rename/delete, undo/redo, remote replay, server commit, and OOXML save. A structural operation must have one authoritative mapping for cells and every reference-bearing owner. Planning is side-effect free; the complete change is validated before the workbook, calculation state, projection, history, collaboration state, persistence snapshot, or native package is changed. Unsupported owners fail with an observable typed error. Unknown OOXML content remains preserved, but preservation must not imply that a structurally stale part is valid.

## Six static review passes — confirmed findings

1. **Structural owners:** (a) `StructuralTransform.apply` has separate axis, cell-shift, and move implementations, while row permutation and sheet identity changes use other command paths; (b) metadata preflight clones every sheet plus workbook tables/sources and replays mutations; (c) `StructuralTransformResult` reports calculation-oriented ranges/owners, not all committed owner deltas.
2. **Reference owners:** (a) `ReferenceIndex` indexes formula dependency geometry and formula-owner positions, not typed non-cell owners; (b) defined names, rules, table formulas, drawing/chart/pivot bindings, spill/filter/protection metadata, and print/style-template references are handled by independent scans; (c) TypeScript AST transforms and Java token transforms are separate semantics.
3. **Calculation/projection/history:** (a) before this slice, row permutation and sheet identity mutations rebuilt the FormulaEngine; (b) the rebuild reset and re-enumerated formula owners and, when formulas existed, value inputs across the workbook; (c) row permutation now returns its exact calculation input range and uses the incremental synchronization path, while sheet identity changes and unified history/projection patch ownership remain open.
4. **Server structural owners:** (a) `applyPublicMutations` used to copy each snapshot before dispatch while every executable reducer independently copied before mutation, adding one redundant full-workbook copy per non-empty replay batch; the wrapper copy is now retained only for empty batches; (b) structural cache invalidation walks all sheets and formula cells; (c) the Java reducer mutates cells and each owner family through operation-specific routines rather than consuming a plan shared with the client.
5. **Collaboration coordinates:** (a) OT maintains duplicate row/column bounds and transform algebra; (b) `transformParams` infers coordinate meaning recursively from field names; (c) rebasing queued drafts across move, sort, and table-resize operations is rejected because no canonical patch exists, while direct committed replay could still apply those operations and leave local undo history untransformed. The replay history gap now fails closed, but full OT support remains open.
6. **OOXML structural state:** (a) source-byte passthrough is guarded by exact snapshot hash, but changed-snapshot export still starts from all imported package parts; (b) registered workbook/worksheet/native object paths are rewritten while opaque parts have no structural transform; (c) an opaque part with an affected reference can therefore be preserved byte-for-byte with stale coordinates instead of causing a typed rejection.

## Target architecture

- `CanonicalStructuralPlanner` resolves the operation and current owner index into an immutable `StructuralPatch`; it does not mutate a model or clone/replay the workbook to discover whether a mutation is valid.
- `ReferenceTransformDomain` owns typed formula, range, point, anchor, and identity transforms. Every supported owner kind registers its read/write contract and index updates. Unknown or untransformable owners reject the plan.
- `StructuralPatch` records cell and metadata deltas, transformed reference owners, formula input deltas, projection invalidations, and reversible history data. Apply is atomic and is the only mutation entrypoint for supported structural operations.
- Calculation, projection, history, collaboration, persistence, server commit, and OOXML consume the same patch. TypeScript and Java share the operation/patch contract and structural mapping vectors; neither layer independently guesses transformations.
- Native-package export may retain opaque parts only when their structural references are known unchanged. If a structural patch could affect an opaque, unmodeled reference owner, the operation/export must fail with an explicit unsupported-feature error while preserving the source artifact.

## Implementation sequence and current slice

1. Stabilize fail-close semantics and establish typed owner/patch contracts.
2. Move axis, cell-shift, move, sort, and table-resize transformations behind the planner and reference domain.
3. Drive calculation, projections, history, OT, and server reduction from patch deltas; remove full-workbook rebuild/copy paths where the patch provides exact changes.
4. Make OOXML package participation explicit and reject structurally stale opaque references.
5. Review every operation in the scope list against success, rejection, inverse, remote replay, and native-save contracts.

This PR slice fixes statically confirmed server regressions and connects row permutation to exact incremental formula-input synchronization. No local build or test was run in this slice; repository CI remains the external verification source.

## Six-round static self-review — reducer snapshot isolation

1. **Production callers:** historical reconstruction invokes `applyPublicMutations` once per committed operation, so the extra root copy scales with replay length and workbook snapshot size.
2. **Descriptor coverage:** all registered executable descriptor implementations were traced; each copies its input before mutation, including pivot reducers through `canonicalSnapshot`. Restore and unavailable descriptors reject without mutation.
3. **Failure isolation:** the first reducer returns a detached candidate, so later reducers receive that candidate; removing the wrapper copy does not expose the caller's input to partial writes.
4. **Empty-batch behavior:** the prior method returned a detached copy even when no mutation was supplied. The fast path explicitly preserves that observable isolation contract.
5. **Extension boundary:** `MutationDescriptor.apply` now documents the no-input-mutation/independent-result contract. No production call site dynamically registers external descriptors; built-in registrations were enumerated and reviewed.
6. **Regression surface:** source-level coverage asserts a successful reduction leaves the input unchanged and an empty reduction remains detached. These tests were added but not executed, as requested; the patch receives static review only.

## Six-round static self-review — remote history integrity

1. **Replay order:** `CommandRuntime.applyRemoteMutations` preflights and applies committed mutations, then transforms local undo/redo entries against each applied mutation.
2. **Queue boundary:** `CollaborationSession.assertPendingCanRebase` checks queued drafts only; it does not protect already-recorded local history when the queue is empty.
3. **Unknown structural effect:** the old `structuralDelta` inferred row/column semantics from mutation ID substrings. For `range.move`, `rows.permuted`, and cell shifts it returned no delta, and `transformHistoryEntry` treated that as an identity transform.
4. **Reachability:** same-sheet `range.move` is registered with a real `StructuralTransform` handler and is accepted by the committed-operation path when no queued draft conflicts.
5. **User-visible consequence:** local undo replays the stored inverse against the current workbook without a matching remote-history revision guard. If remote move overwrites the local edit's destination, stale undo can overwrite the moved value.
6. **Repair and scope:** mutation registrations now declare an explicit typed axis transform or an invalidation reason for worksheet identity, cell shifts, range moves, cut/paste moves, row permutations, and sheet-table changes; overlapping non-structural remote writes invalidate conflicting entries. Regression cases were added to the feature-level collaboration suite but were not run. This is a temporary safety boundary, not full rebase support: structural payload coordinates are still recursively inferred by field name, and canonical StructuralPatch history rebasing remains required.
