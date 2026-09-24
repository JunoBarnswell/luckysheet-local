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

## Six-round static self-review — replay envelope and paste footprint

1. **Revision precondition:** invalid remote revisions were checked only after mutation handlers had run and local history had been transformed. A malformed revision therefore raised an error after changing the workbook. Validation now runs before preflight or apply; a rejection-path test was added but not executed.
2. **Structural target identity:** row/column history transforms used `MutationInfo.sheetId`, while the registered reducer uses `params.sheetId`; the axis metadata allowed the two to disagree. The registry now rejects this mismatch before replay; a two-sheet rejection case was added but not executed.
3. **Paste range declaration:** `range.paste` used `mode: 'declared'`, which accepted an empty or shortened affected-range list even though the resolver can determine the canonical footprint. It now requires exact ranges.
4. **Paste snapshot boundary:** cell and point metadata from the replay snapshot were not checked against the clipboard destination/source, so a validly shaped snapshot could write outside its declared cell footprint. Replay validation now checks cell, note, hyperlink, comment, and clear-range coordinates against the canonical paste footprint; column-width coordinates are separately bounded by the worksheet extent.
5. **Paste-owned collections:** validation and conditional-format snapshots replace complete sheet collections, while column widths affect complete columns. Target-cell-only ranges understated those writes. Paste affected ranges now cover the whole sheet for complete rule collections, or the complete columns for width changes; move and workbook-theme changes invalidate history without attempting a field-name transform.
6. **Unscoped history collision:** defined-name mutations and other workbook-scoped operations have empty affected ranges, so two conflicting no-range operations previously appeared disjoint and local undo could overwrite the remote value. Two empty scopes now fail closed and invalidate the local history entry. Feature-level cases cover defined-name replay and paste footprint rejection; they were added but not run.

This slice remains intentionally incomplete: recursive payload coordinate inference and a shared structural patch/conflict-key model are still open. In particular, workbook-theme paste has workbook-wide scope that a cell-range list cannot fully express for queued edits on other sheets; the current history invalidation prevents stale local undo, but the canonical collaboration scope contract is still required. No local tests or builds were run, as requested.

## Six-round static self-review — paste snapshot type integrity

1. **Spec-to-snapshot binding:** the replay handler applies every optional snapshot field regardless of the Paste Special flags. A payload could therefore apply validation, width, or theme changes while the selected spec disabled them. Snapshot field presence and clear ranges now have to match the canonical spec.
2. **Worksheet ownership:** rule snapshots could contain ranges and `sheetId` values for another worksheet, then be assigned to the target worksheet's complete rule collection. Rule ownership and every rule range are now checked against the target sheet and worksheet bounds.
3. **Rule structure:** `validations` and `conditionalFormats` were previously checked only as arrays. Each entry now passes the canonical sheet-rule normalizer before it can enter a replay snapshot.
4. **Note values:** a note entry with a key but malformed or missing note fields passed the schema and could be stored in the review index. Note identity, author/text/timestamp types, and visibility are now required.
5. **Hyperlink targets:** arbitrary target discriminators and incomplete URL/email/sheet/name variants previously passed and were stored in the model. The replay schema now validates the discriminated target shape.
6. **Comment and width entries:** comment arrays and column-width arrays were accepted without validating entry fields; invalid thread identity/scope or non-positive/non-finite widths could reach model state. Thread/reply fields and bounded positive width entries are now checked.

Static regression cases were extended for disabled-spec metadata injection, invalid widths, and foreign-sheet rule ownership. They remain unexecuted per the static-only instruction. Workbook-theme cross-sheet conflict scope and full formula-reference owner/rewrite validation remain part of the canonical owner/patch work, not claimed solved here.

## Nine-round static self-review — paste replay and inverse integrity

1. **Local/replay parity:** the command precondition rejected overlapping same-sheet cut/paste, but the replay schema did not. The permission guard already covers protected mutations in remote replay, so replay now adds the missing semantic overlap rejection in the shared payload validator instead of re-running caller-specific permission checks.
2. **Selected metadata only:** `applyPasteSnapshot` cleared notes, hyperlinks, and comments for every metadata clear range, even when Paste Special selected only validation, widths, or conditional formats. Clearing now follows which typed metadata collections are present in the snapshot.
3. **Column-width footprint:** width entries were bounded by the worksheet, not the paste target/source columns or the clipboard's mapped target columns. Since replay writes the declared column directly, a payload could alter a distant column. Width coordinates now must match a target-column mapping from the source metadata or a same-sheet paste/source range; transposed pastes retain the existing width mapping semantics.
4. **Move tombstones:** cut planning intentionally encodes deleted source notes and hyperlinks as entries with `value: undefined`; the stricter snapshot schema rejected those generated entries. The validator now accepts only these typed deletion markers or valid metadata values, preserving move and undo semantics.
5. **Canonical metadata keys:** leading-zero coordinates passed numeric footprint checks but hyperlink replay stores the original key string, creating map entries normal cell-address lookups cannot find. Metadata keys must now equal their canonical `row:column` serialization.
6. **Rule discriminants:** the canonical normalizer validates rule-specific constraints but does not reject every unknown runtime `type`/`operator` value. Paste replay now checks known discriminants and the data-validation list-source/value shapes before normalization.
7. **Sparse-paste inverse:** with `skipBlanks`, the forward snapshot preserves untouched destination cells and has no target clear range. For a previously empty destination cell that the paste writes, the inverse snapshot had no deletion record, so undo could leave the pasted value behind. The inverse now records a scoped empty-cell tombstone for that case without changing the forward snapshot.
8. **Overlapping width ownership:** when a cut stays in the same column, the width plan wrote the pasted source width and then a source-clear entry for that same column, deleting the width. The source clear now excludes columns also owned by the target mapping.
9. **Clipboard source bounds/identity:** range validation allowed coordinates beyond worksheet limits, and copy source ranges were not part of the mutation's affected cell ranges. A malformed or deleted source worksheet could therefore be used to generate impossible formula addresses while the target range remained valid. Local preconditions and replay schema reject empty sheet identities and out-of-bounds source extents; replay also resolves the source worksheet before applying the snapshot.

Regression test source now covers excluded metadata preservation, cut metadata tombstones, skip-blanks undo, transposed and same-column widths, invalid source/rule/width/key snapshots, and overlapping cut replay. No tests or builds were run. Remaining architectural blockers are unchanged: workbook-theme changes still need a workbook-wide conflict scope, recursive coordinate inference remains, and a shared StructuralPatch/conflict-key model is not yet implemented.
