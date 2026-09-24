# Structural Editing & Reference Integrity Runtime

## Scope and invariants

This redesign covers axis and cell insert/delete, move/copy/cut/paste, fill, drag, sort, table resize, worksheet rename/delete, undo/redo, remote replay, server commit, and OOXML save. A structural operation must have one authoritative mapping for cells and every reference-bearing owner. Planning is side-effect free; the complete change is validated before the workbook, calculation state, projection, history, collaboration state, persistence snapshot, or native package is changed. Unsupported owners fail with an observable typed error. Unknown OOXML content remains preserved, but preservation must not imply that a structurally stale part is valid.

## Six static review passes — confirmed findings

1. **Structural owners:** (a) `StructuralTransform.apply` has separate axis, cell-shift, and move implementations, while row permutation and sheet identity changes use other command paths; (b) metadata preflight clones every sheet plus workbook tables/sources and replays mutations; (c) `StructuralTransformResult` reports calculation-oriented ranges/owners, not all committed owner deltas.
2. **Reference owners:** (a) `ReferenceIndex` indexes formula dependency geometry and formula-owner positions, not typed non-cell owners; (b) defined names, rules, table formulas, drawing/chart/pivot bindings, spill/filter/protection metadata, and print/style-template references are handled by independent scans; (c) TypeScript AST transforms and Java token transforms are separate semantics.
3. **Calculation/projection/history:** (a) before this slice, row permutation and sheet identity mutations rebuilt the FormulaEngine; (b) the rebuild reset and re-enumerated formula owners and, when formulas existed, value inputs across the workbook; (c) row permutation uses exact incremental input ranges, and worksheet rename now applies its formula-owner delta incrementally when the new name cannot resolve previously-invalid cell or defined-name formulas; add/remove/reorder still rebuild, and unified history/projection patch ownership remains open.
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

## Incremental worksheet rename calculation path

The rename plan returns the exact formula-cell owners whose cell formula, preserved formula provenance, or barcode formula changed. Runtime updates the FormulaEngine's sheet-name identity table in place, then routes those owners through the same incremental input/index synchronization used by structural transforms. If a cell formula owner or defined-name formula already refers to the proposed new sheet name, rename retains the full context rebuild: that previously unresolved reference can become valid even though its formula text is unchanged. Worksheet add/remove/reorder retain their existing rebuild boundary because their address-space or 3D-reference semantics differ.

Six static self-review passes:

1. **Mutation effect propagation:** local command application and history/remote replay both deliver the rename plan's structural effect to the runtime listener.
2. **Direct formula owners and indexes:** rewritten cell/provenance/barcode formula owners enter incremental input and auxiliary-reference synchronization; FormulaEngine and RangeIndex receive the same renamed sheet identity table.
3. **Dynamic references:** `INDIRECT` is volatile and is included in automatic smart recalculation; manual mode retains its explicit recalculation semantics.
4. **Previously unresolved names:** this pass found the real gap—defined-name formulas referring to the proposed new sheet name were omitted from the rebuild predicate. The predicate now includes those definitions, since unchanged name text otherwise leaves dependencies stale.
5. **Other formula-bearing metadata:** worksheet table lookup inputs contain table identity/range/column metadata rather than formulas; table-sheet, view, and drawing formulas remain separate owners and are rewritten by the identity plan, not indexed as FormulaEngine cell inputs.
6. **Rebuild and scheduling boundaries:** add/remove/reorder continue to rebuild; rename fallback runs before incremental synchronization; worker cancellation/context generation and mutation replay remain aligned. Local tests/builds were not run.

The real issue found in this pass was fixed. Remaining architecture risks include integrating this transitional identity effect into the canonical StructuralPatch and applying the same patch atomically to history, remote replay, and server state.

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

## Eight-round static self-review — row-permutation owner integrity

1. **Scope parity:** the server transforms row-addressed rule and protection ranges through `affectedColumnEnd`, while the client used only the selected sort rectangle. The canonical client plan now carries the same full-row metadata scope.
2. **Rule formula semantics:** conditional-format and validation formulas are evaluated relative to `formulaAnchor`; moving an anchor without shifting relative A1 references changed rule meaning. Both runtimes now offset the formula fields when that owner moves.
3. **Defined-name ownership:** anchored workbook- and sheet-scoped names were omitted from row permutation, leaving the formula engine's anchor stale. Anchors and formulas now move together, and the server updates its workbook-name projection.
4. **Reusable validation templates:** formula anchors and relative formulas in workbook cell-style templates were omitted. Both runtimes now include these owners in the scope and transform.
5. **Declared extent and bounds:** the extent omitted formula-owner anchors beyond the materialized grid, while client schema accepted columns beyond Excel's maximum. Client/server extent calculation now includes in-scope anchors and the client rejects out-of-bounds declarations.
6. **Failure atomicity:** newly transformed formulas could fail parsing or row-bound checks after cell writes. Rule/name/template changes are staged and checked before client mutation; server validation evaluates copies before reduction.
7. **Exact outline mapping:** a row outline group could become discontiguous; the client discovered this only during application and the server widened it by mapping endpoints. Both now reject fragmentation during preflight. A suspected area-bound discrepancy was checked against both implementations and was not a defect, so the existing matching cap remains unchanged.
8. **Anchor identity and implicit origins:** formula owners on another worksheet must not be shifted by matching row/column numbers, and a rule without an explicit anchor inherits its first range's origin. The client now scopes workbook owners by `sheetId`; when a non-identity permutation can reorder exact range fragments, both runtimes materialize the mapped implicit rule anchor even if that row itself is fixed.

Static regression source now covers anchor/formula rebasing, cross-sheet owner isolation, implicit-anchor range fragmentation, workbook-owner extent, invalid formula rejection before writes, maximum column rejection, and fragmented outline rejection. No tests or builds were run; the current draft PR's CI remains the external verification boundary.

## Six-round static self-review — collaboration structural fail-close

1. **Capability inventory:** the generated mutation contract identifies sheet add/remove/rename/duplicate/restore/reorder as workbook-scope structural operations with no cell `affectedRanges`.
2. **Classification:** those mutation IDs had no collaboration kind, so `classifyMutation` reduced them to `unknown` despite their explicit structural capability.
3. **Committed-operation branch:** `rebaseMutation` treated an unrecognized committed mutation as unchanged unless it matched the short unsupported list; a sheet rename could therefore pass through without touching pending sheet-qualified formulas.
4. **Preflight interaction:** `CollaborationSession.assertPendingCanRebase` detects overlap only through ranges. Workbook-scope identity operations carry none, so that check could not compensate for the missing classification.
5. **Pending operation branch:** move-range/sort/table-resize were rejected when committed but still passed through recursive coordinate inference when pending across row/column changes, despite lacking canonical patch semantics.
6. **Atomic replay boundary:** `applyRemote` runs the preflight before runtime mutation, so explicit classification plus rejection preserves the queued draft and prevents the remote identity/unsupported structural change from being locally applied as if the queue were safe.

Confirmed repairs: worksheet identity mutations now have a dedicated collaboration kind; unknown committed mutations fail closed; unsupported structural kinds are rejected on both sides of an axis rebase instead of guessing coordinates from field names. Regression sources cover unknown commits, worksheet identity IDs, unsupported pending range moves, and the existing supported axis-rebase path. No tests/builds were run. Sheet identity/move/sort/table-resize OT remains intentionally blocked until the shared StructuralPatch provides a canonical transform.

## Six-round static self-review — protection bounds for sparse structural scopes

1. **CI failure reproduction path:** row-permutation metadata legitimately extends through column 8 while the canonical sheet currently materializes only two columns. `affectedRanges` includes the complete row-remap/conflict extent, not only the sort selection.
2. **Server protection validation:** `ProtectionResolver` checked the full affected range against `rowCount`/`columnCount` before examining protection rules. This rejected a valid sparse metadata scope as `VALIDATION_ERROR`, even when no protection rule intersected it.
3. **Affected-range contract:** the broad range is required by conflict detection, history, invalidation, and metadata/protection-rule transformation. Narrowing it to the selected sort cells would silently weaken an existing contract.
4. **Protection behavior cross-check:** the existing protected-metadata regression explicitly requires a locked range anywhere in the remapped extent to block sorting. Its fixture now places that lock beyond the materialized column count, preserving the intended rejection path while verifying sparse bounds.
5. **Client/server parity:** the frontend resolver validates permission ranges against Excel's maximum row/column counts, while the backend used current materialized dimensions. The backend now retains canonical-dimension presence checks but validates coordinates against the same Excel address-space limits.
6. **Failure ordering and fail-close:** the second CI failure was the downstream symptom: the bounds error masked the row-permutation formula-owner preflight error (`SERVICE_UNAVAILABLE`). Existing source coverage asserts the latter, and the owner-transform fixture asserts that the prepared affected range remains wide. Regression sources were updated; no tests or builds were run.

The initial idea to separate permission scope from affected scope was rejected after the protected-metadata test established that wide protection checking is intentional. The confirmed defect was specifically the backend's use of materialized dimensions as address-space maxima; the correction preserves both broad protection coverage and sparse metadata support.
