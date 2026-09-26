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

## 2026-09-26 continuation — local-only structural gate

### Confirmed issue and scoped repair

`WorkbookSession` previously rejected mutations only when `localOnly` was false and `remoteConnected` was false. That condition therefore allowed local-only workbooks to execute address-changing structural mutations through the TypeScript reducers, despite the user's decision that Java owns structural planning and may be required online. The continuation adds generated contract sets for 27 planner-sensitive mutation IDs and 33 corresponding command IDs. The command gate rejects before context resolution/reducer execution, `canExecute` reports those commands unavailable offline, and the central mutation guard provides a second check for compound/history paths. Pivot “Show Details” now checks service availability before preparing or staging detail blocks, and its menu entry is disabled offline. The changes are committed to PR #345 as `21e43340`, `72d516e3`, and the current continuation; test-only session fixtures explicitly model transport-ready state and do not claim to exercise Java planning.

### End-to-end authority gap remains open

The source chain is still: `WorkbookSession.runCommand` → `CommandRuntime.execute` applies the TypeScript mutation → runtime mutation/command listeners capture its `MutationInfo` → `CollaborationSession.enqueueLocalMutations` creates an envelope containing mutation IDs and params → `WorkbookApiClient.commitOperation` sends it → `WorkbookOperationService.commitInternal` runs Java reducers and stores the committed operation. The committed structural patch currently carries formula/defined-name owner deltas, not complete cell and metadata after-facts. Own-client acknowledgement therefore reconciles those deltas against the already-mutated TS model; another client replays the committed intent through its TS reducer. Server persistence is authoritative, but structural planning and replay are not Java-only.

The required next slice is not another availability check: it must establish an intent-first transaction boundary and complete, reversible cell/metadata/reference facts produced by Java, then make local apply, undo/redo, remote replay, persistence, and OOXML consume those facts without independently rerunning structural transforms. Until that contract is implemented, the user-approved single-Java-planner objective remains incomplete.

### Intent-first transaction and owner-complete patch — implementation design

**Confirmed source boundary.** The web client currently calls `WorkbookSession.runCommand → CommandRuntime.execute`; inside the synchronous command transaction each registered `mutation.apply` runs before listeners capture its ID/params. `runtime.ts` later creates/enqueues the operation and calls `WorkbookApiClient.commitOperation`. The sole workbook operation POST then runs `WorkbookOperationService.commitInternal` and Java reducers. Its v3 patch has formula/defined-name owner deltas only. Own-client acknowledgement reconciles those deltas after local mutation; remote replay calls the TS mutation path. This confirms that “online” is only an availability gate, not Java-first planning.

**Chosen transaction boundary.** Keep one idempotent operation-commit request rather than a plan endpoint followed by a second commit: validate operation identity, ACL/protection, base revision and conflicts; derive the plan; apply the patch to the transaction-owned Java snapshot; persist the snapshot/history/outbox atomically; and return the exact committed patch. A command must produce typed, serializable intent without invoking a workbook reducer. The UI model stays unchanged until it has the committed patch; transport failure or rejected intent therefore cannot leave an optimistic structural mutation to reconcile. The command/intent API must preserve multi-mutation envelope atomicity and operation idempotency.

**Canonical patch contract.** Replace the partial v3 runtime contract with one explicitly versioned, immutable, reversible patch containing: (1) typed cell owner before/after address and value/presence facts; (2) typed metadata/reference owner identity plus before/after presence and values, including workbook/sheet identity and order; (3) calculation input additions/removals and exact projection invalidations; and (4) base/result revision plus conflict/authorization scope. Formula and defined-name facts become owner families in this same patch, not parallel delta channels. Owner keys are schema-defined by kind, never inferred recursively from payload field names. Patch application validates every precondition before its first write, preserves absent-vs-null and stable ordering, and rejects duplicate/conflicting owner facts. A structural mutation with an unindexed or unsupported owner fails with a typed error before persistence.

**Indexes and complexity.** Extend the existing TS `ReferenceIndex`/`RangeIndex` rather than create a sibling formula index. It owns formula/range dependencies and non-cell owners (CF/DV, names, tables/structured references, filter, Pivot, Sparkline, chart/drawing anchors, spill, protection, print/layout, and data-source ranges). The Java index is a bounded, revision/checksum-keyed derived cache built from the canonical workbook on a verified cache miss and advanced by the same committed patch. It stores only typed owner identities/geometries, not copied cell or metadata payloads; stale/mismatched entries are never used to authorize a write and may be discarded/rebuilt. This does not introduce an independent H2 read model or second source of truth. The planner queries only owners intersecting the affected geometry and emits facts proportional to affected cells/owners; it must not deep-copy the workbook or compute generic before/after JSON diffs. `ReferenceTransformDomain` supplies the shared row/column, cell-shift, move, permutation, anchor, and identity semantics; TS and Java consume the same wire cases and vectors.

**Consumers and failure recovery.** The post-commit TS applier is a generic patch consumer, not a mutation-ID switch or a second structural planner. It updates WorkbookModel/CellMatrix, FormulaEngine, projection caches and selection as one validated transaction. History stores/inverts the same facts; remote replay and own ACK apply that same patch once; persistence and the coordination outbox store/broadcast the exact committed envelope. OOXML maps known package owners through the same patch and preserves unknown bytes only when evidence proves their references cannot be affected; otherwise planning/export fails as unsupported. If applying a valid committed patch locally fails, reload exactly its committed revision; if that recovery cannot be verified, stop editing and expose the failure rather than rerunning a TS reducer or silently accepting divergent state.

**Clean break and migration.** Bump the patch/snapshot history protocol at an explicit Flyway migration boundary; runtime readers accept only the new canonical version. The migration must verify checkpoint checksums and contiguous history, replay v3 intents with the Java planner, verify their old formula/name deltas against derived facts, rewrite operation-log and pending-outbox envelopes, and update checksums atomically. If history/checkpoints cannot reproduce a fact or a legacy owner is unsupported, migration fails with the workbook/revision and recovery source; it must not retain a v3 runtime fallback or invent missing before-values. No dual-write/shim phase.

Six design self-review checks: (1) operation identity/base-revision validation and patch persistence stay in one idempotent transaction, avoiding a plan/commit TOCTOU gap; (2) no workbook command or history entry becomes visible before server commit; (3) every patch fact has an explicit typed owner, presence, and before/after value, so inverse does not rerun coordinate logic; (4) the client validates the full patch against its base before the first write and has an exact-revision recovery route; (5) the ReferenceIndex cache is reconstructible and revision/checksum gated, never a second authority; (6) old operation/outbox history and opaque OOXML owners either migrate with proof or reject at the owning boundary. This section is a design change only: no code, tests, builds, browser, Excel, or performance measurements were run; final gates are listed below.

**Acceptance boundary.** First implement the transaction/patch contract and axis insert/delete + row permutation end-to-end, while keeping the feature disabled unless the server and exact patch consumer are available. Then migrate cell shifts/move/cut/paste/fill, sheet identity, tables and remaining owners by deleting their old reducer paths in the same delivery. For each family require success, rejection-before-write, inverse, remote replay, projection/calculation and package-owner coverage. Final acceptance still includes shared TS/Java vectors, backend/frontend gates, real in-app interaction, native Excel round-trip corpus, and large sparse/dense CPU/heap measurements; none is inferred from static review or CI alone.

### Six static review lenses and evidence

1. **Contract coverage**: generator checks uniqueness and canonical permission ownership; a Java regression requires every classified mutation to resolve to a registered reducer. Static source comparison found all 27 IDs in the Java registry ID set.
2. **Command ordering**: planner commands are rejected before parameter-context resolution; Pivot drill-down is rejected before reading/preparing/uploading details. The mutation guard remains before each mutation's `apply` for callers or compound commands outside the explicit command list.
3. **History**: undo/redo invoke mutation-guard preflight; offline structural history therefore rejects before replay. Successful remote committed replay is intentionally exempt from the local service-availability gate.
4. **Atomic rejection**: command runtime rolls back already-applied mutations if a later step fails; the new simple rejection regression checks unchanged model snapshot and history depth. Those new regressions were not run locally.
5. **UI state and test scope**: all 33 classified command IDs return unavailable from `canExecute` while offline; Pivot’s Show Details menu uses that result. The test fixture only toggles runtime transport flags and is explicitly not a live service test. A direct `runtime.commands.execute('sheet.add')` test call was included in the fixture audit.
6. **Architecture and acceptance**: the online path still executes a TS reducer first, as the end-to-end trace above demonstrates. Contract generation and `git diff --check` passed; both PR `canonical-build` checks passed on head `72d516e3`. The app/browser remains gated by the previously observed `/api/auth/config` 500, and native Excel round-trip/performance acceptance remain outstanding.

This pass confirms one local-only structural-editing root cause; the 27/33 contract coverage counts are not defect counts and do not satisfy the requested 30-distinct-issue batch. No same-root symptoms were double-counted.

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

## Six-round static self-review — shared axis point mapping

1. **Core structural application:** `structural-transform.ts` owned an insertion/deletion point mapper used by cells, anchors, dimensions, and metadata. Its interval boundary is half-open for deletion; the new core-model function preserves that exact mapping.
2. **Undo/redo history:** `CommandRuntime.transformIndex` independently implemented the same insertion and deletion mapping. The caller still validates the coordinate and converts a deleted/overflowed result to its existing `undefined` invalidation result.
3. **Collaboration rebase:** `ot-rebase.ts` had a third copy. Its deleted-coordinate conflict and worksheet-bound errors remain at the collaboration boundary; only the arithmetic moved to core-model.
4. **Boundary comparison:** insertion before/at/after and deletion before/inside/after were checked against all three previous implementations. No change was made to each caller's `at`/`count` validation or address-space maximum.
5. **Dependency direction:** both consumers already depend on `core-model`; the extracted module has no dependencies, so this removes duplicate ownership without introducing a cycle or a second state path.
6. **Regression surface:** source-level cases cover insertion's unchanged and shifted sides and deletion's unchanged, removed, and gap-closing coordinates. The tests were added but not executed, per the static-only instruction.

Confirmed repair: TypeScript structural application, local history rebase, and collaboration rebase now call one scalar axis-mapping function. This does not unify range transforms, recursive payload rewrites, or Java/server structural semantics; those remain open parts of the Canonical StructuralPatch work.

## Six-round static self-review — OOXML untouched-export contract

1. **Fast-path reachability:** an imported OPC artifact with the same filename and snapshot hash returned source bytes before writer options or output limits were evaluated.
2. **Macro policy:** `preserveMacros: false` is consumed by the OPC writer to remove VBA parts and relationships; the byte shortcut returned the original macro project instead. Reuse is now disabled when the source contains VBA and the caller opts out.
3. **Cell-cache and date serialization:** `includeCachedValues: false` removes cached formula values, and `dateSystem` controls workbook XML and date serialization. Both were bypassed; overrides now force normal export.
4. **Asset authority:** `assetBytes` is the writer's authoritative AssetStore input for image/embedded-object parts. The shortcut ignored supplied bytes; any supplied asset map now routes through the writer.
5. **Compatibility report:** the shortcut returned the import-time report, so an export-level or date-system override could yield a stale report despite the requested output contract. Source reuse now requires matching report level and date system.
6. **Resource boundary:** export options carry finite resource limits, but the emitted package was reloaded with `{}`. The output validation now receives the caller's limits; custom limits also disable the raw-byte shortcut.

Confirmed repair: source bytes are reused only when the snapshot, filename, compatibility/date report, and output-affecting options all match the source; regenerated packages are validated with the requested limits. Regression source covers the combined macro/cache/date/asset/report path and a low-limit rejection. No tests or builds were run, per the static-only instruction.

## Six-round static self-review — Java axis point mapping

1. **Formula references:** `FormulaReferenceTransformer.mapAxisPoint` and the reducer's `shiftIndex` had independent insertion/deletion arithmetic.
2. **Insertion boundary:** both left coordinates before `at` unchanged and map `at` itself to `at + count`; this remains the shared point contract.
3. **Deletion boundary:** both preserve points before `at`, return the existing `-1` removed-point sentinel inside `[at, at + count)`, and subtract `count` after the interval.
4. **Overflow responsibility:** formula mapping retains its Excel row/column maximum check after the common long-valued map; reducer mapping retains its existing sentinel API, and axis bounds are validated before reducer application.
5. **Call-path validation:** `remapAxis` validates `at/count` before formula rewrite, and `StructuralSnapshotReducer` validates worksheet bounds before owner/cell remapping. The helper therefore remains a pure mapper rather than a competing validator.
6. **Regression surface:** Java source cases cover before/at/after insertion and before/inside/after deletion, matching the TypeScript core-model cases. They were added but not executed, per the static-only instruction.

Confirmed repair: Java cell/metadata reduction and formula-reference rewriting now use one point-mapping owner, with language-specific validation and error boundaries retained. This aligns the scalar rule with TypeScript but does not yet establish shared cross-language vectors or unify interval/StructuralPatch semantics.

## Six-round static self-review — data-only formula owner patches

1. **Plan representation:** `StagedStructuralFormulaChange` stored executable `validate`/`apply` closures, so the plan had no inspectable owner identity or serializable before/after values. Replaced it with a discriminated data record.
2. **Owner lifetime:** closures retained the exact nested object seen during preflight and could not resolve the canonical owner again when applying. Replaced captures with workbook-local locators and current-state resolution.
3. **Array identity:** a bare table-column or data-view field index could address a different owner after collection changes. Locators now pair the index with the stable field ID; static call-path inspection confirmed the structural helpers in this operation do not reorder either collection.
4. **Variant/path integrity:** drawing payload maps are heterogeneous and validation list sources are a discriminated union. Reads and writes now reject a non-shape payload or a non-formula list source instead of targeting an invalid variant.
5. **Apply ordering:** validating and writing each staged item in one pass could leave earlier formula owners changed if a later owner had drifted. Application retains a full precondition pass before any write.
6. **Address snapshot:** the first data-record version still held the live mutable `formulaAnchor` object as `before`, allowing an in-place coordinate mutation to alter the expected value itself. Anchor before/after coordinates are now copied into `Readonly<CellAddress>` values.

Confirmed repair: workbook-level formula owner rewrites are represented as locator-based data patches with value snapshots; all locators are validated before any patch is applied. Existing structural source coverage exercises successful template-anchor/formula rewrites and rejection of invalid/deleted anchors. No local tests or builds were run, per the static-only instruction. This is a bounded repair within the planner path, not completion of the cross-layer `StructuralPatch` redesign.

## Six-round static self-review — ReportSheet structural coordinate ownership

1. **Consumer evidence:** `buildReportProjection` reads each persisted `binding.cell` as the destination coordinate on the report worksheet. Structural edits previously changed worksheet cells without changing these persisted coordinates, so projection could target the wrong template cells.
2. **Axis operations:** row/column insert and delete are the canonical point-mapping path. Report bindings now use the same insertion/deletion point rule; repeated-header rows use it for row operations. Deleting a binding anchor fails before cell mutation; deleting a repeated-header row removes that row reference.
3. **Cell-shift ordering:** the original metadata routine ran after cell relocation. Report-owner mapping is now checked on staged metadata and materialized before changing cells; removal or out-of-bounds results reject rather than leave a partially shifted report definition.
4. **Range-move ownership:** a binding inside the source follows the moved block. A binding already at the destination would be overwritten, so the operation rejects before source/destination cells change. Non-overlapping source and destination bounds are already enforced by the move planner.
5. **Row-permutation scope:** bindings are remapped only inside the operation's declared metadata scope; unrelated report bindings remain fixed. Repeated-header rows inside the permuted row interval follow the same source-to-destination row map.
6. **Cross-runtime and failure boundary:** the Java snapshot reducer now plans the same four transformations before cell mutation, validates report coordinates against worksheet limits, and returns typed validation/unsupported-reference failures. Source regressions cover successful axis/cell-shift/move/permutation mapping and deletion/overwrite rejection; tests and builds were not run.

Confirmed issue: one omitted coordinate owner manifested across axis shifts, cell shifts, range moves, and row permutations in both runtimes. The six passes validate that shared root cause and its operation-specific failure modes; they are not presented as six unrelated defects. Static review only; no local test/build execution.

## Remote CI correction — ReportSheet cell-shift type boundary

1. **Authoritative failure:** CI run `36076281198` passed the frontend build and failed Java compilation at `StructuralSnapshotReducer.java:101` with `RangeRef cannot be converted to FormulaReferenceTransformer.Range`.
2. **Signature tracing:** `FormulaReferenceTransformer.remapCellShiftCoordinate` explicitly accepts its own range record, not the API `RangeRef` contract.
3. **Existing canonical conversion:** the reducer already owns `formulaRange(RangeRef)` and uses it at other formula-transform boundaries; no new adapter or duplicate conversion was introduced.
4. **Scope check:** only the new ReportSheet cell-shift mapper passed the wrong static type; axis, move, and permutation paths use their native coordinate records and were not changed.
5. **Failure ordering:** conversion is performed while building the ReportSheet plan, before cell extraction/clear, so the correction preserves the pre-mutation rejection boundary.
6. **Verification discipline:** the remote compile evidence was used as a static diagnostic; no local test/build was run. The corrected commit's remote CI must be checked before treating the fix as verified.

## Static review update — Sheet Table width integrity

Six independent review passes confirmed two real defects in one integrity chain:

1. **Model invariant:** `validateSheetTableModel` requires the table column count to equal the range width.
2. **TypeScript axis path:** insertion preflight returned immediately for every insert; the later table reducer shifted only `range` and AutoFilter.
3. **Java axis path:** insert dispatch skipped delete-preservation checks; its table reducer likewise shifted range/filter only, before any schema update.
4. **Snapshot ingress:** both canonical snapshot validators accepted a Sheet Table with a range/column-count mismatch; the frontend hydration path did not invoke the feature-layer table validator.
5. **Persistence consequence:** OOXML writes `ref` from the table range and `tableColumns count` from the column list, so the invalid model can produce contradictory package metadata.
6. **Failure boundary:** both axis reducers can reject the unsupported interior-column case before touching cells or metadata; insertion before a table remains a supported range translation with stable width.

The fixes enforce the width invariant at TS and Java snapshot boundaries and reject interior worksheet-column insertion in both reducers until a canonical table-column patch supplies new stable column identities and names. Regression sources cover valid snapshot/table translation and width mismatch/interior-insert rejection with unchanged state. No local tests or builds were run; remote CI remains the verification gate.

Confirmed defect: one mismatched range type at the new Java cell-shift call site. Fix: pass `formulaRange(selection)` to the existing transformer contract. The correction is committed separately so the failed CI head remains auditable.

## Six-round static self-review — OOXML unknown worksheet-node fail-close

1. **Import detection:** `detectWorksheetCapabilities` classifies unmodeled worksheet-root children as `unknown-worksheet-node`; balanced import retains the source artifact while strict import reports/rejects unsupported capabilities.
2. **Raw-byte path:** exact-source reuse is allowed only before regeneration and returns the original package bytes, so unsupported nodes remain intact when no output-affecting change is requested.
3. **Writer coverage:** `buildWorksheetXml` rebuilds each worksheet and copies only an explicit list of modeled opaque children; arbitrary unknown root nodes are not serialized. This is the concrete data-loss path.
4. **Failure semantics:** the export report was computed after package generation and did not prevent omission. The export boundary now checks source detections before invoking the writer and raises `NATIVE_DOCUMENT_UNCHANGED_SAVE_REQUIRED` with the part location and recovery guidance.
5. **Boundary review:** package-level opaque parts and worksheet `extLst` use distinct preservation paths; the new gate is limited to unknown worksheet-root nodes and does not block those existing paths.
6. **Regression reachability:** the previous “editable export” opaque-part case kept the snapshot unchanged and therefore exercised raw-byte reuse, not the writer. Its source now changes a cell; separate regressions check unchanged-source preservation and edited-export rejection for unsupported worksheet XML. The Java CI fixture supplies the required `pane` and four `review` index objects used by the reducer.

Confirmed defects: unknown worksheet-root nodes could be silently omitted by regenerated exports, and the ReportSheet reducer regression failed before reaching its assertions because its fixture omitted `pane`. The fixes are source-reviewed only; no local tests or builds were run. Remote CI after push is required to validate the new test sources.

### Six-round self-review correction — worksheet extension reference boundary

1. **Feature detection:** `extLst` is a known worksheet container, so it is not classified as an unknown root node; unmodeled child `<ext>` elements instead produce `unknown-extension`, while extension-hosted CF/DV receive dedicated feature IDs.
2. **Writer behavior:** `serializeWorksheetControlExtensions` clones the original `extLst` node and reserializes unrecognized extension children without updating their formulas/ranges; it is not the source-byte fast path.
3. **Structural-owner path:** TypeScript/Java structural reducers only traverse modeled snapshot owners; the extension XML remains outside `ReferenceIndex` and no reference transform is applied to it.
4. **Export information boundary:** `NativeDocumentExportRequest` supplies a snapshot and source artifact, not the executed mutation or a structural patch. Export therefore cannot prove a regenerated snapshot changed only ordinary values.
5. **Feature-scope cross-check:** `unknown-extension`, `extended-validation`, and `extended-conditional-format` all lack a complete canonical writer/structural owner; known controls and sparklines have their own transforms and remain outside this rejection set.
6. **Recovery and regression:** unchanged-source reuse retains the extension exactly; regenerated output now fails with `NATIVE_DOCUMENT_UNCHANGED_SAVE_REQUIRED` and a feature location. Source coverage checks both unchanged-byte preservation and rejection after snapshot change.

Confirmed issue: unchanged opaque worksheet extensions could retain stale address references after structural edits because export serialized raw extension XML with no owner transform. The boundary is now fail-closed for any non-byte-reuse export. No local tests/builds were run; the preserved-only chart slice is reviewed below, and other package-level owner types remain an explicit audit item.

### CI fixture follow-up — ReportSheet reducer canonical inputs

CI runs `36076501168`, `36077170083`, and `36077447582` successively advanced the ReportSheet Java source regression through missing fixture requirements: `pane`, then `review` indexes, then the worksheet name consumed by the formula identity table. The fixture now provides those canonical fields, including `kind: report-sheet`; this is test-data correction only, not a production runtime behavior change. No local tests/builds were run.

### Six-round static self-review — preserved-only chart reference boundary

1. **Package owner discovery:** `readNativeChartGraph` finds chart parts through worksheet/drawing relationships and records the source chart identity/part; chart identity can explicitly be `editable: false`.
2. **Projection boundary:** `projectNativeCharts` skips non-editable definitions, so their drawing anchors and chart series formulas never enter the canonical worksheet snapshot.
3. **Structural reachability:** core structural transforms enumerate snapshot-owned drawings and reference owners only; the skipped chart has no typed anchor/range owner to transform.
4. **Writer behavior:** `synchronizeNativeCharts` keeps preserved-only anchors/relationships unless a matching active canonical payload owns them, leaving the original drawing XML unchanged across regenerated exports.
5. **Export boundary:** the request contains a snapshot/artifact but no operation patch; it cannot distinguish a structural edit from an ordinary value edit. Rewriting a package with a preserved-only or unindexed chart could therefore silently retain stale ranges/anchors.
6. **Scope and regressions:** unchanged-source export remains byte-preserving; regeneration now fails closed for preserved-only/unindexed chart parts. Opaque binary parts without a reference owner still exercise the regenerated writer in a separate source regression.

Confirmed issue: preserved-only charts were outside the structural owner graph while regenerated packages retained their old relationship-backed anchors/references. The export guard now blocks regeneration until the chart has a canonical structural owner. Static-only source review; local tests/builds were not run.

### Six-pass static self-review — owned worksheet-extension classification and deletion

1. **Detector-to-manifest comparison:** `detectWorksheetCapabilities` initially marked every worksheet `<ext>` as unknown. The writer emits canonical x14 `sparklineGroups`, slicer lists, and timeline references, so this incorrectly rejected edited workbooks using those supported owners.
2. **Import-owner trace:** sparkline formulas are imported into typed `sourceRange` and cell `anchor` fields; native slicer/timeline relationships are resolved into validated `NativePivotControlDefinition` entries. Their structural behavior is therefore represented in the current snapshot/graph when those parsers accept the source.
3. **Structural-transform trace:** `structural-transform.ts` maps sparkline source ranges and anchors; drawing participants map the control anchors. Unsupported control identities remain invalid in the native graph and are not classified as canonical ownership.
4. **Writer replacement trace:** the worksheet writer regenerates sparkline groups from canonical models and slicer/timeline refs from native control identities. It previously retained the old sparkline extension when the last canonical sparkline was removed, causing a deleted object/reference to reappear in output.
5. **Fail-close countercheck:** only the exact owned extension URI, expected element tree, recognized attributes/values, resolvable sparkline sheet references, and graph-backed control relationship IDs are accepted. Unknown URIs, extra nodes/attributes, unresolved formulas, and invalid controls remain `unknown-extension` and block regeneration.
6. **Reachability and regression-source review:** changed-snapshot export now covers owned sparkline/control extensions; source cases cover removing the final sparkline and rejecting an unowned child. Existing unknown-worksheet-extension rejection remains. `git diff --check` passed; local tests/builds were not run.

Confirmed defects: the detector blocked valid canonically-owned sparkline/slicer/timeline extensions, and the writer could retain the source sparkline group after canonical deletion. The capability boundary now validates ownership instead of treating all extensions equally; the writer removes only source extensions proven to have canonical owners before rebuilding them from current state. Rollback: revert the single follow-up commit in PR #345; no persisted schema migration is introduced.

### Six-pass static review — workbook-level extension gap

1. **Serializer path:** `buildWorkbookXml` copies workbook `<extLst>` through `serializeWorkbookControlExtensions`; unrecognized extension children survive regeneration without reference rewriting.
2. **Gate path:** the export guard previously called only `detectWorksheetCapabilities`, which walks `sheetPartById`; `workbook.xml` extensions were therefore absent from the gate.
3. **Mutation ambiguity:** export receives the current snapshot and source artifact, not an operation patch, so it cannot prove an arbitrary workbook extension is unaffected by a structural edit.
4. **Known-owner comparison:** native workbook slicer-cache/timeline-cache references are rewritten from `NativePivotControlDefinition.cacheRelationshipId`; only exact extension URIs, element shapes, and graph-backed IDs can safely be regenerated.
5. **Preservation counterexample:** an unknown workbook `<ext>` carrying a worksheet address was copied raw while structural state was regenerated. The source-byte path correctly retains it unchanged, but the changed-snapshot path must reject rather than carry a potentially stale reference.
6. **Regression reachability:** import and export now both detect workbook extensions; source coverage checks unchanged-byte retention, changed-snapshot rejection, and successful known Slicer/Timeline cache-reference regeneration. Static diff checks remain the only local checks.

Confirmed defect: unknown workbook-level extensions bypassed the fail-close boundary despite raw serialization during regeneration. The source artifact is now scanned before regeneration, and the workbook writer removes/rebuilds only cache-reference extensions validated against the source control graph. No local tests/builds were run; remote CI is pending.

### Six-pass static review — workbook extension repair

1. **Detection callsites:** import and regeneration report paths now both include workbook-level capability detections; the mutation gate consumes source detections before package writing.
2. **Supported control graph:** slicer-cache and timeline-cache `r:id` values are compared with validated `cacheRelationshipId` values, not accepted from URI alone.
3. **Writer ownership:** the workbook serializer drops only source control extensions proven canonical and regenerates them from current controls; unknown extension children remain copied only on lower-level direct writing, while the public export gate rejects them.
4. **Unchanged path:** byte reuse remains unchanged and preserves unknown workbook extensions exactly; it avoids false refusal when the source snapshot has not changed.
5. **Changed path:** unknown workbook extensions now raise `NATIVE_DOCUMENT_UNCHANGED_SAVE_REQUIRED` before regeneration; the regression source includes a sheet-address-bearing opaque child.
6. **Negative/positive balance:** regression source covers successful native control workbook ext regeneration and unknown workbook ext rejection. Static `git diff --check` is required before commit; no local tests/builds are run.

Confirmed defect: workbook-level opaque extensions were omitted from export gating and could be re-emitted with stale structural references. The fix stays within the same PR and uses the source native-control graph as the only bypass condition.

### Six-pass static review — workbook reconstruction boundaries

1. **Workbook child inventory:** `buildWorkbookXml` emits a selected set of root children, while the earlier capability scan did not classify other root children. Unknown parts such as external references could therefore disappear during snapshot regeneration; they now produce `unknown-workbook-node` and block edited export.
2. **Canonical pivot-cache subtree:** `<pivotCaches>` was allowlisted even though its writer rebuilds only canonical cache nodes. Unknown children/attributes in this subtree could be silently dropped; detection now validates the container and each `<pivotCache>` before allowing regeneration.
3. **Attribute ownership:** workbook, `workbookPr`, `sheets`, and `definedNames` are regenerated from a bounded canonical shape. Unsupported attributes/children are now detected. The shared XML-namespace exception was also narrowed from any `xmlns*` spelling to actual `xmlns` / `xmlns:` declarations, so ordinary attributes such as `xmlnsfuture` are not mistaken for namespace declarations.
4. **Visibility enum:** import maps both `hidden` and `veryHidden` to one boolean, while the writer can emit only `hidden`; arbitrary `state` values also became visible. Noncanonical visibility states now fail-close on edit instead of being downgraded.
5. **Workbook-view indexes:** sheet order is reconstructed from the current snapshot, but raw `<bookViews>` indices previously stayed at source positions. `activeTab` and `firstSheet` now follow source worksheet relationship identity through reorder/insert/delete; deletion of the referenced sheet and malformed indices return `NATIVE_DOCUMENT_UNCHANGED_SAVE_REQUIRED`.
6. **Boundary and regression review:** the source-byte shortcut already required an identical snapshot and output-affecting options, but excluded packages with a native Pivot graph; this forced unchanged Pivot packages through regeneration. That exclusion is removed, so exact unchanged saves preserve original bytes, including opaque nodes. Changed snapshots still hit the source capability gate. Regression source covers unchanged opaque pivot-cache preservation plus edited rejection, unowned roots, visibility states, namespace-like attributes, and workbook-view reorder/deleted-view behavior. No tests/builds were run, per the static-only instruction.

Confirmed defects: workbook-root data was silently omitted; `veryHidden` and unrecognized sheet states lost visibility semantics; `xmlns*` attribute classification could hide ordinary attributes from the guard; workbook-view indices became stale after sheet structure changes; and exact unchanged Pivot saves were needlessly regenerated, bypassing source-byte preservation. Each has a fail-close, canonical remap, or unchanged-byte path and regression source. `git diff --check` is the only local validation planned for this batch; PR #345 remains the delivery vehicle.

### Six-pass static review — sparse CellMatrix range traversal

1. **Core traversal:** `CellMatrix.forEachInRange` filtered a `Map` by row bounds but visited every stored row first. The Map is insertion-ordered, not coordinate-sorted, so it cannot stop after crossing `endRow`.
2. **Structural call path:** axis transforms call this iterator to repopulate calculation inputs after moving cells. A small affected suffix still paid for scanning every unrelated row on the sheet.
3. **Move/copy call path:** `getRegion` duplicated the same whole-map traversal for range moves and overwritten-cell snapshots, so narrowing the requested rectangle did not narrow row discovery.
4. **Index capability check:** `SparseAxisBounds` exposes only global min/max; it cannot enumerate rows in an interval. The range API had no range-aware row index despite the sparse model.
5. **Cost boundary:** for hydrated matrices, added a lazily sorted row-coordinate cache and binary lower bound. After cache construction, traversal costs O(log R + Rᵣ + Cᵣ), where Rᵣ is stored rows in the selected interval and Cᵣ is stored cells in those rows; rebuilding the cache after row-key changes costs O(R log R). It still scans columns within matching rows, and deferred-JSON iteration still hydrates before callbacks, so this change does not claim full lazy range access.
6. **Mutation/index coherence:** new-row insertion, last-cell deletion, clear, and existing shift operations invalidate the cache through the canonical `set`/`delete` paths. `getRegion` now shares the range iterator. Regression source covers out-of-order row insertion, deletion, shift invalidation, and row/column output; no tests/builds were run.

Confirmed performance defect: on hydrated matrices, sparse range reads and structural input synchronization were O(total sheet rows/cells) even when the requested row interval was small. The row cache fixes range discovery without changing cell storage or persisted snapshots. Remaining separate cost: deferred-JSON range reads still materialize the matrix, and matching rows still scan their stored columns. Local tests/builds are intentionally not run.

### Six-pass static review — structural row-shift index reuse

1. **Call-chain confirmation:** `StructuralTransform.applyAxis` invokes `CellMatrix.shiftRows` for row insertions and for the surviving suffix after row deletions; the transform computes occupied bounds first but does not enumerate cells to warm the sorted-row cache.
2. **Cold-path cost check:** unconditionally calling `forEachInRange` would sort every stored row on a cold cache, changing a linear Map scan into O(R log R) before moving the affected cells. That candidate was rejected.
3. **Warm-path opportunity:** when a prior range query has already built `sortedRowCoordinates`, the Map scan still visits all prefix rows. The shift can reuse the existing lower-bound lookup and visit only the suffix rows.
4. **Invalidation check:** shifting still snapshots cells before mutation; subsequent `delete`/`set` invalidates the cached row coordinates through the canonical mutation paths, so no stale index is read after the shift.
5. **Coordinate-boundary check:** the indexed query uses infinite endpoints because the preexisting shift iterated all numeric columns and all rows at or after `at`; finite `MAX_SAFE_INTEGER`/zero bounds would have narrowed that behavior for noncanonical direct callers.
6. **Semantic/regression check:** the existing source regression creates rows out of coordinate order, confirms prefix preservation, exercises insertion and deletion, and checks that the row index refreshes. No tests/builds were run, as required for this static-only pass.

Confirmed performance issue: a warmed row index was ignored by `shiftRows`, forcing a full prefix scan. The fix reuses it only when already available; a cold shift retains the original O(R) scan rather than paying an O(R log R) cache build. This is a bounded warm-path optimization, not a claim that every structural row shift is sublinear. The same six checks rejected the unconditional candidate before it entered the patch.

### Six-pass static review — atomic CellMatrix writes

1. **Call-site ownership:** `WorksheetModel` constructs `CellMatrix` with an `onWrite` callback that grows the authoritative `SheetExtent`.
2. **Failure ordering:** `CellMatrix.set` invoked that callback before normalizing `style.fontFamily`; `normalizeFontFamily` throws for whitespace-only or control-character names.
3. **Observable reproduction by source path:** a write to a valid but out-of-current-extent coordinate first enlarged row/column counts, then threw on invalid font metadata. No cell was inserted, and the caller received failure without an extent rollback.
4. **Derived-state consistency:** `count`/occupied bounds therefore still described an empty matrix while the worksheet extent had changed. The same ordering could notify other future write observers before rejecting the payload.
5. **Mutation boundary:** font normalization is pure and does not depend on worksheet extent. Moving it before `onWrite` and row-map allocation makes rejected metadata side-effect free; valid writes still grow extent before storing the cell.
6. **Regression-source balance:** added a rejection case proving no cell or extent change and a succeeding write at the same coordinates proving canonical font normalization and expected growth. Tests/builds remain unexecuted under the static-only instruction.

Confirmed integrity defect: invalid cell metadata could throw after the worksheet extent had already grown. Validation now precedes extent notification and sparse row creation. No snapshot/schema migration. `git -c core.whitespace=cr-at-eol diff --check` is the only local validation for this follow-up.

### Six-pass static review — deferred cell normalization failure boundary

1. **Persisted-entry path:** `WorkbookModel.fromSnapshot` stores each sheet's cell object in `CellMatrix.deferJSON`; font normalization is deferred until a read or write hydrates that sheet.
2. **Canonical gate:** `assertCanonicalWorkbookSnapshot` already inspects persisted cells for metadata invariants but did not reject font values that `CellMatrix.set` cannot normalize. It now checks that same field before accepting a snapshot.
3. **Hydration ordering and peak memory:** `hydrate` previously cleared its only deferred source before calling `set` per cell. A bad font later in the input could throw after earlier cells had been installed, permanently losing the deferred source. It now prevalidates every font before clearing the source, then streams writes; it retains only changed font spellings rather than an O(N) temporary cell-entry array.
4. **Mutation-triggered hydration:** an invalid new `set` payload previously called `hydrate` before validating its own font. It now rejects first, leaving an unrelated lazy sheet unmaterialized.
5. **State/extent check:** only prevalidated cells reach the extent callback and sparse row allocation; the earlier extent-growth rejection case now checks both unchanged bounds and a successful subsequent write.
6. **Regression-source balance:** added persisted-snapshot rejection and deferred-hydration preservation cases. The review rejected a full prepared-entry buffer because it would raise peak memory for large sheets. No local tests/builds were run; static diff check is the only local validation.

Confirmed defect: invalid persisted font metadata could be accepted into deferred state and later make first access partially hydrate and discard its source. The canonical snapshot boundary rejects it; direct `CellMatrix` hydration retains deferred data if normalization fails. No persisted schema change.

### Six-pass static review — canonical worksheet hyperlink ownership

1. **Runtime laziness:** `WorkbookModel.fromSnapshot` first deferred every sheet's cells, then used `Object.values(...).some(...)` to search for legacy links; if any existed it iterated and rewrote every cell in that sheet. This contradicted the lazy-hydration contract and made workbook open cost depend on the full populated range.
2. **Single owner/type review:** `CellData` still declared both old cell-level hyperlink fields while `WorksheetModel.hyperlinks` and `SheetSnapshot.hyperlinks` already represented the canonical owner. The runtime compatibility path therefore remained type-valid instead of being confined to migration.
3. **Migration reachability:** TypeScript upgraded v8 directly to its current revision, and Java treated its current revision as already canonical; neither had a v9→v10 transition. The stored-data upgrade now extracts both legacy fields and deletes them before entering the v10 runtime contract; Java Flyway V11 applies the same conversion to workbooks and checkpoints.
4. **Fail-close validation:** TypeScript and Java accepted legacy fields on cells, while Java did not validate the worksheet hyperlink array, duplicate anchors, coordinates, or internal targets. Canonical validation now requires the collection, rejects old carriers, and verifies anchor bounds, unique coordinates, target shape, and sheet/name references. The wire validator requires the v10 collection without traversing all cell maps.
5. **Precedence/data fidelity:** the former load path applied cell links first and canonical worksheet links afterward, so the canonical record won on coordinate collisions. Migration preserves that precedence without collapsing duplicate canonical entries; the validator rejects duplicate canonical anchors rather than silently selecting one.
6. **Producer/consumer audit:** the generated TypeScript/Java contract, model snapshots, OOXML import output, protocol fixtures, service fixtures, and each database migration dialect were checked together. All current producers emit v10 and an explicit hyperlink array; the explicit v9 migration is the only path that reads the old cell fields.

Confirmed defects fixed: eager per-sheet materialization during model construction; legacy hyperlink fields still admitted by the canonical type; missing v9→v10 conversion on browser/server persistence paths; absent canonical rejection of stale cell carriers; and incomplete persisted hyperlink reference validation. Regression source covers legacy URL/detail conversion, canonical-over-legacy precedence, deferred-sheet preservation, stale-field rejection, dangling targets, and duplicate anchors. Local tests/builds were not run. Static `git diff --check` is the only local validation in this turn.

Migration note: V11 rewrites stored workbook/checkpoint snapshots and checkpoint checksums through the existing explicit migration boundary. The old-to-new extraction is one-way; rollback requires restoring the database backup taken before V11 and deploying the prior application version. A v9 browser-storage migration scans legacy cell maps once to discover old links; ordinary `WorkbookModel.fromSnapshot` no longer scans or hydrates cells to discover them.

### Six-pass self-review — hyperlink consumer convergence

1. **Render ownership:** the canvas projection still fell back to `CellData.hyperlinkDetail` after reading the worksheet map. This preserved a second runtime owner and let malformed in-memory cells render links absent from canonical snapshots. The fallback is removed; a source-only legacy field no longer produces a UI hyperlink.
2. **Clear and inverse cost:** metadata-only clear still captured and replayed every cell in the selected range, hydrating deferred data and rewriting unchanged cells on undo. Hyperlink/comment clear snapshots now omit cell state; local and Java inverse reducers leave the cell owner untouched when it is absent. Cell-affecting clear operations use the range iterator instead of traversing all stored cells.
3. **Clipboard semantics:** Paste Special still copied legacy link properties into cell payloads despite clipboard capture and paste planning already carrying worksheet hyperlink metadata independently. Removed that path so only the canonical metadata transfer can paste hyperlinks.
4. **Data-block patch boundary:** `CellPatch` still declared and accepted both deleted cell hyperlink fields, causing invalid indexed `CellData` types and permitting overlays to create a competing owner. Removed both fields from the type and patch-field allowlist.
5. **Wire/schema parity:** the schema closes hyperlink target unions, but manual validators admitted undeclared properties; Java also accepted numeric link IDs, padded target discriminators, and stale `definedNames` projections when canonical `definedNameModels` existed. TypeScript and Java now enforce closed shapes and canonical name authority, with static regression source for each rejected payload.
6. **OOXML reference normalization:** the importer accepted absolute/lowercase A1 addresses during parsing but preserved their raw spelling, then produced a snapshot rejected by the canonical target validator; sheet-name lookup was also case-sensitive even though Excel sheet references are case-insensitive. Imported locations now resolve names case-insensitively and canonicalize addresses before snapshot validation.

All six findings were confirmed against producer/consumer code paths before modification; speculative candidates were not counted. Regression source was updated, but tests/builds/browser verification remain unrun per the static-only instruction. Static contract generation/checking and whitespace validation are recorded separately for this batch.

#### Supplemental fail-close review

A subsequent audit found metadata-only inverse snapshots intentionally omit `cells`, but the inverse mutation did not carry its clear family. That made omission indistinguishable from a truncated cell-clearing undo payload. The backend also removed notes/threads before treating missing metadata arrays as absent, and could silently skip rule restoration when formats/all snapshots omitted rule arrays. The inverse now carries the family; both runtimes enforce the exact cell/rule snapshot shape and require note, hyperlink, and comment arrays before applying any restoration. Regression sources cover the accepted metadata path and rejected missing-cell/missing-metadata paths. No tests/builds were run.

### Six-pass self-review — reference-transform domain

1. **Duplicate coordinate owners:** the core model and Java reducer each had a point helper, while TypeScript AST and Java formula code each duplicated interval algebra. Removed those helpers; point and interval mapping now route through `ReferenceTransformDomain` in each runtime.
2. **Formula AST bypass:** the formula AST still implemented insert/delete point semantics inline even after interval mapping moved to the domain. Replaced that branch with the domain point mapper while preserving `#REF!` generation for deleted or out-of-range formula references.
3. **TypeScript owner overflow:** structural owner transforms accepted shifted points beyond the Excel row/column limit; interval callers also treated out-of-bounds as deletion, so print titles or individual rule ranges could disappear. Point and interval mapping now distinguish `mapped`, `deleted`, and `out-of-bounds`; structural metadata preflight rejects the last case.
4. **Java metadata overflow:** the snapshot reducer's point helper returned an integer without applying row/column maxima, while interval wrappers returned `null` for both deletion and overflow. The axis-aware reducer and formula-range boundary now return validation failures on overflow.
5. **Deletion/overflow ambiguity:** the former nullable/sentinel results could not distinguish deleted points/ranges from references moved outside the worksheet. Added distinct point and interval outcomes; formula references retain `#REF!` behavior, OT rejects out-of-bounds state, and structural metadata fails closed instead of looking deleted.
6. **Bounds ownership:** command history and collaboration each carried their own worksheet maximum constants. They now consume the formula-engine domain constants and handle its typed point outcomes; repository-wide source search found no references to the removed point helpers or duplicated interval transformer.

All six findings were traced to executable call sites before counting. TS and Java regression source covers insert/delete vectors, deletion, point/range overflow, interval boundaries, and invalid inputs. As requested, no tests, builds, or browser checks were run; only static source and diff checks are used for this slice.

### CI correction — TypeScript compile diagnostics

1. **Remote evidence:** the PR's Linux and Windows `canonical-build` jobs reported the same three TypeScript diagnostics; no local build was run.
2. **Deleted module boundary:** the transform refactor removed `axis-coordinate-transform.ts` but left its core-model barrel export. Removed the stale export and searched the frontend packages for remaining references.
3. **Type ownership:** `ClearFamily` is declared and exported by `clear-planner.ts`; the command restore payload used it without importing the type. Added the direct type import from that owner.
4. **Accessor contract:** `CellMatrix.isHydrated` is a boolean getter, as confirmed by its declaration and existing callers. Corrected the three new assertions that invoked it as a method.
5. **Call-site countercheck:** direct structural-transform tests either build a `RangeIndex` in their local adapter or supply the required owner index; no missing third-argument migration was found in production callsites.
6. **Patch boundary:** the changes are limited to the stale export, missing type import, and three getter assertions; `git diff --check` passed. Local tests/builds remain intentionally unrun, and GitHub CI must rerun on the pushed commit.

Confirmed defects: one stale export, one missing type import, and one accessor misuse repeated at three assertions. These are counted by root cause, not as five independent runtime defects.

### Six-round static review — collaboration mutation payload coordinates

1. **Range-write producer/consumer:** `range.set` declares its write origin as `startRow`/`startColumn`, while generic OT point inference only remapped `row`/`column`; the affected range moved but replay kept writing at the old origin. Rebase now maps the exact origin and rejects a delete/split intersecting the implicit value grid.
2. **Fill write authority:** `fill.applied` and `fill.restored` replay `writes[].row/column` after validating against `sourceRange`/`targetRange`. The generic walker shifted both ranges but not the write records. Rebase now maps every write and rejects deletion through either source or target.
3. **Find patch authority:** `find.replaced` replay reads `patch.match.row/column` to write cells, notes, and locate comments; only the nested range and affected ranges were shifted. Match points and their canonical keys now move together.
4. **Per-sheet formula ownership:** workbook-scoped find patches can contain matches on sheets other than the mutation envelope's sheet. Formula snapshots previously inherited the envelope identity and rewrote unqualified references against the wrong worksheet. Each patch now uses its match's sheet identity.
5. **Nested comment anchor:** `comment.add` validates outer row/column against `thread.row/column` and applies the embedded thread. OT moved only the outer address, making the rebased payload invalid or stale. Both coordinates now follow the same point mapping.
6. **Formula provenance and unsupported groups:** `formulaMetadata.sourceFormula` is a structural formula owner but is not a field named `formula`; group ranges/preserved-only formulas have no canonical OT mapper here. Normal provenance and barcode formulas are transformed from their real owner; unsupported formula-group snapshots now fail closed.
7. **Cell write-authority duplicate address:** canonical `cell.set` carries the same row/column inside `writeAuthority.target`, and its replay validator requires that address and candidate to equal the top-level mutation. Mapping only the top-level cell left the authority stale; both authority target and formula-bearing candidate now move with the write.
8. **Local-only mutation classification:** data-region materialization is declared `durability: local`, `remote: false`, and `rebasePolicy: none`, but was listed as a collaboration `cell-value` kind despite carrying absolute cell coordinates. It is no longer assigned a remote rebase kind; direct attempts fail closed as unknown.
9. **Typed snapshot anchors:** removing field-name coordinate inference also removed the only mapping of pasted validation/conditional-format `formulaAnchor` points. These anchors now transform explicitly inside the two canonical rule collections; the existing source regression fixture covers both axes.
10. **Snapshot formula provenance:** pasted cell snapshots contain `CellData`, including `formulaMetadata.sourceFormula`, while only ordinary formula keys are found by recursive formula handling. Paste snapshots now transform cell formula owners from the original immutable payload and fail closed on unsupported formula groups.
11. **Detached clipboard identity:** a pending copy's source `clipboard` is a captured payload, not a live range to rewrite. The generic range walk moved `clipboard.range` while retaining `sourceExtent`, potentially making the mutation schema invalid after deletion; rebase now leaves the source clipboard snapshot unchanged and transforms only the replayable target snapshot.
12. **Hyperlink destination references:** a paste snapshot stores both the hyperlink's source-cell key and its destination. The key moved with the destination range, but same-workbook worksheet targets remained stale. A1 and row/column target forms now map by target sheet identity, and deletion of the target fails closed.
13. **Sparse paste footprint deletion:** a pending paste's rectangular target extent can include blank cells absent from its sparse cell snapshot. Deleting a row inside the footprint shrank ranges but left target origin and source extent unchanged; rebase now rejects any deletion intersecting the canonical target footprint, including blank rows.
14. **Rebase/test contract agreement:** two paste fixtures used `transfer: move` and source snapshots while expecting rebase, although move-range is explicitly unsupported without a canonical patch. The success fixture now exercises copy-only paste; the move fixture asserts fail-close.

Confirmed runtime defects: thirteen distinct coordinate/formula-owner/classifier paths across seven mutation payload families. A separate regression-fixture mismatch was aligned with the fail-close contract. Six iterative review rounds compared generated durability policy, classifiers, producers, replay consumers, snapshot schemas, duplicate authorization state, detached payload identity, hyperlink references, sparse footprints, and rejection boundaries. No local test/build was run.

### Six follow-up self-audit rounds — review and hyperlink mutations

1. **Mutation registry/classifier coverage:** review commands register nine additional mutation IDs (`note.*`, `hyperlink.*`, and comment reply/resolve/remove), but none had a collaboration kind. They were rejected as `unknown` before any structural transform; all registered review mutation IDs now classify explicitly.
2. **Cell-owned review coordinates:** notes and hyperlinks are attached to top-level `sheetId/row/column` fields, not `RangeRef` values. Their source cell and affected range now move together; deleted source cells fail closed.
3. **Stable comment identities:** reply, reply removal, resolve, and thread removal use stable thread IDs rather than cell coordinates. Their payloads remain unchanged while their declared conflict ranges move with the thread anchor.
4. **Hyperlink target ownership:** a worksheet hyperlink has a second address owned by its target sheet, independent of the source cell. Both A1 and row/column target forms now follow a structural edit on the target sheet, including when the source is on another sheet.
5. **Inverse-path symmetry:** review command producers create inverses using the same mutation IDs (`note.set/remove`, `hyperlink.set/remove`, and comment add/remove/reply pairs). The classifier and explicit cell/target transforms now cover both forward and inverse payload shapes.
6. **Deletion boundaries:** a structural deletion of a linked destination is not a valid destination remap. The target-address transform uses the same bounded point mapping as cell writes and rejects a deleted destination instead of retaining a stale hyperlink.

This follow-up pass confirmed ten additional operation-specific stale/rejected paths: nine review mutation IDs lacked a rebase classification, and direct internal hyperlinks had an untransformed target address. Static regression cases cover cell coordinates, stable comment IDs/conflict ranges, cross-sheet hyperlink targets, and deleted targets. Local tests/builds remain intentionally unrun.

### Seventh static audit cycle — visibility-state indices

The protocol marks `rows.visibility` and `columns.visibility` remote with range rebasing, and the command producers persist the selected row/column identifiers in `states[]`. Neither mutation had a collaboration kind after generic coordinate inference was removed, so pending visibility changes were rejected as unknown. They now have an explicit visibility kind and transform only their matching axis; deleting a selected row/column fails closed. Source regression cases cover both-axis insertion and deletion rejection.

Confirmed additional operation paths: 12 (nine review mutation IDs, direct internal hyperlink destinations, and row/column visibility). Across both collaboration review sections, this slice records 25 operation-specific static defects plus one fixture-contract mismatch; these are path counts and include shared classifier causes rather than claiming 25 unrelated root causes. No local tests/builds were run.

### Eighth static audit cycle — remote range-policy coverage

Compared the generated remote `rebasePolicy: range` mutation contracts against both `mutationCapability()` and the app classifier. `cell.editor.set` was the only remaining range-policy mutation without a collaboration kind; its producer and reducer both use `RangeRef[]`, which the canonical range mapper already transforms. It now classifies as `cell-style`; the source regression case checks an insertion before the edited rectangle. Exact-policy remote operations remain outside this mapping and fail closed where no structural contract exists.

Confirmed additional operation paths: 13 in the follow-up audit (the previous 12 plus `cell.editor.set`). The cumulative static audit records 26 operation-specific paths and one fixture-contract mismatch, not 26 unrelated root causes. No local tests/builds were run.

### 六轮自审 — preserved-only 公式缓存值

1. **导入边界：** OOXML 导入保留不可计算公式单元格的缓存 `value`，并同时保留 `formulaMetadata.preservedOnly`；该值是合法的计算输入，不是公式引擎应重新计算的结果。
2. **输入分类：** `calculationInputUpdate` 已把 preserved-only 公式映射为值输入，说明运行时契约明确要求公式引擎可读取其缓存值。
3. **初次装载：** `loadFormulaInputs` 将这类公式登记为结构引用 owner，但紧接着的普通值扫描按 `formula !== undefined` 排除了同一单元格；依赖它的可计算公式因此读不到缓存值。
4. **延迟建图路径：** 值-only 工作簿首次新增可计算公式时，`synchronizeCellMutation` 会补载其他值，但原过滤条件同样跳过 preserved-only 公式单元格，复现同一根因。
5. **消费者确认：** `FormulaEngine.synchronizeInputs` 对 `kind: 'value'` 调用 `loadValue`，并保留后续公式重算根；缺失发生在输入收集，不是 evaluator 的单元格读取或重算调度。
6. **改动边界复核：** 初次 `loadFormulaInputs`、首次出现公式的 `synchronizeCellMutation`、对应的结构变更同步，共三个值收集路径都改为复用 `calculationInputUpdate` 的 canonical 分类，仅追加值输入；普通空单元格仍不复制，公式输入仍由公式路径装载。新增两条回归用例分别覆盖首次 hydration 与首次公式创建。

六轮交叉审查确认的是一个真实根因、两个受影响入口，不把重复表现计作独立缺陷。只做静态检查；本轮未运行测试、构建或浏览器验收。

### 六轮静态审查 — TS/Java 结构映射向量契约

1. **运行时入口：** TypeScript `ReferenceTransformDomain` 接收 `direction`，Java 对应入口接收 `insert`；向量以 `operation` 表达，再由各自测试适配到真实 API，不在生产运行时增加跨语言桥接。
2. **坐标域：** 行与列最大索引来自两侧各自的公开常量；向量只保存轴，不重复维护一个可能漂移的 `maximum` 数值。
3. **点映射结果：** 共同数据覆盖未受影响、插入后移、删除、越界和已有越界点；复审发现 Java 原先用 `-1` 表示删除坐标，而 TypeScript 的 `deleted` 结果不含坐标。
4. **闭区间语义：** 区间数据覆盖插入前/内/后、删除前/相交/全删、列轴和反向输入端点；输出按映射域规范化为闭区间。
5. **加载路径：** 前端测试从仓库根 `contracts` 读取同一 JSON；Maven 仅将该文件作为 test resource 加入类路径，同时保留原 `src/test/resources`，不改生产资源或运行时。
6. **漂移修复：** 两个 Java 消费点先检查映射 kind，再读取坐标；将 Java deleted 坐标改为 `null` 后，不引入空值解引用，并与 TypeScript 的无坐标结果对齐。共享向量现在直接检查这一语义，不再把 `-1` 当作跨语言契约。

首组六轮静态复审确认并修复点映射结果的跨语言契约差异；随后补充审查又确认区间删除结果存在相同类别的表示漂移（详见下节）。此契约只统一并锁定坐标映射样例，不宣称 Java/TypeScript 已共享实现，也不替代仍待完成的不可变 `StructuralPatch` 与端到端消费者迁移。测试源码已接线，但按静态审查要求未执行测试或构建。

### 补充六轮静态复审 — 删除区间结果契约

1. **返回形状对照：** TypeScript `IntervalTransformResult` 的 `deleted` 分支不含端点；Java `IntervalMapping` 对所有 kind 都携带原始类型端点，删除时以 `-1,-1` 填充。
2. **消费点追踪：** Java 两个生产消费点都只在 `MAPPED` 时读取端点；没有空值解引用故障，但这不消除运行时 API 结果形状与类型契约的差别。
3. **测试掩盖定位：** 共享向量测试曾把缺少 `start/end` 自动转成 `-1`，使两种不同表示通过同一断言；这是确认契约漂移的直接证据。
4. **删除语义核对：** 只有映射结果为空的全删区间生成 `DELETED`；把端点改为 `null` 不改变“调用方按 kind 分支”的既有控制流。
5. **其他结果核对：** `MAPPED` 和 `OUT_OF_BOUNDS` 仍保留真实端点，最大行列边界及规范化后的区间计算不变。
6. **修复边界复核：** Java 改为可空端点并在删除分支返回 `null,null`；共享向量测试按字段存在性构造期望值，直接比较 TS/Java 相同的删除形状。未修改公式改写和结构变更消费者。

补充六轮复审确认第二个可复现的类型契约差异（删除区间坐标哨兵与无坐标联合分支不一致），并在同一结果域内修复。生产消费者现有 kind 检查保持安全；静态检查之外未执行测试或构建。

推送后远端编译门禁进一步暴露点映射 `Long` 装箱后的一个真实编译错误：公式转换器仍使用不能应用于 `Long` 的直接 `(int)` 强转。现改为仅在 `MAPPED` 分支执行 `Math.toIntExact`；本地仍不运行构建或测试，等待新提交的远端门禁确认。

随后远端门禁通过编译但发现共享向量测试未将 JSON 的 `out-of-bounds` 规范化为 Java 枚举 `OUT_OF_BOUNDS`，造成一个测试错误。点与区间向量解析现都先将连字符转换为下划线；等待包含此修复的新门禁结果。

## 当前源码全链路与复杂度审计（2026-09-25）

本节按 `origin/main=a2a6140a` 与结构整改 PR 的当前源码记录入口和已有边界，不以文档中的目标设计代替实现事实。复杂度是按数据结构及循环静态推导，未做性能基准。符号：`M`=全 workbook 的结构元数据对象数，`C`=本次涉及的已物化单元格数，`F`=被查出或改写的公式引用 owner 数，`V/E`=公式/值输入数及依赖边数，`N/K`=排序行数/排序键数，`P/R`=待重放/已提交 mutation 数，`S/B`=JSON snapshot 大小/原生包字节与 parts 大小。

| 操作 | 当前入口与运行时 owner 路径 | History / Collaboration / Server / OOXML 边界 | 静态复杂度与已证实缺口 |
|---|---|---|---|
| Insert/Delete Rows/Columns | `sheet-features/data-features.ts` 发出 `rows.*`/`columns.*`；客户端进入 `StructuralTransform.applyAxis`，回传清除/重载公式输入范围及改写 owner；`runtime.ts` 增量同步并失效投影。 | 删除 inverse 另外快照已占用 cells；OT 对轴变更有独立 `StructuralDelta`；Java `StructuralMutationDescriptor` 分派 `StructuralSnapshotReducer`；OOXML 在保存时另作 owner 检查。 | 客户端至少 `O(M+C+F)`；`preflightAxisMetadata` 对每个 sheet 深拷贝结构 metadata，临时内存 `O(M)`，随后又逐 owner 应用。TS 与 Java 仍各自计算完整 owner 变换。 |
| Insert/Delete Cells | `sheet-features/editing/index.ts` 生成 `CellShiftPlan`，再由 `applyCellShift` 改动 cell band 和 metadata；命令保存 inverse snapshot。 | 服务端有独立 cell-shift reducer；OT 仅对登记过的 cell/range mutation 执行参数变换；历史以 cell snapshot 与 inverse mutation 重放。 | 客户端约 `O(M+C+F)`，临时空间包含全 metadata staging 与 band cells；范围和公式 owner 由不同机制检索，尚非一份 patch。 |
| Move / Copy / Cut / Paste | `range.move` 进入 `applyMoveRange`；`range.paste` 走 `applyPasteSnapshot`，剪贴板/填充公式另用 formula offset 和数据命令。 | 移动类 history rebase 失效；Java 对跨 sheet `clearSource` 明确返回 `UNSUPPORTED_FEATURE`；OT 不猜测 move/cut 坐标。 | Move 随源/目标 cells、相关 `F` 和检查过的 metadata 增长；Paste 随 payload cells/metadata 增长。两路尚未汇合到统一逆 patch，跨 sheet cut/paste 不支持。 |
| Drag / Fill | command runtime 对公式使用 `offsetAst`，再生成普通 cell/range 写入；不是 `StructuralTransform` 操作。 | 历史记录 cell/range mutation；远端按具体 mutation schema 重放。 | 与输出 cell 数 `C` 及每个复制公式引用数成正比；同类坐标偏移与结构插入仍由不同 API 定义。 |
| Sort | `data.sort.rows` 计算 `sourceRows` 后发出 `rows.permuted`，客户端走 `applyRowPermutation`（独立于 `StructuralTransform`）。 | runtime 用排序范围同步计算；mutation metadata 将 row permutation 标为 history-rebase invalidate；Java reducer 有另一套 permutation 逻辑；OT 对 sort fail-close。 | 排序本身约 `O(N log N × K)`，随后还要映射 cells/owners；执行、inverse/history 与远端 OT 均没有共用 permutation patch。 |
| Sheet Rename / Delete / Reorder | `sheet.rename/remove/add` 经 `WorkbookModel` 与 `SheetIdentityTransform`；reorder 直接改 sheet identity/order。 | rename 有精确公式 owner delta，但存在新名称解析歧义时重建 FormulaEngine；add/remove/reorder 在 `CALCULATION_CONTEXT_REBUILDS` 中全量重建；OT 对 sheet identity fail-close。 | 身份变更需要遍历 workbook owners，约 `O(M+F)`；重建另需枚举 `V` 与依赖边 `E`。history、OT 和 server 并未消费同一个身份 patch。 |
| Table Resize | 未找到独立 `table.resize` mutation；sheet table 通过 `sheetTable.update`，workbook table 通过 `table.add/remove`，轴插删又各自改 table ranges。 | 各命令有各自 inverse；OT 将 table-resize 列为无 canonical patch；服务端/OOXML 以另一组 table owner reducers 处理。 | 专用跨层 resize 事务和统一 range/reference delta 缺失；当前局部代价取决于被扫描/复制的 table 与 filter owners。 |
| Undo / Redo | CommandRuntime 重放 mutation 的 inverse/redo mutation；没有记录 `StructuralTransformResult` 的完整 owner before/after 值。 | row delete 保存 removed cells；move、sort 等被标记不能跨结构历史安全 rebase；server 接收到的是 mutation 序列而不是历史 patch。 | 回放成本为对应操作成本再加 inverse payload；delete/move 的快照空间随被保存对象数增长。inverse 语义由每种命令重复维护。 |
| Remote Replay | `ot-rebase.ts` 以 `rebaseAgainstHistory` 顺序叠加已提交操作；当前只对轴 delta 做完整坐标映射，其他结构 kinds 拒绝猜测。 | `transformParams` 仍递归检查字段名并结合 mutation-specific 转换；move/sort/table resize/sheet identity 已 fail-close；server commit 与本地 replay 各自执行 reducer。 | 单个 pending 跨 `R` 条 history 是 `O(R × payload)`；`P` 个 pending 约 `O(P×R×payload)`。缺少统一 patch 是支持被拒绝与字段推断仍存在的共同根因。 |
| Server Commit | `MutationDescriptorRegistry.applyPublicMutations` 顺序调用 descriptor；结构 descriptor 对每次 mutation `snapshot.deepCopy()` 后由 Java reducer 执行。 | 同一 batch 的后续 mutation 接收前一结果；Java `FormulaReferenceTransformer` 和 TS AST transform 是独立实现。 | `K` 次结构 mutation 的 snapshot copy 上界为 `O(K×S)` 时间/累计分配，再加每次局部 reducer；尚无 server 可消费的共享 StructuralPatch。 |
| OOXML Structural Save | `exportOoxmlDocument` 对 snapshot hash 匹配的未改文档回传原 bytes；已改文档经 package graph、能力检测、serializer、输出重载检查。 | 部分 unknown nodes/extensions/未建模 chart 会阻止重建；其余 opaque parts 由导入 artifact 保留，包层并非结构 planner 的参与者。 | untouched 路径仍需计算 `O(S)` hash 和复制 `O(B)` bytes；改动后成本涉及 `O(S+B)` 序列化/扫描。未知 part 是否含坐标 owner 不是所有结构 mutation 的前置检查。 |

### 六轮交叉审计结论与收敛方案

1. **入口审计：** Rows/columns、cell shifts、move、row permutation、sheet identity、table updates 来自不同 mutation/command；并不存在一个涵盖上述 kinds 的 planner 入口。
2. **副作用边界：** `StructuralTransform.apply` 在 live `WorkbookModel` 上写 cells/metadata，`StructuralTransformResult` 只携带 calculation deltas 和少量标记，不是可复用的 immutable patch。
3. **Owner 查询：** FormulaEngine `ReferenceIndex` 覆盖公式依赖/公式 owner；其他 metadata owners 由结构函数遍历/深拷贝，无法用一个 owner index 查询并产生变更集。
4. **计算与投影：** runtime 以 mutation ID 集合、临时 structural effect、affected ranges 三种信息决定 rebuild、计算同步和投影失效；这些 effect 不随持久化/历史 mutation 一同保存。
5. **历史、协作、服务端：** undo inverse、OT 的字段推断/明确拒绝、TS mutation reducer 和 Java snapshot reducer 分别编码语义；共享点/区间向量只能约束原子 mapping，不能保证各 owner 结果一致。
6. **OOXML 与失败原子性：** OOXML 的部分 opaque-owner 拒绝发生在 export 而不是 planning；planner 若不先纳入 owner capability，用户可能先成功修改内存，再在保存时才遇到 fail-close。

收敛顺序据此固定为：先让结构请求经 typed owner index 进入 side-effect-free `CanonicalStructuralPlanner`，输出包含 cell、metadata、formula、projection、history/inverse 和协作影响的不可变 `StructuralPatch`；再由客户端 runtime、OT、Java commit 和 OOXML capability boundary 消费同一版本化 patch 语义。第一条纵向迁移应覆盖 whole-axis insert/delete（成功、拒绝、inverse、remote replay、server reducer、native save），完成后移除对应旧的 live-mutation 分支，而不是增加并行 wrapper。复杂度目标是 planning/commit 随 affected cells 与 affected owners `O(C+F+M_affected)`，避免每次全 workbook metadata clone `O(M)`；无法证明 opaque owner 不受影响时在计划阶段 fail-close。

以上是源码级复杂度推导，不是 benchmark，也未运行本地测试/构建。PR head `4034ae14` 的两项远端 `canonical-build` 均成功。完整 StructuralPatch/owner-index 迁移仍未完成。

### Spill 障碍快照与执行语义 — 六轮自审

1. **合并区域路径：** 主线程 `createSpillEnvironment` 曾通过 `sheet.isMerged` 阻止 Spill；Worker snapshot 只包含占用单元格坐标，`fromCalculationSnapshot` 重建时没有合并几何，故同一数组公式在线程内外可产生不同结果。
2. **Table 路径：** snapshot 虽另存 `sheetTables`，Worker Spill resolver 并不查询该列表；原 `isOccupied` 闭包中的 Table 范围无法跨 Worker 边界，Table 内的空白单元格因而错误地允许 Spill。
3. **既有 Spill 路径：** Workbook snapshot 持久化 `spillRanges`，主线程占用闭包会将它们当障碍；计算快照没有活动 Spill 状态，Worker 重建后遗漏其它公式的投影占用。
4. **自身投影路径：** 重建引擎时先加载的模型仍带着已保存 `spillRanges`，而公式引擎重新解析该锚点后会检查旧子格；既有公式可能在首次重算时把自己的输出判成阻塞。新环境不再把模型输出投影误作 authored occupancy，Resolver 明确排除同一锚点的旧 Spill。
5. **历史预览路径：** `hydratePreviewFormula` 原先只以 `sheet.cells.get(...)` 判占用，语义不同于正常运行时，且没有合并/Table/动态 Spill 统一边界；预览现复用主运行时创建的 Spill environment。
6. **同批重算路径：** 多个受影响 Spill owner 按地址依次重算时，较早 owner 曾会被较晚 owner 的旧范围阻塞，即使后者在同一轮会缩小并释放该区域；现在仅尚未处理的受影响旧投影不参与冲突，新投影仍按稳定顺序互斥。
7. **性能与协议路径：** 原几何检查在候选 Spill 的每个子格重复扫描 merges/tables/spills，成本随输出面积与障碍物数量相乘；现在范围障碍每个候选只做一次范围相交，普通 authored cell 仍用坐标查询，且快照校验/Worker 协议版本一同升级。
8. **失败结果投影：** `spillValueAt` 对 `blocked` Spill 曾只将锚点变为 `#SPILL!`，仍向子格读取者返回被拒绝矩阵中的值；现在 blocked 状态只暴露锚点错误，子格不产生任何值。
9. **边界溢出：** `spill-error` 曾把越过工作表 extent 的矩阵裁到可见范围并显示首值，导致越界结果部分落地；现越界锚点返回 `#SPILL!`，范围内子格也不读取截断矩阵。
10. **障碍恢复路径：** runtime/session 曾把 `blocked` 或 `spill-error` 的意向范围当作只读 spill 子格，包含导致 `#SPILL!` 的 authored blocker 本身，用户无法清除障碍；只成功投影的 Spill 子格现在才只读。
11. **What-If 投影路径：** Goal Seek/Scenario 的 `isSpillCell` 仍以状态无关的矩形包含判断拒绝 blocked Spill 的空闲子格；锚点继续受保护，但只有成功投影的子格现在会被当作 Spill cell。

本轮修复把合并区、Table 区和成功 Spill 范围作为范围障碍传输/查询；Worker 快照同时恢复活动 Spill 投影，静态快照校验拒绝无公式锚点的投影；阻塞或越界 Spill 不再向子格泄漏矩阵值、显示部分数组或锁住障碍单元格，What-If 允许写入未投影的空闲子格。新增了范围阻塞、快照往返、自身投影重算、同批 owner 更新、失败子格不投影、越界拒绝、障碍恢复、What-If blocked-range、工作簿重建和合并/Table 环境回归用例。按用户要求仅静态审查，未执行测试或构建；完整 StructuralPatch/owner-index 跨层迁移仍未完成。

### 六轮跨层自审 — 轴结构变更的保护范围

1. **客户端命令入口：** 行/列插入删除命令以请求的 `at/count` 构造 axis band，列维度取当前 `rowCount`，行维度取当前 `columnCount`。
2. **权限语义：** `features/permission/policy.test.ts` 已明确验证：删除被保护列应拒绝，而删除与其不相交的列应允许；这排除了“任一结构变更都按整张表授权”的当前客户端语义。
3. **本地执行：** `WorkbookSession` 的 mutation guard 将命令给出的 `affectedRanges` 传入 `PermissionService.checkMutation`；因此客户端只按选中的轴 band 检查范围保护。
4. **服务端范围：** `StructuralMutationDescriptor.affectedRanges` 对四种 rows/columns 插入删除 mutation 曾统一返回 `wholeSheetRange`，不读取 `at/count` 来界定受保护目标。
5. **服务端授权：** `MutationDescriptorRegistry.prepare` 用该范围调用 `ProtectionResolver`；range-scope 的锁定规则只需与范围相交即可拒绝，所以位于未操作列/行上的锁会让服务端拒绝客户端已允许的操作。
6. **失败边界：** 结构 mutation 的服务端 rebase policy 是 `EXACT_BASE`；全表范围并非并发重放保护所必需，却改变了保护判定语义并造成跨端结果不一致。

**确认缺陷与修复：** 服务端现从结构 mutation 的 `axis/at/count` 生成与客户端相同的请求轴 band，并在权限判断前验证工作表 extent 与 Excel 坐标上界。Java 回归用例覆盖不相交的锁列可删除、相交锁列拒绝及行插入范围。该源码测试尚未执行；本地测试/构建仍按要求跳过。

### 六轮自审 — 删除轴变换的公式历史可逆性

1. **引用语义：** 删除覆盖公式引用时，`mapAstStructuralReferences` 生成 `invalid-reference`（`#REF!`）；该 AST 分支后续结构变换保持不变，因此反向插入不能推回原引用。
2. **Owner 覆盖：** `preflightFormulaRewrite` 从结构引用索引取出受影响公式 owner，确认这种情况不是漏扫公式，而是 owner 已被找到并改写。
3. **计划数据：** `FormulaRewritePlan.cells` 仅记录新公式字段和原坐标；没有可持久化的 before/after owner 值供历史消费。
4. **提交效果：** `applyFormulaRewritePlan` 改写活动 workbook 并返回 owner 地址；`StructuralTransformResult` 不包含这些 owner 的可逆数据。
5. **命令逆序：** 行/列删除命令的 inverse 只包括反向插入和被删轴带内单元格恢复；带外公式 owner 不在快照中。
6. **历史重放：** `CommandRuntime.applyMutation` 把 mutation 声明的 inverse 写入 history；effect 只传给 mutation listeners。Undo 因而只重放轴插入和已声明快照，无法还原变为 `#REF!` 的公式。

**确认缺陷：** 删除包含直接公式引用的行/列后，Undo 会恢复工作表地址空间和被删单元格，但不会恢复存活公式 owner 原先的引用。相同缺口适用于由结构变更改写的其他公式型 owner；修复验收不能只覆盖 cell formula。

**方案约束：** 禁止在命令层事后拼接 `cell.restore`。它会把结构副作用变成普通单元格写入，造成权限动作/范围不匹配；若不把 owner 影响范围并入 history conflict keys，远端对公式 owner 的并发写还可能被 Undo 覆盖。服务端 `OperationMutation` 当前只有 mutation id、sheet id、params，并由 Java reducer 独立重推结构语义，所以仅修改客户端历史也不能恢复持久化及其他协作者状态。

**服务端权威边界：** `OperationEnvelope` 是客户端提交的 intent；`WorkbookOperationService` 逐项 prepare/apply 后才生成 `CommittedOperationMutation`，其中目前只有原 mutation、服务端 affected ranges 和 revision。幂等检查比较 request intent，不应把服务端产生的 patch 当作可由客户端伪造的请求字段。因而在线结构提交应由服务端从精确 base revision 规划并持久化 patch；committed envelope 返回 patch 与 owner-impact conflict keys。客户端 optimistic preview 可继续使用 TS planner，但 ACK/remote replay 必须消费服务端 patch；hash/precondition 不匹配时停止并重载，不再运行第二套 reducer 修补结果。离线队列只保存结构 intent，重连时必须针对当前 revision 重新规划，不能重放旧 patch。

**下一步实施契约：** 把 whole-axis insert/delete 迁移为版本化 `StructuralPatch` 事务。Patch 必须在提交前包含 typed owner locator、before/after、坐标映射、授权目标范围与冲突范围，以及 calculation/projection/history 影响；逆操作由同一 patch 产生。服务端对结构 intent 做唯一权威规划并提交 patch，客户端本地预览与协作重放消费已提交 patch；拒绝未知 owner、旧 patch 版本、失配 revision 或不完整 owner delta，而不是各层重新解析 mutation 再补救。随后把同一消费入口扩到 cells/move/permutation 与 OOXML capability preflight。此结论是静态设计约束，不表示相关实现已完成；本轮未运行本地测试或构建。

**代码迁移边界（由当前调用链确定）：** 将 `MutationDescriptor.apply` 改为返回包含 snapshot 与可选 `StructuralPatch` 的单一 `MutationApplication`；`MutationDescriptorRegistry` 只传递该结果，不再另设结构专用旁路。`StructuralMutationDescriptor`/`StructuralSnapshotReducer` 负责 axis patch 规划，`WorkbookOperationService` 在同一 revision transaction 内提交 patch，并将其放入 `CommittedOperationMutation`。前端 protocol validator、`CollaborationSession` 和 `CommandRuntime.HistoryEntry` 统一消费 committed patch；`assertOperationResultMatches` 仍只比较请求意图，不能允许客户端控制服务端 patch。任何结构 mutation 未返回完整 patch、patch base revision 不匹配或 owner precondition 不符，都必须在提交前 fail-close。完成 axis vertical slice 后，删除 axis reducer 的远端重推和客户端历史逆 mutation 分支，再按相同接口迁移其他结构操作。

### 六轮自审回合 — 非单元格结构引用的可逆性 fail-close（2026-09-25）

1. **改动时序：** `applyAxis` 与 `applyCellShift` 在改 live `WorkbookModel` 前执行公式重写预检；名称、规则、工作簿公式 owner 的往返检查放在该阶段，拒绝不会留下半改模型。
2. **Owner 覆盖：** 预检与既有写入路径字段对齐：defined-name、CF/DV 的 `value1/value2/formula1/formula2/listSource.formula`，以及 table-sheet、shape drawing、data-view、cell-style-template formula/anchor。轴/单元格删除若变换后不能被反向变换还原，会返回 `UNSUPPORTED_STRUCTURAL_REFERENCE`。
3. **服务端事务边界：** Java `StructuralMutationDescriptor.applyWithPatch` 先深拷贝 snapshot；reducer 在本地候选快照上的 fail-close 错误会中止提交，不写 operation/revision/checkpoint。
4. **比较语义：** AST 客户端比较公式规范化后的结构，Java 端规范化 A1 引用 token；避免把公式空格、大小写或 renderer 格式差异误报成不可逆。Java/TS 都只在前向结果变化时执行逆向检查。
5. **双端对齐：** Java 对 axis、cell shift 和 `range.move` 的公式 cells/rules/names/persisted owners 使用正向映射与反向映射；客户端同样在对应操作的预检阶段检查 cell provenance/barcode、hyperlink address 和上述非单元格 owners。
6. **剩余边界：** `rows.permuted`、fill/paste、sheet identity、table resize、OOXML owner relocation 尚未统一到可逆 `StructuralPatch`；本次不把 fail-close 子集宣称为完整 runtime。删除受影响引用的非单元格公式现会明确拒绝，直到完整 owner delta 与保护/冲突作用域接入同一 patch。

**本轮交付边界：** 这是现有单元格公式 patch 的临时 fail-close 闭环，避免不可逆公式文本被结构 undo 留成 `#REF!`；不是完整修复非单元格 owner 的提交载荷/冲突/保护范围。严格按用户要求只静态审查，未运行测试、构建或 UI 验收。

### 六轮自审复核 — CF/DV 公式 owner 纳入 axis/cell-shift patch（2026-09-25）

1. **模型回合：** `StructuralFormulaOwnerDelta` 仅有 cell 地址与 cell state；真实变更的 CF/DV 规则公式没有 locator 或前后状态。加入 rule-kind、sheet/rule/field 稳定身份、公式 before/after 与适用范围 before/after。
2. **变换回合：** axis 变换先移动规则范围再改写公式，cell shift 也分阶段改写范围、anchor 与公式；仅在改写前捕获公式和范围，不能构造可逆 patch。两条路径现捕获 pre-state 并以最终规则状态产出 delta；重复规则身份在规划阶段 fail-close。
3. **历史回合：** cell-only `applyFormulaOwnerDelta` 不能回放规则公式，逆结构操作还可能已恢复适用范围但未恢复公式。历史/远端应用现定位具体规则字段，要求唯一 owner 和逆操作后的目标范围，再执行幂等写入或拒绝漂移。
4. **协议回合：** protocol validator 与 committed impact 计算硬编码 cell 地址，新增规则 delta 会被拒绝或把范围影响算错。已按 kind 做精确字段校验、拒绝重复 owner key 和越界/异表范围，并将规则的前后适用范围作为去重后的 impact ranges。
5. **服务端回合：** Java reducer 曾改写 CF/DV 公式但仅返回 cell delta，提交 patch 因而无法描述这部分状态；patch 应用和 inverse 也只支持 cell。现捕获规则 pre-state、生成/反转 rule delta，并在服务端 reducer 中以规则身份、公式与 post-transform ranges 做 fail-close 应用。
6. **权限/冲突回合：** `committedRanges`、`structuralImpactRanges` 与 patch merge 以 `afterAddress` 为唯一键，无法授权/隔离 rule owner 或合并同一规则字段。现将前后适用范围纳入保护与冲突范围，并以 rule identity + field 对齐 reducer 与 inverse patch；前端历史影响范围也按完整 range 去重，避免每个公式字段重复放大范围列表。

本次只覆盖 axis 和 cell-shift 上 CF/DV 公式字段；move-range、permutation、fill/paste、sheet identity、table resize、其他非单元格公式 owner 与 OOXML 尚未统一到同一 patch，不能据此宣称目标完成。按要求只做静态审查，未运行测试、构建或 UI 验收。

### 六轮自审 — range.move owner patch 与 undo 闭环（2026-09-25）

1. **CI 类型诊断复核：** PR 远端编译日志证实 command-runtime 导入了不存在的 `StructuralFormulaOwnerIndex`、影响范围变量名与定义不一致，且 `MoveFormulaRule` 缺少使用中的 `id`。已按实际导出名、变量名和稳定 owner 字段修正；未在本地重跑构建。
2. **单元格 owner 复核：** move 会重写源区域公式及外部依赖公式，但 effect 过去只报告重写地址，不生成 before/after formula-cell delta。现分别记录源地址到目标地址的 moved-owner delta，以及同地址 dependent-owner delta；计划已排除 source/destination owner，避免重复 key。
3. **CF/DV owner 复核：** move 的规则公式已有可逆预检与写入，但没有进入历史/保护/冲突 patch。现于任何范围或公式变更前捕获稳定 rule identity、公式和范围，完成变更后生成 formula-rule delta。
4. **Java reducer 复核：** `range.move` 的 descriptor 原先返回 null patch，snapshot reducer 也丢弃单元格与规则公式重写结果。现 reducer 依据变换前后 owner 状态生成 patch，descriptor 将其返回给提交与 journal replay 链。
5. **协议与影响范围复核：** protocol mutation-id allowlist 明确拒绝 `range.move` patch；补入 move 后，cell 与 rule 前后范围会走既有严格字段/边界/重复 owner 校验及去重后的 impact range 计算。
6. **Undo/授权复核：** 服务端 undo 原来既不把 `range.move` 纳入 structural-patch mutation，也不能匹配其反向 source/destination；因此不会取回/校验目标 patch。现验证逆源范围等于前向目标几何、逆 targetOrigin 等于原 source 起点，只有精确逆操作才可复用 inverse patch；patch 应用对 reducer 已恢复的状态保持幂等。

本轮静态改动覆盖 move 的 cell 与 CF/DV formula owner patch、protocol 接受范围、Java patch 派生及服务端 structural undo 匹配。Formula names、table/drawing/template 等其他 persisted formula owners、permutation、fill/paste、table resize 与 OOXML 尚未进入同一 owner-delta/impact 契约；整体整改仍未完成。按用户要求未运行测试、构建或 UI 验收。

### 六轮静态复审 — 删除锚点的拒绝错误优先级（2026-09-25）

1. **CI 结果复核：** 两个 `canonical-build` 独立运行均有相同的 2 个断言失败，均为锚点删除预期 `VALIDATION_ERROR`、实际先收到 `SERVICE_UNAVAILABLE`；不是编译失败。
2. **轴变更用例：** `structuralAxisRewriteCoversPersistedFormulaOwnersAndRejectsRemovedTemplateAnchors` 删除模板 formula anchor 所在行；模板公式逆变换失败先于锚点检查，遮住了 owner 坐标的确定性拒绝。
3. **调用顺序：** `applyAxis` 原先先搬移 cells/metadata，再进入公式 owner 改写，模板 anchor 只在 `rewritePersistedFormulaOwners` 内后验映射；现于结构写入前预检名称和模板 anchors。
4. **单元格位移用例：** `cellInsertAndRowPermutationHaveDeterministicInverseFriendlySnapshots` 删除定义名称 anchor 所在单元格；同样由名称公式不可逆错误抢先，现先用相同 cell-shift coordinate mapper 判定 anchor 删除并返回 `VALIDATION_ERROR`。
5. **边界对齐：** axis preflight 复用 `shiftIndex` 与 `definedNameAnchorCoordinate` 的工作表边界语义；cell-shift preflight 复用 `remapCellShiftCoordinate`，模板与名称 owner 使用相同变换方向/范围。
6. **失败原子性：** 预检在 Java 深拷贝候选快照上运行；静态复核还发现 `SnapshotMutationSupport.array` 会为缺失的可选模板字段创建空数组，故预检改为只读可选字段、仅验证其存在时的形状，避免成功结构编辑凭空改写快照结构。

本轮仅按远端 CI 失败及源码路径做静态修复，未在本地运行测试或构建；需由 PR 后续门禁确认。此前远端失败的两个断言均仍待新 head 验证。

### 六轮静态复审 — rows.permuted 引用与影响范围闭环（2026-09-25）

1. **显式公式锚点：** 对照 TypeScript `remapRuleForPermutation` 与 Java `remapPermutationRuleFormulaOwners`，确认后端偏移 CF/DV 公式时没有把显式 `formulaAnchor.row` 写到置换后的行；现与公式偏移共用同一目标行。
2. **隐式公式锚点：** 对照已有“range fragments reorder”前端契约，确认后端物化隐式锚点时写入旧行；现写入映射行，并保留无行变化时的原坐标语义。
3. **BandedRule：** 前端在排序前校验并重映射 `bandedRule.range`，后端原先既不拒绝碎片化结果，也不写回；现添加预检与精确范围写回。
4. **ReportSheet：** 绑定单元格按行跟随置换，但列可能在当前 materialized grid 之外；旧 canonical `affectedColumnEnd` 未计入绑定列，导致回调因范围不足而原地保留。现前后端都纳入所选行内的绑定列，绑定先映射后应用。
5. **绘图引用：** 排序原先漏掉跨工作表拥有的 camera/screenshot/chart/form-control payload 引用。现对工作簿全部 drawing payload owner 做同一行映射；区间不能精确表示为单一区间时在写入前 fail-close。仅复制受影响 payload，避免对无关的大型对象做深拷贝。
6. **影响范围篡改：** 服务端和客户端原先只要求 `affectedColumnEnd` 覆盖下界，允许提交者多报列并扩大行元数据迁移范围；现服务端要求与独立 canonical extent 精确相等，客户端在应用前也校验 metadata scope 精确匹配。

已补充前端与 Java 回归用例，覆盖显式/隐式规则锚点、远网格 ReportSheet 绑定、banded range、跨工作表 camera source 及 over-reported extent。按要求未执行本地测试、构建或 UI 验收；`git diff --check` 通过，仍需 PR CI 验证。本轮没有统一 rows.permuted 与 axis/cell-shift/move 的服务端公式 owner patch 协议，其他结构入口和 OOXML 仍需后续静态审查；本目标未完成。

### 六轮静态复审 — rows.permuted 远端门禁回归（2026-09-25）

1. **失败真实性：** 核对 CI head 为 `8cc7c628`，失败发生在两项后端回归断言，不是 Actions 弃用告警或类型编译失败；一项暴露置换后 camera source range 被错误拒绝，另一项暴露显式 CF anchor 被映回旧行。
2. **映射方向：** `validatePermutation` 返回 target→source，而区间映射消费 source→target。preflight 误收前者；对 `[2,0,1]` 会把源行 0、1映到 2、0并产生两个区间。现从同一 `targetRowsBySource` 构造点映射后复用于 preflight 与 apply。
3. **区间闭环：** 对照 camera 跨 sheet owner 的静态用例，确认 apply 与 preflight 必须使用完全相同的单区间精确映射；否则可在 prepare 成功后 apply 才失败。已让校验和变换共享同一方向的行映射。
4. **锚点单一所有权：** 跟踪 CF/DV 从 `remapPermutationRuleFormulaOwners` 到 `SheetRuleLifecycle.transformStructuralFields`，确认显式锚点被连续映射两次；反转排序会偶然掩盖该缺陷。现在锚点与公式偏移在规则置换路径中一次性处理，生命周期阶段仅处理 DV range list source。
5. **跨端规则语义：** 对照前端 `remapRuleForPermutation` 与 Java 规则 owner 逻辑，确认两端都会为字面值规则无条件生成 `formulaAnchor`；现按各自 CF/DV 公式字段选择器判定真实公式 owner，只有公式规则才物化隐式锚点。
6. **失败原子性与覆盖：** 再核对 reducer 先在 descriptor 的 snapshot 深拷贝上完成 exact-range preflight，失败不写入持久状态；前后端回归用例覆盖非对称置换、显式/隐式公式 owner、camera range 和字面规则。修复后 CI 暴露新增后端用例遗漏规范 `review` 对象；已补齐空 review 结构，这是测试夹具完整性问题，不计为生产缺陷。未运行本地测试/构建；由后续 PR CI 负责验证。

该轮只修复上述静态追踪和 CI 明确证实的问题；没有把未验证猜测计入问题数。若远端门禁继续暴露实际失败，将在同一 PR 上继续修复；本目标仍未完成。

### 六轮自审复核 — 结构统一链未闭环（2026-09-25，HEAD `8155b081`）

以下 30 项是源码可定位的架构缺口，不等同于 30 个独立崩溃缺陷；每轮复核后均对照调用方、mutation 契约及后端 reducer，确认不是旧文档推测。未运行本地测试、构建或 UI。

**第 1 轮：操作入口与语义所有权**

1. 整行/列插删进入 `StructuralTransform.applyAxis`，而不是一个全操作 planner。
2. 单元格插删另走 `planCellShift` / `applyCellShift`，范围位移契约与轴变换分离。
3. `range.move` 进入 `applyMoveRange`，粘贴/剪切则由 `applyPasteSnapshot` 重建状态；跨表剪切明确拒绝。
4. 排序使用独立 `rows.permuted` / `applyRowPermutation`，前后端另有排序 reducer。
5. Fill/drag 使用公式 offset 后发普通 cell/range mutation；Sheet identity 使用 `SheetIdentityTransform`，表格仅有 add/update/remove 路径而没有专用 resize mutation。

**第 2 轮：引用 Owner 与索引边界**

6. `StructuralReferenceOwnerIndex` 原先只提供 cell-address 查询；本轮已把 defined name 身份、引用几何与 anchor 查询纳入同一个 `ReferenceIndex`，规则、表、绘图和范围对象仍未统一。
7. `preflightFormulaRewrite` 原先扫描全部 `definedNameModels`；轴插删、单元格位移与 move 现在只查询受影响名称，并通过 typed before/after delta 增量同步 FormulaEngine。Sheet identity/full rebuild 生命周期仍有全量名称同步。
8. CF/DV 公式和范围以 `workbook.getSheets()` 全表循环重写、快照，不是按目标 sheet/range 命中 owner。
9. table-sheet 列、shape payload、data-view 字段、cell-style-template 公式在 `preflightWorkbookFormulaOwners` 全量枚举。
10. Hyperlink、chart/pivot/sparkline、filter、print、spill、protection 等坐标 owner 由多个 `shift*` / `relocate*` 过程逐类扫描，缺少统一可查询 owner 索引。

**第 3 轮：预检成本与提交原子性**

11. `preflightAxisMetadata` 对每张 worksheet 深拷贝结构元数据，即使其与编辑范围无交集。
12. `preflightCellShiftMetadata` 也复制所有 worksheet、workbook tables 和 data sources。
13. 轴/单元格位移对公式规则先对全 workbook capture snapshot，再独立遍历规则执行实际变换。
14. `StructuralTransform.apply*` 的验证计划仍以 live `WorkbookModel` 为输入并在验证后原位修改；没有可提交/可丢弃的只读 immutable plan。
15. `StructuralTransformResult` 主要描述计算输入范围、删除的 cells 和公式 owner delta；并非完整 cell/metadata/object before-after patch。`CommandRuntime.applyMutation` 在 mutation callback 返回后才把该 mutation 的 inverse 放入事务，因此 callback 中途失败时事务没有该 mutation 的可重放逆项。

**第 4 轮：Patch 协议及跨端同义性**

16. TypeScript `StructuralPatch` / Java `StructuralPatch` 只有 `formulaOwnerDeltas`，不承载 cell 和普通结构元数据 before/after。
17. Defined-name before/after 已加入客户端结构 effect，但尚未进入版本化跨端 `StructuralPatch`；persisted formula participants（table-sheet、drawing payload、data view、template validation）仍缺少 effect/patch delta。
18. Java `StructuralMutationDescriptor` 的 `rows.permuted` 分支调用会改写规则、名称、模板、绘图的 reducer 后，仍返回 `structuralPatch == null`；客户端排序 effect 也只给计算清除/重载范围。
19. 协议 structural-patch allowlist 不含 `rows.permuted`、table resize 或 sheet identity mutation；这些操作因此不能由同一 patch 契约确认 owner 结果。
20. TypeScript 与 Java 分别实现轴、cell shift、move、sort 与 metadata mapping；共享的点/区间向量只验证原子坐标映射，不校验完整 workbook owner patch 等价。

**第 5 轮：History、OT 与远端重放**

21. `range.move` history policy 是 `invalidate`，不能将已有本地 undo/redo 坐标变换到新结构。
22. cell shift 同样因没有 canonical history transform 而 invalidate。
23. `rows.permuted` 也 invalidate history；其 inverse 仍是另一组 permutation mutation，不是同一结构 patch 的逆。
24. `ot-rebase.ts` 将 move/sort/table-resize/sheet-identity 归为需 canonical patch 的结构类型；缺少 patch 时 fail-close，而不是由共享 ReferenceTransformDomain 变换。
25. 远端 replay 仍调用本地 mutation handler 重算结构，再把服务端公式 delta 应用到结果；服务端并未发送可直接提交的完整结构 patch。

**第 6 轮：计算、投影、持久化与 OOXML 边界**

26. 计算 context rebuild 由 mutation-id 集合和 effect 上的布尔字段共同决定，重建策略没有与结构变更的 typed owner effect 同源。
27. 结构 effect 是 command listener 的临时返回值；runtime 持久化/提交 mutation 时只复制 mutation 参数、affected ranges 和现有 impact ranges。
28. ProjectionRuntime 通过 mutation 分类与 affected ranges 推导失效，不消费完整结构 owner patch；chart 依赖也另有索引与失效过程。
29. 未建模/opaque OOXML capability 的主要拒绝发生在 `exportOoxmlDocument`；结构 planner 不先证明 opaque part 的坐标 owner 可安全映射，因此编辑可能先成功、到保存才拒绝重建。
30. Java `applyPublicMutations` 对 mutation 顺序调用 descriptor；结构 descriptor 的 `applyWithPatch` 对每次调用深拷贝完整 JSON snapshot，批次成本随 mutation 数重复承担 snapshot 分配。

**修复收敛方案：** 以上问题收敛到同一条 clean-break 路径：先建立 side-effect-free `CanonicalStructuralPlanner` 和 typed `ReferenceIndex`，让其对已注册 owner 产生版本化、不可变且可逆的完整 `StructuralPatch`；随后本地命令、计算与投影、history/OT、Java commit/replay、持久化和 OOXML capability preflight 全部消费同一 patch。迁移先覆盖整行/列轴变更的完整纵向链，并在同一提交中删除被替代的直接 live-transform 分支；再按同一契约迁移 cell shift、move/paste、permutation、identity、table resize。预检工作量以受影响 cells/owners 为界，服务端一批 mutation 只建立一个事务工作快照；不能登记或安全映射的 owner 在提交任何 live 状态前以 typed error 拒绝。

本轮已完成六次独立静态复核并记录 30 项经源码确认的架构缺口；它们尚未被本节所替代的统一实现修复，故不声称整改完成。下一实施批次必须一次闭合首个轴变更纵向链（成功/拒绝、逆 patch、远端重放、server 和 OOXML capability），本地仍按用户要求不运行测试或构建；修复提交继续进入现有草稿 PR #345。

### 六轮静态自审 — rows.permuted 公式 owner patch 接线（2026-09-25）

本轮按调用链分六轮复核，每轮记录可由源码直接确认的断点，不把同一断点的重复表现扩充成 30 个独立缺陷：

1. `applyRowPermutation` 在重写移动单元格公式后返回 `void`，调用方无法记录 before/after cell formula owner；已改为产出可逆 owner delta。
2. 同一 permutation 会重写 CF/DV 公式及其 ranges/anchor，但没有把规则公式变化放入 patch；已按现有 structural rule 字段选择器生成规则 delta，并对不稳定规则身份 fail-close。
3. 已注册的 `rows.permuted` mutation handler 丢弃 cell/rule delta，计算 effect 因而不能进入 runtime history；已将 delta 接入 effect。
4. `data.sort.rows` 是另一条独立入口，同样丢弃 delta；已同步接入，避免仅修 replay handler 而漏掉交互排序。
5. Java `permuteRows` 改写公式后原先不返回 patch，协议 allowlist 和 undo patch 分类也不承认该 mutation；已补服务端 patch、协议许可及结构撤销分类。
6. 撤销端此前无法证明 `sourceRows` 是原 permutation 的逆映射；已加入精确范围/列界/逆序映射核验，并让 undo 使用目标 patch 的 inverse。

静态追踪确认 TS/Java 对移动 cell formula、source formula、barcode formula、CF/DV formula 字段的 delta 类型及范围形状一致；history 按 owner delta 内容排序比较，impact ranges、server replay 幂等应用和 undo merge 均沿现有 `StructuralPatch` 链处理。补充了 TypeScript 与 Java 回归用例源码。仅执行 `git diff --check`，未运行任何测试、构建或 UI；仍需 PR CI 验证。

本批次只闭合公式 owner delta 的排序接线，不等于实现完整 `CanonicalStructuralPlanner`/`ReferenceIndex`：名称、模板、绘图、普通元数据仍不在完整可逆 patch 中；history rebase 仍会 invalidate 受后续 permutation 影响的历史项；更早已提交且没有 server patch 的旧排序操作也不能据此获得可逆 patch。原六轮列出的其余架构缺口仍开放，PR #345 与总目标均未完成。

### 六轮静态自审 — 计算上下文影响的类型化传播（2026-09-25）

本批次按六条独立调用链复核，每轮只记录能由当前源码与真实调用方确认的问题：

1. **影响契约：** `runtime.ts` 用 mutation-id 集合与 `StructuralTransformResult.requiresCalculationContextRebuild` 布尔值两套来源分派重建；结构结果没有可区分“重建/同步名称/同步表”的动作。新增 `WorkbookCalculationContextEffect`，并将结构结果的公开布尔字段替换为 typed effect。
2. **工作表身份入口：** `sheet.add/remove/restore/duplicated/reordered` 注册契约没有声明重建影响，执行、撤销与协作 replay 只能依赖运行时 ID 清单。上述五种 mutation 现显式声明 `rebuild`。
3. **公式上下文数据入口：** `table.add/remove`、`sheetTable.add/remove/update` 与 `name.set/remove` 的表/名称同步规则只在运行时 ID 分支中存在；现分别声明 `sync-tables` 或 `sync-defined-names`。
4. **本地 effect 传递：** `addSheet`、`removeSheet`、`duplicateSheet`、`removeTable` 的命令回调会返回模型对象；直接采用 `apply()` 返回值时，metadata effect 的 undefined 兜底被这些非 effect 值遮蔽。相关回调现明确返回 `void`，保留结构 effect 的优先级。
5. **历史与协作传递：** `applyHistory` 只转发 replay handler 返回值，不读取 mutation metadata；所以 undo/redo/remote 对没有显式 handler effect 的身份、名称及表操作会漏掉上下文更新。replay 现与本地 apply 使用同一 metadata 兜底；新增了 command/undo/redo 传播及非法 metadata 拒绝用例源码。
6. **重算所有权：** 即使同步动作改为 typed effect，`FORMULA_SYNC_MUTATIONS` 仍重复列出相同 mutation ID，导致 effect 契约与重算触发清单继续分叉。计算上下文 effect 现直接进入调度分支，这些重复 ID 已从该集合移除；简单 sheet rename 等仍需独立增量同步的入口保留原分类。

本批次仅静态检查并执行 `git diff --check`；按用户要求未运行测试、构建或 UI。修复覆盖旧 30 项清单中的计算上下文路由缺口及上述直接传播缺陷，不代表其余结构 patch、owner index、Java/OOXML 纵向链已完成；PR #345 和总目标仍未完成。

### CI fixture follow-up — row-permutation worksheet identity

head `4a4ba321` 的两条 Java CI 都在 row-permutation 用例中以 `name is required` 失败。下载并检查既有 CI Surefire 报告后，调用栈定位到 `captureRuleFormulaSnapshots` 对每张 sheet 读取 canonical `(id, name)`；`WorkbookSnapshotValidator` 也要求 worksheet name 非空。失败的两个测试快照漏了 `name`，并非生产运行时应接受缺失身份的情况。已补齐两个 fixture 的 worksheet name，保留 fail-close 行为；未在本地运行测试，需由新 PR head 的 CI 确认。

### 六轮静态复核 — worksheet rename drawing payload 联合类型

head `b0217386` 的两个远端 `canonical-build` 都报告同一组 TypeScript 错误：`sheet-identity-transform.ts` 的 shape/chart drawing 变更集合推断为 shape 专属类型，导致 chart 分支不兼容，并使后续 delta flatten/apply 中的 `change` 降为 `unknown`。六轮复核逐一确认：

1. 两个 CI job 是同一个编译根因，不是两组独立问题。
2. `DrawingPayload` 是 shape 与 chart 等绘图 payload 的判别联合；变更集合确实同时产生不同分支。
3. shape 分支返回 `ShapeDrawingPayload`，chart 分支返回 `ChartDrawingPayload`，原先无上下文类型的嵌套 `flatMap` 被首个分支过度收窄。
4. 非目标 payload 与无变化公式返回空数组；显式标注回调结果为 `DrawingPayloadChange[]` 保持 flatMap 的零到多契约。
5. 收集公式 owner delta 与 apply 写回都依赖变更记录的共同字段；将 payload 字段定为完整 `DrawingPayload` 联合类型可为两处提供同一正确契约。
6. 类型修正只约束既有预计算结果，不改变 clone、公式映射、delta 生成或 live apply 顺序；已有 rename linked-chart-formula 回归用例覆盖该业务路径。

已修复为显式联合类型记录及 flatMap 数组结果类型。此次 CI 的失败已确认是编译期真实错误；没有本地重跑测试或构建，修复后的远端 CI 状态待新 head 确认。

### 六轮静态复核 — linked chart title projection 与公式失效

源码追踪确认图表标题 `linkedFormula` 不能只作为 OOXML 字符串保存：

1. **渲染消费：** Canvas 使用 `buildChartLayout`，而原 `baseLayout` 只读取 `elements.title`，不解析 `titleText.linkedFormula`；新增从公式单元格 projection 派生标题文本，不写回持久模型。
2. **公式域：** Excel 的图表标题链接是单个字符串单元格引用；新增共享解析器只接受可解析的单 cell AST，其他表达式以 `UNSUPPORTED_FEATURE` fail-close。
3. **依赖范围：** `chartSourceRanges` 原来只含数据系列范围；标题、legend、axis、data-table 文字公式的源 cell 现加入同一依赖范围计算，命令 affected ranges 与 active-sheet projection 因此包含跨表来源。
4. **普通输入失效：** projection 图表范围索引由上述共享 range 函数重建，来源 cell mutation 会使图表 owner sheet projection 失效。
5. **公式结果失效：** `onCalculationApplied` 原先将精确 changed addresses 降为 sheet-ID 集合，跨表公式依赖的图表 owner 不会失效；现在传递 cell 地址并按单 cell range 查图表依赖。
6. **不支持文件边界：** 不能投影为单 cell 的 native title reference 不进入可编辑 canonical chart，而保留原始 chart part 并记录原因；其他尚无 Canvas renderer 的 chart-text formula 在画布上显示明确 unsupported 状态。

已加入 snapshot 单 cell 引用拒绝路径及图表 source-range/title-value 用例源码。`git diff --check` 是本轮唯一执行的本地验证；未运行测试、构建或浏览器验收。此批仍复用 projection 中现有 chart range 索引；将其与 FormulaEngine `ReferenceIndex` 合一、以及跨层 `CanonicalStructuralPlanner/StructuralPatch` 仍未完成。

### 六轮自审复核 — linked chart formula 的结构失效与 PR 阻断（2026-09-25）

1. **CI 失败根因复核：** 两个最新 Java CI 都在同一测试源码的 `JsonNode.putObject` 编译失败，逐处核实两个父节点实际是 `ObjectNode`；改用父节点 `putObject` 创建对象，未弱化 JSON 类型或校验。
2. **实际 mutation identity：** 顺着 `sheet.rows.insert` command 到 `rows.inserted` mutation 确认，ProjectionRuntime 收到的是 mutation ID，而非 UI command ID；原 chart-source-index 重建条件只认 `sheet.rows.*`，因此真实行/列/cell-shift/sort/move 变更不会重建索引。
3. **触发后果：** 结构变换会改写 chart linkedFormula 或普通系列 RangeRef；失效流程只按 `mutation.affectedRanges` 求交会漏掉随结构移动的引用。仅重建索引仍不够：如源引用 `A10` 随第 5 行插入移到 `A11`，插入带与新地址不相交，跨表 owner 不会被命中。现对结构变更源 sheet 上新索引中的图表 owner 一并失效；并把结构 effect 的 `formula-object` owner delta 仅投递到 projection 队列，精确失效图表所在 sheet，不混入客户端 operation。
4. **恢复路径：** 核对 `rows.deleted`/`columns.deleted` 的 inverse 会先执行对应插入 mutation，cell-shift 的 restore 使用独立 mutation ID；这些正向、逆向结构 mutation 都加入索引失效集合，避免 undo/replay 后再次遗留陈旧坐标。
5. **sheet identity：** 源码实际重排 mutation 为 `sheet.reordered`，不是旧判断中的 `sheet.move`；按注册 ID 覆盖 add/remove/restore/rename/reordered/duplicated，索引 owner 集合及 sheet identity 变化都会重建。
6. **投影域与旁路：** 结构 mutation 现归入完整 projection domain；图表自身、table 和 sheet-table mutation 仍分别触发既有重建入口。公式重算仍使用 exact changed cell addresses；owner delta 只沿内存 projection queue 传递，不改变 history/server operation envelope；本次不把这一 cache 索引包装成已统一的 FormulaEngine `ReferenceIndex`。

本轮仅读取远端既有 CI 日志、源码与 mutation 注册链，并运行 `git diff --check`；没有在本地执行测试、构建、lint 或浏览器验证。Java fixture 编译修复和 projection index mutation-ID 修复均进入现有 #345 草稿 PR；结构 planner 与统一 ReferenceIndex 的总体整改仍未完成。

**结构删除边界复核：** `mapAstStructuralReferences` 会将被删除的单元格引用明确序列化为 `=#REF!`，而不是继续保留一个可索引地址。因此 chart-text range helper 对该 AST 返回“无 cell dependency”，快照继续保留合法错误公式，Canvas 使用公式错误标记而不回退到陈旧缓存标题；其他非单 cell AST 仍拒绝。补充了 snapshot 与标题解析的回归用例源码，未在本地执行。

### CI 静态跟进 — chart formula projection 类型

新 head `34abb455` 的两个 CI 在前端类型检查报告同一组两个确定错误：`RangeRef` 的所有者是 core-model `index.ts`，并非声明它为本地导入但未导出的 `domain.ts`；此外 `isStructuralTransformResult` 原调用结果是 boolean，不能再当结构 effect 读取 `formulaOwnerDeltas`。现改为从类型所有者导入 `RangeRef`，并保留经 type guard 收窄的 effect 值供既有同步逻辑和 projection delta 使用。修复只依据远端失败日志与声明类型，不做本地类型检查或测试；新 PR head 的自动 CI 待确认。

### CI 静态跟进 — Java drawing fixture 对象覆盖

head `1d8508ea` 的两个后端 CI 已通过 test-compile，随后在 `sheetRenameRewritesAllPersistedFormulaOwnerCategories` 的既有 shape 公式断言失败。对照 fixture 逐行确认：为通过 `JsonNode` 的静态类型编译而将第二次 `source.putObject("drawingPayloads")` 改成 `ObjectNode.putObject`，会替换而非复用原子对象，丢掉 `formula-shape`；external drawing map 也有相同覆盖风险。现为每张 sheet 创建一次具名 `ObjectNode` 并向其追加 shape 与 chart，保持 fixture 的兄弟 owner。未本地运行测试；新 head CI 待确认。

### 六轮自审复核 — 完整结构引用链（2026-09-25）

以下六轮各自沿一条不同的源码链复核目标契约；相同 CI job 或同一根因没有重复计数。复核时撤回了两个过宽结论：正常 `hydrateRuntime` 已注入增量维护的 `FormulaEngine.dependencies`，因此 `buildStructuralReferenceIndex` 的无 provider fallback 不能描述为每次生产结构编辑都会扫描全部 cells；`ProjectionRuntime.chartSourceIndex` 管理的是图表数据源投影依赖，与公式引用索引职责不同，分开维护本身不构成重复实现。六项均有源码证据；图表索引每次结构编辑全量扫描的问题已在 follow-up 局部修复，其余结构整改仍未闭合。

1. **Planner / ReferenceIndex 成本：** `runtime.ts` 的生产 hydrate 与远端重建路径为 `CommandRuntime` 注入增量维护的 `runtime.formula.dependencies`；`buildStructuralReferenceIndex` 全量构建仅是未配置 provider 时的 fallback，不能算作每次正常编辑成本。当前可确认的全量扫描是 `core-model/src/structural-transform.ts` 的 `preflightFormulaRewrite` 遍历全部 `definedNameModels`，且 `ReferenceIndex` owner 仍以 `CellAddress` 为核心，未为 name 等非 cell owner 提供稳定类型化身份与范围查询。方案：扩展 canonical `ReferenceIndex` 的 typed owner identity 与增量注册/移除契约；名称按引用目标范围查询，保留 workbook-scoped name 的解析上下文语义，不用伪造 cell address。
2. **前端事务原子性：** `applyAxis` 先移动 cells 并逐类原位改写元数据，随后 `applyFormulaRewritePlan` 仍可因 missing cell、owner 类型变化等 `STRUCTURAL_PATCH_INVARIANT` 抛错；`CommandRuntime.applyMutation` 只在 `mutation.apply(context)` 成功返回后才登记 inverse，外层 rollback 因而不能撤销这个半完成 mutation。方案：所有 owner 变更先进入 side-effect-free plan，校验完整 before/after 与 inverse 后再原子提交；提交前失败不触碰 live model。
3. **协同结构变换未闭合：** protocol envelope 能带 server `structuralPatch`，但 `operation-types.ts` 的 `ClassifiedMutation` 不含 patch，`committedMutationToClassified` 只保留 id/params/ranges；`ot-rebase.ts` 对 `move-range`、`sort`、`table-resize`、`sheet-identity` 明确拒绝 rebase。拒绝本身是 fail-close，现有 cut-paste 用例也明确验证该安全边界；但用户目标明确要求并发结构编辑、远端 replay 与统一 structural patch，且该限制使有效操作不能进入支持链。方案：先把版本化地址变换事实纳入 canonical patch，再让 classification、queued rebase、history 与 remote replay 消费同一事实；在 patch 尚不支持某 owner 时，保留局部、可恢复的冲突，不得拒绝无关远端事务。
4. **图表投影索引全工作簿重建（已局部修复）：** 原确认的问题是每次结构 mutation 都扫描 `O(all sheets + all drawings)`；projection follow-up 已改为按来源范围、table/pivot ID 命中 owner，并只替换受影响图表的 postings。首次建索引及 sheet identity/order 生命周期仍做全量重建；该索引仍独立于 FormulaEngine `ReferenceIndex`，不能描述为两者已统一。
5. **Java 结构语义重复且批量提交重复 snapshot clone：** `StructuralMutationDescriptor` 与 `StructuralSnapshotReducer` 在 Java 中再次执行结构与引用变换，与前端 `StructuralTransform` 形成两个语义 owner，违反“Java 只执行/校验 canonical result”的约束。另有可量化成本：`WorkbookOperationService.commit` 逐项 reducer，结构 descriptor 每次 `snapshot.deepCopy()`，N 个结构 mutation 会重复复制全 workbook N 次。方案：服务端只验证并持久化 canonical structural patch；在迁移完成前，至少以事务级私有 snapshot 承载整批 reducer 变更，批次通过后一次发布。
6. **OOXML capability 拒绝过晚：** `StructuralTransform.apply` 的输入只有 `WorkbookModel` 与 formula owner index，不包含 native package capability；`exchange-excel-ooxml/src/export.ts` 则在重建导出时才检测 unknown worksheet/workbook nodes、extensions 与没有 canonical chart owner 的 parts 并拒绝。现有 OOXML 用例确认：导入后编辑普通 cell 会成功，但保存阶段抛 `NATIVE_DOCUMENT_UNCHANGED_SAVE_REQUIRED`。这是避免丢失未知内容的安全拒绝，却没有在结构事务前揭示本轮不支持的修改边界。方案：将 native-part owner/capability preflight 纳入 canonical planner，不能安全映射的相关编辑在 live commit 前以对象定位和 typed error 拒绝；未知内容仍原样保留，不能为通过编辑而删除。

**收敛实施顺序：** 先定义可逆、版本化的 canonical `StructuralPatch` 与 `ReferenceIndex` owner contract；以一个事务快照生成纯 planner 结果并先闭合轴编辑纵向链（本地提交/撤销、OT/replay、Java 提交/回放、计算/投影、OOXML capability preflight）；再迁移 cell shift、move/paste、permutation、sheet identity、table resize。每个阶段补成功与拒绝行为测试源码，但本任务按用户要求不运行本地测试、构建、lint 或浏览器；用 PR CI 验证。图表索引全量扫描已局部收敛，但整体 canonical planner/index、其余跨层迁移及验收仍未完成，不能把当前 PR 标为完成或合并。

### Projection follow-up — deleted table chart owner invalidation

对 workbook table 删除的后置监听顺序再做静态跟踪，确认了跨表图表缓存的真实失效遗漏：chart source index 在删除后按新模型重建，已不存在的 table 不再产生 worksheet range binding，随后旧受影响范围无法找到仍引用该 table id 的 chart owner。现将 range/table/pivot chart source owner 记录在同一 namespaced chart source index；table binding 独立于当前 table range 是否存在，table.add/remove 以稳定 table id 精确失效 owner，并由 Set 合并同一 mutation 的范围和 table 命中。补充跨 sheet table removal projection-cache 回归测试源码；未本地运行测试或构建，需等待 PR CI。该修复收敛的是 ProjectionRuntime 内部 chart source 索引，不代表其已与 FormulaEngine 的 canonical `ReferenceIndex` 合并。

### Incremental chart-source index — six static self-audit rounds (2026-09-25)

本轮继续对索引生命周期、mutation 载荷、owner 身份、表/透视表引用、批量改名和失败原子性做了六轮独立静态复核，确认并修复以下问题：图形 mutation 的占位 `affectedRanges` 曾误命中 A1 图表来源；非图表 `drawing.payload.update` 曾被错误登记为 chart owner；pivot 输出刷新曾无谓地重算稳定 source postings；数组 posting 删除在高扇出时会重复线性查找；同一 mutation batch 中多个 sheet rename 曾只刷新最后一个名称；批量 owner 删除可能在后续 invariant 错误时留下部分索引；改名表和 owner postings 曾在解析新绑定成功前更新，失败后可能形成不一致缓存。

现在图表 postings 使用按引用身份可直接删除的集合；非 cell 图形/审阅 mutation 不再参与 chart source range 求交，chart add/remove/update 仍按 drawing/payload identity 更新；pivot refresh 只使相关 owner projection 失效；owner 重索引从候选解析开始即标记 dirty，批量 rename 先构造候选 sheet order，所有 replacement bindings 与旧 posting 预校验后才提交，只有完整成功才清除 dirty，任意失败后的下一次查询都从 canonical model 重建。静态检查确认首次初始化及 sheet identity/order 生命周期仍全量建索引，日常结构编辑只读取受影响 source-sheet postings 并重索引命中的 chart owners。

补充确认：`drawing.payload.update` 也承载 form-control 等非图表载荷；只有 `before` 或 `after` 的 `kind` 为 `chart` 时才将其加入 chart owner 重索引。

### Defined-name reference owners — six static self-audit rounds (2026-09-25)

六轮复核逐轮沿真实调用链确认问题并修复：

1. 从 axis/move preflight 及 row-permutation 锚点计划追到全表循环：所有名称都会被无差别解析/枚举，改为按目标几何查询。
2. 沿 FormulaEngine 与 CommandRuntime 两条索引构建路径核对 owner 身份：名称引用从未进入 ReferenceIndex，补入 typed name postings。
3. 对比公式引用与 anchor 坐标的独立变化：无公式引用的 anchor 也必须参与位移，增加单独 anchor 索引。
4. 检查 workbook 名称解析上下文：无 anchor 的未限定 workbook 引用无法确定 sheet，登记失败并在结构操作前 fail-close。
5. 反查模型所有写入口和精确查找：公开列表允许 push/splice 且查找是线性扫描，改为 identity map 与不可变投影。
6. 从结构 apply 追到 FormulaEngine 提交同步：全量名称归一化抵消了预检索引收益，轴、cell-shift、move、row permutation 改用精确 before/after 增量。

现在 `ReferenceIndex` 用 typed defined-name owner posting 同时支持目标范围、结构轴和 anchor 位置查询；`WorkbookModel` 使用受控 identity map 与不可变列表投影；轴、cell-shift、move、row permutation 的 effect 包含名称 before/after，FormulaEngine 以批量原子 delta 更新名称及 posting，不再在这些编辑后扫描整张名称表。无 anchor 且包含未限定引用的 workbook 名称登记为 `unresolved-context`，结构操作 fail-close；无法解析的名称引用也明确拒绝。CommandRuntime 的冷路径索引及结构测试 fixture 同样登记名称 owners。

#### Row-permutation follow-up — seven static review rounds (2026-09-25)

1. **范围规划全量扫描**：基线 `rowPermutationAffectedColumnEnd` 遍历 `workbook.definedNameModels`，名称总量增长会线性拖慢排序，即使绝大多数 anchor 不在目标行。改为用 `getDefinedNamesAnchoredInRange` 查询所选行、合法列范围内的 owner。
2. **owner 预检再次全量扫描**：基线 `validatePermutationMetadata` 又对完整名称表执行 `flatMap`，使仅修复范围规划仍保留第二个 O(全部名称) 热点。改为按目标行和完整合法列范围查询 anchor owner，再核对 metadata extent。
3. **命令链未传 canonical index**：基线本地排序规划及 `rows.permuted` 回放都调用不带索引的 API，无法沿用 `CommandContext.structuralReferenceOwners`。两条生产路径现显式传入同一 typed index。
4. **模型已改写但结果丢失名称 delta**：基线 `applyRowPermutation` 只返回公式 owner delta，尽管之后通过 `setDefinedName` 更新名称；runtime 因而无法增量维护名称引用 postings。现返回名称 identity 和精确 before/after 状态。
5. **本地与回放没有传播名称变更**：两条执行入口都只把公式 delta 交给计算 effect，导致 owner 索引及依赖公式只能走完整名称快照同步。现在共享 `RowPermutationResult` 并把名称 delta 纳入 effect。
6. **空 delta 与缺失 delta 混淆**：`synchronizeStructuralMutation` 明确把 `undefined` 解释为调用 `setDefinedNameModels(workbook.definedNameModels)`；旧 effect 对“本次没有名称变化”也省略字段，仍会全量同步。新 effect 始终带数组（包括 `[]`），表示名称 owner 已由增量路径完整处理。
7. **预检重复查询同一命中集**：修复全量扫描后，`validatePermutationMetadata` 仍先通过 extent 计算查询并精确查找 anchor owners，随后又查一次相同范围来生成变更。现预检只解析一次 selected-row/full-column 命中集，并复用于 extent 与 name remap；回归测试源码断言规划查询加预检查询共两次，而预检本身只查询一次（未运行）。

复核索引几何覆盖完整 metadata scope；每个查询命中通过 WorkbookModel identity map 精确解析，缺失 owner 或位置不符在首次写入前 fail-close。名称投影冻结后，写回只走 `setDefinedName`。名称变化产生的公式 roots 继续进入 calculation-result/projection invalidation。成功 delta、metadata extent、单次预检查询及 stale-index 拒绝回归断言已加入测试源码；拒绝路径断言 workbook snapshot 不变，测试未运行。

该路径覆盖 CommandRuntime 的本地排序 mutation 与数据排序 command。测试源码未运行；协同 StructuralPatch 仍没有传输 defined-name deltas，属于跨端 patch 未收敛的开放项。

六轮后的反向审查还捕获并修正五处实现缺陷：delta 归一化曾丢掉后续校验必需的 owner identity；行可见性递归仍引用已移除的局部变量；运行时无效 scope 曾可被 identity 计算误当作 workbook scope；row-permutation 仍直接修改现在冻结的名称投影；快照 DTO 的可写数组类型曾接收到只读投影类型。分别恢复身份字段、规范化名称 token、拒绝未知 scope、改走 `WorkbookModel.setDefinedName`，并在 snapshot 边界复制成独立数组；同时增加 delta 拒绝后索引保持不变的回归测试源码。

新增成功路径与拒绝路径测试源码，但按任务约束未运行测试、构建、lint 或浏览器。静态检查仅确认 axis/cell-shift/move/row-permutation 名称选择无全量名称循环、row-permutation 单次预检复用 anchor 命中、公开名称列表无直接写入且 `git diff --check` 无 whitespace error。Sheet identity/full calculation rebuild 仍通过完整快照同步名称；跨端版本化 StructuralPatch、其他 owner families、Java 共用结构语义及整体验收仍是开放项，本轮不代表整体目标完成。

本轮仅执行静态源码审查与 `git diff --check`；未运行测试、构建、lint 或浏览器验收。上述变更尚需 PR CI；不能据此宣称完整 Structural Runtime 整改已经完成。

### Java transaction-owned snapshot replay — six static review passes (2026-09-25)

六轮静态复核在 Java 结构 mutation / committed replay 的快照边界确认并处理以下问题：

1. **独立结构 API 每次复制整本 workbook**：`StructuralMutationDescriptor.applyWithPatch` 必须隔离调用方输入，但 registry 私有批次也只能走这个纯 API。新增包内 `OwnedSnapshotMutationDescriptor` 能力；独立 API 仍复制，owned 入口只处理调用方已隔离的 root。
2. **公开 mutation batch 连续结构编辑重复复制**：`applyPublicMutations` 原先每次结构 mutation 都从上一步重新 deep-copy。现先沿用纯 reducer 返回的独立结果；结构 reducer 只在尚未拥有结果时复制一次，随后连续结构编辑复用 transaction-owned root。
3. **committed replay 在批次开始和每个结构 reducer 各复制一次**：replay 入口改为延迟取得所有权；首个纯 reducer 的输出或首次结构前的 detached copy 成为批次 root，连续结构 mutation 不再逐项复制。
4. **replay 的公式 owner patch 为每个 mutation 再复制整本快照**：`applyFormulaOwnerPatch` 的公开纯入口保留复制语义；registry replay 改调用独立的 owned-root reducer，与本次结构 mutation 共用同一 root。
5. **直接去除 `deepCopy` 会破坏输入纯度和拒绝原子性**：现在纯入口与 owned 入口分离；任何 patch 不匹配或 owner precondition 拒绝都只会丢弃 registry 私有 snapshot，调用方对象不被部分写入。服务端 commit 暂未接入 owned 入口，因为 `committedRanges` 和 data-block 检查需要结构编辑前后的两个状态；需要先设计显式 before/after transaction owner，不能通过别名当前快照来规避复制。
6. **replay 成功/拒绝的隔离契约缺少回归覆盖**：增加源码测试，覆盖 standalone purity、连续结构编辑、公式 owner replay 成功、损坏 patch fail-close，以及成功和拒绝后原输入保持不变。按静态审查要求，测试未运行。

本轮没有更改 mutation 协议或 workbook 数据格式；owned 接口为 package-private，普通 descriptor 继续遵守 `MutationDescriptor.apply` 的独立结果契约。静态检查仅运行 `git diff --check` 和源码核对；未运行测试、构建、lint、浏览器或 OOXML 验收。WorkbookOperationService commit 的全快照复制仍是已知边界，不声称本 slice 解决了完整结构运行时目标。

### Defined-name deltas in local history — six static review passes (2026-09-25)

本轮沿 command → inverse history → undo/redo → remote rebase 做六轮静态复核并修复本地历史链：

1. `StructuralTransformResult` 已生成 `StructuralDefinedNameOwnerDelta`，但 core-model 公共出口未导出该类型；现向 command runtime 暴露同一 canonical delta 类型。
2. `CommandRuntime.applyMutation` 只读取 `formulaOwnerDeltas`，丢弃 effect 中的名称 before/after，故 `MutationInfo` 与 undo plan 没有名称 owner 身份；现写入独立 typed 字段，不伪装成 cell range。
3. inverse plan 只保存公式 owner delta；现把名称 owner delta 放在对应逆 mutation 上，并在 undo/redo 复用它。
4. `applyHistory` 原先只对 formula cell/rule/chart owner 作 precondition 检查；现按 scope/name/sheetId 定位名称，比较公式与 anchor 的预期状态，保留 hidden/comment 等非引用元数据，只改引用状态；陈旧 owner 在 preview preflight 阶段 fail-close。
5. `historyRebase` 的 axis 防护只检查公式 patch；现也检查名称 patch。远端同身份 `name.set/remove` 会使冲突历史失效；sheet rename/remove/restore/duplicate 与 workbook restore 对带名称 patch 的历史采取保守失效，避免空 affected-range 绕过冲突判断。
6. 即使 patch 已改模型，mutation listener 原先只拿到 handler effect；若 handler 未返回名称 delta，FormulaEngine 可能只能退回全量名称同步。undo/redo/remote 现向 listener 传递对应方向的精确 delta；source-level 回归覆盖同名/异名远端写、undo/redo 和 stale-state 拒绝。

协同协议层仍未闭合：server-owned `StructuralPatch` v1 仍只承载 formula owner deltas，未携带上述名称 delta。不能把旧 v1 patch 静默解释为 v2 空名称列表；需先设计操作日志/Outbox 的显式迁移或 canonical patch 版本边界，再贯通 Java、协议校验与协作 ACK/replay。本轮新增测试源码但未执行任何本地测试、构建、lint 或浏览器验收；只做静态源码检查与 `git diff --check`。

### Six-pass static audit — StructuralPatch persistence and version boundary (2026-09-25, HEAD `21c67dda`)

本轮沿 protocol → Java reducer → collaboration replay → persisted history 对照当前源码，六轮分别确认：

1. TypeScript `StructuralPatch` 固定为 version 1，精确 contract 只有 `mutationId` 与 `formulaOwnerDeltas`；validator 会拒绝未知字段，不能在该版本悄悄附加 defined-name state。
2. Java `StructuralPatch` 同样把 version 固定为 1，构造器要求公式 delta 列表；`inverse()` 仅反转公式 owner。服务器/客户端版本不一致会在反序列化或协议校验处失败。
3. Java axis、cell-shift、move 与 row-permutation reducers 会直接改写 `definedNameModels` 的 formula/anchor（例如 `StructuralSnapshotReducer` 的轴变换、`rewriteCellShiftFormulas`、`moveDefinedNameAnchors` 和 `remapPermutationDefinedNames`），但返回 patch 不保留 name before/after 状态；server commit 因此无法向客户端交付该 owner 的精确 delta。
4. collaboration ACK/replay 只把 `structuralPatch.formulaOwnerDeltas` 映射到 `MutationInfo.structuralFormulaOwnerDeltas`；即使后端补出名称 delta，当前入口仍会丢弃它。
5. committed envelope 被序列化到 `operation_log.envelope_json`，并可能进入 `coordination_outbox.payload_json`；重放与结构 undo 会再次读取已提交 envelope/target patch。仅改 DTO 而不迁移这些记录会使已有操作在新版读取或 undo 时失败。
6. 当前 `CanonicalRowsPermutedMigration` 已证明持久层有显式 envelope rewrite 边界，但它对 operation log 与未发布 outbox 分别处理；已发布协同事件、checkpoint 与旧 patch 的可逆语义也必须纳入版本切换，而不是把历史缺失 owner delta 填成空数组。

收敛方案：先定一个带完整 owner coverage 的 canonical patch v2 和 fail-close 规则；实现 Java/TypeScript 同构验证及 owner delta 后，再在显式版本迁移边界从每个 workbook 最早可用且校验通过的 checkpoint 重放连续 operation log，重新导出可证明的 owner before/after，并校验各现存 checkpoint。迁移必须同事务更新 operation log 与待发布 outbox；若 operation 缺失、checkpoint 不匹配、opaque owner 无法判定或历史 patch 不能重建，迁移应中止并报告 workbook/revision，而不能制造空 delta。上线边界还需保证已发布旧事件不会与 v2 客户端混读。当前源码尚未实现该迁移，本节是经验证的阻断条件与实施约束，不宣称修复完成。

### 六轮自审复核 — structural patch v2 的持久化升级前置条件（2026-09-25，HEAD `f34c0228`）

本轮不是重复计算旧清单中的问题数，而是针对“只补 owner 字段即可升级 v1”的方案做六次独立反证；每一轮都检查真实生产入口，静态审查，无测试/构建/lint/UI 执行。

1. **TypeScript 协议面：** `protocol/src/index.ts` 的 `validateStructuralPatch` 对 key 做精确校验，固定 `version === 1`，只返回 `formulaOwnerDeltas`。因此在 v1 上直接增加名称、cell 或 metadata delta 会被客户端拒绝，不是无害扩字段。
2. **Java 协议面：** `StructuralPatch` record 也固定 v1，只有公式 owner 列表，`inverse()` 逐项反转的仍只有公式状态。TypeScript 单边扩字段不能形成跨端同一语义。
3. **提交/重放面：** `WorkbookOperationService` 先由 `StructuralMutationDescriptor` 按 mutation intent 派生候选 snapshot，再合并部分 patch；`MutationDescriptorRegistry.applyCommittedMutations` 重放时重新运行 descriptor 并比较当前可表达的 patch。服务端目前没有可供客户端直接应用的完整结构结果。
4. **客户端/历史面：** `collaboration-session.ts` 把 committed mutation 的 `params` 交给本地 handler 重放，并只把 `formulaOwnerDeltas` 送进 `MutationInfo`；协议逆 patch 同样没有名称和普通 owner 状态。故本地 handler 与服务端 reducer 仍可独立计算同一意图。
5. **持久化面：** `operation_log.envelope_json` 保存完整已提交 envelope，`coordination_outbox.payload_json` 保存待发布 envelope；现有 `CanonicalRowsPermutedMigration` 只规范化 mutation params/ranges，不能为历史操作凭空生成缺失的 owner before/after。直接切换 DTO 会令历史读取、结构 undo 或待发布消息失败。
6. **重建证据面：** 服务端 snapshot 按至少 50 个操作或 512,000 字节批量 checkpoint，并非逐 revision 快照。升级必须沿可验证 checkpoint + 连续 operation log 重放，校验中间 checkpoint 与当前 workbook snapshot；遇到 revision 缺口、校验不匹配或不支持的 owner 时必须整批中止，不能以空 delta 填充。

**已确认问题与修复方案：** 真实缺口是 v1 patch 的 owner 覆盖不完整，且服务端、协作重放和历史撤销依赖该 patch；这是一个跨层 root cause，不把六次验证冒充六个独立 bug。下一实现批次应定义完整且可逆的 v2 patch，再建立显式、事务化的历史迁移：从每个 workbook 最早可用且校验通过的 checkpoint 重放连续日志、重新计算 patch，并同事务改写 operation log 与未发布 outbox；迁移全部校验通过后才允许 runtime 只接受 v2。任何无法重建的 workbook 必须保留原数据并报告 unit/revision，禁止降级到客户端重算或默认空 owner 列表。当前没有改协议或数据，PR #345 与整改目标均未完成。

### Local snapshot frequency — six static safety passes (2026-09-25)

确认本地编辑在每个操作 journal 已持久化后仍无条件安排 1 秒完整 workbook snapshot，导致长期编辑的大 workbook 即使 journal 可恢复也持续承担全量序列化成本。本轮六次反证检查后将完整快照改为每 50 个 local revision 执行一次；journal 写入仍逐次保留，显式保存仍立即产生完整快照。

1. `runtime.ts` command-completion listener 每个根事务递增 local revision 并触发本地持久化，确认原有 timer 对单次编辑也会启动。
2. `commitLocalOperationJournal` 先把完整 pending operations 写入 workspace operation store；只有成功后才可能调度 snapshot，因此完整快照不是 journal durability 的先决条件。
3. journal 提交同时更新 workspace head 的 local revision 与 pending log；启动恢复比较 `pending.snapshotRevision` 和 `localRevision`，差值存在时重放 pending operations。
4. 完整 checkpoint 将当前 workbook snapshot 与 pending journal 一起保存，并把 pending journal 的 snapshot revision 推进到该快照 revision；runtime 在成功后同步新的 compaction baseline。
5. `saveWorkbook` 提交 native artifact 时走 `checkpointWithArtifact` 的立即 checkpoint 路径，不受 50-revision 的自动 compaction 阈值限制。
6. snapshot 写入失败不会清除已持久化 journal；阈值未到时 journal 仍可恢复，达到阈值后的 checkpoint 失败会保留旧 baseline，后续 journal 提交仍可重试。

实现位于 `spreadsheet-app/src/runtime.ts`，新增边界测试源码但未执行。静态 diff 检查通过；没有改变 operation journal schema、重放次序、显式保存语义或服务端 checkpoint policy。整项 Excel runtime 重构及前述 StructuralPatch v2 迁移仍未完成。

### Local checkpoint overlap regression — six static review passes (2026-09-25)

对上轮 50-revision compaction 再做六轮反向审查，确认异步 checkpoint 与新 mutation 交叠时的真实 journal 丢失路径，并一起修正：

1. checkpoint 在 await 前捕获 snapshot、revision 与 pending journal；期间新根事务会同步更新同一个 `OperationJournalStore`。
2. checkpoint 成功回调原先把捕获的旧 `record.pending` 无条件写回 journal cache，会抹掉期间追加的操作；现改为按 checkpoint baseline 重基当前 journal。
3. 重基保留当前 operation IDs 与单调 sequence，采用 checkpoint 的 `snapshotRevision`；因此继续保留未确认操作，但不会复活已经从当前 pending journal 移除的已确认操作。
4. journal commit 仍在 checkpoint 链后排队；若 runtime 在 checkpoint 完成前卸载，旧 disposed guard 会直接丢弃已入队的 mutation journal commit。
5. checkpoint 返回后即使 runtime 已卸载，也必须同步新 storage revision 与 journal baseline，后续已入队的 journal commit 才能用正确 CAS revision 写入；用户界面回调及 asset reconcile 则可以停止。
6. 旧 checkpoint 若晚于较新 journal baseline 才完成，不得把 baseline 倒退；新增 `STALE_SNAPSHOT_CHECKPOINT` fail-close，并保证拒绝前 cache 不变。

代码在 `features/persistence/storage.ts` 与 `runtime.ts`；测试源码覆盖新操作保留、checkpoint baseline 重基及陈旧 checkpoint 拒绝，未执行。静态审查确认了 queued journal 在 dispose 后继续完成存储写入、snapshot completion 在 dispose 后仍同步 storage revision，且不再向已卸载 runtime 派发 UI 通知。只执行 `git diff --check`；未运行测试、构建、lint 或 UI。StructuralPatch v2 与其历史迁移仍是总体整改的开放项。

### Six-pass static review — worksheet pane structural coordinates (2026-09-25)

六轮逐一做了如下静态复核：

1. 对照前端 `WorksheetPane` union 和服务端快照规则，确认服务端只校验 pane kind/state，没有实现相同坐标与分支字段约束。
2. 沿 `freeze.set` 生产入口确认其接受越过 Excel 地址空间的视口坐标、frozen split 小数及非法 `activePane`，并直接写入 candidate snapshot。
3. 反查 checkpoint/current snapshot validator，确认快照里的缺失坐标、越界坐标及 `kind: none` 携带 split-only state 原先均可通过 canonical 校验。
4. 沿 axis reducer 检查失败语义，确认 `asInt(0)` 会把缺失或非整数坐标静默转成 0，而不是拒绝损坏的 pane。
5. 分别推演 TS 与 Java 删除映射，确认删除区间穿过 frozen split 或 viewport start 时，原算法减去完整 count，会把坐标移到删除区间之前。
6. 反查 staged metadata preflight 与服务端 owned-snapshot 提交边界，确认位移后未校验 pane 上限，越界状态能随结构编辑提交；两侧应共享映射规则并在事务发布前 fail-close。

这些证据确认了 pane contract 缺口、删除区间映射错误及位移后缺少边界验证，不把同一字段的多个坏值虚增成 30 个独立问题。`freeze.set` 曾接受超出 Excel 地址空间的视口坐标、冻结计数小数及非法 `activePane`；canonical snapshot validator 只检查 kind/state；结构 reducer 再用 `asInt(0)` 把缺失或非整数坐标静默变为 0。该缺口会在结构编辑后制造与前端 `WorksheetPane` 不同的状态。

现在 `WorkbookSnapshotValidator.requireCanonicalPane` 统一校验服务端 freeze mutation ingress、快照入口及 Java 结构 reducer；reducer 不再提供 0 默认值。复核还发现并修复第二个真实语义错误：删除区间若穿过 frozen split 或 viewport start，旧算法减去完整 count 会把边界映射到删除区间之前；TS/Java 现在均将区间内边界钳到删除起点，并在映射后检查坐标上限。新增 validator、mutation 及客户端 pane boundary 成功/拒绝路径测试源码；按本任务要求未运行测试、构建、lint 或 UI。本节记录两个独立缺陷，且六轮审查不等于声称本轮满足每轮 30 项的总体审计目标。总体结构 patch v2、历史迁移与其余 owner families 仍未完成。

### Six-pass follow-up — local freeze mutation ingress (2026-09-25)

在上节完成后追加六轮交叉自审，并据此修正文档中的边界描述：

1. 对照 `WorksheetPane` 类型和 core-model 结构前置校验，确认它们使用同一行列上限，但检查入口不同。
2. 沿 `sheet.freeze.set` command → `CommandRuntime.assertMutation` → mutation schema → handler，确认本地命令确实经过 schema preflight，不是服务端验证的重复路径。
3. 检查旧 `isWorksheetPane` 的 frozen 分支，确认其把任意非负 finite number（包括小数）当作合法冻结计数。
4. 检查 start 坐标，确认旧入口只要求非负整数、不限制 Excel 最大行列索引。
5. 检查 kind/state 与 activePane，确认旧入口未验证二者与 `WorksheetPane` 分支契约一致，`kind: none` 也会接受附加冻结字段。
6. 对照 Java `requireCanonicalPane` 的每个分支与既有 TS structural preflight，确认无意扩宽 split 的 native fractional units，唯一需要修改的是未被调用的本地 mutation ingress；成功/拒绝路径测试现覆盖相同入口及 history 原子性。

此轮确认一个真实缺陷：前端 `freeze.set` schema 接受非规范 pane 并可将它写入本地模型/undo history，结构变换随后才拒绝；服务端已拒绝不构成本地路径的保护。现将完整 TS pane 判定集中在 core-model，结构变换与 feature mutation schema 共用同一规则。未把其他被六轮排除的候选计为问题。新增测试源码未执行；本轮仍只允许静态审查，整体 patch v2 / 历史迁移保持未完成。

### 六轮静态复审 — pane schema exactness and TS snapshot ingress (2026-09-25)

继续沿已修复的 pane 边界做六轮独立检查：

1. 对照 `WorksheetPane` union，确认 pane 只允许 `kind/state/xSplit/ySplit/startRow/startColumn/activePane`，none 分支只允许 `kind`。
2. 反查前端 mutation 判定，确认 exact-key 缺口会令 `referenceHint` 等未知 own field 仍被当作 canonical，并随 frozen/split spread 保留。
3. 追到 Java `requireCanonicalPane`，确认服务端也只验证已知值，既不拒绝 frozen/split 的额外字段，也不拒绝 none 的未知字段。
4. 追到 `assertCanonicalWorkbookSnapshot` 与公开 `WorkbookModel.fromSnapshot`，确认前者有第三份更弱规则，后者可被直接调用绕过 snapshot assertion；二者都未检查完整 pane contract。
5. 核对正常 OOXML unknown-part 保留走原始包 capability/export 路径，不依赖在 canonical pane JSON 任意塞字段；exact-key 拒绝不会替代或删除 raw OOXML parts。
6. 核对拒绝发生在 local mutation schema、TS snapshot hydration 及 Java snapshot/mutation 边界，测试源码分别覆盖有效本地 pane与损坏/未知 pane 拒绝。

确认的根因是 pane canonical schema 在多个入口各自实现且未约束对象 key。现在 TS core-model 的单一 validator 拒绝未声明 key 与缺失的必需字段，结构变换、freeze mutation、`assertCanonicalWorkbookSnapshot` 及公开 `WorkbookModel.fromSnapshot` 共用它；Java 对 frozen/split 与 none 执行相同的 exact-key 校验。增加 TS/Java snapshot 与 mutation 拒绝用例源码，未运行测试或构建。该项不代表 StructuralPatch v2、ReferenceIndex 收敛、OOXML planner preflight 或整体目标已完成。

### 六轮静态复审 — operation history continuity and identity (2026-09-26)

本轮只确认并修复一个根因，不把多个调用点重复计数：

1. 核对提交入口，确认每次成功操作将 workbook revision 精确加一；因此 checkpoint revision 到目标 revision 之间的 operation history 必须逐号连续。
2. 核对 `WorkbookStore.listOperationsBetween` 与 repository 查询，确认它只按 revision 区间过滤并升序返回，不检查缺失 revision、重复 revision 或记录身份。
3. 追踪 `currentSnapshot`，确认它直接应用查询返回的记录并返回快照，没有验证首尾覆盖；缺项会静默留下旧状态。
4. 独立追踪 `snapshotAtRevision`，确认同样只遍历返回记录；snapshot canonical validator 只验证快照结构，不能检测事件历史缺项。
5. 追踪 stale-base commit 的冲突检查，确认缺失的 intervening operation 会被跳过，可能漏掉结构变化或重叠范围冲突。
6. 检查数据影响与反例：`currentSnapshot` 会被后续 commit 复用并增加新 revision，故不完整重建可能被继续固化；同时静态复核修复本身，发现并显式保留 `SYSTEM` restore 的服务伪主体与 envelope 操作者差异，避免误拒绝合法 restore 记录。

根因已在 `WorkbookOperationService` 收敛到连续历史读取：current snapshot 重建、历史 revision 重建与 stale-base 冲突检查共用完整性校验，缺口、重复、错误边界及数据库行/envelope 身份不一致均以 `STORAGE_CORRUPT` fail-close；系统 restore 只接受其专用 `system:workbook-restore` 行身份。静态 diff 检查之外未运行测试、构建、lint 或 UI。本轮仅确认一个独立根因；不声称满足此前每轮至少 30 个问题的总体验收偏好，StructuralPatch v2、历史迁移及整体架构目标仍未完成。

### 六轮补充复审 — 非顺序 operation lookup 的 envelope 身份（2026-09-26）

在补入连续日志校验后，又对所有持久化 envelope 读取点做了六轮静态复核：

1. 检查 `operationResult`，确认单条查询只验证调用者与数据库行归属，随后直接返回 envelope。
2. 检查 commit 幂等重试，确认已存在 operation id 时直接用 envelope 比对请求，未校验 envelope 与 operation row 的 revision/sequence/session/actor 一致性。
3. 检查 undo target 读取，确认只校验数据库行主体与 revision，原先仅比较 envelope 的 operation id；结构 undo 会进一步依赖该 patch 推导逆 patch。
4. 检查 replay 中的 undo target，确认即便当前事件来自已验证的连续日志，目标 operation 仍是独立查找并使用，不能依赖当前行校验间接覆盖。
5. 检查 `revisionRecord`，确认返回值同时包含数据库行 revision/time 与 envelope 内容；身份分叉可能对外暴露相互矛盾的 revision 记录。
6. 对照 SYSTEM restore 的行主体约定，确认通用行/envelope 校验必须保留服务伪主体映射，并由数据库列与 envelope 内的 session 值彼此相等，而不能硬编码只接受新写入的 `system` session。

确认的是一个根因：非顺序读取绕过了 operation-row/envelope 身份一致性边界。所有直接业务读取现统一经过 `readCommittedHistoryRow`，原始 JSON 反序列化只留在该校验器内部；SYSTEM restore 继续按其专用行主体验证。未运行测试、构建、lint 或 UI；该检查和前一轮的连续性校验均只做静态复核，不把两个根因虚增为 30 项，也不代表完整结构架构已完成。

### 六轮静态自审 — coordination outbox revision identity (2026-09-26)

沿 operation commit → Outbox → Redis → WebSocket 做六轮核查，确认：

1. `WorkbookOperationService.enqueueRevisionEvent` 从同一 committed envelope 构造 Outbox 行，但 unit/operation/revision 另存为独立数据库列。
2. `CoordinationOutboxPublisher` 将数据库列写入 Redis 外层消息，同时原样附上 payload；此前没有比较两套身份。
3. `RedisCoordinationSubscriber` 只将外层 `operation` 解析为 envelope，忽略外层 unitId、operationId、revision。
4. `WebSocketSessionRegistry.broadcastRevision` 按内层 envelope 的 operationId 去重，并按内层 unitId 选择接收会话，故身份不一致会改变实际投递目标。
5. WorkbookOperationService 的 operation_log 校验器不覆盖 outbox 发布路径；Outbox 的独立 payload 即使与其行身份不一致仍可进入 Redis。
6. 失败语义已有安全落点：publisher 异常会释放 Outbox 供既有重试，subscriber 异常会忽略该消息；因此可在消息边界拒绝不匹配事件，不必将坏数据转成 workbook 状态。

现新增 package-private `RevisionCoordinationEvent` 作为两端共用的强类型事件，构造时核对 UUID、kind、unitId、operationId 与 revision；publisher 先解析规范 envelope 再构造事件，subscriber 反序列化同一 DTO 后才广播。六轮确认的是一个独立身份边界缺口，并非六个问题。只做静态源码/diff 检查；不运行测试、构建、lint 或 UI，StructuralPatch v2 和整体架构目标仍未完成。

### 六轮静态自审 — Sheet Table 工作簿级身份与复制引用（2026-09-26）

六轮分别沿公式解析、依赖索引、前端命令、快照导入、服务端 mutation 与 Sheet 复制做反证，确认一个根因及其入口表现：Sheet Table 名称被公式引擎当成工作簿级 key，实际校验却局限于单 Sheet 或完全缺失。

1. `normalizeSheetTables` 用 `trim().toUpperCase()` 建 map，重复名称原先后写覆盖前写；公式 AST 只保存 tableName，无法用当前 worksheet 消除歧义。
2. `FormulaEngine.setSheetTables` 和 `tableReferenceIndex` 都以该名称作为受影响公式的 key；冲突会把公式依赖和 evaluator 解析指向同一个被覆盖的表，而不是仅造成显示重名。
3. 前端 `sheetTable.add`/`update` mutation 与 command 只检查目标 sheet；`update` 原先甚至不检查同一 sheet 的名称重名，跨 sheet 更无唯一性门禁。
4. TS `assertCanonicalWorkbookSnapshot` 与 Java `WorkbookSnapshotValidator` 原先只核对范围/列宽，不检查跨 sheet 表 ID/名称唯一性；导入和历史快照因此可把歧义对象带入运行时。
5. Java `SheetDataMutationDescriptor` 的 upsert 也只验证目标 sheet 元数据，协作/持久化路径不会替客户端补做 workbook 级唯一性检查。
6. `duplicateSheet` 已全局重分配表 ID，却原样克隆 Table 名及副本公式；直接采用全局唯一规则会让过去的静默错指暴露为冲突，因此复制边界必须同时派生新名并改写副本内引用。

已修复：公式引擎拒绝重复 ID/名称而不再覆盖；前端命令、服务端 mutation 及 TS/Java canonical snapshot validator 统一执行 workbook 级唯一性校验；复制 Sheet 为克隆表分配不冲突名称，并用 AST 重写副本单元格、保留公式、验证规则、Sheet-scoped names、shape/chart 文本公式中的结构化引用。追加复核还确认服务端 `sheetTable.update` 原先能把不存在的 ID 当新增项 upsert，现与前端一致地要求目标身份已存在；非规范空白 ID 也在两端边界拒绝。遇到无法安全改写的 preserved-only 公式时复制 fail-close，不留下部分 Sheet。新增了公式索引与 Sheet 复制的回归测试源码，但按本轮要求没有执行测试、构建、lint 或 UI；只进行源码及 diff 静态审查。该六轮验证确认的是一个根因的六条证据，不虚报成六个独立缺陷；整个 Structural Editing & Reference Integrity 目标仍未完成。

### 六轮静态自审 — Sheet Table rename 的公式 owner 完整性（2026-09-26）

六轮分别沿命令入口、公式重算、Undo/Redo、服务端 reducer、协作 patch 与 OOXML 序列化反证，确认一个真实跨层缺陷：表名更新只修改 Sheet Table 元数据，没有把结构化引用的 formula owner before/after 纳入同一事务。

1. `WorkbookSession.setActiveSheetTableName` 只将新名称送入 `sheetTable.update`；命令 mutation 的逆操作只保存旧表模型，没有公式 owner 状态。
2. `FormulaEngine.setSheetTables` 按表名变化查询并标记依赖，但公式 AST 仍持有旧 `tableName`；运行时随后重算时旧名已从表索引消失，相关公式解析为 `#NAME?`，不是单纯缓存未刷新。
3. 本地 Undo/Redo 重放 `sheetTable.update` 的前后表模型，不带结构化引用 owner delta，因此正向 rename、Undo、Redo 没有共享一个可逆引用变换。
4. Java `SheetDataMutationDescriptor` 对 `sheetTable.update` 只验证并 upsert 表对象；该 mutation 不走生成 StructuralPatch 的结构 reducer，服务端持久快照会保留旧公式文本。
5. TS/Java StructuralPatch v1 对 key 与版本做精确校验，只覆盖 cell、rule、chart-text owner，且 mutation 白名单不含 `sheetTable.update`；协作 ACK/Replay 仅转交 patch 中已有的 formula owner deltas，不能补出缺失变换。
6. OOXML cell writer 从当前公式文本生成 `<f>`，table writer 则输出新名称；因此保存会把新 Table 名与仍引用旧名的公式一起写入文件，持久化不会自行修复。

这些是同一 owner-transaction 缺口在六个真实消费者上的证据，不计作六个独立根因。修复边界必须是完整切片：由 canonical rename planner 生成全部受影响 owner 的 typed before/after delta；服务端提交、协作、Undo/Redo 与 OOXML 都消费该 delta；同时在显式迁移边界从校验通过的 checkpoint 和连续 operation log 重建旧 rename 的 owner delta，并事务更新历史 envelope 与未发布 outbox。当前 v1 精确契约无法承载完整 owner 集，也没有安全的纯字段升级方式；在该迁移与 v2 消费链闭合前，不提交单端公式改写或仅拒绝常见 rename 的局部补丁，以免客户端、服务端和历史快照产生新的语义分叉。

本轮只完成静态自审和有界修复方案记录；未改运行时代码，未运行测试、构建、lint 或 UI，也未将该缺陷标记为已修复。后续实现需在同一 PR 中完成协议、迁移、所有 owner 家族及回归用例后再验收。

### 六轮静态自审 — Defined Name owner identity（2026-09-26）

另沿公式引擎归一化、模型存储、快照 ingress、服务端快照校验、服务端名称 mutation 与跨层 identity key 做六轮复核，确认一个独立根因：同一 scope/sheet/name 的重复定义被 Map/首项查找静默折叠，且不同层对 sheetId 大小写使用了不一致的 identity 规则。

1. `normalizeDefinedNameModels` 通过 `Map.set` 处理同 key 定义，后项覆盖前项；公式依赖和求值因此取决于输入顺序。
2. `WorkbookModel.replaceDefinedNames` 及 `fromSnapshot` 的逐项 `setDefinedName` 同样覆盖重复 canonical identity，没有拒绝边界。
3. TS `assertCanonicalWorkbookSnapshot` 原先未检查 `definedNameModels` 的 key、scope、sheet 引用或唯一性；快照校验通过后模型初始化才折叠数据。
4. Java `WorkbookSnapshotValidator` 原先只在 hyperlink target lookup 中读取定义名，没有验证 `definedNameModels` 本身，服务端可接受客户端随后无法无损加载的快照。
5. Java `WorkbookStateMutationDescriptor.nameIndex` 首次命中即返回；遇到已有重复定义时 `name.set`/`name.remove` 只操作第一项，留下未显式报告的第二个 owner。
6. FormulaEngine 原先对 sheetId 也做大写折叠，但 `WorkbookModel.definedNameStoreKey` 和 `ReferenceIndex` 使用精确 sheetId；两个不同且大小写不同的 worksheet ID 会被公式引擎误认为同一 owner。

第六轮复核另确认 `definedNames` 派生映射入口会绕过上述检查：快照缺少 `definedNameModels` 时 `fromSnapshot` 仍会从映射逐项 upsert；两字段并存时也未验证映射与 workbook-scope models 一致。已同步修复 TS/Java 快照边界：映射-only 输入检查 canonical key/value 与大小写不敏感唯一性；models 并存时要求映射与 workbook-scope projection 精确一致。回归源码补足 sheet-scope 重复、失败替换保留原状态、映射重复与投影不一致。该映射仍是现有 snapshot 兼容边界，后续 clean-break 应在显式 snapshot migration 中移除 fallback；本次未擅自扩大为协议删除。

第七轮跨语言边界复核发现 Java `String.trim()/isBlank()` 与 TS `String.trim()` 对 NBSP、BOM 等边界空白的定义不同，可能导致服务端接受前端 canonical ingress 会拒绝的 formula。Java 校验现显式匹配 ECMAScript trim 字符集合，补充 NBSP 拒绝用例源码；未运行测试。

第八轮沿 `name.set/name.remove` 的写入链反查发现后端 projection 按原大小写精确增删，而 canonical identity 按大小写不敏感查找；仅修改名称大小写会留下旧 map key。mutation 现先清理同一大小写折叠 identity 再写入/删除派生映射，新增 set 后再以不同大小写 remove 的成功路径源码用例；未运行测试。

第九轮核对 projection 构造器时确认合法名称 `__proto__` 会被普通对象的属性 setter 特殊处理，导致 TypeScript/OOXML 投影漏项并与 canonical models 不一致。两处现以 `Object.fromEntries` 创建 own data properties，并新增快照保留该名称的静态用例。

第十轮继续追查 `FormulaEngine` 的真实读写索引，确认其内部 `definedNameIdentity` 与 `findDefinedName` 仍将 worksheet ID 大小写折叠；仅修正 normalizer 并不能阻止后续 Map 覆盖或跨表误读。内部 key 与查找现改为精确 sheetId，新增两个大小写不同 worksheet 上同名 local name 的独立求值用例源码；未运行测试。

推送后自动 CI 两个 job 均在前端 TypeScript 检查同样失败：用于故意构造无效快照的测试变量被窄化为 `Record<string, any>`，未在调用静态 validator/fromSnapshot 前恢复 `WorkbookSnapshot` 类型。已在这些负例调用点显式加边界 cast；依据 CI 日志静态修复，未本地运行编译或测试，等待新 head CI。

已修复该切片：FormulaEngine 与 WorkbookModel 批量替换、TS/Java 快照 ingress 对重复 identity fail-close；sheet-scoped name key 在所有求值与引用索引处都对 name 大小写不敏感、对 sheetId 精确区分；`definedNames` 派生投影校验 canonical shape、唯一性和与 models 的一致性，Java mutation 对大小写变更执行同步清理/更新，TS/OOXML 投影保留 `__proto__` own key；TS/Java 校验对齐 ECMAScript trim 边界空白。新增 TS/Java 成功与拒绝路径测试源码，覆盖同名不同 scope 可并存、大小写不同的 sheet ID 独立求值、identity 冲突、缺失 worksheet、投影同步与 fail-close；按本轮要求未执行测试、构建、lint 或 UI。只做静态源码审查与 diff 检查。历史中若存在重复定义，无法从被静默覆盖后的公式行为推断原作者意图；不自动删除/合并任一项，严格 ingress 会 fail-close，需从用户确认的备份恢复后再迁移。StructuralPatch v2、历史迁移与完整整改仍未完成。

### 六轮自审复核 — StructuralPatch 定义名称 owner 缺口与 CI 负例（2026-09-26，HEAD `489359a3`）

本轮按用户要求进行了六轮独立边界复核；每轮均回到实际生产路径，不把同一结论重复计作多个不同缺陷：

1. **协议入口**：`protocol.validateStructuralPatch` 只接受 v1 的精确字段集合和 `formulaOwnerDeltas`。`StructuralDefinedNameOwnerDelta` 虽已存在于 core-model，却不在共享提交合同内；名称公式/anchor 前后态无法由服务端提交协议表达。
2. **服务端归约**：`StructuralSnapshotReducer` 的轴变换、单元格带移动、行置换会改写 `definedNameModels`，而 patch 构造器只返回公式 owner delta。快照结果因此含名称变化，patch 却无法描述这些变化；这是服务端提交、差异校验与逆变换的真实缺口。
3. **协作传输**：`collaboration-session` 在重复 ACK 与新远端操作两条路径均只复制 `formulaOwnerDeltas`。客户端虽有 defined-name replay/precondition 支持，但服务端提交数据没有进入该路径，远端结果不能依赖同一份名称 owner 前后态。
4. **Undo/Redo**：`WorkbookOperationService.inverseStructuralPatch` 仅反转公式 delta；本地 `command-runtime` 已有名称 owner delta 的历史处理。对于删除/位移导致 anchor 归并或公式变换的操作，两端撤销语义缺少统一的名称逆 patch。
5. **历史重放**：`MutationDescriptorRegistry.applyCommittedMutations` 将存储 patch 与当前 reducer 结果精确比较；`WorkbookOperationService` 从 checkpoint 加连续 operation log 重放。因而提升 patch 版本不能保留旧 patch 的运行时兜底，也不能只改线上新写入；必须在迁移边界验证 revision-0 基线、连续日志、每个中间 checkpoint 与当前快照，再重写历史 patch。
6. **outbox**：已应用的 repeatable Java migration 会直接改写 `operation_log.envelope_json` 和未发布 `coordination_outbox.payload_json`；发布器随后按持久化 row 原样发出。schema 升级必须在同一迁移边界同步重建未发布事件，且对缺失日志、checkpoint 不一致或无法确定 inverse target fail-close。

**修复方案**：一次性提升为严格 StructuralPatch v2，增加身份明确的 defined-name owner before/after（公式与 anchor）并在 TypeScript/Java 校验、reducer、impact、collaboration、Undo/Redo 与 remote replay 中使用同一合同；历史迁移从已验证的 revision-0 checkpoint 按连续 revision 重放并派生 v2 patch，同时核验中间 checkpoints/current snapshot、改写 operation log 和 unpublished outbox。迁移遇到缺失连续性、不可验证快照或模糊 undo target 必须中止，不允许 v1 fallback。此架构纵切仍在实现中，不能将本轮两处夹具修复误报为该缺口已解决。

本轮远程 CI 另确认两条测试源码问题：一条结构行变换 fixture 缺少 canonical pane 必需的 `state`；另一条越界插入负例误用 `rows.deleted` descriptor，因而并未执行所声称的插入。已分别补齐 `state: "frozen"` 并使用 `rows.inserted` descriptor。只基于 CI 失败输出与静态源码修复；未运行本地测试、构建、lint 或 UI。

### 六轮自审复核 — StructuralPatch v2 与定义名称回放 (2026-09-26)

按用户要求再次进行六轮彼此独立的静态自审；每轮都沿真实调用链核实，不将同一个根因按多个调用点重复计数：

1. **Java v2 迁移入口**：逐项核对导入、迁移包装类与 replay helper，确认新 migration 的 `Map` 声明缺少导入，会直接阻断 Java 编译；已补齐并移除未用参数/导入。
2. **快照与 owner 契约**：从快照校验追到 `WorkbookModel.fromSnapshot`，确认 Java 仍接受只有 `definedNames`、没有 `definedNameModels` 的历史投影快照；当前客户端已折叠成模型，但服务端 patch 无法表达对应名称变化。迁移现先核对快照/检查点，再把所有检查点、当前快照及 restore 内嵌快照折叠为模型；未模型化且仍含名称的结构 mutation 现在 fail-close。
3. **公式引擎调用路径**：沿本地结构命令、FormulaEngine 同步、服务端 ACK 回放确认相同 delta 会第二次到达；旧实现只接受 before 状态，会把正常 ACK 报成冲突。FormulaEngine 现对当前已等于 after 的重复 delta 幂等，对第三种状态仍拒绝，并补回归源码。
4. **服务器提交与重放**：检查 v2 exact contract、reducer diff、undo inverse 和 command-runtime 对齐，确认公式/anchor owner 的 before/after 与身份进入同一 patch，前后置条件保留；增加服务端 reducer、逆 patch 和破坏性前置条件测试源码。
5. **restore/checkpoint 连续性与历史版本边界**：逐条核对 revision-0、当前 checkpoint、全部中间 checkpoint、restore 与 operation log 重放。旧 restore 参数内嵌快照也会与规范化后的检查点失配；迁移现同步规范化其快照，并在操作日志重放前验证原 current snapshot/checkpoint 配对。回查 v1 patch 引入提交的父版本后确认，既有历史结构操作本就没有 patch/impact 字段；旧迁移因此会在真实历史上中止。现仅在重放与 checkpoint 全部吻合且两个字段均缺失时从历史派生 v2，已存在 v1 patch 则继续严格比对，部分字段状态 fail-close。
6. **未发布 outbox 与数据库副本完整性**：核对 H2/MySQL/PostgreSQL repeatable wrapper 和发布器原样发送持久 payload 的路径，确认不能忽略 outbox 自身 patch，也不能把 operation log 已升级后的新字段差异误判为旧 envelope 内容损坏。现先在任何写入前验证待发 outbox 与 operation log 源 envelope 一致，再逐 mutation 校验旧 v1 patch 字段、身份、formula delta 和 impact；对两个副本都缺失的历史格式规范升级，对副本漂移 fail-close，最后以 operation log 的 v2 envelope 重写。已发布事件不重写。

本轮六个审查边界中确认并修复九项独立的实现/数据完整性问题：迁移编译缺陷、投影-only 名称缺少 patch、该旧形状仍可进入结构写入、重复 ACK 非幂等、restore 内嵌快照未迁移、pending outbox patch 差异被忽略、迁移不兼容真实存在的 patchless 历史、升级后的 impact 字段导致旧 outbox 无法规范化、outbox 与 operation log 源副本漂移可能被掩盖。新增测试源码未执行；只允许静态审查，后续仍需远程 CI/迁移门禁与真实应用验收。当前完成的是定义名称 owner 的 v2 纵切，不代表其它结构编辑 owner families 或整体整改目标已完成。

### 六轮自审复核 — Sheet Table rename 的 owner 原子性与懒加载（2026-09-26）

本轮按用户要求沿六条独立路径复核并确认实际缺陷；不是把同一根因重复计数：

1. **公式改写器**：Java 改名扫描器在多个同名结构化引用间复用错误的源切片边界，会把前一个引用后缀重复输出；现逐段输出改名 token 与原结构化引用，并用多引用源码用例约束。TS/Java 边界同时避免把 Unicode 标识符中的 ASCII 后缀误认成目标表名、误改外部 workbook 表引用；正常字符串字面量保留。
2. **延迟单元格**：公式 owner 遍历本身不 hydrate deferred JSON，但 rename apply 与通用 Undo/Redo formula-cell patch 原先都会调用 `CellMatrix.get/set`，仍把整张工作表物化。现增加不 hydrate 的稀疏单元格读写与按行 copy-on-write，维持 deferred 状态、隔离输入快照、递增内容 revision；rename 与 Undo/Redo 源码用例都断言不 hydrate。
3. **公式来源所有权**：Java patch reducer 拒绝修改 `preservedOnly` 的 `formulaMetadata.sourceFormula`，与 TS Undo/Redo 和 rename planner 的 owner 模型冲突。现仅要求 metadata owner 仍存在，允许按前置状态更新 source formula，并覆盖 inverse/redo 源码用例。
4. **提交与 patch**：`WorkbookOperationService` 将 `applyWithPatch().snapshot()` 作为提交快照；Sheet Table descriptor 原先只返回新表元数据，公式 owner delta 只在回放阶段应用，实时提交会持久化旧公式。descriptor 现于同一 detached snapshot 应用 patch 后再返回，提交与重放使用同一状态。
5. **对象身份与协议**：table rename 必须携带 cell、rule、defined-name、chart/shape、TableSheet、data-view 和 cell-style-template owners；现由 TS/Java v3 精确契约约束新增 owner 身份，服务端查找 table ID 时要求全工作簿唯一，避免跨 Sheet 重复身份被局部查找遮蔽。
6. **迁移/恢复边界**：逐项复核 v1/v2 patch、无 patch 的旧 rename、checkpoint、restore 与待发 outbox 的同一迁移链；restore `targetRevision` 现在要求可安全转换为 `long`，避免超大整数溢出后误指向有效历史 revision。迁移仍以连续日志、已验证 checkpoint 和 operation-log/outbox 源一致性 fail-close。

本轮新增 TS/Java 成功与拒绝路径测试源码，未执行测试、构建、lint 或 UI；只允许静态审查。`StructuralPatch` 现提升为 v3，前端/服务端实现及 v2 历史升级仍需后续逐层静态复核与 CI/真实互操作验收；这只是 rename owner 纵切，不能标记整个 Structural Editing & Reference Integrity 目标完成。

### 六轮静态复审 — v3 结构 patch 类型流与定义名称 wire exactness（2026-09-26，CI baseline `d3f0fc34`）

1. **远程编译证据**：push 自动运行的两个 `canonical-build` job 在同一三个 TS 错误处失败：移动公式计划与普通结构公式计划各将 `formula-cell` 推入只推断为 `formula-object` 的数组；定义名称 anchor 校验引用了 map 回调内部的 `address`。
2. **移动引用路径**：`applyMovedFormulaRewritePlan` 同时收集 chart-text object delta 与公式单元格 delta，数组必须声明为完整 `StructuralFormulaOwnerDelta[]`，否则移动工作簿快照的前后态 patch 无法构建。
3. **结构轴/单元格变换路径**：`applyFormulaRewritePlan` 具有相同的完整联合类型契约；单独修移动路径不足以恢复插删/单元格位移相关结构入口。
4. **协议闭包与地址约束**：`validateStructuralPatch` 的 cell 地址与 defined-name anchor 共用同一 bounds/exact-key 校验；解析器现位于 patch validator 作用域，消除编译错误且不绕过边界验证。
5. **TS/Java wire 形状**：defined-name owner identity 的 `sheetId` 仅 sheet scope 必需；Java identity/state record 使用 `NON_NULL`，因此 workbook scope 的 `sheetId` 与无 anchor 状态在 JSON 中合法省略。TS validator 原先仍将这些 optional key 列为必需，导致服务端合法 v3 patch 被客户端拒绝。
6. **提交入口与正反例**：`validateCommittedOperationEnvelope` 在历史和实时 committed envelope 入口调用同一 patch validator；测试源码现覆盖 workbook/sheet scope、anchor 缺省/存在、identity 漂移、重复 owner 和额外字段拒绝。

已修复三条编译诊断对应的两个类型/作用域根因，以及一个 TS/Java optional wire 字段契约分叉；没有把相同的根因重复包装成更多问题。当前只做源码检查与 `git diff --check`，未在本地运行测试、构建、lint 或 UI。CI baseline `d3f0fc34` 的失败已确认；修复后的新 head 仍需远端 checks 验证，整体整改目标继续开放。

### 六轮静态自审 — 跨工作表结构 owner 的投影失效（2026-09-26）

六轮分别检查结构计划产出、mutation 目标身份、提交/远程回放、projection 失效策略、Canvas 快照捕获方式与 WorkbookSession 刷新顺序，确认一个真实根因：结构 patch 可改写其他工作表的公式 owner，但投影缓存原先只失效 mutation 自身工作表，且除 chart-text 外不消费 formula-owner delta。这里记录的是一个跨层失效缺陷，不按 owner 类型重复计数：

1. **结构计划产出**：`planSheetTableRename` 遍历工作簿全部工作表的公式单元格和条件格式/验证规则，delta 的 owner sheet 可不同于 Sheet Table 所属 sheet。
2. **mutation 身份**：rename mutation 的 `sheetId` 固定指向被改名表所在 sheet，不能代表所有公式 owner 的投影归属；结构变换与移动引用也可能产生跨表 owner。
3. **提交及回放**：本地 `runtime`、服务端 committed patch 与 collaboration ACK/remote replay 都把 typed formula-owner deltas 放入 mutation 通知，信息在到达投影层前没有丢失。
4. **失效策略**：`ProjectionRuntime` 原先按 mutation.sheetId 递增 revision，只对 chart-text delta 额外失效 owner drawing；formula-cell、formula-rule、shape-property、table-sheet-column owner 被忽略。
5. **快照实物**：Canvas 快照深拷贝 drawing payload 与 Table Sheet 定义，并捕获 rule 集合/条件格式运行时；缓存 snapshot 身份不变会让 React 消费者继续持有旧投影。公式单元格值可能另由 calculation callback 触发更新，不能单独作为足以证明此缺陷的证据。
6. **刷新顺序**：WorkbookSession 有 mutation 时调用 mutation 投影失效分支，而不执行“无 mutation 时”的全局公式投影失效；因此必须由结构 delta 精确标出额外 owner sheet，不能依赖无关的全局刷新。

现已按 owner 地址失效公式单元格前后工作表和公式依赖图；按 owner sheet 失效 formula-rule、shape-property、chart-text 与 table-sheet-column 的对应投影域。data-view-field 和 cell-style-template 没有被扩大到所有 Canvas sheet cache：它们不在 per-sheet Canvas 快照中，贸然全量清缓存会扩大性能成本而没有快照证据；其非 Canvas 消费者仍须在后续 owner 审计中单独核实。新增跨表单元格、规则、图形及 Table Sheet 投影的回归测试源码；只做源码审查和 `git diff --check`，没有执行测试、构建、lint 或 UI。本次只修复该投影传播切片，整体架构整改及 PR 验收仍未完成。

### 六轮静态复审 — StructuralPatch 远端 Java 编译阻断（2026-09-26，head `327f88c5`）

PR 上两个 `canonical-build` job 使用相同 head，前端依赖安装与前端构建成功，均在 backend-package 编译阶段失败。六轮沿 CI 日志、迁移 slot 契约、可选 patch 的索引语义、JSON 节点类型边界、定义名称解析及坐标拒绝路径交叉核对，确认两类源码根因；重复的 14 条 javac 诊断不按独立缺陷计数：

1. **CI job 归因**：两个失败 job 都报告同一 backend-package 步骤；前端阶段退出码为 0，因此此次失败不是刚修改的 TS projection 源码导致。
2. **迁移返回契约**：`StructuralPatchMigrationReplay` 要求每条 mutation 对应一个 `Optional<StructuralPatch>` slot；局部集合原先却声明为非 Optional，导致空 patch slot 无法表达并使方法无法编译。
3. **slot 对齐语义**：不能滤掉没有结构 patch 的普通 mutation，否则后续 patch 与原 mutation 索引错位。现保留 `Optional.empty()`，并补充连续 patchless mutation 的回归源码。
4. **JSON 对象边界**：`SnapshotMutationSupport.text` 明确只接受 `ObjectNode`；Sheet Table 和 defined-name parser 却把通用 `JsonNode` 直接传入。现先由 `requireObject` 执行类型检查再读取，畸形 owner 保持 typed validation failure。
5. **定义名称 anchor**：结构 patch 计算前后定义名状态时必须解析模型与 anchor 对象；原路径未完成对象窄化，多个调用因此触发 javac 类型错误。
6. **坐标数值边界**：anchor row/column 读取调用了不存在的 `integer(JsonNode, String)`，且不能把小数、负数、超出 Excel 上界值传给 `CellAddress` 构造器后变成非业务异常。现显式要求整数并按 row/column 上界 fail-close，新增小数和越界拒绝测试源码。

已修复 migration slot 泛型及 StructuralSnapshotReducer 的对象/坐标解析，补充 patchless slot 与 anchor 拒绝路径测试源码。只依据远端 CI 日志和源码契约静态修复，未运行本地测试或构建；head `327f88c5` 的远端检查已失败，修复提交后的远端重验仍待完成，整体目标继续开放。

### 六轮静态自审 — Sheet Table rename 与保留型 data-table 来源公式（2026-09-26，head `9ffd4051`）

六轮围绕同一个 CI 失败点逐层核对，并把同根因的客户端/服务端门禁作为一个问题记录：

1. **CI 与失败夹具**：后端编译通过，Java 测试中仅 Sheet Table rename 断言失败；输入是 `kind=dataTable`、`preservedOnly=true` 且只有 `sourceFormula` 的 owner，失败确由通用公式组门禁触发。
2. **OOXML 数据语义**：导入器将 Excel data-table 公式保存为 preserved-only provenance，`sourceFormula` 与 `range` 分开；本操作只对可识别 structured reference 的 table token 改名，不改 data-table range 或计算语义。
3. **owner 枚举**：复核 `forEachFormulaOwner` 后确认其包括 provenance-only `sourceFormula`，因此 owner 遗漏不是缺陷，不作无证据修复。
4. **TypeScript 规划器**：确认 `planSheetTableRename` 已得到完整 owner delta，但 `hasFormulaGroupMetadata` 不区分明确的 dataTable source-only 重写与需要组操作的 shared/array，造成真实的本地拒绝。
5. **Java reducer**：独立沿 `renameSheetTableReferences` 复核同一宽门禁；这也是当前 CI 失败的服务端拒绝点，提交 patch 的 before/after state 本身可表达 sourceFormula。
6. **提交、撤销与失败边界**：客户端命令应用及 Java patch setter 保留 metadata 其余字段并支持 inverse；现有 Java 用例覆盖 undo/redo。shared/array 仍须拒绝，结构公式解析失败仍须原子拒绝，worksheet rename 的 preserved-only 拒绝也保持不变。

现仅允许 `preservedOnly dataTable` 且没有可执行公式、仅其来源公式因 table token 改名的窄场景；`range`、bar-code owner 和其他公式字段均不得随之改变。新增 shared-formula 拒绝路径源码，并强化 dataTable range 保持断言。只运行 `git diff --check` 作静态补丁检查；没有运行本地测试、构建、lint 或 UI。修复后的远端 CI 尚待新 head 验证，Structural Editing & Reference Integrity 总体目标仍继续开放。

后续 head `d336b525` 的 CI 将该路径推进到 owner 变换后，暴露原 Java 测试快照缺少两张 worksheet 的 canonical `name`；调用栈落在 `StructuralSnapshotReducer.identity` 的必需字段校验。现只补齐 fixture 名称，没有放宽生产校验；新 head 的远端检查仍待确认。

### 六轮静态复审 — Sheet Table patch 推导的输入不可变性（2026-09-26，head `42d8b0aa`）

后续 CI 已越过 owner identity 并报告原快照在成功的 `applyWithPatch` 后被修改。六轮沿差异内容和读写边界确认一个生产根因：

1. **失败断言**：失败位于原快照等值断言，不是变换结果断言；变更发生在调用者提供的 `snapshot`。
2. **变化字段**：Surefire expected/actual 首个差异是原始 sheet 新增空 `conditionalFormats` 与 `dataValidations`，证明 patch 推导带来旁路写入。
3. **对象所有权**：`SnapshotMutationSupport.root` 返回相同 `ObjectNode`，不是副本；传入 before-snapshot 的读取路径必须只读。
4. **helper 语义**：`SnapshotMutationSupport.array` 在字段缺失时会 `parent.set` 创建数组；它适用于拥有中的 mutation reducer，不适用于 immutable patch preflight。
5. **调用范围**：table 查找遍历可缺省的 `sheetTables`，formula-rule 扫描遍历可缺省的规则数组，`ruleRanges` 又可能为缺失必需 ranges 写入空数组；三条路径都要避免静默修改 before state。
6. **修复边界**：使用非变更式 optional-array reader 读取 table/rule 列表；公式规则 `ranges` 缺失时 typed validation fail-close。这样避免为只读计划额外深拷贝整本 workbook，也不生成伪空 owner state。

已把 fixture 的第二张 sheet 设为省略 `sheetTables`，保留首张 sheet 缺省 rule arrays；成功路径的原快照等值断言因此同时覆盖三种可选字段。新增带公式但缺少 ranges 的拒绝且不变更输入的测试源码。只做静态审查和 `git diff --check`，未运行本地测试/构建/lint/UI；新 head 远端检查待运行结果。

### 六轮静态复核 — 结构预检的非目标工作表克隆边界（2026-09-26）

沿目标架构中的全表预检成本做六轮相互独立的静态核对，确认一个真实性能问题及其最窄安全修复：

1. **调用入口**：`preflightAxisMetadata` 与 `preflightCellShiftMetadata` 都曾对每张工作表调用完整元数据克隆器；两条路径均属结构编辑热路径。
2. **Axis owner 枚举**：行列预检跨表仅传递 conditional format、data validation、pivot、sparkline、drawing payload 与 hyperlink target；非目标表的 merge、drawing anchor、review、freeze、filter、print 等集合并未被该遍历消费。
3. **Cell-shift owner 枚举**：单元格位移的跨表阶段同样只读写规则、pivot、sparkline、drawing payload 与 hyperlink；工作表表格、绘图锚点、review、spill 和 print 范围只属于目标表。
4. **暂存隔离**：上述被访问的跨表 owner 集合必须深拷贝，否则预检会改写真实工作簿；实现仍结构化克隆这些集合，但 workbook tables 与 sources 只克隆 sourceRange 指向目标表的对象，print document 只暂存目标表对象。
5. **目标表完整性**：目标表仍使用完整元数据克隆器，避免缩窄目标 owner 集合而改变既有成功/拒绝语义。
6. **负向回归边界**：新增测试源码用非目标工作表 `drawings` 的读取陷阱证明预检不再触碰不相关元数据；仅静态检查该用例可达结构入口，未执行测试。

**修复**：两条预检路径现在对目标表完整暂存，对非目标表只暂存六类实际跨表引用 owner；workbook tables 与 sources 也只克隆指向目标表的对象，避免复制无关的大型元数据集合。`git diff --check` 通过；没有运行本地测试、构建、lint 或 UI。此修复不消除逐表 owner 遍历，也未替代全类型 `ReferenceIndex`、Canonical Structural Planner 或 patch 原子应用；剩余架构目标继续开放。

### 六轮静态复核 — 结构变换遗漏非图表公式对象 patch（2026-09-26）

六轮分别核对状态生成、delta 编码、历史/投影消费者、Java reducer 与 wire 契约，只确认一个跨端根因：

1. **Owner 收集**：轴插入/删除、单元格位移和 move 的公式预检都会收集 table-sheet 列、shape property、chart text、data-view field 与 cell-style-template 的公式变化；这些字段已实际写回模型。
2. **TS delta 编码**：`structuralFormulaObjectDelta` 已能为五类公式对象生成类型化 delta，但轴/cell-shift 的 `applyFormulaRewritePlan` 与 move 的 `applyMovedFormulaRewritePlan` 都只保留 chart-text，丢弃其余四类。
3. **History 契约**：`CommandRuntime.applyMutation` 用返回 delta 记录 inverse history；提交 patch 与本地 history owner 集不一致时会将条目作废。因此缺失 delta 会让已有结构操作在服务器确认后失去可撤销历史。
4. **Projection/remote 契约**：projection runtime 对 chart-text、shape-property、table-sheet-column 使用公式对象 delta；本地 effect 未带这些 owner 时，提交前缺少对应 owner 失效，服务器 patch 又会与本地记录不一致。另追踪 data-view/template 调用点后确认它们没有 per-sheet 投影缓存，分别由模型集合和 session getter 直接读取，不因其缺少 projection 分支再增加全表失效逻辑。
5. **Java patch 生成**：Java axis、cell-shift、move reducer 会同步改写五类 owner，但调用默认 `includeNonChartObjectDeltas=false` 的重载；Sheet Table rename 已证明同一 reducer 可发出这些 owner 类型。
6. **协议与拒绝边界**：StructuralPatch v3、TS protocol 与 Java reducer 已支持全部五类公式对象 owner，因此无需变更 wire shape；既有 anchor 删除用例继续验证拒绝不写入快照。

现已让 TS 两条应用路径对所有 staged 公式对象统一生成 delta，并让 Java axis/cell-shift/move patch 发出同一组 owner 变化。TS fixture 覆盖三类操作的 owner-kind 集；Java fixture 覆盖 13 个轴插入 owner delta、轴 patch inverse 以及 cell-shift/move 的完整对象 delta。只做静态源码审查和 `git diff --check`，未运行本地测试、构建、lint 或 UI。没有更改 patch 字段或版本；本修复不补齐 permutation、sheet identity、OOXML 或普通非公式 metadata patch，Canonical Planner/ReferenceIndex 与完整跨层链仍未完成。

### 六轮静态自审 — 定义名称的协同结构重放（2026-09-26）

六轮从不同边界核对同一个候选问题，只计为一个协同契约缺口：

1. **模型契约**：`DefinedNameModel` 的相对公式 anchor 是 `anchor: CellAddress`，不是普通 `formulaAnchor` 字段；结构式计算按 `anchor`、sheet scope 的 A1 默认上下文解析。
2. **生产 mutation 路径**：注册的领域 mutation ID 是 `name.set`/`name.remove`；`name.set` 的 inverse 会携带此前完整模型，因此已锚定公式确实能进入协同事件 payload。
3. **分类边界**：协同 kind map 未列这两个已注册 mutation，generated capability 也没有提供可用的 collaboration kind，故它们一直落入 `unknown`。
4. **失败行为**：`rebaseMutation` 对结构历史中的 pending `unknown` 立即抛 `STRUCTURAL_REBASE_CONFLICT`。这是可观察的 fail-close，不会静默损坏数据；真实影响是带定义名称的有效离线操作被拒绝重放。
5. **坐标/公式所有权**：通用 `transformParams` 只读取 `formulaAnchor`，不会移动 `model.anchor`；同时会把 `model.formula` 错按 envelope 的主工作表解析。单纯注册 kind 会产生错误公式 owner 和过期 anchor。
6. **缺省上下文与拒绝路径**：FormulaEngine 使用显式 name anchor，sheet-scope 默认 A1；无 anchor 的 workbook name 没有可推断的相对 owner。因此只变换有明确 owner 的公式；无 anchor 的 workbook 名称仅允许安全映射全限定引用，遇到相对引用即 fail-close；删除 anchor 的结构变更也必须拒绝。

**修复**：将 `name.set`/`name.remove` 分类为 defined-name；名称 set 使用专用变换，按模型 anchor 或 sheet-scope owner 改写公式并映射 anchor，remove 保持坐标不变。新增成功与拒绝路径回归测试源码，覆盖 anchored 公式/anchor 同步移动、全限定 workbook 名称、无 anchor 的相对引用、anchor 删除及 name removal。仅静态审查；未运行测试、构建、lint 或 UI。此修复不声明其它非结构性名称冲突已解决，Canonical Structural Planner/ReferenceIndex 与完整跨层迁移继续开放。

### 六轮静态复核 — workbook 名称的单前缀范围引用（2026-09-26）

对上节新增的 owner-context 识别再次执行六轮交叉复核，确认并修复一个真实的合法公式拒绝回归：

1. **词法形式**：`Target!A10:A20` 是一个显式指向 `Target` 的范围，而不是一个目标表引用加一个 owner-relative 端点。
2. **AST 形状**：parser 把 `Target` 放在 range start 的 reference 上，右端点 `A20` 的 `sheetId` 留空；这是合法范围的标准 AST 表示。
3. **现有变换语义**：`transformStructuralRange` 以 `start.sheetId ?? end.sheetId` 判定范围 owner，明确支持单端带 qualifier 的同表范围。
4. **失败触发点**：workbook name 无 anchor 的前置检查曾对两个端点分别检查并用 OR 汇总，因而把任一空字段误当作相对 owner，阻断本可安全转换的全限定范围。
5. **影响边界**：此问题仅影响无 anchor 的 workbook name 范围引用；带 anchor、sheet-scoped 默认 owner、相对引用 fail-close 与删除 anchor 拒绝路径不受该修正放宽。
6. **回归锁定**：新增 `SUM(Target!A10:A20)` 跨行插入的期望输出源码，要求左右端点随引用目标一起移动且只保留一处 qualifier；未执行测试。

现改为仅在 range 两端都没有 qualifier 时判定为 owner-relative。`git diff --check` 通过；未运行本地测试、构建、lint 或 UI。该复核修复的是上一提交引入的一处合法输入拒绝，不新增第二个独立架构根因。

### 六轮静态自审 — 原生文件 untouched-save 快照身份（2026-09-26）

1. **影响路径**：OOXML、文本、ODS、SSJSON、SJS、BIFF 与 XLSB 均可在快路径返回 artifact 的原始 `sourceBytes`；这条路径的唯一快照判据是 `sourceSnapshotHash`。
2. **摘要强度**：旧 `nativeSnapshotHash` 是 FNV-1a 32 位值，输出空间只有 2^32；不同快照必然存在碰撞，且该算法不具抗碰撞性，不能充当“快照完全未变”的证明。
3. **实际后果**：一旦不同快照碰撞，导出会绕过格式写入器并返回旧字节，当前快照编辑因此静默丢失；源字节自身 SHA-256 checksum 不能证明快照身份。
4. **边界区分**：对 `unitId` 与 print-document `unitId` 的规范化是既有跨会话等价规则，保留不变；缺陷是归一化结果被 32 位摘要压缩，而不是这些字段被忽略。
5. **所有权范围**：OOXML 与其它 codec 的快路径复用同一摘要 helper，必须一次升级公共 artifact 契约并将调用者改为异步验证，不能只补 OOXML 单一路径。
6. **持久化迁移**：旧 artifact 可能在 session-memory record 中出现；迁移不得把旧 FNV 值“升格”为 SHA-256。v1->v2 显式迁移校验 artifact 后移除旧快路径身份，保留原字节与 package graph；新 artifact 只接受 SHA-256 身份。

**修复方案**：快路径改用 Web Crypto SHA-256，codec revision 和 native-document record 升至 v2；仅在 `LocalNativeDocumentStore.load` 的显式迁移边界识别 v1，成功校验后清除旧快路径身份并原子写入 v2，非法旧 artifact 保持拒绝且不改存储。新增 hash 变更与迁移成功/拒绝路径测试源码。静态执行 `git diff --check`；未运行本地测试、构建、lint 或 UI。旧 artifact 在完成一次正常导出之前不会走原字节捷径；若其内容不能安全重写，既有 fail-close 规则仍会拒绝，不声称本次已完成 OOXML opaque-owner 的结构变换。

附加边界检查还发现存储读取用 truthiness 把损坏的 falsy record 当成“无记录”；现在只有 key 缺失（`undefined`）返回空，`null`/其它错误值 fail-close，回归源码同时覆盖无效 v1 hash 与 `false`/`null` record 且确认失败不写入。远端首次编译进一步确认 TypeScript 将 `null` 保留为可能值；显式分支已补齐，该处未在本地构建。

### 六轮静态自审 — 行置换漏更新跨表 Pivot/Sparkline 源区域（2026-09-26）

六轮独立核对确认一个真实 owner 枚举缺口，影响 Pivot 与 Sparkline 两类工作表源引用：

1. **模型契约**：Sparkline 插入 helper 明确说明目标工作表持有对象、`sourceRange` 可属于另一工作表；Pivot 的单源/多源 worksheet source 均使用带 `sheetId` 的 `RangeRef`，没有同表限制。
2. **TypeScript 预检**：`validatePermutationMetadata` 只读取被排序工作表的 `sparklines`/`pivots`。映射器按 source `sheetId` 判断是否受影响，因此跨表 owner 根本不会进入预检，也就不会在非连续结果时 fail-close。
3. **TypeScript 写回**：`applyRowPermutation` 同样只写目标表集合；对照之下，跨表 drawing payload 已遍历全工作簿并改写引用范围。对象的锚点仍须按其实际 `sheetId`/pivot target 决定，不能随着源区域一起移动。
4. **Java 预检**：`validatePermutationMetadataExact` 已接收 workbook root，但 Sparkline 与 Pivot 遍历仍限于目标 `sheet`；`PivotMutationDescriptor.forEachWorksheetSourceRange` 只枚举 range，不会弥补 owner 遍历边界。
5. **Java 写回**：`remapPermutationMetadata` 也只更新目标表的 Sparkline/Pivot 集合。`writeSingleRange` 已正确按 range sheet ID 映射，根因是没有访问跨表 owner，而不是坐标映射代数。
6. **对照与回归边界**：现有跨表 drawing-source 行置换用例证明引用源应跟随源单元格移动；旧 Pivot/Sparkline 用例缺失。新增成功用例断言源范围移动而 owner 锚点及另一表源范围不变；拒绝用例分别断言无法表达为单一区域时两类 owner 都在快照变更前被拒绝。

**修复**：TypeScript 与 Java 的预检和写回现在均枚举 workbook 中所有 Pivot/Sparkline 源 owner；只映射与被排序范围相交的源区域，以保留其它 `RangeRef` 的对象身份；Sparkline anchor 只跟随其所在 sheet，Pivot anchor 只按 target sheet 映射。新增 TS/Java 成功与 split-range 拒绝回归测试源码。实现自审还确认若无相交判断，TypeScript 的精确映射器会克隆并重赋无关范围，可能引起无必要的下游刷新；此路径已收窄并用对象身份断言覆盖。进一步核对发现服务端 `SnapshotMutationSupport.array` 会把缺省集合物化为空数组，因此全表只读预检/枚举改用 `existingArray`，并用无关 sheet 的字段缺省断言防止快照被扩写。仅静态审查；未运行本地测试、构建、lint 或 UI；`git diff --check` 通过。全局 StructuralPatch/ReferenceIndex 整改仍未完成。

**远端门禁反馈**：首个 head 的两个 `canonical-build` 均只报告新增 Java 用例的 `affectedColumnEnd` 夹具错误：sheet `columnCount=2`，测试却传入 0。按 `SheetRuleLifecycle.affectedColumnEnd` 的 `columnCount - 1` 下界改为 1；仅测试上下文修正，不改生产逻辑。第二个 head 随后显示负向用例只调用 `registry.prepare`，并未调用执行结构重映射的 descriptor `apply`；两条拒绝断言现都触发 `apply` 并继续检查原快照不变。以上均为测试夹具/触发路径修正，不改生产逻辑；本地仍未执行测试或构建。

### 六轮静态自审 — 行置换遗漏工作簿级数据源范围（2026-09-26）

六轮分别从模型契约、下游消费、TS 变换、Java 变换、操作重放与失败边界复核，确认一个跨端 owner 漏洞，涉及 workbook table 与 data-source 两种范围：

1. **模型所有权**：`WorkbookTableModel.sourceRange` 被定义为 sheet-backed table 的 canonical source range；`DataSourceManifest.sourceRange` 必须与 `sourceSheetId` 配对，并以范围尺寸校验 `rowCount` 与字段宽度，不是可忽略的提示字段。
2. **实际消费者**：OOXML native chart writer 对 `source.kind === 'table'` 通过 table id 读取 `WorkbookTableModel.sourceRange` 并据此生成分类/系列引用；表范围落后于行移动会导出错误行。
3. **TS 结构路径对照**：轴变更与 cell-shift 会预检并更新 workbook-table/data-source 范围；`validatePermutationMetadata` 和 `applyRowPermutation` 原先没有枚举这两类 owner。置换一段连续数据行时，表/源若只覆盖其中一部分，应跟随那部分原始记录；若置换后不能表达为单矩形，应拒绝。
4. **Java 结构路径对照**：`StructuralSnapshotReducer` 的轴与 cell-shift reducer 同样维护这两种全局范围；`validatePermutationMetadataExact` 与 `remapPermutationMetadata` 原先遗漏，导致 server 接受置换后仍保留旧地址。
5. **同一操作路径**：`data.sort.rows` 将已校验顺序写成 `rows.permuted`；前端 mutation handler 和服务端 `permuteRows` 都执行上述转换。因此遗漏同时影响本地首次执行与持久化/重放，不是未调用的辅助函数差异。
6. **触发与 fail-close**：置换范围精确映射器可将完整连续目标映射为一个范围，也会对拆分结果返回多个片段。两类 owner 原先既没有更新也没有 split 检查；新增成功路径校验范围与 cell 同步移动，拒绝路径校验在快照变化前 fail-close。

**修复**：TS 预检暂存相交 table/data-source 的精确新范围，并在所有单元格预计算后应用；未变化及非相交范围不重赋。Java 在写入前验证范围可表示为单矩形，再于 reducer 写回；读取可选 `dataModel.sources` 使用 non-mutating 的 `existingDataModelArray`，不物化缺省数组。新增 TS/Java 成功与 split-range 拒绝回归用例。仅静态源码审查与 `git diff --check`；没有运行本地测试、构建、lint 或 UI。远端 PR CI 尚待新 head 结果；Canonical Structural Planner/ReferenceIndex、其它 owner 家族、OOXML opaque-owner 结构迁移及总目标仍未完成。

### 六轮静态自审 — 复制/转置引用与 owner 完整性（2026-09-26）

六轮分别以坐标代数、公式 owner、结果缓存、稀疏元数据、矩阵变换边界、反向重放为视角，核验后只计独立根因，不把同一缺陷按字段重复计数：

1. **复制公式坐标**：`sheet.range.paste` 遍历稀疏单元格时，原先用目标 cell 减 source range 起点，而非减该公式自己的源地址；多行/多列及转置公式因此多移一次偏移。现改为逐 cell 的 destination-source 坐标差，并为混合绝对引用保留 AST 规则。
2. **公式结果缓存**：Paste 复制/公式/算术写入会继承目标或源的 `formulaValue`/`displayValue`；矩阵变换也会把旧位置缓存随公式一并写到新位置。现对内容写入和公式 owner 移动清除旧缓存，公式规则的纯引用改写则保留 owner 与 anchor。
3. **Paste Values 语义**：公式 cell 的 UI 结果由 `formulaValue` 承载，但旧 values 分支只复制 `value`，且 `null` 结果不能用 nullish fallback 判断为“无缓存”。现按 `undefined` 判断缓存缺失，写入标量常量并保留错误结果的公式值投影。
4. **规则公式 owner**：`SheetRuleRegistry.cloneForPaste` 只 remap rule ranges/anchor，遗漏 conditional-format `value1/value2`、validation `formula1/formula2` 与 `listSource.formula` 的相对引用。现从 `structuralRuleFormulaFields` 取得 canonical 字段，在源/目标 formula anchor 间用公式 AST 转译；无法解析时 fail-close。
5. **转置及 Skip Blanks 元数据**：notes/comments/hyperlinks 原先按未转置 offset 写回；copy + skipBlanks 又清除整块目标元数据并从空 cell 复制 metadata。现将三类 cell owner 映射到转置目标，并仅清理/复制非空来源对应的目标 metadata；空白目标项保留原 owner。现存 column-width 的转置映射按既有专项契约保持不变。
6. **矩阵变换所有权与失败边界**：相对公式引用若在选区外，旧逻辑不随 formula owner 位移；block-backed data region 未在 source/target 预检；hyperlink anchors 会留在旧坐标；review/drawing/sparkline 预检有的只看 row、有的全表拒绝；且只检查 target 上界，未验证 source/target 的安全整数与下界。现按 formula owner delta 平移选区外相对引用、移动公式 owner 时清缓存；source 与 transposed target 都拒绝相交 data region；无法迁移的 hyperlink owner 显式 `UNSUPPORTED_FEATURE` 拒绝；anchor guard 使用 row+column 精确范围，并在任何变换前验证 source/target 完整边界。

新增回归测试源码覆盖多格/转置公式、规则公式、结果缓存、Paste Values 的标量和显式空值、转置 metadata、Skip Blanks、矩阵相对引用、data-region/hyperlink fail-close、二维 anchor 与非法边界。静态复核追到 `range.paste` 单 mutation snapshot、undo inverse 与 remote replay 均消费同一快照；本轮未执行测试、构建、lint 或 UI，`git diff --check` 通过。六轮共确认并修复 12 个独立问题（将 paste/matrix 的缓存失效分别计数，规则公式字段按同一 owner 根因合并）；未凑数宣称达到此前提出的 30 项门槛。其余结构 owner 与完整 Canonical Structural Planner/ReferenceIndex 目标继续开放。

额外公式引用边界核对确认第 13 个独立问题：whole-row/whole-column AST 虽可解析，但 copy/fill `offsetAst` 不移动其相对轴端点，且 formatter 曾把 `A:A` / `1:1` 折叠成不再是范围的 `A` / `1`，混合 `$` 端点也未保留。现 AST 按端点保留绝对标记，解析器让数值行范围（如 `1:1`）进入引用分支并支持 `$1`，formatter 保留范围语法，copy/fill 与矩阵 owner 位移按相对轴偏移并对越界生成 `#REF!`。新增回归测试源码覆盖轴引用复制、混合绝对端点、越界和结构插入；仍仅静态审查，未运行测试、构建、lint 或 UI。

第七个补充交叉核对确认第 14 个独立问题：矩阵转置的目标区域可能超出源区域（如 2×1 转置为 1×2），预检只检查源区内的 review/drawing/sparkline owner，目标扩展区的锚点可能留在覆盖后的旧坐标。现这些 owner 按源区和目标区的并集作二维相交检查；新增每种 owner 各自 fail-close 且快照不变的回归测试源码。仍未运行测试、构建、lint 或 UI。

第八个补充交叉核对确认第 15 个独立问题：Paste Skip Blanks 把 `value:null` 且仅有 canonical `formulaValue`（如常量错误值）的单元格当作空白，既跳过值又漏掉其目标元数据清理。现共同的稀疏剪贴板空白判定同时检查 `value`、`formula` 与 `formulaValue`；新增错误常量、便笺和超链接一同粘贴的回归测试源码。仍仅静态审查，未运行测试、构建、lint 或 UI。

第九个补充交叉核对确认第 16 个独立问题：动态数组 spill child 不占据 `CellMatrix`，矩阵转置的目标空格检查无法发现该 owner，可能把写入落到溢出投影并破坏数组结果。现矩阵操作在变更前检查源区与目标区是否相交任一 spill range，并 fail-close；新增目标扩展区命中 spill projection 的原子拒绝回归测试源码。仍未运行测试、构建、lint 或 UI。

第十个补充交叉核对确认第 17 个独立问题：复制规则覆盖区的子区间时，原 formula anchor 可能位于剪贴板 source 之外；把该 anchor 的相对坐标直接映射到 target 会造成公式少偏移被裁剪距离（例如从 A5 复制规则到 C5 却仍引用 C1）。现将规则 anchor 落在实际复制交集的目标坐标，并从原公式 anchor 精确平移公式 AST；新增 CF/DV 子区间复制源码断言公式和 anchor。仍未运行测试、构建、lint 或 UI。

### 六轮静态复核 — 完整整轴引用与跨端公式契约（2026-09-26）

**边界与方案**：本轮限于 formula parser、Java formula reference transformer、既有规则复制类型边界及这些路径的回归源码。整行/整列引用必须作为含工作表限定符的完整 token 进入变换；共享点/区间向量扩展到完整公式，以同时约束 TS AST 和 Java 服务端实现。Java 删除“先改 A1、再扫整轴”的旧路径，在同一个扫描器中分派完整 A1/整轴引用；不增加兼容桥或第二套读写状态。事务、权限、快照版本与 StructuralPatch v3 字段不变，没有持久化迁移。

六个独立复核视角与证据：

1. **解析入口**：`parseReference`、3D 与外部引用三个入口都用 `check('colon')` 判断当前端点 token，而非下一 token，因此 `A:A`/`1:1` 的首端点未被允许进入 whole-axis 分支。上一节“whole-axis AST 虽可解析”的结论不准确；本轮追到真实入口后修正为 `checkNext`，并使用 canonical reference domain 限制整轴端点，避免把 `Revenue:Other` 或越界端点误当有效轴范围。
2. **失败与限定符分派**：标识符解析曾在引用语法已消费后吞掉异常并退回 name，`A:` 因而可能被截断为 `A`；现在仅未进入引用语法的单独标识符可作为 name。另确认字符串/布尔常量分派先于工作表限定符，导致 `'Budget A1':Other!A:A` 与 `TRUE!A:A` 无法进入正确分支；现在有效限定符先进入引用解析，普通字符串与 `TRUE` 常量保持原语义。这是两个根因。
3. **Java 完整 token 所有权**：旧 A1 扫描器不能消费 `'Budget A1'!B:B`，随后会扫描并改写工作表名内部的 `A1`。受影响入口包括轴变换与 row-permutation offset。新扫描器一次消费限定符和整轴 body；外部引用/字符串仍保持完整，rename/delete 也复用此入口。删除重复整轴扫描、未被生产调用的通用 `offset` 及其 forwarding overloads。
4. **输出坐标与绝对标记**：旧 Java 整轴删除/溢出只替换 body，可能留下 `'Budget A1'!#REF!`，与 TS 的完整 `#REF!` 不一致；现在替换完整 token。反向轴范围的输出又与 TS 正序 AST 不同；现在按端点一起移动 `$` 标记并统一顺序，canonicalization 使用相同 renderer。这是两个根因，不按行/列和不同限定符重复计数。
5. **服务器 patch、逆向与拒绝边界**：沿 `rows.inserted`、`rows.permuted` 的 descriptor/reducer 追到 formulaOwnerDelta、canonical inverse 比对与 replay。新增源码用例断言完整限定符、公式 owner 新地址、patch formula、置换回放恢复与输入快照不变；保留对 whole-row/external row-offset 的 fail-close，并补充带 cell-like 工作表名的拒绝用例，不扩大排序支持范围。
6. **验收源码与已有门禁反馈**：`reference-transform-vectors.json` 新增 10 个完整公式向量，TS 与 Java 都读取同一份输入和期望值，覆盖限定符、反向端点、删除/溢出、字符串、外部/3D 与布尔工作表名。另读取上一个 head `6c14437e` 已有的远端 CI（run `36198413697`）：唯一已报告的构建错误为 `rule-lifecycle.ts` 在 `SheetRule` union 上直接访问 `listSource`（TS2339）；现通过已有 conditional-format 类型判别先排除非 validation owner，再写回 formula list source。未执行或重跑任何本地测试/构建。

本轮确认并修复 **7 个独立问题**；坐标上界为解析入口修正后的附加拒绝约束，不另凑问题数。成功/拒绝回归测试源码已补充，但没有执行，不能把静态推导当成测试通过。验收仅包括六轮源码复核、共享 JSON 语法读取及 `git diff --check`；本地 tests/build/lint/typecheck/UI 均未运行。上一 head 的 CI 失败是已知事实，本轮 head 的自动 CI 结果须单独记录，不能沿用旧成功记录。

**剩余工作与回退**：这只是跨端公式语义收敛的一步；完整无副作用 `CanonicalStructuralPlanner`、typed metadata ReferenceIndex、全量可逆 StructuralPatch、OT/OOXML opaque-owner 参与仍未完成。PR 继续 draft，不宣称通过浏览器或原生 Excel 互操作验收。回退需整体 revert 本轮 production、共享向量和回归源码的提交；不需要数据降级迁移。

### 六轮静态自审 — 轴 metadata 从重复推导收敛到规划结果提交（2026-09-26）

**本轮基线**：重新读取 GitHub main 分支，HEAD 仍为 `a2a6140a90351b38f1e6f5fbc167d09b4f6ecc7f`；开发分支从 `21325eea58f83683e3c3be3a8895f6bbea6f1a81` 继续。该上一提交的自动 PR/push workflows `36209046085` / `36209043773` 均已成功，本轮未执行或重跑本地验证。

**有界方案**：整行/整列插删的 metadata 在单元格写入前只执行一次变换，把已校验的结果交给应用阶段。删除旧 `preflightAxisMetadata` 丢弃结果后在 live 模型上再次调用全部 shift helpers 的路径；内部 `AxisMetadataPlan` 仅由同一次同步调用使用，不暴露为 wire patch，也不冒称完整 Canonical Planner。单元格算法、公式计划、权限、mutation IDs、StructuralPatch v3 和持久化格式保持不变。

六个独立视角的源码证据与自审修正：

1. **真实入口及推导次数**：`rows/columns.inserted/deleted` 的 command apply 和 mutation replay 均进入 `StructuralTransform.apply → applyAxis`。原来 metadata 先在 detached owners 上变换，再于 cell shift 之后对 live owners 重算；现在 `planAxisMetadata` 生成结果，`applyAxisMetadataPlan` 只应用值，不调用任何坐标/引用变换。`shiftDataRegionAxis`、`shiftWorkbookTables` 不再提供默认 live workbook collections，避免重新引入隐式写入目标。
2. **参与者覆盖**：对照删除的 live 调用逐项核验，计划提交包含六类跨表 owner 集合（CF、DV、Pivot、Sparkline、drawing payload、hyperlink）和目标表 dataRegions、merge、Sheet Table、drawing anchors、spill、protection、filter、banded、outline、pane、隐藏行列、尺寸、notes/threads；workbook tables/sources 和 PrintDocument 同批提交。ReportSheet 与 formula owners 保持原有独立预计算结果，不由 metadata staging 覆盖。
3. **失败原子性**：所有 shift、范围有效性判断及 metadata 差异比较都发生在 cell shift 前。自审中将最初位于 apply 阶段的比较移回规划阶段，避免序列化失败出现在部分写入之后；不可表示为 snapshot data 的值保留 `STRUCTURAL_PATCH_INVARIANT` 错误。新增源码用例在最后的 print page-break 映射溢出时验证 cells/隐藏行/notes/print state 全部不变，再改为有效 page break 验证成功提交。
4. **canonical 索引同步**：`DataRegionBoundsIndex.add` 保存的是 `region.range` 的 clone，而旧轴路径直接修改 live `region.range`，因此 `sheet.usedRange` 会保留旧边界。新提交通过 `WorksheetModel.replaceDataRegions` 同步写入范围与索引；回归源码同时覆盖 row/column insert-delete 往返，不能仅以 snapshot 范围正确作为索引正确的证据。
5. **跨端坐标拒绝与对象身份**：旧 TS data-region range/header 使用裸 `+= count`，当 chunked-table projection 位于 Excel 边缘且没有 worksheet source range 代为检查时可能越界；Java `shiftDataRegions` 已使用 `requireShiftedRange` 和 `shiftIndex`。TS 现也走 `ReferenceTransformDomain`，在计划阶段拒绝溢出。changed-field 集合仅提交变化字段；数组/Map 集合实例及仍位于同一槽位/键下的未变化 owner 保留身份，避免无关 owner 被 cloned staging 替换。新增跨表 Sparkline/hyperlink 身份与冻结窗格/尺寸/notes 一次移动、整体逆向恢复源码断言。
6. **history/replay/持久化与验收边界**：mutation replay 复用同一轴入口，formula-owner 应用仍按 owner ID 解析已提交 metadata，不持有旧 staging 引用。现有 mutation/result/protocol 不变；本轮并没有把 metadata 计划写进服务器 operation log 或 history，因此 metadata 的完整事实逆向回放仍需 owner-complete patch 迁移。只做静态 review，新增测试源码未运行；不会把上一个 head 的 CI 通过当成本轮验收。

本轮确认 **3 个独立问题**：轴 metadata 重复推导、data-region bounds index 过期、TS data-region 坐标越界缺少领域检查。验证范围为六轮源码复核及 `git diff --check`，不运行 tests/build/lint/typecheck/browser。没有 schema/data migration；回退为整体 revert 本轮实现和回归源码提交。

**明确未完成**：内部计划仍暂存 detached worksheet metadata，并枚举既有跨表 owner families；差异比较也有线性开销。虽然删去第二次结构推导，但尚未建立 affected-owner metadata index，也没有证明达到 `O(affected references + affected objects + moved cells)`，本轮没有 benchmark。下一步仍须用 typed owner delta 替代 coarse staging，并将 cells、formula、metadata、calculation/projection、history、Java authority 和 OOXML 纳入完整可逆 StructuralPatch，移除余下独立 reducers。本步骤不缩减最终目标，PR 继续 draft。

### 六轮静态自审 — 三类范围 owner 改为精确几何事实（2026-09-26）

**基线及真实失败**：main 仍为 `a2a6140a90351b38f1e6f5fbc167d09b4f6ecc7f`，从开发 head `a7af29ae11c47365a9b789c16b1b3c9484274481` 继续。上一 head 的自动 workflows `36210201193` / `36210198645` 已失败；PR job `108314937919` 报 `structural-transform.ts(891,62)` TS2322：`noUncheckedIndexedAccess` 下数组索引值可能为 `undefined`。现将索引值保存在局部变量并在赋回前明确收窄。未执行本地编译，也不把源码修正写成 CI 已通过。

**有界设计与实施**：轴操作仅移动范围时，上一版对 workbook table/data-source 连同 fields、blocks 等内容整体 clone、比较并替换 owner，付出了与范围变换无关的复制成本。此次只迁移 data-region range/header、workbook-table sourceRange、data-source sourceRange 三类无公式几何 owner：先生成冻结、可序列化的 `StructuralRangeOwnerDelta` before/after 事实，内部 apply 只提交这些值；删除这三类 owner 的 axis staging、`shiftDataRegionAxis`、`shiftWorkbookTables` 旧路径。其它 metadata 暂不迁移，避免整块 CF/DV/chart 快照覆盖独立公式计划。权限、事务入口、持久化 schema、StructuralPatch v3 与 Java wire shape 不变。

六轮有不同检查对象，回看同一根因不重复计数：

1. **入口与旧路径删除**：`applyAxis` 在 cells 写入前规划三类 owner；`applyAxisMetadataPlan` 只消费范围事实。检查 `cloneStructuralPreflightSheets` 两个调用者，cell-shift 已在 live 模型预检阶段拒绝与 data region 相交，`shiftCellBandMetadata` 不读取 dataRegions，因此删除公共 staging 中的 dataRegions 复制不会使 cell-shift 失去校验。
2. **owner 与载荷边界**：table/source 的 fields、blocks、rowOrder、revision 都不是此次几何变换所有者，不能随坐标变化替换；新事实只包含稳定 owner ID、range 和必要 headerRow。table/source 本体与其字段、数据块数组保持身份，未变化范围不发 delta，也不替换范围对象。仍枚举现有 owner collections，不宣称 affected-only 复杂度。
3. **读取与写入契约**：复核实际 getter 后发现 `getDataSource()` 返回 clone，而 `getTable()` 返回 live owner。本轮未提交草稿最初曾错误地向 data-source getter 的副本赋值，已在静态自审中撤掉；apply 现在按规划身份访问 canonical sources Map 并仅更新 sourceRange。回归源码分别锁定真实 owner 移动与公共读取快照仍然独立，不更改 getter 契约。
4. **事实独立性与索引**：事实数组、条目、before/after 及 region 的嵌套 range 都冻结且不引用 live model。apply 给 live range 写入新值，dataRegions 仍通过 `replaceDataRegions` 同步 `DataRegionBoundsIndex`。回归源码覆盖 JSON 往返、冻结层级、后续 row/column 往返不能修改旧事实以及 usedRange 变化；不声称 region 整体替换已变成增量索引更新。
5. **拒绝与写入时序**：范围/头行映射溢出、region 重复/空身份或错误 sheet、table/source Map key 与 owner ID 不一致均在 cells 写入前拒绝。新增三类 owner 的损坏身份源码用例，断言完整 snapshot 与 usedRange 不变；保留已有 block-backed 边缘溢出、跨 region 编辑拒绝及末端 print-plan 失败用例。mapInterval/shiftIndex 的范围语义与原路径一致，不在 apply 中重算。
6. **消费者、回放与类型边界**：`StructuralTransformResult.rangeOwnerDeltas` 只由当前轴操作发出，用于暴露本轮已迁移事实；CommandRuntime、服务器 operation log 与 StructuralPatch v3 仍只处理已有公式/name facts，尚未消费这些范围事实。不能用“先反向执行旧算法、再覆盖整块 metadata”冒充可逆 patch。本轮不更改 undo、OT、Java、OOXML 协议，也不宣称其已经统一。静态检查可选结果字段的断言、判别联合及 `noUncheckedIndexedAccess` 收窄，未运行 typecheck。

**验证与问题计数**：本轮修复两个已确认问题（上一提交的编译阻塞、范围变换不必要的整 owner 复制/替换）；三类 owner 是同一迁移的覆盖面，不拆成三个独立根因，也不把自审拦下的未提交回归计为额外完成项。新增/补充成功、未变化、拒绝、身份与事实不可变性的回归源码；按用户要求只做静态审查及 diff whitespace 检查，不执行 tests/build/lint/typecheck/browser。自动 CI 仅记录相应 head 的实际状态，未主动重跑。

**剩余工作与回退**：这仍不是完整 Canonical Structural Planner、metadata ReferenceIndex 或可逆跨端 StructuralPatch。其它 detached metadata staging、cells/公式与 metadata 的计划聚合、history/Java authority/OT/OOXML 的事实消费仍待迁移。没有 schema/data migration；回退须整体 revert 本轮类型、实现与回归源码，不能只删结果字段却保留半套 apply。PR 保持 draft，浏览器、原生 Excel 文件与性能验收仍未完成。

### 基础操作性能优先 — 单元格插删 metadata 计划（2026-09-26）

用户进一步要求优先基础操作、大数据量、内存和计算时间，并参考微软官方设计。本轮先处理 Insert/Delete Cells：当前 `preflightCellShiftMetadata` 在 detached owners 上变换后丢弃结果，`applyCellShift` 再遍历 live metadata 重算；report bindings 同样映射两遍。预检还深拷贝所有同表 workbook tables/data sources，但 `planCellShift → validateDataRegionCellShift` 已拒绝它们与 affected band 相交，合法情况下这些对象不可能移动。

上述最初方案在边界自审中发现一个前提缺口：原校验的 affected band 只覆盖当前 rowCount/columnCount，block-backed owner 可以合法地超出这个已物化范围，而引用位移覆盖 Excel 坐标域的整个尾部。因此先同步修正 TS/Java 的 owner guard，让它检查完整语义尾部；只有通过这个校验，删除 table/source 的无效范围变换才成立。不是简单假设“屏幕外/未加载对象不会受影响”。

**实施边界**：让轴操作和单元格插删共用一个 metadata 计划提交器；cell-shift 保留完整预检结果，report binding 进入同一计划，删除第二次 live 变换和无变化 table/source 的整对象复制/范围变换。单元格移动、公式引用计算、权限和 wire/history schema 本轮不改。任何 planning 失败必须在 cells/metadata 写入前返回；成功提交保持未变化集合/owner 身份。验收源码覆盖行/列插删、跨表引用、report/review/print、无关数据 owner 和失败时 snapshot 不变；仍按本轮静态审查要求不执行本地 tests/build/browser。

**微软依据与推导边界**：微软 [Excel calculation performance](https://learn.microsoft.com/en-us/office/vba/excel/concepts/excel-performance/excel-improving-calculation-performance) 说明依赖跟踪、smart recalculation 和重用计算顺序；[performance and limit improvements](https://learn.microsoft.com/en-us/office/vba/excel/concepts/excel-performance/excel-performance-and-limit-improvements) 记录了整列引用、多工作表，以及过滤/排序/复制粘贴在 CPU、内存和响应时间方面的改进。我们据此把受影响对象工作量、重复计算和分配量作为基础操作优化重点；这不是声称微软内部采用本仓库的 metadata plan。单元格位移与整行/整列操作的区别按 [Microsoft insert/delete guidance](https://support.microsoft.com/en-us/excel/get-started/insert-or-delete-rows-and-columns-in-excel) 保留。尚未实测延迟、峰值内存或百万行承载能力，不从源码调用次数推算加速倍数。

六轮静态自审与本轮结果：

1. **真实基础操作入口**：`sheet-features/src/editing/index.ts` 的 `sheet.cells.insert/delete`、对应 `cells.inserted/deleted` replay 与 restore 都进入 `StructuralTransform.apply → applyCellShift`。现在 `planCellShiftMetadata` 只调用一次 metadata 变换，并保留 report/print 结果；删除 cells 移动后的第二次 `shiftCellBandMetadata` 和第二次 report 映射。轴操作与 cell-shift 共用 `collectStructuralMetadataPlan → applyStructuralMetadataPlan`，未引入另一套 mutation 入口。
2. **工作量与内存存活**：删除 cell-shift 对 workbook table/source（含 fields、blocks、rowOrder）的整对象克隆；应用阶段也不再访问这些未变化载荷。metadata 计划改存 typed changed-field values，不再持有完整 detached WorksheetModel；无变化工作表不进入提交列表。数组/Map/Set/尺寸容器实例与未变化 owner 保留身份，空集合仍会正确清空 live 内容。规划期间仍有 metadata 克隆和差异比较，不声称峰值分配已经消除。
3. **参与者与公式顺序**：对照旧 live helper，范围/anchor、CF/DV、Chart、Pivot、Sparkline、drawing、filter/sort、spill、protection、banded、review、hyperlink、PrintDocument 均由一次计划覆盖；ReportSheet 加入 typed local fields。公式计划仍按稳定 owner identity 在已提交 geometry 上写公式，不能被 staged metadata 覆盖。回归源码覆盖两轴 insert/delete 往返、验证公式/anchor、柱状图源、跨表 Sparkline/hyperlink、report/note，以及无关数据源与集合身份。
4. **canonical 身份与重复查表**：`shiftCellRangeReference` 之前把 RangeRef.sheetId 放进公式 qualifier，再构造完整 sheetOrder。公式解析按 display name 优先查找，另一张表的名称等于目标 ID 时会误判范围 owner，导致漏移。生成的局部 AST 现在使用已解析 owner 上下文，不再按名字解析，也不再为每个范围复制 sheet identities。超链接 address 是真实公式文本，仍保留名称语义，但身份列表移到操作级只构造一次。按源码调用数，该 metadata 子路径不再有“每个范围构造 S 个工作表身份 + 每张 owner sheet 再构造 S 个身份”的分配；不据此宣称整个结构操作已达 affected-only。
5. **懒加载范围与失败原子性**：TS/Java 的 block region、Sheet Table、workbook table、data-source guard 都改查完整结构尾部，Java 将 guard 移到 cells 写入前；删除 Java metadata 阶段重复的 guard 与无效 table/source 移位。未物化范围与选区语义相交时明确要求专用 block/table 事务，不静默忽略或加载数据块。TS/Java 新增两轴、三种 owner 的尾部拒绝与输入不变源码；Java 同时覆盖移到不相关范围后可成功且数据 owner 内容不变。PrintDocument 非连续范围和末端 ReportSheet 拒绝源码确认已规划的 notes/cells 不先写入。
6. **跨端与回放边界**：Java `mapCellShiftRange` 原本就按 canonical sheetId 比较后处理坐标，TS 身份修复恢复这一约束。当前 StructuralPatch v3/history 仍未存放完整 metadata/cell facts，不能把两个 reducer 的同向结果当成最终统一语义。未更改 protocol 或持久化版本；只删除已迁移的重复路径，不新增回放字段去掩盖剩余独立推导。新 generic typed-field 收集、可选字段存在性、空集合及 formula-owner 写入目标均完成静态复核；源码用例未运行。

本轮确认 **5 个独立问题**：重复 metadata/report 变换、无关数据载荷复制及暂存保留、canonical 范围身份误解析、重复 sheet identity 列表构造、已物化 extent 导致尾部 owner guard 漏查。同根因的行/列或不同对象类型不重复计数。验证仅为六轮源码审查及 diff whitespace 检查；没有本地 tests/build/lint/typecheck/browser，也没有速度、峰值内存或承载规模的实测结论。前一 head `53891062` 的自动 workflows `36211353463` / `36211351079` 均成功，不代表当前修改已通过 CI。

剩余 priority 是基础操作的 typed affected-owner index、稀疏 cell/metadata planning、精确可逆 history、Java authority/OT/OOXML 消费同一完整 patch，以及 100k formulas/500k–1m occupied cells 的实测。没有 schema/data migration；回退应整体 revert 本轮 TS/Java 实现和回归源码，不能只恢复一个端的尾部边界。PR 继续 draft，完整目标不变。

### 基础操作继续 — 预规划公式写入与多 Sheet 稀疏读取（2026-09-26）

**基线与真实 CI 证据**：从 `6cda327dc8e7ddb52b2e12c7a3a398680d480f1d` 继续。其 PR/push workflows `36212947081` / `36212944892` 均失败；PR job `108323128071` 的代表性错误是 `MutationDescriptorRegistryTest.cellShiftRejectsDataOwnersBeyondExtentAndPreservesUnrelatedSources:2072`、`review must be an object`。新增测试输入缺少 canonical review 对象，成功路径才到达该校验。本轮补齐 notes/thread 四个索引的空对象，不放宽生产校验。没有把上轮的“CI 进行中”沿用成已通过。

**有界实施方案与影响**：当前 command history 仍先执行 inverse reducer，再覆盖公式/name facts；只把范围字段加入协议不能消除反向推导。因此先收敛轴编辑与 cell-shift 的公式计划：在首个 cell write 前计算目标地址、最终 CellData 和 before/after formula facts；CF/DV 公式写入 detached metadata，与几何变化一起收集和提交。range.move 的外部公式 owner 复用同一个准备/提交器，移除重复的 live-cell 组装路径。公式提交器不再接收轴操作或 cell-shift 参数，也不再解释地址变化。没有新增结构命令、兼容 reader、wire 字段、权限入口或持久化迁移。

追踪实际调用方确认：仅替换模型中的 `cells.get/set` 不能维持惰性，因为 `synchronizeStructuralMutation`/auxiliary reindex 会再次强制读取整表，formula bootstrap/首个公式又使用 `cells.forEach` 物化全部 worksheet。此次同步引入 CellMatrix 所有的 `forEachWithoutHydration` 稀疏读取契约，并让两个已有公式枚举入口共用其存储遍历；计算输入加载、首次公式的完整值输入同步、跨表公式更新和 spill occupancy 都读取同一 canonical sparse storage。不复制第二份 workbook，不跳过普通值，也不把未加载单元格视为空。

六轮静态自审：

1. **入口及坐标所有权**：核对 `StructuralTransform.apply → applyAxis/applyCellShift` 与 range.move 外部引用入口。前两者的 formula-cell 目标地址和 before/after 状态在预检生成，删除后的 owner 进入 removedCells 而不是幸存者 delta。`applyFormulaRewritePlan` 不再接收 shift、targetSheetId 或 cellShiftPlan，`applyPlannedFormulaCellChanges` 只检查原 formula state 并写规划值；range.move 复用这条单元格写入链。source/destination 内被移动的实体仍由原 move planner 负责，不借本轮宣称整个 move/history 已迁移。
2. **规则公式与几何合并**：CF/DV after-formula 在 `stageMetadataFormulaRules` 写入 detached rule；其 afterRanges 也在 cell write 前复制到 delta。之后才进行 changed-field 收集，因此公式单独变化、公式和范围同时变化都进入唯一 metadata 提交。删除轴/cell-shift apply 阶段的规则查找、字段写入和从 live model 生成 afterRanges；未变化 owner 身份仍由公共收集器保留。回归源码覆盖两轴插删/单元格位移的 CF/DV 公式与范围、后续修改不污染 history range facts。
3. **多 Sheet 数据载荷**：公式索引命中的 owner 通过 `getFormulaOwnerWithoutHydration` 读取；`replaceFormulaOwnerWithoutHydration` 使用现有 copy-on-write sparse storage。初始加载、first-formula transition 的 ordinary-value 输入枚举改为无 hydration，但仍喂给 engine 所有普通值，包括远端稀疏尾部和 preserved-only 缓存。公式同步复用已读取 CellData 建 auxiliary index，删除第二次读取。没有仅对测试绕过 runtime 的专门入口。
4. **规范化与失败原子性**：无 hydration 读取不负责修改字体；准备需要写入的 cell 时提前执行既有 `normalizeFontFamily`，并检查 provenance/barcode owner 可写性。控制字符字体在任意 cells/metadata 写入前拒绝，源码断言完整 snapshot 和 deferred 状态均不变。新的稀疏枚举不修改 revision、输入 JSON 或 storage ownership，callback 抛错向上传播；没有 catch/default 吞错。既有 `forEach` 的 materializing 契约保持不变。
5. **计算、历史及 spill 链**：实际 runtime 的 rewritten-owner 查询改用稀疏读，bootstrap/首个公式路径保留完整值输入；spill occupancy 每次读取当前 sparse cell，清除 blocker 后会释放，公式单元格仍阻挡。新增 WorkbookSession 源码用例覆盖 snapshot load、insert row、undo、redo、引擎实例不被替换、远端公式值及 inactive sheet 状态；另覆盖首次公式使用本表/远端/极端坐标普通值以及实时 blocker 改变。未执行这些用例，不用它们的存在冒充行为通过。
6. **跨端、类型与剩余工作量**：同一 formula owner delta shape 仍供 CommandRuntime 和 StructuralPatch v3 使用，没有 runtime legacy reader 或字段补丁。TS 类型、所有改动调用点、可选字段、被删除 owner、稀疏与 hydrated 两种存储读取逐项静态复核。Java 本轮只有测试 fixture 修正，仍是独立 structural reducer；history 精确 cell/metadata facts、server authority、OT 和 OOXML 尚未收敛。不能把这批内部预规划称为完整 CanonicalStructuralPlanner。

**本轮结果与验证边界**：收敛两个已证实的架构缺口（写入后才决定公式 owner 目标/最终状态、规则公式与 metadata 结果分离），修复一组跨表强制物化路径和一个 CI 测试输入错误。不把多个调用点或行列变体凑成 30 个独立 bug。新增/补充成功与拒绝回归源码，按用户要求只做静态审查与 diff whitespace 检查，不执行本地 tests/build/lint/typecheck/browser。测试中的 600001 行/XFD 是稀疏边界坐标，只有少量 occupied cells，绝不是大数据量 benchmark。

**性能与回退**：已删除公式 owner 更新引发的整表 Map 物化及重复 live 写入组装，但首次计算仍要枚举值、FormulaEngine 仍持有普通输入，deferred 写入仍有目录/行的 copy-on-write 成本，metadata 仍有 cloning/owner-family 枚举。没有毫秒、heap、加速倍数或百万 occupied cells 验收结论。无 schema/data migration；回退需整体 revert 模型公式计划、稀疏枚举及 runtime/spill 消费者和回归源码。完整目标保持 active，PR 继续 draft。

### 基础操作性能继续 — StructuralPatch owner 合并（2026-09-26）

本轮沿 undo/redo 服务链静态跟进 `WorkbookOperationService.commitInternal → MutationDescriptorRegistry.mergeStructuralPatches`。当 reducer 与 undo inverse 都有 owner facts 时，旧实现对每个 inverse formula/name owner 用 stream 从头扫描已生成列表；两边各有 `F`、`I` 个 owner 时，身份比较上界为 `O(F×I)`。Formula-owner deltas 可能随大范围公式引用变换增长，因此这是可由代码路径直接确认的 CPU 放大点，不是基准测试推断。

将 owner 合并改为类型化身份 key 的索引查找，期望比较成本为 `O(F+I)`；临时索引建在两侧较小的列表上，为 `O(min(F,I))`，任一侧为空则不建索引。输出仍按原生成 delta 顺序，再接 inverse-only owners 的原始顺序；同一 owner 的完全相同事实去重，冲突仍以 `CONFLICT` fail-close。formula-cell 按 after address、formula-rule 按 sheet/kind/id/field、五类 formula-object 按各自稳定 ID 组合、defined-name 按 scope + `Locale.ROOT` 大小写折叠 + sheet ID 建 key。随后一次架构自审发现 identity 曾在 patch 构造验证和 merge 两处重复定义；现在由 `StructuralPatch.formulaOwnerKey/definedNameOwnerKey` 单点产生，重复校验与合并共同消费，新增 owner 不再需要维护第二套身份 comparator。仅 Java 内部 API 增加 identity helper 类型/方法，不改 wire/schema/history。

六轮静态自审：

1. **根因与调用量**：检查 commit、committed replay、migration replay 三个调用点；只有 generated 与 inverse 同时非空才执行交叉匹配，线性扫描被替换为散列表查找。
2. **公式 cell 身份**：核对旧 comparator 与 `StructuralPatch` uniqueness key 都以 after address 标识同一目标 owner；保留同地址不同事实的冲突拒绝。
3. **规则与对象身份**：逐项比对 formula-rule 及 chart-text、shape-property、table-sheet-column、data-view-field、cell-style-template key 字段；身份只在 `StructuralPatch.formulaOwnerKey` 定义，patch 构造器去重和 registry 合并共同调用；新增回归源码覆盖全部类别和同图表不同字段。
4. **defined-name 身份**：scope 与 sheet scope 必须区分；名称仅允许 ASCII 标识符，`Locale.ROOT` 大小写归一与原 `equalsIgnoreCase` 语义一致；同名不同大小写但事实不一致仍拒绝。
5. **输出契约**：较小侧建索引时仍先校验所有交集，再按 inverse 输入顺序追加未匹配项；不从 HashMap 迭代生成 wire 列表，避免 nondeterministic ordering。
6. **分配与边界**：空侧走无索引路径，非空时索引空间随较小输入线性增长；成功合并、公式冲突、大小写名称冲突及 StructuralPatch wire field 集合不变的测试源码已新增。未运行测试/build，仅执行静态审查和 `git diff --check`，不能据此宣称 wall-clock 加速。

本轮只确认并修复这一项二次复杂度根因，不把 owner 类型、key 字段或测试断言拆成多个 bug。大型 structural patch 的实测 CPU/heap 仍待最终验收；完整 Java planner、客户端 intent-first、精确历史 patch、协作与 OOXML 消费仍未完成。

### 基础操作性能继续 — commit snapshot 所有权

静态追踪 `WorkbookOperationService.commitInternal → MutationDescriptorRegistry → StructuralMutationDescriptor` 确认：`currentSnapshot(row)` 已从持久化 JSON 解析出本次事务独占的 canonical tree，但每个结构 mutation 仍调用默认 `applyWithPatch`，再对整个 snapshot `deepCopy()`。一个 envelope 含 Q 个结构 mutation、snapshot 大小为 S 时，这条路径额外复制约 Q 份完整快照（O(Q×S) 字节复制），与真正修改的 cells/owner 数无关。

现在 commit 对无保护/OWNER descriptor，以及编辑者下明确列入规则型 allowlist 的 `OwnedSnapshotMutationDescriptor`，直接规约事务独占的 candidate：行/列插删、行置换的编辑者路径只复制各 sheet 的维度与 protectionRules；数据块发布 guard 只复制每个 source 的 block-reference manifest；普通 `edit-cell` 保护仍保留独立完整 snapshot，因为它需要旧 cell 的显式 unlocked style。其余/未来未列入的受保护 descriptor 默认继续使用 detached reducer。应用后强制检查返回 root identity。数据块引用变化检测、公式 owner 的二次保护检查、undo inverse patch、冲突范围和提交序列仍处于同一调用顺序；失败发生在持久化写入前，事务 candidate 不会变成可见状态。

同一边界检查还发现 `ProtectionResolver.assertAllowed` 先对空 affected-ranges 返回、后验证 action；缺失或未知 action 会绕过契约校验。现在先验证 action，再允许空范围结束，并增加拒绝用例；这不是把空范围改成全表授权。

六个独立静态复审角度：

1. **所有权来源**：确认 candidate 是 `currentSnapshot` 新解析的 JSON tree，不是共享缓存或 `WorkbookRow` 中可变对象；失败不会污染已存 snapshot。
2. **descriptor 路由**：仅实现 owned capability 的 reducer可原地运行；显式 protection-action allowlist 控制编辑者路径，cell-edit 和未分类动作保留 detached copy。
3. **保护前像**：post-reducer owner ranges 仍按 mutation 前的 dimensions/rules 检查；显式 unlocked cell 计数不被裁剪前像替代。
4. **数据块引用**：在原地规约前复制 manifests，因此候选树被改写时仍可逐 source 检出新增/变更引用；未改变引用时不访问 block store。
5. **事务与 undo**：inverse StructuralPatch 仍在 candidate 上应用，提交前校验及版本冲突失败均早于持久化写入；未改变 operation wire/history schema。
6. **复杂度与未覆盖调用方**：完整 snapshot 的读取/JSON parse 仍是 O(S)，保护规则及 block manifests 仍按其各自规模复制；优化仅去掉每 mutation 的整快照副本，不声称已实现 affected-only planner 或实测提速。公共 replay/migration 路径继续保持其原有 detach/ownership 规则。

本轮确认并修复 **2 个独立问题**：每个结构 mutation 重复复制完整 candidate；空范围会跳过保护 action 契约验证。不把行/列、各 descriptor 或断言拆分凑数。新增源码回归覆盖 owned root 身份、cell-protection detached 选择、规则 preimage、无变化/变化数据引用，以及空范围非法 action 拒绝。按本阶段静态优先要求，未运行本地 tests/build；推送后的 CI 另行记录。没有性能基准、浏览器交互或桌面 Excel 互操作证据，这些仍属于最终验收。无数据库/schema/protocol migration；回退为整体 revert 本轮实现、回归源码与本节记录。

### 基础操作协同链 — 权威 StructuralPatch ACK 与待提交队列（2026-09-26）

**基线与范围**：从 `1ed7f6e2` 继续，分支仍基于 `main` 的 `a2a6140a`，目标是现有 draft PR #345。本轮只审查已经进入服务端提交结果后的客户端确认、恢复查询和 outbox，不把它误称为 intent-first planner 或完整 owner patch。

六个独立静态复核视角与结论：

1. **REST 正常提交闭环**：协议返回 `CommittedOperationEnvelope`，内含 server-derived `structuralPatch`；runtime 原先只检查 request identity、确认 recovery journal 并 ACK，未消费 patch。相同 operation 的 websocket 通知又因 `ownOperationIds` 被跳过，形成权威 formula/name owner facts 永远不进入本地 runtime 的闭环缺口。现在先验证并应用服务端 owner facts，再确认 recovery journal 与 outbox ACK。
2. **恢复结果查询**：`getOperationResult` 恢复路径同样 ACK 并移除 pending 项，却不消费响应 patch；现与正常 REST 提交共享同一个权威 patch 入口，并在确认后推进 `remoteRevision`。
3. **ACK 失败顺序**：websocket 收到尚未确认的本地 operation 时，旧顺序先移除 outbox、标记 committed，再应用结构 patch；patch precondition 失败会使已提交事实丢失重试入口。现在先应用/校验 patch，成功后才将 operation 变成 terminal ACK。新增成功与拒绝路径源码用例检查 formula owner、pending 队列和 revision。
4. **本地 echo 生命周期**：runtime 只在 transport 失败时删除 `ownOperationIds`；成功提交后永久保留，集合随长期编辑增长。正常及恢复结果成功处理后均删除对应 ID，revision 先推进，因此迟到的同 revision 广播仍会被 revision gate 安全忽略。
5. **稀疏 patch 与模型分配**：`CommandRuntime.applyCommittedStructuralPatches` 原来把已定义但为空的 delta 数组视为有 patch，触发整本 `snapshot → fromSnapshot`。现在只对非空公式/name delta 预检；新增源码用例将任何意外 snapshot 变成显式失败。非空 patch 仍需完整工作簿副本以保障批量 precondition 原子性，尚未解决。
6. **outbox 查询与批处理**：协同点查询原先调用 `getPending()` 深拷贝整条队列；历史逐条匹配会重复拷贝，hydration API 还二次 clone 同一 envelope；批量 ACK/丢弃逐项持久化会反复重写剩余队列。现提供目标 operation/ID 读取、history 单次 ID 集合匹配、ACK/丢弃批处理单次持久化，并移除 hydration 的重复 clone。`rebaseQueuedOperations` 仍需一次完整快照，因为它会改写所有可变 pending intents。

此外，`applyRemote` 已先完整校验 committed envelope，再走不重复校验的内部 patch 应用路径，避免大型 patch 被全量做两遍协议验证。以上按独立根因计 **6 项修复**；不把正常/恢复两个入口、公式/name 两类 delta 或各个队列查询逐一拆数凑成 30。完整 intent-first planner、非空 ACK patch 的 affected-owner-only 预检、owner-complete history/OT/Java/OOXML 仍未完成，本目标继续 active，PR 继续 draft。

新增 ACK success/rejection 与 queue isolation/batch persistence 测试源码；按当前静态阶段要求，未运行本地 tests/build/lint/typecheck/browser。仅允许 `git diff --check` 静态检查；无 schema 或数据迁移。最终仍需 PR CI、真实浏览器协同、原生 Excel 文件往返及 CPU/heap 基准，不能用源码复杂度推导替代实测。

### Follow-up — 非空 ACK owner patch 稀疏预检（2026-09-26）

继续上一节未解决的性能项：非空 server-owned patch 原先为了批量失败原子性复制整个 workbook，再在副本上应用 owner deltas；其额外时间和峰值内存随工作簿总模型规模增长，而不是仅随 patch 规模增长。现改为公式 cell/rule/object 与 defined-name 的稀疏状态 overlay，按 owner identity 顺序模拟 before/after；不再调用 `snapshot()` / `fromSnapshot()`。预检和实际 formula-cell 写入共用同一状态转换，CellMatrix 与预检也共用 canonical cell-storage normalizer；defined name 在写入前使用同一规范化校验。规则 owner 的定位暂仍按所属规则数组扫描，因此这项变更只去掉 workbook 级复制，不宣称每条 lookup 都是 O(1) 或整体已达 affected-only。

六轮静态复核：1) 预检不再随无关 sheets/cells 的 workbook snapshot 扩张；2) overlay 顺序与实际 item/delta 顺序一致，重复 owner 使用同一 staged state；3) formula、provenance、barcode、formulaValue 清理及字体规范化与真实 CellMatrix 写入共用准备路径；4) rule owner 仍要求唯一 identity 与精确 after-ranges；5) formula-object precondition 必须先读到 before/after 中一个有效字符串，避免写入时 owner 消失；6) defined-name anchor/identity 经规范化后再进入 overlay，且拒绝路径在 live 写入前完成。确认并修复 **1 个独立根因**：非空 ACK patch 的全 workbook 预检副本。没有把 cell/rule/object 或多入口拆成问题数；30 个问题的规模目标不据此虚报。

新增源码回归覆盖：成功的多 owner patch 不得调用 `snapshot()`；第二个 owner precondition 失败时首个 owner 不得变化；owner 状态合法但 cell storage normalization 失败时也不得部分写入。按静态优先指令未执行本地测试/build/lint/typecheck/browser；`git diff --check` 通过。尚无该新 head 的 CI 结果、浏览器协同实测、原生 Excel 往返或 CPU/heap benchmark。无 schema/data/protocol migration；失败仍 fail-close，回滚方式为整体 revert 本 follow-up。

**CI follow-up (`5fe481de`)**：两条 `canonical-build` 均因 `formula-rule` 预检状态将 Conditional Formatting 的 `value1` 误窄为 `string | undefined`，而模型契约允许 `number`，报同一 TS2322。状态类型已保留 `number`；比较仍针对 `beforeFormula` / `afterFormula` 字符串，因此数值 owner 继续 fail-close，与原应用路径相同。此纯类型修正尚待新 head 的远端 CI 确认；未在本地运行 typecheck。

### Follow-up — Java range-owner writer、undo 与 v4 历史迁移（2026-09-26）

**边界**：把 `data-region`、`workbook-table`、`data-source` 几何事实从 core/command-runtime 的本地 history 延伸到 Java 权威 reducer、exact StructuralPatch v4、协作 ACK/remote consumer 与 operation-log/outbox replay。Runtime 仅接受 v4；Flyway repeatable migration 在明示迁移边界内验证并升级 v1–v3，不加 runtime fallback。

六轮静态复核：

1. **轴插删生产端**：`applyAxis` 已在变换前后采集 owner，但构造补丁时丢弃了结果。现在 v4 patch 返回 range deltas，覆盖 before/after impact 与 owner protection preconditions。
2. **移动与排序生产端**：`range.move` 会搬移 region/table/source ranges；`rows.permuted` 会重映射 workbook table/source ranges；旧返回值只含 formula owners。现在两条路径都采集并发出 before/after facts。排序还漏掉了 `dataRegions` 的实际几何与 headerRow 更新；现增加单值范围预检、header 归属校验和一致的 remap。
3. **服务端消费与拒绝原子性**：按稳定 owner identity 建目标索引，先预检所有当前状态，再批量更新目标 owner；缺失、重复、漂移或非预期 before/after 均 fail-close。历史 patch inverse 可恢复同一 owner facts；测试源码覆盖 axis、move、permutation 往返。
4. **wire 与跨端 impact**：Java、protocol validator、ACK/remote `MutationInfo` 统一要求精确 v4 字段；formula/name 与三类 range owners 的影响范围按同一顺序去重。旧 v3 客户端输入被拒绝，不让双方以缺字段的空 range delta 假装一致。
5. **持久化迁移**：迁移 checksum 前进到 v4；v1–v3 的旧 impact 列表仍按其 formula-only 语义验证，再将日志改写为 v4 impact；v2 精确字段比较排除新 range 字段；v4 重放必须与 reducer 完全相等。待发布 outbox 从已迁移 operation-log source 重写，避免两份 envelope 分叉。迁移测试源码覆盖 v1/v2/v3 升级、v3 impact 篡改拒绝及 v4 幂等接受。
6. **边界与失败语义**：该协议要求物理 extent 不变；服务端原先允许在 workbook table 内插入并扩展范围，随后会触发非业务型构造异常。现于 mutation 前以可观察 `UNSUPPORTED_FEATURE` 拒绝这类尚无 data-block/table transaction 的扩容；同时修正静态签名复核发现的无效 header-row 校验调用，以及 wire key 断言仍停留在 v3 的测试源码。

本轮确认并修复 **8 个实现/契约根因**，另修正 **1 个过期测试契约断言**。未把轴类型、range owner 类型、版本分支或成功/拒绝断言拆分计数；本轮仍未达到目标中提出的 30 个独立问题，因此不能据此宣称整个 Structural Editing & Reference Integrity 目标已完成。

新增 Java reducer、merge、migration 回归源码及既有 TS protocol/ACK/history 用例；按当前静态阶段要求没有运行本地 tests/build/lint/typecheck，也没有实测。该 commit 的 operation-log payload 升级是持久化写入；部署/迁移前需备份数据库。若 v4 repeatable migration 已应用，回滚应用代码必须同时恢复迁移前数据库快照；仅回退代码不会把 v4 历史降回旧版本。`aa38d3c4` 的两个远端 `canonical-build` 因 command-runtime 测试缺少 `sheet` 局部变量、issue-code union 漏项及 range-delta inverse 联合类型返回不匹配而失败；`bfeeea3d` 修正了这些 TS 问题，但其两个远端 `canonical-build` 又发现 Java reducer 的重复 `rangeNode` 定义和 cell-shift 未传入 `rangeOwnerDeltas`。`ede9bd62` 删除重复定义并修复了类型/参数链，远端编译通过；Java 用例因测试子类名不符合 Flyway migration 命名规则而失败，已将测试子类改为 `V4__TestMigration`；两个 `canonical-build` 在 `3a5e19d2` 均通过。后续静态复核确认 cell-shift 在 reducer 前 fail-close 排除受支持 range owners，且 metadata reducer 不写这些 owner，因此 patch 显式携带空 range facts，避免额外的两轮 owner 快照扫描；该改动尚无 CI 结果。PR #345 仍为 draft，最终浏览器/Excel 实测待后续执行。

**尚未覆盖的边界**：当前 v4 range-owner union 仍只覆盖以上三类。Chart/pivot/drawing source ranges 等其它 metadata references 即使被 structural reducer 重映射，也尚未纳入同一 owner-fact 协议；OOXML 往返、浏览器协同和原生 Excel corpus 也未验收。它们仍属于未完成目标，不能由本节结果代替。

### Historical stage — range-owner history 与 ACK 消费（早于上方 v4 follow-up）

**当时范围**：把 core-model 已生成的 `data-region`、`workbook-table`、`data-source` 三类几何 facts 接入本地 `CommandRuntime` history/undo/redo 和 committed-ACK consumer；该阶段尚未扩展 StructuralPatch wire v3、协作协议、Java reducer/日志或 OOXML。上方 v4 follow-up 已补入 Java 与持久化迁移；chart/pivot/drawing source ranges 及 OOXML 仍未完成。

六轮静态自审：

1. **生产者至消费者**：轴变换返回的 `rangeOwnerDeltas` 之前既不保存在 `MutationInfo`，也不进入 inverse history；已将精确 before/after facts 写入 mutation/inverse，并随 undo/redo/remote replay effect 暴露。ACK 与本地操作先按完整 facts 比较，server patch 不同会使该 history entry 失效；服务端目前仍未传这些 facts。
2. **状态与身份**：data-region 按 `(sheetId, regionId)`，table/source 按各自稳定 ID 定位；读取必须唯一，sheet/range identity 要一致。facts 在 mutation registry、ACK sparse preflight 和 apply 前验证 Excel 坐标上限、sheet identity、轴平移不改变几何尺寸及 data-region header 边界；owner 缺失、重复、漂移均 fail-close。
3. **提交原子性**：ACK 的所有 owner facts 先做 sparse overlay preflight，再可能失效 history 或写 live model；第二个 owner 的拒绝不会留下第一个 owner 的部分范围更新。历史失败仍由 `preflightHistory` 在 detached workbook 上先验证，拒绝后不移动 undo/redo 栈。
4. **data-region 扇出与内存**：旧消费草案按每个 region 重新线性搜索 owner 并重复调用 `replaceDataRegions`，D 个变化可能引发 D 次全 sheet 扫描/索引重建。现先为每个受影响 sheet 建一次 ID 索引，批量比较所有 delta，再每张 sheet 单次替换；`WorksheetModel.replaceDataRegions` 复用已经完整验证的 bounds index，不再清空后第二次重建。仍有一次随该 sheet region 数线性增长的 replace，不宣称 affected-only 或测得提速。
5. **其他 range owners 与规则兼容**：table/source 只写 `sourceRange`，不 clone/替换 fields、blocks、rowOrder 或 revision；范围不变时不写。混合 table/source/region facts 在同一 ACK item 先整体核验，且旧公式/name sparse preflight 保持顺序。回归源码增加两个 region 同批只 replace 一次、table/source 成功及第二个 owner 拒绝原子性断言；修正数据源 fixture，使字段数/范围宽度/rowCount 满足 manifest 契约。
6. **协议与接受边界（历史状态）**：当时本地 ACK consumer 即使已支持，也不能从 exact-key StructuralPatch v3 收到 range deltas；Java patch constructor、reducer 和持久化迁移尚未接入。所需的 v4 wire/log/outbox 与 Java authority 落点已在上方 follow-up 实施；本段只保留该先后阶段的审计记录，不代表当前仍停留于 v3。

本轮按独立根因确认并修复 **4 项**：range facts 没有进入客户端 history/ACK 消费；region fan-out 触发重复线性扫描/索引重建；ACK precondition 失败前先改写 history 状态；authoritative geometry 缺少运行时边界/身份校验。数据源测试夹具错误是测试输入修正，不计产品问题。没有将各 owner 类型、不同 replay 入口或断言拆分凑数，本轮未达到“至少 30 个独立问题”；要达到该数量必须继续扩大到完整协议/Java/OT/持久化链路并只计有证据的根因。

按要求只完成源码静态检查并新增回归源码；未运行 tests/build/lint/typecheck/browser，也没有本轮 head 的 CI 结果。只读 diff/whitespace 检查尚待提交前执行。没有协议/数据迁移；PR #345 保持 draft。此分步修改需与 range-owner wire/Java 迁移保持同一 PR，若回退应整体 revert 本节实现、类型消费与回归源码，不能留下客户端半套 replay。

## 2026-09-26 continuation — remote replay and reference-index complexity

Static source tracing confirms `CommandRuntime.applyRemoteMutations` performs detached `preflightHistory` by exporting the entire workbook snapshot, reconstructing a second `WorkbookModel`, and replaying all mutation handlers before replaying them again on the live workbook. This is required today for atomic rejection; deleting the preflight alone would be unsafe. The intended replacement is complete sparse facts with all-owner preconditions checked before one apply, not another full snapshot diff.

The app supplies `FormulaEngine.dependencies` to `StructuralReferenceOwnerIndex`, but every resolution still calls `indexStructuralFormulaRules`, which enumerates every conditional-format/data-validation formula and range across every worksheet before returning the index. The detached history runtime has no persistent provider and calls `buildStructuralReferenceIndex`, which walks all occupied cells/formulas, names, and rules. The current index contract also omits chart, pivot, sparkline, drawing, filter, table, and spill owner geometry, which structural code still scans separately. `ReferenceTransformDomain` point/interval behavior already shares cross-language vectors; move/permutation and owner-family semantics remain outside that proof.

Six non-overlapping static review passes confirmed two independent performance roots (full detached remote replay and full CF/DV rule enumeration) plus the incomplete typed owner index. No local tests/build/browser or performance benchmark ran. At audit time, both `canonical-build` checks for PR #345 head `291f4b09` were green; Java-only planning, sparse authoritative facts, facts-only remote/history replay, browser, native Excel, and large-workbook performance acceptance remain open.
