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
