# Linked calculation architecture

## Canonical ownership

- WorkbookSnapshot 11 owns external link bindings and Record table field definitions. Values read from other workbooks are transient, tied to the authenticated subject, source revision and access revision; snapshots and OOXML metadata do not persist this cache.
- External formula workbook/sheet tokens resolve through a binding to stable source workbook/sheet identities. Java independently authorizes the target and source read and removes hidden inputs. Hidden range dependencies return `#BLOCKED!`; missing sources return `#REF!`; unavailable reads clear cached inputs and return `#N/A`.
- Record identity is a stored, unique text field. Editing resolves `(tableId, recordId, fieldId)` before applying the canonical cell write authority. Identity and calculated fields are read only. Physical coordinate writes to Record data are rejected. Field IDs, ordinals and source dimensions are immutable during configuration, and Record source ranges cannot overlap another Record owner.
- `record.restore` belongs to authenticated undo: Java checks the original operation and the complete restored preimage. Its stored cell/null payload restores metadata and sparse cell absence exactly. A normal write cannot invoke it. The ID column is an authorization read precondition; edit permission applies to the addressed field.
- Formula, Lookup and Rollup definitions live once per field. Their derived Record owners and relationship membership indexes project into the existing FormulaEngine graph and calculation Worker protocol 4. There is no second evaluator or persisted computed value store.
- Sheet deletion and endpoint crossing rewrite 3D references through the shared formula reference representation. Surviving endpoints retain orientation. Ordinary references to a deleted sheet become `#REF!`; external references are opaque. StructuralPatch 10 records reversible cell, rule, name and Record field formula facts. Java commits these facts and validates undo against the complete preimage.

## Migration and interoperability

Flyway V13 explicitly upgrades stored snapshots and verifies/replays operation history into the current structural patch contract. This boundary normalizes missing worksheet defaults in stored snapshots and historical sheet restore payloads. Runtime sheet creation emits the complete canonical shape; undo comparison remains exact. The browser snapshot import boundary upgrades older snapshots; runtime readers require version 11. ReactSheets OOXML metadata 3 stores bindings, identity fields and calculated definitions. Old metadata 2 upgrades at import only.

Native Excel external-link parts are not converted into service bindings. Desktop Excel acceptance is Blocked in this environment. Record fields currently require a worksheet source with at most 100,000 records; block sources and native Excel calculated-column interoperability are outside the supported contract. External input graphs and nested cached calculation values are bounded to 100,000 inputs/values. Source refresh uses periodic authorized reads; recursive external workbook graphs and external structured-table references are not supported.

An external source containing Record tables requires read permission for each complete Record source range. Partial hidden Record inputs reject the graph with `ACCESS_HIDDEN`; they cannot remove definitions and produce blank calculated values. Ordinary worksheet hidden inputs retain blocked geometry and calculate `#BLOCKED!`. Owner access continues to follow the canonical range resolver.

External refresh coalesces only within the same workbook model and FormulaEngine. Hydration starts a fresh request immediately, and late results from the old context are ignored. Initial collaboration synchronization preserves an already authorized access projection when no local cache has been written; changed or unknown access still purges browser-owned data before hydration.

## Verification

The user requested pushing the implementation to PR #346 before continuing acceptance. The initial checkpoint was pushed, and subsequent fixes retain remote formula catalog and dependency ownership changes. The final commands, results, real HTTP and browser scenarios, screenshot evidence and release blockers are recorded in [linked-calculation-acceptance.md](linked-calculation-acceptance.md).

## Rollback

Retain a database backup before applying V13. Rolling back requires restoring the matching database and executable together; older binaries cannot safely read version 11 / StructuralPatch 10. Do not downgrade schemas or modify migration checksums in a live database.
