# Linked calculation architecture

## Canonical ownership

- WorkbookSnapshot 11 owns external link bindings and Record table field definitions. Values read from other workbooks are transient, tied to the authenticated subject, source revision and access revision; snapshots and OOXML metadata do not persist this cache.
- External formula workbook/sheet tokens resolve through a binding to stable source workbook/sheet identities. Java independently authorizes the target and source read and removes hidden inputs. Hidden range dependencies return `#BLOCKED!`; missing sources return `#REF!`; unavailable reads clear cached inputs and return `#N/A`.
- Record identity is a stored, unique text field. Editing resolves `(tableId, recordId, fieldId)` before applying the canonical cell write authority. Identity and calculated fields are read only. Physical coordinate writes to Record data are rejected.
- Formula, Lookup and Rollup definitions live once per field. Their derived Record owners and relationship membership indexes project into the existing FormulaEngine graph and calculation Worker protocol 4. There is no second evaluator or persisted computed value store.
- Sheet deletion and endpoint crossing rewrite 3D references through the shared formula reference representation. Surviving endpoints retain orientation. Ordinary references to a deleted sheet become `#REF!`; external references are opaque. StructuralPatch 10 records reversible cell, rule, name and Record field formula facts. Java commits these facts and validates undo against the complete preimage.

## Migration and interoperability

Flyway V13 explicitly upgrades stored snapshots and verifies/replays operation history into the current structural patch contract. The browser snapshot import boundary upgrades older snapshots; runtime readers require version 11. ReactSheets OOXML metadata 3 stores bindings, identity fields and calculated definitions. Old metadata 2 upgrades at import only.

Native Excel external-link parts are not converted into service bindings. Desktop Excel acceptance is Blocked in this environment. Record fields currently require a worksheet source with at most 100,000 records; block sources and native Excel calculated-column interoperability are outside the supported contract. External input graphs are bounded to 100,000 nonempty inputs. Source refresh uses periodic authorized reads; recursive external workbook graphs and external structured-table references are not supported.

## Verification checkpoint before continuing acceptance

The user requested pushing this checkpoint to PR #346 before continuing development.

- Java full suite previously passed 310 tests; subsequent linked descriptor/integration suite passed 4 tests.
- Focused calculation, Record and OOXML tests previously passed; final rerun remains pending after the latest runtime corrections.
- Real Chromium against Spring Boot Java 21 and file-backed H2 passed external binding, source update and rename, 3D endpoint deletion, undo, and reopening after undo.
- Real Record UI acceptance exposed loading empty physical cells over calculated owners. The runtime loader now preserves calculated owners; browser verification is continuing.
- The latest full frontend suite was 1,229 tests / 1,145 pass / 84 fail. These 84 failures match the recorded previous baseline. The PR remains draft; this checkpoint is not a release acceptance claim.
- QA scripts, screenshots, request traces and H2 data are stored outside the repository in `/workspace` and `/tmp`.

## Rollback

Retain a database backup before applying V13. Rolling back requires restoring the matching database and executable together; older binaries cannot safely read version 11 / StructuralPatch 10. Do not downgrade schemas or modify migration checksums in a live database.
