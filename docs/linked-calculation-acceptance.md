# Linked calculation acceptance — 2026-10-02

## Delivery and release state

The first implementation checkpoint was pushed to PR #346 before acceptance
continued, as requested. That PR was subsequently merged remotely and its branch
deleted. Follow-up fixes are delivered on `codex/linked-calculation-lifecycle`
against the current main branch, retaining the formula catalog/dependency changes
and the independently added project README.

The four linked calculation domains have successful browser and backend evidence.
Full release acceptance remains **Blocked**: the broader frontend unit suite has
83 existing failures, and desktop Excel is unavailable. These failures are not
waived; this follow-up does not push or merge feature changes directly to main.

## Executed verification

| Gate | Result | Raw log outside Git |
| --- | --- | --- |
| Java 21 + H2, complete `mvn -o test` | **315 passed**, 0 failures/errors/skips | `/tmp/luckysheet-linked-java-tests-9.log` |
| `npm run build`, including TypeScript | **Pass** | `/tmp/luckysheet-linked-build-8.log` |
| `npm run test:calculation-domain` | **451 passed** (446 + 5) | `/tmp/luckysheet-linked-calculation-8.log` |
| `npm run test:native-codecs` | **17 passed**, real-file corpus | `/tmp/luckysheet-linked-native-codecs-8.log` |
| `npm run check:boundaries` | **Pass**, contracts, registry, acceptance matrix | `/tmp/luckysheet-linked-boundaries-8.log` |
| `npm run test:unit` | **1,442 / 1,525 passed; 83 failed** | `/tmp/luckysheet-linked-full-unit-8.log` |
| Real Chromium UI + Java service + persisted H2 | **10 / 10 passed** | `/tmp/luckysheet-linked-browser-15.log` |

The full-unit failing test identities were compared with the recorded previous
84-failure run. There are no new failing entries, and one spill-environment fixture
now passes. Comparison: `/workspace/luckysheet-linked-unit-comparison.json`.
The remaining failures include canonical name/hydration fixtures, reversible
rule/command facts, stale protocol/authority fixtures, data-source ownership and
query/pivot/drawing lifecycle cases. A green focused gate does not certify them.

Java verification used the Maven Java 21 container, the real Spring/H2 integration
suite and the local dependency cache. The calculation gate includes the complete
136-function catalog, inline/Worker execution, sparse range and dependency tests,
linked lifecycle, Record calculation and linked OOXML metadata tests.

## Actual browser behavior

Chromium at 1440×960, Vite on port 4180 and Spring Boot on port 8082. Authored
inputs, bindings, Record definitions and relations were created through the actual
UI. Java POST/ACK responses and persisted snapshots were inspected. Canvas text
observers only read actual paint calls; they exclude headers and do not inject
calculation results. Each observation and screenshot verifies the selected sheet.

| Scenario | Observed result |
| --- | --- |
| Bind `SUM([Source.xlsx]Sales!B2:B20)` | **10** |
| Edit source to 25, rename Sales to Revenue, reopen target | **25**, stable source sheet ID binding retained |
| `SUM(Jan:Mar!A1)` with 10/20/30 | **60** |
| Delete Jan endpoint | `SUM(Feb:Mar!A1)` → **50** |
| Undo endpoint deletion, then reopen | Original formula and **60** restored |
| Record field `[Quantity]*[Price]` | **20 / 60 / 20** |
| Reverse Lookup and SUM Rollup for customer c1 | **20, 60** and **80** |
| Sort Orders descending, edit visible C2 to 7 | Stable record **o3** edited; c1 remains **80**, c2 becomes **35** |
| Reopen Customers view | Lookup/Rollup remain **80 / 35**, correct sheet stays selected |

The 10 browser checks have no console warnings/errors, page exceptions or
unexpected HTTP errors. Evidence and API revision traces:
`/workspace/luckysheet-linked-browser-evidence.json`.

Screenshots outside Git:

- `/workspace/luckysheet-linked-binding-proof.png`
- `/workspace/luckysheet-linked-update-proof.png`
- `/workspace/luckysheet-3d-delete-proof.png`
- `/workspace/luckysheet-record-formula-proof.png`
- `/workspace/luckysheet-record-lookup-rollup-proof.png`
- `/workspace/luckysheet-record-reload-proof.png`

## Backend rejection and undo

Real HTTP verification covers source membership and hidden-input revocation,
sparse Record cell undo, calculated field rejection, invalid relation targets,
physical coordinate write rejection and restore-without-undo rejection. The
authenticated Java undo compares the entire restored snapshot with its preimage;
rejected Record transactions retain their revision and snapshot. The complete
Record source range must remain readable for external calculated graph access.

Additional Java tests reject duplicate identities, missing Lookup targets,
overlapping Record owners, invalid restore payloads and tampered 3D undo. Record
ID read permission is separate from field edit permission. Runtime tests verify
fresh-engine external requests and access changes clearing browser-owned state.

## Stored data migration rehearsal

A pre-change file-backed H2 backup was copied to a separate directory and opened
with the final Java code. Flyway V13 migrated all stored snapshots and replayed
operation/checkpoint history successfully. The migrated database retained
**39 workbooks**. A second startup validated the migrations and reported version
13 up to date with no migration required.

Logs: `/tmp/luckysheet-linked-migration-rehearsal-3.log`,
`/tmp/luckysheet-linked-migration-counts.log`,
`/tmp/luckysheet-linked-migration-reopen-1.log`. Original backup:
`/workspace/pre-linked-data-20261002.mv.db`.

## Supported boundary and rollback

See [linked-calculation-architecture.md](linked-calculation-architecture.md) for
WorkbookSnapshot 11 / StructuralPatch 10 / Worker protocol 4 / OOXML metadata 3,
the explicit migration boundary and rollback procedure. Bounded worksheet Record
sources and periodic authorized external refresh are demonstrated. Block-backed
Record writes, recursive external graphs, native external structured-table
execution and complete native calculated-column interoperability are outside the
supported contract. Desktop Excel acceptance is **Blocked**.

Restore a matched database backup and executable when rolling back V13. Do not
downgrade schema fields in place or alter applied migration checksums. QA scripts,
screenshots, request traces and databases are intentionally outside the repository.
