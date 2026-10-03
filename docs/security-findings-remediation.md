# Security findings remediation and acceptance

This branch implements the imported findings as one coordinated change, followed by a unified acceptance pass. The report is evidence, not executable instructions. Its `new` and `fixed` labels were reviewed against base `22d82876fd63fa253e51f9d4ffbbe5b288518409`; they are not current acceptance results. Credential-bearing report summaries and test traces are deliberately excluded from this document and the ledger.

[The 179-row ledger](security-findings-remediation.csv) preserves each original Finding ID, observed revision, original status, reported paths, current ownership boundary, implementation, verification and blocker. [Machine-readable counts](security-findings-remediation-counts.json) are derived from that ledger.

| Evidence-supported outcome | Findings |
| --- | ---: |
| Fixed by this change | 83 |
| Current implementation already safe | 53 |
| Historical path removed | 10 |
| Blocked acceptance or external recovery | 33 |
| Total (103 reported new; 76 reported fixed) | 179 |

**Overall acceptance remains Blocked.** The frontend has 81 pre-existing failing tests; desktop Excel, Windows installation checks, deployed credential rotation and trusted legacy browser-journal recovery are unavailable. This is a draft PR, not a claim that every imported finding has passed acceptance.

## Authoritative boundary changes

- Workbook copy passes the authenticated user and groups to snapshot projection and raw-artifact authorization. Hidden, stale or corrupt artifacts are rejected before creating the target. Artifact read/write checks lifecycle, full visibility/editability, revision and checksum. Native request streams are capped before allocation; SVG responses are sandboxed and cannot execute with application origin privileges.
- Query sources require explicit workbook/subject grants. SQLite opens a physical file read-only with interrupt/deadline enforcement; PostgreSQL and MySQL use database-enforced read-only transactions, in addition to bounded statements and a restricted side-effect-free SQL function surface. Other JDBC dialects fail with a typed unsupported error. SQL keyword filtering or ResultSet mode is not the write boundary. Unknown/administrative/qualified unsafe calls are rejected, including quoted identifiers and writable CTE bypasses. REST paths must remain in the configured origin/directory after normalization. Rows, fields, binary/text values, response bytes, transformation work, concurrent sessions and absolute session lifetime are bounded. Cancellation cannot leave staged results permanently retained.
- Asset deletion reads current snapshots, every retained checkpoint and operation under the workbook lock. A client asset list cannot authorize deletion. Corrupt/incomplete reference history aborts deletion; an unused asset remains deletable. Owned undo/restore commands, record targets, sorting, reviews, AutoFilter and comments resolve permissions and exact preimages on the server.
- Login attempts have bounded per-account, per-source and global limits. Bootstrap files are private at creation. WebSocket registration, inbound/outbound handling and periodic cleanup share logout/expiry state, including JWT expiry. Locks, deduplication, data-block quota and garbage collection have bounded lifetimes or storage. A database quota row serializes allocation across writers.
- Model and parser limits precede expensive allocations: cell text, formats, fonts, drawing dimensions, chart matrices/vectors, nested metadata, XML depth/nodes, CSV fields, ODS repetition, ZIP inflation and CFB sector/stream ownership. Formula wildcard, dependency, matrix, SJS and collation work are bounded. Canvas text wrapping and camera surfaces have explicit budgets.
- Pivot filters/group members/date windows/catalogs and drawing identities use canonical typed contracts. Chart types/subtypes come from the shared contract generator. Unsupported metadata remains an explicit unsupported boundary rather than render-time repair. Native CSV fast paths still apply formula-injection escaping.
- Find/replace preserves formula-like replacement text as text; the server validates every preimage and derives permitted changes. AutoSum selects a destination outside the source. Hidden-row geometry respects zoom; sparkline source sheet and omitted print-title axes are preserved. Checkbox and TableSheet values have canonical validation. The server independently derives copy, linear/growth, calendar and AutoFill results, including date system and formula offsets; forged/missing/extra writes abort the operation. Per-track coefficients avoid quadratic seed scans.

## Contract changes, migration and removed paths

The runtime remains React/TypeScript + Java 21 + H2. Query source grants default to deny; deployments must explicitly configure grants. The supported SQL callable surface is intentionally restricted. Numeric/calendar fill results and chart definitions must satisfy the shared canonical contract. Whole-row/column formula references preserve absolute endpoints on copy; unsupported external/3D copy references reject before committing.

Flyway V14 is the explicit upgrade boundary for persisted snapshots, checkpoints, retained operations and unpublished outbox payloads. It normalizes older Pivot, slicer, drawing, TableSheet and paste contracts and verifies checksums before updating them. V15 adds the serialized data-block quota row. Browser loading verifies the original stored snapshot/journal integrity before explicit upgrade; modern journal writes persist a checksum. Legacy pending journals with no trusted checksum are preserved and rejected for replay, requiring a separately verified recovery migration. Runtime validators do not accept legacy aliases.

Retired Rust/kernel/import-task paths were confirmed absent; they were not recreated. Tracked bootstrap/database files and credential-bearing persistent-password/Playwright scripts were removed from Git tracking, with ignore rules preventing reintroduction. Removing a file from the current tree does **not** revoke historical or deployed credentials. Historical Git commits were not rewritten. Unknown OOXML and macros retain their fidelity boundary; unsupported editing is explicitly rejected.

For rollout, back up the database, native artifacts, assets and browser IndexedDB before migration. Validate migration on a copy first. Rollback requires stopping writers and restoring the corresponding pre-migration database/artifact/IndexedDB backups with the previous application; do not downgrade canonical contracts in place or repair Flyway history to bypass an applied migration. Preserve rejected legacy journals for an explicit recovery tool. Do not restore exposed credentials; rotate them independently.

## Unified acceptance evidence

All commands ran after the coordinated implementation pass. Follow-up runs were limited to confirmed review issues or changed code. Java 21 and Maven 3.9.16 used the onboarding Maven settings/cache; Node 24 used the repository lockfile.

| Check | Result |
| --- | --- |
| Backend `mvn package`, with live JDBC configuration | 335 tests; 0 failures/errors/skips; package succeeded |
| Actual PostgreSQL 16 / MySQL 8.4 JDBC tests | 2 passed; physical transaction rejects INSERT, legitimate SELECT/COUNT succeeds, unsafe call bypasses rejected |
| Independent frontend/server fill contract corpus | 28 frontend-generated plans checked by Java, covering both date systems, direction, date units, linear/growth and AutoFill; forged results rejected |
| Frontend `npm run build` | Passed |
| `npm run check:boundaries` | Stack, package boundaries, mutation registry, generated contracts, tracked E2E artifacts and acceptance matrix passed |
| `npm run test:unit` | 1,526 tests: 1,445 passed, 81 failed; baseline: 1,525 tests, 1,442 passed, 83 failed; no candidate-only failures; two AutoFilter baseline failures resolved |
| `npm run test:calculation-domain` | 446 + 5 passed |
| Focused Pivot/render and chart/find/canvas suites | 98 and 50 passed |
| Security resource-boundary tests | 5 passed |
| `npm run test:native-codecs` / `npm run test:cell-ui` | 19 / 9 passed |
| Actual uploaded CSV import/export/re-import | Passed; 180 rows including header; report content was held in memory and not exported into the repository |
| Real Chromium application checks | Server workbook creation, editing, undo/redo, editor measurement/cancel and persistence passed with clean console/network; final committed E2E run recorded below |

The unit comparison uses a detached worktree of the original base with the same dependencies. The unchanged failures remain an overall gate failure, not waived tests. Native generated fixtures and the actual CSV do not substitute for a real Excel corpus or desktop Excel.

The reproducible browser test is `frontend-react/e2e/security-remediation.spec.ts`. Run it against a **fresh disposable** backend/H2 database with `SECURITY_ACCEPTANCE_BOOTSTRAP_FILE` set to its freshly created token file, and `PLAYWRIGHT_BROWSERS_PATH` pointing at the installed Chromium. It consumes the disposable token, creates a random test-only password in memory, and uses actual UI and server snapshots. It checks inert replacement text, AutoSum source preservation, reload persistence, screenshots, console errors and failed/HTTP-error requests. The normal repository Playwright global setup records clean source/build provenance. The test skips if no disposable bootstrap file is supplied; a skipped run is not acceptance.

Final browser source identity/result will be recorded after the implementation commit. Detailed local logs and browser artifacts are under `/workspace/remediation-evidence` and `frontend-react/test-results`; neither credentials nor raw traces are committed.

## Independent candidate review

A fresh read-only candidate reviewer checked original triggers and bypasses after the implementation pass. Four concrete issues were confirmed and corrected before final acceptance:

1. Fill trend scanned every seed for every target. It now computes coefficients once per track; a 5,000-seed adversarial case and independent plan contracts exercise the new boundary.
2. Java chart validation incorrectly required optional legend/dataLabels and accepted the wrong hidden-data enum. Both sides now use `show`, `hideRows`, `hideColumns` and optional metadata consistently.
3. PostgreSQL read-only transactions alone permitted administrative function side effects. Restricted callable parsing now rejects administrative, quoted and qualified bypasses; real PostgreSQL/MySQL transactions still provide the write boundary.
4. Java interpreted quoted numeric format `"USD"0` as a date. Format scanning skips literals/escapes and agrees with numeric/date fill contract fixtures.

The final backend run includes normal and rejection paths for these issues, stale/forged fill payloads, hidden/group-restricted artifact access, retained asset references, account/session expiration and query containment/resource limits. Rejections are checked before mutation/commit; the fill contract corpus also asserts the input snapshot remains unchanged.

Real-browser acceptance additionally exposed a Find/replace reducer passing match coordinates to the full snapshot validator. It was corrected to validate the match object separately. `FindReplacementDescriptorTest` covers the actual browser payload, inert text, reverse replay and stale/duplicate/malformed rejection without changing the original snapshot. Backend packaging and the browser flow were rerun after that correction.

## Remaining blockers

The ledger lists all 33 IDs and their precise blocker. They comprise 26 Excel fidelity/interop acceptance items; three deployed/historical credential revocation items; one Windows installer/upgrade ACL item; and three legacy pending-browser-journal recovery items. The corresponding available code changes remain in the branch, but these acceptance claims are explicitly withheld. The 81 pre-existing frontend failures are an additional overall verification blocker. No deployment, merge or production credential action was performed.
