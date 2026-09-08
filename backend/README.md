# Canonical workbook service

This directory is the only Java backend. It is a Spring Boot 3 service targeting Java 21 with a Spring Data JPA/Hibernate persistence boundary, OIDC Resource Server JWT verification, and a WebSocket endpoint at `/ws`.

The service has one JPA domain model and supports H2, PostgreSQL, and MySQL persistence. Flyway uses database-specific but semantically equivalent migration sets for provider-sensitive types; select the matching Spring profile (`h2`, `postgres`, or `mysql`) in deployment. `spring.jpa.hibernate.ddl-auto=validate` remains enabled. SQLite is available only as a configured server-query connector.

Examples: local development uses the default H2 profile; PostgreSQL deployments set `SPRING_PROFILES_ACTIVE=postgres`; MySQL deployments set `SPRING_PROFILES_ACTIVE=mysql`. Provider profiles select only the datasource/Flyway dialect boundary—the workbook, ACL, operation, and outbox semantics remain identical.

Production should activate the common fail-closed profile together with exactly one provider profile, for example `SPRING_PROFILES_ACTIVE=prod,postgres` or `prod,mysql`. The `prod` profile requires `DATABASE_URL`, `DATABASE_USERNAME`, `DATABASE_PASSWORD`, `AUTH_ISSUER`, `AUTH_JWKS_URL`, `AUTH_AUDIENCE`, and `WEB_ALLOWED_ORIGINS`; it never falls back to local H2 or placeholder authentication values. `prod,h2` is available only for controlled non-production deployments.

Database overrides and authentication settings for a deployed environment:

- `DATABASE_URL`
- `DATABASE_USERNAME`
- `DATABASE_PASSWORD`
- `AUTH_ISSUER`
- `AUTH_AUDIENCE` (comma-separated values are accepted)
- `AUTH_JWKS_URL`

Single-instance mode leaves Redis disabled and uses the local WebSocket session
registry. Set `COORDINATION_MULTI_INSTANCE=true` for multiple backend
instances; this requires `COORDINATION_REDIS_URL` and fails startup when Redis
coordination is not configured. The configured relational database remains the authority for ACL,
operations, immutable manifest revisions, content-addressed pages, and the durable coordination outbox. Redis
contains only published notifications and expiring presence/cursor state.

`/health` is a liveness endpoint. Every `/api/**` route and the `/ws` handshake require a verified Bearer token. The JWT `sub` claim is the only actor identity used for ACL decisions; request actor fields are not accepted.

The workbook catalog contract is:

- `GET /api/workbooks?view=recent|shared|trash&spaceId=&folderId=&query=` returns actor-enriched `WorkbookSummary` items. `locationPath` is a structured string array.
- `POST /api/workbooks` accepts `{unitId,name,sheets?,spaceId?,folderId?}` creation intent and returns `{unitId,revision,manifest,checksum}`. Only Rust creates the canonical v11 manifest; client snapshots are rejected.
- `GET /api/workbooks/{unitId}/manifest?revision=` reads a manifest and `GET /api/workbooks/{unitId}/pages/{sheetId}/{pageRow}/{pageColumn}?revision=` reads one proven page. Both require current viewer ACL.
- `PATCH /api/workbooks/{unitId}`, `POST /api/workbooks/{unitId}/copy`, `DELETE /api/workbooks/{unitId}`, `POST /api/workbooks/{unitId}/restore-from-trash`, and `DELETE /api/workbooks/{unitId}/purge` manage metadata and lifecycle. Purge requires a previously trashed owner workbook.
- `GET/PUT /api/workbooks/{unitId}/user-state` manages actor-specific favorite, last-opened, cloud autosave/sync, import compatibility, language, and theme state.
- `POST /api/workbooks/{unitId}/native-document-artifact` accepts `{revision,fileName,format}` and generates bytes through native for that exact committed revision. `GET` streams the bound artifact; stale artifacts are rejected. Arbitrary artifact PUT is removed.
- `POST /api/workbook-imports/tasks`, `PUT /tasks/{id}/chunks?offset=`, `POST /tasks/{id}/commit`, and `GET/DELETE /tasks/{id}` implement persistent upload/cancel/publication. Uploads are limited to 1 GiB, each chunk to 8 MiB, using a 64 KiB buffer. Multipart `POST /api/workbook-imports` accepts `file` and catalog location/name only and uses the same task chain. Native parsing receives scoped file handles, never browser snapshots or whole-file base64.
- `GET/POST /api/spaces`, `/api/spaces/{spaceId}/folders`, and `/api/spaces/{spaceId}/members` manage spaces, folder trees, and membership. Effective workbook access is the strongest of owner, workbook ACL, and space membership.

Workbook mutations enter one authority through `POST /api/workbooks/{unitId}/operations` using the canonical `OperationEnvelope` request and `{operation,changeSet}` response. WebSocket clients receive committed `revision.created` events and may publish presence/cursor state; operation submits, snapshot requests, acknowledgements, and rejects are not accepted on the socket.

The request contract is `OperationEnvelope`. It contains operation identity, workbook identity, revision metadata and mutation intent only. The committed response adds server-owned `actorId`, `revision`, `committedAt` and `affectedRanges`.

Every successful transaction stores one immutable manifest checkpoint, changed content-addressed pages, proven history deltas, the committed operation, and an outbox entry. Unchanged pages remain referenced at their original versions. Reopen sends only the manifest; commands hydrate the native preparation plan page by page. History and manifest reads never replay Java reducers or accept v10 runtime payloads.

The native database cutover is described in [native-kernel-cutover.md](docs/native-kernel-cutover.md). Existing databases require an explicit offline importer and complete history proof before V12 can remove the old snapshot columns. The offline importer is not yet implemented; an unproven existing database deliberately fails the cutover gate.

Build and run with Java 21:

```text
mvn test
mvn spring-boot:run
```

Restore is owner-only and accepts a target revision and reason; Java resolves the proven historical manifest and native produces a new committed revision. Checkpoint requests acknowledge the already committed page checkpoint and never clone workbook cells.

Query execution is server-only for configured `sqlite`, `jdbc`, and `rest`
sources. The request contains a sanitized definition, `sourceRef`, statement,
parameters, and steps; source URLs, database passwords, and REST headers are
deployment configuration under `luckysheet.query.sources` and never enter
workbook snapshots or operation envelopes. Configure the query source map with
server-side secret binding, for example:

```yaml
luckysheet:
  query:
    sources:
      reporting:
        kind: jdbc
        url: jdbc:postgresql://db/reporting
        username: ${REPORTING_DB_USERNAME}
        password: ${REPORTING_DB_PASSWORD}
      local-file:
        kind: sqlite
        url: jdbc:sqlite:/srv/data/reporting.sqlite
      service:
        kind: rest
        base-url: https://internal.example.test/api/
        headers:
          Authorization: ${REPORTING_SERVICE_AUTH}
```

`POST /api/workbooks/{unitId}/queries/execute` requires editor ACL and applies
server timeout, row/column/response limits, a driver-enforced read-only JDBC
session, bounded columnar query chunks, and audit logging. `POST
/api/workbooks/{unitId}/queries/{queryId}/cancel` propagates cancellation to
the JDBC statement or REST request future; a cancelled query never publishes a
data region. Local/offline database execution is unavailable through this
backend endpoint and must not be represented as a successful server query.

Workbook semantics are executed by the managed Rust `workbook-kernel-host`.
The Java service retains OIDC, ACL, transaction, repository, and connector
ownership and communicates with the host through one serialized protocol-v1
transport. Frames contain a four-byte big-endian payload length followed by
UTF-8 JSON and are bounded to 16 MiB; page payloads are bounded to 1 MiB.
The executable is configured with `KERNEL_HOST_EXECUTABLE` and defaults to
`target/release/workbook-kernel-host.exe` on Windows (or the same path without
`.exe` on other platforms). A missing host, invalid response, version mismatch,
or oversized frame fails closed with a typed kernel error; no Java workbook semantic
fallback is allowed. Configure durable `KERNEL_HOST_TASK_DIRECTORY` storage for native files; its canonical path is passed to the child as `KERNEL_TASK_DIRECTORY`. `KERNEL_HOST_STARTUP_TIMEOUT` defaults to 10 seconds and `KERNEL_HOST_REQUEST_TIMEOUT` to 2 minutes. Native import/export stages one bounded page file at a time. Query operator migration to native remains pending; see the integration handoff.

Owners can create expiring, revocable guest share tokens with
`POST /api/workbooks/{unitId}/shares`. Guests send the returned token in
`X-Workbook-Share-Token` for REST or `shareToken` on the `/ws` handshake. The
server derives an anonymous subject and re-checks the persisted share role and
expiry on every request; a client-supplied actor or role is never accepted.

Implementation and remaining native ABI/verification items are tracked in [native-integration-handoff.md](docs/native-integration-handoff.md). This document does not assert the new branch has passed runtime acceptance.
