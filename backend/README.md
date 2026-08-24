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
operations, revisions, snapshots, and the durable coordination outbox. Redis
contains only published notifications and expiring presence/cursor state.

`/health` is a liveness endpoint. Every `/api/**` route and the `/ws` handshake require a verified Bearer token. The JWT `sub` claim is the only actor identity used for ACL decisions; request actor fields are not accepted.

The workbook catalog contract is:

- `GET /api/workbooks?view=recent|shared|trash&spaceId=&folderId=&query=` returns actor-enriched `WorkbookSummary` items. `locationPath` is a structured string array.
- `POST /api/workbooks` creates a native workbook in the actor's personal space unless `spaceId`/`folderId` are supplied.
- `PATCH /api/workbooks/{unitId}`, `POST /api/workbooks/{unitId}/copy`, `DELETE /api/workbooks/{unitId}`, `POST /api/workbooks/{unitId}/restore-from-trash`, and `DELETE /api/workbooks/{unitId}/purge` manage metadata and lifecycle. Purge requires a previously trashed owner workbook.
- `GET/PUT /api/workbooks/{unitId}/user-state` manages actor-specific favorite, last-opened, autosave/sync, default cloud location, offline cache, import compatibility, language, and theme state.
- `GET/PUT /api/workbooks/{unitId}/native-package-state` streams the native package state with an SHA-256 header. `POST /api/workbook-imports` accepts `file`, `format`, `nativeMetadata`, and a browser-parsed `snapshot` multipart part and creates a new workbook atomically.
- `GET/POST /api/spaces`, `/api/spaces/{spaceId}/folders`, and `/api/spaces/{spaceId}/members` manage spaces, folder trees, and membership. Effective workbook access is the strongest of owner, workbook ACL, and space membership.

Workbook mutations are written only through `POST /api/workbooks/{unitId}/operations`. WebSocket clients receive committed `revision.created` events and may publish presence/cursor state; operation submits, snapshot requests, acknowledgements, and rejects are not accepted on the socket.

The request contract is `OperationEnvelope`. It contains operation identity, workbook identity, revision metadata and mutation intent only. The committed response adds server-owned `actorId`, `revision`, `committedAt` and `affectedRanges`.

Build and run with Java 21:

```text
mvn test
mvn spring-boot:run
```

Snapshot replacement is not exposed to ordinary editors. Checkpoints are server-generated and restore accepts only a target revision and reason; the server loads the historical snapshot and records a committed restore operation.

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
server timeout, row/column/response limits, read-only SQL validation, and
audit logging. `POST /api/workbooks/{unitId}/queries/{queryId}/cancel` cancels
an active server execution. Local/offline database execution is unavailable
through this backend endpoint and must not be represented as a successful
server query.

Owners can create expiring, revocable guest share tokens with
`POST /api/workbooks/{unitId}/shares`. Guests send the returned token in
`X-Workbook-Share-Token` for REST or `shareToken` on the `/ws` handshake. The
server derives an anonymous subject and re-checks the persisted share role and
expiry on every request; a client-supplied actor or role is never accepted.
