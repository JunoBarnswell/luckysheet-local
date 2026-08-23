# Canonical workbook service

This directory is the only Java backend. It is a Spring Boot 3 service targeting Java 21 with PostgreSQL, Flyway migrations, OIDC Resource Server JWT verification, and a WebSocket endpoint at `/ws`.

Required environment:

- `DATABASE_URL`
- `DATABASE_USERNAME`
- `DATABASE_PASSWORD`
- `AUTH_ISSUER`
- `AUTH_AUDIENCE` (comma-separated values are accepted)
- `AUTH_JWKS_URL`

Single-instance mode leaves Redis disabled and uses the local WebSocket session
registry. Set `COORDINATION_MULTI_INSTANCE=true` for multiple backend
instances; this requires `COORDINATION_REDIS_URL` and fails startup when Redis
coordination is not configured. PostgreSQL remains the authority for ACL,
operations, revisions, snapshots, and the durable coordination outbox. Redis
contains only published notifications and expiring presence/cursor state.

`/health` is a liveness endpoint. Every `/api/**` route and the `/ws` handshake require a verified Bearer token. The JWT `sub` claim is the only actor identity used for ACL decisions; request actor fields are not accepted.

The request contract is `OperationEnvelope`. It contains operation identity, workbook identity, revision metadata and mutation intent only. The committed response adds server-owned `actorId`, `revision`, `committedAt` and `affectedRanges`.

Build and run with Java 21:

```text
mvn test
mvn spring-boot:run
```

PostgreSQL schema creation is handled by the Flyway migration in `src/main/resources/db/migration`. Snapshot replacement is not exposed to ordinary editors. Checkpoints are server-generated and restore accepts only a target revision and reason; the server loads the historical snapshot and records a committed restore operation.

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
