CREATE TABLE IF NOT EXISTS coordination_outbox (
    event_id UUID PRIMARY KEY,
    unit_id TEXT NOT NULL REFERENCES workbooks(unit_id) ON DELETE CASCADE,
    operation_id TEXT NOT NULL,
    revision BIGINT NOT NULL,
    payload_json JSONB NOT NULL,
    created_at TIMESTAMPTZ NOT NULL,
    published_at TIMESTAMPTZ,
    lease_until TIMESTAMPTZ,
    next_attempt_at TIMESTAMPTZ NOT NULL,
    attempts INTEGER NOT NULL DEFAULT 0,
    CONSTRAINT coordination_outbox_revision_positive CHECK (revision > 0),
    CONSTRAINT coordination_outbox_attempts_nonnegative CHECK (attempts >= 0),
    CONSTRAINT coordination_outbox_unit_revision_unique UNIQUE (unit_id, revision)
);

CREATE INDEX IF NOT EXISTS coordination_outbox_pending_idx
    ON coordination_outbox(next_attempt_at, created_at)
    WHERE published_at IS NULL;

CREATE INDEX IF NOT EXISTS coordination_outbox_lease_idx
    ON coordination_outbox(lease_until)
    WHERE published_at IS NULL;

CREATE TABLE IF NOT EXISTS workbook_share (
    share_id UUID PRIMARY KEY,
    unit_id TEXT NOT NULL REFERENCES workbooks(unit_id) ON DELETE CASCADE,
    token_hash TEXT NOT NULL UNIQUE,
    role TEXT NOT NULL,
    expires_at TIMESTAMPTZ NOT NULL,
    revoked_at TIMESTAMPTZ,
    created_by TEXT NOT NULL,
    created_at TIMESTAMPTZ NOT NULL,
    CONSTRAINT workbook_share_role_valid CHECK (role IN ('VIEWER', 'COMMENTER', 'EDITOR')),
    CONSTRAINT workbook_share_expiry_valid CHECK (expires_at > created_at)
);

CREATE INDEX IF NOT EXISTS workbook_share_unit_idx ON workbook_share(unit_id, expires_at);
CREATE INDEX IF NOT EXISTS workbook_share_active_token_idx ON workbook_share(token_hash, revoked_at, expires_at);
