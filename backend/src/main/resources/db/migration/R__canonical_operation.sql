CREATE TABLE IF NOT EXISTS workbooks (
    unit_id TEXT PRIMARY KEY,
    name TEXT NOT NULL,
    snapshot_json JSONB NOT NULL,
    snapshot_revision BIGINT NOT NULL DEFAULT 0,
    revision BIGINT NOT NULL DEFAULT 0,
    created_at TIMESTAMPTZ NOT NULL,
    updated_at TIMESTAMPTZ NOT NULL,
    CONSTRAINT workbooks_revision_nonnegative CHECK (revision >= 0),
    CONSTRAINT workbooks_snapshot_revision_nonnegative CHECK (snapshot_revision >= 0),
    CONSTRAINT workbooks_snapshot_revision_lte_revision CHECK (snapshot_revision <= revision)
);

CREATE TABLE IF NOT EXISTS workbook_acl (
    unit_id TEXT NOT NULL REFERENCES workbooks(unit_id) ON DELETE CASCADE,
    subject TEXT NOT NULL,
    role TEXT NOT NULL,
    created_at TIMESTAMPTZ NOT NULL,
    updated_at TIMESTAMPTZ NOT NULL,
    PRIMARY KEY (unit_id, subject),
    CONSTRAINT workbook_acl_role_valid CHECK (role IN ('OWNER', 'EDITOR', 'COMMENTER', 'VIEWER'))
);

CREATE INDEX IF NOT EXISTS workbook_acl_subject_idx ON workbook_acl(subject, unit_id);

CREATE TABLE IF NOT EXISTS operation_log (
    operation_id TEXT PRIMARY KEY,
    unit_id TEXT NOT NULL REFERENCES workbooks(unit_id) ON DELETE CASCADE,
    revision BIGINT NOT NULL,
    actor_subject TEXT NOT NULL,
    client_sequence BIGINT NOT NULL,
    base_revision BIGINT NOT NULL,
    envelope_json JSONB NOT NULL,
    committed_at TIMESTAMPTZ NOT NULL,
    UNIQUE (unit_id, revision),
    CONSTRAINT operation_revision_nonnegative CHECK (revision > 0),
    CONSTRAINT operation_base_revision_nonnegative CHECK (base_revision >= 0),
    CONSTRAINT operation_client_sequence_positive CHECK (client_sequence > 0)
);

CREATE INDEX IF NOT EXISTS operation_log_unit_revision_idx ON operation_log(unit_id, revision);

CREATE TABLE IF NOT EXISTS snapshot_checkpoint (
    unit_id TEXT NOT NULL REFERENCES workbooks(unit_id) ON DELETE CASCADE,
    revision BIGINT NOT NULL,
    snapshot_json JSONB NOT NULL,
    checksum CHAR(64) NOT NULL,
    created_at TIMESTAMPTZ NOT NULL,
    PRIMARY KEY (unit_id, revision)
);

CREATE TABLE IF NOT EXISTS operation_audit (
    audit_id UUID PRIMARY KEY,
    operation_id TEXT,
    unit_id TEXT,
    actor_subject TEXT NOT NULL,
    event_type TEXT NOT NULL,
    outcome TEXT NOT NULL,
    reason TEXT,
    details_json JSONB NOT NULL DEFAULT '{}'::jsonb,
    occurred_at TIMESTAMPTZ NOT NULL,
    CONSTRAINT operation_audit_outcome_valid CHECK (outcome IN ('ACCEPTED', 'REJECTED'))
);

CREATE INDEX IF NOT EXISTS operation_audit_unit_time_idx ON operation_audit(unit_id, occurred_at DESC);
