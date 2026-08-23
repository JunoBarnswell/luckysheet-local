create table workbooks (
    unit_id varchar(200) not null primary key,
    name varchar(500) not null,
    snapshot_json text not null,
    snapshot_revision bigint not null,
    revision bigint not null,
    created_at timestamp with time zone not null,
    updated_at timestamp with time zone not null
);

create table workbook_acl (
    unit_id varchar(200) not null,
    subject varchar(500) not null,
    role varchar(16) not null,
    created_at timestamp with time zone not null,
    updated_at timestamp with time zone not null,
    primary key (unit_id, subject)
);

create table workbook_share (
    share_id varchar(36) not null primary key,
    unit_id varchar(200) not null,
    token_hash varchar(128) not null unique,
    role varchar(16) not null,
    expires_at timestamp with time zone not null,
    revoked_at timestamp with time zone,
    created_by varchar(500) not null,
    created_at timestamp with time zone not null
);

create table operation_log (
    operation_id varchar(200) not null primary key,
    unit_id varchar(200) not null,
    revision bigint not null,
    actor_subject varchar(500) not null,
    client_sequence bigint not null,
    base_revision bigint not null,
    envelope_json text not null,
    committed_at timestamp with time zone not null,
    constraint operation_log_unit_actor_sequence_uk unique (unit_id, actor_subject, client_sequence)
);

create table snapshot_checkpoint (
    unit_id varchar(200) not null,
    revision bigint not null,
    snapshot_json text not null,
    checksum varchar(64) not null,
    created_at timestamp with time zone not null,
    primary key (unit_id, revision)
);

create table operation_audit (
    audit_id varchar(36) not null primary key,
    operation_id varchar(200),
    unit_id varchar(200),
    actor_subject varchar(500) not null,
    event_type varchar(100) not null,
    outcome varchar(16) not null,
    reason varchar(2000),
    details_json text not null,
    occurred_at timestamp with time zone not null
);

create table coordination_outbox (
    event_id varchar(36) not null primary key,
    unit_id varchar(200) not null,
    operation_id varchar(200) not null,
    revision bigint not null,
    payload_json text not null,
    created_at timestamp with time zone not null,
    published_at timestamp with time zone,
    lease_until timestamp with time zone,
    next_attempt_at timestamp with time zone not null,
    attempts integer not null,
    constraint coordination_outbox_unit_revision_uk unique (unit_id, revision)
);

create table workbook_data_block (
    unit_id varchar(200) not null,
    source_id varchar(200) not null,
    block_id varchar(200) not null,
    checksum varchar(64) not null,
    byte_length integer not null,
    content varbinary(33554432) not null,
    created_at timestamp with time zone not null,
    updated_at timestamp with time zone not null,
    primary key (unit_id, source_id, block_id)
);

create index workbook_acl_subject_idx on workbook_acl(subject, unit_id);
create index workbook_share_unit_idx on workbook_share(unit_id, expires_at);
create index workbook_share_active_token_idx on workbook_share(token_hash, revoked_at, expires_at);
create index operation_log_unit_revision_idx on operation_log(unit_id, revision);
create index snapshot_checkpoint_unit_revision_idx on snapshot_checkpoint(unit_id, revision);
create index operation_audit_unit_time_idx on operation_audit(unit_id, occurred_at);
create index coordination_outbox_pending_idx on coordination_outbox(next_attempt_at, created_at);
create index coordination_outbox_lease_idx on coordination_outbox(lease_until);
create index workbook_data_block_lookup_idx on workbook_data_block(unit_id, source_id, checksum);
