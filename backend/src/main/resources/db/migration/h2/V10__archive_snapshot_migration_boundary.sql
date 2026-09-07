-- Explicit migration boundary. Stop application writers before applying V10.
-- Archive tables are for the offline migration tool only; runtime has no readers.
create table workbook_v10_migration_archive (
    source_kind varchar(16) not null,
    unit_id varchar(200) not null,
    revision bigint not null,
    snapshot_revision bigint not null,
    snapshot_json text not null,
    source_checksum varchar(64) not null,
    original_checksum varchar(64),
    source_created_at timestamp with time zone not null,
    source_updated_at timestamp with time zone not null,
    archived_at timestamp with time zone not null,
    primary key (source_kind, unit_id, revision)
);
insert into workbook_v10_migration_archive
    (source_kind, unit_id, revision, snapshot_revision, snapshot_json, source_checksum,
     original_checksum, source_created_at, source_updated_at, archived_at)
select 'WORKBOOK', unit_id, revision, snapshot_revision, snapshot_json, lower(rawtohex(hash('SHA-256', stringtoutf8(snapshot_json)))),
       null, created_at, updated_at, current_timestamp
from workbooks;
insert into workbook_v10_migration_archive
    (source_kind, unit_id, revision, snapshot_revision, snapshot_json, source_checksum,
     original_checksum, source_created_at, source_updated_at, archived_at)
select 'CHECKPOINT', unit_id, revision, revision, snapshot_json, lower(rawtohex(hash('SHA-256', stringtoutf8(snapshot_json)))),
       checksum, created_at, created_at, current_timestamp
from snapshot_checkpoint;

-- Preserve the complete old operation input for source/history verification.
create table workbook_v10_operation_archive (
    operation_id varchar(200) primary key,
    unit_id varchar(200) not null,
    revision bigint not null,
    actor_subject varchar(500) not null,
    client_sequence bigint not null,
    base_revision bigint not null,
    envelope_json text not null,
    source_checksum varchar(64) not null,
    committed_at timestamp with time zone not null
);
insert into workbook_v10_operation_archive
    (operation_id, unit_id, revision, actor_subject, client_sequence, base_revision,
     envelope_json, source_checksum, committed_at)
select operation_id, unit_id, revision, actor_subject, client_sequence, base_revision,
       envelope_json, lower(rawtohex(hash('SHA-256', stringtoutf8(envelope_json)))), committed_at
from operation_log;

-- Only the offline importer may write these attestations, after native page and
-- complete archived checkpoint/operation history validation. A converter alone
-- must never claim pages_verified or history_verified.
create table workbook_v11_migration_proofs (
    unit_id varchar(200) primary key,
    proof_version int not null,
    source_revision bigint not null,
    source_snapshot_revision bigint not null,
    source_checksum varchar(64) not null,
    manifest_checksum varchar(64) not null,
    checkpoint_count bigint not null,
    operation_count bigint not null,
    history_checksum varchar(64) not null,
    pages_verified boolean not null,
    history_verified boolean not null,
    verified_at timestamp with time zone not null,
    constraint workbook_v11_proof_version_ck check (proof_version = 1),
    constraint workbook_v11_proof_counts_ck check (checkpoint_count >= 0 and operation_count >= 0)
);

