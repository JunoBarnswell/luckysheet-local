create table workbook_query_execution (
    execution_token varchar(36) primary key,
    unit_id varchar(200) not null,
    query_id varchar(200) not null,
    actor_subject varchar(500) not null,
    source_revision bigint not null,
    status varchar(16) not null,
    result_hash varchar(64),
    result_json text,
    created_at timestamp with time zone not null,
    expires_at timestamp with time zone not null,
    constraint workbook_query_execution_query_uq unique (unit_id, query_id),
    constraint workbook_query_execution_workbook_fk foreign key (unit_id) references workbooks(unit_id) on delete cascade
);
-- Native artifacts from v10 remain an offline migration input only.
-- Runtime never interprets legacy content or guesses its revision binding.
alter table workbook_source_artifact rename to workbook_native_artifact_migration_archive;
create table workbook_source_artifact (
    unit_id varchar(200) primary key,
    file_name varchar(500) not null,
    mime_type varchar(200) not null,
    checksum varchar(64) not null,
    workbook_revision bigint not null,
    byte_length bigint not null,
    storage_path varchar(2000) not null,
    native_metadata_json text not null,
    created_at timestamptz not null,
    updated_at timestamptz not null
);
create index workbook_native_artifact_sha_idx on workbook_source_artifact(checksum);
