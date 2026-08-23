create table workbooks (
    unit_id varchar(200) not null primary key,
    name varchar(255) not null,
    snapshot_json longtext not null,
    snapshot_revision bigint not null,
    revision bigint not null,
    entity_version bigint not null default 0,
    created_at datetime(6) not null,
    updated_at datetime(6) not null,
    owner_subject varchar(500) not null,
    space_id varchar(200),
    folder_id varchar(200),
    storage_location varchar(16) not null default 'REMOTE',
    source varchar(32) not null default 'NATIVE',
    lifecycle varchar(16) not null default 'ACTIVE',
    deleted_at datetime(6)
) engine=InnoDB;

create table workbook_acl (
    unit_id varchar(200) not null,
    subject varchar(500) not null,
    role varchar(16) not null,
    created_at datetime(6) not null,
    updated_at datetime(6) not null,
    primary key (unit_id, subject)
) engine=InnoDB;

create table workspace_space (
    space_id varchar(200) not null primary key,
    name varchar(200) not null,
    type varchar(16) not null,
    owner_subject varchar(500) not null,
    personal_owner_subject varchar(500) generated always as
        (case when type = 'PERSONAL' then owner_subject else null end) stored,
    created_at datetime(6) not null,
    updated_at datetime(6) not null,
    unique key workspace_space_single_personal_owner_idx (personal_owner_subject)
) engine=InnoDB;

create table workspace_folder (
    folder_id varchar(200) not null primary key,
    space_id varchar(200) not null,
    parent_id varchar(200),
    name varchar(200) not null,
    created_by varchar(500) not null,
    created_at datetime(6) not null,
    updated_at datetime(6) not null
) engine=InnoDB;

create table space_member (
    space_id varchar(200) not null,
    subject varchar(500) not null,
    role varchar(16) not null,
    created_at datetime(6) not null,
    updated_at datetime(6) not null,
    primary key (space_id, subject)
) engine=InnoDB;

create table workbook_user_state (
    unit_id varchar(200) not null,
    subject varchar(500) not null,
    favorite boolean not null default false,
    last_opened_at datetime(6),
    auto_save boolean not null default true,
    auto_sync boolean not null default true,
    default_create_location varchar(16) not null default 'remote',
    import_compatibility_level varchar(16) not null default 'standard',
    language varchar(32),
    offline_cache boolean not null default true,
    theme varchar(16) not null default 'system',
    updated_at datetime(6) not null,
    primary key (unit_id, subject)
) engine=InnoDB;

create table user_preference (
    subject varchar(500) not null primary key,
    default_space_id varchar(200),
    default_folder_id varchar(200),
    auto_save boolean not null default true,
    auto_sync boolean not null default true,
    offline_cache boolean not null default true,
    import_compatibility varchar(32) not null default 'B',
    language varchar(32),
    theme varchar(32) not null default 'system',
    updated_at datetime(6) not null,
    entity_version bigint not null default 0
) engine=InnoDB;

create table workbook_source_artifact (
    unit_id varchar(200) not null primary key,
    file_name varchar(500) not null,
    mime_type varchar(200) not null,
    checksum varchar(64) not null,
    byte_length bigint not null,
    content longblob not null,
    detected_features_json longtext not null,
    created_at datetime(6) not null,
    updated_at datetime(6) not null
) engine=InnoDB;

create table workbook_share (
    share_id varchar(36) not null primary key,
    unit_id varchar(200) not null,
    token_hash varchar(128) not null unique,
    role varchar(16) not null,
    expires_at datetime(6) not null,
    revoked_at datetime(6),
    created_by varchar(500) not null,
    created_at datetime(6) not null
) engine=InnoDB;

create table operation_log (
    operation_id varchar(200) not null primary key,
    unit_id varchar(200) not null,
    revision bigint not null,
    actor_subject varchar(500) not null,
    client_sequence bigint not null,
    base_revision bigint not null,
    envelope_json longtext not null,
    committed_at datetime(6) not null,
    unique key operation_log_unit_actor_sequence_uk (unit_id, actor_subject, client_sequence),
    unique key operation_log_unit_revision_uk (unit_id, revision)
) engine=InnoDB;

create table snapshot_checkpoint (
    unit_id varchar(200) not null,
    revision bigint not null,
    snapshot_json longtext not null,
    checksum varchar(64) not null,
    created_at datetime(6) not null,
    primary key (unit_id, revision)
) engine=InnoDB;

create table operation_audit (
    audit_id varchar(36) not null primary key,
    operation_id varchar(200),
    unit_id varchar(200),
    actor_subject varchar(500) not null,
    event_type varchar(100) not null,
    outcome varchar(16) not null,
    reason varchar(2000),
    details_json longtext not null,
    occurred_at datetime(6) not null
) engine=InnoDB;

create table coordination_outbox (
    event_id varchar(36) not null primary key,
    unit_id varchar(200) not null,
    operation_id varchar(200) not null,
    revision bigint not null,
    payload_json longtext not null,
    created_at datetime(6) not null,
    published_at datetime(6),
    lease_until datetime(6),
    next_attempt_at datetime(6) not null,
    attempts integer not null,
    unique key coordination_outbox_unit_revision_uk (unit_id, revision)
) engine=InnoDB;

create table workbook_data_block (
    unit_id varchar(200) not null,
    source_id varchar(200) not null,
    block_id varchar(200) not null,
    checksum varchar(64) not null,
    byte_length integer not null,
    content longblob not null,
    created_at datetime(6) not null,
    updated_at datetime(6) not null,
    primary key (unit_id, source_id, block_id)
) engine=InnoDB;

alter table workspace_folder add constraint workspace_folder_space_fk
    foreign key (space_id) references workspace_space(space_id) on delete restrict;
alter table workspace_folder add constraint workspace_folder_parent_fk
    foreign key (parent_id) references workspace_folder(folder_id) on delete restrict;
alter table space_member add constraint space_member_space_fk
    foreign key (space_id) references workspace_space(space_id) on delete cascade;
alter table workbooks add constraint workbooks_space_fk
    foreign key (space_id) references workspace_space(space_id) on delete restrict;
alter table workbooks add constraint workbooks_folder_fk
    foreign key (folder_id) references workspace_folder(folder_id) on delete restrict;
alter table workbook_acl add constraint workbook_acl_workbook_fk
    foreign key (unit_id) references workbooks(unit_id) on delete cascade;
alter table workbook_user_state add constraint workbook_user_state_workbook_fk
    foreign key (unit_id) references workbooks(unit_id) on delete cascade;
alter table workbook_source_artifact add constraint workbook_source_artifact_workbook_fk
    foreign key (unit_id) references workbooks(unit_id) on delete cascade;
alter table workbook_share add constraint workbook_share_workbook_fk
    foreign key (unit_id) references workbooks(unit_id) on delete cascade;
alter table operation_log add constraint operation_log_workbook_fk
    foreign key (unit_id) references workbooks(unit_id) on delete cascade;
alter table snapshot_checkpoint add constraint snapshot_checkpoint_workbook_fk
    foreign key (unit_id) references workbooks(unit_id) on delete cascade;
alter table coordination_outbox add constraint coordination_outbox_workbook_fk
    foreign key (unit_id) references workbooks(unit_id) on delete cascade;
alter table workbook_data_block add constraint workbook_data_block_workbook_fk
    foreign key (unit_id) references workbooks(unit_id) on delete cascade;

create index workbooks_owner_updated_idx on workbooks(owner_subject, updated_at);
create index workbooks_space_folder_idx on workbooks(space_id, folder_id, deleted_at);
create index workspace_space_owner_idx on workspace_space(owner_subject, updated_at);
create index workspace_folder_space_parent_idx on workspace_folder(space_id, parent_id, name);
create index space_member_subject_idx on space_member(subject, space_id);
create index workbook_acl_subject_idx on workbook_acl(subject, unit_id);
create index workbook_share_unit_idx on workbook_share(unit_id, expires_at);
create index workbook_share_active_token_idx on workbook_share(token_hash, revoked_at, expires_at);
create index operation_log_unit_revision_idx on operation_log(unit_id, revision);
create index snapshot_checkpoint_unit_revision_idx on snapshot_checkpoint(unit_id, revision);
create index operation_audit_unit_time_idx on operation_audit(unit_id, occurred_at);
create index coordination_outbox_pending_idx on coordination_outbox(next_attempt_at, created_at);
create index coordination_outbox_lease_idx on coordination_outbox(lease_until);
create index workbook_data_block_lookup_idx on workbook_data_block(unit_id, source_id, checksum);
