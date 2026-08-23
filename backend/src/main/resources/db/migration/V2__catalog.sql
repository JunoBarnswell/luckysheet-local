alter table workbooks add column if not exists owner_subject varchar(500) not null default '';
alter table workbooks add column if not exists space_id varchar(200);
alter table workbooks add column if not exists folder_id varchar(200);
alter table workbooks add column if not exists storage_location varchar(16) not null default 'REMOTE';
alter table workbooks add column if not exists source varchar(32) not null default 'NATIVE';
alter table workbooks add column if not exists lifecycle varchar(16) not null default 'ACTIVE';
alter table workbooks add column if not exists deleted_at timestamp with time zone;

-- Existing workbooks already have an OWNER row in workbook_acl on supported
-- deployments. Reuse that proven identity instead of inventing ownership.
update workbooks
set owner_subject = coalesce((
    select min(subject)
    from workbook_acl
    where workbook_acl.unit_id = workbooks.unit_id
      and workbook_acl.role = 'OWNER'
), owner_subject)
where owner_subject = '';

create table workspace_space (
    space_id varchar(200) not null primary key,
    name varchar(200) not null,
    type varchar(16) not null,
    owner_subject varchar(500) not null,
    created_at timestamp with time zone not null,
    updated_at timestamp with time zone not null
);

create table workspace_folder (
    folder_id varchar(200) not null primary key,
    space_id varchar(200) not null,
    parent_id varchar(200),
    name varchar(200) not null,
    created_by varchar(500) not null,
    created_at timestamp with time zone not null,
    updated_at timestamp with time zone not null
);

create table space_member (
    space_id varchar(200) not null,
    subject varchar(500) not null,
    role varchar(16) not null,
    created_at timestamp with time zone not null,
    updated_at timestamp with time zone not null,
    primary key (space_id, subject)
);

create table workbook_user_state (
    unit_id varchar(200) not null,
    subject varchar(500) not null,
    favorite boolean not null default false,
    last_opened_at timestamp with time zone,
    auto_save boolean not null default true,
    auto_sync boolean not null default true,
    default_create_location varchar(16) not null default 'remote',
    import_compatibility_level varchar(16) not null default 'standard',
    language varchar(32),
    offline_cache boolean not null default true,
    theme varchar(16) not null default 'system',
    updated_at timestamp with time zone not null,
    primary key (unit_id, subject)
);

create table user_preference (
    subject varchar(500) not null primary key,
    default_space_id varchar(200),
    default_folder_id varchar(200),
    auto_save boolean not null default true,
    auto_sync boolean not null default true,
    offline_cache boolean not null default true,
    import_compatibility varchar(32) not null default 'warn',
    language varchar(32),
    theme varchar(32),
    updated_at timestamp with time zone not null
);

create table workbook_source_artifact (
    unit_id varchar(200) not null primary key,
    file_name varchar(500) not null,
    mime_type varchar(200) not null,
    checksum varchar(64) not null,
    byte_length bigint not null,
    content varbinary(52428800) not null,
    detected_features_json text not null,
    created_at timestamp with time zone not null,
    updated_at timestamp with time zone not null
);

-- Backfill every legacy owner into an isolated personal root. A per-workbook
-- root is deliberate: it preserves access without guessing a subject-wide
-- folder hierarchy; owners can consolidate folders after upgrade.
insert into workspace_space (space_id, name, type, owner_subject, created_at, updated_at)
select concat('legacy-', unit_id), '我的云文档', 'PERSONAL', owner_subject, current_timestamp, current_timestamp
from workbooks
where owner_subject <> ''
  and space_id is null
  and not exists (
    select 1 from workspace_space
    where workspace_space.space_id = concat('legacy-', workbooks.unit_id)
  );

insert into space_member (space_id, subject, role, created_at, updated_at)
select concat('legacy-', unit_id), owner_subject, 'OWNER', current_timestamp, current_timestamp
from workbooks
where owner_subject <> ''
  and space_id is null
  and not exists (
    select 1 from space_member
    where space_member.space_id = concat('legacy-', workbooks.unit_id)
      and space_member.subject = workbooks.owner_subject
  );

update workbooks
set space_id = concat('legacy-', unit_id)
where owner_subject <> ''
  and space_id is null;

create index workbooks_owner_updated_idx on workbooks(owner_subject, updated_at);
create index workbooks_space_folder_idx on workbooks(space_id, folder_id, deleted_at);
create index workspace_space_owner_idx on workspace_space(owner_subject, updated_at);
create index workspace_folder_space_parent_idx on workspace_folder(space_id, parent_id, name);
create index space_member_subject_idx on space_member(subject, space_id);
create index workbook_user_state_subject_idx on workbook_user_state(subject, unit_id);
create index workbook_source_artifact_checksum_idx on workbook_source_artifact(checksum);
