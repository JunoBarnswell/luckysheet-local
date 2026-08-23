alter table workbooks add column if not exists entity_version bigint not null default 0;
alter table user_preference add column if not exists entity_version bigint not null default 0;

update workspace_folder set space_id = (
    select min(winner.space_id)
    from workspace_space legacy
    join workspace_space winner on winner.type = 'PERSONAL' and winner.owner_subject = legacy.owner_subject
    where legacy.space_id = workspace_folder.space_id and legacy.type = 'PERSONAL'
)
where exists (
    select 1 from workspace_space legacy
    where legacy.space_id = workspace_folder.space_id and legacy.type = 'PERSONAL'
);

update workbooks set space_id = (
    select min(candidate.space_id) from workspace_space candidate
    where candidate.type = 'PERSONAL' and candidate.owner_subject = workbooks.owner_subject
)
where owner_subject <> '' and exists (
    select 1 from workspace_space candidate
    where candidate.type = 'PERSONAL' and candidate.owner_subject = workbooks.owner_subject
);

merge into space_member (space_id, subject, role, created_at, updated_at)
key(space_id, subject)
select min(winner.space_id), member.subject,
       case max(case member.role when 'OWNER' then 4 when 'EDITOR' then 3 when 'COMMENTER' then 2 else 1 end)
           when 4 then 'OWNER' when 3 then 'EDITOR' when 2 then 'COMMENTER' else 'VIEWER' end,
       min(member.created_at), max(member.updated_at)
from space_member member
join workspace_space legacy on legacy.space_id = member.space_id and legacy.type = 'PERSONAL'
join workspace_space winner on winner.type = 'PERSONAL' and winner.owner_subject = legacy.owner_subject
group by legacy.owner_subject, member.subject;

delete from space_member where space_id in (
    select duplicate.space_id from workspace_space duplicate
    where duplicate.type = 'PERSONAL' and exists (
        select 1 from workspace_space winner
        where winner.type = 'PERSONAL'
          and winner.owner_subject = duplicate.owner_subject
          and winner.space_id < duplicate.space_id
    )
);

delete from workspace_space where type = 'PERSONAL' and exists (
    select 1 from workspace_space winner
    where winner.type = 'PERSONAL'
      and winner.owner_subject = workspace_space.owner_subject
      and winner.space_id < workspace_space.space_id
);

insert into space_member (space_id, subject, role, created_at, updated_at)
select space.space_id, space.owner_subject, 'OWNER', current_timestamp, current_timestamp
from workspace_space space
where space.type = 'PERSONAL' and not exists (
    select 1 from space_member member
    where member.space_id = space.space_id and member.subject = space.owner_subject
);

create index if not exists workbooks_folder_reference_idx on workbooks(folder_id, space_id);
create unique index if not exists operation_log_unit_revision_unique_idx on operation_log(unit_id, revision);

alter table workspace_folder add constraint workspace_folder_space_fk foreign key (space_id)
    references workspace_space(space_id) on delete restrict;
alter table workspace_folder add constraint workspace_folder_parent_fk foreign key (parent_id)
    references workspace_folder(folder_id) on delete restrict;
alter table space_member add constraint space_member_space_fk foreign key (space_id)
    references workspace_space(space_id) on delete cascade;
alter table workbooks add constraint workbooks_space_fk foreign key (space_id)
    references workspace_space(space_id) on delete restrict;
alter table workbooks add constraint workbooks_folder_fk foreign key (folder_id)
    references workspace_folder(folder_id) on delete restrict;
alter table workbook_acl add constraint workbook_acl_workbook_fk foreign key (unit_id)
    references workbooks(unit_id) on delete cascade;
alter table workbook_user_state add constraint workbook_user_state_workbook_fk foreign key (unit_id)
    references workbooks(unit_id) on delete cascade;
alter table workbook_source_artifact add constraint workbook_source_artifact_workbook_fk foreign key (unit_id)
    references workbooks(unit_id) on delete cascade;
alter table workbook_share add constraint workbook_share_workbook_fk foreign key (unit_id)
    references workbooks(unit_id) on delete cascade;
alter table operation_log add constraint operation_log_workbook_fk foreign key (unit_id)
    references workbooks(unit_id) on delete cascade;
alter table snapshot_checkpoint add constraint snapshot_checkpoint_workbook_fk foreign key (unit_id)
    references workbooks(unit_id) on delete cascade;
alter table coordination_outbox add constraint coordination_outbox_workbook_fk foreign key (unit_id)
    references workbooks(unit_id) on delete cascade;
alter table workbook_data_block add constraint workbook_data_block_workbook_fk foreign key (unit_id)
    references workbooks(unit_id) on delete cascade;

alter table workspace_space add column personal_owner_subject varchar(500)
    as (case when type = 'PERSONAL' then owner_subject else null end);
create unique index workspace_space_single_personal_owner_idx on workspace_space(personal_owner_subject);
