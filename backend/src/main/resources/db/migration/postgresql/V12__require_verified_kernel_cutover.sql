-- CUTOVER GATE: run only after V10, offline conversion/import, and proof validation.
-- CHECK failure deliberately aborts before any destructive DDL. On MySQL V10
-- archives remain durable even if V12 fails; retain them and the external backup.
create table if not exists workbook_v11_cutover_guard (
    unverified_sources bigint not null,
    constraint workbook_v11_migration_required_ck check (unverified_sources = 0)
);
delete from workbook_v11_cutover_guard;
insert into workbook_v11_cutover_guard (unverified_sources)
select
    (select count(*) from workbooks w where not exists (
        select 1 from workbook_v10_migration_archive a
        join workbook_v11_migration_proofs p on p.unit_id = a.unit_id
        join workbook_manifests m on m.unit_id = p.unit_id and m.revision = p.source_revision
        where a.source_kind = 'WORKBOOK' and a.unit_id = w.unit_id
          and a.revision = w.revision and a.snapshot_revision = w.snapshot_revision
          and a.source_checksum = encode(sha256(convert_to(w.snapshot_json, 'UTF8')), 'hex')
          and a.source_checksum = encode(sha256(convert_to(a.snapshot_json, 'UTF8')), 'hex')
          and p.proof_version = 1 and p.source_revision = a.revision
          and p.source_snapshot_revision = a.snapshot_revision
          and p.source_checksum = a.source_checksum
          and m.manifest_version = 11 and m.checksum = p.manifest_checksum
          and m.checksum = encode(sha256(convert_to(m.manifest_json, 'UTF8')), 'hex')
          and p.pages_verified = true and p.history_verified = true
          and length(p.history_checksum) = 64
          and p.checkpoint_count = (select count(*) from workbook_v10_migration_archive c
              where c.source_kind = 'CHECKPOINT' and c.unit_id = w.unit_id)
          and p.operation_count = (select count(*) from workbook_v10_operation_archive o where o.unit_id = w.unit_id)
    ))
    + (select count(*) from snapshot_checkpoint c where not exists (
        select 1 from workbook_v10_migration_archive a
        where a.source_kind = 'CHECKPOINT' and a.unit_id = c.unit_id and a.revision = c.revision
          and a.source_checksum = encode(sha256(convert_to(c.snapshot_json, 'UTF8')), 'hex')
          and a.source_checksum = encode(sha256(convert_to(a.snapshot_json, 'UTF8')), 'hex')
          and a.original_checksum = c.checksum and a.source_created_at = c.created_at
    ))
    + (select count(*) from operation_log o where not exists (
        select 1 from workbook_v10_operation_archive a where a.operation_id = o.operation_id
          and a.unit_id = o.unit_id and a.revision = o.revision
          and a.actor_subject = o.actor_subject and a.client_sequence = o.client_sequence
          and a.base_revision = o.base_revision and a.committed_at = o.committed_at
          and a.source_checksum = encode(sha256(convert_to(o.envelope_json, 'UTF8')), 'hex')
          and a.source_checksum = encode(sha256(convert_to(a.envelope_json, 'UTF8')), 'hex')
    ))
    + (select count(*) from workbook_native_artifact_migration_archive a where not exists (
        select 1 from workbook_source_artifact n
        join workbooks w on w.unit_id = n.unit_id
        join workbook_manifests m on m.unit_id = n.unit_id and m.revision = n.workbook_revision
        where n.unit_id = a.unit_id and n.workbook_revision = w.revision
          and m.manifest_version = 11 and n.checksum = a.checksum
          and a.checksum = encode(sha256(a.content), 'hex')
          and n.byte_length = a.byte_length and a.byte_length = octet_length(a.content)
          and n.storage_path is not null and length(trim(n.storage_path)) > 0
    ));

-- No destructive statement may be moved above the guard insertion.
alter table workbooks drop column snapshot_json;
alter table workbooks drop column snapshot_revision;
drop table snapshot_checkpoint;
drop table workbook_v11_cutover_guard;
-- Keep archives and proofs permanently until a separately authorized retention
-- decision. They are not a runtime fallback, and rollback requires restoring the
-- external pre-upgrade database backup (particularly on nontransactional DDL).
