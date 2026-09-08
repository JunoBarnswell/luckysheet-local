-- Cloud-only catalog contract. Historical columns are removed at this explicit
-- migration boundary after the runtime DTO/entity clean break.
alter table workbooks drop column if exists storage_location;
alter table workbook_user_state drop column if exists default_create_location;
alter table workbook_user_state drop column if exists offline_cache;
alter table user_preference drop column if exists offline_cache;
