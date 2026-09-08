-- Cloud-only catalog contract. Historical columns are removed at this explicit
-- migration boundary after the runtime DTO/entity clean break.
alter table workbooks drop column storage_location;
alter table workbook_user_state drop column default_create_location;
alter table workbook_user_state drop column offline_cache;
alter table user_preference drop column offline_cache;
