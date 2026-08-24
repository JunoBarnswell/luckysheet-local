alter table workbook_source_artifact add column if not exists native_metadata_json text;
update workbook_source_artifact set native_metadata_json = coalesce(native_metadata_json, detected_features_json, '{}');
alter table workbook_source_artifact alter column native_metadata_json set not null;
alter table workbook_source_artifact drop column if exists detected_features_json;
