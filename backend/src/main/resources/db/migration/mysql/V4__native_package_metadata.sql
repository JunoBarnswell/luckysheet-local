alter table workbook_source_artifact add column native_metadata_json longtext null;
update workbook_source_artifact set native_metadata_json = coalesce(native_metadata_json, detected_features_json, '{}');
alter table workbook_source_artifact modify native_metadata_json longtext not null;
alter table workbook_source_artifact drop column detected_features_json;
