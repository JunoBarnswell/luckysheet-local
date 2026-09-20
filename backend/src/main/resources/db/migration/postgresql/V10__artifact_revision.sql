-- Existing files remain unbound: their export revision cannot be inferred safely.
ALTER TABLE workbook_source_artifact ADD COLUMN source_revision BIGINT;
