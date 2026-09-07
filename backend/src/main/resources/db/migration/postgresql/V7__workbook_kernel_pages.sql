create table workbook_manifests (
    manifest_id varchar(64) primary key,
    unit_id varchar(200) not null,
    revision bigint not null,
    manifest_version int not null,
    manifest_json text not null,
    checksum varchar(64) not null,
    created_at timestamp with time zone not null,
    constraint workbook_manifests_workbook_fk foreign key (unit_id) references workbooks(unit_id) on delete cascade
);
create unique index workbook_manifest_revision_idx on workbook_manifests(unit_id, revision);
create table workbook_pages (
    page_id varchar(64) primary key,
    unit_id varchar(200) not null,
    checksum varchar(64) not null,
    byte_length bigint not null,
    payload_base64 text not null,
    created_at timestamp with time zone not null,
    constraint workbook_pages_workbook_fk foreign key (unit_id) references workbooks(unit_id) on delete cascade
);
create unique index workbook_page_content_idx on workbook_pages(unit_id, checksum);
