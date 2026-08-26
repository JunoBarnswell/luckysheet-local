create table workbook_asset (
    unit_id varchar(200) not null,
    asset_id varchar(200) not null,
    content_hash varchar(64) not null,
    mime_type varchar(127) not null,
    byte_length integer not null,
    width integer,
    height integer,
    content longblob not null,
    created_at datetime(6) not null,
    updated_at datetime(6) not null,
    primary key (unit_id, asset_id),
    constraint workbook_asset_workbook_fk foreign key (unit_id) references workbooks(unit_id) on delete cascade,
    constraint workbook_asset_hash_uk unique (unit_id, content_hash)
) engine=InnoDB;
