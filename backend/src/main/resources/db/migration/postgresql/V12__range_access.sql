create table workbook_access_revision (
    unit_id varchar(200) primary key references workbooks(unit_id) on delete cascade,
    revision bigint not null
);

create table workbook_access_region (
    region_id varchar(200) primary key,
    unit_id varchar(200) not null references workbooks(unit_id) on delete cascade,
    sheet_id varchar(200) not null,
    start_row integer not null,
    end_row integer not null,
    start_column integer not null,
    end_column integer not null,
    default_access varchar(16) not null,
    created_by varchar(500) not null,
    created_at timestamptz not null,
    updated_at timestamptz not null
);

create table workbook_access_grant (
    region_id varchar(200) not null references workbook_access_region(region_id) on delete cascade,
    principal_kind varchar(16) not null,
    principal_id varchar(500) not null,
    access_level varchar(16) not null,
    primary key (region_id, principal_kind, principal_id)
);

create index workbook_access_region_sheet_idx on workbook_access_region(unit_id, sheet_id, start_row, start_column);
