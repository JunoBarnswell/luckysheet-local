create table workbook_access_revision (
    unit_id varchar(200) not null primary key,
    revision bigint not null,
    constraint workbook_access_revision_workbook_fk foreign key (unit_id) references workbooks(unit_id) on delete cascade
) engine=InnoDB;

create table workbook_access_region (
    region_id varchar(200) not null primary key,
    unit_id varchar(200) not null,
    sheet_id varchar(200) not null,
    start_row integer not null,
    end_row integer not null,
    start_column integer not null,
    end_column integer not null,
    default_access varchar(16) not null,
    created_by varchar(500) not null,
    created_at datetime(6) not null,
    updated_at datetime(6) not null,
    constraint workbook_access_region_workbook_fk foreign key (unit_id) references workbooks(unit_id) on delete cascade
) engine=InnoDB;

create table workbook_access_grant (
    region_id varchar(200) not null,
    principal_kind varchar(16) not null,
    principal_id varchar(500) not null,
    access_level varchar(16) not null,
    primary key (region_id, principal_kind, principal_id),
    constraint workbook_access_grant_region_fk foreign key (region_id) references workbook_access_region(region_id) on delete cascade
) engine=InnoDB;

create index workbook_access_region_sheet_idx on workbook_access_region(unit_id, sheet_id, start_row, start_column);
