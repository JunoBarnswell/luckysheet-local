create table workbook_kernel_history (
    history_id varchar(64) primary key,
    unit_id varchar(200) not null,
    operation_id varchar(200) not null,
    base_revision bigint not null,
    revision bigint not null,
    history_json longtext not null,
    result_metadata_json longtext not null,
    checksum varchar(64) not null,
    created_at timestamp(6) not null,
    constraint workbook_history_workbook_fk foreign key (unit_id) references workbooks(unit_id) on delete cascade
);
create unique index workbook_history_operation_idx on workbook_kernel_history(unit_id, operation_id);
