create table native_document_task (
    task_id varchar(36) primary key,
    actor_subject varchar(500) not null,
    file_name varchar(500) not null,
    workbook_name varchar(500),
    space_id varchar(200),
    folder_id varchar(200),
    byte_length bigint not null,
    sha256 varchar(64),
    uploaded_bytes bigint not null,
    state varchar(16) not null,
    result_json longtext,
    error_code varchar(100),
    error_message longtext,
    created_at timestamp(6) not null,
    updated_at timestamp(6) not null,
    constraint native_document_task_length_ck check (byte_length between 1 and 1073741824),
    constraint native_document_task_upload_ck check (uploaded_bytes between 0 and byte_length),
    constraint native_document_task_state_ck check (state in ('uploading', 'importing', 'completed', 'cancelled', 'failed'))
);
create index native_document_task_actor_idx on native_document_task(actor_subject, updated_at);
