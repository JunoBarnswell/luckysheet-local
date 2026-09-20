create table app_user (
    user_id varchar(36) not null primary key,
    username varchar(200) not null unique,
    password_hash varchar(255) not null,
    display_name varchar(200) not null,
    enabled boolean not null,
    admin boolean not null,
    credential_version bigint not null default 0,
    created_at datetime(6) not null,
    updated_at datetime(6) not null
) engine=InnoDB;

create index app_user_enabled_idx on app_user(enabled);
