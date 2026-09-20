create table app_user (
    user_id varchar(36) primary key,
    username varchar(200) not null unique,
    password_hash varchar(255) not null,
    display_name varchar(200) not null,
    enabled boolean not null,
    admin boolean not null,
    credential_version bigint not null default 0,
    created_at timestamptz not null,
    updated_at timestamptz not null
);

create index app_user_enabled_idx on app_user(enabled);
