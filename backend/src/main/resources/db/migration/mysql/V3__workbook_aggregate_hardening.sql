-- V1 creates the aggregate hardening columns, foreign keys, and personal
-- space invariant. Keep the common migration version in lockstep with H2 and
-- PostgreSQL so profile changes cannot silently skip a contract boundary.
select 1;
