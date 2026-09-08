# Native Excel corpus acceptance manifest

This directory is the acceptance boundary for real Excel interoperability. A
runner should populate it with source files and record the SHA-256, Excel
version, feature set, expected sheet/cell counts, and the result of opening the
export in desktop Excel. The Rust codec must pass import, unchanged export,
and edited export checks for each applicable row.

| file | sha256 | format | features | import | unchanged export | edited export | Excel open |
|---|---|---|---|---|---|---|---|
| _(fixture required)_ | | xlsx/xlsm | | BLOCKED | BLOCKED | BLOCKED | BLOCKED |

No desktop Excel corpus is bundled in this repository, so these acceptance
rows remain blocked until real files and an Excel verification environment are
available. PDF/XPS are intentionally outside this manifest and rejected as
export-only presentation inputs.
