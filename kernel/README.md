# Shared spreadsheet kernel

The Rust workspace owns workbook semantics for both native and WASM hosts. The host boundary uses protocol version 1 and workbook manifest version 11. Core coordinates are zero-based; ranges are inclusive. Page dimensions are 1024 rows by 32 columns. `Scalar` serializes as null, boolean, finite number, string, or the canonical `{kind:"error",code,message}` object. `Cell` retains authored metadata without creating a second interpreted model.

`CellReader` is revision-pinned. An absent authored cell is `Ok(None)`; an unavailable page is `Err(DATA_PAGE_UNAVAILABLE)`. Visibility never modifies this distinction. All public failures return `KernelError {code,message,object,recovery}`.

Module ownership: core (model/pages/transaction) and host (dispatch/transport) are coordinated centrally; formula, analytics, geometry, and native-document own their separate semantic domains. No production host may route a failed Rust call to a TypeScript or Java semantic implementation.

The implementation and acceptance ledger is `frontend-react/docs/web-excel-shared-kernel-prd.md`. A crate existing or compiling is not evidence that its production consumers have migrated.
