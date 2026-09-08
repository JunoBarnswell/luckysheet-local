# Verification entrypoints

The repository has one build and verification chain. From the repository root, run:

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File scripts/build.ps1
powershell -NoProfile -ExecutionPolicy Bypass -File scripts/verify.ps1
```

The build requires Node `24.18.0`, Rust `1.97.1` with the `wasm32-unknown-unknown` target, Java 21, and Maven `3.9.9`. It builds `workbook-kernel-host` for the native host and `kernel_host.wasm` using Rust's raw WASM ABI. No `wasm-bindgen-cli` step is part of the chain. The WASM artifact is copied to `frontend-react/apps/web/public/kernel/` and checked against `kernel-manifest.json`.

## Excel corpus

`node scripts/verify-excel-corpus.mjs docs/verification/excel-corpus.manifest.json` validates a supplied corpus by SHA-256 and byte length. The manifest must declare `source: "real-excel-corpus"` and `synthetic: false`; generated or synthetic files are rejected. The repository intentionally does not include proprietary Excel files. A run without the manifest reports `BLOCKED` and exits with status 2. A reproducible run with a licensed corpus is:

```powershell
node scripts/verify-excel-corpus.mjs C:\path\to\excel-corpus.manifest.json
```

Desktop Excel interoperability remains blocked until the configured Excel executable and a real corpus are available. Browser tests do not substitute for that acceptance.
