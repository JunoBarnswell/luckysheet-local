import { createHash } from 'node:crypto';
import { readFile, writeFile } from 'node:fs/promises';
import path from 'node:path';

const directory = path.resolve(process.argv[2] ?? 'frontend-react/apps/web/public/kernel');
const fileName = 'kernel_host.wasm';
const bytes = await readFile(path.join(directory, fileName));
const manifest = { schemaVersion: 1, artifact: fileName, bytes: bytes.byteLength, sha256: createHash('sha256').update(bytes).digest('hex'), source: 'cargo build -p kernel-host --target wasm32-unknown-unknown --release' };
await writeFile(path.join(directory, 'kernel-manifest.json'), `${JSON.stringify(manifest, null, 2)}\n`, 'utf8');
console.log(`Kernel WASM manifest: ${manifest.sha256}`);
