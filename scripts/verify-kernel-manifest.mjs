import { createHash } from 'node:crypto';
import { readFile } from 'node:fs/promises';
import path from 'node:path';
const repositoryRoot = path.resolve(import.meta.dirname, '..');
const directory = path.resolve(process.argv[2] ?? path.join(repositoryRoot, 'frontend-react/apps/web/public/kernel'));
const manifest = JSON.parse(await readFile(path.join(directory, 'kernel-manifest.json'), 'utf8'));
if (manifest.schema !== 'react-sheets.kernel-build.v1' || typeof manifest.expectedSha256 !== 'string' || !/^[a-f0-9]{64}$/.test(manifest.expectedSha256)) {
  throw new Error('Kernel manifest has an invalid build-bound schema or expectedSha256');
}
const bytes = await readFile(path.join(directory, manifest.artifact));
const sha256 = createHash('sha256').update(bytes).digest('hex');
if (manifest.bytes !== bytes.byteLength || manifest.expectedSha256 !== sha256) throw new Error('Kernel WASM manifest does not match artifact');
console.log(`Kernel manifest verified: ${sha256}`);
