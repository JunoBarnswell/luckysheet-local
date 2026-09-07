import { createHash } from 'node:crypto';
import { access, readFile } from 'node:fs/promises';
import path from 'node:path';

const manifestPath = path.resolve(process.argv[2] ?? 'docs/verification/excel-corpus.manifest.json');
try { await access(manifestPath); } catch { console.error(`BLOCKED: Excel corpus manifest is unavailable: ${manifestPath}`); process.exitCode = 2; }
if (process.exitCode === 2) process.exit();
const manifest = JSON.parse(await readFile(manifestPath, 'utf8'));
if (manifest.synthetic === true || manifest.source !== 'real-excel-corpus') throw new Error('Excel corpus must declare source=real-excel-corpus and synthetic=false');
const base = path.dirname(manifestPath);
for (const entry of manifest.files ?? []) {
  const filePath = path.resolve(base, entry.path);
  const bytes = await readFile(filePath);
  const hash = createHash('sha256').update(bytes).digest('hex');
  if (hash !== entry.sha256 || bytes.byteLength !== entry.bytes) throw new Error(`Excel corpus mismatch: ${entry.path}`);
  console.log(`verified ${entry.path} (${entry.bytes} bytes)`);
}
if (!Array.isArray(manifest.files) || manifest.files.length === 0) throw new Error('Excel corpus manifest contains no real files');
