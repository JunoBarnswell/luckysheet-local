import { readFile } from 'node:fs/promises';
import { readdir } from 'node:fs/promises';
import { join } from 'node:path';

const forbidden = [
  '@univerjs',
  '@univerjs-pro',
  'frontend/src',
  'jquery',
  'jfrefreshgrid',
  'controlHistory',
  'new Function',
  'eval(',
];

async function walk(directory) {
  const entries = await readdir(directory, { withFileTypes: true });
  const files = [];
  for (const entry of entries) {
    if (entry.name === 'node_modules' || entry.name === 'dist') continue;
    const target = join(directory, entry.name);
    if (entry.isDirectory()) files.push(...(await walk(target)));
    else if (/\.(ts|tsx|js|mjs|json)$/.test(entry.name)) files.push(target);
  }
  return files;
}

const files = [
  ...(await walk(join(process.cwd(), 'apps'))),
  ...(await walk(join(process.cwd(), 'packages'))),
];
const violations = [];
for (const file of files) {
  const source = await readFile(file, 'utf8');
  for (const token of forbidden) {
    if (source.includes(token)) violations.push(`${file}: ${token}`);
  }
}

if (violations.length) {
  console.error(violations.join('\n'));
  process.exitCode = 1;
} else {
  console.log(`React boundary check passed: ${files.length} source files scanned.`);
}
