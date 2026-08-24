import { readFile } from 'node:fs/promises';
import { resolve } from 'node:path';

const root = process.cwd();
const contracts = JSON.parse(await readFile(resolve(root, '../contracts/workbook-contract.json'), 'utf8'));
const snapshotSchema = JSON.parse(await readFile(resolve(root, '../contracts/workbook-snapshot.schema.json'), 'utf8'));
const javaRegistry = await readFile(resolve(root, '../backend/src/main/java/com/xc/luckysheet/server/mutation/MutationDescriptorRegistry.java'), 'utf8');
const generatedTypeScript = await readFile(resolve(root, 'packages/protocol/src/generated-contract.ts'), 'utf8');
const generatedJava = await readFile(resolve(root, '../backend/src/main/java/com/xc/luckysheet/server/contract/GeneratedWorkbookContract.java'), 'utf8');
const protocol = await readFile(resolve(root, 'packages/protocol/src/index.ts'), 'utf8');
const catalog = await readFile(resolve(root, 'packages/spreadsheet-app/src/features/workbook-catalog/service.ts'), 'utf8');
const hub = await readFile(resolve(root, 'apps/web/src/containers/WorkbookHubContainer.tsx'), 'utf8');
const frontend = await Promise.all([
  readFile(resolve(root, 'packages/sheet-features/src/index.ts'), 'utf8'),
  readFile(resolve(root, 'packages/sheet-features/src/editing/index.ts'), 'utf8'),
  readFile(resolve(root, 'packages/spreadsheet-app/src/features/review/commands.ts'), 'utf8'),
]);
const violations = [];
if (snapshotSchema?.properties?.version?.const !== contracts.workbook.snapshotVersion) violations.push('WorkbookSnapshot JSON Schema version is out of sync with workbook-contract.json');
if (!generatedTypeScript.includes(`WORKBOOK_CONTRACT_API_VERSION = ${JSON.stringify(contracts.apiVersion)}`)
  || !generatedJava.includes(`API_VERSION = ${JSON.stringify(contracts.apiVersion)}`)) {
  violations.push('generated contract outputs are stale; run npm run generate:contracts');
}
for (const required of ['listWorkbookPage', 'listRevisionPage', 'validateUserPreferences', 'validateCursorPage']) {
  if (!protocol.includes(required)) violations.push(`protocol boundary is missing ${required}`);
}
if (!catalog.includes('listPage') || !catalog.includes('listWorkbookPage')) {
  violations.push('workbook catalog must consume cursor pages through listPage');
}
if (!hub.includes('AbortController') || !hub.includes('loadGeneration')) {
  violations.push('workbook hub must cancel stale loads and guard request generations');
}
for (const [id, capability] of Object.entries(contracts.mutations)) {
  for (const field of ['durability', 'remote', 'schema', 'minRole', 'rebasePolicy', 'javaReducer']) {
    if (!(field in capability)) violations.push(`mutation ${id} is missing manifest field ${field}`);
  }
  if (capability.remote && !capability.javaReducer) violations.push(`remote mutation ${id} must declare a Java reducer`);
  const visible = frontend.some((source) => source.includes(`'${id}'`));
  if (capability.remote && visible && javaRegistry.includes(`Map.entry("${id}"`)) {
    violations.push(`remote visible mutation ${id} is still marked unavailable by Java`);
  }
  if (capability.durability === 'transient' && !javaRegistry.includes(`Map.entry("${id}"`)) {
    violations.push(`transient mutation ${id} must be explicitly rejected by Java`);
  }
}
if (violations.length > 0) {
  console.error(violations.join('\n'));
  process.exitCode = 1;
} else {
  console.log(`Contract capability check passed for ${Object.keys(contracts.mutations).length} declared mutations.`);
}
