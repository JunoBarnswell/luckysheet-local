import { readFile } from 'node:fs/promises';
import { resolve } from 'node:path';

const root = process.cwd();
const contracts = JSON.parse(await readFile(resolve(root, '../contracts/workbook-contract.json'), 'utf8'));
const snapshotSchema = JSON.parse(await readFile(resolve(root, '../contracts/workbook-snapshot.schema.json'), 'utf8'));
const javaRegistry = await readFile(resolve(root, '../backend/src/main/java/com/xc/luckysheet/server/mutation/MutationDescriptorRegistry.java'), 'utf8');
const generatedTypeScript = await readFile(resolve(root, 'packages/protocol/src/generated-contract.ts'), 'utf8');
const generatedJava = await readFile(resolve(root, '../backend/src/main/java/com/xc/luckysheet/server/contract/GeneratedWorkbookContract.java'), 'utf8');
const permissionService = await readFile(resolve(root, 'packages/spreadsheet-app/src/permission-service.ts'), 'utf8');
const protectionResolver = await readFile(resolve(root, '../backend/src/main/java/com/xc/luckysheet/server/mutation/ProtectionResolver.java'), 'utf8');
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
  || !generatedJava.includes(`API_VERSION = ${JSON.stringify(contracts.apiVersion)}`)
  || !generatedTypeScript.includes('MUTATION_PERMISSION_POLICIES')
  || !generatedJava.includes('MUTATION_PERMISSIONS')
  || !generatedTypeScript.includes('COMMAND_PERMISSION_PREFIXES')) {
  violations.push('generated contract outputs are stale; run npm run generate:contracts');
}
if (permissionService.includes('isFormatOnlyRestore') || permissionService.includes('protectionActionForCommand') || permissionService.includes('protectionActionForMutation')) {
  violations.push('permission service must consume generated command and mutation policies');
}
if (protectionResolver.includes('boundedDimension') || protectionResolver.includes('for (int row') || protectionResolver.includes('for (int column')) {
  violations.push('protection resolver must use canonical extents and sparse exception lookup');
}
const protectionActions = new Set(['none', 'edit-cell', 'format', 'insert-rows', 'insert-columns', 'delete-rows', 'delete-columns', 'sort', 'auto-filter', 'edit-objects', 'select-locked', 'select-unlocked']);
const permissionCapabilities = new Set(['navigate', 'edit-cell', 'format', 'structure', 'drawing', 'protect', 'share', 'comment', 'restore', 'query', 'script']);
if (!contracts.permissions?.commands || !Array.isArray(contracts.permissions?.commandPrefixes) || !contracts.permissions?.mutations) {
  violations.push('permission contract must declare exact commands, command prefixes, and mutation policies');
}
for (const [id, policy] of Object.entries(contracts.permissions?.mutations ?? {})) {
  if (!permissionCapabilities.has(policy.capability) || !protectionActions.has(policy.protectionAction) || typeof policy.checksProtection !== 'boolean') {
    violations.push(`mutation permission policy ${id} is invalid`);
  }
}
for (const [id, policy] of Object.entries(contracts.permissions?.commands ?? {})) {
  if (!permissionCapabilities.has(policy.capability) || !protectionActions.has(policy.protectionAction) || typeof policy.checksProtection !== 'boolean') {
    violations.push(`command permission policy ${id} is invalid`);
  }
}
for (const policy of contracts.permissions?.commandPrefixes ?? []) {
  if (typeof policy.prefix !== 'string' || policy.prefix.length === 0 || !permissionCapabilities.has(policy.capability)
    || !protectionActions.has(policy.protectionAction) || typeof policy.checksProtection !== 'boolean') {
    violations.push(`command permission prefix policy is invalid: ${policy.prefix ?? '<empty>'}`);
  }
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
  if (!contracts.permissions?.mutations?.[id]) violations.push(`mutation ${id} is missing its permission policy`);
}
if (violations.length > 0) {
  console.error(violations.join('\n'));
  process.exitCode = 1;
} else {
  console.log(`Contract capability check passed for ${Object.keys(contracts.mutations).length} declared mutations.`);
}
