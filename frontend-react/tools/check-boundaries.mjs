import { readFile } from 'node:fs/promises';
import { readdir } from 'node:fs/promises';
import { join, relative } from 'node:path';

const root = process.cwd();

const forbiddenTokens = [
  '@univerjs',
  '@univerjs-pro',
  'frontend/src',
  'jquery',
  'jfrefreshgrid',
  'controlHistory',
  'new Function',
  'eval(',
  'useWorkspaceState',
  'toSheetView',
  'handleRibbonAction',
];

const forbiddenArchitectureNames = [
  /\bbridge\b/i,
  /\badapter\b/i,
  /\bshim\b/i,
  /\balias\b/i,
  /\bpro\./i,
  /@react-sheets\/(?:pro-features|storage)\b/i,
  /\b(?:WorkbookSnapshot|OperationEnvelope|CommittedOperationEnvelope|WorkspaceRecord|PendingOperationJournal|PersistenceSession|Snapshot|Protocol|Api|Backend)V[0-9]+\b/i,
  /\bschemaVersion\b/i,
  /\/api\/v[0-9]+\b/i,
];

const testFilePattern = /(?:^|[/.])(?:[^/]+\.)?(?:test|spec)\.[^/]+$/i;
const uiSourcePattern = /^apps\/web\//;
const bridgeImportPattern = /(?:from|import\s*\()\s*['"][^'"]*-bridge(?:\.[^'"]+)?['"]/;
const directWorkbookMutationPattern = /\b(?:workbook|model)\s*\.\s*(?:[A-Za-z_$][\w$]*\s*=|(?:set|add|delete|remove|insert|update|splice|push)\s*\()/;
const nativeUiElementPattern = /<\s*(div|span|button|input|select|textarea|table|thead|tbody|tr|td|th|label|form|ul|ol|li|a|canvas)\b/g;

// Business UI must render through @react-sheets/ui-system. Exact-file
// allowlists are reserved for infrastructure that owns a browser primitive;
// adding a page or panel here is an architecture violation. SheetCanvas may
// host a canvas surface when the render engine requires a direct DOM canvas.
const nativeUiInfrastructureAllowlist = new Map([
  ['apps/web/src/components/SheetCanvas.tsx', new Set(['canvas'])],
]);

function isTestFile(relPath) {
  return testFilePattern.test(relPath) || /(?:^|\/)(?:test|tests|__tests__)(?:\/|$)/i.test(relPath);
}

function collectUiArchitectureViolations(relPath, source) {
  // The first gate is intentionally scoped to the browser application. The
  // remaining feature packages still own legacy implementations during the
  // destructive migration, while new UI consumers must not add another path.
  if (!uiSourcePattern.test(relPath) || isTestFile(relPath)) return [];

  const violations = [];
  if (/\bRibbonAction\b|ribbon-(?:actions|command-map)/.test(source)) {
    violations.push(`${relPath}: UI must dispatch CommandDescriptor directly; RibbonAction/map is removed`);
  }
  if (/\bui\./.test(source)) {
    violations.push(`${relPath}: UI chrome must use typed UiSessionIntent callbacks, not ui.* command descriptors`);
  }
  if (/PivotDefinition|pivot-layout-ops/.test(source)) {
    violations.push(`${relPath}: UI must consume the canonical PivotModel and pivot feature commands`);
  }
  if (/@react-sheets\/pro-features|packages\/pro-features/.test(source)) {
    violations.push(`${relPath}: UI must consume canonical spreadsheet-app feature modules, not pro-features`);
  }
  if (/\bpro\./.test(source)) {
    violations.push(`${relPath}: UI must not depend on the legacy pro. command namespace`);
  }
  if (relPath.endsWith('-bridge.ts') || relPath.endsWith('-bridge.tsx') || bridgeImportPattern.test(source)) {
    violations.push(`${relPath}: UI must not define or import *-bridge compatibility paths`);
  }
  if (directWorkbookMutationPattern.test(source)) {
    violations.push(`${relPath}: UI must not write Workbook/WorkbookModel directly; dispatch a command`);
  }
  return violations;
}

function collectNativeUiElementViolations(relPath, source) {
  if (!uiSourcePattern.test(relPath) || isTestFile(relPath)) return [];
  const allowlistedTags = nativeUiInfrastructureAllowlist.get(relPath) ?? new Set();
  const violations = [];
  for (const match of source.matchAll(nativeUiElementPattern)) {
    const tag = match[1];
    if (tag && allowlistedTags.has(tag)) continue;
    violations.push(`${relPath}: native <${tag}> is not allowed in apps/web business UI; use @react-sheets/ui-system`);
  }
  return violations;
}

const packageImportPattern = /(?:from\s*|import\s*\()\s*['"]@react-sheets\/([^'"]+)['"]/g;
const appsImportPattern = /(?:from|import)\s+['"](?:\.\.?\/)+apps\//;

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

async function readDeclaredDependencies(owner) {
  const packageFile = join(root, owner, 'package.json');
  try {
    const manifest = JSON.parse(await readFile(packageFile, 'utf8'));
    return new Set([
      ...Object.keys(manifest.dependencies ?? {}),
      ...Object.keys(manifest.devDependencies ?? {}),
      ...Object.keys(manifest.peerDependencies ?? {}),
    ]);
  } catch (error) {
    if (error && typeof error === 'object' && 'code' in error && error.code === 'ENOENT') return null;
    throw error;
  }
}

function packageFromFile(file) {
  const rel = relative(root, file).replaceAll('\\', '/');
  const appsMatch = rel.match(/^apps\/([^/]+)\//);
  if (appsMatch) return `apps/${appsMatch[1]}`;
  const packageMatch = rel.match(/^packages\/([^/]+)\//);
  if (packageMatch) return `packages/${packageMatch[1]}`;
  return null;
}

function collectPackageImports(source) {
  const imports = new Set();
  for (const match of source.matchAll(packageImportPattern)) {
    imports.add(match[1]);
  }
  return imports;
}

function detectCycles(graph) {
  const cycles = [];
  const visiting = new Set();
  const visited = new Set();
  const stack = [];

  function visit(node) {
    if (visited.has(node)) return;
    if (visiting.has(node)) {
      const start = stack.indexOf(node);
      cycles.push([...stack.slice(start), node]);
      return;
    }
    visiting.add(node);
    stack.push(node);
    for (const next of graph.get(node) ?? []) visit(next);
    stack.pop();
    visiting.delete(node);
    visited.add(node);
  }

  for (const node of graph.keys()) visit(node);
  return cycles;
}

const files = [
  ...(await walk(join(root, 'apps'))),
  ...(await walk(join(root, 'packages'))),
];

const violations = [];
const packageGraph = new Map();
const declaredDependencies = new Map();

for (const file of files) {
  const source = await readFile(file, 'utf8');
  const relPath = relative(root, file).replaceAll('\\', '/');
  const skipForbiddenTokens = relPath.includes('features/automation/') || isTestFile(relPath);

  if (!skipForbiddenTokens) {
    for (const token of forbiddenTokens) {
      if (source.includes(token)) violations.push(`${relPath}: forbidden token "${token}"`);
    }
  }
  if (!isTestFile(relPath)) {
    for (const pattern of forbiddenArchitectureNames) {
      if (pattern.test(source)) violations.push(`${relPath}: forbidden architecture name ${pattern}`);
    }
  }

  violations.push(...collectUiArchitectureViolations(relPath, source));
  violations.push(...collectNativeUiElementViolations(relPath, source));

  const owner = packageFromFile(file);
  if (!owner || !/\.(ts|tsx|mjs|js)$/.test(file)) continue;

  if (owner.startsWith('packages/') && appsImportPattern.test(source)) {
    violations.push(`${relPath}: package must not import from apps/`);
  }

  const ownerPackage = owner.replace(/^(apps|packages)\//, '');
  if (!declaredDependencies.has(owner)) declaredDependencies.set(owner, await readDeclaredDependencies(owner));
  const ownerDependencies = declaredDependencies.get(owner);
  for (const imported of collectPackageImports(source)) {
    if (owner.startsWith('packages/ui-system') && imported !== 'ui-system') {
      violations.push(`${relPath}: ui-system must not import @react-sheets/${imported}`);
      continue;
    }
    if (owner.startsWith('packages/') && imported === ownerPackage) continue;
    const packageName = `@react-sheets/${imported}`;
    if (!ownerDependencies?.has(packageName)) {
      violations.push(`${relPath}: ${owner} imports ${packageName} without declaring it in package.json`);
    }
    if (!packageGraph.has(ownerPackage)) packageGraph.set(ownerPackage, new Set());
    packageGraph.get(ownerPackage).add(imported);
  }
}

for (const [from, targets] of packageGraph.entries()) {
  for (const target of targets) {
    if (!from || !target) continue;
  }
}

const cycles = detectCycles(packageGraph);
for (const cycle of cycles) {
  violations.push(`package cycle: ${cycle.join(' -> ')}`);
}

if (violations.length) {
  console.error(violations.join('\n'));
  process.exitCode = 1;
} else {
  console.log(`React boundary check passed: ${files.length} source files scanned, ${packageGraph.size} packages in dependency graph.`);
}
