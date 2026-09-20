import { existsSync, readFileSync, readdirSync } from 'node:fs';
import { dirname, join, relative, resolve } from 'node:path';
import { fileURLToPath } from 'node:url';

const forbiddenEntries = [
  'Cargo.toml', 'Cargo.lock', 'rust-toolchain', 'rust-toolchain.toml', '.cargo', 'kernel',
  'frontend-react/packages/kernel-client', 'frontend-react/apps/web/public/kernel',
  'backend/src/main/java/com/xc/luckysheet/server/kernel',
  'target/release/workbook-kernel-host', 'target/release/workbook-kernel-host.exe',
  'target/wasm32-unknown-unknown',
];
const ignoredDirectories = new Set(['.git', 'node_modules', 'target', 'dist', '.tools', '.gitnexus']);
const executableRoots = ['backend/src/main', 'frontend-react/apps', 'frontend-react/packages', 'scripts', 'installer', 'contracts', '.github/workflows'];
const sourceExtensions = /\.(?:[cm]?js|tsx?|java|json|ya?ml|ps1|xml|toml|nsi|nsh|cmd|bat|sh|properties)$/i;
const retiredRuntime = /@react-sheets\/kernel-client|workbook-kernel-host|kernel_host\.wasm|wasm32-unknown-unknown|\bKernelHost(?:Gateway|Properties|Configuration)\b|\bKernelPersistenceService\b|\b(?:cargo|rustc|rustup)\s+(?:build|test|run|install|--version)/;

/** Build-tool policy, not a runtime compatibility path. */
export function inspectRuntimeStack(root) {
  const problems = [];
  for (const entry of forbiddenEntries) if (existsSync(join(root, entry))) problems.push(`Retired runtime entry: ${entry}`);
  const visit = (directory) => {
    if (!existsSync(directory)) return;
    for (const entry of readdirSync(directory, { withFileTypes: true })) {
      if (entry.isDirectory() && ignoredDirectories.has(entry.name)) continue;
      const path = join(directory, entry.name);
      const name = relative(root, path).replaceAll('\\', '/');
      if (entry.isSymbolicLink()) { problems.push(`Source entry must not hide dependencies behind a symlink: ${name}`); continue; }
      if (entry.isDirectory()) { visit(path); continue; }
      if (/\.rs$|\.wasm$|^Cargo\.(?:toml|lock)$|^rust-toolchain(?:\.toml)?$/i.test(entry.name)) {
        problems.push(`Retired runtime artifact: ${name}`);
      }
      if (name === 'scripts/check-runtime-stack.mjs' || name === 'scripts/check-runtime-stack.test.mjs') continue;
      if (sourceExtensions.test(entry.name) && retiredRuntime.test(readFileSync(path, 'utf8'))) {
        problems.push(`Retired runtime dependency: ${name}`);
      }
    }
  };
  for (const directory of executableRoots) visit(join(root, directory));
  for (const entry of ['frontend-react/package.json', 'frontend-react/package-lock.json', 'frontend-react/vite.config.ts', 'backend/pom.xml']) {
    const path = join(root, entry);
    if (existsSync(path) && retiredRuntime.test(readFileSync(path, 'utf8'))) problems.push(`Retired runtime dependency: ${entry}`);
  }
  const requiredOwners = [
    'backend/src/main/java/com/xc/luckysheet/server/service/WorkbookOperationService.java',
    'frontend-react/packages/core-model/src/domain.ts',
    'frontend-react/packages/exchange-excel-ooxml/src/ooxml.ts',
  ];
  for (const owner of requiredOwners) if (!existsSync(join(root, owner))) problems.push(`Canonical owner missing: ${owner}`);
  const frontend = join(root, 'frontend-react/package.json');
  if (!existsSync(frontend) || JSON.parse(readFileSync(frontend, 'utf8')).engines?.node !== '24.x') problems.push('Frontend must declare Node 24.x');
  const pom = join(root, 'backend/pom.xml');
  if (!existsSync(pom) || !/<java\.version>21<\/java\.version>/.test(readFileSync(pom, 'utf8'))) problems.push('Backend must declare Java 21');
  return problems;
}

if (process.argv[1] && resolve(process.argv[1]) === fileURLToPath(import.meta.url)) {
  const root = resolve(dirname(fileURLToPath(import.meta.url)), '..');
  const problems = inspectRuntimeStack(root);
  if (problems.length) {
    console.error(`RUNTIME_STACK_INVALID\n${problems.join('\n')}`);
    process.exitCode = 1;
  } else console.log('Runtime stack verified: React / Node 24 / Java 21; no retired native or WASM runtime.');
}
