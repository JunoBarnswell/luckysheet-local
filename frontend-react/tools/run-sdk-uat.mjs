import { spawn } from 'node:child_process';
import { createWriteStream, existsSync } from 'node:fs';
import { mkdtemp, mkdir } from 'node:fs/promises';
import { createServer } from 'node:net';
import { tmpdir } from 'node:os';
import path from 'node:path';

const frontend = path.resolve(import.meta.dirname, '..');
const jar = path.resolve(frontend, '../backend/target/react-sheets.jar');
if (!existsSync(jar)) throw new Error('Build backend/target/react-sheets.jar before running SDK UAT.');

async function requireFreePort(port) {
  const server = createServer();
  await new Promise((resolve, reject) => {
    server.once('error', reject);
    server.listen(port, '127.0.0.1', resolve);
  });
  await new Promise((resolve, reject) => server.close(error => error ? reject(error) : resolve()));
}
await requireFreePort(8082);
await requireFreePort(4180);
const isolated = await mkdtemp(path.join(tmpdir(), 'sdk-product-uat-'));
const evidence = process.env.SDK_UAT_EVIDENCE_DIR ?? path.join(isolated, 'evidence');
await mkdir(evidence, { recursive: true });
const backendLog = createWriteStream(path.join(evidence, 'backend.log'));
const java = process.env.JAVA_HOME
  ? path.join(process.env.JAVA_HOME, 'bin', process.platform === 'win32' ? 'java.exe' : 'java') : 'java';
const backend = spawn(java, ['-jar', jar], {
  cwd: frontend, env: { ...process.env, SHEETS_DATA_DIR: isolated, DATABASE_URL: `jdbc:h2:file:${isolated.replaceAll('\\', '/')}/luckysheet_canonical;DB_CLOSE_DELAY=-1;CASE_INSENSITIVE_IDENTIFIERS=TRUE`, DATABASE_USERNAME: 'sa', DATABASE_PASSWORD: '', SERVER_PORT: '8082', AUTH_MODE: 'local' },
  stdio: ['ignore', 'pipe', 'pipe'],
});
backend.stdout.pipe(backendLog); backend.stderr.pipe(backendLog);
let launchError;
backend.once('error', error => { launchError = error; });
let acceptance;
const stop = () => { acceptance?.kill(); backend.kill(); };
process.once('SIGINT', stop); process.once('SIGTERM', stop);
try {
  const deadline = Date.now() + 120_000;
  let ready = false;
  while (Date.now() < deadline) {
    if (launchError) throw launchError;
    if (backend.exitCode !== null) throw new Error(`Java UAT service exited ${backend.exitCode}; see ${evidence}/backend.log`);
    try {
      ready = (await fetch('http://127.0.0.1:8082/health', { signal: AbortSignal.timeout(1_000) })).ok;
    } catch { /* Only bounded service-start polling; never retries a workbook action. */ }
    if (ready) break;
    await new Promise(resolve => setTimeout(resolve, 250));
  }
  if (!ready) throw new Error(`Java UAT service did not become ready; see ${evidence}/backend.log`);
  acceptance = spawn(process.execPath, [path.join(frontend, 'node_modules/@playwright/test/cli.js'), 'test', 'e2e/sdk-product.spec.ts', '--workers=1', '--retries=0'], {
    cwd: frontend, stdio: 'inherit', env: { ...process.env, SDK_UAT_ENABLED: '1',
      SDK_UAT_BOOTSTRAP_FILE: path.join(isolated, 'bootstrap-token'), SDK_UAT_EVIDENCE_DIR: evidence,
      E2E_COMMAND: 'npm run test:sdk-uat',
    },
  });
  const exitCode = await new Promise((resolve, reject) => {
    acceptance.once('error', reject);
    acceptance.once('exit', code => resolve(code ?? 1));
  });
  process.exitCode = exitCode;
  console.log(`SDK UAT evidence: ${evidence}`);
} finally {
  stop();
  if (backend.exitCode === null) await Promise.race([
    new Promise(resolve => backend.once('exit', resolve)),
    new Promise(resolve => setTimeout(() => { backend.kill('SIGKILL'); resolve(); }, 5_000)),
  ]);
  backendLog.end();
}
