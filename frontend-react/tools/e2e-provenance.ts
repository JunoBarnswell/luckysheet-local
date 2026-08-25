import crypto from 'node:crypto';
import { execFileSync } from 'node:child_process';
import fs from 'node:fs';
import path from 'node:path';
import process from 'node:process';
import { chromium, type FullConfig } from '@playwright/test';

const repoRoot = path.resolve(import.meta.dirname, '../..');
const frontendRoot = path.join(repoRoot, 'frontend-react');
const manifestPath = path.join(frontendRoot, 'test-results', 'provenance.json');

export interface E2EProvenanceManifest {
  schema: 'luckysheet-local.e2e-provenance.v1';
  runId: string;
  sourceSha: string;
  buildSha: string;
  sourceDirty: boolean;
  packageLockSha256: string;
  nodeVersion: string;
  playwrightVersion: string;
  browserVersion: string;
  backendBuildIdentity: string;
  locale: string;
  viewport: string;
  baseURL: string;
  command: string;
  startedAt: string;
  artifactRoot: string;
}

function git(...args: string[]): string {
  return execFileSync('git', args, { cwd: repoRoot, encoding: 'utf8' }).trim();
}

function packageLockSha256(): string {
  return crypto.createHash('sha256').update(fs.readFileSync(path.join(frontendRoot, 'package-lock.json'))).digest('hex');
}

function packageJson(): { version?: string; devDependencies?: Record<string, string> } {
  return JSON.parse(fs.readFileSync(path.join(frontendRoot, 'package.json'), 'utf8')) as { version?: string; devDependencies?: Record<string, string> };
}

async function installedBrowserVersion(): Promise<string> {
  const browser = await chromium.launch({ headless: true });
  try {
    return browser.version();
  } finally {
    await browser.close();
  }
}

async function buildManifest(): Promise<E2EProvenanceManifest> {
  const startedAt = new Date().toISOString();
  const sourceSha = git('rev-parse', 'HEAD');
  const sourceDirty = git('status', '--porcelain', '--untracked-files=all').length > 0;
  const packageInfo = packageJson();
  const playwrightVersion = packageInfo.devDependencies?.['@playwright/test'] ?? 'unknown';
  const browserVersion = process.env.PLAYWRIGHT_BROWSER_VERSION ?? await installedBrowserVersion();
  return {
    schema: 'luckysheet-local.e2e-provenance.v1',
    runId: `${sourceSha}-${startedAt.replaceAll(/[-:.TZ]/g, '')}`,
    sourceSha,
    buildSha: process.env.E2E_BUILD_SHA ?? sourceSha,
    sourceDirty,
    packageLockSha256: packageLockSha256(),
    nodeVersion: process.version,
    playwrightVersion,
    browserVersion,
    backendBuildIdentity: process.env.E2E_BACKEND_BUILD_ID ?? 'local-dev',
    locale: process.env.E2E_LOCALE ?? 'en-US',
    viewport: process.env.E2E_VIEWPORT ?? '1440x960',
    baseURL: process.env.E2E_BASE_URL ?? 'http://127.0.0.1:4180',
    command: process.env.E2E_COMMAND ?? 'npm run test:e2e',
    startedAt,
    artifactRoot: 'frontend-react/test-results',
  };
}

export function validateE2EProvenance(manifest: E2EProvenanceManifest): void {
  const required = ['runId', 'sourceSha', 'buildSha', 'packageLockSha256', 'nodeVersion', 'playwrightVersion', 'browserVersion', 'backendBuildIdentity', 'locale', 'viewport', 'baseURL', 'command', 'startedAt', 'artifactRoot'] as const;
  for (const key of required) {
    if (typeof manifest[key] !== 'string' || manifest[key].trim() === '') throw new Error(`E2E provenance is missing ${key}`);
  }
  if (manifest.schema !== 'luckysheet-local.e2e-provenance.v1') throw new Error(`Unsupported E2E provenance schema: ${manifest.schema}`);
  if (manifest.sourceSha !== manifest.buildSha) throw new Error(`E2E build/source SHA mismatch: build=${manifest.buildSha}, source=${manifest.sourceSha}`);
  if (manifest.sourceDirty) throw new Error('E2E provenance requires a clean source tree');
  if (!/^[0-9a-f]{64}$/i.test(manifest.packageLockSha256)) throw new Error('E2E provenance has an invalid package-lock digest');
  if (manifest.browserVersion.toLowerCase() === 'unknown' || manifest.browserVersion.toLowerCase() === 'unresolved') throw new Error('E2E provenance requires an installed browser version');
  if (!/^\d+x\d+$/.test(manifest.viewport)) throw new Error(`E2E provenance has an invalid viewport: ${manifest.viewport}`);
  if (Number.isNaN(Date.parse(manifest.startedAt))) throw new Error('E2E provenance has an invalid start time');
}

export async function writeE2EProvenance(): Promise<E2EProvenanceManifest> {
  const manifest = await buildManifest();
  validateE2EProvenance(manifest);
  fs.mkdirSync(path.dirname(manifestPath), { recursive: true });
  fs.writeFileSync(manifestPath, `${JSON.stringify(manifest, null, 2)}\n`, 'utf8');
  return manifest;
}

function assertNoTrackedArtifacts(): void {
  const tracked = git('ls-files', '--', '**/test-results/**', '**/playwright-report/**', '**/*.trace.zip', '**/*.webm', '**/*.png');
  const forbidden = tracked.split(/\r?\n/).filter((entry) => /(^|\/)(test-results|playwright-report)\//.test(entry) || /\.(trace\.zip|webm)$/i.test(entry));
  if (forbidden.length > 0) throw new Error(`Generated Playwright artifacts are tracked:\n${forbidden.join('\n')}`);
}

function readManifest(): E2EProvenanceManifest {
  if (!fs.existsSync(manifestPath)) throw new Error(`E2E provenance manifest is missing: ${manifestPath}`);
  return JSON.parse(fs.readFileSync(manifestPath, 'utf8')) as E2EProvenanceManifest;
}

export default async function globalSetup(_config: FullConfig): Promise<void> {
  await writeE2EProvenance();
}

if (process.argv[1] && path.resolve(process.argv[1]) === path.resolve(import.meta.filename)) {
  const command = process.argv[2] ?? 'write';
  if (command === 'check-tracked') {
    assertNoTrackedArtifacts();
  } else if (command === 'check') {
    assertNoTrackedArtifacts();
    validateE2EProvenance(readManifest());
  } else if (command === 'write') {
    await writeE2EProvenance();
  } else {
    throw new Error(`Unknown E2E provenance command: ${command}`);
  }
}
