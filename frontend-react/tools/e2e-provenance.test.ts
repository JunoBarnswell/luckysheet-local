import assert from 'node:assert/strict';
import { describe, it } from 'node:test';
import { validateE2EProvenance, type E2EProvenanceManifest } from './e2e-provenance';

const validManifest: E2EProvenanceManifest = {
  schema: 'luckysheet-local.e2e-provenance.v1',
  runId: 'abc-20260825000000',
  sourceSha: 'a'.repeat(40),
  buildSha: 'a'.repeat(40),
  sourceDirty: false,
  packageLockSha256: 'b'.repeat(64),
  nodeVersion: 'v24.0.0',
  playwrightVersion: '^1.51.0',
  browserVersion: 'Chromium 140.0.0.0',
  backendBuildIdentity: 'local-dev',
  locale: 'en-US',
  viewport: '1440x960',
  baseURL: 'http://127.0.0.1:4180',
  command: 'npm run test:e2e',
  startedAt: '2026-08-25T00:00:00.000Z',
  artifactRoot: 'frontend-react/test-results',
};

describe('E2E provenance contract', () => {
  it('accepts a complete clean manifest tied to one source/build SHA', () => {
    assert.doesNotThrow(() => validateE2EProvenance(validManifest));
  });

  it('rejects a build SHA mismatch before product tests can run', () => {
    assert.throws(() => validateE2EProvenance({ ...validManifest, buildSha: 'c'.repeat(40) }), /SHA mismatch/);
  });

  it('rejects dirty source and unresolved browser identity', () => {
    assert.throws(() => validateE2EProvenance({ ...validManifest, sourceDirty: true }), /clean source/);
    assert.throws(() => validateE2EProvenance({ ...validManifest, browserVersion: 'unresolved' }), /browser version/);
  });
});
