import assert from 'node:assert/strict';
import { describe, it } from 'node:test';
import { hasUnexpectedSourceChanges, validateE2EProvenance, validateKernelBuildArtifact, type E2EProvenanceManifest } from './e2e-provenance';

const validManifest: E2EProvenanceManifest = {
  schema: 'luckysheet-local.e2e-provenance.v2',
  runId: 'abc-20260825000000',
  sourceSha: 'a'.repeat(40),
  buildSha: 'a'.repeat(40),
  sourceDirty: false,
  packageLockSha256: 'b'.repeat(64),
  kernelArtifactSha256: 'c'.repeat(64),
  kernelArtifactBytes: 1024,
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

  it('allows only verified kernel build outputs to differ from the checkout', () => {
    assert.equal(hasUnexpectedSourceChanges(' M frontend-react/apps/web/public/kernel/kernel_host.wasm\n M frontend-react/apps/web/public/kernel/kernel-manifest.json'), false);
    assert.equal(hasUnexpectedSourceChanges(' M frontend-react/apps/web/public/kernel/kernel_host.wasm\n M frontend-react/tools/e2e-provenance.ts'), true);
  });

  it('binds the generated kernel payload to its manifest', () => {
    const payload = Buffer.from('canonical kernel');
    const expectedSha256 = '437807f09eaf2edb00545c1430015531e17ff097c63309788c37d0949d36487b';
    const manifest = { schema: 'react-sheets.kernel-build.v1' as const, artifact: 'kernel_host.wasm', bytes: payload.byteLength, expectedSha256 };
    assert.equal(validateKernelBuildArtifact(manifest, payload), expectedSha256);
    assert.throws(() => validateKernelBuildArtifact({ ...manifest, bytes: payload.byteLength + 1 }, payload), /byte length mismatch/);
    assert.throws(() => validateKernelBuildArtifact({ ...manifest, expectedSha256: '0'.repeat(64) }, payload), /SHA-256 mismatch/);
  });
});
