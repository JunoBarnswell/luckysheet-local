import assert from 'node:assert/strict';
import test from 'node:test';
import { KernelInvocationError, initializeKernelFromManifest, parseKernelBuildManifest, versionedKernelWasmUrl } from './index';

const hash = 'a'.repeat(64);

test('kernel build manifest produces a hash-bound WASM URL', () => {
  const manifest = parseKernelBuildManifest({
    schema: 'react-sheets.kernel-build.v1',
    artifact: 'kernel_host.wasm',
    bytes: 12,
    expectedSha256: hash,
  });
  assert.equal(versionedKernelWasmUrl(`https://sheets.test/kernel/kernel-manifest.json?build=${hash}`, manifest).toString(), `https://sheets.test/kernel/kernel_host.wasm?sha256=${hash}`);
});

test('legacy or mismatched kernel manifests fail closed', async () => {
  assert.throws(
    () => parseKernelBuildManifest({ artifact: 'kernel_host.wasm', bytes: 12, sha256: hash }),
    (cause: unknown) => cause instanceof KernelInvocationError && cause.code === 'KERNEL_MANIFEST_INVALID',
  );

  const previousFetch = globalThis.fetch;
  globalThis.fetch = async () => new Response(JSON.stringify({
    schema: 'react-sheets.kernel-build.v1',
    artifact: 'kernel_host.wasm',
    bytes: 12,
    expectedSha256: 'b'.repeat(64),
  }), { status: 200, headers: { 'content-type': 'application/json' } });
  try {
    await assert.rejects(
      initializeKernelFromManifest(`https://sheets.test/kernel/kernel-manifest.json?build=${hash}`),
      (cause: unknown) => cause instanceof KernelInvocationError && cause.code === 'KERNEL_BUILD_CONTRACT_MISMATCH',
    );
  } finally {
    globalThis.fetch = previousFetch;
  }
});
