import { expect, test } from '@playwright/test';

test('serves the canonical kernel manifest and raw wasm artifact', async ({ page }) => {
  const manifestResponse = await page.request.get('/kernel/kernel-manifest.json');
  expect(manifestResponse.ok()).toBeTruthy();
  const manifest = await manifestResponse.json() as { schema: string; artifact: string; bytes: number; expectedSha256: string };
  expect(manifest.schema).toBe('react-sheets.kernel-build.v1');
  expect(manifest.artifact).toBe('kernel_host.wasm');
  expect(manifest.bytes).toBeGreaterThan(0);
  expect(manifest.expectedSha256).toMatch(/^[a-f0-9]{64}$/);
  const wasmResponse = await page.request.get(`/kernel/${manifest.artifact}?sha256=${manifest.expectedSha256}`);
  expect(wasmResponse.ok()).toBeTruthy();
  expect((await wasmResponse.body()).byteLength).toBe(manifest.bytes);
});
