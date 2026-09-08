const KERNEL_BUILD_ID = import.meta.env.VITE_KERNEL_BUILD_ID;

if (!/^[a-f0-9]{64}$/.test(KERNEL_BUILD_ID ?? '')) {
  throw new Error('The frontend build is missing its kernel build binding.');
}

export function kernelManifestUrl(): string {
  return `/kernel/kernel-manifest.json?build=${encodeURIComponent(KERNEL_BUILD_ID)}`;
}

export function versionedKernelResourceUrls(): string[] {
  return [kernelManifestUrl(), `/kernel/kernel_host.wasm?sha256=${encodeURIComponent(KERNEL_BUILD_ID)}`];
}

export function versionedServiceWorkerUrl(): string {
  return `/sw.js?build=${encodeURIComponent(KERNEL_BUILD_ID)}`;
}
