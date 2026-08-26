import 'fake-indexeddb/auto';

// Node exposes BroadcastChannel, but the IndexedDB unit harness has no browser
// page lifecycle to close it. Cross-page coordination is covered by explicit
// coordinator tests; keep the fake storage environment deterministic here.
Object.defineProperty(globalThis, 'BroadcastChannel', {
  configurable: true,
  writable: true,
  value: undefined,
});
