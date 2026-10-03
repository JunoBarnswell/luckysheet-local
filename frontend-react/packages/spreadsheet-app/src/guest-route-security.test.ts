import assert from 'node:assert/strict';
import { test } from 'node:test';
import { resolveShareToken } from './runtime';

test('fragment credential is removed synchronously and memory is workbook scoped', () => {
  const original = Object.getOwnPropertyDescriptor(globalThis, 'window');
  let location = new URL('https://app.test/workbooks/book#share=secret');
  const values = new Map<string, string>();
  const history = { state: null, replaceState(_s: unknown, _t: string, value: string) { location = new URL(value, location); } };
  const sessionStorage = { getItem: (key: string) => values.get(key) ?? null, setItem: (key: string, value: string) => { values.set(key, value); }, removeItem: (key: string) => { values.delete(key); } };
  Object.defineProperty(globalThis, 'window', { configurable: true, value: { get location() { return location; }, history, sessionStorage } });
  try {
    assert.equal(resolveShareToken(), 'secret'); assert.equal(location.hash, ''); assert.equal(location.search, '');
    assert.equal(resolveShareToken(), 'secret');
    assert.equal(values.get('share:/workbooks/book'), 'secret', 'credential is kept tab-scoped for refresh after fragment cleanup');
    location = new URL('https://app.test/workbooks/other'); assert.equal(resolveShareToken(), null);
    location = new URL('https://app.test/workbooks/book?share=legacy'); assert.equal(resolveShareToken(), null);
  } finally { if (original) Object.defineProperty(globalThis, 'window', original); else Reflect.deleteProperty(globalThis, 'window'); }
});
