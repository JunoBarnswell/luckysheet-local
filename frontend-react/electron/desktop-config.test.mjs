import assert from 'node:assert/strict';
import test from 'node:test';
import {
  collaborationArgument,
  parseDesktopConfig,
  resolveBackendOrigin,
  resolveCollaborationUrl,
} from './desktop-config.mjs';

test('desktop config maps one validated API origin to its canonical websocket endpoint', () => {
  const origin = resolveBackendOrigin('https://sheets.example.test');
  const collaborationUrl = resolveCollaborationUrl(origin);
  const config = parseDesktopConfig(['electron', collaborationArgument(collaborationUrl)]);
  assert.equal(origin.toString(), 'https://sheets.example.test/');
  assert.equal(config.collaborationUrl, 'wss://sheets.example.test/ws');
  assert.equal(Object.isFrozen(config), true);
});

test('desktop config fails closed for credentials, paths, and missing renderer arguments', () => {
  assert.throws(() => resolveBackendOrigin('https://user:secret@sheets.example.test'), /uncredentialed/);
  assert.throws(() => resolveBackendOrigin('https://sheets.example.test/api'), /origin/);
  assert.throws(() => collaborationArgument('https://sheets.example.test/ws'), /invalid/);
  assert.throws(() => collaborationArgument('wss://sheets.example.test/not-ws'), /invalid/);
  assert.throws(() => parseDesktopConfig(['electron']), /missing/);
});
