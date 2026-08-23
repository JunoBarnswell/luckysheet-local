import test from 'node:test';
import assert from 'node:assert/strict';
import { generateKeyPair } from 'node:crypto';
import { exportJWK, SignJWT } from 'jose';
import {
  AuthenticationConfigurationError,
  AuthenticationError,
  JwtAuthenticator,
  extractBearerToken,
  loadAuthConfig,
} from './auth';

test('auth configuration fails closed when OIDC settings are missing', () => {
  assert.throws(
    () => loadAuthConfig({} as NodeJS.ProcessEnv),
    (error: unknown) => error instanceof AuthenticationConfigurationError && error.code === 'AUTH_CONFIGURATION_ERROR',
  );
});

test('bearer extraction rejects absent and malformed credentials', () => {
  assert.throws(() => extractBearerToken({}), AuthenticationError);
  assert.equal(extractBearerToken({ authorization: 'Bearer token-value' }), 'token-value');
  assert.throws(() => extractBearerToken({ authorization: 'Basic token-value' }), AuthenticationError);
  assert.throws(() => extractBearerToken({ authorization: ['Bearer one', 'Bearer two'] } as never), AuthenticationError);
});

test('JWT subject is verified against issuer, audience and remote JWKS', async () => {
  const { privateKey, publicKey } = await new Promise<{ privateKey: import('node:crypto').KeyObject; publicKey: import('node:crypto').KeyObject }>((resolve, reject) => {
    generateKeyPair('rsa', { modulusLength: 2048 }, (error, publicKeyResult, privateKeyResult) => {
      if (error) reject(error);
      else resolve({ publicKey: publicKeyResult, privateKey: privateKeyResult });
    });
  });
  const jwk = await exportJWK(publicKey);
  jwk.kid = 'test-key';
  const issuer = 'https://issuer.example.test';
  const token = await new SignJWT({ scope: 'workbook:write' })
    .setProtectedHeader({ alg: 'RS256', kid: 'test-key' })
    .setIssuer(issuer)
    .setAudience('react-sheets')
    .setSubject('oidc-user-1')
    .setIssuedAt()
    .setExpirationTime('5m')
    .sign(privateKey);

  const originalFetch = globalThis.fetch;
  globalThis.fetch = async () => new Response(JSON.stringify({ keys: [jwk] }), {
    status: 200,
    headers: { 'content-type': 'application/json' },
  });
  try {
    const authenticator = new JwtAuthenticator({ issuer, audience: 'react-sheets', jwksUrl: 'https://jwks.example.test/.well-known/jwks.json' });
    const principal = await authenticator.authenticateHeaders({ authorization: `Bearer ${token}` });
    assert.equal(principal.subject, 'oidc-user-1');
    await assert.rejects(
      authenticator.authenticateHeaders({ authorization: 'Bearer invalid.token.value' }),
      AuthenticationError,
    );
  } finally {
    globalThis.fetch = originalFetch;
  }
});
