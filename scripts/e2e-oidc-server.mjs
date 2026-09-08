import { generateKeyPairSync, sign } from 'node:crypto';
import { writeFileSync } from 'node:fs';
import { createServer } from 'node:http';

const output = process.argv[2];
if (!output) throw new Error('Usage: node scripts/e2e-oidc-server.mjs <oidc-user-file>');

const issuer = process.env.AUTH_ISSUER ?? 'http://127.0.0.1:8090';
const audience = process.env.AUTH_AUDIENCE ?? 'react-sheets-api';
const clientId = process.env.OIDC_CLIENT_ID ?? 'react-sheets-web';
const subject = process.env.OIDC_SUBJECT ?? 'ci-browser-owner';
const { privateKey, publicKey } = generateKeyPairSync('rsa', { modulusLength: 2048 });
const jwk = publicKey.export({ format: 'jwk' });
jwk.kid = 'ci-browser';
jwk.alg = 'RS256';
jwk.use = 'sig';

const encoded = (value) => Buffer.from(JSON.stringify(value)).toString('base64url');
const issuedAt = Math.floor(Date.now() / 1000);
const expiresAt = issuedAt + 3600;
const header = encoded({ alg: 'RS256', typ: 'JWT', kid: jwk.kid });
const claims = encoded({ iss: issuer, aud: audience, sub: subject, iat: issuedAt, exp: expiresAt });
const signingInput = `${header}.${claims}`;
const accessToken = `${signingInput}.${sign('RSA-SHA256', Buffer.from(signingInput), privateKey).toString('base64url')}`;
writeFileSync(output, JSON.stringify({
  access_token: accessToken,
  token_type: 'Bearer',
  scope: 'openid profile email',
  expires_at: expiresAt,
  profile: { sub: subject },
}), 'utf8');

const discovery = JSON.stringify({
  issuer,
  authorization_endpoint: `${issuer}/authorize`,
  token_endpoint: `${issuer}/token`,
  jwks_uri: `${issuer}/.well-known/jwks.json`,
  response_types_supported: ['code'],
  subject_types_supported: ['public'],
  id_token_signing_alg_values_supported: ['RS256'],
  client_id: clientId,
});
const jwks = JSON.stringify({ keys: [jwk] });
const server = createServer((request, response) => {
  if (request.url === '/.well-known/jwks.json') {
    response.writeHead(200, { 'content-type': 'application/json' });
    response.end(jwks);
    return;
  }
  if (request.url === '/.well-known/openid-configuration') {
    response.writeHead(200, { 'content-type': 'application/json' });
    response.end(discovery);
    return;
  }
  response.writeHead(404);
  response.end();
});

server.listen(8090, '127.0.0.1', () => {
  process.stdout.write(`E2E OIDC fixture ready at ${issuer}; user=${output}\n`);
});
