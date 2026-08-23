import type { IncomingHttpHeaders, IncomingMessage } from 'node:http';
import {
  createRemoteJWKSet,
  jwtVerify,
  type JWTPayload,
  type RemoteJWKSet,
} from 'jose';

export interface AuthenticatedPrincipal {
  /** OIDC subject. This is the only actor identity accepted by the server. */
  subject: string;
  issuer: string;
  claims: JWTPayload;
}

export interface AuthConfig {
  issuer: string;
  audience: string | string[];
  jwksUrl: string;
  clockToleranceSeconds?: number;
}

export class AuthenticationError extends Error {
  readonly code = 'UNAUTHENTICATED' as const;
  readonly status = 401 as const;

  constructor(message = 'Bearer authentication is required') {
    super(message);
    this.name = 'AuthenticationError';
  }
}

export class AuthenticationConfigurationError extends Error {
  readonly code = 'AUTH_CONFIGURATION_ERROR' as const;
  readonly status = 503 as const;

  constructor(message: string) {
    super(message);
    this.name = 'AuthenticationConfigurationError';
  }
}

function nonEmptyEnvironmentValue(env: NodeJS.ProcessEnv, name: string): string {
  const value = env[name]?.trim();
  if (!value) throw new AuthenticationConfigurationError(`${name} must be configured`);
  return value;
}

export function loadAuthConfig(env: NodeJS.ProcessEnv = process.env): AuthConfig {
  const issuer = nonEmptyEnvironmentValue(env, 'AUTH_ISSUER');
  const audienceValue = nonEmptyEnvironmentValue(env, 'AUTH_AUDIENCE');
  const jwksUrl = nonEmptyEnvironmentValue(env, 'AUTH_JWKS_URL');
  try {
    // Validate URLs at startup/configuration time, not after an unauthenticated
    // request has already reached business code.
    new URL(issuer);
    const jwks = new URL(jwksUrl);
    if (jwks.protocol !== 'https:' && env.NODE_ENV === 'production') {
      throw new Error('AUTH_JWKS_URL must use HTTPS in production');
    }
  } catch (error) {
    throw new AuthenticationConfigurationError(error instanceof Error ? error.message : 'Invalid auth URL configuration');
  }
  const audience = audienceValue.split(',').map((value) => value.trim()).filter(Boolean);
  if (audience.length === 0) throw new AuthenticationConfigurationError('AUTH_AUDIENCE must contain a value');
  return {
    issuer,
    audience: audience.length === 1 ? audience[0]! : audience,
    jwksUrl,
    clockToleranceSeconds: Number(env.AUTH_CLOCK_TOLERANCE_SECONDS ?? 5) || 5,
  };
}

export function extractBearerToken(headers: Pick<IncomingHttpHeaders, 'authorization'>): string {
  const value = headers.authorization;
  if (Array.isArray(value)) throw new AuthenticationError('Exactly one bearer token is required');
  if (!value) throw new AuthenticationError('Bearer authentication is required');
  const match = /^Bearer\s+(\S+)$/i.exec(value.trim());
  if (!match?.[1]) throw new AuthenticationError('Authorization must use the Bearer scheme');
  return match[1];
}

export class JwtAuthenticator {
  private readonly config: AuthConfig | null;
  private readonly configurationError: AuthenticationConfigurationError | null;
  private readonly remoteJwks: RemoteJWKSet | null;

  constructor(config?: AuthConfig) {
    if (config) {
      this.config = config;
      this.configurationError = null;
      this.remoteJwks = createRemoteJWKSet(new URL(config.jwksUrl));
      return;
    }
    try {
      this.config = loadAuthConfig();
      this.configurationError = null;
      this.remoteJwks = createRemoteJWKSet(new URL(this.config.jwksUrl));
    } catch (error) {
      this.config = null;
      this.configurationError = error instanceof AuthenticationConfigurationError
        ? error
        : new AuthenticationConfigurationError('Invalid authentication configuration');
      this.remoteJwks = null;
    }
  }

  async authenticateRequest(request: IncomingMessage): Promise<AuthenticatedPrincipal> {
    return this.authenticateHeaders(request.headers);
  }

  async authenticateHeaders(headers: Pick<IncomingHttpHeaders, 'authorization'>): Promise<AuthenticatedPrincipal> {
    if (this.configurationError) throw this.configurationError;
    if (!this.config || !this.remoteJwks) throw new AuthenticationConfigurationError('Authentication is not configured');
    const token = extractBearerToken(headers);
    try {
      const verified = await jwtVerify(token, this.remoteJwks, {
        issuer: this.config.issuer,
        audience: this.config.audience,
        // Do not allow jose to infer an algorithm from attacker-controlled
        // token headers.  These are the OIDC algorithms supported by the
        // server's remote JWKS configuration.
        algorithms: ['RS256', 'RS384', 'RS512', 'PS256', 'PS384', 'PS512', 'ES256', 'ES384', 'ES512', 'EdDSA'],
        clockTolerance: this.config.clockToleranceSeconds ?? 5,
      });
      const subject = verified.payload.sub;
      if (typeof subject !== 'string' || subject.trim().length === 0) {
        throw new AuthenticationError('Bearer token does not contain a subject');
      }
      return {
        subject: subject.trim(),
        issuer: typeof verified.payload.iss === 'string' ? verified.payload.iss : this.config.issuer,
        claims: verified.payload,
      };
    } catch (error) {
      if (error instanceof AuthenticationError) throw error;
      throw new AuthenticationError('Bearer token is invalid or expired');
    }
  }
}
