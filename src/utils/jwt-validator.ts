import { createRemoteJWKSet, decodeJwt, jwtVerify, type JWTPayload } from "jose";
import {
  GRAPH_TENANT_ID,
  MANAGED_IDENTITY_ENABLED,
  MANAGED_IDENTITY_AUDIENCE,
  VAULT_BEARER_AUTH_ENABLED,
  VAULT_ALLOWED_APP_IDS,
  VAULT_EXPECTED_AUDIENCES,
} from "./config.js";
import { log } from "./logger.js";

export interface ValidatedToken {
  oid: string;
  roles: string[];
  appid: string;
}

export interface ValidatedVaultToken {
  oid: string;
  appid: string;
  upn: string;
}

interface MIJwtPayload extends JWTPayload {
  oid?: string;
  roles?: string[];
  appid?: string;
  azp?: string;
}

interface VaultJwtPayload extends JWTPayload {
  oid?: string;
  appid?: string;
  azp?: string;
  upn?: string;
  preferred_username?: string;
}

// Lazy-initialized JWKS — created once on first validation call
let jwks: ReturnType<typeof createRemoteJWKSet> | null = null;

function getJWKS(): ReturnType<typeof createRemoteJWKSet> {
  if (!jwks) {
    const jwksUrl = new URL(
      `https://login.microsoftonline.com/${GRAPH_TENANT_ID}/discovery/v2.0/keys`
    );
    jwks = createRemoteJWKSet(jwksUrl);
  }
  return jwks;
}

/**
 * Quick structural check: is this a JWT (3 dot-separated base64url segments)?
 * Does NOT validate — just checks shape so we can distinguish JWTs from opaque tokens.
 */
export function looksLikeJwt(token: string): boolean {
  const parts = token.split(".");
  return parts.length === 3 && parts.every((p) => p.length > 0);
}

/**
 * Validates a Managed Identity JWT issued by Azure AD.
 *
 * Expected claims:
 *   aud: api://b8157430-c8f3-4760-beaa-a4b95cfc20a7 (or MANAGED_IDENTITY_AUDIENCE)
 *   iss: https://login.microsoftonline.com/{tenant}/v2.0
 *   roles: ["MCP.Invoke"]
 *
 * @throws if validation fails or required claims are missing
 */
export async function validateManagedIdentityToken(
  token: string
): Promise<ValidatedToken> {
  if (!MANAGED_IDENTITY_ENABLED) {
    throw new Error("Managed Identity auth is not enabled");
  }

  const expectedIssuer = `https://login.microsoftonline.com/${GRAPH_TENANT_ID}/v2.0`;

  const { payload } = await jwtVerify(token, getJWKS(), {
    audience: MANAGED_IDENTITY_AUDIENCE,
    issuer: expectedIssuer,
    algorithms: ["RS256"],
  });

  const claims = payload as MIJwtPayload;
  const roles = claims.roles ?? [];

  if (!roles.includes("MCP.Invoke")) {
    throw new Error(
      `JWT missing required role 'MCP.Invoke'. Roles present: [${roles.join(", ")}]`
    );
  }

  const oid = claims.oid ?? "";
  const appid = claims.appid ?? claims.azp ?? "";

  if (!oid) {
    throw new Error("JWT missing 'oid' claim");
  }

  log(`MI JWT validated: oid=${oid}, appid=${appid}, roles=[${roles.join(", ")}]`);

  return { oid, roles, appid };
}

/**
 * Validates a delegated user / vault Bearer token for Microsoft Graph.
 *
 * Accepts tokens issued by configured Entra apps (default: Graph-MCP-User)
 * targeting the Microsoft Graph API. Same validator handles two callers:
 *   - Local Claude Code users (token from get-user-token.py)
 *   - Anthropic Managed Agents Vaults (token injected as Bearer header)
 *
 * Validation:
 *   - iss matches Majans tenant (v1 sts.windows.net or v2 login.microsoftonline.com)
 *   - aud matches Microsoft Graph (URL or app-ID GUID form)
 *   - appid (or azp) is in VAULT_ALLOWED_APP_IDS
 *   - exp/nbf claims valid (not expired, not future-dated)
 *
 * **Why no signature verification:** Microsoft Graph access tokens cannot be
 * validated by third parties. Microsoft inserts a hashed nonce into the JWT
 * header before signing, with explicit guidance:
 *
 *   "Don't write any code that depends on the ability to validate the
 *    signature of an access token in your API implementations."
 *   — https://learn.microsoft.com/en-us/entra/identity-platform/access-tokens#validate-tokens
 *
 * The downstream Graph API call is the real authentication gate — Graph
 * rejects forged tokens regardless of what we accept here. This validator
 * filters obviously-invalid tokens before they reach Graph (saves a round
 * trip and surfaces a 401 with a useful error) but does not — and per
 * Microsoft cannot — replace Graph's own signature check.
 *
 * @throws if any claim check fails
 */
export async function validateVaultBearerToken(
  token: string
): Promise<ValidatedVaultToken> {
  if (!VAULT_BEARER_AUTH_ENABLED) {
    throw new Error("Vault Bearer auth is not enabled");
  }

  // decodeJwt parses + verifies structure (3 base64url segments, JSON payload),
  // throws on malformed input. It does NOT verify the signature — that's the
  // intentional choice for Graph tokens (see docstring).
  let payload: JWTPayload;
  try {
    payload = decodeJwt(token);
  } catch (err) {
    throw new Error(`Bearer token is not a parseable JWT: ${(err as Error).message}`);
  }

  // Issuer
  const issuerV1 = `https://sts.windows.net/${GRAPH_TENANT_ID}/`;
  const issuerV2 = `https://login.microsoftonline.com/${GRAPH_TENANT_ID}/v2.0`;
  if (payload.iss !== issuerV1 && payload.iss !== issuerV2) {
    throw new Error(
      `Bearer token issuer '${payload.iss}' does not match Majans tenant`
    );
  }

  // Audience — `aud` may be a string or string[].
  const audClaim: string[] = Array.isArray(payload.aud)
    ? payload.aud
    : payload.aud
      ? [payload.aud]
      : [];
  const audOk = audClaim.some((a) => VAULT_EXPECTED_AUDIENCES.has(a));
  if (!audOk) {
    throw new Error(
      `Bearer token audience [${audClaim.join(", ")}] does not match Graph`
    );
  }

  // App ID allow-list
  const claims = payload as VaultJwtPayload;
  const appid = claims.appid ?? claims.azp ?? "";
  if (!VAULT_ALLOWED_APP_IDS.has(appid)) {
    throw new Error(
      `Bearer token appid '${appid || "<missing>"}' not in allow-list [${[...VAULT_ALLOWED_APP_IDS].join(", ")}]`
    );
  }

  // Expiry / not-before
  const nowSec = Math.floor(Date.now() / 1000);
  if (typeof payload.exp === "number" && payload.exp < nowSec) {
    throw new Error("Bearer token has expired");
  }
  if (typeof payload.nbf === "number" && payload.nbf > nowSec + 60) {
    throw new Error("Bearer token nbf is in the future");
  }

  const oid = claims.oid ?? "";
  const upn = claims.upn ?? claims.preferred_username ?? "";

  log(`Vault Bearer auth (claims-only): appid=${appid}, oid=${oid}, upn=${upn}`);

  return { oid, appid, upn };
}
