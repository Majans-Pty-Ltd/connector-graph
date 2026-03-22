import { createRemoteJWKSet, jwtVerify, type JWTPayload } from "jose";
import { GRAPH_TENANT_ID, MANAGED_IDENTITY_ENABLED, MANAGED_IDENTITY_AUDIENCE } from "./config.js";
import { log, logError } from "./logger.js";

export interface ValidatedToken {
  oid: string;
  roles: string[];
  appid: string;
}

interface MIJwtPayload extends JWTPayload {
  oid?: string;
  roles?: string[];
  appid?: string;
  azp?: string;
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
 *   aud: api://connector-graph (or MANAGED_IDENTITY_AUDIENCE)
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
