import { AsyncLocalStorage } from "node:async_hooks";
import type { IncomingMessage } from "http";
import {
  GRAPH_API_KEY,
  MANAGED_IDENTITY_ENABLED,
  VAULT_BEARER_AUTH_ENABLED,
} from "./config.js";
import { log, logError } from "./logger.js";
import { looksLikeJwt, validateManagedIdentityToken, validateVaultBearerToken } from "./jwt-validator.js";

/**
 * Per-request user token storage.
 *
 * When a user connects via /mcp with a Bearer token, the HTTP handler
 * stores it here. GraphClient checks this before falling back to the
 * Service Principal token, so Graph API enforces the user's own permissions.
 *
 * Node.js equivalent of Python's contextvars.ContextVar pattern used
 * in connector-fabric.
 */
export const userTokenStorage = new AsyncLocalStorage<string | undefined>();

/** Returns the current request's user token, or undefined if using SP auth. */
export function getUserToken(): string | undefined {
  return userTokenStorage.getStore();
}

/** Outcome of an auth check. See `authenticateRequest` precedence rules. */
export interface AuthCheckResult {
  /** True if the request may proceed. */
  allowed: boolean;
  /** When set, downstream Graph calls should run with this user token (per-user
   *  permissions). When undefined, run on the SP path. */
  userToken?: string;
  /** When allowed=false, a short error message safe to include in 401 response. */
  error?: string;
}

/** Extract Bearer token from Authorization header, or undefined if absent. */
function extractBearerToken(req: IncomingMessage): string | undefined {
  const authHeader = req.headers.authorization ?? "";
  if (authHeader.toLowerCase().startsWith("bearer ")) {
    return authHeader.slice(7);
  }
  return undefined;
}

/**
 * Apply the connector's auth precedence rules. Used by /mcp and /call-tool.
 *
 * Precedence (highest first):
 *   1. X-API-Key  → SP path. Must match server key if header present —
 *                   no silent fall-through to Bearer when key is wrong.
 *   2. Bearer MI JWT (when MANAGED_IDENTITY_ENABLED) → SP path.
 *   3. Bearer vault/user token → user path. When VAULT_BEARER_AUTH_ENABLED,
 *                   token is signature-checked against Microsoft's JWKS
 *                   and its `iss`/`aud`/`appid` claims must match.
 *   4. No auth + no API key configured → dev mode (allow as SP).
 *   5. Otherwise → reject.
 */
export async function authenticateRequest(
  req: IncomingMessage
): Promise<AuthCheckResult> {
  // 1. X-API-Key takes precedence — must match if present.
  const apiKeyHeader = req.headers["x-api-key"] as string | undefined;
  if (apiKeyHeader) {
    if (!GRAPH_API_KEY || apiKeyHeader === GRAPH_API_KEY) {
      return { allowed: true };
    }
    return { allowed: false, error: "Invalid X-API-Key" };
  }

  // 2 + 3. Bearer token paths.
  const bearerToken = extractBearerToken(req);
  if (bearerToken) {
    // 2. MI JWT first (only attempted on JWT-shaped tokens).
    if (MANAGED_IDENTITY_ENABLED && looksLikeJwt(bearerToken)) {
      try {
        await validateManagedIdentityToken(bearerToken);
        log("MI JWT auth succeeded");
        return { allowed: true };
      } catch (err) {
        // Fall through to vault validation. Distinguish "definitely an MI
        // token that failed" from "looks like a delegated user token" by
        // the role/scp claim shape — vault validation will accept the
        // latter and reject the former.
      }
    }

    // 3. Vault / delegated user token. Strict validation when enabled.
    if (VAULT_BEARER_AUTH_ENABLED) {
      try {
        await validateVaultBearerToken(bearerToken);
      } catch (err) {
        logError(`Bearer token validation failed: ${(err as Error).message}`);
        return {
          allowed: false,
          error:
            "Bearer token failed validation (issuer/audience/appid/signature)",
        };
      }
    }

    return { allowed: true, userToken: bearerToken };
  }

  // 4. No auth at all — dev mode if no API key configured.
  if (!GRAPH_API_KEY) {
    return { allowed: true };
  }

  // 5. Otherwise reject.
  return { allowed: false, error: "provide Bearer token or X-API-Key" };
}
