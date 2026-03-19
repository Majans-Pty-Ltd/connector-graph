import { AsyncLocalStorage } from "node:async_hooks";

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
