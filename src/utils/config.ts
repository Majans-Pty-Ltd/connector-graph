import { config } from "dotenv";
import { resolve, dirname } from "path";
import { fileURLToPath } from "url";

const __dirname = dirname(fileURLToPath(import.meta.url));
const projectRoot = resolve(__dirname, "..", "..");

config({ path: resolve(projectRoot, ".env") });

export const GRAPH_TENANT_ID = process.env.GRAPH_TENANT_ID ?? "";
export const GRAPH_CLIENT_ID = process.env.GRAPH_CLIENT_ID ?? "";
export const GRAPH_CLIENT_SECRET = process.env.GRAPH_CLIENT_SECRET ?? "";
export const GRAPH_API_KEY = process.env.GRAPH_API_KEY ?? "";
export const GRAPH_API_BASE = "https://graph.microsoft.com/v1.0/";

// Managed Identity JWT auth (agent-to-connector via Azure MI)
export const MANAGED_IDENTITY_ENABLED = process.env.MANAGED_IDENTITY_ENABLED === "true";
export const MANAGED_IDENTITY_AUDIENCE = process.env.MANAGED_IDENTITY_AUDIENCE || "api://b8157430-c8f3-4760-beaa-a4b95cfc20a7";

// Vault / delegated user Bearer token validation.
//
// Default ON: tokens issued to Graph-MCP-User (the only delegated app we
// support today) pass validation, so existing local Claude Code users are
// unaffected. Set VAULT_BEARER_AUTH_ENABLED=false to fall back to legacy
// Bearer passthrough — debug only.
export const VAULT_BEARER_AUTH_ENABLED =
  (process.env.VAULT_BEARER_AUTH_ENABLED ?? "true").toLowerCase() === "true";

// Allowed Entra app IDs that may issue Bearer tokens accepted by this connector.
// Comma-separated list, defaults to Graph-MCP-User (the public-client delegated app).
// Override via env var to add additional vault-credential apps.
const _DEFAULT_VAULT_APP_IDS = "02fa0ea1-4b30-4bd9-9c4a-483f97d63b21";
export const VAULT_ALLOWED_APP_IDS: ReadonlySet<string> = new Set(
  (process.env.VAULT_ALLOWED_APP_IDS ?? _DEFAULT_VAULT_APP_IDS)
    .split(",")
    .map((s) => s.trim())
    .filter((s) => s.length > 0)
);

// Microsoft Graph audience as it appears in tokens issued by Azure AD.
// The .default scope and per-permission scopes both yield audience
// "https://graph.microsoft.com" (no trailing slash) or with trailing slash.
export const VAULT_EXPECTED_AUDIENCES: ReadonlySet<string> = new Set([
  "https://graph.microsoft.com",
  "https://graph.microsoft.com/",
  "00000003-0000-0000-c000-000000000000", // Graph app ID (sometimes appears as aud)
]);

export function validateConfig(): void {
  if (!GRAPH_TENANT_ID) {
    throw new Error("GRAPH_TENANT_ID not set. Copy .env.template to .env and fill in credentials.");
  }
  if (!GRAPH_CLIENT_ID) {
    throw new Error("GRAPH_CLIENT_ID not set.");
  }
  if (!GRAPH_CLIENT_SECRET) {
    throw new Error("GRAPH_CLIENT_SECRET not set.");
  }
}
