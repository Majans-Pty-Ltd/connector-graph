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
export const MANAGED_IDENTITY_AUDIENCE = process.env.MANAGED_IDENTITY_AUDIENCE || "api://connector-graph";

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
