#!/usr/bin/env node
import { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { StdioServerTransport } from "@modelcontextprotocol/sdk/server/stdio.js";
import { StreamableHTTPServerTransport } from "@modelcontextprotocol/sdk/server/streamableHttp.js";
import { validateConfig, GRAPH_CLIENT_ID, GRAPH_API_KEY, MANAGED_IDENTITY_ENABLED } from "./utils/config.js";
import { log, logError } from "./utils/logger.js";
import { userTokenStorage } from "./utils/auth.js";
import { looksLikeJwt, validateManagedIdentityToken } from "./utils/jwt-validator.js";
import { GraphClient, isDelegatedAuth } from "./api/client.js";
import { registerUserTools } from "./tools/users.js";
import { registerGroupTools } from "./tools/groups.js";
import { registerLicenseTools } from "./tools/licenses.js";
import { registerMailTools } from "./tools/mail.js";
import { registerOneDriveTools } from "./tools/onedrive.js";
import { registerCalendarTools } from "./tools/calendar.js";
import { registerSharePointTools } from "./tools/sharepoint.js";
import { registerPlannerTools } from "./tools/planner.js";
import { registerTodoTools } from "./tools/todo.js";
import { registerTeamsTools } from "./tools/teams.js";
import { createServer, type IncomingMessage, type ServerResponse } from "http";

function createMcpServer(): McpServer {
  const client = new GraphClient();

  const server = new McpServer({
    name: "graph",
    version: "2.0.0",
  });

  // Auth status tool — reports whether running as delegated user or SP
  server.tool(
    "graph_auth_status",
    "Verify Graph API authentication is working. Returns auth mode (delegated/app-only), tenant info, and client ID.",
    {},
    async () => {
      try {
        const delegated = isDelegatedAuth();

        if (delegated) {
          // With delegated token, call /me to identify the user
          const me = await client.get<{ displayName: string; mail: string; id: string }>(
            "me",
            { $select: "displayName,mail,id" }
          );
          return {
            content: [
              {
                type: "text" as const,
                text: JSON.stringify(
                  {
                    authenticated: true,
                    auth_mode: "delegated",
                    user: me.displayName,
                    email: me.mail,
                    user_id: me.id,
                  },
                  null,
                  2
                ),
              },
            ],
          };
        }

        // App-only: list organization
        const result = await client.get<{ value: Array<{ displayName: string; id: string }> }>(
          "organization",
          { $select: "displayName,id" }
        );
        const org = result.value[0];
        return {
          content: [
            {
              type: "text" as const,
              text: JSON.stringify(
                {
                  authenticated: true,
                  auth_mode: "app-only",
                  client_id: GRAPH_CLIENT_ID,
                  organization: org?.displayName ?? "Unknown",
                  tenant_id: org?.id ?? "Unknown",
                },
                null,
                2
              ),
            },
          ],
        };
      } catch (err) {
        return {
          content: [
            {
              type: "text" as const,
              text: JSON.stringify(
                { authenticated: false, error: (err as Error).message },
                null,
                2
              ),
            },
          ],
          isError: true,
        };
      }
    }
  );

  // Register all tool categories
  registerUserTools(server, client);
  registerGroupTools(server, client);
  registerLicenseTools(server, client);
  registerMailTools(server, client);
  registerOneDriveTools(server, client);
  registerCalendarTools(server, client);
  registerSharePointTools(server, client);
  registerPlannerTools(server, client);
  registerTodoTools(server, client);
  registerTeamsTools(server, client);

  return server;
}

/**
 * Extract Bearer token from Authorization header.
 * Returns undefined if no Bearer token is present.
 */
function extractBearerToken(req: IncomingMessage): string | undefined {
  const authHeader = req.headers.authorization ?? "";
  if (authHeader.toLowerCase().startsWith("bearer ")) {
    return authHeader.slice(7);
  }
  return undefined;
}

/**
 * Validate API key from X-API-Key header (for agent /call-tool path).
 * Returns true if GRAPH_API_KEY is not configured (disabled) or if the key matches.
 */
function validateApiKey(req: IncomingMessage): boolean {
  if (!GRAPH_API_KEY) return true; // API key auth disabled
  const provided = req.headers["x-api-key"] as string | undefined;
  return provided === GRAPH_API_KEY;
}

async function startStdio(): Promise<void> {
  const server = createMcpServer();
  const transport = new StdioServerTransport();
  await server.connect(transport);
  log("Graph MCP server started (stdio)");
}

async function startHttp(port: number): Promise<void> {
  const httpServer = createServer(async (req, res) => {
    // Health check
    if (req.method === "GET" && req.url === "/health") {
      res.writeHead(200, { "Content-Type": "application/json" });
      res.end(JSON.stringify({ status: "ok", transport: "streamable-http" }));
      return;
    }

    // MCP endpoint — supports delegated (Bearer), Managed Identity (Bearer JWT), and app-only auth
    if (req.url === "/mcp" || req.url === "/") {
      const bearerToken = extractBearerToken(req);
      let userToken: string | undefined = undefined;

      if (bearerToken) {
        // Bearer token present — determine if it's an MI JWT or a delegated user token
        if (MANAGED_IDENTITY_ENABLED && looksLikeJwt(bearerToken)) {
          try {
            await validateManagedIdentityToken(bearerToken);
            log("MCP request authenticated via Managed Identity JWT");
            // MI path: do NOT store token in userTokenStorage — GraphClient uses SP fallback
          } catch (miErr) {
            // MI validation failed — check if this was clearly an MI token (not a delegated user token)
            // MI tokens have `roles` claim and no `scp` claim; delegated tokens have `scp` claim
            let isMiToken = false;
            try {
              const payloadB64 = bearerToken.split(".")[1];
              const payload = JSON.parse(Buffer.from(payloadB64, "base64url").toString());
              isMiToken = payload.roles && !payload.scp;
            } catch {
              // Can't decode — treat as opaque delegated token
            }

            if (isMiToken) {
              // Token was clearly an MI token but failed validation — check for API key fallback
              if (validateApiKey(req)) {
                log("MI JWT validation failed, falling back to valid X-API-Key auth");
              } else {
                logError("MI JWT validation failed and no valid X-API-Key", miErr as Error);
                res.writeHead(401, { "Content-Type": "application/json" });
                res.end(JSON.stringify({ error: "Invalid Managed Identity token and no valid X-API-Key" }));
                return;
              }
            } else {
              // Token looks like a delegated user token (has `scp` or undecodable) — existing path
              log("Bearer token is not an MI JWT, treating as delegated user token");
              userToken = bearerToken;
            }
          }
        } else {
          // MI not enabled or token doesn't look like a JWT — delegated user token
          userToken = bearerToken;
        }

        if (userToken) {
          log("MCP request with delegated user token");
        }
      } else {
        // No Bearer token — require API key
        if (!validateApiKey(req)) {
          res.writeHead(401, { "Content-Type": "application/json" });
          res.end(JSON.stringify({ error: "Invalid or missing X-API-Key header" }));
          return;
        }
      }

      // Run MCP handler inside AsyncLocalStorage context with the user's token
      // For MI auth: userToken is undefined, so GraphClient falls back to SP token
      await userTokenStorage.run(userToken, async () => {
        const transport = new StreamableHTTPServerTransport({
          sessionIdGenerator: undefined,
        });
        const server = createMcpServer();

        res.on("close", () => {
          transport.close().catch(() => {});
          server.close().catch(() => {});
        });

        await server.connect(transport);
        await transport.handleRequest(req, res);
      });
      return;
    }

    // 404 for everything else
    res.writeHead(404, { "Content-Type": "application/json" });
    res.end(JSON.stringify({ error: "Not found. Use /mcp for MCP or /health for health check." }));
  });

  httpServer.listen(port, () => {
    log(`Graph MCP server started (StreamableHTTP on port ${port})`);
    log(`  MCP endpoint: http://localhost:${port}/mcp`);
    log(`  Health check: http://localhost:${port}/health`);
    log(`  Auth modes: delegated (Bearer token) + app-only (X-API-Key / SP)`);
  });
}

async function main(): Promise<void> {
  validateConfig();

  const port = process.env.PORT ? parseInt(process.env.PORT, 10) : undefined;

  if (port) {
    await startHttp(port);
  } else {
    await startStdio();
  }
}

main().catch((err) => {
  logError("Failed to start MCP server", err);
  process.exit(1);
});
