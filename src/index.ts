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
import { z, type ZodRawShape } from "zod";

// ── Tool handler registry (populated by intercepting server.tool() calls) ────
// Stores each tool's zod schema + handler so /call-tool can invoke tools
// directly without going through the MCP protocol.
type ToolHandler = (args: Record<string, unknown>) => Promise<any>;
interface RegisteredTool {
  schema: ZodRawShape;
  handler: ToolHandler;
}
const toolHandlers = new Map<string, RegisteredTool>();

function createMcpServer(): McpServer {
  const client = new GraphClient();

  const server = new McpServer({
    name: "graph",
    version: "2.0.0",
  });

  // Intercept server.tool() to capture registrations in toolHandlers
  const originalTool = server.tool.bind(server);
  (server as any).tool = (name: string, description: string, schema: ZodRawShape, handler: ToolHandler) => {
    toolHandlers.set(name, { schema, handler });
    return originalTool(name, description, schema, handler);
  };

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
  // Create once at startup to populate the toolHandlers registry.
  // Per-request MCP sessions create their own servers, but the registry
  // is module-level and only needs to be filled once.
  createMcpServer();

  const httpServer = createServer(async (req, res) => {
    // Health check
    if (req.method === "GET" && req.url === "/health") {
      res.writeHead(200, { "Content-Type": "application/json" });
      res.end(JSON.stringify({ status: "ok", transport: "streamable-http" }));
      return;
    }

    // ── GET /tools — list all registered tool names ──────────────────────
    if (req.method === "GET" && req.url === "/tools") {
      res.writeHead(200, { "Content-Type": "application/json" });
      res.end(JSON.stringify({ tools: Array.from(toolHandlers.keys()) }));
      return;
    }

    // ── POST /call-tool — invoke a tool directly via REST ──────────────
    if (req.method === "POST" && req.url === "/call-tool") {
      // Auth: require Bearer token (delegated) OR X-API-Key (app-only)
      const bearerToken = extractBearerToken(req);
      let userToken: string | undefined = undefined;

      if (bearerToken) {
        userToken = bearerToken;
      } else if (!validateApiKey(req)) {
        res.writeHead(401, { "Content-Type": "application/json" });
        res.end(JSON.stringify({ error: "Invalid or missing auth. Provide Authorization: Bearer <token> or X-API-Key header." }));
        return;
      }

      // Parse request body
      let body: string;
      try {
        body = await new Promise<string>((resolve, reject) => {
          const chunks: Buffer[] = [];
          req.on("data", (chunk: Buffer) => chunks.push(chunk));
          req.on("end", () => resolve(Buffer.concat(chunks).toString("utf-8")));
          req.on("error", reject);
        });
      } catch {
        res.writeHead(400, { "Content-Type": "application/json" });
        res.end(JSON.stringify({ error: "Failed to read request body" }));
        return;
      }

      let parsed: { name?: string; arguments?: Record<string, unknown> };
      try {
        parsed = JSON.parse(body);
      } catch {
        res.writeHead(400, { "Content-Type": "application/json" });
        res.end(JSON.stringify({ error: "Invalid JSON in request body" }));
        return;
      }

      const toolName = parsed.name;
      const toolArgs = parsed.arguments ?? {};

      if (!toolName || typeof toolName !== "string") {
        res.writeHead(400, { "Content-Type": "application/json" });
        res.end(JSON.stringify({ error: "Missing or invalid 'name' field in request body" }));
        return;
      }

      const registered = toolHandlers.get(toolName);
      if (!registered) {
        res.writeHead(404, { "Content-Type": "application/json" });
        res.end(JSON.stringify({ error: `Unknown tool: ${toolName}`, available: Array.from(toolHandlers.keys()) }));
        return;
      }

      // Validate arguments against the tool's zod schema
      let validatedArgs: Record<string, unknown>;
      try {
        const zodSchema = z.object(registered.schema);
        validatedArgs = zodSchema.parse(toolArgs);
      } catch (err) {
        res.writeHead(400, { "Content-Type": "application/json" });
        res.end(JSON.stringify({ error: "Invalid tool arguments", details: (err as Error).message }));
        return;
      }

      // Execute the tool inside the correct auth context
      try {
        const result = await userTokenStorage.run(userToken, async () => {
          return await registered.handler(validatedArgs);
        });
        res.writeHead(200, { "Content-Type": "application/json" });
        res.end(JSON.stringify(result));
      } catch (err) {
        res.writeHead(500, { "Content-Type": "application/json" });
        res.end(JSON.stringify({
          content: [{ type: "text", text: `Tool execution error: ${(err as Error).message}` }],
          isError: true,
        }));
      }
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
    res.end(JSON.stringify({ error: "Not found. Use /mcp for MCP, /call-tool for REST, /tools for tool list, or /health for health check." }));
  });

  httpServer.listen(port, () => {
    log(`Graph MCP server started (StreamableHTTP on port ${port})`);
    log(`  MCP endpoint: http://localhost:${port}/mcp`);
    log(`  REST /tools:  http://localhost:${port}/tools`);
    log(`  REST /call-tool: http://localhost:${port}/call-tool`);
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
