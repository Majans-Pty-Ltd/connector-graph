#!/usr/bin/env node
import { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { StdioServerTransport } from "@modelcontextprotocol/sdk/server/stdio.js";
import { StreamableHTTPServerTransport } from "@modelcontextprotocol/sdk/server/streamableHttp.js";
import { validateConfig, GRAPH_CLIENT_ID } from "./utils/config.js";
import { log, logError } from "./utils/logger.js";
import { GraphClient } from "./api/client.js";
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
import { createServer } from "http";

function createMcpServer(): McpServer {
  const client = new GraphClient();

  const server = new McpServer({
    name: "graph",
    version: "2.0.0",
  });

  // Auth status tool
  server.tool(
    "graph_auth_status",
    "Verify Graph API authentication is working. Returns tenant info and client ID.",
    {},
    async () => {
      try {
        // Quick test: list organization
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

    // MCP endpoint
    if (req.url === "/mcp" || req.url === "/") {
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
