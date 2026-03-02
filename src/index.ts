#!/usr/bin/env node
import { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { StdioServerTransport } from "@modelcontextprotocol/sdk/server/stdio.js";
import { validateConfig, GRAPH_CLIENT_ID } from "./utils/config.js";
import { log, logError } from "./utils/logger.js";
import { GraphClient } from "./api/client.js";
import { registerUserTools } from "./tools/users.js";
import { registerGroupTools } from "./tools/groups.js";
import { registerLicenseTools } from "./tools/licenses.js";

async function main(): Promise<void> {
  validateConfig();

  const client = new GraphClient();

  const server = new McpServer({
    name: "graph",
    version: "1.0.0",
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

  registerUserTools(server, client);
  registerGroupTools(server, client);
  registerLicenseTools(server, client);

  const transport = new StdioServerTransport();
  await server.connect(transport);
  log("Graph MCP server started (stdio)");
}

main().catch((err) => {
  logError("Failed to start MCP server", err);
  process.exit(1);
});
