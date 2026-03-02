import { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { z } from "zod";
import type { GraphClient } from "../api/client.js";
import type { GraphUser, ODataResponse } from "../api/types.js";

const USER_SELECT = "id,displayName,userPrincipalName,mail,jobTitle,department,officeLocation,accountEnabled,createdDateTime";

export function registerUserTools(server: McpServer, client: GraphClient): void {
  server.tool(
    "graph_list_users",
    "List users in the directory. Supports OData $filter, $select, $top, $orderby, and pagination.",
    {
      filter: z.string().optional().describe("OData $filter (e.g. \"department eq 'IT'\")"),
      select: z.string().optional().describe("Comma-separated fields to return"),
      top: z.number().optional().describe("Max results per page (default 100, max 999)"),
      orderby: z.string().optional().describe("OData $orderby (e.g. \"displayName\")"),
      all_pages: z.boolean().optional().describe("Fetch all pages (default false)"),
    },
    async ({ filter, select, top, orderby, all_pages }) => {
      try {
        const params: Record<string, string> = {
          $select: select || USER_SELECT,
          $top: String(top ?? 100),
          $count: "true",
        };
        if (filter) params.$filter = filter;
        if (orderby) params.$orderby = orderby;

        if (all_pages) {
          const users = await client.getAll<GraphUser>("/users", params);
          return {
            content: [
              {
                type: "text" as const,
                text: JSON.stringify({ total: users.length, users }, null, 2),
              },
            ],
          };
        }

        const result = await client.get<ODataResponse<GraphUser>>("/users", params);
        return {
          content: [
            {
              type: "text" as const,
              text: JSON.stringify(
                {
                  count: result["@odata.count"],
                  has_more: !!result["@odata.nextLink"],
                  users: result.value,
                },
                null,
                2
              ),
            },
          ],
        };
      } catch (err) {
        return {
          content: [{ type: "text" as const, text: `Error: ${(err as Error).message}` }],
          isError: true,
        };
      }
    }
  );

  server.tool(
    "graph_get_user",
    "Get full details of a specific user by ID or UPN. Includes signInActivity if available.",
    {
      user_id: z.string().describe("User ID (GUID) or userPrincipalName (e.g. amit@majans.com)"),
    },
    async ({ user_id }) => {
      try {
        const result = await client.get<GraphUser>(`/users/${encodeURIComponent(user_id)}`, {
          $select: `${USER_SELECT},signInActivity`,
        });
        return {
          content: [{ type: "text" as const, text: JSON.stringify(result, null, 2) }],
        };
      } catch (err) {
        return {
          content: [{ type: "text" as const, text: `Error: ${(err as Error).message}` }],
          isError: true,
        };
      }
    }
  );

  server.tool(
    "graph_update_user",
    "Update cloud-only user properties (jobTitle, department, officeLocation, accountEnabled, etc.).",
    {
      user_id: z.string().describe("User ID (GUID) or UPN"),
      properties: z
        .record(z.unknown())
        .describe("Properties to update (e.g. {\"jobTitle\": \"Manager\", \"accountEnabled\": false})"),
    },
    async ({ user_id, properties }) => {
      try {
        await client.patch(`/users/${encodeURIComponent(user_id)}`, properties);
        return {
          content: [
            {
              type: "text" as const,
              text: `User ${user_id} updated successfully. Properties changed: ${Object.keys(properties).join(", ")}`,
            },
          ],
        };
      } catch (err) {
        return {
          content: [{ type: "text" as const, text: `Error: ${(err as Error).message}` }],
          isError: true,
        };
      }
    }
  );
}
