import { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { z } from "zod";
import type { GraphClient } from "../api/client.js";
import type { GraphGroup, GraphDirectoryObject, ODataResponse } from "../api/types.js";

export function registerGroupTools(server: McpServer, client: GraphClient): void {
  server.tool(
    "graph_list_groups",
    "List groups in the directory. Supports filtering by mailEnabled, securityEnabled, displayName.",
    {
      filter: z.string().optional().describe("OData $filter (e.g. \"securityEnabled eq true\")"),
      top: z.number().optional().describe("Max results per page (default 100)"),
      all_pages: z.boolean().optional().describe("Fetch all pages (default false)"),
    },
    async ({ filter, top, all_pages }) => {
      try {
        const params: Record<string, string> = {
          $select: "id,displayName,description,mailEnabled,mailNickname,securityEnabled,groupTypes,createdDateTime",
          $top: String(top ?? 100),
          $count: "true",
        };
        if (filter) params.$filter = filter;

        if (all_pages) {
          const groups = await client.getAll<GraphGroup>("groups", params);
          return {
            content: [
              {
                type: "text" as const,
                text: JSON.stringify({ total: groups.length, groups }, null, 2),
              },
            ],
          };
        }

        const result = await client.get<ODataResponse<GraphGroup>>("groups", params);
        return {
          content: [
            {
              type: "text" as const,
              text: JSON.stringify(
                {
                  count: result["@odata.count"],
                  has_more: !!result["@odata.nextLink"],
                  groups: result.value,
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
    "graph_get_group_members",
    "List members of a group.",
    {
      group_id: z.string().describe("Group ID (GUID)"),
      top: z.number().optional().describe("Max results (default 100)"),
    },
    async ({ group_id, top }) => {
      try {
        const params: Record<string, string> = {
          $top: String(top ?? 100),
          $select: "id,displayName,userPrincipalName",
        };
        const result = await client.get<ODataResponse<GraphDirectoryObject>>(
          `groups/${group_id}/members`,
          params
        );
        return {
          content: [
            {
              type: "text" as const,
              text: JSON.stringify(
                { count: result.value.length, members: result.value },
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
    "graph_add_group_member",
    "Add a user to a group.",
    {
      group_id: z.string().describe("Group ID (GUID)"),
      user_id: z.string().describe("User ID (GUID) to add"),
    },
    async ({ group_id, user_id }) => {
      try {
        await client.post(`groups/${group_id}/members/$ref`, {
          "@odata.id": `https://graph.microsoft.com/v1.0/directoryObjects/${user_id}`,
        });
        return {
          content: [
            { type: "text" as const, text: `User ${user_id} added to group ${group_id}.` },
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
    "graph_remove_group_member",
    "Remove a user from a group.",
    {
      group_id: z.string().describe("Group ID (GUID)"),
      user_id: z.string().describe("User ID (GUID) to remove"),
    },
    async ({ group_id, user_id }) => {
      try {
        await client.delete(`groups/${group_id}/members/${user_id}/$ref`);
        return {
          content: [
            { type: "text" as const, text: `User ${user_id} removed from group ${group_id}.` },
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
