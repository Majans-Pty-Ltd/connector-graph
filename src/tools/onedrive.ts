import { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { z } from "zod";
import type { GraphClient } from "../api/client.js";
import type { GraphDriveItem, ODataResponse } from "../api/types.js";

export function registerOneDriveTools(server: McpServer, client: GraphClient): void {
  server.tool(
    "graph_list_drive_items",
    "List files and folders at a OneDrive path for a user. Use '/' or omit path for root.",
    {
      user_id: z.string().describe("User ID (GUID) or UPN (e.g. amit@majans.com)"),
      path: z.string().optional().describe("Folder path (e.g. '/Documents/Invoices'). Omit for root."),
      top: z.number().optional().describe("Max items to return (default 50)"),
    },
    async ({ user_id, path, top }) => {
      try {
        const folderPath = path && path !== "/" ? `/root:${path}:/children` : "/root/children";
        const result = await client.get<ODataResponse<GraphDriveItem>>(
          `users/${encodeURIComponent(user_id)}/drive${folderPath}`,
          {
            $select: "id,name,size,lastModifiedDateTime,webUrl,folder,file",
            $top: String(top ?? 50),
            $orderby: "name",
          }
        );
        return {
          content: [
            {
              type: "text" as const,
              text: JSON.stringify(
                {
                  count: result.value.length,
                  has_more: !!result["@odata.nextLink"],
                  items: result.value.map((item) => ({
                    ...item,
                    type: item.folder ? "folder" : "file",
                  })),
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
    "graph_get_drive_item_content",
    "Get metadata and download URL for a specific OneDrive file by path.",
    {
      user_id: z.string().describe("User ID (GUID) or UPN"),
      path: z.string().describe("File path (e.g. '/Documents/Invoices/invoice-001.pdf')"),
    },
    async ({ user_id, path }) => {
      try {
        const result = await client.get<GraphDriveItem>(
          `users/${encodeURIComponent(user_id)}/drive/root:${path}`,
          {
            $select: "id,name,size,lastModifiedDateTime,webUrl,file,@microsoft.graph.downloadUrl",
          }
        );
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
}
