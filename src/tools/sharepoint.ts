import { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { z } from "zod";
import type { GraphClient } from "../api/client.js";
import type { GraphSite, GraphDrive, GraphDriveItem, GraphSearchResult, ODataResponse } from "../api/types.js";

export function registerSharePointTools(server: McpServer, client: GraphClient): void {
  server.tool(
    "graph_list_sites",
    "Search SharePoint sites by keyword. Returns matching sites with ID, name, URL, and description.",
    {
      search: z.string().describe("Search keyword to find sites (e.g. 'Intranet', 'HR')"),
      top: z.number().optional().describe("Max results to return (default 25)"),
    },
    async ({ search, top }) => {
      try {
        const result = await client.get<ODataResponse<GraphSite>>(
          "sites",
          {
            $search: `"${search}"`,
            $select: "id,displayName,name,webUrl,description,createdDateTime,lastModifiedDateTime",
            $top: String(top ?? 25),
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
                  sites: result.value,
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
    "graph_get_site",
    "Get a SharePoint site by ID. Returns full site details including webUrl and description.",
    {
      site_id: z.string().describe("Site ID (e.g. 'contoso.sharepoint.com,guid,guid' or just the site path like 'root')"),
    },
    async ({ site_id }) => {
      try {
        const result = await client.get<GraphSite>(
          `sites/${encodeURIComponent(site_id)}`,
          {
            $select: "id,displayName,name,webUrl,description,createdDateTime,lastModifiedDateTime,siteCollection",
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

  server.tool(
    "graph_list_site_drives",
    "List document libraries (drives) in a SharePoint site.",
    {
      site_id: z.string().describe("Site ID"),
      top: z.number().optional().describe("Max results to return (default 50)"),
    },
    async ({ site_id, top }) => {
      try {
        const result = await client.get<ODataResponse<GraphDrive>>(
          `sites/${encodeURIComponent(site_id)}/drives`,
          {
            $select: "id,name,driveType,webUrl,description,createdDateTime,lastModifiedDateTime,quota",
            $top: String(top ?? 50),
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
                  drives: result.value,
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
    "graph_list_site_drive_items",
    "List files and folders in a SharePoint site's document library. Use '/' or omit path for root.",
    {
      site_id: z.string().describe("Site ID"),
      drive_id: z.string().optional().describe("Drive ID (document library). Omit for the site's default drive."),
      path: z.string().optional().describe("Folder path (e.g. '/General/Reports'). Omit for root."),
      top: z.number().optional().describe("Max items to return (default 50)"),
    },
    async ({ site_id, drive_id, path, top }) => {
      try {
        const folderPath = path && path !== "/" ? `/root:${path}:/children` : "/root/children";
        const drivePart = drive_id
          ? `sites/${encodeURIComponent(site_id)}/drives/${encodeURIComponent(drive_id)}`
          : `sites/${encodeURIComponent(site_id)}/drive`;

        const result = await client.get<ODataResponse<GraphDriveItem>>(
          `${drivePart}${folderPath}`,
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
    "graph_get_site_file_content",
    "Get metadata and download URL for a file in a SharePoint site's document library by path.",
    {
      site_id: z.string().describe("Site ID"),
      drive_id: z.string().optional().describe("Drive ID (document library). Omit for the site's default drive."),
      path: z.string().describe("File path (e.g. '/General/Reports/Q1-2026.xlsx')"),
    },
    async ({ site_id, drive_id, path }) => {
      try {
        const drivePart = drive_id
          ? `sites/${encodeURIComponent(site_id)}/drives/${encodeURIComponent(drive_id)}`
          : `sites/${encodeURIComponent(site_id)}/drive`;

        const result = await client.get<GraphDriveItem>(
          `${drivePart}/root:${path}`,
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

  server.tool(
    "graph_search_site_files",
    "Search for files within a SharePoint site by keyword. Returns matching files with metadata.",
    {
      site_id: z.string().describe("Site ID"),
      query: z.string().describe("Search query (e.g. 'budget 2026')"),
      top: z.number().optional().describe("Max results (default 25)"),
    },
    async ({ site_id, query, top }) => {
      try {
        const result = await client.get<ODataResponse<GraphSearchResult>>(
          `sites/${encodeURIComponent(site_id)}/drive/root/search(q='${encodeURIComponent(query)}')`,
          {
            $top: String(top ?? 25),
            $select: "id,name,size,webUrl,lastModifiedDateTime,file,folder,parentReference",
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
                  results: result.value.map((item) => ({
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
}
