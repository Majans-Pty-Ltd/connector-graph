import { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { z } from "zod";
import type { GraphClient } from "../api/client.js";
import type { GraphMailMessage, GraphMailMessageFull, GraphAttachment, ODataResponse } from "../api/types.js";

const MAIL_SELECT = "id,subject,bodyPreview,from,toRecipients,receivedDateTime,isRead,hasAttachments,importance";

export function registerMailTools(server: McpServer, client: GraphClient): void {
  server.tool(
    "graph_search_mail",
    "Search a user's mailbox by query string. Supports OData $filter and $search for finding emails by subject, sender, date, etc.",
    {
      user_id: z.string().describe("User ID (GUID) or UPN (e.g. amit@majans.com)"),
      search: z.string().optional().describe("Free-text search query (e.g. 'Coral Enterprises invoice')"),
      filter: z.string().optional().describe("OData $filter (e.g. \"receivedDateTime ge 2024-01-01\")"),
      top: z.number().optional().describe("Max results to return (default 25, max 100)"),
      orderby: z.string().optional().describe("OData $orderby (default 'receivedDateTime desc')"),
    },
    async ({ user_id, search, filter, top, orderby }) => {
      try {
        const params: Record<string, string> = {
          $select: MAIL_SELECT,
          $top: String(top ?? 25),
          $orderby: orderby ?? "receivedDateTime desc",
        };
        if (search) params.$search = `"${search}"`;
        if (filter) params.$filter = filter;

        const result = await client.get<ODataResponse<GraphMailMessage>>(
          `users/${encodeURIComponent(user_id)}/messages`,
          params
        );
        return {
          content: [
            {
              type: "text" as const,
              text: JSON.stringify(
                {
                  count: result.value.length,
                  has_more: !!result["@odata.nextLink"],
                  messages: result.value,
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
    "graph_read_mail",
    "Get full email content (including body) by message ID.",
    {
      user_id: z.string().describe("User ID (GUID) or UPN"),
      message_id: z.string().describe("Message ID from search results"),
    },
    async ({ user_id, message_id }) => {
      try {
        const result = await client.get<GraphMailMessageFull>(
          `users/${encodeURIComponent(user_id)}/messages/${encodeURIComponent(message_id)}`,
          {
            $select: `${MAIL_SELECT},body,ccRecipients`,
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
    "graph_list_attachments",
    "List attachments on an email message.",
    {
      user_id: z.string().describe("User ID (GUID) or UPN"),
      message_id: z.string().describe("Message ID to list attachments for"),
    },
    async ({ user_id, message_id }) => {
      try {
        const result = await client.get<ODataResponse<GraphAttachment>>(
          `users/${encodeURIComponent(user_id)}/messages/${encodeURIComponent(message_id)}/attachments`,
          {
            $select: "id,name,contentType,size,isInline",
          }
        );
        return {
          content: [
            {
              type: "text" as const,
              text: JSON.stringify(
                { count: result.value.length, attachments: result.value },
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
