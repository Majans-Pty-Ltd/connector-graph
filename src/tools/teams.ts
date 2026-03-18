import { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { z } from "zod";
import type { GraphClient } from "../api/client.js";
import type { GraphChat, GraphChatMessage, ODataResponse } from "../api/types.js";

export function registerTeamsTools(server: McpServer, client: GraphClient): void {
  server.tool(
    "graph_list_chats",
    "List a user's Teams chats (1:1, group, and meeting chats). Returns chat ID, type, topic, and last updated time.",
    {
      user_id: z.string().describe("User ID (GUID) or UPN (e.g. amit@majans.com)"),
      chat_type: z
        .enum(["oneOnOne", "group", "meeting"])
        .optional()
        .describe("Filter by chat type"),
      top: z.number().optional().describe("Max results (default 50)"),
    },
    async ({ user_id, chat_type, top }) => {
      try {
        const params: Record<string, string> = {
          $top: String(top ?? 50),
          $orderby: "lastUpdatedDateTime desc",
          $select: "id,topic,chatType,createdDateTime,lastUpdatedDateTime,webUrl",
        };
        if (chat_type) {
          params.$filter = `chatType eq '${chat_type}'`;
        }

        const result = await client.get<ODataResponse<GraphChat>>(
          `users/${encodeURIComponent(user_id)}/chats`,
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
                  chats: result.value,
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
    "graph_list_chat_messages",
    "List messages in a Teams chat. Returns message content, sender, and timestamps. Supports pagination.",
    {
      chat_id: z.string().describe("Chat ID (from graph_list_chats)"),
      top: z.number().optional().describe("Max messages to return (default 50)"),
      filter: z.string().optional().describe("OData $filter (e.g. filter by date)"),
    },
    async ({ chat_id, top, filter }) => {
      try {
        const params: Record<string, string> = {
          $top: String(top ?? 50),
          $orderby: "createdDateTime desc",
        };
        if (filter) params.$filter = filter;

        const result = await client.get<ODataResponse<GraphChatMessage>>(
          `chats/${encodeURIComponent(chat_id)}/messages`,
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
                  messages: result.value.map((msg) => ({
                    id: msg.id,
                    createdDateTime: msg.createdDateTime,
                    messageType: msg.messageType,
                    from: msg.from?.user?.displayName
                      ?? msg.from?.application?.displayName
                      ?? "Unknown",
                    body: msg.body.content,
                    bodyType: msg.body.contentType,
                    importance: msg.importance,
                    attachmentCount: msg.attachments?.length ?? 0,
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
    "graph_send_chat_message",
    "Send a message to a Teams chat. Supports HTML or plain text content.",
    {
      chat_id: z.string().describe("Chat ID (from graph_list_chats)"),
      content: z.string().describe("Message content (HTML or plain text)"),
      content_type: z.enum(["html", "text"]).optional().describe("Content type (default html)"),
    },
    async ({ chat_id, content, content_type }) => {
      try {
        const result = await client.post<GraphChatMessage>(
          `chats/${encodeURIComponent(chat_id)}/messages`,
          {
            body: {
              contentType: content_type ?? "html",
              content,
            },
          }
        );
        return {
          content: [
            {
              type: "text" as const,
              text: JSON.stringify(
                {
                  sent: true,
                  messageId: result.id,
                  chatId: chat_id,
                  createdDateTime: result.createdDateTime,
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
