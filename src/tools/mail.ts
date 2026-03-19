import { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { z } from "zod";
import type { GraphClient } from "../api/client.js";
import type { GraphMailMessage, GraphMailMessageFull, GraphAttachment, GraphSendMailRecipient, ODataResponse } from "../api/types.js";

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
    "graph_send_mail",
    "Send an email on behalf of a user via Microsoft Graph API. Requires Mail.Send application permission. Supports HTML or plain text body, multiple recipients, CC, and importance level.",
    {
      sender: z.string().describe("Sender user ID (GUID) or UPN (e.g. amit@majans.com). The email is sent from this mailbox."),
      to: z.array(z.object({
        address: z.string().describe("Recipient email address"),
        name: z.string().optional().describe("Recipient display name"),
      })).describe("To recipients (at least one required)"),
      subject: z.string().describe("Email subject line"),
      body: z.string().describe("Email body content (HTML or plain text)"),
      body_type: z.enum(["HTML", "Text"]).optional().describe("Body content type (default HTML)"),
      cc: z.array(z.object({
        address: z.string().describe("CC email address"),
        name: z.string().optional().describe("CC display name"),
      })).optional().describe("CC recipients"),
      importance: z.enum(["low", "normal", "high"]).optional().describe("Email importance (default normal)"),
      save_to_sent: z.boolean().optional().describe("Save to Sent Items (default true)"),
    },
    async ({ sender, to, subject, body, body_type, cc, importance, save_to_sent }) => {
      try {
        const toRecipients: GraphSendMailRecipient[] = to.map(r => ({
          emailAddress: { address: r.address, ...(r.name ? { name: r.name } : {}) },
        }));

        const ccRecipients: GraphSendMailRecipient[] | undefined = cc?.map(r => ({
          emailAddress: { address: r.address, ...(r.name ? { name: r.name } : {}) },
        }));

        const payload: Record<string, unknown> = {
          message: {
            subject,
            body: { contentType: body_type ?? "HTML", content: body },
            toRecipients,
            ...(ccRecipients ? { ccRecipients } : {}),
            ...(importance ? { importance } : {}),
          },
          saveToSentItems: save_to_sent ?? true,
        };

        await client.post<void>(
          `users/${encodeURIComponent(sender)}/sendMail`,
          payload
        );

        const recipientList = to.map(r => r.address).join(", ");
        return {
          content: [
            {
              type: "text" as const,
              text: JSON.stringify(
                {
                  sent: true,
                  from: sender,
                  to: recipientList,
                  subject,
                },
                null,
                2
              ),
            },
          ],
        };
      } catch (err) {
        return {
          content: [{ type: "text" as const, text: `Error sending email: ${(err as Error).message}` }],
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
