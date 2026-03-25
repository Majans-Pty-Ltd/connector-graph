import { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { z } from "zod";
import type { GraphClient } from "../api/client.js";
import type { GraphMailMessage, GraphMailMessageFull, GraphMailFolder, GraphAttachment, GraphFileAttachment, GraphSendMailRecipient, ODataResponse } from "../api/types.js";
import { extractContent } from "../utils/content-extractor.js";

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
    "graph_create_draft",
    "Create an Outlook draft email with optional file attachments. Draft is saved to the user's Drafts folder but NOT sent. Use this when the user wants to review before sending, or when attachments are needed.",
    {
      sender: z.string().describe("Sender user ID (GUID) or UPN (e.g. amit@majans.com). Draft is created in this mailbox."),
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
      attachments: z.array(z.object({
        filename: z.string().describe("Attachment filename (e.g. 'report.pdf')"),
        content_base64: z.string().describe("Base64-encoded file content"),
        content_type: z.string().describe("MIME type (e.g. 'application/vnd.openxmlformats-officedocument.presentationml.presentation')"),
      })).optional().describe("File attachments (base64-encoded)"),
    },
    async ({ sender, to, subject, body, body_type, cc, importance, attachments }) => {
      try {
        const toRecipients: GraphSendMailRecipient[] = to.map(r => ({
          emailAddress: { address: r.address, ...(r.name ? { name: r.name } : {}) },
        }));

        const ccRecipients: GraphSendMailRecipient[] | undefined = cc?.map(r => ({
          emailAddress: { address: r.address, ...(r.name ? { name: r.name } : {}) },
        }));

        const payload: Record<string, unknown> = {
          subject,
          body: { contentType: body_type ?? "HTML", content: body },
          toRecipients,
          ...(ccRecipients ? { ccRecipients } : {}),
          ...(importance ? { importance } : {}),
          ...(attachments && attachments.length > 0 ? {
            attachments: attachments.map(a => ({
              "@odata.type": "#microsoft.graph.fileAttachment",
              name: a.filename,
              contentBytes: a.content_base64,
              contentType: a.content_type,
            })),
          } : {}),
        };

        const result = await client.post<{ id: string; subject: string; webLink?: string }>(
          `users/${encodeURIComponent(sender)}/messages`,
          payload
        );

        const recipientList = to.map(r => r.address).join(", ");
        return {
          content: [
            {
              type: "text" as const,
              text: JSON.stringify(
                {
                  created: true,
                  draft_id: result.id,
                  from: sender,
                  to: recipientList,
                  subject,
                  attachment_count: attachments?.length ?? 0,
                  ...(result.webLink ? { web_link: result.webLink } : {}),
                },
                null,
                2
              ),
            },
          ],
        };
      } catch (err) {
        return {
          content: [{ type: "text" as const, text: `Error creating draft: ${(err as Error).message}` }],
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

  server.tool(
    "graph_get_attachment",
    "Download an email attachment's content by message ID and attachment ID. Returns decoded text for text-based types, or base64 for binary files.",
    {
      user_id: z.string().describe("User ID (GUID) or UPN"),
      message_id: z.string().describe("Message ID that contains the attachment"),
      attachment_id: z.string().describe("Attachment ID (from graph_list_attachments)"),
    },
    async ({ user_id, message_id, attachment_id }) => {
      try {
        const result = await client.get<GraphFileAttachment>(
          `users/${encodeURIComponent(user_id)}/messages/${encodeURIComponent(message_id)}/attachments/${encodeURIComponent(attachment_id)}`
        );

        const contentType = result.contentType ?? "";
        const isTextType =
          contentType.startsWith("text/") ||
          contentType.includes("json") ||
          contentType.includes("xml") ||
          contentType.includes("csv");

        let content: string;
        if (isTextType && result.contentBytes) {
          // Decode base64 to UTF-8 text for text-based attachments
          content = Buffer.from(result.contentBytes, "base64").toString("utf-8");
        } else {
          // Return raw base64 for binary files (images, PDFs, etc.)
          content = result.contentBytes ?? "";
        }

        return {
          content: [
            {
              type: "text" as const,
              text: JSON.stringify(
                {
                  id: result.id,
                  name: result.name,
                  contentType: result.contentType,
                  size: result.size,
                  encoding: isTextType ? "utf-8" : "base64",
                  content,
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
    "graph_read_attachment",
    "Download an email attachment and extract readable text content. Supports PDF, Word (.docx), Excel (.xlsx), HTML, CSV, EML, images, and plain text files. For unsupported types, use graph_get_attachment instead.",
    {
      user_id: z.string().describe("User ID (GUID) or UPN (e.g. amit@majans.com)"),
      message_id: z.string().describe("Message ID (from graph_search_mail or graph_read_mail)"),
      attachment_id: z.string().describe("Attachment ID (from graph_list_attachments)"),
    },
    async ({ user_id, message_id, attachment_id }) => {
      try {
        const result = await client.get<GraphFileAttachment>(
          `users/${encodeURIComponent(user_id)}/messages/${encodeURIComponent(message_id)}/attachments/${encodeURIComponent(attachment_id)}`
        );

        if (!result.contentBytes) {
          return {
            content: [
              {
                type: "text" as const,
                text: JSON.stringify(
                  {
                    name: result.name,
                    contentType: result.contentType,
                    size: result.size,
                    extractedText: "No content available (attachment may be a reference or item attachment).",
                    format: "unsupported",
                  },
                  null,
                  2
                ),
              },
            ],
          };
        }

        const extraction = await extractContent(
          result.contentBytes,
          result.contentType ?? "",
          result.name ?? "unknown"
        );

        return {
          content: [
            {
              type: "text" as const,
              text: JSON.stringify(
                {
                  name: result.name,
                  contentType: result.contentType,
                  size: result.size,
                  extractedText: extraction.extractedText,
                  format: extraction.format,
                  ...(extraction.metadata ? { metadata: extraction.metadata } : {}),
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
    "graph_list_mail_folders",
    "List mail folders for a user (Inbox, Sent Items, Drafts, Archive, etc.). Returns folder IDs, display names, and unread/total item counts.",
    {
      user_id: z.string().describe("User ID (GUID) or UPN (e.g. amit@majans.com)"),
      top: z.number().optional().describe("Max folders to return (default 50)"),
    },
    async ({ user_id, top }) => {
      try {
        const result = await client.get<ODataResponse<GraphMailFolder>>(
          `users/${encodeURIComponent(user_id)}/mailFolders`,
          {
            $select: "id,displayName,parentFolderId,childFolderCount,unreadItemCount,totalItemCount",
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
                  folders: result.value,
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
    "graph_list_folder_messages",
    "List messages from a specific mail folder. Use well-known folder names (inbox, sentitems, drafts, archive, deleteditems, junkemail) or a folder ID from graph_list_mail_folders. Supports OData $filter and $search.",
    {
      user_id: z.string().describe("User ID (GUID) or UPN (e.g. amit@majans.com)"),
      folder: z.string().describe("Well-known folder name (inbox, sentitems, drafts, archive, deleteditems, junkemail) or folder ID (GUID)"),
      search: z.string().optional().describe("Free-text search query (e.g. 'invoice')"),
      filter: z.string().optional().describe("OData $filter (e.g. \"receivedDateTime ge 2024-01-01\")"),
      top: z.number().optional().describe("Max results to return (default 25, max 100)"),
      orderby: z.string().optional().describe("OData $orderby (default 'receivedDateTime desc')"),
    },
    async ({ user_id, folder, search, filter, top, orderby }) => {
      try {
        const params: Record<string, string> = {
          $select: MAIL_SELECT,
          $top: String(top ?? 25),
          $orderby: orderby ?? "receivedDateTime desc",
        };
        if (search) params.$search = `"${search}"`;
        if (filter) params.$filter = filter;

        const result = await client.get<ODataResponse<GraphMailMessage>>(
          `users/${encodeURIComponent(user_id)}/mailFolders/${encodeURIComponent(folder)}/messages`,
          params
        );
        return {
          content: [
            {
              type: "text" as const,
              text: JSON.stringify(
                {
                  folder,
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
    "graph_reply_mail",
    "Reply to an email in-thread. The reply is sent from the user's mailbox and preserves the conversation thread. Use message_id from graph_search_mail or graph_list_folder_messages.",
    {
      user_id: z.string().describe("User ID (GUID) or UPN (e.g. amit@majans.com). Reply is sent from this mailbox."),
      message_id: z.string().describe("Message ID of the email to reply to (from graph_search_mail)"),
      comment: z.string().describe("Reply body content (HTML supported)"),
    },
    async ({ user_id, message_id, comment }) => {
      try {
        await client.post<void>(
          `users/${encodeURIComponent(user_id)}/messages/${encodeURIComponent(message_id)}/reply`,
          { comment }
        );
        return {
          content: [
            {
              type: "text" as const,
              text: JSON.stringify(
                {
                  replied: true,
                  from: user_id,
                  in_reply_to: message_id,
                },
                null,
                2
              ),
            },
          ],
        };
      } catch (err) {
        return {
          content: [{ type: "text" as const, text: `Error replying: ${(err as Error).message}` }],
          isError: true,
        };
      }
    }
  );
}
