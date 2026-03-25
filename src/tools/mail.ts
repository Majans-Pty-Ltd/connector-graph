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
    "Create an Outlook draft email with optional file attachments. Draft is saved to the user's Drafts folder but NOT sent. Use this when the user wants to review before sending, or when attachments are needed. To create an in-thread draft reply, provide reply_to_message_id — this uses createReplyAll to preserve the conversation thread and pre-populate recipients.",
    {
      sender: z.string().describe("Sender user ID (GUID) or UPN (e.g. amit@majans.com). Draft is created in this mailbox."),
      to: z.array(z.object({
        address: z.string().describe("Recipient email address"),
        name: z.string().optional().describe("Recipient display name"),
      })).optional().describe("To recipients. Required for new drafts, optional for replies (defaults to reply-all recipients)."),
      subject: z.string().optional().describe("Email subject line. Required for new drafts, optional for replies (defaults to RE: original subject)."),
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
      reply_to_message_id: z.string().optional().describe("Message ID to reply to. Creates an in-thread draft reply-all instead of a new draft. Recipients default to reply-all but can be overridden with to/cc."),
    },
    async ({ sender, to, subject, body, body_type, cc, importance, attachments, reply_to_message_id }) => {
      try {
        const toRecipients: GraphSendMailRecipient[] | undefined = to?.map(r => ({
          emailAddress: { address: r.address, ...(r.name ? { name: r.name } : {}) },
        }));

        const ccRecipients: GraphSendMailRecipient[] | undefined = cc?.map(r => ({
          emailAddress: { address: r.address, ...(r.name ? { name: r.name } : {}) },
        }));

        let result: { id: string; subject: string; webLink?: string };

        if (reply_to_message_id) {
          // Create in-thread draft reply-all
          const message: Record<string, unknown> = {
            body: { contentType: body_type ?? "HTML", content: body },
            ...(toRecipients ? { toRecipients } : {}),
            ...(ccRecipients ? { ccRecipients } : {}),
            ...(subject ? { subject } : {}),
            ...(importance ? { importance } : {}),
          };

          result = await client.post<{ id: string; subject: string; webLink?: string }>(
            `users/${encodeURIComponent(sender)}/messages/${encodeURIComponent(reply_to_message_id)}/createReplyAll`,
            { message }
          );
        } else {
          // Create new draft
          if (!to || to.length === 0) {
            return {
              content: [{ type: "text" as const, text: "Error: 'to' recipients are required for new drafts." }],
              isError: true,
            };
          }

          const payload: Record<string, unknown> = {
            subject: subject ?? "",
            body: { contentType: body_type ?? "HTML", content: body },
            toRecipients,
            ...(ccRecipients ? { ccRecipients } : {}),
            ...(importance ? { importance } : {}),
          };

          result = await client.post<{ id: string; subject: string; webLink?: string }>(
            `users/${encodeURIComponent(sender)}/messages`,
            payload
          );
        }

        // Add attachments to the draft if any
        if (attachments && attachments.length > 0) {
          for (const a of attachments) {
            await client.post<unknown>(
              `users/${encodeURIComponent(sender)}/messages/${encodeURIComponent(result.id)}/attachments`,
              {
                "@odata.type": "#microsoft.graph.fileAttachment",
                name: a.filename,
                contentBytes: a.content_base64,
                contentType: a.content_type,
              }
            );
          }
        }

        const recipientList = to?.map(r => r.address).join(", ") ?? "(reply-all recipients)";
        return {
          content: [
            {
              type: "text" as const,
              text: JSON.stringify(
                {
                  created: true,
                  is_reply: !!reply_to_message_id,
                  draft_id: result.id,
                  from: sender,
                  to: recipientList,
                  subject: result.subject ?? subject ?? "",
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

  // ── Move message to folder ──

  server.tool(
    "graph_move_mail",
    "Move an email to a different folder. Use graph_list_mail_folders to get folder IDs. Common folders: 'DeletedItems', 'Archive', 'Inbox', 'Drafts', 'JunkEmail'.",
    {
      user_id: z.string().describe("User ID (GUID) or UPN"),
      message_id: z.string().describe("Message ID to move"),
      destination_folder_id: z.string().describe("Destination folder ID (from graph_list_mail_folders) or well-known name: 'DeletedItems', 'Archive', 'Inbox', 'Drafts', 'JunkEmail', 'SentItems'"),
    },
    async ({ user_id, message_id, destination_folder_id }) => {
      try {
        const result = await client.post<{ id: string; parentFolderId: string }>(
          `users/${encodeURIComponent(user_id)}/messages/${encodeURIComponent(message_id)}/move`,
          { destinationId: destination_folder_id }
        );
        return {
          content: [{
            type: "text" as const,
            text: JSON.stringify({
              moved: true,
              new_message_id: result.id,
              destination_folder: destination_folder_id,
            }, null, 2),
          }],
        };
      } catch (err) {
        return {
          content: [{ type: "text" as const, text: `Error moving message: ${(err as Error).message}` }],
          isError: true,
        };
      }
    }
  );

  // ── Delete message ──

  server.tool(
    "graph_delete_mail",
    "Delete an email. Moves to Deleted Items by default. Use permanently=true to hard-delete (skip Deleted Items).",
    {
      user_id: z.string().describe("User ID (GUID) or UPN"),
      message_id: z.string().describe("Message ID to delete"),
      permanently: z.boolean().optional().describe("Hard-delete permanently (default false — moves to Deleted Items)"),
    },
    async ({ user_id, message_id, permanently }) => {
      try {
        if (permanently) {
          await client.delete(
            `users/${encodeURIComponent(user_id)}/messages/${encodeURIComponent(message_id)}`
          );
        } else {
          await client.post<unknown>(
            `users/${encodeURIComponent(user_id)}/messages/${encodeURIComponent(message_id)}/move`,
            { destinationId: "DeletedItems" }
          );
        }
        return {
          content: [{
            type: "text" as const,
            text: JSON.stringify({
              deleted: true,
              permanent: permanently ?? false,
              message_id,
            }, null, 2),
          }],
        };
      } catch (err) {
        return {
          content: [{ type: "text" as const, text: `Error deleting message: ${(err as Error).message}` }],
          isError: true,
        };
      }
    }
  );

  // ── Update message properties (read/unread, flag, importance, categories) ──

  server.tool(
    "graph_update_mail",
    "Update email properties: mark as read/unread, flag for follow-up, set importance, or assign categories.",
    {
      user_id: z.string().describe("User ID (GUID) or UPN"),
      message_id: z.string().describe("Message ID to update"),
      is_read: z.boolean().optional().describe("Mark as read (true) or unread (false)"),
      importance: z.enum(["low", "normal", "high"]).optional().describe("Set importance level"),
      flag_status: z.enum(["notFlagged", "flagged", "complete"]).optional().describe("Follow-up flag: 'flagged' to flag, 'complete' to mark done, 'notFlagged' to clear"),
      categories: z.array(z.string()).optional().describe("Set category labels (e.g. ['Red category', 'Important']). Pass empty array to clear."),
    },
    async ({ user_id, message_id, is_read, importance, flag_status, categories }) => {
      try {
        const payload: Record<string, unknown> = {
          ...(is_read !== undefined ? { isRead: is_read } : {}),
          ...(importance ? { importance } : {}),
          ...(flag_status ? { flag: { flagStatus: flag_status } } : {}),
          ...(categories !== undefined ? { categories } : {}),
        };

        await client.patch<unknown>(
          `users/${encodeURIComponent(user_id)}/messages/${encodeURIComponent(message_id)}`,
          payload
        );

        return {
          content: [{
            type: "text" as const,
            text: JSON.stringify({
              updated: true,
              message_id,
              changes: {
                ...(is_read !== undefined ? { is_read } : {}),
                ...(importance ? { importance } : {}),
                ...(flag_status ? { flag_status } : {}),
                ...(categories !== undefined ? { categories } : {}),
              },
            }, null, 2),
          }],
        };
      } catch (err) {
        return {
          content: [{ type: "text" as const, text: `Error updating message: ${(err as Error).message}` }],
          isError: true,
        };
      }
    }
  );
}
