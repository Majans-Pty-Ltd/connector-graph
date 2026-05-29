import { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { z } from "zod";
import type { GraphClient } from "../api/client.js";
import type { GraphMailMessage, GraphMailMessageFull, GraphMailFolder, GraphAttachment, GraphFileAttachment, GraphSendMailRecipient, ODataResponse } from "../api/types.js";
import { extractContent } from "../utils/content-extractor.js";

const MAIL_SELECT = "id,subject,bodyPreview,from,toRecipients,receivedDateTime,isRead,hasAttachments,importance";

// Zod shape for an @mention entry — used by send_mail, create_draft, reply_mail.
const mentionShape = z.object({
  address: z.string().describe("Email address of the person being mentioned"),
  name: z.string().optional().describe("Display name (recommended — appears next to the @ in Outlook)"),
});

// Convert input mentions to Microsoft Graph payload format.
// Body must contain matching @Name text (HTML) for Outlook to render the highlight.
type MentionInput = { address: string; name?: string };
function toGraphMentions(input: MentionInput[] | undefined): Array<Record<string, unknown>> | undefined {
  if (!input || input.length === 0) return undefined;
  return input.map(m => ({
    "@odata.type": "#microsoft.graph.mention",
    mentioned: {
      address: m.address,
      ...(m.name ? { name: m.name } : {}),
    },
  }));
}

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
        };
        // Graph API does not support $orderby with $search
        if (search) {
          params.$search = `"${search}"`;
        } else {
          params.$orderby = orderby ?? "receivedDateTime desc";
        }
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
    "Send an email on behalf of a user via Microsoft Graph API. Requires Mail.Send application permission. Supports HTML or plain text body, multiple recipients, CC, importance level, and file attachments. NOTE: each attachment is limited to ~4MB via /sendMail (Graph total request limit). For larger files use graph_create_draft + the createUploadSession API.",
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
      attachments: z.array(z.object({
        name: z.string().describe("Attachment filename (e.g. 'report.pdf')"),
        contentBytes: z.string().describe("Base64-encoded file content"),
        contentType: z.string().optional().describe("MIME type (e.g. 'application/pdf'). Optional — Graph infers from filename if omitted."),
      })).optional().describe("File attachments. Each attachment ≤4MB; total request must stay under Graph's ~4MB /sendMail limit. For larger files, use graph_create_draft and the upload session API."),
      internet_message_headers: z.array(z.object({
        name: z.string().describe("Header name. Microsoft Graph requires custom headers to start with 'X-' (e.g. 'X-Majans-Workflow')."),
        value: z.string().describe("Header value (string). Long values are accepted but may be truncated by some mail clients."),
      })).optional().describe("Custom internet message headers (RFC 5322 / SMTP). Used for email provenance tracking — e.g. X-Majans-Workflow, X-Majans-Run-Id, X-Majans-Repo. Only X-prefixed names are accepted by Graph for custom headers."),
      mentions: z.array(mentionShape).optional().describe("@mentions in the email body. Each mention attaches to the message metadata so Outlook renders the @Name highlight, fires a notification badge for the mentioned user, and the recipient can filter their inbox by '@me'. The body MUST contain matching @Name text (e.g. '@Phoebe Zhai') for the highlight to render — the connector does not insert it for you."),
    },
    async ({ sender, to, subject, body, body_type, cc, importance, save_to_sent, attachments, internet_message_headers, mentions }) => {
      try {
        const toRecipients: GraphSendMailRecipient[] = to.map(r => ({
          emailAddress: { address: r.address, ...(r.name ? { name: r.name } : {}) },
        }));

        const ccRecipients: GraphSendMailRecipient[] | undefined = cc?.map(r => ({
          emailAddress: { address: r.address, ...(r.name ? { name: r.name } : {}) },
        }));

        const graphAttachments = attachments?.map(a => ({
          "@odata.type": "#microsoft.graph.fileAttachment",
          name: a.name,
          contentBytes: a.contentBytes,
          ...(a.contentType ? { contentType: a.contentType } : {}),
        }));

        const internetMessageHeaders = internet_message_headers?.map(h => ({
          name: h.name,
          value: h.value,
        }));

        const graphMentions = toGraphMentions(mentions);

        const payload: Record<string, unknown> = {
          message: {
            subject,
            body: { contentType: body_type ?? "HTML", content: body },
            toRecipients,
            ...(ccRecipients ? { ccRecipients } : {}),
            ...(importance ? { importance } : {}),
            ...(graphAttachments && graphAttachments.length > 0 ? { attachments: graphAttachments } : {}),
            ...(internetMessageHeaders && internetMessageHeaders.length > 0 ? { internetMessageHeaders } : {}),
            ...(graphMentions ? { mentions: graphMentions } : {}),
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
                  attachment_count: attachments?.length ?? 0,
                  mention_count: mentions?.length ?? 0,
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
        name: z.string().describe("Attachment filename (e.g. 'report.pdf')"),
        contentBytes: z.string().describe("Base64-encoded file content"),
        contentType: z.string().optional().describe("MIME type (e.g. 'application/pdf'). Optional — Graph infers from filename if omitted."),
      })).optional().describe("File attachments (base64-encoded). Each attachment ≤3MB when added via /messages/{id}/attachments. For larger files, use the createUploadSession API."),
      reply_to_message_id: z.string().optional().describe("Message ID to reply to. Creates an in-thread draft reply-all instead of a new draft. Recipients default to reply-all but can be overridden with to/cc."),
      mentions: z.array(mentionShape).optional().describe("@mentions in the email body. Each mention attaches to the message metadata so Outlook renders the @Name highlight, fires a notification badge for the mentioned user, and the recipient can filter their inbox by '@me'. The body MUST contain matching @Name text (e.g. '@Phoebe Zhai') for the highlight to render — the connector does not insert it for you."),
    },
    async ({ sender, to, subject, body, body_type, cc, importance, attachments, reply_to_message_id, mentions }) => {
      try {
        const toRecipients: GraphSendMailRecipient[] | undefined = to?.map(r => ({
          emailAddress: { address: r.address, ...(r.name ? { name: r.name } : {}) },
        }));

        const ccRecipients: GraphSendMailRecipient[] | undefined = cc?.map(r => ({
          emailAddress: { address: r.address, ...(r.name ? { name: r.name } : {}) },
        }));

        const graphMentions = toGraphMentions(mentions);

        // Map to Graph fileAttachment shape once. Reused for inline attach on
        // new drafts and the per-item POST loop on the reply path.
        const graphAttachments = attachments?.map(a => ({
          "@odata.type": "#microsoft.graph.fileAttachment",
          name: a.name,
          contentBytes: a.contentBytes,
          ...(a.contentType ? { contentType: a.contentType } : {}),
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
            ...(graphMentions ? { mentions: graphMentions } : {}),
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
            ...(graphMentions ? { mentions: graphMentions } : {}),
            // Attach inline on creation so body + attachments commit in ONE
            // atomic POST. The previous create-then-POST-each-attachment flow
            // left a body-only draft whenever the follow-up attachment POST
            // stalled or failed — Graph accepts fileAttachments inline here.
            ...(graphAttachments && graphAttachments.length > 0 ? { attachments: graphAttachments } : {}),
          };

          result = await client.post<{ id: string; subject: string; webLink?: string }>(
            `users/${encodeURIComponent(sender)}/messages`,
            payload
          );
        }

        // Reply drafts are materialised via createReplyAll, which does NOT accept
        // inline attachments — so add them with a follow-up POST per attachment.
        // New drafts already attached inline in the create POST above.
        if (reply_to_message_id && graphAttachments && graphAttachments.length > 0) {
          for (const a of graphAttachments) {
            await client.post<unknown>(
              `users/${encodeURIComponent(sender)}/messages/${encodeURIComponent(result.id)}/attachments`,
              a
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
                  mention_count: mentions?.length ?? 0,
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
    "graph_list_child_folders",
    "List child (sub) folders of a specific mail folder. Use a folder ID from graph_list_mail_folders to get its subfolders.",
    {
      user_id: z.string().describe("User ID (GUID) or UPN (e.g. amit@majans.com)"),
      folder_id: z.string().describe("Parent folder ID (from graph_list_mail_folders)"),
      top: z.number().optional().describe("Max folders to return (default 100)"),
    },
    async ({ user_id, folder_id, top }) => {
      try {
        const result = await client.get<ODataResponse<GraphMailFolder>>(
          `users/${encodeURIComponent(user_id)}/mailFolders/${encodeURIComponent(folder_id)}/childFolders`,
          {
            $select: "id,displayName,parentFolderId,childFolderCount,unreadItemCount,totalItemCount",
            $top: String(top ?? 100),
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
        };
        // Graph API does not support $orderby with $search
        if (search) {
          params.$search = `"${search}"`;
        } else {
          params.$orderby = orderby ?? "receivedDateTime desc";
        }
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
    "Reply to an email in-thread. The reply is sent from the user's mailbox and preserves the conversation thread. Use message_id from graph_search_mail or graph_list_folder_messages. Supports optional file attachments — when provided, uses createReply (draft) → add attachments → send. Each attachment ≤3MB; for larger files use graph_create_draft + the createUploadSession API.",
    {
      user_id: z.string().describe("User ID (GUID) or UPN (e.g. amit@majans.com). Reply is sent from this mailbox."),
      message_id: z.string().describe("Message ID of the email to reply to (from graph_search_mail)"),
      comment: z.string().describe("Reply body content (HTML supported)"),
      attachments: z.array(z.object({
        name: z.string().describe("Attachment filename (e.g. 'report.pdf')"),
        contentBytes: z.string().describe("Base64-encoded file content"),
        contentType: z.string().optional().describe("MIME type (e.g. 'application/pdf'). Optional — Graph infers from filename if omitted."),
      })).optional().describe("File attachments. Each ≤3MB. When provided, the reply is built as a draft (createReply) so attachments can be added before sending."),
      mentions: z.array(mentionShape).optional().describe("@mentions in the reply body. Each mention attaches to the message metadata so Outlook renders the @Name highlight, fires a notification badge for the mentioned user, and the recipient can filter their inbox by '@me'. The comment body MUST contain matching @Name text (e.g. '@Phoebe Zhai') for the highlight to render — the connector does not insert it for you. When provided, the reply is built as a draft (createReply) so mentions can be set before sending."),
    },
    async ({ user_id, message_id, comment, attachments, mentions }) => {
      try {
        const graphMentions = toGraphMentions(mentions);
        const useDraftFlow = (attachments && attachments.length > 0) || !!graphMentions;

        if (useDraftFlow) {
          // Draft-based flow: createReply → modify draft → send.
          // /reply doesn't accept attachments or mentions inline, so we materialise the draft first.
          const createReplyPayload: Record<string, unknown> = { comment };
          if (graphMentions) {
            createReplyPayload.message = { mentions: graphMentions };
          }

          const draft = await client.post<{ id: string }>(
            `users/${encodeURIComponent(user_id)}/messages/${encodeURIComponent(message_id)}/createReply`,
            createReplyPayload
          );

          if (attachments && attachments.length > 0) {
            for (const a of attachments) {
              await client.post<unknown>(
                `users/${encodeURIComponent(user_id)}/messages/${encodeURIComponent(draft.id)}/attachments`,
                {
                  "@odata.type": "#microsoft.graph.fileAttachment",
                  name: a.name,
                  contentBytes: a.contentBytes,
                  ...(a.contentType ? { contentType: a.contentType } : {}),
                }
              );
            }
          }

          await client.post<void>(
            `users/${encodeURIComponent(user_id)}/messages/${encodeURIComponent(draft.id)}/send`,
            {}
          );
        } else {
          await client.post<void>(
            `users/${encodeURIComponent(user_id)}/messages/${encodeURIComponent(message_id)}/reply`,
            { comment }
          );
        }

        return {
          content: [
            {
              type: "text" as const,
              text: JSON.stringify(
                {
                  replied: true,
                  from: user_id,
                  in_reply_to: message_id,
                  attachment_count: attachments?.length ?? 0,
                  mention_count: mentions?.length ?? 0,
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
