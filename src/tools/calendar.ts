import { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { z } from "zod";
import type { GraphClient } from "../api/client.js";
import type {
  GraphCalendarEvent,
  GraphOnlineMeeting,
  GraphTranscript,
  ODataResponse,
} from "../api/types.js";

const EVENT_SELECT =
  "id,subject,start,end,organizer,attendees,isOnlineMeeting,onlineMeetingUrl,onlineMeeting,bodyPreview";

export function registerCalendarTools(server: McpServer, client: GraphClient): void {
  server.tool(
    "graph_list_events",
    "List calendar events for a user within a date range. Returns meetings with attendees and online meeting info.",
    {
      user_id: z.string().describe("User ID (GUID) or UPN (e.g. amit@majans.com)"),
      start: z.string().describe("Start date/time in ISO 8601 (e.g. 2026-03-01T00:00:00Z)"),
      end: z.string().describe("End date/time in ISO 8601 (e.g. 2026-03-08T23:59:59Z)"),
      filter: z
        .string()
        .optional()
        .describe("Additional OData $filter (e.g. \"isOnlineMeeting eq true\")"),
      top: z.number().optional().describe("Max results (default 25, max 100)"),
    },
    async ({ user_id, start, end, filter, top }) => {
      try {
        const params: Record<string, string> = {
          startDateTime: start,
          endDateTime: end,
          $select: EVENT_SELECT,
          $top: String(top ?? 25),
          $orderby: "start/dateTime asc",
        };
        if (filter) params.$filter = filter;

        const result = await client.get<ODataResponse<GraphCalendarEvent>>(
          `users/${encodeURIComponent(user_id)}/calendarview`,
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
                  events: result.value,
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
    "graph_get_online_meeting",
    "Get online meeting details by join URL. Use this to get the meeting ID needed for transcript retrieval.",
    {
      user_id: z.string().describe("Organizer's User ID (GUID) or UPN"),
      join_url: z
        .string()
        .describe("Teams meeting join URL (from calendar event's onlineMeeting.joinUrl)"),
    },
    async ({ user_id, join_url }) => {
      try {
        const filter = `joinWebUrl eq '${join_url.replace(/'/g, "''")}'`;
        const result = await client.get<ODataResponse<GraphOnlineMeeting>>(
          `users/${encodeURIComponent(user_id)}/onlineMeetings`,
          {
            $filter: filter,
            $select: "id,subject,startDateTime,endDateTime,joinWebUrl,chatInfo,participants",
          }
        );
        if (result.value.length === 0) {
          return {
            content: [
              {
                type: "text" as const,
                text: "No online meeting found for this join URL. The user must be the organizer.",
              },
            ],
          };
        }
        return {
          content: [{ type: "text" as const, text: JSON.stringify(result.value[0], null, 2) }],
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
    "graph_list_meeting_transcripts",
    "List available transcripts for an online meeting. Requires the meeting ID from graph_get_online_meeting.",
    {
      user_id: z.string().describe("Organizer's User ID (GUID) or UPN"),
      meeting_id: z.string().describe("Online meeting ID (from graph_get_online_meeting)"),
    },
    async ({ user_id, meeting_id }) => {
      try {
        const result = await client.get<ODataResponse<GraphTranscript>>(
          `users/${encodeURIComponent(user_id)}/onlineMeetings/${encodeURIComponent(meeting_id)}/transcripts`,
          { $select: "id,meetingId,meetingOrganizerId,createdDateTime,transcriptContentUrl" }
        );
        return {
          content: [
            {
              type: "text" as const,
              text: JSON.stringify(
                { count: result.value.length, transcripts: result.value },
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
    "graph_get_meeting_transcript_content",
    "Get the text content of a meeting transcript (VTT format). Returns the full transcript text with speaker names and timestamps.",
    {
      user_id: z.string().describe("Organizer's User ID (GUID) or UPN"),
      meeting_id: z.string().describe("Online meeting ID"),
      transcript_id: z.string().describe("Transcript ID (from graph_list_meeting_transcripts)"),
    },
    async ({ user_id, meeting_id, transcript_id }) => {
      try {
        const text = await client.getText(
          `users/${encodeURIComponent(user_id)}/onlineMeetings/${encodeURIComponent(meeting_id)}/transcripts/${encodeURIComponent(transcript_id)}/content`,
          { $format: "text/vtt" }
        );
        return {
          content: [{ type: "text" as const, text }],
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
