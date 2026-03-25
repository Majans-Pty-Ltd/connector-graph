import { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { z } from "zod";
import type { GraphClient } from "../api/client.js";
import type {
  GraphCalendarEvent,
  GraphOnlineMeeting,
  GraphTranscript,
  GraphScheduleInfo,
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

  // ── Create calendar event ──

  server.tool(
    "graph_create_event",
    "Create a new calendar event for a user. Supports attendees, online Teams meeting, and recurrence.",
    {
      user_id: z.string().describe("User ID (GUID) or UPN (e.g. amit@majans.com)"),
      subject: z.string().describe("Event subject/title"),
      start: z.string().describe("Start date/time in ISO 8601 (e.g. 2026-04-01T09:00:00)"),
      end: z.string().describe("End date/time in ISO 8601 (e.g. 2026-04-01T10:00:00)"),
      time_zone: z.string().optional().describe("Time zone (default 'Australia/Brisbane')"),
      body: z.string().optional().describe("Event body/description (HTML supported)"),
      location: z.string().optional().describe("Location display name (e.g. 'Board Room')"),
      attendees: z.array(z.object({
        address: z.string().describe("Attendee email address"),
        name: z.string().optional().describe("Attendee display name"),
        type: z.enum(["required", "optional", "resource"]).optional().describe("Attendee type (default required)"),
      })).optional().describe("Event attendees"),
      is_online_meeting: z.boolean().optional().describe("Create a Teams online meeting (default false)"),
      is_all_day: z.boolean().optional().describe("All-day event (default false). Use date-only format for start/end (e.g. 2026-04-01)."),
    },
    async ({ user_id, subject, start, end, time_zone, body, location, attendees, is_online_meeting, is_all_day }) => {
      try {
        const tz = time_zone ?? "Australia/Brisbane";
        const payload: Record<string, unknown> = {
          subject,
          start: { dateTime: start, timeZone: tz },
          end: { dateTime: end, timeZone: tz },
          ...(body ? { body: { contentType: "HTML", content: body } } : {}),
          ...(location ? { location: { displayName: location } } : {}),
          ...(is_online_meeting ? { isOnlineMeeting: true, onlineMeetingProvider: "teamsForBusiness" } : {}),
          ...(is_all_day ? { isAllDay: true } : {}),
          ...(attendees ? {
            attendees: attendees.map(a => ({
              emailAddress: { address: a.address, ...(a.name ? { name: a.name } : {}) },
              type: a.type ?? "required",
            })),
          } : {}),
        };

        const result = await client.post<GraphCalendarEvent & { webLink?: string }>(
          `users/${encodeURIComponent(user_id)}/events`,
          payload
        );

        return {
          content: [{
            type: "text" as const,
            text: JSON.stringify({
              created: true,
              event_id: result.id,
              subject: result.subject,
              start: result.start,
              end: result.end,
              is_online_meeting: result.isOnlineMeeting,
              join_url: result.onlineMeeting?.joinUrl ?? null,
              ...(result.webLink ? { web_link: result.webLink } : {}),
            }, null, 2),
          }],
        };
      } catch (err) {
        return {
          content: [{ type: "text" as const, text: `Error creating event: ${(err as Error).message}` }],
          isError: true,
        };
      }
    }
  );

  // ── Update calendar event ──

  server.tool(
    "graph_update_event",
    "Update an existing calendar event. Only provided fields are changed.",
    {
      user_id: z.string().describe("User ID (GUID) or UPN"),
      event_id: z.string().describe("Event ID (from graph_list_events)"),
      subject: z.string().optional().describe("New subject"),
      start: z.string().optional().describe("New start date/time ISO 8601"),
      end: z.string().optional().describe("New end date/time ISO 8601"),
      time_zone: z.string().optional().describe("Time zone (default 'Australia/Brisbane')"),
      body: z.string().optional().describe("New body (HTML supported)"),
      location: z.string().optional().describe("New location"),
      is_cancelled: z.boolean().optional().describe("Set true to cancel the event"),
    },
    async ({ user_id, event_id, subject, start, end, time_zone, body, location, is_cancelled }) => {
      try {
        const tz = time_zone ?? "Australia/Brisbane";
        const payload: Record<string, unknown> = {
          ...(subject ? { subject } : {}),
          ...(start ? { start: { dateTime: start, timeZone: tz } } : {}),
          ...(end ? { end: { dateTime: end, timeZone: tz } } : {}),
          ...(body ? { body: { contentType: "HTML", content: body } } : {}),
          ...(location ? { location: { displayName: location } } : {}),
          ...(is_cancelled ? { isCancelled: true } : {}),
        };

        await client.patch<unknown>(
          `users/${encodeURIComponent(user_id)}/events/${encodeURIComponent(event_id)}`,
          payload
        );

        return {
          content: [{
            type: "text" as const,
            text: JSON.stringify({ updated: true, event_id }, null, 2),
          }],
        };
      } catch (err) {
        return {
          content: [{ type: "text" as const, text: `Error updating event: ${(err as Error).message}` }],
          isError: true,
        };
      }
    }
  );

  // ── Delete calendar event ──

  server.tool(
    "graph_delete_event",
    "Delete a calendar event. This removes the event and sends cancellation notices to attendees.",
    {
      user_id: z.string().describe("User ID (GUID) or UPN"),
      event_id: z.string().describe("Event ID to delete (from graph_list_events)"),
    },
    async ({ user_id, event_id }) => {
      try {
        await client.delete(
          `users/${encodeURIComponent(user_id)}/events/${encodeURIComponent(event_id)}`
        );
        return {
          content: [{
            type: "text" as const,
            text: JSON.stringify({ deleted: true, event_id }, null, 2),
          }],
        };
      } catch (err) {
        return {
          content: [{ type: "text" as const, text: `Error deleting event: ${(err as Error).message}` }],
          isError: true,
        };
      }
    }
  );

  // ── Free/busy schedule lookup ──

  server.tool(
    "graph_get_schedule",
    "Get free/busy availability for one or more users within a time range. Useful for finding meeting slots.",
    {
      user_id: z.string().describe("Requesting user ID (GUID) or UPN — must have calendar access"),
      schedules: z.array(z.string()).describe("Email addresses to check availability for (e.g. ['amit@majans.com', 'kam@majans.com'])"),
      start: z.string().describe("Start date/time in ISO 8601 (e.g. 2026-04-01T08:00:00)"),
      end: z.string().describe("End date/time in ISO 8601 (e.g. 2026-04-01T18:00:00)"),
      time_zone: z.string().optional().describe("Time zone (default 'Australia/Brisbane')"),
      interval: z.number().optional().describe("Availability interval in minutes (default 30)"),
    },
    async ({ user_id, schedules, start, end, time_zone, interval }) => {
      try {
        const tz = time_zone ?? "Australia/Brisbane";
        const result = await client.post<{ value: GraphScheduleInfo[] }>(
          `users/${encodeURIComponent(user_id)}/calendar/getSchedule`,
          {
            schedules,
            startTime: { dateTime: start, timeZone: tz },
            endTime: { dateTime: end, timeZone: tz },
            availabilityViewInterval: interval ?? 30,
          }
        );

        const formatted = result.value.map(s => ({
          user: s.scheduleId,
          availability_view: s.availabilityView,
          availability_legend: "0=free, 1=tentative, 2=busy, 3=OOF, 4=working-elsewhere",
          busy_slots: s.scheduleItems.map(item => ({
            status: item.status,
            subject: item.subject ?? "(private)",
            start: item.start.dateTime,
            end: item.end.dateTime,
          })),
        }));

        return {
          content: [{
            type: "text" as const,
            text: JSON.stringify({ count: formatted.length, schedules: formatted }, null, 2),
          }],
        };
      } catch (err) {
        return {
          content: [{ type: "text" as const, text: `Error getting schedule: ${(err as Error).message}` }],
          isError: true,
        };
      }
    }
  );
}
