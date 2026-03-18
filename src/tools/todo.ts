import { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { z } from "zod";
import type { GraphClient } from "../api/client.js";
import type { GraphTodoList, GraphTodoTask, ODataResponse } from "../api/types.js";

export function registerTodoTools(server: McpServer, client: GraphClient): void {
  server.tool(
    "graph_list_todo_lists",
    "List a user's Microsoft To Do lists. Returns list names and IDs.",
    {
      user_id: z.string().describe("User ID (GUID) or UPN (e.g. amit@majans.com)"),
    },
    async ({ user_id }) => {
      try {
        const result = await client.get<ODataResponse<GraphTodoList>>(
          `users/${encodeURIComponent(user_id)}/todo/lists`
        );
        return {
          content: [
            {
              type: "text" as const,
              text: JSON.stringify(
                { count: result.value.length, lists: result.value },
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
    "graph_list_todo_tasks",
    "List tasks in a Microsoft To Do list. Returns task titles, status, due dates, and importance.",
    {
      user_id: z.string().describe("User ID (GUID) or UPN"),
      list_id: z.string().describe("To Do list ID (from graph_list_todo_lists)"),
      filter: z.string().optional().describe("OData $filter (e.g. \"status ne 'completed'\")"),
      top: z.number().optional().describe("Max results (default 50)"),
    },
    async ({ user_id, list_id, filter, top }) => {
      try {
        const params: Record<string, string> = {
          $top: String(top ?? 50),
          $orderby: "createdDateTime desc",
        };
        if (filter) params.$filter = filter;

        const result = await client.get<ODataResponse<GraphTodoTask>>(
          `users/${encodeURIComponent(user_id)}/todo/lists/${encodeURIComponent(list_id)}/tasks`,
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
                  tasks: result.value,
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
    "graph_create_todo_task",
    "Create a new task in a Microsoft To Do list.",
    {
      user_id: z.string().describe("User ID (GUID) or UPN"),
      list_id: z.string().describe("To Do list ID"),
      title: z.string().describe("Task title"),
      body_content: z.string().optional().describe("Task body/notes (plain text)"),
      due_date: z.string().optional().describe("Due date in ISO 8601 (e.g. 2026-04-01T00:00:00Z)"),
      importance: z.enum(["low", "normal", "high"]).optional().describe("Task importance (default normal)"),
      is_reminder_on: z.boolean().optional().describe("Enable reminder (default false)"),
    },
    async ({ user_id, list_id, title, body_content, due_date, importance, is_reminder_on }) => {
      try {
        const body: Record<string, unknown> = { title };
        if (body_content) {
          body.body = { content: body_content, contentType: "text" };
        }
        if (due_date) {
          body.dueDateTime = { dateTime: due_date, timeZone: "UTC" };
        }
        if (importance) body.importance = importance;
        if (is_reminder_on !== undefined) body.isReminderOn = is_reminder_on;

        const result = await client.post<GraphTodoTask>(
          `users/${encodeURIComponent(user_id)}/todo/lists/${encodeURIComponent(list_id)}/tasks`,
          body
        );
        return {
          content: [
            {
              type: "text" as const,
              text: JSON.stringify(
                {
                  created: true,
                  id: result.id,
                  title: result.title,
                  status: result.status,
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
    "graph_update_todo_task",
    "Update a Microsoft To Do task (title, status, due date, body, importance).",
    {
      user_id: z.string().describe("User ID (GUID) or UPN"),
      list_id: z.string().describe("To Do list ID"),
      task_id: z.string().describe("Task ID"),
      title: z.string().optional().describe("New task title"),
      status: z
        .enum(["notStarted", "inProgress", "completed", "waitingOnOthers", "deferred"])
        .optional()
        .describe("New task status"),
      body_content: z.string().optional().describe("New task body/notes (plain text)"),
      due_date: z.string().optional().describe("New due date in ISO 8601 (or empty string to clear)"),
      importance: z.enum(["low", "normal", "high"]).optional().describe("Task importance"),
    },
    async ({ user_id, list_id, task_id, title, status, body_content, due_date, importance }) => {
      try {
        const body: Record<string, unknown> = {};
        if (title !== undefined) body.title = title;
        if (status !== undefined) body.status = status;
        if (body_content !== undefined) {
          body.body = { content: body_content, contentType: "text" };
        }
        if (due_date !== undefined) {
          body.dueDateTime = due_date ? { dateTime: due_date, timeZone: "UTC" } : null;
        }
        if (importance !== undefined) body.importance = importance;

        await client.patch(
          `users/${encodeURIComponent(user_id)}/todo/lists/${encodeURIComponent(list_id)}/tasks/${encodeURIComponent(task_id)}`,
          body
        );

        const updated = Object.keys(body);
        return {
          content: [
            {
              type: "text" as const,
              text: `Task ${task_id} updated successfully. Properties changed: ${updated.join(", ")}`,
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
    "graph_complete_todo_task",
    "Mark a Microsoft To Do task as completed.",
    {
      user_id: z.string().describe("User ID (GUID) or UPN"),
      list_id: z.string().describe("To Do list ID"),
      task_id: z.string().describe("Task ID"),
    },
    async ({ user_id, list_id, task_id }) => {
      try {
        await client.patch(
          `users/${encodeURIComponent(user_id)}/todo/lists/${encodeURIComponent(list_id)}/tasks/${encodeURIComponent(task_id)}`,
          { status: "completed" }
        );
        return {
          content: [
            { type: "text" as const, text: `Task ${task_id} marked as completed.` },
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
