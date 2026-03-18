import { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { z } from "zod";
import type { GraphClient } from "../api/client.js";
import type {
  GraphPlannerPlan,
  GraphPlannerBucket,
  GraphPlannerTask,
  ODataResponse,
} from "../api/types.js";

export function registerPlannerTools(server: McpServer, client: GraphClient): void {
  server.tool(
    "graph_list_plans",
    "List Planner plans for an M365 group. Planner plans belong to M365 groups.",
    {
      group_id: z.string().describe("M365 Group ID (GUID) that owns the plans"),
    },
    async ({ group_id }) => {
      try {
        const result = await client.get<ODataResponse<GraphPlannerPlan>>(
          `groups/${encodeURIComponent(group_id)}/planner/plans`
        );
        return {
          content: [
            {
              type: "text" as const,
              text: JSON.stringify(
                { count: result.value.length, plans: result.value },
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
    "graph_get_plan",
    "Get details of a specific Planner plan by ID.",
    {
      plan_id: z.string().describe("Planner plan ID"),
    },
    async ({ plan_id }) => {
      try {
        const result = await client.get<GraphPlannerPlan>(
          `planner/plans/${encodeURIComponent(plan_id)}`
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
    "graph_list_buckets",
    "List buckets (columns) in a Planner plan.",
    {
      plan_id: z.string().describe("Planner plan ID"),
    },
    async ({ plan_id }) => {
      try {
        const result = await client.get<ODataResponse<GraphPlannerBucket>>(
          `planner/plans/${encodeURIComponent(plan_id)}/buckets`
        );
        return {
          content: [
            {
              type: "text" as const,
              text: JSON.stringify(
                { count: result.value.length, buckets: result.value },
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
    "graph_list_plan_tasks",
    "List all tasks in a Planner plan. Returns task title, status, assignees, due dates, and etags for updates.",
    {
      plan_id: z.string().describe("Planner plan ID"),
    },
    async ({ plan_id }) => {
      try {
        const tasks = await client.getAll<GraphPlannerTask>(
          `planner/plans/${encodeURIComponent(plan_id)}/tasks`
        );
        return {
          content: [
            {
              type: "text" as const,
              text: JSON.stringify(
                {
                  count: tasks.length,
                  tasks: tasks.map((t) => ({
                    id: t.id,
                    title: t.title,
                    bucketId: t.bucketId,
                    percentComplete: t.percentComplete,
                    priority: t.priority,
                    startDateTime: t.startDateTime,
                    dueDateTime: t.dueDateTime,
                    completedDateTime: t.completedDateTime,
                    assignees: Object.keys(t.assignments ?? {}),
                    etag: t["@odata.etag"],
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
    "graph_get_task",
    "Get full details of a specific Planner task by ID. Includes the etag needed for updates.",
    {
      task_id: z.string().describe("Planner task ID"),
    },
    async ({ task_id }) => {
      try {
        const result = await client.get<GraphPlannerTask>(
          `planner/tasks/${encodeURIComponent(task_id)}`
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
    "graph_create_task",
    "Create a new Planner task in a plan. Optionally assign to a bucket and/or users.",
    {
      plan_id: z.string().describe("Planner plan ID"),
      title: z.string().describe("Task title"),
      bucket_id: z.string().optional().describe("Bucket ID to place the task in"),
      assignments: z
        .array(z.string())
        .optional()
        .describe("Array of user IDs (GUIDs) to assign the task to"),
      due_date: z.string().optional().describe("Due date in ISO 8601 (e.g. 2026-04-01T00:00:00Z)"),
      start_date: z.string().optional().describe("Start date in ISO 8601"),
      priority: z
        .number()
        .optional()
        .describe("Priority: 0=none, 1=urgent, 3=important, 5=medium, 9=low"),
    },
    async ({ plan_id, title, bucket_id, assignments, due_date, start_date, priority }) => {
      try {
        const body: Record<string, unknown> = {
          planId: plan_id,
          title,
        };
        if (bucket_id) body.bucketId = bucket_id;
        if (due_date) body.dueDateTime = due_date;
        if (start_date) body.startDateTime = start_date;
        if (priority !== undefined) body.priority = priority;
        if (assignments && assignments.length > 0) {
          const assignObj: Record<string, { "@odata.type": string; orderHint: string }> = {};
          for (const userId of assignments) {
            assignObj[userId] = {
              "@odata.type": "#microsoft.graph.plannerAssignment",
              orderHint: " !",
            };
          }
          body.assignments = assignObj;
        }

        const result = await client.post<GraphPlannerTask>("planner/tasks", body);
        return {
          content: [
            {
              type: "text" as const,
              text: JSON.stringify(
                {
                  created: true,
                  id: result.id,
                  title: result.title,
                  planId: result.planId,
                  bucketId: result.bucketId,
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
    "graph_update_task",
    "Update a Planner task. Requires the task's etag (from graph_get_task or graph_list_plan_tasks) for concurrency control.",
    {
      task_id: z.string().describe("Planner task ID"),
      etag: z.string().describe("Task @odata.etag value (required for If-Match header)"),
      title: z.string().optional().describe("New task title"),
      percent_complete: z
        .number()
        .optional()
        .describe("Completion percentage (0, 25, 50, 75, 100)"),
      due_date: z.string().optional().describe("New due date in ISO 8601 (or empty string to clear)"),
      start_date: z.string().optional().describe("New start date in ISO 8601 (or empty string to clear)"),
      bucket_id: z.string().optional().describe("Move task to a different bucket"),
      priority: z
        .number()
        .optional()
        .describe("Priority: 0=none, 1=urgent, 3=important, 5=medium, 9=low"),
      assignments: z
        .record(z.unknown())
        .optional()
        .describe("Assignments object — set userId key to plannerAssignment or null to unassign"),
    },
    async ({ task_id, etag, title, percent_complete, due_date, start_date, bucket_id, priority, assignments }) => {
      try {
        const body: Record<string, unknown> = {};
        if (title !== undefined) body.title = title;
        if (percent_complete !== undefined) body.percentComplete = percent_complete;
        if (due_date !== undefined) body.dueDateTime = due_date || null;
        if (start_date !== undefined) body.startDateTime = start_date || null;
        if (bucket_id !== undefined) body.bucketId = bucket_id;
        if (priority !== undefined) body.priority = priority;
        if (assignments !== undefined) body.assignments = assignments;

        await client.patch(
          `planner/tasks/${encodeURIComponent(task_id)}`,
          body,
          { "If-Match": etag }
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
    "graph_delete_task",
    "Delete a Planner task. Requires the task's etag for concurrency control.",
    {
      task_id: z.string().describe("Planner task ID"),
      etag: z.string().describe("Task @odata.etag value (required for If-Match header)"),
    },
    async ({ task_id, etag }) => {
      try {
        await client.delete(
          `planner/tasks/${encodeURIComponent(task_id)}`,
          { "If-Match": etag }
        );
        return {
          content: [
            { type: "text" as const, text: `Task ${task_id} deleted successfully.` },
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
