import type { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js"
import { z } from "zod"
import { makeGraphRequest, getAccessToken, MS_GRAPH_BASE } from "../graph-client.js"
import { formatTask } from "../helpers.js"
import type { Task, TaskList } from "../types.js"

export function register(server: McpServer) {
  // Convenience tools
  server.tool(
    "complete-task",
    "Mark a task as completed. A shortcut for update-task with status 'completed'.",
    {
      listId: z.string().describe("ID of the task list"),
      taskId: z.string().describe("ID of the task to complete"),
    },
    async ({ listId, taskId }) => {
      try {
        const token = await getAccessToken()
        if (!token) {
          return {
            content: [{ type: "text", text: "Failed to authenticate with Microsoft API" }],
          }
        }

        const response = await makeGraphRequest<Task>(
          `${MS_GRAPH_BASE}/me/todo/lists/${listId}/tasks/${taskId}`,
          token,
          "PATCH",
          { status: "completed" },
        )

        if (!response) {
          return {
            content: [{ type: "text", text: `Failed to complete task: ${taskId}` }],
          }
        }

        return {
          content: [{ type: "text", text: `Task completed!\nTitle: ${response.title}\nID: ${response.id}` }],
        }
      } catch (error) {
        return {
          content: [{ type: "text", text: `Error completing task: ${error}` }],
        }
      }
    },
  )

  server.tool(
    "search-tasks",
    "Search for tasks across all lists by keyword, status, importance, or due date range. Returns matching tasks with their list names.",
    {
      keyword: z.string().optional().describe("Search keyword to match in task titles (case-insensitive)"),
      status: z
        .enum(["notStarted", "inProgress", "completed", "waitingOnOthers", "deferred"])
        .optional()
        .describe("Filter by task status"),
      importance: z.enum(["low", "normal", "high"]).optional().describe("Filter by importance"),
      dueBefore: z.string().optional().describe("Filter tasks due before this date (ISO format, e.g. 2024-12-31)"),
      dueAfter: z.string().optional().describe("Filter tasks due after this date (ISO format, e.g. 2024-01-01)"),
      includeCompleted: z.boolean().optional().describe("Include completed tasks in results (default: false)"),
    },
    async ({ keyword, status, importance, dueBefore, dueAfter, includeCompleted = false }) => {
      try {
        const token = await getAccessToken()
        if (!token) {
          return {
            content: [{ type: "text", text: "Failed to authenticate with Microsoft API" }],
          }
        }

        // Get all lists
        const listsResponse = await makeGraphRequest<{ value: TaskList[] }>(`${MS_GRAPH_BASE}/me/todo/lists`, token)
        if (!listsResponse) {
          return {
            content: [{ type: "text", text: "Failed to retrieve task lists" }],
          }
        }

        const allResults: { listName: string; task: Task }[] = []

        for (const list of listsResponse.value || []) {
          // Build OData filter
          const filters: string[] = []
          if (status) filters.push(`status eq '${status}'`)
          if (!includeCompleted && !status) filters.push("status ne 'completed'")
          if (importance) filters.push(`importance eq '${importance}'`)

          const queryParams = new URLSearchParams()
          if (filters.length > 0) queryParams.append("$filter", filters.join(" and "))
          queryParams.append("$expand", "linkedResources,checklistItems")

          const queryString = queryParams.toString()
          const url = `${MS_GRAPH_BASE}/me/todo/lists/${list.id}/tasks${queryString ? "?" + queryString : ""}`

          const tasksResponse = await makeGraphRequest<{ value: Task[] }>(url, token)
          if (!tasksResponse) continue

          for (const task of tasksResponse.value || []) {
            // Client-side keyword filter (OData doesn't support contains on title for To Do)
            if (keyword && !task.title.toLowerCase().includes(keyword.toLowerCase())) continue

            // Client-side date range filters
            if (dueBefore && task.dueDateTime) {
              const dueDate = task.dueDateTime.dateTime.slice(0, 10)
              if (dueDate > dueBefore.slice(0, 10)) continue
            }
            if (dueAfter && task.dueDateTime) {
              const dueDate = task.dueDateTime.dateTime.slice(0, 10)
              if (dueDate < dueAfter.slice(0, 10)) continue
            }
            // If filtering by date range, skip tasks without due dates
            if ((dueBefore || dueAfter) && !task.dueDateTime) continue

            allResults.push({ listName: list.displayName, task })
          }
        }

        if (allResults.length === 0) {
          return {
            content: [{ type: "text", text: "No tasks found matching your criteria." }],
          }
        }

        const formatted = allResults.map(({ listName, task }) => `[${listName}]\n${formatTask(task)}`).join("\n")

        return {
          content: [{ type: "text", text: `Found ${allResults.length} task(s):\n\n${formatted}` }],
        }
      } catch (error) {
        return {
          content: [{ type: "text", text: `Error searching tasks: ${error}` }],
        }
      }
    },
  )

  server.tool(
    "get-todays-tasks",
    "Get all tasks due today or overdue across all lists. Provides a unified daily view of what needs attention.",
    {
      includeOverdue: z.boolean().optional().describe("Include overdue tasks (default: true)"),
      includeNoDueDate: z.boolean().optional().describe("Include tasks with no due date (default: false)"),
    },
    async ({ includeOverdue = true, includeNoDueDate = false }) => {
      try {
        const token = await getAccessToken()
        if (!token) {
          return {
            content: [{ type: "text", text: "Failed to authenticate with Microsoft API" }],
          }
        }

        const today = new Date()
        const todayStr = `${today.getUTCFullYear()}-${String(today.getUTCMonth() + 1).padStart(2, "0")}-${String(today.getUTCDate()).padStart(2, "0")}`

        const listsResponse = await makeGraphRequest<{ value: TaskList[] }>(`${MS_GRAPH_BASE}/me/todo/lists`, token)
        if (!listsResponse) {
          return {
            content: [{ type: "text", text: "Failed to retrieve task lists" }],
          }
        }

        const todayTasks: { listName: string; task: Task }[] = []
        const overdueTasks: { listName: string; task: Task }[] = []
        const noDueDateTasks: { listName: string; task: Task }[] = []

        for (const list of listsResponse.value || []) {
          const queryParams = new URLSearchParams()
          queryParams.append("$filter", "status ne 'completed'")
          queryParams.append("$expand", "linkedResources,checklistItems")

          const url = `${MS_GRAPH_BASE}/me/todo/lists/${list.id}/tasks?${queryParams.toString()}`
          const tasksResponse = await makeGraphRequest<{ value: Task[] }>(url, token)
          if (!tasksResponse) continue

          for (const task of tasksResponse.value || []) {
            if (!task.dueDateTime) {
              if (includeNoDueDate) {
                noDueDateTasks.push({ listName: list.displayName, task })
              }
              continue
            }

            const dueDate = task.dueDateTime.dateTime.slice(0, 10)

            if (dueDate === todayStr) {
              todayTasks.push({ listName: list.displayName, task })
            } else if (dueDate < todayStr && includeOverdue) {
              overdueTasks.push({ listName: list.displayName, task })
            }
          }
        }

        let output = ""

        if (overdueTasks.length > 0) {
          output += `OVERDUE (${overdueTasks.length}):\n\n`
          output += overdueTasks.map(({ listName, task }) => `[${listName}]\n${formatTask(task)}`).join("\n")
          output += "\n\n"
        }

        if (todayTasks.length > 0) {
          output += `DUE TODAY (${todayTasks.length}):\n\n`
          output += todayTasks.map(({ listName, task }) => `[${listName}]\n${formatTask(task)}`).join("\n")
          output += "\n\n"
        }

        if (noDueDateTasks.length > 0) {
          output += `NO DUE DATE (${noDueDateTasks.length}):\n\n`
          output += noDueDateTasks.map(({ listName, task }) => `[${listName}]\n${formatTask(task)}`).join("\n")
          output += "\n\n"
        }

        if (!output) {
          output = "No tasks due today and no overdue tasks. You're all caught up!"
        } else {
          output = `Daily Task Summary (${todayStr})\n${"=".repeat(40)}\n\n${output}`
        }

        return {
          content: [{ type: "text", text: output.trim() }],
        }
      } catch (error) {
        return {
          content: [{ type: "text", text: `Error fetching today's tasks: ${error}` }],
        }
      }
    },
  )
}
