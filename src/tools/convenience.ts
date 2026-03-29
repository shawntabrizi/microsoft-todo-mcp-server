import type { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js"
import { z } from "zod"
import { makeGraphRequest, getAccessToken, MS_GRAPH_BASE } from "../graph-client.js"
import { formatTask } from "../helpers.js"
import type { Task, TaskList } from "../types.js"

// Fetch all pages of a paginated Graph API response
async function fetchAllPages<T>(url: string, token: string): Promise<{ items: T[]; errors: string[] }> {
  const items: T[] = []
  const errors: string[] = []
  let nextUrl: string | undefined = url

  while (nextUrl) {
    const response = await makeGraphRequest<{ value: T[]; "@odata.nextLink"?: string }>(nextUrl, token)
    if (!response) {
      errors.push(nextUrl)
      break
    }
    items.push(...(response.value || []))
    nextUrl = response["@odata.nextLink"]
  }

  return { items, errors }
}

// Fetch all tasks from all lists, with pagination and error tracking
async function fetchTasksAcrossLists(
  token: string,
  buildTaskUrl: (listId: string) => string,
): Promise<{ results: { listName: string; task: Task }[]; warnings: string[] }> {
  const warnings: string[] = []

  const { items: lists, errors: listErrors } = await fetchAllPages<TaskList>(`${MS_GRAPH_BASE}/me/todo/lists`, token)
  if (listErrors.length > 0) {
    warnings.push("Failed to fetch some task list pages")
  }

  const results: { listName: string; task: Task }[] = []

  for (const list of lists) {
    const url = buildTaskUrl(list.id)
    const { items: tasks, errors: taskErrors } = await fetchAllPages<Task>(url, token)
    if (taskErrors.length > 0) {
      warnings.push(`Failed to fetch some tasks from list "${list.displayName}"`)
    }
    for (const task of tasks) {
      results.push({ listName: list.displayName, task })
    }
  }

  return { results, warnings }
}

export function register(server: McpServer) {
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
      linkedApp: z
        .string()
        .optional()
        .describe(
          "Filter to tasks with a linked resource from this application name (case-insensitive, e.g. 'Ninety')",
        ),
    },
    async ({ keyword, status, importance, dueBefore, dueAfter, includeCompleted = false, linkedApp }) => {
      try {
        const token = await getAccessToken()
        if (!token) {
          return {
            content: [{ type: "text", text: "Failed to authenticate with Microsoft API" }],
          }
        }

        const { results: allTaskEntries, warnings } = await fetchTasksAcrossLists(token, (listId) => {
          const filters: string[] = []
          if (status) filters.push(`status eq '${status}'`)
          if (!includeCompleted && !status) filters.push("status ne 'completed'")
          if (importance) filters.push(`importance eq '${importance}'`)

          const queryParams = new URLSearchParams()
          if (filters.length > 0) queryParams.append("$filter", filters.join(" and "))
          // Expand linkedResources when filtering by app name
          if (linkedApp) queryParams.append("$expand", "linkedResources")

          const queryString = queryParams.toString()
          return `${MS_GRAPH_BASE}/me/todo/lists/${listId}/tasks${queryString ? "?" + queryString : ""}`
        })

        // Client-side filters (OData doesn't support contains on title or linked resource filtering)
        const filtered = allTaskEntries.filter(({ task }) => {
          if (keyword && !task.title.toLowerCase().includes(keyword.toLowerCase())) return false

          if (dueBefore && task.dueDateTime) {
            if (task.dueDateTime.dateTime.slice(0, 10) > dueBefore.slice(0, 10)) return false
          }
          if (dueAfter && task.dueDateTime) {
            if (task.dueDateTime.dateTime.slice(0, 10) < dueAfter.slice(0, 10)) return false
          }
          if ((dueBefore || dueAfter) && !task.dueDateTime) return false

          if (linkedApp) {
            if (!task.linkedResources || task.linkedResources.length === 0) return false
            const appLower = linkedApp.toLowerCase()
            const hasMatch = task.linkedResources.some(
              (r) =>
                r.applicationName?.toLowerCase().includes(appLower) || r.displayName?.toLowerCase().includes(appLower),
            )
            if (!hasMatch) return false
          }

          return true
        })

        let warningText = ""
        if (warnings.length > 0) {
          warningText = `\n\nWarnings (results may be incomplete):\n${warnings.map((w) => `- ${w}`).join("\n")}`
        }

        if (filtered.length === 0) {
          return {
            content: [{ type: "text", text: `No tasks found matching your criteria.${warningText}` }],
          }
        }

        const formatted = filtered.map(({ listName, task }) => `[${listName}]\n${formatTask(task)}`).join("\n")

        return {
          content: [{ type: "text", text: `Found ${filtered.length} task(s):\n\n${formatted}${warningText}` }],
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
    "Get all tasks due today or overdue across all lists. Provides a unified daily view of what needs attention. Uses local system time for date comparison.",
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

        // Use local time, not UTC, so "today" matches the user's perspective
        const today = new Date()
        const todayStr = `${today.getFullYear()}-${String(today.getMonth() + 1).padStart(2, "0")}-${String(today.getDate()).padStart(2, "0")}`

        const { results: allTaskEntries, warnings } = await fetchTasksAcrossLists(token, (listId) => {
          const queryParams = new URLSearchParams()
          queryParams.append("$filter", "status ne 'completed'")
          return `${MS_GRAPH_BASE}/me/todo/lists/${listId}/tasks?${queryParams.toString()}`
        })

        const todayTasks: { listName: string; task: Task }[] = []
        const overdueTasks: { listName: string; task: Task }[] = []
        const noDueDateTasks: { listName: string; task: Task }[] = []

        for (const { listName, task } of allTaskEntries) {
          if (!task.dueDateTime) {
            if (includeNoDueDate) {
              noDueDateTasks.push({ listName, task })
            }
            continue
          }

          const dueDate = task.dueDateTime.dateTime.slice(0, 10)

          if (dueDate === todayStr) {
            todayTasks.push({ listName, task })
          } else if (dueDate < todayStr && includeOverdue) {
            overdueTasks.push({ listName, task })
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

        let warningText = ""
        if (warnings.length > 0) {
          warningText = `\nWarnings (results may be incomplete):\n${warnings.map((w) => `- ${w}`).join("\n")}\n`
        }

        if (!output) {
          output = `No tasks due today and no overdue tasks. You're all caught up!${warningText ? `\n${warningText}` : ""}`
        } else {
          output = `Daily Task Summary (${todayStr})\n${"=".repeat(40)}\n\n${output}${warningText}`
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
