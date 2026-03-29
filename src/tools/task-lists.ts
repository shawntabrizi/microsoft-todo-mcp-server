import type { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js"
import { z } from "zod"
import { makeGraphRequest, getAccessToken, MS_GRAPH_BASE, isAllowedGraphUrl } from "../graph-client.js"
import type { TaskList, DeltaResponse } from "../types.js"

export function register(server: McpServer) {
  // Task List tools
  server.tool(
    "get-task-lists",
    "Get all Microsoft Todo task lists (the top-level containers that organize your tasks). Shows list names, IDs, and indicates default or shared lists.",
    {},
    async () => {
      try {
        const token = await getAccessToken()
        if (!token) {
          return {
            content: [{ type: "text", text: "Failed to authenticate with Microsoft API" }],
          }
        }

        const response = await makeGraphRequest<{ value: TaskList[] }>(`${MS_GRAPH_BASE}/me/todo/lists`, token)

        if (!response) {
          return {
            content: [{ type: "text", text: "Failed to retrieve task lists" }],
          }
        }

        const lists = response.value || []
        if (lists.length === 0) {
          return {
            content: [{ type: "text", text: "No task lists found." }],
          }
        }

        const formattedLists = lists.map((list) => {
          let wellKnownInfo = ""
          if (list.wellknownListName && list.wellknownListName !== "none") {
            if (list.wellknownListName === "defaultList") {
              wellKnownInfo = " (Default Tasks List)"
            } else if (list.wellknownListName === "flaggedEmails") {
              wellKnownInfo = " (Flagged Emails)"
            }
          }

          let sharingInfo = ""
          if (list.isShared) {
            sharingInfo = list.isOwner ? " (Shared by you)" : " (Shared with you)"
          }

          return `ID: ${list.id}\nName: ${list.displayName}${wellKnownInfo}${sharingInfo}\n---`
        })

        return {
          content: [{ type: "text", text: `Your task lists:\n\n${formattedLists.join("\n")}` }],
        }
      } catch (error) {
        return {
          content: [{ type: "text", text: `Error fetching task lists: ${error}` }],
        }
      }
    },
  )

  server.tool(
    "get-task-list",
    "Get a single Microsoft Todo task list by ID.",
    {
      listId: z.string().describe("ID of the task list"),
      select: z.string().optional().describe("Comma-separated list of properties to include"),
    },
    async ({ listId, select }) => {
      try {
        const token = await getAccessToken()
        if (!token) {
          return {
            content: [{ type: "text", text: "Failed to authenticate with Microsoft API" }],
          }
        }

        const queryParams = new URLSearchParams()
        if (select) queryParams.append("$select", select)

        const queryString = queryParams.toString()
        const url = `${MS_GRAPH_BASE}/me/todo/lists/${listId}${queryString ? "?" + queryString : ""}`
        const list = await makeGraphRequest<TaskList>(url, token)

        if (!list) {
          return {
            content: [{ type: "text", text: `Failed to retrieve task list: ${listId}` }],
          }
        }

        const metadata = []
        if (list.wellknownListName && list.wellknownListName !== "none")
          metadata.push(`Type: ${list.wellknownListName}`)
        if (list.isShared !== undefined) metadata.push(`Shared: ${list.isShared ? "Yes" : "No"}`)
        if (list.isOwner !== undefined) metadata.push(`Owner: ${list.isOwner ? "Yes" : "No"}`)

        return {
          content: [
            {
              type: "text",
              text: `Task list details:\n\nID: ${list.id}\nName: ${list.displayName}${metadata.length ? `\n${metadata.join("\n")}` : ""}`,
            },
          ],
        }
      } catch (error) {
        return {
          content: [{ type: "text", text: `Error fetching task list: ${error}` }],
        }
      }
    },
  )

  server.tool(
    "get-task-lists-delta",
    "Track changes to Microsoft Todo task lists using the Graph delta API.",
    {
      deltaUrl: z
        .string()
        .optional()
        .describe("Full @odata.nextLink or @odata.deltaLink URL from a previous delta response"),
      deltaToken: z.string().optional().describe("Delta token from a previous delta response"),
      skipToken: z.string().optional().describe("Skip token from a previous delta response"),
      select: z.string().optional().describe("Comma-separated list of properties to include on the initial request"),
      maxPageSize: z.number().int().min(1).optional().describe("Preferred maximum number of lists returned"),
    },
    async ({ deltaUrl, deltaToken, skipToken, select, maxPageSize }) => {
      try {
        const token = await getAccessToken()
        if (!token) {
          return {
            content: [{ type: "text", text: "Failed to authenticate with Microsoft API" }],
          }
        }

        let url = deltaUrl || `${MS_GRAPH_BASE}/me/todo/lists/delta`
        if (deltaUrl && !isAllowedGraphUrl(deltaUrl)) {
          return {
            content: [
              {
                type: "text",
                text: "Invalid deltaUrl: must be a Microsoft Graph API URL (https://graph.microsoft.com/...)",
              },
            ],
          }
        }
        if (!deltaUrl) {
          const queryParams = new URLSearchParams()
          if (deltaToken) queryParams.append("$deltatoken", deltaToken)
          if (skipToken) queryParams.append("$skiptoken", skipToken)
          if (select) queryParams.append("$select", select)
          const queryString = queryParams.toString()
          if (queryString) url += `?${queryString}`
        }

        const data = await makeGraphRequest<DeltaResponse<TaskList>>(url, token)
        if (!data) {
          return {
            content: [{ type: "text", text: "Failed to fetch task list delta" }],
          }
        }
        const formattedLists = (data.value || [])
          .map((list) => `ID: ${list.id}\nName: ${list.displayName}\n---`)
          .join("\n")

        return {
          content: [
            {
              type: "text",
              text:
                `Task list delta results:\n\n${formattedLists || "No changed lists returned."}` +
                `${data["@odata.nextLink"] ? `\n\nNext Link:\n${data["@odata.nextLink"]}` : ""}` +
                `${data["@odata.deltaLink"] ? `\n\nDelta Link:\n${data["@odata.deltaLink"]}` : ""}`,
            },
          ],
        }
      } catch (error) {
        return {
          content: [{ type: "text", text: `Error fetching task list delta: ${error}` }],
        }
      }
    },
  )

  server.tool(
    "create-task-list",
    "Create a new task list (top-level container) in Microsoft Todo to help organize your tasks into categories or projects.",
    {
      displayName: z.string().describe("Name of the new task list"),
    },
    async ({ displayName }) => {
      try {
        const token = await getAccessToken()
        if (!token) {
          return {
            content: [{ type: "text", text: "Failed to authenticate with Microsoft API" }],
          }
        }

        const response = await makeGraphRequest<TaskList>(`${MS_GRAPH_BASE}/me/todo/lists`, token, "POST", {
          displayName,
        })

        if (!response) {
          return {
            content: [{ type: "text", text: `Failed to create task list: ${displayName}` }],
          }
        }

        return {
          content: [
            {
              type: "text",
              text: `Task list created successfully!\nName: ${response.displayName}\nID: ${response.id}`,
            },
          ],
        }
      } catch (error) {
        return {
          content: [{ type: "text", text: `Error creating task list: ${error}` }],
        }
      }
    },
  )

  server.tool(
    "update-task-list",
    "Update the name of an existing task list (top-level container) in Microsoft Todo.",
    {
      listId: z.string().describe("ID of the task list to update"),
      displayName: z.string().describe("New name for the task list"),
    },
    async ({ listId, displayName }) => {
      try {
        const token = await getAccessToken()
        if (!token) {
          return {
            content: [{ type: "text", text: "Failed to authenticate with Microsoft API" }],
          }
        }

        const response = await makeGraphRequest<TaskList>(`${MS_GRAPH_BASE}/me/todo/lists/${listId}`, token, "PATCH", {
          displayName,
        })

        if (!response) {
          return {
            content: [{ type: "text", text: `Failed to update task list with ID: ${listId}` }],
          }
        }

        return {
          content: [{ type: "text", text: `Task list updated successfully!\nNew name: ${response.displayName}` }],
        }
      } catch (error) {
        return {
          content: [{ type: "text", text: `Error updating task list: ${error}` }],
        }
      }
    },
  )

  server.tool(
    "delete-task-list",
    "Delete a task list (top-level container) from Microsoft Todo. This will remove the list and all tasks within it.",
    {
      listId: z.string().describe("ID of the task list to delete"),
    },
    async ({ listId }) => {
      try {
        const token = await getAccessToken()
        if (!token) {
          return {
            content: [{ type: "text", text: "Failed to authenticate with Microsoft API" }],
          }
        }

        await makeGraphRequest<null>(`${MS_GRAPH_BASE}/me/todo/lists/${listId}`, token, "DELETE")

        return {
          content: [{ type: "text", text: `Task list with ID: ${listId} was successfully deleted.` }],
        }
      } catch (error) {
        return {
          content: [{ type: "text", text: `Error deleting task list: ${error}` }],
        }
      }
    },
  )
}
