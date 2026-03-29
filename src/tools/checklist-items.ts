import type { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js"
import { z } from "zod"
import { makeGraphRequest, getAccessToken, MS_GRAPH_BASE } from "../graph-client.js"
import type { Task, ChecklistItem } from "../types.js"

export function register(server: McpServer) {
  // Checklist item tools
  server.tool(
    "get-checklist-items",
    "Get checklist items (subtasks) for a specific task. Checklist items are smaller steps or components that belong to a parent task.",
    {
      listId: z.string().describe("ID of the task list"),
      taskId: z.string().describe("ID of the task"),
    },
    async ({ listId, taskId }) => {
      try {
        const token = await getAccessToken()
        if (!token) {
          return {
            content: [{ type: "text", text: "Failed to authenticate with Microsoft API" }],
          }
        }

        const taskResponse = await makeGraphRequest<Task>(
          `${MS_GRAPH_BASE}/me/todo/lists/${listId}/tasks/${taskId}`,
          token,
        )

        const taskTitle = taskResponse ? taskResponse.title : "Unknown Task"

        const response = await makeGraphRequest<{ value: ChecklistItem[] }>(
          `${MS_GRAPH_BASE}/me/todo/lists/${listId}/tasks/${taskId}/checklistItems`,
          token,
        )

        if (!response) {
          return {
            content: [{ type: "text", text: `Failed to retrieve checklist items for task: ${taskId}` }],
          }
        }

        const items = response.value || []
        if (items.length === 0) {
          return {
            content: [{ type: "text", text: `No checklist items found for task "${taskTitle}" (ID: ${taskId})` }],
          }
        }

        const formattedItems = items.map((item) => {
          const status = item.isChecked ? "\u2713" : "\u25CB"
          let itemInfo = `${status} ${item.displayName} (ID: ${item.id})`

          if (item.createdDateTime) {
            itemInfo += `\nCreated: ${new Date(item.createdDateTime).toLocaleString()}`
          }

          if (item.checkedDateTime) {
            itemInfo += `\nChecked: ${new Date(item.checkedDateTime).toLocaleString()}`
          }

          return itemInfo
        })

        return {
          content: [
            {
              type: "text",
              text: `Checklist items for task "${taskTitle}" (ID: ${taskId}):\n\n${formattedItems.join("\n\n")}`,
            },
          ],
        }
      } catch (error) {
        return {
          content: [{ type: "text", text: `Error fetching checklist items: ${error}` }],
        }
      }
    },
  )

  server.tool(
    "get-checklist-item",
    "Get a single checklist item (subtask) by ID.",
    {
      listId: z.string().describe("ID of the task list"),
      taskId: z.string().describe("ID of the task"),
      checklistItemId: z.string().describe("ID of the checklist item"),
    },
    async ({ listId, taskId, checklistItemId }) => {
      try {
        const token = await getAccessToken()
        if (!token) {
          return {
            content: [{ type: "text", text: "Failed to authenticate with Microsoft API" }],
          }
        }

        const item = await makeGraphRequest<ChecklistItem>(
          `${MS_GRAPH_BASE}/me/todo/lists/${listId}/tasks/${taskId}/checklistItems/${checklistItemId}`,
          token,
        )

        if (!item) {
          return {
            content: [{ type: "text", text: `Failed to retrieve checklist item: ${checklistItemId}` }],
          }
        }

        const status = item.isChecked ? "Checked" : "Not checked"
        let details = `${status}: ${item.displayName}\nID: ${item.id}`
        if (item.createdDateTime) details += `\nCreated: ${new Date(item.createdDateTime).toLocaleString()}`
        if (item.checkedDateTime) details += `\nChecked: ${new Date(item.checkedDateTime).toLocaleString()}`

        return {
          content: [{ type: "text", text: `Checklist item details:\n\n${details}` }],
        }
      } catch (error) {
        return {
          content: [{ type: "text", text: `Error fetching checklist item: ${error}` }],
        }
      }
    },
  )

  server.tool(
    "create-checklist-item",
    "Create a new checklist item (subtask) for a task. Checklist items help break down a task into smaller, manageable steps.",
    {
      listId: z.string().describe("ID of the task list"),
      taskId: z.string().describe("ID of the task"),
      displayName: z.string().describe("Text content of the checklist item"),
      isChecked: z.boolean().optional().describe("Whether the item is checked off"),
      checkedDateTime: z.string().optional().describe("Completion timestamp in ISO format"),
      createdDateTime: z.string().optional().describe("Creation timestamp in ISO format"),
    },
    async ({ listId, taskId, displayName, isChecked, checkedDateTime, createdDateTime }) => {
      try {
        const token = await getAccessToken()
        if (!token) {
          return {
            content: [{ type: "text", text: "Failed to authenticate with Microsoft API" }],
          }
        }

        const requestBody: any = { displayName }
        if (isChecked !== undefined) requestBody.isChecked = isChecked
        if (checkedDateTime !== undefined) requestBody.checkedDateTime = checkedDateTime
        if (createdDateTime !== undefined) requestBody.createdDateTime = createdDateTime

        const response = await makeGraphRequest<ChecklistItem>(
          `${MS_GRAPH_BASE}/me/todo/lists/${listId}/tasks/${taskId}/checklistItems`,
          token,
          "POST",
          requestBody,
        )

        if (!response) {
          return {
            content: [{ type: "text", text: `Failed to create checklist item for task: ${taskId}` }],
          }
        }

        return {
          content: [
            {
              type: "text",
              text: `Checklist item created successfully!\nContent: ${response.displayName}\nID: ${response.id}`,
            },
          ],
        }
      } catch (error) {
        return {
          content: [{ type: "text", text: `Error creating checklist item: ${error}` }],
        }
      }
    },
  )

  server.tool(
    "update-checklist-item",
    "Update an existing checklist item (subtask). Allows changing the text content or completion status of the subtask.",
    {
      listId: z.string().describe("ID of the task list"),
      taskId: z.string().describe("ID of the task"),
      checklistItemId: z.string().describe("ID of the checklist item to update"),
      displayName: z.string().optional().describe("New text content of the checklist item"),
      isChecked: z.boolean().optional().describe("Whether the item is checked off"),
      checkedDateTime: z
        .string()
        .optional()
        .describe("Completion timestamp in ISO format. Pass an empty string to clear it."),
      createdDateTime: z.string().optional().describe("Creation timestamp in ISO format"),
    },
    async ({ listId, taskId, checklistItemId, displayName, isChecked, checkedDateTime, createdDateTime }) => {
      try {
        const token = await getAccessToken()
        if (!token) {
          return {
            content: [{ type: "text", text: "Failed to authenticate with Microsoft API" }],
          }
        }

        const requestBody: any = {}
        if (displayName !== undefined) requestBody.displayName = displayName
        if (isChecked !== undefined) requestBody.isChecked = isChecked
        if (checkedDateTime !== undefined) requestBody.checkedDateTime = checkedDateTime === "" ? null : checkedDateTime
        if (createdDateTime !== undefined) requestBody.createdDateTime = createdDateTime

        if (Object.keys(requestBody).length === 0) {
          return {
            content: [
              {
                type: "text",
                text: "No properties provided for update. Please specify at least one checklist item property to change.",
              },
            ],
          }
        }

        const response = await makeGraphRequest<ChecklistItem>(
          `${MS_GRAPH_BASE}/me/todo/lists/${listId}/tasks/${taskId}/checklistItems/${checklistItemId}`,
          token,
          "PATCH",
          requestBody,
        )

        if (!response) {
          return {
            content: [{ type: "text", text: `Failed to update checklist item with ID: ${checklistItemId}` }],
          }
        }

        const statusText = response.isChecked ? "Checked" : "Not checked"

        return {
          content: [
            {
              type: "text",
              text: `Checklist item updated successfully!\nContent: ${response.displayName}\nStatus: ${statusText}`,
            },
          ],
        }
      } catch (error) {
        return {
          content: [{ type: "text", text: `Error updating checklist item: ${error}` }],
        }
      }
    },
  )

  server.tool(
    "delete-checklist-item",
    "Delete a checklist item (subtask) from a task. This removes just the specific subtask, not the parent task.",
    {
      listId: z.string().describe("ID of the task list"),
      taskId: z.string().describe("ID of the task"),
      checklistItemId: z.string().describe("ID of the checklist item to delete"),
    },
    async ({ listId, taskId, checklistItemId }) => {
      try {
        const token = await getAccessToken()
        if (!token) {
          return {
            content: [{ type: "text", text: "Failed to authenticate with Microsoft API" }],
          }
        }

        await makeGraphRequest<null>(
          `${MS_GRAPH_BASE}/me/todo/lists/${listId}/tasks/${taskId}/checklistItems/${checklistItemId}`,
          token,
          "DELETE",
        )

        return {
          content: [
            {
              type: "text",
              text: `Checklist item with ID: ${checklistItemId} was successfully deleted from task: ${taskId}`,
            },
          ],
        }
      } catch (error) {
        return {
          content: [{ type: "text", text: `Error deleting checklist item: ${error}` }],
        }
      }
    },
  )
}
