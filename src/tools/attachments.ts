import type { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js"
import { z } from "zod"
import { makeGraphRequest, getAccessToken, MS_GRAPH_BASE } from "../graph-client.js"
import { isTaskFileAttachment } from "../helpers.js"
import type { TaskFileAttachment, UploadSession } from "../types.js"

export function register(server: McpServer) {
  // Attachment tools
  server.tool(
    "get-attachments",
    "List file attachments for a Microsoft Todo task.",
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

        const response = await makeGraphRequest<{ value: TaskFileAttachment[] }>(
          `${MS_GRAPH_BASE}/me/todo/lists/${listId}/tasks/${taskId}/attachments`,
          token,
        )

        if (!response) {
          return {
            content: [{ type: "text", text: `Failed to retrieve attachments for task: ${taskId}` }],
          }
        }

        const attachments = response.value || []
        if (attachments.length === 0) {
          return {
            content: [{ type: "text", text: `No attachments found for task: ${taskId}` }],
          }
        }

        const formattedAttachments = attachments.map((attachment) => {
          const details = [
            `ID: ${attachment.id}`,
            `Name: ${attachment.name}`,
            `Size: ${attachment.size ?? "Unknown"} bytes`,
          ]

          if (attachment.contentType) details.push(`Content Type: ${attachment.contentType}`)
          if (attachment.lastModifiedDateTime) {
            details.push(`Last Modified: ${new Date(attachment.lastModifiedDateTime).toLocaleString()}`)
          }

          return `${details.join("\n")}\n---`
        })

        return {
          content: [{ type: "text", text: `Attachments for task ${taskId}:\n\n${formattedAttachments.join("\n")}` }],
        }
      } catch (error) {
        return {
          content: [{ type: "text", text: `Error fetching attachments: ${error}` }],
        }
      }
    },
  )

  server.tool(
    "get-attachment",
    "Get a single file attachment for a Microsoft Todo task.",
    {
      listId: z.string().describe("ID of the task list"),
      taskId: z.string().describe("ID of the task"),
      attachmentId: z.string().describe("ID of the attachment"),
    },
    async ({ listId, taskId, attachmentId }) => {
      try {
        const token = await getAccessToken()
        if (!token) {
          return {
            content: [{ type: "text", text: "Failed to authenticate with Microsoft API" }],
          }
        }

        const response = await makeGraphRequest<{ value?: TaskFileAttachment } | TaskFileAttachment>(
          `${MS_GRAPH_BASE}/me/todo/lists/${listId}/tasks/${taskId}/attachments/${attachmentId}`,
          token,
        )

        let attachment: TaskFileAttachment | undefined
        if (response && typeof response === "object" && "value" in response) {
          attachment = response.value
        } else if (isTaskFileAttachment(response)) {
          attachment = response
        }

        if (!attachment) {
          return {
            content: [{ type: "text", text: `Failed to retrieve attachment: ${attachmentId}` }],
          }
        }

        return {
          content: [
            {
              type: "text",
              text:
                `Attachment details:\n\nID: ${attachment.id}\nName: ${attachment.name}` +
                `${attachment.contentType ? `\nContent Type: ${attachment.contentType}` : ""}` +
                `${attachment.size !== undefined ? `\nSize: ${attachment.size} bytes` : ""}` +
                `${attachment.lastModifiedDateTime ? `\nLast Modified: ${new Date(attachment.lastModifiedDateTime).toLocaleString()}` : ""}` +
                `${attachment.contentBytes ? `\nContent Bytes Present: Yes (${attachment.contentBytes.length} base64 chars)` : ""}`,
            },
          ],
        }
      } catch (error) {
        return {
          content: [{ type: "text", text: `Error fetching attachment: ${error}` }],
        }
      }
    },
  )

  server.tool(
    "create-attachment",
    "Create a small file attachment on a Microsoft Todo task. For files larger than 3 MB, use create-attachment-upload-session.",
    {
      listId: z.string().describe("ID of the task list"),
      taskId: z.string().describe("ID of the task"),
      name: z.string().describe("Attachment display name"),
      contentBytes: z.string().describe("Base64-encoded file contents"),
      contentType: z.string().optional().describe("Attachment content type"),
    },
    async ({ listId, taskId, name, contentBytes, contentType }) => {
      try {
        const token = await getAccessToken()
        if (!token) {
          return {
            content: [{ type: "text", text: "Failed to authenticate with Microsoft API" }],
          }
        }

        const response = await makeGraphRequest<TaskFileAttachment>(
          `${MS_GRAPH_BASE}/me/todo/lists/${listId}/tasks/${taskId}/attachments`,
          token,
          "POST",
          {
            "@odata.type": "#microsoft.graph.taskFileAttachment",
            name,
            contentBytes,
            contentType,
          },
        )

        if (!response) {
          return {
            content: [{ type: "text", text: `Failed to create attachment for task: ${taskId}` }],
          }
        }

        return {
          content: [
            {
              type: "text",
              text:
                `Attachment created successfully!\nID: ${response.id}\nName: ${response.name}` +
                `${response.size !== undefined ? `\nSize: ${response.size} bytes` : ""}`,
            },
          ],
        }
      } catch (error) {
        return {
          content: [{ type: "text", text: `Error creating attachment: ${error}` }],
        }
      }
    },
  )

  server.tool(
    "create-attachment-upload-session",
    "Create an upload session for a large file attachment on a Microsoft Todo task.",
    {
      listId: z.string().describe("ID of the task list"),
      taskId: z.string().describe("ID of the task"),
      name: z.string().describe("Attachment display name"),
      size: z.number().int().min(0).describe("Attachment size in bytes"),
    },
    async ({ listId, taskId, name, size }) => {
      try {
        const token = await getAccessToken()
        if (!token) {
          return {
            content: [{ type: "text", text: "Failed to authenticate with Microsoft API" }],
          }
        }

        const response = await makeGraphRequest<UploadSession>(
          `${MS_GRAPH_BASE}/me/todo/lists/${listId}/tasks/${taskId}/attachments/createUploadSession`,
          token,
          "POST",
          {
            attachmentInfo: {
              attachmentType: "file",
              name,
              size,
            },
          },
        )

        if (!response) {
          return {
            content: [{ type: "text", text: `Failed to create attachment upload session for task: ${taskId}` }],
          }
        }

        return {
          content: [
            {
              type: "text",
              text: `Attachment upload session created successfully!\nUpload URL: ${response.uploadUrl}\nExpiration: ${response.expirationDateTime}\nNext Expected Ranges: ${response.nextExpectedRanges.join(", ")}`,
            },
          ],
        }
      } catch (error) {
        return {
          content: [{ type: "text", text: `Error creating attachment upload session: ${error}` }],
        }
      }
    },
  )

  server.tool(
    "delete-attachment",
    "Delete a file attachment from a Microsoft Todo task.",
    {
      listId: z.string().describe("ID of the task list"),
      taskId: z.string().describe("ID of the task"),
      attachmentId: z.string().describe("ID of the attachment"),
    },
    async ({ listId, taskId, attachmentId }) => {
      try {
        const token = await getAccessToken()
        if (!token) {
          return {
            content: [{ type: "text", text: "Failed to authenticate with Microsoft API" }],
          }
        }

        await makeGraphRequest<null>(
          `${MS_GRAPH_BASE}/me/todo/lists/${listId}/tasks/${taskId}/attachments/${attachmentId}`,
          token,
          "DELETE",
        )

        return {
          content: [
            { type: "text", text: `Attachment with ID: ${attachmentId} was successfully deleted from task: ${taskId}` },
          ],
        }
      } catch (error) {
        return {
          content: [{ type: "text", text: `Error deleting attachment: ${error}` }],
        }
      }
    },
  )
}
