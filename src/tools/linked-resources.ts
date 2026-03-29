import type { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js"
import { z } from "zod"
import { makeGraphRequest, getAccessToken, MS_GRAPH_BASE } from "../graph-client.js"
import type { LinkedResource } from "../types.js"

export function register(server: McpServer) {
  // Linked resource tools
  server.tool(
    "get-linked-resources",
    "Get linked resources for a Microsoft Todo task.",
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

        const response = await makeGraphRequest<{ value: LinkedResource[] }>(
          `${MS_GRAPH_BASE}/me/todo/lists/${listId}/tasks/${taskId}/linkedResources`,
          token,
        )

        if (!response) {
          return {
            content: [{ type: "text", text: `Failed to retrieve linked resources for task: ${taskId}` }],
          }
        }

        const linkedResources = response.value || []
        if (linkedResources.length === 0) {
          return {
            content: [{ type: "text", text: `No linked resources found for task: ${taskId}` }],
          }
        }

        const formattedResources = linkedResources.map((resource) => {
          const details = [
            `ID: ${resource.id || "Unknown"}`,
            `Display Name: ${resource.displayName || "Unknown"}`,
            `Application: ${resource.applicationName || "Unknown"}`,
          ]

          if (resource.webUrl) details.push(`URL: ${resource.webUrl}`)
          if (resource.externalId) details.push(`External ID: ${resource.externalId}`)

          return `${details.join("\n")}\n---`
        })

        return {
          content: [{ type: "text", text: `Linked resources for task ${taskId}:\n\n${formattedResources.join("\n")}` }],
        }
      } catch (error) {
        return {
          content: [{ type: "text", text: `Error fetching linked resources: ${error}` }],
        }
      }
    },
  )

  server.tool(
    "create-linked-resource",
    "Create a linked resource for a Microsoft Todo task.",
    {
      listId: z.string().describe("ID of the task list"),
      taskId: z.string().describe("ID of the task"),
      webUrl: z.string().optional().describe("Deep link to the linked item"),
      applicationName: z.string().optional().describe("Source application name"),
      displayName: z.string().optional().describe("Display name for the linked item"),
      externalId: z.string().optional().describe("External identifier from the source system"),
    },
    async ({ listId, taskId, webUrl, applicationName, displayName, externalId }) => {
      try {
        const token = await getAccessToken()
        if (!token) {
          return {
            content: [{ type: "text", text: "Failed to authenticate with Microsoft API" }],
          }
        }

        const response = await makeGraphRequest<LinkedResource>(
          `${MS_GRAPH_BASE}/me/todo/lists/${listId}/tasks/${taskId}/linkedResources`,
          token,
          "POST",
          { webUrl, applicationName, displayName, externalId },
        )

        if (!response) {
          return {
            content: [{ type: "text", text: `Failed to create linked resource for task: ${taskId}` }],
          }
        }

        return {
          content: [
            {
              type: "text",
              text:
                `Linked resource created successfully!\nID: ${response.id || "Unknown"}` +
                `${response.displayName ? `\nDisplay Name: ${response.displayName}` : ""}` +
                `${response.webUrl ? `\nURL: ${response.webUrl}` : ""}`,
            },
          ],
        }
      } catch (error) {
        return {
          content: [{ type: "text", text: `Error creating linked resource: ${error}` }],
        }
      }
    },
  )

  server.tool(
    "get-linked-resource",
    "Get a single linked resource by ID for a Microsoft Todo task.",
    {
      listId: z.string().describe("ID of the task list"),
      taskId: z.string().describe("ID of the task"),
      linkedResourceId: z.string().describe("ID of the linked resource"),
    },
    async ({ listId, taskId, linkedResourceId }) => {
      try {
        const token = await getAccessToken()
        if (!token) {
          return {
            content: [{ type: "text", text: "Failed to authenticate with Microsoft API" }],
          }
        }

        const resource = await makeGraphRequest<LinkedResource>(
          `${MS_GRAPH_BASE}/me/todo/lists/${listId}/tasks/${taskId}/linkedResources/${linkedResourceId}`,
          token,
        )

        if (!resource) {
          return {
            content: [{ type: "text", text: `Failed to retrieve linked resource: ${linkedResourceId}` }],
          }
        }

        const details = [`ID: ${resource.id || "Unknown"}`, `Display Name: ${resource.displayName || "Unknown"}`]
        if (resource.applicationName) details.push(`Application: ${resource.applicationName}`)
        if (resource.webUrl) details.push(`URL: ${resource.webUrl}`)
        if (resource.externalId) details.push(`External ID: ${resource.externalId}`)

        return {
          content: [{ type: "text", text: `Linked resource details:\n\n${details.join("\n")}` }],
        }
      } catch (error) {
        return {
          content: [{ type: "text", text: `Error fetching linked resource: ${error}` }],
        }
      }
    },
  )

  server.tool(
    "update-linked-resource",
    "Update an existing linked resource on a Microsoft Todo task.",
    {
      listId: z.string().describe("ID of the task list"),
      taskId: z.string().describe("ID of the task"),
      linkedResourceId: z.string().describe("ID of the linked resource to update"),
      webUrl: z.string().optional().describe("New deep link URL"),
      applicationName: z.string().optional().describe("New source application name"),
      displayName: z.string().optional().describe("New display name"),
      externalId: z.string().optional().describe("New external identifier"),
    },
    async ({ listId, taskId, linkedResourceId, webUrl, applicationName, displayName, externalId }) => {
      try {
        const token = await getAccessToken()
        if (!token) {
          return {
            content: [{ type: "text", text: "Failed to authenticate with Microsoft API" }],
          }
        }

        const body: any = {}
        if (webUrl !== undefined) body.webUrl = webUrl
        if (applicationName !== undefined) body.applicationName = applicationName
        if (displayName !== undefined) body.displayName = displayName
        if (externalId !== undefined) body.externalId = externalId

        if (Object.keys(body).length === 0) {
          return {
            content: [{ type: "text", text: "No properties provided for update." }],
          }
        }

        const response = await makeGraphRequest<LinkedResource>(
          `${MS_GRAPH_BASE}/me/todo/lists/${listId}/tasks/${taskId}/linkedResources/${linkedResourceId}`,
          token,
          "PATCH",
          body,
        )

        if (!response) {
          return {
            content: [{ type: "text", text: `Failed to update linked resource: ${linkedResourceId}` }],
          }
        }

        return {
          content: [
            {
              type: "text",
              text:
                `Linked resource updated successfully!\nID: ${response.id || "Unknown"}` +
                `${response.displayName ? `\nDisplay Name: ${response.displayName}` : ""}` +
                `${response.webUrl ? `\nURL: ${response.webUrl}` : ""}`,
            },
          ],
        }
      } catch (error) {
        return {
          content: [{ type: "text", text: `Error updating linked resource: ${error}` }],
        }
      }
    },
  )

  server.tool(
    "delete-linked-resource",
    "Delete a linked resource from a Microsoft Todo task.",
    {
      listId: z.string().describe("ID of the task list"),
      taskId: z.string().describe("ID of the task"),
      linkedResourceId: z.string().describe("ID of the linked resource to delete"),
    },
    async ({ listId, taskId, linkedResourceId }) => {
      try {
        const token = await getAccessToken()
        if (!token) {
          return {
            content: [{ type: "text", text: "Failed to authenticate with Microsoft API" }],
          }
        }

        await makeGraphRequest<null>(
          `${MS_GRAPH_BASE}/me/todo/lists/${listId}/tasks/${taskId}/linkedResources/${linkedResourceId}`,
          token,
          "DELETE",
        )

        return {
          content: [{ type: "text", text: `Linked resource ${linkedResourceId} deleted from task: ${taskId}` }],
        }
      } catch (error) {
        return {
          content: [{ type: "text", text: `Error deleting linked resource: ${error}` }],
        }
      }
    },
  )
}
