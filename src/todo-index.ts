import { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js"
import { StdioServerTransport } from "@modelcontextprotocol/sdk/server/stdio.js"
import dotenv from "dotenv"
import { tokenManager } from "./token-manager.js"
import { isPersonalMicrosoftAccount } from "./graph-client.js"
import type { ServerConfig } from "./types.js"

import { register as registerAuth } from "./tools/auth.js"
import { register as registerTaskLists } from "./tools/task-lists.js"
import { register as registerTasks } from "./tools/tasks.js"
import { register as registerChecklistItems } from "./tools/checklist-items.js"
import { register as registerLinkedResources } from "./tools/linked-resources.js"
import { register as registerAttachments } from "./tools/attachments.js"
import { register as registerConvenience } from "./tools/convenience.js"
import { register as registerBulk } from "./tools/bulk.js"

// Load environment variables
dotenv.config()

// Create server instance
const server = new McpServer({
  name: "mstodo",
  version: "2.0.0",
})

// Register all tools
registerAuth(server)
registerTaskLists(server)
registerTasks(server)
registerChecklistItems(server)
registerLinkedResources(server)
registerAttachments(server)
registerConvenience(server)
registerBulk(server)

// Main function to start the server
export async function startServer(config?: ServerConfig): Promise<void> {
  try {
    if (config?.tokenFilePath) {
      process.env.MSTODO_TOKEN_FILE = config.tokenFilePath
      tokenManager.configure({ tokenFilePath: config.tokenFilePath })
    }

    if (config?.accessToken) {
      process.env.MS_TODO_ACCESS_TOKEN = config.accessToken
    }

    if (config?.refreshToken) {
      process.env.MS_TODO_REFRESH_TOKEN = config.refreshToken
    }

    // Check if using a personal Microsoft account and show warning if needed
    await isPersonalMicrosoftAccount()

    // Start the server
    const transport = new StdioServerTransport()
    await server.connect(transport)

    console.error("Server started and listening")
  } catch (error) {
    console.error("Error starting server:", error)
    throw error
  }
}

// Main entry point when executed directly
if (import.meta.url === `file://${process.argv[1]}`) {
  startServer().catch((error) => {
    console.error("Fatal error in main():", error)
    process.exit(1)
  })
}
