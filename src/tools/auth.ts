import type { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js"
import { tokenManager } from "../token-manager.js"
import { isPersonalMicrosoftAccount } from "../graph-client.js"

export function register(server: McpServer) {
  // Auth status
  server.tool(
    "auth-status",
    "Check if you're authenticated with Microsoft Graph API. Shows current token status and expiration time.",
    {},
    async () => {
      const tokens = await tokenManager.getTokens()

      if (!tokens) {
        return {
          content: [
            {
              type: "text",
              text: "Not authenticated. Please run 'npx mstodo-setup' or 'pnpm run setup' to authenticate with Microsoft.",
            },
          ],
        }
      }

      const isExpired = Date.now() > tokens.expiresAt
      const expiryTime = new Date(tokens.expiresAt).toLocaleString()

      const isPersonal = await isPersonalMicrosoftAccount()
      let accountMessage = ""

      if (isPersonal) {
        accountMessage =
          "\n\nWARNING: You are using a personal Microsoft account. " +
          "Microsoft To Do API access is typically not available for personal accounts " +
          "through the Microsoft Graph API. You may encounter 'MailboxNotEnabledForRESTAPI' errors."
      }

      if (isExpired) {
        return {
          content: [
            {
              type: "text",
              text: `Authentication expired at ${expiryTime}. Will attempt to refresh when you call any API.${accountMessage}`,
            },
          ],
        }
      } else {
        return {
          content: [
            {
              type: "text",
              text: `Authenticated. Token expires at ${expiryTime}.${accountMessage}`,
            },
          ],
        }
      }
    },
  )

  server.tool(
    "refresh-auth-token",
    "Force a Microsoft Graph token refresh using the stored refresh token and report the new expiration time.",
    {},
    async () => {
      const previousTokens = await tokenManager.getTokens()

      if (!previousTokens) {
        return {
          content: [
            {
              type: "text",
              text: "Not authenticated. Please run 'npx mstodo-setup' or 'pnpm run setup' to authenticate with Microsoft.",
            },
          ],
        }
      }

      const refreshedTokens = await tokenManager.getTokens({ forceRefresh: true })

      if (!refreshedTokens) {
        return {
          content: [
            {
              type: "text",
              text:
                "Failed to refresh the Microsoft Graph token. Reauthentication may be required." +
                `\nToken file: ${tokenManager.getTokenFilePath()}`,
            },
          ],
        }
      }

      const refreshedExpiryTime = new Date(refreshedTokens.expiresAt).toLocaleString()
      const previousExpiryTime = new Date(previousTokens.expiresAt).toLocaleString()
      const didExpiryChange = refreshedTokens.expiresAt !== previousTokens.expiresAt

      return {
        content: [
          {
            type: "text",
            text: didExpiryChange
              ? `Authentication refreshed successfully. Previous expiry: ${previousExpiryTime}. New expiry: ${refreshedExpiryTime}.\nToken file: ${tokenManager.getTokenFilePath()}`
              : `Authentication is already current. Current expiry: ${refreshedExpiryTime}.\nToken file: ${tokenManager.getTokenFilePath()}`,
          },
        ],
      }
    },
  )
}
