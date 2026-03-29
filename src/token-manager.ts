// src/token-manager.ts
import { readFileSync, writeFileSync, existsSync, mkdirSync } from "fs"
import { dirname, join } from "path"
import { homedir } from "os"

interface TokenData {
  accessToken: string
  refreshToken: string
  expiresAt: number
}

interface StoredTokenData extends TokenData {
  clientId?: string
  clientSecret?: string
  tenantId?: string
}

function decodeJwtExpiry(accessToken: string): number | null {
  try {
    const parts = accessToken.split(".")
    if (parts.length !== 3) return null

    const payload = JSON.parse(Buffer.from(parts[1], "base64url").toString("utf8")) as { exp?: number }
    if (!payload.exp) return null

    return payload.exp * 1000
  } catch {
    return null
  }
}

function getDefaultConfigDir(): string {
  return process.platform === "win32"
    ? join(process.env.APPDATA || join(homedir(), "AppData", "Roaming"), "microsoft-todo-mcp")
    : join(homedir(), ".config", "microsoft-todo-mcp")
}

export class TokenManager {
  private tokenFilePath: string
  private currentTokens: StoredTokenData | null = null
  private configuredTokenFilePath?: string

  constructor() {
    this.tokenFilePath = this.resolveTokenFilePath()
  }

  private resolveTokenFilePath(): string {
    const configuredPath = this.configuredTokenFilePath || process.env.MSTODO_TOKEN_FILE
    if (configuredPath) {
      return configuredPath
    }

    const configDir = getDefaultConfigDir()
    if (!existsSync(configDir)) {
      mkdirSync(configDir, { recursive: true })
    }

    return join(configDir, "tokens.json")
  }

  configure(options?: { tokenFilePath?: string }): void {
    if (options?.tokenFilePath) {
      this.configuredTokenFilePath = options.tokenFilePath
    }

    const nextTokenFilePath = this.resolveTokenFilePath()
    this.tokenFilePath = nextTokenFilePath

    const tokenDir = dirname(nextTokenFilePath)
    if (!existsSync(tokenDir)) {
      mkdirSync(tokenDir, { recursive: true })
    }

    console.error(`Token file path: ${this.tokenFilePath}`)
  }

  getTokenFilePath(): string {
    this.configure()
    return this.tokenFilePath
  }

  private buildEnvTokens(): StoredTokenData | null {
    if (!process.env.MS_TODO_ACCESS_TOKEN || !process.env.MS_TODO_REFRESH_TOKEN) {
      return null
    }

    const expiresAt = decodeJwtExpiry(process.env.MS_TODO_ACCESS_TOKEN) ?? 0

    return {
      accessToken: process.env.MS_TODO_ACCESS_TOKEN,
      refreshToken: process.env.MS_TODO_REFRESH_TOKEN,
      expiresAt,
      clientId: process.env.CLIENT_ID,
      clientSecret: process.env.CLIENT_SECRET,
      tenantId: process.env.TENANT_ID,
    }
  }

  private readStoredTokens(pathToRead: string): StoredTokenData | null {
    if (!existsSync(pathToRead)) {
      return null
    }

    try {
      const data = readFileSync(pathToRead, "utf8")
      return JSON.parse(data) as StoredTokenData
    } catch (error) {
      console.error(`Error reading token file ${pathToRead}:`, error)
      return null
    }
  }

  // Try to get tokens from multiple sources
  async getTokens(options?: { forceRefresh?: boolean }): Promise<TokenData | null> {
    this.configure()
    const forceRefresh = options?.forceRefresh === true

    // 1. Prefer the configured token file when using repo-local/file-based auth.
    const fileTokens = this.readStoredTokens(this.tokenFilePath)
    if (fileTokens) {
      this.currentTokens = fileTokens

      if (forceRefresh || Date.now() > fileTokens.expiresAt) {
        const refreshed = await this.refreshToken(fileTokens.refreshToken)
        if (refreshed) {
          return refreshed
        }
      }

      if (!forceRefresh) {
        return fileTokens
      }
    }

    // 2. Check environment variables for backward compatibility.
    const envTokens = this.buildEnvTokens()
    if (envTokens) {
      this.currentTokens = envTokens

      if (forceRefresh || Date.now() > envTokens.expiresAt) {
        const refreshed = await this.refreshToken(envTokens.refreshToken)
        if (refreshed) {
          return refreshed
        }
      }

      if (!forceRefresh && envTokens.expiresAt > Date.now()) {
        return envTokens
      }
    }

    // 3. Check legacy token file location
    const legacyPath = join(process.cwd(), "tokens.json")
    const legacyTokens = this.readStoredTokens(legacyPath)
    if (legacyTokens) {
      this.currentTokens = legacyTokens

      if (forceRefresh || Date.now() > legacyTokens.expiresAt) {
        const refreshed = await this.refreshToken(legacyTokens.refreshToken)
        if (refreshed) {
          return refreshed
        }
      }

      this.saveTokens(legacyTokens)

      if (!forceRefresh) {
        return legacyTokens
      }
    }

    return null
  }

  async refreshToken(refreshToken: string): Promise<TokenData | null> {
    try {
      // Get client credentials from stored tokens or environment
      const clientId = this.currentTokens?.clientId || process.env.CLIENT_ID
      const clientSecret = this.currentTokens?.clientSecret || process.env.CLIENT_SECRET
      const tenantId = this.currentTokens?.tenantId || process.env.TENANT_ID || "organizations"

      if (!clientId || !clientSecret) {
        console.error("Missing client credentials for token refresh")
        return null
      }

      const tokenEndpoint = `https://login.microsoftonline.com/${tenantId}/oauth2/v2.0/token`

      const formData = new URLSearchParams({
        client_id: clientId,
        client_secret: clientSecret,
        refresh_token: refreshToken,
        grant_type: "refresh_token",
        scope: "offline_access Tasks.Read Tasks.ReadWrite Tasks.Read.Shared Tasks.ReadWrite.Shared User.Read",
      })

      const response = await fetch(tokenEndpoint, {
        method: "POST",
        headers: {
          "Content-Type": "application/x-www-form-urlencoded",
        },
        body: formData,
      })

      if (!response.ok) {
        const errorText = await response.text()
        console.error(`Token refresh failed: ${errorText}`)

        // If refresh fails, prompt for re-authentication
        this.promptForReauth()
        return null
      }

      const data = await response.json()

      const newTokens: StoredTokenData = {
        accessToken: data.access_token,
        refreshToken: data.refresh_token || refreshToken,
        expiresAt: Date.now() + data.expires_in * 1000 - 5 * 60 * 1000, // 5 min buffer
        clientId,
        clientSecret,
        tenantId,
      }

      // Save the refreshed tokens
      this.saveTokens(newTokens)

      // Also update Claude config if possible
      await this.updateClaudeConfig(newTokens)

      return newTokens
    } catch (error) {
      console.error("Error refreshing token:", error)
      this.promptForReauth()
      return null
    }
  }

  saveTokens(tokens: StoredTokenData): void {
    this.currentTokens = tokens
    writeFileSync(this.tokenFilePath, JSON.stringify(tokens, null, 2), "utf8")
  }

  // Update Claude config automatically
  async updateClaudeConfig(_tokens: TokenData): Promise<void> {
    try {
      const claudeConfigPath =
        process.platform === "win32"
          ? join(process.env.APPDATA || "", "Claude", "claude_desktop_config.json")
          : process.platform === "darwin"
            ? join(homedir(), "Library", "Application Support", "Claude", "claude_desktop_config.json")
            : join(homedir(), ".config", "Claude", "claude_desktop_config.json")

      if (!existsSync(claudeConfigPath)) {
        return
      }

      const config = JSON.parse(readFileSync(claudeConfigPath, "utf8"))

      if (config.mcpServers) {
        for (const serverName of ["microsoftTodo", "microsoft-todo"]) {
          if (config.mcpServers[serverName]) {
            const serverEnv = {
              ...(config.mcpServers[serverName].env || {}),
              MSTODO_TOKEN_FILE: this.tokenFilePath,
            }

            delete serverEnv.MS_TODO_ACCESS_TOKEN
            delete serverEnv.MS_TODO_REFRESH_TOKEN

            config.mcpServers[serverName].env = serverEnv
          }
        }

        writeFileSync(claudeConfigPath, JSON.stringify(config, null, 2), "utf8")
        console.error("Updated Claude config with new tokens")
      }
    } catch (error) {
      console.error("Could not update Claude config:", error)
    }
  }

  promptForReauth(): void {
    console.error(`
=================================================================
TOKEN REFRESH FAILED - REAUTHENTICATION REQUIRED

Your Microsoft To Do tokens have expired and could not be refreshed.

To fix this:
1. Open a new terminal
2. Navigate to the microsoft-todo-mcp-server directory
3. Run: pnpm run setup
4. Complete the authentication in your browser
5. Restart Claude Desktop to use the new tokens

Your tokens are stored in: ${this.tokenFilePath}
=================================================================
    `)
  }

  // Store client credentials with tokens for future refreshes
  async storeCredentials(clientId: string, clientSecret: string, tenantId: string): Promise<void> {
    if (this.currentTokens) {
      this.currentTokens.clientId = clientId
      this.currentTokens.clientSecret = clientSecret
      this.currentTokens.tenantId = tenantId
      this.saveTokens(this.currentTokens)
    }
  }
}

export const tokenManager = new TokenManager()
