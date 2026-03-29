#!/usr/bin/env node

import fs from "fs"
import path from "path"
import { homedir } from "os"

// Resolve the token file path (same logic as cli.ts and token-manager.ts)
function resolveTokenFilePath(): string {
  if (process.argv[2]) {
    return path.resolve(process.argv[2])
  }

  if (process.env.MSTODO_TOKEN_FILE) {
    return process.env.MSTODO_TOKEN_FILE
  }

  const configDir =
    process.platform === "win32"
      ? path.join(process.env.APPDATA || path.join(homedir(), "AppData", "Roaming"), "microsoft-todo-mcp")
      : path.join(homedir(), ".config", "microsoft-todo-mcp")

  const configPath = path.join(configDir, "tokens.json")
  if (fs.existsSync(configPath)) {
    return configPath
  }

  // Legacy fallback
  const cwdPath = path.join(process.cwd(), "tokens.json")
  if (fs.existsSync(cwdPath)) {
    return cwdPath
  }

  return configPath
}

const tokenPath = resolveTokenFilePath()
const outputPath = process.argv[3] || path.join(process.cwd(), "mcp.json")

console.log(`Token file: ${tokenPath}`)
console.log(`Writing config to: ${outputPath}`)

try {
  // Verify the token file exists and is readable
  if (!fs.existsSync(tokenPath)) {
    console.error(`Token file not found: ${tokenPath}`)
    console.error("Run 'pnpm run auth' or 'npx mstodo-setup' to authenticate first.")
    process.exit(1)
  }

  // Point the MCP config at the token file instead of embedding secrets
  const mcpConfig = {
    mcpServers: {
      microsoftTodo: {
        command: "npx",
        args: ["--yes", "microsoft-todo-mcp-server"],
        env: {
          MSTODO_TOKEN_FILE: tokenPath,
        },
      },
    },
  }

  // Write the config with restrictive permissions (owner read/write only)
  fs.writeFileSync(outputPath, JSON.stringify(mcpConfig, null, 2), {
    encoding: "utf8",
    mode: 0o600,
  })

  console.log("MCP configuration file created successfully!")
  console.log("You can now use the service with Claude or Cursor by referencing this mcp.json file.")
} catch (error) {
  const errorMessage = error instanceof Error ? error.message : String(error)
  console.error("Error creating MCP config:", errorMessage)
  process.exit(1)
}
