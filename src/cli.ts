#!/usr/bin/env node

import { startServer } from "./todo-index.js"
import fs from "fs"
import path from "path"
import { homedir } from "os"

// Resolve token file path: env var > platform config dir > cwd fallback
function resolveTokenFilePath(): string {
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

const TOKEN_FILE_PATH = resolveTokenFilePath()

// Check for tokens in environment variables
let accessToken = process.env.MS_TODO_ACCESS_TOKEN
let refreshToken = process.env.MS_TODO_REFRESH_TOKEN

console.error("Microsoft Todo MCP CLI")
console.error(`Token file: ${TOKEN_FILE_PATH}`)

// Check if tokens are missing from environment but available in file
if ((!accessToken || !refreshToken) && fs.existsSync(TOKEN_FILE_PATH)) {
  try {
    const tokenData = JSON.parse(fs.readFileSync(TOKEN_FILE_PATH, "utf8"))

    if (!accessToken && tokenData.accessToken) {
      accessToken = tokenData.accessToken
    }

    if (!refreshToken && tokenData.refreshToken) {
      refreshToken = tokenData.refreshToken
    }
  } catch (error) {
    const errorMessage = error instanceof Error ? error.message : String(error)
    console.error("Error reading token file:", errorMessage)
  }
}

// Start the MCP server with the available tokens
startServer({
  accessToken,
  refreshToken,
  tokenFilePath: TOKEN_FILE_PATH,
}).catch((error) => {
  const errorMessage = error instanceof Error ? error.message : String(error)
  console.error("Error starting server:", errorMessage)
  process.exit(1)
})
