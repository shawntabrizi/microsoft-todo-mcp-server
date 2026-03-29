# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Common Development Commands

### Build and Development

```bash
pnpm install         # Install dependencies
pnpm run build       # Build with tsup to dist/ directory
pnpm run dev         # Build and run CLI in one command
```

### Authentication and Setup

```bash
pnpm run auth        # Start OAuth authentication server (default port 3000, configurable via AUTH_PORT)
pnpm run create-config # Generate mcp.json from tokens.json
```

### Running the Server

```bash
pnpm run cli         # Run MCP server via CLI wrapper
pnpm start           # Run MCP server directly
```

## Architecture Overview

This is a Model Context Protocol (MCP) server that enables AI assistants to interact with Microsoft To Do via the Microsoft Graph API. The codebase follows a modular architecture with four main components:

1. **MCP Server** (`src/todo-index.ts`): Core server implementing the MCP protocol with 33 tools for Microsoft To Do operations
2. **CLI Wrapper** (`src/cli.ts`): Executable entry point that handles token loading from environment or file
3. **Token Manager** (`src/token-manager.ts`): Token storage, refresh, JWT expiry decoding, and Claude config auto-update
4. **Auth Server** (`src/auth-server.ts`): Express server implementing OAuth 2.0 flow with MSAL
5. **Config Generator** (`src/create-mcp-config.ts`): Utility to create MCP configuration files

### Key Architectural Patterns

- **Token Management**: Tokens stored in platform-specific config dir with JWT expiry decoding and automatic refresh
- **Request Deduplication**: In-flight cache prevents duplicate API calls on POST/PATCH/DELETE
- **Multi-tenant Support**: Configurable for different Microsoft account types via TENANT_ID
- **Error Handling**: Special handling for personal Microsoft accounts (MailboxNotEnabledForRESTAPI), 401 auto-retry
- **Type Safety**: Strict TypeScript with Zod schemas for parameter validation

### Microsoft Graph API Integration

The server communicates with Microsoft Graph API v1.0:

- Base URL: `https://graph.microsoft.com/v1.0`
- Hierarchy: Lists → Tasks → Checklist Items, Attachments, Linked Resources
- Supports OData query parameters for filtering and sorting
- Delta queries for efficient change tracking

### Environment Configuration

- `MSTODO_TOKEN_FILE`: Custom path for tokens.json (defaults to ~/.config/microsoft-todo-mcp/tokens.json)
- `AUTH_PORT`: Custom port for auth server (defaults to 3000)
- `.env` file required for authentication with CLIENT_ID, CLIENT_SECRET, TENANT_ID, REDIRECT_URI

## Important Notes

- Always run `pnpm run build` after modifying TypeScript files (uses tsup for bundling to dist/)
- The auth server runs on port 3000 by default (configurable via AUTH_PORT)
- Tokens are automatically refreshed using the refresh token when needed
- Personal Microsoft accounts have limited API access compared to work/school accounts
- Node.js 18+ required (target set in tsup.config.ts)
