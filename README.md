# Microsoft To Do MCP Server

A [Model Context Protocol](https://modelcontextprotocol.io/) (MCP) server that enables AI assistants like Claude to interact with Microsoft To Do via the Microsoft Graph API. Provides **30 tools** for comprehensive task management through a secure OAuth 2.0 authentication flow.

## About This Fork

This is a community-maintained fork of [jordanburke/microsoft-todo-mcp-server](https://github.com/jordanburke/microsoft-todo-mcp-server). The original project has gone a bit stale with several open PRs and community contributions that haven't been merged. This fork consolidates fixes and features from across the community so they're available in one place.

If you've contributed to the original repo or a fork and want your work included here, open an issue or PR.

### Community Contributions

The following fixes and features were pulled from open PRs and community forks:

| Contribution | Author | Source |
|---|---|---|
| Configurable auth server port via `AUTH_PORT` env var | [@vavdb](https://github.com/vavdb) | [PR #4](https://github.com/jordanburke/microsoft-todo-mcp-server/pull/4) |
| Fix ESM dynamic `require('fs')` with static import | [@jleaders](https://github.com/jleaders) | [PR #3](https://github.com/jordanburke/microsoft-todo-mcp-server/pull/3) |
| Full task metadata: timestamps, recurrence, attachments | [@ThePlasmak](https://github.com/ThePlasmak) | [PR #2](https://github.com/jordanburke/microsoft-todo-mcp-server/pull/2) |
| Request deduplication preventing duplicate task creation | [@Mcp20091](https://github.com/Mcp20091) | [Fork](https://github.com/Mcp20091/microsoft-todo-mcp-server) |
| DELETE 204 No Content handling fix | [@Mcp20091](https://github.com/Mcp20091) | [Fork](https://github.com/Mcp20091/microsoft-todo-mcp-server) |
| Improved token manager (JWT expiry, forceRefresh, configure) | [@Mcp20091](https://github.com/Mcp20091) | [Fork](https://github.com/Mcp20091/microsoft-todo-mcp-server) |
| Full task display without description truncation | [@Mcp20091](https://github.com/Mcp20091) | [Fork](https://github.com/Mcp20091/microsoft-todo-mcp-server) |
| Proper recurrence handling for recurring task date updates | [@Mcp20091](https://github.com/Mcp20091) | [Fork](https://github.com/Mcp20091/microsoft-todo-mcp-server) |
| `auth-status` and `refresh-auth-token` tools | [@Mcp20091](https://github.com/Mcp20091) | [Fork](https://github.com/Mcp20091/microsoft-todo-mcp-server) |
| Single-item getters: `get-task`, `get-task-list` | [@Mcp20091](https://github.com/Mcp20091) | [Fork](https://github.com/Mcp20091/microsoft-todo-mcp-server) |
| Delta query tools: `get-tasks-delta`, `get-task-lists-delta` | [@Mcp20091](https://github.com/Mcp20091) | [Fork](https://github.com/Mcp20091/microsoft-todo-mcp-server) |
| `skip-task-to-current` recurring task tool | [@Mcp20091](https://github.com/Mcp20091) | [Fork](https://github.com/Mcp20091/microsoft-todo-mcp-server) |
| Linked resource tools: `get-linked-resources`, `create-linked-resource` | [@Mcp20091](https://github.com/Mcp20091) | [Fork](https://github.com/Mcp20091/microsoft-todo-mcp-server) |
| Attachment tools: `get-attachments`, `get-attachment`, `create-attachment`, `delete-attachment`, `create-attachment-upload-session` | [@Mcp20091](https://github.com/Mcp20091) | [Fork](https://github.com/Mcp20091/microsoft-todo-mcp-server) |
| `archive-completed-tasks` bulk operation | [@Mcp20091](https://github.com/Mcp20091) | [Fork](https://github.com/Mcp20091/microsoft-todo-mcp-server) |
| Enhanced `create-task` / `update-task` with recurrence, linkedResources, categories | [@Mcp20091](https://github.com/Mcp20091) | [Fork](https://github.com/Mcp20091/microsoft-todo-mcp-server) |
| `move-task` tool (cross-list with metadata preservation) | [@steven-pribilinskiy](https://github.com/steven-pribilinskiy) | [Fork](https://github.com/steven-pribilinskiy/microsoft-todo-mcp-server) |
| `reorganize-list` tool (bulk task restructuring with dry-run) | [@commit21](https://github.com/commit21) | [Fork](https://github.com/commit21/microsoft-todo-mcp-server) |
| tsup target bump to `node18` with explicit `platform: 'node'` | [@jleaders](https://github.com/jleaders) | [PR #3](https://github.com/jordanburke/microsoft-todo-mcp-server/pull/3) |

## Features

- **30 MCP Tools** covering lists, tasks, checklist items, attachments, linked resources, and bulk operations
- **Automatic Token Refresh** with JWT expiry decoding and 5-minute buffer
- **Request Deduplication** prevents duplicate task creation when tools are double-invoked
- **OAuth 2.0 Authentication** via MSAL with configurable tenant support
- **Delta Queries** for efficient change tracking
- **Recurring Task Support** including skip-to-current and proper recurrence PATCH handling
- **Full Metadata** display including timestamps, recurrence patterns, categories, and linked resources
- **TypeScript + ESM** with strict typing and Zod schema validation

## Prerequisites

- Node.js 18 or higher
- pnpm package manager
- A Microsoft account (personal, work, or school)
- Azure App Registration (see [Setup](#azure-app-registration))

## Installation

### Option 1: npm / npx

```bash
# Run directly (no install)
npx microsoft-todo-mcp-server

# Or install globally
npm install -g microsoft-todo-mcp-server
```

### Option 2: Clone and Build

```bash
git clone https://github.com/shawntabrizi/microsoft-todo-mcp-server.git
cd microsoft-todo-mcp-server
pnpm install
pnpm run build
```

## Azure App Registration

1. Go to the [Azure Portal](https://portal.azure.com)
2. Navigate to **App registrations** and create a new registration
3. Name your application (e.g., "To Do MCP")
4. For **Supported account types**, choose based on your needs:
   - **Single tenant** - One organization only
   - **Multitenant** - Any Azure AD directory
   - **Multitenant + personal** - Work/school and personal Microsoft accounts
5. Set the **Redirect URI** to `http://localhost:3000/callback` (or use `AUTH_PORT` env var for a different port)
6. Under **Certificates & secrets**, create a new client secret
7. Under **API permissions**, add these Microsoft Graph delegated permissions:
   - `Tasks.Read`
   - `Tasks.ReadWrite`
   - `User.Read`
8. Click **Grant admin consent**

## Configuration

### Environment Variables

Create a `.env` file in the project root:

```env
CLIENT_ID=your_client_id
CLIENT_SECRET=your_client_secret
TENANT_ID=your_tenant_setting
REDIRECT_URI=http://localhost:3000/callback
```

**`TENANT_ID` options:**

| Value | Use Case |
|---|---|
| `organizations` | Multi-tenant work/school accounts (default) |
| `consumers` | Personal Microsoft accounts only |
| `common` | Both work/school and personal accounts |
| `<your-tenant-id>` | Single-tenant / specific organization |

### Token Storage

Tokens are stored in `~/.config/microsoft-todo-mcp/tokens.json` (Linux/macOS) or `%APPDATA%\microsoft-todo-mcp\tokens.json` (Windows) with automatic refresh.

Override with environment variables:

```bash
# Custom token file path
export MSTODO_TOKEN_FILE=/path/to/tokens.json

# Or pass tokens directly
export MS_TODO_ACCESS_TOKEN=your_access_token
export MS_TODO_REFRESH_TOKEN=your_refresh_token
```

### Auth Server Port

The authentication server defaults to port 3000. To use a different port:

```env
AUTH_PORT=3002
REDIRECT_URI=http://localhost:3002/callback
```

## Usage

### Step 1: Authenticate

```bash
pnpm run auth
# Or: pnpm run setup
```

This opens a browser window for Microsoft authentication and saves tokens locally.

### Step 2: Configure Your AI Assistant

**Claude Desktop** (`~/Library/Application Support/Claude/claude_desktop_config.json` on macOS):

```json
{
  "mcpServers": {
    "microsoftTodo": {
      "command": "npx",
      "args": ["--yes", "microsoft-todo-mcp-server"],
      "env": {
        "MSTODO_TOKEN_FILE": "/path/to/tokens.json"
      }
    }
  }
}
```

**Cursor:**

```bash
cp mcp.json ~/.cursor/mcp-servers.json
```

### Development Scripts

```bash
pnpm run build         # Build TypeScript to dist/
pnpm run dev           # Build and run in one command
pnpm start             # Run MCP server directly
pnpm run cli           # Run via CLI wrapper
pnpm run auth          # Start OAuth authentication server
pnpm run create-config # Generate mcp.json from tokens
pnpm run format        # Format code with Prettier
```

## MCP Tools

### Authentication (2 tools)

| Tool | Description |
|---|---|
| `auth-status` | Check authentication status, token expiration, and account type |
| `refresh-auth-token` | Force a token refresh and report the new expiration time |

### Task Lists (6 tools)

| Tool | Description |
|---|---|
| `get-task-lists` | Get all task lists with metadata (default, shared, etc.) |
| `get-task-list` | Get a single task list by ID |
| `get-task-lists-delta` | Track changes to task lists via delta queries |
| `create-task-list` | Create a new task list |
| `update-task-list` | Rename an existing task list |
| `delete-task-list` | Delete a task list and all its contents |

### Tasks (8 tools)

| Tool | Description |
|---|---|
| `get-tasks` | Get tasks with OData filtering, sorting, and pagination (`$filter`, `$select`, `$orderby`, `$top`, `$skip`, `$count`) |
| `get-task` | Get a single task by ID |
| `get-tasks-delta` | Track changes to tasks via delta queries |
| `create-task` | Create a task with title, body, due date, start date, importance, reminders, recurrence, status, categories, and linked resources |
| `update-task` | Update any task properties, with proper handling for recurring task date adjustments |
| `delete-task` | Delete a task and all its checklist items |
| `move-task` | Move a task between lists, preserving checklist items and metadata |
| `skip-task-to-current` | Advance a recurring task to the next occurrence on or after today |

### Checklist Items / Subtasks (4 tools)

| Tool | Description |
|---|---|
| `get-checklist-items` | Get subtasks for a specific task |
| `create-checklist-item` | Add a subtask with optional checked state and timestamps |
| `update-checklist-item` | Update subtask text, completion status, or timestamps |
| `delete-checklist-item` | Remove a specific subtask |

### Attachments (5 tools)

| Tool | Description |
|---|---|
| `get-attachments` | List file attachments on a task |
| `get-attachment` | Get a single attachment by ID |
| `create-attachment` | Attach a small file (base64-encoded, under 3 MB) |
| `create-attachment-upload-session` | Create an upload session for large files |
| `delete-attachment` | Remove a file attachment |

### Linked Resources (2 tools)

| Tool | Description |
|---|---|
| `get-linked-resources` | Get linked resources for a task |
| `create-linked-resource` | Link an external resource (URL, app, external ID) to a task |

### Bulk Operations (3 tools)

| Tool | Description |
|---|---|
| `archive-completed-tasks` | Move completed tasks older than N days to an archive list (supports dry-run) |
| `reorganize-list` | Restructure flat tasks into category tasks with checklist items (supports dry-run and idempotency checks) |

## Architecture

```
src/
  todo-index.ts        # Core MCP server with all 30 tools
  cli.ts               # CLI entry point with token loading
  token-manager.ts     # Token storage, refresh, and JWT decoding
  auth-server.ts       # Express OAuth 2.0 server
  create-mcp-config.ts # MCP config file generator
  setup.ts             # Interactive setup wizard
```

**Key design decisions:**
- **Request deduplication**: In-flight cache prevents duplicate API calls when MCP tools are double-invoked
- **401 auto-retry**: Transparent token refresh on authentication failures
- **Recurring task handling**: Temporarily clears recurrence before date updates to work around Graph API limitations
- **Delta queries**: Efficient change tracking without re-fetching entire lists

## Limitations

### Personal Microsoft Accounts

Personal accounts (outlook.com, hotmail.com, live.com) may receive `MailboxNotEnabledForRESTAPI` errors. This is a [Microsoft Graph API limitation](https://learn.microsoft.com/en-us/graph/api/resources/todo-overview), not an issue with this server. Work/school accounts have full access.

### API Constraints

- Microsoft rate limits apply
- `move-task` uses copy+delete (no native move endpoint); tasks with attachments cannot be moved
- Some shared list operations may be restricted

## Troubleshooting

**Authentication failures**: Verify `CLIENT_ID`, `CLIENT_SECRET`, and `TENANT_ID` in `.env`. Ensure the redirect URI matches exactly (including port).

**Token expiry**: Use the `auth-status` tool to check, or `refresh-auth-token` to force a refresh. If refresh fails, re-run `pnpm run auth`.

**Personal account errors**: Set `TENANT_ID=consumers` or `TENANT_ID=common` in your `.env` file.

**Port conflicts**: Set `AUTH_PORT=3002` (or any open port) in `.env` and update `REDIRECT_URI` to match.

**Debug logs**: The server logs to stderr. Capture with `mstodo 2> debug.log`.

## License

MIT License - See [LICENSE](LICENSE) file for details.

## Acknowledgments

- Originally forked from [jordanburke/microsoft-todo-mcp-server](https://github.com/jordanburke/microsoft-todo-mcp-server) (itself a fork of [@jhirono/todomcp](https://github.com/jhirono/todomcp))
- Community contributors: [@Mcp20091](https://github.com/Mcp20091), [@steven-pribilinskiy](https://github.com/steven-pribilinskiy), [@commit21](https://github.com/commit21), [@vavdb](https://github.com/vavdb), [@jleaders](https://github.com/jleaders), [@ThePlasmak](https://github.com/ThePlasmak)
- Built on the [Model Context Protocol SDK](https://github.com/modelcontextprotocol/sdk)
- Uses [Microsoft Graph API](https://developer.microsoft.com/en-us/graph)
