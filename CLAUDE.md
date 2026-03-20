# connector-graph

MCP server exposing 44 Microsoft Graph API tools for user management, groups, licenses, mail, OneDrive, calendar, meeting transcripts, SharePoint, Planner, To Do, and Teams Chat.

## Tech Stack

- **Runtime**: Node.js (ES2022, ES modules)
- **Language**: TypeScript 5.7+
- **Protocol**: MCP SDK (stdio + StreamableHTTP transport)
- **Auth**: Dual — delegated (user Bearer token) + app-only (Service Principal)
- **API**: Microsoft Graph REST API v1.0

## Key Commands

```bash
npm run build     # Compile TypeScript -> dist/
npm start         # Start MCP server (stdio)
npm run dev       # Watch mode (tsc --watch)
PORT=8030 npm start  # Start as HTTP server (StreamableHTTP)
```

## Architecture

```
src/
├── index.ts              # MCP server entry point + graph_auth_status tool + dual transport (stdio/HTTP) + dual auth (delegated/SP)
├── api/
│   ├── client.ts         # GraphClient — dual auth (user token → SP fallback), token cache, 429 retry, OData pagination, getText, If-Match
│   └── types.ts          # All TypeScript interfaces
├── tools/
│   ├── users.ts          # graph_list_users, graph_get_user, graph_update_user
│   ├── groups.ts         # graph_list_groups, graph_get_group_members, graph_add_group_member, graph_remove_group_member
│   ├── licenses.ts       # graph_list_subscribed_skus, graph_list_user_licenses
│   ├── mail.ts           # graph_send_mail, graph_search_mail, graph_read_mail, graph_list_attachments, graph_get_attachment, graph_read_attachment
│   ├── onedrive.ts       # graph_list_drive_items, graph_get_drive_item_content
│   ├── calendar.ts       # graph_list_events, graph_get_online_meeting, graph_list_meeting_transcripts, graph_get_meeting_transcript_content
│   ├── sharepoint.ts     # graph_list_sites, graph_get_site, graph_list_site_drives, graph_list_site_drive_items, graph_get_site_file_content, graph_search_site_files
│   ├── planner.ts        # graph_list_plans, graph_get_plan, graph_list_buckets, graph_list_plan_tasks, graph_get_task, graph_create_task, graph_update_task, graph_delete_task
│   ├── todo.ts           # graph_list_todo_lists, graph_list_todo_tasks, graph_create_todo_task, graph_update_todo_task, graph_complete_todo_task
│   └── teams.ts          # graph_list_chats, graph_list_chat_messages, graph_send_chat_message
├── types/
│   └── mammoth.d.ts      # Type declarations for mammoth (no bundled types)
├── utils/
│   ├── auth.ts           # AsyncLocalStorage for per-request user tokens (delegated auth context)
│   ├── config.ts         # GRAPH_TENANT_ID, CLIENT_ID, CLIENT_SECRET, API_KEY validation
│   ├── content-extractor.ts  # Attachment content extraction (PDF, Word, Excel, HTML, CSV, EML, images)
│   └── logger.ts         # Stderr logger
get-user-token.py          # MSAL device-code flow — acquires delegated Graph token for users
start-mcp.cmd              # Wrapper script — gets token + launches mcp-remote for Claude Code
.github/
└── workflows/
    └── deploy.yml        # CI/CD -> ACR -> Container Apps
```

## MCP Tools (44 total)

### Auth (1)
| Tool | Description |
|------|-------------|
| `graph_auth_status` | Verify auth — reports auth mode (delegated/app-only), user identity or org info |

### Users (3)
| Tool | Description |
|------|-------------|
| `graph_list_users` | OData filter/select/top/orderby + pagination |
| `graph_get_user` | By ID or UPN, includes signInActivity |
| `graph_update_user` | Cloud-only properties (jobTitle, accountEnabled, etc.) |

### Groups (4)
| Tool | Description |
|------|-------------|
| `graph_list_groups` | Filter by mailEnabled, securityEnabled, displayName |
| `graph_get_group_members` | List group members |
| `graph_add_group_member` | Add user to group |
| `graph_remove_group_member` | Remove user from group |

### Licenses (2)
| Tool | Description |
|------|-------------|
| `graph_list_subscribed_skus` | License inventory (consumed/prepaid/available) |
| `graph_list_user_licenses` | Licenses assigned to a specific user |

### Mail (6)
| Tool | Description |
|------|-------------|
| `graph_send_mail` | Send email from a user's mailbox (HTML/text, to/cc, importance) |
| `graph_search_mail` | Search mailbox with $search/$filter |
| `graph_read_mail` | Get full email content by message ID |
| `graph_list_attachments` | List attachments on an email |
| `graph_get_attachment` | Download attachment content (decoded text or base64 binary) |
| `graph_read_attachment` | Download attachment and extract readable text (PDF, Word, Excel, HTML, CSV, EML, images, plain text) |

### OneDrive (2)
| Tool | Description |
|------|-------------|
| `graph_list_drive_items` | List OneDrive files/folders at a path |
| `graph_get_drive_item_content` | Get file metadata and download URL |

### Calendar & Meetings (4)
| Tool | Description |
|------|-------------|
| `graph_list_events` | Calendar events in a date range with attendees and online meeting info |
| `graph_get_online_meeting` | Get online meeting details by join URL (needed for transcript retrieval) |
| `graph_list_meeting_transcripts` | List available transcripts for an online meeting |
| `graph_get_meeting_transcript_content` | Get full meeting transcript text (VTT format with speakers and timestamps) |

### SharePoint (6)
| Tool | Description |
|------|-------------|
| `graph_list_sites` | Search SharePoint sites by keyword |
| `graph_get_site` | Get site details by ID |
| `graph_list_site_drives` | List document libraries in a site |
| `graph_list_site_drive_items` | List files/folders in a site's document library |
| `graph_get_site_file_content` | Get file metadata and download URL from a site |
| `graph_search_site_files` | Search files within a site by keyword |

### Planner (8)
| Tool | Description |
|------|-------------|
| `graph_list_plans` | List plans for an M365 group |
| `graph_get_plan` | Get plan details by ID |
| `graph_list_buckets` | List buckets (columns) in a plan |
| `graph_list_plan_tasks` | List all tasks in a plan with status and assignees |
| `graph_get_task` | Get full task details including etag |
| `graph_create_task` | Create a task with optional bucket, assignees, due date |
| `graph_update_task` | Update task properties (requires etag for If-Match) |
| `graph_delete_task` | Delete a task (requires etag for If-Match) |

### To Do (5)
| Tool | Description |
|------|-------------|
| `graph_list_todo_lists` | List a user's To Do lists |
| `graph_list_todo_tasks` | List tasks in a To Do list |
| `graph_create_todo_task` | Create a To Do task |
| `graph_update_todo_task` | Update a To Do task (title, status, due date, body) |
| `graph_complete_todo_task` | Mark a To Do task as completed |

### Teams Chat (3)
| Tool | Description |
|------|-------------|
| `graph_list_chats` | List a user's Teams chats (1:1, group, meeting) |
| `graph_list_chat_messages` | List messages in a chat with pagination |
| `graph_send_chat_message` | Send a message to a Teams chat |

## Configuration

Required env vars:
- `GRAPH_TENANT_ID` — Azure AD tenant ID
- `GRAPH_CLIENT_ID` — Entra app client ID (SP: `Majans-Graph-MCP-Agent`)
- `GRAPH_CLIENT_SECRET` — Entra app client secret
- `GRAPH_API_KEY` — (optional) API key for agent access when no Bearer token. If set, requests without Bearer token must provide `X-API-Key` header.
- `PORT` — (optional) Set to enable StreamableHTTP transport (e.g. `8030`). Omit for stdio.

Use 1Password: `op run --env-file=.env.template -- npm start`

## Dual Auth Architecture

Two auth paths coexist on the same server. Tool functions are identical — auth switching is transparent.

```
Majans Users (Claude Code)              Agents (Container Apps)
        │                                    │
        │ MSAL device-code flow              │ Service Principal
        │ (get-user-token.py)                │ (client credentials)
        ▼                                    ▼
    /mcp endpoint                        /mcp endpoint
    Authorization: Bearer <user_token>   X-API-Key header (no Bearer)
    Graph sees the user's identity       Graph sees SP identity
    User only accesses own data          SP has full tenant access
```

### How It Works

1. **HTTP request arrives** at `/mcp`
2. `index.ts` checks for `Authorization: Bearer` header
3. If present → stores user token in `AsyncLocalStorage` (per-request context)
4. If absent → validates `X-API-Key` header (agent path)
5. `GraphClient.getToken()` checks `AsyncLocalStorage` first, falls back to SP token
6. Graph API enforces permissions based on whichever token is used

### Auth Mode Detection

`graph_auth_status` tool reports which mode is active:
- **Delegated**: returns `auth_mode: "delegated"`, user name, email
- **App-only**: returns `auth_mode: "app-only"`, org name, client ID

### Per-User Permissions (Delegated)

With delegated auth, Graph API enforces the user's own permissions:
- **Mail**: user can only read/send their own email
- **Calendar**: user can only see their own events
- **OneDrive**: user can only access their own files
- **To Do**: user can only see their own task lists
- **Teams**: user can only see chats they're part of
- **SharePoint**: user can only access sites they have permission to
- **Planner**: user can only see plans for groups they belong to
- **Users/Groups**: limited by the user's directory role (most users get read-only basic profiles)

### User Setup (Claude Code)

1. Set `GRAPH_MCP_CLIENT_ID` env var to the `Graph-MCP-User` Entra app client ID
2. Register in Claude Code MCP config:
   ```json
   {
     "graph": {
       "type": "stdio",
       "command": "cmd",
       "args": ["/c", "C:\\path\\to\\connector-graph\\start-mcp.cmd"]
     }
   }
   ```
3. First run: browser opens for device-code login (Entra ID)
4. Subsequent runs: token refreshes silently from cache (`~/.connector-graph/token_cache.bin`)

Prerequisites: Python 3.10+ (`pip install msal`), Node.js (`npx mcp-remote`)

## Transport Modes

### stdio (default)
Used when `PORT` is not set. Standard MCP stdio transport for local Claude Desktop / Claude Code usage. Uses SP auth (no delegated path in stdio mode).

### StreamableHTTP
Used when `PORT` is set. Starts an HTTP server with:
- `POST /mcp` — MCP StreamableHTTP endpoint (supports Bearer token for delegated auth)
- `GET /health` — Health check for container probes

### Docker Deployment
```bash
npm run build
docker build -t connector-graph .
docker run -p 8030:8030 \
  -e GRAPH_TENANT_ID=... \
  -e GRAPH_CLIENT_ID=... \
  -e GRAPH_CLIENT_SECRET=... \
  -e GRAPH_API_KEY=... \
  -e PORT=8030 \
  connector-graph
```

## Deployment

**Live** on Azure Container Apps.

| Resource | Value |
|----------|-------|
| Resource Group | `rg-majans-agents` |
| Container Registry | `acrmajansagents.azurecr.io` |
| Environment | `cae-majans-agents` (Australia East) |
| Container App | `connector-graph` (external ingress, port 8030) |
| CI/CD | `.github/workflows/deploy.yml` -- push to master -> ACR build -> Container Apps update |
| Secrets | 1Password via `1password/load-secrets-action@v2` at deploy time |

### GitHub Org Secrets Required

| Secret | Source |
|--------|--------|
| `OP_SERVICE_ACCOUNT_TOKEN` | 1Password `Claude-CLI-2` service account |
| `AZURE_CREDENTIALS` | Service principal `github-majans-agents` JSON |

### 1Password References

| Env Var | 1Password Reference |
|---------|-------------------|
| `GRAPH_TENANT_ID` | `op://Majans Dev/Graph MCP Agent/tenant_id` |
| `GRAPH_CLIENT_ID` | `op://Majans Dev/Graph MCP Agent/client_id` |
| `GRAPH_CLIENT_SECRET` | `op://Majans Dev/Graph MCP Agent/client_secret` |
| `GRAPH_API_KEY` | `op://Majans Dev/Graph MCP API Key/credential` |

## Entra App Requirements

### App 1: `Majans-Graph-MCP-Agent` (app-only, for agents)

**Application** permissions with admin consent:

- `User.ReadWrite.All` — list/get/update users
- `Group.ReadWrite.All` — list/manage groups and members
- `Directory.Read.All` — directory metadata, signInActivity
- `Mail.Read` — read any user's mailbox
- `Mail.Send` — send email from any user's mailbox
- `Files.Read.All` — read any user's OneDrive files
- `Calendars.Read` — read any user's calendar events
- `OnlineMeetings.Read.All` — read online meeting details
- `OnlineMeetingTranscript.Read.All` — read meeting transcripts
- `Sites.ReadWrite.All` — SharePoint site access, document library read/write
- `Tasks.ReadWrite.All` — Planner plans, buckets, tasks (app-only)
- `Tasks.ReadWrite` — Microsoft To Do lists and tasks
- `Chat.Read.All` — read Teams chats and messages
- `Chat.ReadWrite.All` — send messages to Teams chats

### App 2: `Graph-MCP-User` (delegated, for users)

**Setup in Azure Portal:**
1. App registrations > New registration
2. Name: `Graph-MCP-User`
3. Supported account types: Single tenant (Majans only)
4. Redirect URI: leave blank (device-code doesn't need it)
5. Authentication > Advanced settings > Allow public client flows: **Yes**

**Delegated** permissions (admin consent recommended for smooth UX):

- `User.Read` — read own profile
- `User.ReadBasic.All` — read basic profiles of other users
- `Mail.ReadWrite` — read/manage own email
- `Mail.Send` — send email as self
- `Files.ReadWrite` — own OneDrive files
- `Calendars.ReadWrite` — own calendar events
- `Sites.Read.All` — SharePoint sites (scoped by site permissions)
- `Tasks.ReadWrite` — Planner + To Do tasks
- `Chat.ReadWrite` — Teams chats user is part of
- `OnlineMeetings.Read` — own meeting details

**Client ID**: `02fa0ea1-4b30-4bd9-9c4a-483f97d63b21`
**1Password**: `op://Majans Dev/Graph MCP User/client_id`

## Meeting Transcript Workflow

To fetch a meeting transcript programmatically:
1. `graph_list_events` — find meetings in a date range (filter `isOnlineMeeting eq true`)
2. `graph_get_online_meeting` — get meeting ID from the join URL (user must be organizer)
3. `graph_list_meeting_transcripts` — list available transcripts for the meeting
4. `graph_get_meeting_transcript_content` — download VTT transcript text

The VTT transcript includes speaker names and timestamps. Agents can summarize it with Claude.

## Planner Workflow

Planner tasks require etag-based concurrency control for updates and deletes:
1. `graph_list_groups` — find the M365 group that owns the plan
2. `graph_list_plans` — list plans for the group
3. `graph_list_buckets` — get bucket IDs for task placement
4. `graph_list_plan_tasks` — list tasks (includes etags)
5. `graph_update_task` / `graph_delete_task` — pass the etag from step 4
