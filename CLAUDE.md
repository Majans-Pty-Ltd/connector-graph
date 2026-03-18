# connector-graph

MCP server exposing 42 Microsoft Graph API tools for user management, groups, licenses, mail, OneDrive, calendar, meeting transcripts, SharePoint, Planner, To Do, and Teams Chat.

## Tech Stack

- **Runtime**: Node.js (ES2022, ES modules)
- **Language**: TypeScript 5.7+
- **Protocol**: MCP SDK (stdio + StreamableHTTP transport)
- **Auth**: Client credentials (Service Principal, app-only)
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
├── index.ts              # MCP server entry point + graph_auth_status tool + dual transport (stdio/HTTP)
├── api/
│   ├── client.ts         # GraphClient — client credentials auth, token cache, 429 retry, OData pagination, getText for non-JSON, If-Match header support
│   └── types.ts          # All TypeScript interfaces (users, groups, licenses, mail, drive, calendar, meetings, transcripts, sharepoint, planner, todo, teams)
├── tools/
│   ├── users.ts          # graph_list_users, graph_get_user, graph_update_user
│   ├── groups.ts         # graph_list_groups, graph_get_group_members, graph_add_group_member, graph_remove_group_member
│   ├── licenses.ts       # graph_list_subscribed_skus, graph_list_user_licenses
│   ├── mail.ts           # graph_send_mail, graph_search_mail, graph_read_mail, graph_list_attachments
│   ├── onedrive.ts       # graph_list_drive_items, graph_get_drive_item_content
│   ├── calendar.ts       # graph_list_events, graph_get_online_meeting, graph_list_meeting_transcripts, graph_get_meeting_transcript_content
│   ├── sharepoint.ts     # graph_list_sites, graph_get_site, graph_list_site_drives, graph_list_site_drive_items, graph_get_site_file_content, graph_search_site_files
│   ├── planner.ts        # graph_list_plans, graph_get_plan, graph_list_buckets, graph_list_plan_tasks, graph_get_task, graph_create_task, graph_update_task, graph_delete_task
│   ├── todo.ts           # graph_list_todo_lists, graph_list_todo_tasks, graph_create_todo_task, graph_update_todo_task, graph_complete_todo_task
│   └── teams.ts          # graph_list_chats, graph_list_chat_messages, graph_send_chat_message
└── utils/
    ├── config.ts         # GRAPH_TENANT_ID, CLIENT_ID, CLIENT_SECRET validation
    └── logger.ts         # Stderr logger
```

## MCP Tools (42 total)

### Auth (1)
| Tool | Description |
|------|-------------|
| `graph_auth_status` | Verify auth, show org name and client ID |

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

### Mail (4)
| Tool | Description |
|------|-------------|
| `graph_send_mail` | Send email from a user's mailbox (HTML/text, to/cc, importance) |
| `graph_search_mail` | Search mailbox with $search/$filter |
| `graph_read_mail` | Get full email content by message ID |
| `graph_list_attachments` | List attachments on an email |

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
- `GRAPH_CLIENT_ID` — Entra app client ID
- `GRAPH_CLIENT_SECRET` — Entra app client secret
- `PORT` — (optional) Set to enable StreamableHTTP transport (e.g. `8030`). Omit for stdio.

Use 1Password: `op run --env-file=.env.template -- npm start`

## Transport Modes

### stdio (default)
Used when `PORT` is not set. Standard MCP stdio transport for local Claude Desktop / Claude Code usage.

### StreamableHTTP
Used when `PORT` is set. Starts an HTTP server with:
- `POST /mcp` — MCP StreamableHTTP endpoint
- `GET /health` — Health check for container probes

### Docker Deployment
```bash
npm run build
docker build -t connector-graph .
docker run -p 8030:8030 \
  -e GRAPH_TENANT_ID=... \
  -e GRAPH_CLIENT_ID=... \
  -e GRAPH_CLIENT_SECRET=... \
  -e PORT=8030 \
  connector-graph
```

## Entra App Requirements

The Entra app (`Majans-Graph-MCP-Agent`) needs these **application** permissions with admin consent:

### Existing
- `User.ReadWrite.All` — list/get/update users
- `Group.ReadWrite.All` — list/manage groups and members
- `Directory.Read.All` — directory metadata, signInActivity
- `Mail.Read` — read any user's mailbox
- `Mail.Send` — send email from any user's mailbox
- `Files.Read.All` — read any user's OneDrive files
- `Calendars.Read` — read any user's calendar events
- `OnlineMeetings.Read.All` — read online meeting details
- `OnlineMeetingTranscript.Read.All` — read meeting transcripts

### New (for Phase 3 tools)
- `Sites.ReadWrite.All` — SharePoint site access, document library read/write
- `Tasks.ReadWrite.All` — Planner plans, buckets, tasks (app-only)
- `Tasks.ReadWrite` — Microsoft To Do lists and tasks
- `Chat.Read.All` — read Teams chats and messages
- `Chat.ReadWrite.All` — send messages to Teams chats

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
