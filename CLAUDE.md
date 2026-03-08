# connector-graph

MCP server exposing 19 Microsoft Graph API tools for user management, groups, licenses, mail, OneDrive, calendar, and meeting transcripts.

## Tech Stack

- **Runtime**: Node.js (ES2022, ES modules)
- **Language**: TypeScript 5.7+
- **Protocol**: MCP SDK (stdio transport)
- **Auth**: Client credentials (Service Principal, app-only)
- **API**: Microsoft Graph REST API v1.0

## Key Commands

```bash
npm run build     # Compile TypeScript -> dist/
npm start         # Start MCP server (stdio)
npm run dev       # Watch mode (tsc --watch)
```

## Architecture

```
src/
├── index.ts              # MCP server entry point + graph_auth_status tool
├── api/
│   ├── client.ts         # GraphClient — client credentials auth, token cache, 429 retry, OData pagination, getText for non-JSON
│   └── types.ts          # All TypeScript interfaces (users, groups, licenses, mail, drive, calendar, meetings, transcripts)
├── tools/
│   ├── users.ts          # graph_list_users, graph_get_user, graph_update_user
│   ├── groups.ts         # graph_list_groups, graph_get_group_members, graph_add_group_member, graph_remove_group_member
│   ├── licenses.ts       # graph_list_subscribed_skus, graph_list_user_licenses
│   ├── mail.ts           # graph_search_mail, graph_read_mail, graph_list_attachments
│   ├── onedrive.ts       # graph_list_drive_items, graph_get_drive_item_content
│   └── calendar.ts       # graph_list_events, graph_get_online_meeting, graph_list_meeting_transcripts, graph_get_meeting_transcript_content
└── utils/
    ├── config.ts         # GRAPH_TENANT_ID, CLIENT_ID, CLIENT_SECRET validation
    └── logger.ts         # Stderr logger
```

## MCP Tools (19 total)

| Tool | Description |
|------|-------------|
| `graph_auth_status` | Verify auth, show org name and client ID |
| `graph_list_users` | OData filter/select/top/orderby + pagination |
| `graph_get_user` | By ID or UPN, includes signInActivity |
| `graph_update_user` | Cloud-only properties (jobTitle, accountEnabled, etc.) |
| `graph_list_groups` | Filter by mailEnabled, securityEnabled, displayName |
| `graph_get_group_members` | List group members |
| `graph_add_group_member` | Add user to group |
| `graph_remove_group_member` | Remove user from group |
| `graph_list_subscribed_skus` | License inventory (consumed/prepaid/available) |
| `graph_list_user_licenses` | Licenses assigned to a specific user |
| `graph_search_mail` | Search mailbox with $search/$filter |
| `graph_read_mail` | Get full email content by message ID |
| `graph_list_attachments` | List attachments on an email |
| `graph_list_drive_items` | List OneDrive files/folders at a path |
| `graph_get_drive_item_content` | Get file metadata and download URL |
| `graph_list_events` | Calendar events in a date range with attendees and online meeting info |
| `graph_get_online_meeting` | Get online meeting details by join URL (needed for transcript retrieval) |
| `graph_list_meeting_transcripts` | List available transcripts for an online meeting |
| `graph_get_meeting_transcript_content` | Get full meeting transcript text (VTT format with speakers and timestamps) |

## Configuration

Required env vars:
- `GRAPH_TENANT_ID` — Azure AD tenant ID
- `GRAPH_CLIENT_ID` — Entra app client ID
- `GRAPH_CLIENT_SECRET` — Entra app client secret

Use 1Password: `op run --env-file=.env.template -- npm start`

## Entra App Requirements

The Entra app (`Majans-Graph-MCP-Agent`) needs these **application** permissions with admin consent:
- `User.ReadWrite.All` — list/get/update users
- `Group.ReadWrite.All` — list/manage groups and members
- `Directory.Read.All` — directory metadata, signInActivity
- `Mail.Read` — read any user's mailbox
- `Files.Read.All` — read any user's OneDrive files
- `Calendars.Read` — read any user's calendar events
- `OnlineMeetings.Read.All` — read online meeting details
- `OnlineMeetingTranscript.Read.All` — read meeting transcripts

## Meeting Transcript Workflow

To fetch a meeting transcript programmatically:
1. `graph_list_events` — find meetings in a date range (filter `isOnlineMeeting eq true`)
2. `graph_get_online_meeting` — get meeting ID from the join URL (user must be organizer)
3. `graph_list_meeting_transcripts` — list available transcripts for the meeting
4. `graph_get_meeting_transcript_content` — download VTT transcript text

The VTT transcript includes speaker names and timestamps. Agents can summarize it with Claude.

## Phase 3 (future)

- SharePoint tools: site access, drive listing, file read/write
- Teams chat tools: read meeting chat messages (for Copilot Facilitator summaries)
- Reports tools: usage analytics
