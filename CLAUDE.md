# connector-graph

MCP server exposing 11 Microsoft Graph API tools for user management, group operations, and license inventory.

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
│   ├── client.ts         # GraphClient — client credentials auth, token cache, 429 retry, OData pagination
│   └── types.ts          # GraphUser, GraphGroup, GraphSubscribedSku, ODataResponse<T>
├── tools/
│   ├── users.ts          # graph_list_users, graph_get_user, graph_update_user
│   ├── groups.ts         # graph_list_groups, graph_get_group_members, graph_add_group_member, graph_remove_group_member
│   └── licenses.ts       # graph_list_subscribed_skus, graph_list_user_licenses
└── utils/
    ├── config.ts         # GRAPH_TENANT_ID, CLIENT_ID, CLIENT_SECRET validation
    └── logger.ts         # Stderr logger
```

## MCP Tools (11 total)

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
- `Reports.Read.All` — usage reports (future)

## Phase 2 (future)

- SharePoint tools: site access, drive listing, file read/write
- Mail tools: list/get messages
- Calendar tools
