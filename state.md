# Current State

## Status
Stable — Dual auth live (delegated + app-only). 43 tools across 10 categories.

## Active Work
Per-user delegated auth deployed. Majans staff can now connect with their own Entra ID via `start-mcp.cmd`.

## Recent Changes
- Added delegated auth path — users authenticate via MSAL device-code flow, Graph enforces their own permissions
- Entra app `Graph-MCP-User` (`02fa0ea1-4b30-4bd9-9c4a-483f97d63b21`) created with 10 delegated permissions
- `src/utils/auth.ts`: AsyncLocalStorage for per-request user token context
- `src/api/client.ts`: `getToken()` checks user token first, falls back to SP
- `src/index.ts`: Bearer token extraction, API key validation for agent path
- `get-user-token.py` + `start-mcp.cmd`: user-facing scripts for Claude Code
- Added `graph_send_mail` tool (43 tools total)
- 1Password item `Graph MCP User` created in Majans Dev vault
- Admin consent granted for all delegated permissions

## Pending
- Roll out `start-mcp.cmd` to Majans staff (see USER-SETUP.md)
- Consider bundling into a Claude Code plugin for easier distribution
- Monitor token refresh reliability across users

## Key Files for Current Work
src/utils/auth.ts, src/api/client.ts, src/index.ts, get-user-token.py, start-mcp.cmd
