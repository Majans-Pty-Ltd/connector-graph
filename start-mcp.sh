#!/bin/bash
# Acquires a Graph user token via MSAL and launches mcp-remote as a
# stdio-to-StreamableHTTP proxy. Add to Claude Code MCP config as:
#
#   "graph": {
#     "type": "stdio",
#     "command": "bash",
#     "args": ["/path/to/connector-graph/start-mcp.sh"]
#   }
#
# First run: opens browser for device-code login.
# Subsequent runs: silently refreshes cached token.

SCRIPT_DIR="$(cd "$(dirname "$0")" && pwd)"

# Get fresh token (--force-refresh ensures full ~60 min lifetime, not a stale
# cached token that's about to expire mid-session)
GRAPH_TOKEN=$(python3 "$SCRIPT_DIR/get-user-token.py" --force-refresh 2>/dev/tty)

if [ -z "$GRAPH_TOKEN" ]; then
    echo "ERROR: Failed to acquire Graph token." >&2
    echo "Run 'python3 $SCRIPT_DIR/get-user-token.py' manually to authenticate." >&2
    exit 1
fi

# Launch mcp-remote as stdio proxy to the remote StreamableHTTP MCP server
exec npx -y mcp-remote@latest https://graph.majans.com/mcp --header "Authorization: Bearer $GRAPH_TOKEN"
