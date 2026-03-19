@echo off
REM Acquires a Graph user token via MSAL device-code and launches mcp-remote
REM as a stdio-to-StreamableHTTP proxy. Register in Claude Code MCP config as:
REM
REM   "graph": {
REM     "type": "stdio",
REM     "command": "cmd",
REM     "args": ["/c", "C:\\path\\to\\connector-graph\\start-mcp.cmd"]
REM   }
REM
REM First run: opens browser for device-code login.
REM Subsequent runs: silently refreshes cached token.
REM
REM Prerequisites:
REM   - Python 3.10+ with msal: pip install msal
REM   - Node.js (for npx mcp-remote)
REM   - GRAPH_MCP_CLIENT_ID env var set to the Graph-MCP-User Entra app client ID

set SCRIPT_DIR=%~dp0

REM Get fresh token (stderr shows device-code prompt if needed)
for /f "usebackq delims=" %%t in (`python "%SCRIPT_DIR%get-user-token.py"`) do set GRAPH_TOKEN=%%t

if "%GRAPH_TOKEN%"=="" (
    echo ERROR: Failed to acquire Graph token. >&2
    echo Run 'python "%SCRIPT_DIR%get-user-token.py"' manually to authenticate. >&2
    exit /b 1
)

REM Launch mcp-remote as stdio proxy to the remote StreamableHTTP MCP server
npx -y mcp-remote@latest https://graph.majans.com/mcp --header "Authorization: Bearer %GRAPH_TOKEN%"
