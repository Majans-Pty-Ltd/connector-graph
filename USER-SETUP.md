# Connect to Microsoft Graph from Claude Code

This guide sets up connector-graph so you can use Microsoft Graph tools (email, calendar, OneDrive, Teams, SharePoint, Planner, To Do) from Claude Code — using **your own Majans account**. You only see your own data.

## Prerequisites

1. **Claude Code** installed and working
2. **Python 3.10+** — check with `python --version`
3. **Node.js 18+** — check with `node --version`
4. **MSAL library** — install once:
   ```
   pip install msal
   ```

## Setup (one-time, ~2 minutes)

### Step 1: Add to Claude Code MCP config

Open your Claude Code settings and add this MCP server entry.

**Option A — Global (all projects):**

Edit `~/.claude.json` and add to the `mcpServers` section:

```json
{
  "mcpServers": {
    "graph": {
      "type": "stdio",
      "command": "cmd",
      "args": ["/c", "C:\\Users\\Amit\\OneDrive - Majans Pty Ltd\\Documents 1\\GitHub\\connector-graph\\start-mcp.cmd"]
    }
  }
}
```

> **Important**: Replace the path above with the actual path to `start-mcp.cmd` on your machine. Ask Amit if you're not sure where the repo is.

**Option B — Per-project:**

Add the same entry to `.mcp.json` in any project where you want Graph access.

### Step 2: First-time authentication

1. Start Claude Code (or restart if already running)
2. Claude Code will launch `start-mcp.cmd` automatically
3. You'll see a message like:
   ```
   To sign in, use a web browser to open the page
   https://login.microsoft.com/device and enter the code XXXXXXXX
   ```
4. Open that URL in your browser
5. Enter the code shown
6. Sign in with your **Majans Entra ID** (e.g. `yourname@majans.com`)
7. Approve the permissions when prompted

That's it. Your token is cached and will refresh silently for ~90 days.

### Step 3: Verify it works

In Claude Code, ask:

> "Check my graph connection"

Claude will call `graph_auth_status` and should return something like:

```json
{
  "authenticated": true,
  "auth_mode": "delegated",
  "user": "Your Name",
  "email": "yourname@majans.com"
}
```

## What you can do

Once connected, you can ask Claude things like:

| Category | Example prompts |
|----------|----------------|
| **Email** | "Search my inbox for emails from supplier X" |
| | "Send an email to john@example.com about the meeting" |
| **Calendar** | "What meetings do I have this week?" |
| **OneDrive** | "List files in my OneDrive Documents folder" |
| **Teams** | "Show my recent Teams chats" |
| **SharePoint** | "Search the Intranet site for the leave policy" |
| **Planner** | "List my tasks in the Operations plan" |
| **To Do** | "Create a to-do to follow up with the supplier" |

## What you CAN'T do (by design)

With delegated auth, you only access **your own data**:

- You can read your own email, not other people's
- You can see your own calendar, not other people's
- You can access OneDrive files you own or that are shared with you
- You can see Teams chats you're part of
- SharePoint access follows your existing site permissions
- Planner shows plans for groups you belong to

This is intentional — it's the same as what you'd see if you logged into Outlook or Teams yourself.

## Troubleshooting

### "Failed to acquire Graph token"
Run the token script manually to see the error:
```
python "C:\path\to\connector-graph\get-user-token.py"
```

### Token expired / need to re-authenticate
Delete the cached token and restart Claude Code:
```
del %USERPROFILE%\.connector-graph\token_cache.bin
```
You'll get a new device-code prompt on next launch.

### "Access denied" on a specific tool
Your Entra ID account doesn't have permission for that resource. This is expected — talk to Amit if you need broader access.

### Claude Code doesn't show Graph tools
1. Check that the MCP entry in `~/.claude.json` has the correct path
2. Restart Claude Code
3. Check that Python and Node.js are on your PATH

## Need help?

Contact Amit (amit@majans.com) for:
- Access issues or permission requests
- Adding the connector to a new machine
- Reporting bugs or unexpected behavior
