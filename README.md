# Google Workspace MCP Server

Python MCP server giving Fred (ReisierX's AI agent) direct access to Google Workspace — Gmail, Drive, Docs, Sheets, Slides, and Calendar.

**86 tools** across 6 services. Runs on the ReisierX Mac mini, connected to OpenClaw via mcporter.

---

## Services & Tools

| Service | Tools |
|---------|-------|
| Google Docs | Create, read, append, format, tables, lists, headings, batch updates |
| Google Sheets | Read, write, create, format, append rows, manage sheets |
| Google Drive | List, search, get, create folders, copy, move, rename, delete, share, upload, download, export |
| Google Calendar | List events, get event, create event, update event, delete event, list calendars |
| Gmail | Search, read, send, reply, create draft, list labels, manage labels |
| Google Slides | Create, read, add slides, add text, add images, export |

---

## Architecture

```
OpenClaw (agent) → mcporter → Python MCP stdio → Google APIs
```

- **Server:** `/Users/reisierx/repos/google-workspace-mcp/server.py`
- **Venv:** `.venv`
- **mcporter config:** `~/.mcporter/mcporter.json` (server registered as `google-workspace`)
- **Auth:** OAuth2 via `credentials.json` + `token.json` (fred@reisierx.com)
- **GCP Project:** `reisierx-fred`

---

## Usage (from within OpenClaw)

```bash
# List all tools
mcporter list google-workspace

# Call a tool
mcporter call google-workspace.<tool_name> --args '{"key": "value"}'
```

Examples:

```bash
mcporter call google-workspace.google_calendar_list --args '{"max_results": 5}'
mcporter call google-workspace.gmail_search --args '{"query": "from:goncalo@reisierx.com"}'
mcporter call google-workspace.google_docs_create --args '{"title": "Meeting Notes"}'
```

---

## Setup (from scratch)

If the server needs to be set up on a new machine:

```bash
# 1. Clone
git clone https://github.com/reisierx/google-workspace-mcp.git
cd google-workspace-mcp

# 2. Create venv and install deps
python3 -m venv .venv
.venv/bin/pip install -r requirements.txt

# 3. Add credentials.json (from GCP project reisierx-fred, OAuth client "Fred")
# Then run once to trigger OAuth:
.venv/bin/python server.py
# → browser opens, sign in as fred@reisierx.com, grant all scopes
# → token.json is created automatically

# 4. Register with mcporter
mcporter config add google-workspace \
  --command "/path/to/repo/.venv/bin/python" \
  --arg "/path/to/repo/server.py" \
  --description "Google Workspace MCP" \
  --scope home

# 5. Test
mcporter list google-workspace
```

---

## Re-authentication

If the token expires or scopes change, delete `token.json` and run `server.py` once manually. It will re-open the browser flow.

```bash
rm token.json
.venv/bin/python server.py
```

---

## Upstream

Forked from `goncaloreis/google-workspace-mcp`. Changes from upstream:
- Google Tasks removed (scope, service, 12 tools)
- Shared Drive support added to all Drive API calls (`supportsAllDrives=True`)
- README rewritten for ReisierX deployment
