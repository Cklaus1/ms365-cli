# ms365-cli

A standalone command-line interface for Microsoft 365, powered by the Microsoft Graph API.

Manage your email, calendar, files, tasks, contacts, Teams, SharePoint, Planner, OneNote, and more — all from the terminal.

```
ms365 mail triage
ms365 calendar events --top 5 -o table
ms365 workflow standup
ms365 drive search "quarterly report"
ms365 --org-mode teams list -o table
```

---

## Table of Contents

- [Features](#features)
- [Requirements](#requirements)
- [Installation](#installation)
- [Quick Start](#quick-start)
- [Authentication](#authentication)
- [Global Options](#global-options)
- [Output Formats](#output-formats)
- [Commands](#commands)
  - [mail](#mail) — Outlook mail (35 commands)
  - [calendar](#calendar) — Calendar events (15 commands)
  - [drive](#drive) — OneDrive files (13 commands)
  - [todo](#todo) — Microsoft To Do (8 commands)
  - [contacts](#contacts) — Outlook contacts (9 commands)
  - [teams](#teams) — Microsoft Teams (8 commands)
  - [chat](#chat) — Teams chat (5 commands)
  - [sharepoint](#sharepoint) — SharePoint sites (12 commands)
  - [planner](#planner) — Microsoft Planner (21 commands)
  - [onenote](#onenote) — OneNote notebooks (5 commands)
  - [excel](#excel) — Excel workbooks (5 commands)
  - [search](#search) — Microsoft 365 search (2 commands)
  - [user](#user) — User profile, presence, OOO (14 commands)
  - [groups](#groups) — M365 Groups & distribution lists (7 commands)
  - [workflow](#workflow) — Productivity workflows (6 commands)
  - [schema](#schema) — API introspection
  - [generate-skills](#generate-skills) — AI agent skill files
- [OData Query Parameters](#odata-query-parameters)
- [Pagination](#pagination)
- [Dry Run](#dry-run)
- [Error Handling](#error-handling)
- [Environment Variables](#environment-variables)
- [MCP Server Compatibility](#mcp-server-compatibility)
- [Architecture](#architecture)
- [Development](#development)
- [License](#license)

---

## Features

- **16 service commands** with **174 subcommands** covering the full Microsoft 365 suite
- **Productivity workflows** — standup report, meeting prep, email-to-task, weekly digest, focus time finder
- **Smart email** — triage view, full message reader, inbox rules, batch operations, follow-up flags, categories
- **Scheduling** — accept/decline/tentative meeting invites, free/busy lookup, room finder
- **File management** — search, upload (streaming + MIME detection), share, create folders, quota
- **Presence & OOO** — get/set availability status and out-of-office auto-replies
- **Schema introspection** — inspect API endpoints, parameters, and scopes
- **AI agent skills** — auto-generate SKILL.md files for Claude Code and other agents
- **Interactive setup wizard** — guided Azure AD app registration
- **5 output formats** — JSON, table (nested object flattening), CSV, YAML, text
- **OData query support** — `--filter`, `--select`, `--top`, `--orderby`, `--search`
- **Auto-pagination** — `--all` with `--page-limit` and `--page-delay` controls
- **Dry-run mode** — preview API requests without executing
- **Retry with backoff** — automatic retry on 429 rate limits and 5xx server errors
- **`.env` file support** — auto-loads from current directory, compatible with MCP server env vars
- **Organization mode** — opt-in access to Teams, SharePoint, shared mailboxes, Groups
- **Graph API version** — switch between `v1.0` and `beta` with `--api-version`
- **Structured logging** — JSON log files with daily rotation
- **Secure token storage** — OS keychain via `keytar` with file-based fallback
- **Zero MCP dependency** — standalone CLI, talks directly to Microsoft Graph API

---

## Requirements

- **Node.js** >= 18 (recommended >= 20)
- A **Microsoft account** (personal or work/school)
- An **Azure AD app registration** with client ID (see [Authentication](#authentication))

---

## Installation

### From source

```bash
git clone <repo-url> ms365-cli
cd ms365-cli
npm install
npm run build
npm link        # makes 'ms365' available globally
```

### Run without installing globally

```bash
npm run dev -- mail list --top 5
# or after building:
node dist/index.js mail list --top 5
```

---

## Quick Start

```bash
# 1. Set up your .env with your Azure AD app credentials
echo 'MS365_CLI_CLIENT_ID=your-client-id' > .env
echo 'MS365_CLI_TENANT_ID=your-tenant-id' >> .env

# 2. Log in (opens device code flow in terminal)
ms365 auth login

# 3. Verify your session
ms365 auth status

# 4. Triage your inbox
ms365 mail triage

# 5. Check today's schedule
ms365 workflow standup

# 6. Send an email
ms365 mail send --to "colleague@example.com" --subject "Hello" --body "Sent from ms365-cli"
```

---

## Authentication

ms365-cli uses **MSAL** (Microsoft Authentication Library) with the **device code flow**.

### Prerequisites

You need an Azure AD app registration. Either:
- Run `ms365 auth setup` for an interactive walkthrough, or
- Register manually in the [Azure Portal](https://portal.azure.com/#blade/Microsoft_AAD_RegisteredApps/ApplicationsListBlade) and set `MS365_CLI_CLIENT_ID` in your `.env`

### Login

```bash
ms365 auth login
```

Prints a URL and code. Open the URL in your browser, enter the code, and sign in. Tokens are cached automatically.

### Token storage

| OS      | Backend                    |
|---------|----------------------------|
| macOS   | Keychain                   |
| Windows | Credential Manager         |
| Linux   | libsecret (GNOME Keyring)  |

Fallback: `~/.config/ms365-cli/token-cache.json`

### Multiple accounts

```bash
ms365 auth accounts            # List cached accounts
ms365 auth select <accountId>  # Switch active account
ms365 auth remove <accountId>  # Remove an account
```

### Bring your own token

```bash
# In .env
MS365_CLI_TOKEN=eyJ0eXAiOiJKV1Qi...
```

---

## Global Options

| Option                   | Description                                                  |
|--------------------------|--------------------------------------------------------------|
| `--version`              | Print version number                                         |
| `--org-mode`             | Enable Teams, SharePoint, shared mailboxes, Groups           |
| `--read-only`            | Restrict to GET operations only                              |
| `--dry-run`              | Preview the API request without executing                    |
| `--api-version <ver>`    | Graph API version: `v1.0` (default) or `beta`               |
| `--help`                 | Show help                                                    |

```bash
ms365 mail list                          # Personal mode
ms365 --org-mode teams list              # Organization mode
ms365 --api-version beta user me         # Beta API
ms365 --dry-run mail send --to "a@b.com" --subject "x" --body "y"   # Preview
```

---

## Output Formats

Every data command supports `-o` / `--output`:

| Format  | Flag          | Description                                     |
|---------|---------------|-------------------------------------------------|
| JSON    | `-o json`     | Pretty-printed JSON (default)                   |
| Table   | `-o table`    | ASCII table with nested object flattening       |
| CSV     | `-o csv`      | Comma-separated values with header row          |
| YAML    | `-o yaml`     | YAML format                                     |
| Text    | `-o text`     | Human-readable key-value pairs                  |

Table and CSV flatten nested objects into dot-notation columns:

```
from.emailAddress.name  from.emailAddress.address  hasAttachments  flag.flagStatus
----------------------  -------------------------  --------------  ---------------
John Doe                john@acme.com              true            flagged
Jane Smith              jane@acme.com              false           notFlagged
```

---

## Commands

### mail

35 commands for Outlook mail.

#### Triage and reading

```bash
ms365 mail triage                          # Unread inbox summary (table)
ms365 mail triage --top 30 --include-read  # Include read messages
ms365 mail read <messageId>                # Full message with body, attachments, flags
ms365 mail read <messageId> --text         # Plain text body
ms365 mail list --top 20 -o table          # Raw message list
ms365 mail list --folder "Sent Items"      # By folder name
ms365 mail list --folder drafts            # Well-known folder names
ms365 mail list --search "quarterly report"
ms365 mail get <messageId>                 # Raw API response
```

#### Sending and replying

```bash
ms365 mail send --to "a@b.com" --subject "Hi" --body "Hello"
ms365 mail send --to "a@b.com,b@c.com" --cc "d@e.com" --subject "Update" --body "<h1>HTML</h1>" --html
ms365 mail reply <messageId> --body "Thanks!"
ms365 mail reply <messageId> --body "Noted." --reply-all
ms365 mail forward <messageId> --to "boss@co.com" --comment "FYI"
```

#### Drafts

```bash
ms365 mail draft --to "a@b.com" --subject "WIP" --body "..."
ms365 mail drafts                          # List drafts
ms365 mail update-draft <id> --subject "Updated" --body "New content"
ms365 mail send-draft <id>                 # Send a draft
ms365 mail add-attachment <id> --name "file.pdf" --content-type "application/pdf" --content <base64>
```

#### Inbox management

```bash
ms365 mail mark-read <id>                  # Mark as read
ms365 mail mark-unread <id>                # Mark as unread
ms365 mail flag <id> --due 2026-03-25      # Follow-up flag
ms365 mail unflag <id>                     # Remove flag
ms365 mail complete <id>                   # Mark flag as complete
ms365 mail categorize <id> --categories "Red category,Blue category"
ms365 mail move <id> --folder "Archive"
ms365 mail delete <id>
```

#### Batch operations

```bash
ms365 mail batch-read "id1,id2,id3"               # Mark multiple as read
ms365 mail batch-move "id1,id2" --folder "Archive" # Move multiple
```

#### Folders

```bash
ms365 mail folders -o table                # List folders with item counts
ms365 mail folders --parent inbox          # Child folders
ms365 mail folder "Sent Items"             # Lookup folder by name
ms365 mail create-folder --name "Projects"
```

#### Inbox rules

```bash
ms365 mail rules                           # List rules
ms365 mail rule <ruleId>                   # Get rule details
ms365 mail create-rule --name "Auto-archive" --from "news@co.com" --move-to archive --mark-read
ms365 mail update-rule <ruleId> --enabled false   # Disable a rule
```

#### Shared mailboxes (requires `--org-mode`)

```bash
ms365 --org-mode mail shared-list shared@company.com
ms365 --org-mode mail shared-get shared@company.com <messageId>
ms365 --org-mode mail shared-send shared@company.com --to "a@b.com" --subject "..." --body "..."
```

---

### calendar

15 commands. Alias: `cal`.

```bash
ms365 cal events --top 20 -o table
ms365 cal events --calendar <calendarId>
ms365 cal view --start 2026-03-17T00:00:00 --end 2026-03-23T23:59:59
ms365 cal get <eventId>

# Create / update / delete
ms365 cal create --subject "Standup" --start "2026-03-17T09:00:00" --end "2026-03-17T09:30:00" --attendees "a@b.com"
ms365 cal update <eventId> --subject "Renamed" --location "Room B"
ms365 cal delete <eventId>

# Respond to invites
ms365 cal accept <eventId> --comment "See you there"
ms365 cal decline <eventId>
ms365 cal tentative <eventId>

# Scheduling
ms365 cal free-busy --emails "alice@co.com,bob@co.com" --start "2026-03-17T08:00:00" --end "2026-03-17T18:00:00"
ms365 --org-mode cal rooms
ms365 --org-mode cal room-lists

# Calendar management
ms365 cal list
ms365 cal create-calendar --name "Side Project" --color lightGreen
ms365 cal update-calendar <calendarId> --name "Renamed"
```

---

### drive

13 commands for OneDrive.

```bash
ms365 drive list -o table                  # List drives
ms365 drive files <driveId> <folderId>     # List files in folder
ms365 drive search "quarterly report"      # Search files
ms365 drive info <driveId> <itemId>        # File/folder metadata
ms365 drive quota                          # Storage usage

# Upload / download
ms365 drive upload <driveId> <parentId> ./report.xlsx   # Auto-detects MIME, streams large files
ms365 drive upload <driveId> <parentId> ./data.bin --content-type "application/gzip"
ms365 drive download <driveId> <itemId> <childId> --out ./report.pdf
ms365 drive download-shared <shareId>

# Organize
ms365 drive mkdir <driveId> <parentId> "New Folder"
ms365 drive delete <driveId> <itemId>

# Sharing
ms365 drive share <driveId> <itemId> --type edit --scope organization
ms365 drive share <driveId> <itemId> --type view --password "secret" --expiry "2026-04-01"
ms365 drive permissions <driveId> <itemId>
```

---

### todo

8 commands for Microsoft To Do.

```bash
ms365 todo lists -o table
ms365 todo create-list --name "Shopping"
ms365 todo update-list <listId> --name "[ARCHIVED] Old List"

ms365 todo tasks <listId> -o table
ms365 todo create <listId> --title "Buy groceries" --due 2026-03-20 --importance high
ms365 todo update <listId> <taskId> --status completed
ms365 todo get <listId> <taskId>
ms365 todo delete <listId> <taskId>
```

---

### contacts

9 commands for Outlook contacts.

```bash
ms365 contacts list -o table
ms365 contacts list --search "John"
ms365 contacts get <contactId>
ms365 contacts create --given-name "Jane" --surname "Doe" --email "jane@co.com" --company "Acme"
ms365 contacts update <contactId> --email "new@co.com"
ms365 contacts delete <contactId>

# Contact folders
ms365 contacts folders
ms365 contacts create-folder --name "VIP"
ms365 contacts update-folder <folderId> --name "[ARCHIVED] Old Folder"
ms365 contacts folder-contacts <folderId>
```

---

### teams

8 commands. Requires `--org-mode`.

```bash
ms365 --org-mode teams list -o table
ms365 --org-mode teams get <teamId>
ms365 --org-mode teams members <teamId>
ms365 --org-mode teams channels <teamId>
ms365 --org-mode teams channel <teamId> <channelId>
ms365 --org-mode teams channel-messages <teamId> <channelId> --top 20
ms365 --org-mode teams send-channel-message <teamId> <channelId> --body "Hello!"
ms365 --org-mode teams reply-channel-message <teamId> <channelId> <messageId> --body "Agreed."
```

---

### chat

5 commands. Requires `--org-mode`.

```bash
ms365 --org-mode chat list
ms365 --org-mode chat messages <chatId> --top 20
ms365 --org-mode chat send <chatId> --body "Quick question..."
ms365 --org-mode chat reply <chatId> <messageId> --body "Got it."
```

---

### sharepoint

12 commands. Alias: `sp`. Requires `--org-mode`.

```bash
ms365 --org-mode sp search "Marketing"
ms365 --org-mode sp site <siteId>
ms365 --org-mode sp site-by-path contoso.sharepoint.com /sites/marketing
ms365 --org-mode sp drives <siteId>
ms365 --org-mode sp lists <siteId> -o table
ms365 --org-mode sp list-items <siteId> <listId> --expand "fields" --all
ms365 --org-mode sp delta                  # Track site changes
```

---

### planner

21 commands for Microsoft Planner.

```bash
ms365 planner plans -o table
ms365 planner my-tasks -o table
ms365 planner tasks <planId>
ms365 planner task <taskId>
ms365 planner task-details <taskId>
ms365 planner buckets <planId>

ms365 planner create-task --plan <planId> --title "Review" --bucket <bucketId> --due 2026-03-25
ms365 planner update-task <taskId> --percent 50 --etag 'W/"abc"'
ms365 planner create-plan --title "Sprint 42" --group <groupId>
ms365 planner create-bucket --plan <planId> --name "To Do"

# Safe archive (rename instead of delete)
ms365 planner update-plan <planId> --title "[ARCHIVED] Old Plan" --etag 'W/"abc"'
```

---

### onenote

5 commands.

```bash
ms365 onenote notebooks -o table
ms365 onenote sections <notebookId>
ms365 onenote pages <sectionId>
ms365 onenote page <pageId>
ms365 onenote create-page <sectionId> --title "Notes" --content "<p>Action items</p>"
```

---

### excel

5 commands for Excel workbooks stored in OneDrive.

```bash
ms365 excel worksheets <driveId> <itemId>
ms365 excel range <driveId> <itemId> <worksheetId> "A1:D10"
ms365 excel create-chart <driveId> <itemId> <worksheetId> --type ColumnClustered --source "A1:D10" --series-by columns
ms365 excel format-range <driveId> <itemId> <worksheetId> --font-bold --font-color "#FF0000"
ms365 excel sort-range <driveId> <itemId> <worksheetId> --fields '[{"key":0,"ascending":true}]'
```

---

### search

2 commands.

```bash
ms365 search query "quarterly report"
ms365 search query "budget" --entity message --top 5
ms365 search query "alpha" --entity "driveItem,message,site"
ms365 search people "John"
```

Entity types: `message`, `driveItem`, `listItem`, `site`, `event`, `chatMessage`.

---

### user

14 commands for profile, presence, OOO, and settings.

```bash
ms365 user me
ms365 --org-mode user list --filter "department eq 'Engineering'" -o table

# Out of Office
ms365 user ooo                             # Get current auto-reply settings
ms365 user set-ooo --status scheduled --internal "I'm OOO until Monday" --start "2026-03-20T00:00:00" --end "2026-03-24T00:00:00"
ms365 user set-ooo --status disabled       # Turn off

# Presence
ms365 user presence                        # Get current status
ms365 user set-presence --availability DoNotDisturb --activity Focusing
ms365 user clear-presence                  # Revert to automatic

# Profile photo
ms365 user photo                           # Metadata
ms365 user download-photo --out avatar.jpg
ms365 user upload-photo ./new-avatar.jpg

# Mailbox settings
ms365 user mailbox-settings
ms365 user set-timezone "America/New_York"
ms365 user set-language "en-US"
ms365 user working-hours
```

---

### groups

7 commands. Requires `--org-mode`.

```bash
ms365 --org-mode groups list -o table
ms365 --org-mode groups my                 # Groups you belong to
ms365 --org-mode groups get <groupId>
ms365 --org-mode groups members <groupId>
ms365 --org-mode groups owners <groupId>
ms365 --org-mode groups distribution-lists # Mail-enabled DLs
ms365 --org-mode groups teams              # Groups with Teams
```

---

### workflow

6 cross-service productivity commands. Alias: `wf`.

```bash
ms365 wf standup                           # Today's calendar + unread count
ms365 wf meeting-prep                      # Next meeting: attendees, agenda, join link
ms365 wf email-to-task <messageId>         # Convert email to To Do task
ms365 wf digest                            # Weekly summary
ms365 wf focus-time                        # Find calendar gaps for deep work
ms365 --org-mode wf file-announce <driveId> <itemId> --team <id> --channel <id>
```

---

### schema

Inspect API endpoints before calling them.

```bash
ms365 schema --list                        # List all documented endpoints
ms365 schema mail.send                     # Parameters, scopes, response shape
ms365 schema calendar.create
ms365 schema drive.upload
```

---

### generate-skills

Generate SKILL.md files for AI agent integration.

```bash
ms365 generate-skills                      # All skills to ./skills/
ms365 generate-skills --output-dir ./my-skills
ms365 generate-skills --filter "mail|calendar"
```

---

## OData Query Parameters

Most list commands support:

| Flag              | OData Parameter | Example                                      |
|-------------------|-----------------|----------------------------------------------|
| `--top <n>`       | `$top`          | `--top 25`                                   |
| `--filter <expr>` | `$filter`       | `--filter "isRead eq false"`                 |
| `--select <fields>` | `$select`     | `--select "subject,from,receivedDateTime"`   |
| `--orderby <field>` | `$orderby`   | `--orderby "receivedDateTime desc"`          |
| `--search <query>` | `$search`      | `--search "quarterly report"`                |

### Filter examples

```bash
ms365 mail list --filter "isRead eq false"
ms365 mail list --filter "from/emailAddress/address eq 'boss@co.com'"
ms365 cal events --filter "start/dateTime ge '2026-03-01T00:00:00Z'"
ms365 todo tasks <listId> --filter "importance eq 'high'"
ms365 --org-mode user list --filter "department eq 'Engineering'"
```

---

## Pagination

```bash
# Single page (default)
ms365 mail list --top 25

# All pages
ms365 mail list --all

# Controlled pagination
ms365 mail list --all --page-limit 5           # Max 5 pages
ms365 contacts list --all --page-delay 200     # 200ms between pages
ms365 drive files d1 f1 --all --page-limit 10 --page-delay 100 -o csv > files.csv
```

Default page limit: **100 pages**. Default delay: **0ms**.

---

## Dry Run

Preview API requests without executing:

```bash
ms365 --dry-run mail list --top 5 --filter "isRead eq false"
# [dry-run] GET https://graph.microsoft.com/v1.0/me/messages?$top=5&$filter=isRead%20eq%20false

ms365 --dry-run mail send --to "a@b.com" --subject "Test" --body "Hello"
# [dry-run] POST https://graph.microsoft.com/v1.0/me/sendMail
# [dry-run] Body: {"message":{"subject":"Test",...}}
```

Dry-run skips authentication, so it works even when not logged in.

---

## Error Handling

- **401** — suggests `ms365 auth login`
- **403** — suggests `--org-mode` when scope-related
- **429** — automatic retry with exponential backoff (up to 3 retries)
- **5xx** — automatic retry with backoff
- All commands wrapped with user-friendly error messages (no stack traces)

---

## Environment Variables

| Variable                  | Description                                          | Required |
|---------------------------|------------------------------------------------------|----------|
| `MS365_CLI_CLIENT_ID`    | Azure AD app client ID                               | Yes      |
| `MS365_CLI_TENANT_ID`    | Azure AD tenant ID                                   | No (defaults to `common`) |
| `MS365_CLI_TOKEN`        | Pre-obtained OAuth2 access token                     | No       |
| `MS365_CLI_ORG_MODE`     | Enable organization mode (`true`)                    | No       |
| `MS365_CLI_API_VERSION`  | Graph API version (`v1.0` or `beta`)                 | No       |
| `MS365_CLI_LOG_LEVEL`    | Log level: `debug`, `info`, `warn`, `error`, `none`  | No       |
| `MS365_CLI_LOG_FILE`     | JSON log file path (enables daily rotation)          | No       |

The CLI auto-loads `.env` from the current working directory.

---

## MCP Server Compatibility

If you're also using the [ms-365-mcp-server](https://github.com/softeria/ms-365-mcp-server), the CLI accepts its env var names as fallbacks:

| MCP Server Variable       | CLI Equivalent          |
|---------------------------|-------------------------|
| `MS365_MCP_CLIENT_ID`    | `MS365_CLI_CLIENT_ID`  |
| `MS365_MCP_TENANT_ID`    | `MS365_CLI_TENANT_ID`  |
| `MS365_MCP_OAUTH_TOKEN`  | `MS365_CLI_TOKEN`      |
| `MS365_MCP_ORG_MODE`     | `MS365_CLI_ORG_MODE`   |

You can share the same `.env` file between both projects.

---

## Architecture

```
ms365-cli/
├── src/
│   ├── index.ts                Entry point — .env loading, CLI setup, command registration
│   ├── auth.ts                 MSAL device code auth, token cache, multi-account
│   ├── graph-client.ts         Graph API client — retry, pagination, dry-run, streaming upload
│   ├── formatter.ts            Output formatting (JSON, table, CSV, YAML, text)
│   ├── errors.ts               Error handling wrapper for all commands
│   ├── logger.ts               Structured JSON logging with daily rotation
│   ├── mime.ts                 MIME type detection (70+ extensions)
│   └── commands/
│       ├── auth.ts             login, logout, status, accounts, setup (7)
│       ├── mail.ts             list, triage, read, send, reply, forward, drafts, rules, batch (35)
│       ├── calendar.ts         events, view, create, respond, free-busy, rooms (15)
│       ├── drive.ts            files, search, upload, share, mkdir, quota (13)
│       ├── todo.ts             lists, tasks, create/update (8)
│       ├── contacts.ts         CRUD + contact folders (9)
│       ├── teams.ts            teams, channels, messages (8)
│       ├── chat.ts             chats, messages, reply (5, registered in teams.ts)
│       ├── sharepoint.ts       sites, lists, items, delta (12)
│       ├── planner.ts          plans, tasks, buckets, details (21)
│       ├── onenote.ts          notebooks, sections, pages (5)
│       ├── excel.ts            worksheets, ranges, charts, format (5)
│       ├── search.ts           query, people (2)
│       ├── user.ts             profile, OOO, presence, photo, settings (14)
│       ├── groups.ts           groups, members, DLs (7)
│       ├── workflow.ts         standup, meeting-prep, digest, focus-time (6)
│       ├── schema.ts           API endpoint introspection
│       ├── generate-skills.ts  SKILL.md generation for AI agents
│       └── setup.ts            Interactive auth setup wizard
├── dist/                       Compiled output (ESM)
├── .gitignore
├── package.json
├── tsconfig.json
├── tsup.config.ts
├── LICENSE                     CC BY-NC 4.0
└── README.md
```

### Key design decisions

- **Standalone** — no MCP server, no MCP SDK. Talks directly to Microsoft Graph API via `fetch`.
- **Commander.js** — CLI framework with subcommands, options, and help generation.
- **MSAL** — Microsoft's official auth library for device code flow and silent token refresh.
- **dotenv** — auto-loads `.env` from working directory, compatible with MCP server env vars.
- **Streaming upload** — files <= 4MB use simple PUT; larger files use chunked upload sessions (10MB chunks).
- **MIME auto-detection** — 70+ file extensions mapped; override with `--content-type`.
- **Structured logging** — JSON-line log files with daily rotation via `MS365_CLI_LOG_FILE`.
- **No destructive batch ops** — batch-delete intentionally omitted for agent safety.
- **Archive by rename** — calendars, todo lists, contact folders can be archived by renaming with `[ARCHIVED]` prefix instead of deleting.

### Runtime dependencies

| Package            | Purpose                                          |
|--------------------|--------------------------------------------------|
| `@azure/msal-node` | Microsoft authentication (device code, token cache) |
| `commander`        | CLI argument parsing and help                    |
| `dotenv`           | `.env` file loading                              |
| `keytar`           | OS keychain access for secure token storage      |

---

## Development

### Build

```bash
npm run build
```

### Dev mode

```bash
npm run dev -- mail list --top 5
```

### Adding a new command

1. Create `src/commands/myservice.ts`
2. Export `registerMyServiceCommands(program, getClient)`
3. Import and call in `src/index.ts`
4. `npm run build`

---

## License

**MIT** — free to use, modify, and distribute, including commercially, with attribution.

See [LICENSE](./LICENSE) for full text.
