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
  - [auth](#auth) — Login and accounts (7 commands)
  - [mail](#mail) — Outlook mail (36 commands)
  - [calendar](#calendar) — Calendar events (15 commands)
  - [drive](#drive) — OneDrive files (16 commands)
  - [todo](#todo) — Microsoft To Do (8 commands)
  - [contacts](#contacts) — Outlook contacts (9 commands)
  - [teams](#teams) — Microsoft Teams (8 commands)
  - [chat](#chat) — Teams chat (5 commands)
  - [sharepoint](#sharepoint) — SharePoint sites (11 commands)
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
- [Read-only Mode](#read-only-mode)
- [Error Handling and Exit Codes](#error-handling-and-exit-codes)
- [Remote Access (Tailscale)](#remote-access-tailscale)
- [Environment Variables](#environment-variables)
- [MCP Server Compatibility](#mcp-server-compatibility)
- [Architecture](#architecture)
- [Development](#development)
- [License](#license)

---

## Features

- **16 command groups** with **175 subcommands** covering the Microsoft 365 suite, plus `schema` and `generate-skills`
- **Productivity workflows** — standup report, meeting prep, email-to-task, weekly digest, focus time finder
- **Smart email** — triage view, full message reader, attachments, inbox rules, batch operations, follow-up flags, categories
- **Scheduling** — accept/decline/tentative meeting invites, free/busy lookup, room finder
- **File management** — path-based `ls`/`cat`/`cp`, search, upload (streaming + MIME detection, conflict handling), share, create folders, quota
- **Presence & OOO** — get/set preferred availability and out-of-office auto-replies
- **Schema introspection** — inspect the Graph request(s) behind every Graph-calling command, with parameters and scopes for the most-used ones
- **AI agent skills** — generate SKILL.md files for Claude Code and other agents from the registered commands
- **Interactive setup wizard** — built-in app (no Azure portal step) or your own app registration
- **5 output formats** — table on a terminal, JSON when piped; CSV, YAML and text on request
- **OData query support** — `--filter`, `--select`, `--top`, `--orderby`, `--search`
- **Auto-pagination** — `--all` with `--page-limit` and `--page-delay` controls
- **Dry-run mode** — preview API requests on stderr without executing them or logging in
- **Read-only mode** — block every write request
- **Retry with backoff** — 429 rate limits are retried for every request; timeouts and 5xx only for requests that are safe to repeat
- **`.env` file support** — auto-loads `.env` from the current directory (see [Environment Variables](#environment-variables))
- **Organization mode** — opt-in access to Teams, SharePoint, shared mailboxes, Groups
- **Graph API version** — switch between `v1.0` and `beta` with `--api-version`
- **Structured logging** — JSON log files with daily rotation
- **File-based token cache** — owner-only file permissions, shared safely between concurrent processes
- **Zero MCP dependency** — standalone CLI, talks directly to Microsoft Graph API

---

## Requirements

- **Node.js** >= 24 (LTS)
- A **Microsoft account** (work/school; personal accounts are untested with the built-in app)
- Optional: your own **Azure AD app registration** (see [Authentication](#authentication)). Without one the CLI uses a built-in Microsoft client ID.

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
# 1. Log in (device code flow; uses the built-in client ID unless you set your own)
ms365 auth login

# 2. Verify your session
ms365 auth status

# 3. Triage your inbox
ms365 mail triage

# 4. Check today's schedule
ms365 workflow standup

# 5. Send an email
ms365 mail send --to "colleague@example.com" --subject "Hello" --body "Sent from ms365-cli"
```

To use your own app registration or enable organization features, run `ms365 auth setup` first.

---

## Authentication

ms365-cli uses **MSAL** (Microsoft Authentication Library) with the **device code flow**.

### Client ID

By default the CLI signs in with the **Microsoft Graph Command Line Tools** public client, `14d82eec-204b-4c2f-b7e8-296a70dab67e` — a Microsoft first-party app (the one Microsoft Graph PowerShell uses). No app registration is needed.

Know the trade-offs:
- Microsoft controls this app ID; it could change or be retired.
- Tenants can block the app or require admin consent for it. When sign-in fails with `AADSTS65001`, `AADSTS50105`, `AADSTS700016`, `AADSTS90094`, `AADSTS53003`, `AADSTS7000112`, `AADSTS50020` or `AADSTS500208`, the CLI points you to `ms365 auth setup` option [2] (your own app).
- Personal Microsoft accounts are untested with the built-in app.

Set `MS365_CLI_CLIENT_ID` (or `MS365_MCP_CLIENT_ID`) to use your own app registration instead. Your app needs the delegated permissions listed below and "Allow public client flows" enabled.

### Setup wizard

```bash
ms365 auth setup
```

- **[1] Quick setup** — uses the built-in app; no Azure Portal step.
- **[2] Custom app** — asks for your app's client ID (a GUID; re-prompts on bad input) and prints `export MS365_CLI_CLIENT_ID=…` for you to set.
- Choosing organization features prints `export MS365_CLI_ORG_MODE=true`.
- If you are already logged in, "Use existing login?" keeps the current login. Unknown menu input re-prompts; if input ends mid-wizard, setup exits 1.

The wizard does not write any config file: set the printed exports in your shell (see [Environment Variables](#environment-variables)).

### Login

```bash
ms365 auth login
```

Prints a URL and code on stderr. Open the URL in your browser, enter the code, and sign in. Tokens are cached automatically.

```bash
ms365 auth status    # exit 0 when logged in; exit 1 with "Not logged in. Run: ms365 auth login"
ms365 auth logout    # removes the token cache and selected account
```

`auth status` also exits 1 with "Could not verify the login: …" when Graph or the network fails.

### Permissions

Scopes are requested per group. Organization groups are requested only with `--org-mode` / `MS365_CLI_ORG_MODE`.

| Group | Scopes | Org mode only |
|-------|--------|---------------|
| User profile | `User.Read` | |
| Mail, mail rules, mailbox settings / auto-replies | `Mail.ReadWrite`, `Mail.Send`, `MailboxSettings.ReadWrite` | |
| Calendar | `Calendars.ReadWrite` | |
| OneDrive / Excel | `Files.ReadWrite` | |
| To Do / Planner | `Tasks.ReadWrite` | |
| Contacts | `Contacts.ReadWrite` | |
| OneNote | `Notes.Read`, `Notes.Create` | |
| People search | `People.Read` | |
| Presence | `Presence.Read`, `Presence.ReadWrite` | |
| Teams and channel messages | `Team.ReadBasic.All`, `Channel.ReadBasic.All`, `TeamMember.Read.All`, `ChannelMessage.Read.All`, `ChannelMessage.Send` | yes |
| Chats | `Chat.Read`, `ChatMessage.Read`, `ChatMessage.Send` | yes |
| SharePoint | `Sites.Read.All` | yes |
| Shared mailboxes and calendars | `Mail.Read.Shared`, `Mail.Send.Shared`, `Calendars.Read.Shared` | yes |
| Directory: users and groups | `User.Read.All`, `Group.Read.All` | yes |

When the requested permissions change (for example after an upgrade, or when switching org mode on), commands fail with "Permissions changed or sign-in expired … run: ms365 auth login". Log in once more.

### Token storage

Tokens are stored only in a file; there is no OS keychain integration.

| File | Mode |
|------|------|
| `~/.config/ms365-cli/` (directory) | `0700` |
| `~/.config/ms365-cli/token-cache.json` | `0600` |
| `~/.config/ms365-cli/selected-account.json` | `0600` |

These modes are re-applied on every start. Several processes can share the cache: writes are atomic (temporary file, then rename) and the last writer wins; there is no lock file. The device-code message is printed to stderr only and never written to disk.

### Multiple accounts

```bash
ms365 auth accounts            # List cached accounts
ms365 auth select <accountId>  # Switch active account
ms365 auth remove <accountId>  # Remove an account
```

### Bring your own token

```bash
export MS365_CLI_TOKEN=eyJ0eXAiOiJKV1Qi...
```

An access token in `MS365_CLI_TOKEN` (or `MS365_MCP_OAUTH_TOKEN`) is used as-is instead of the token cache. Graph commands then do not read or create `~/.config/ms365-cli`, so they also work with a read-only or missing `HOME` (CI, containers).

---

## Global Options

| Option                   | Description                                                  |
|--------------------------|--------------------------------------------------------------|
| `-V`, `--version`        | Print version number                                         |
| `--org-mode`             | Enable Teams, SharePoint, shared mailboxes, Groups           |
| `--read-only`            | Block write requests (see [Read-only Mode](#read-only-mode)) |
| `--dry-run`              | Preview the API request on stderr without executing it       |
| `--api-version <ver>`    | Graph API version: `v1.0` (default) or `beta`               |
| `-q`, `--quiet`          | Suppress non-data output on stderr (confirmations, progress) |
| `-h`, `--help`           | Show help                                                    |

Global options must come **before** the subcommand. After it they are rejected with `unknown option` (exit 1), so an argument value such as `--body --dry-run` can never switch a mode.

```bash
ms365 mail list                          # Personal mode
ms365 --org-mode teams list              # Organization mode
ms365 --api-version beta user me         # Beta API
ms365 --dry-run mail send --to "a@b.com" --subject "x" --body "y"   # Preview
ms365 -q mail send --to "a@b.com" --subject "x" --body "y"     # No "Email sent successfully." on stderr
```

---

## Output Formats

Data commands support `-o` / `--output`. Data goes to stdout; confirmations, progress and errors go to stderr.

| Format  | Flag          | Description                                     |
|---------|---------------|-------------------------------------------------|
| Table   | `-o table`    | ASCII table with nested object flattening (default on a terminal) |
| JSON    | `-o json`     | Pretty-printed JSON (default when stdout is piped or redirected) |
| CSV     | `-o csv`      | Comma-separated values with header row          |
| YAML    | `-o yaml`     | YAML format                                     |
| Text    | `-o text`     | Human-readable key-value pairs                  |

The default depends on stdout: **table** when it is a terminal, **JSON** when it is a pipe or file, so `ms365 mail list | jq` just works. A few human-oriented commands have a fixed default regardless of stdout: `mail triage` (table), `drive quota`, `user presence` and the `workflow` reports `standup`, `meeting-prep`, `digest`, `focus-time` (text). Pass `-o json` in scripts.

Table and CSV flatten nested objects into dot-notation columns:

```
from.emailAddress.name  from.emailAddress.address  hasAttachments  flag.flagStatus
----------------------  -------------------------  --------------  ---------------
John Doe                john@acme.com              true            flagged
Jane Smith              jane@acme.com              false           notFlagged
```

Commands that write files (`--out`) never overwrite an existing file unless you pass `--force`. Without `--out`, `drive download`, `download-shared`, `cat` and `onenote page` stream the raw bytes to stdout.

---

## Commands

### auth

7 commands.

```bash
ms365 auth setup                           # Interactive setup wizard
ms365 auth login                           # Device code login
ms365 auth status                          # Show the signed-in user (exit 1 if not logged in)
ms365 auth logout
ms365 auth accounts
ms365 auth select <accountId>
ms365 auth remove <accountId>
```

---

### mail

36 commands for Outlook mail.

#### Triage and reading

```bash
ms365 mail triage                          # Unread inbox summary (table)
ms365 mail triage --top 30 --include-read  # Include read messages
ms365 mail triage --folder "Projects"      # Another folder (name or ID)
ms365 mail read <messageId>                # Full message with body, attachments, flags
ms365 mail read <messageId> --text         # Plain text body
ms365 mail list --top 20 -o table          # Messages across all folders
ms365 mail list --folder inbox             # Inbox only
ms365 mail list --folder "Sent Items"      # By folder name
ms365 mail list --folder drafts            # Well-known folder names
ms365 mail list --search "quarterly report"
ms365 mail get <messageId>                 # Raw API response
```

- `mail list` lists messages from **all folders** (`/me/messages`); use `--folder inbox` for the inbox only.
- `mail triage` defaults to the **inbox**. Its `--filter` (default `isRead eq false`) is combined with a `receivedDateTime` condition so Graph accepts the sort order; `--include-read` drops only the unread condition and keeps your `--filter`. It also takes `--all`, `--page-limit` and `--page-delay`.
- `--folder` (on `list`, `triage`, `shared-list`, `move`, `batch-move`), `folders --parent`, `create-folder --parent` and `create-rule --move-to` accept a well-known name (`inbox`, `drafts`, `sentitems`, `archive`, …), a folder display name (exact match, case-insensitive, nested folders included) or a folder ID. An unknown name, or a name that matches two folders at the same depth, is an error.
- `mail read -o json` / `-o yaml` print the structured message including an `attachments` array; `table`, `text` and `csv` print the human view.

#### Sending and replying

```bash
ms365 mail send --to "a@b.com" --subject "Hi" --body "Hello"
ms365 mail send --to "a@b.com,b@c.com" --cc "d@e.com" --subject "Update" --body "<h1>HTML</h1>" --html
ms365 mail reply <messageId> --body "Thanks!"
ms365 mail reply <messageId> --body "<b>Noted.</b>" --html --reply-all
ms365 mail forward <messageId> --to "boss@co.com" --comment "FYI"
```

Send, reply and forward requests are never retried after a timeout or server error (they might already have been delivered); the error says so.

#### Drafts

```bash
ms365 mail draft --to "a@b.com" --subject "WIP" --body "..."
ms365 mail drafts                          # List drafts
ms365 mail update-draft <id> --subject "Updated" --body "New content"
ms365 mail send-draft <id>                 # Send a draft
```

#### Attachments

```bash
ms365 mail attachments <messageId>                     # List attachments (metadata only, no content)
ms365 mail attachment <messageId> <attachmentId>       # Attachment metadata/content
ms365 mail add-attachment <messageId> --name "file.pdf" --content-type "application/pdf" --content <base64>
ms365 mail download-attachment <messageId> <attachmentId> --out ./downloads/
ms365 mail download-attachment <messageId> <attachmentId> --out invoice.pdf --force
ms365 mail delete-attachment <messageId> <attachmentId>  # Asks for confirmation
```

`attachments` returns `id`, `name`, `contentType`, `size`, `isInline` and `lastModifiedDateTime`; use `attachment` or `download-attachment` for the content. `download-attachment` saves the raw bytes. `--out` may be a file or an existing directory; without `--out` the file goes into the current directory. The sender's file name is reduced to its base name, so it cannot escape the target directory. Existing files are never overwritten without `--force`. Reference (link) attachments are refused.

#### Inbox management

```bash
ms365 mail mark-read <id>                  # Mark as read
ms365 mail mark-unread <id>                # Mark as unread
ms365 mail flag <id> --due 2026-03-25      # Follow-up flag (start: today in local time, or the due date if earlier)
ms365 mail flag <id> --due 2026-03-25 --start 2026-03-20
ms365 mail unflag <id>                     # Remove flag
ms365 mail complete <id>                   # Mark flag as complete
ms365 mail categorize <id> --categories "Red category,Blue category"
ms365 mail move <id> --folder "Archive"
ms365 mail delete <id>                     # Asks for confirmation
```

`flag --start`/`--due` are sent as midnight in your local time zone when Graph documents that zone, otherwise as midnight UTC.

#### Batch operations

```bash
ms365 mail batch-read "id1,id2,id3"               # Mark multiple as read
ms365 mail batch-move "id1,id2" --folder "Archive" # Move multiple
```

If any message fails, the batch exits 1 and prints each failed full ID with its error.

#### Folders

```bash
ms365 mail folders -o table                # List top-level folders with item counts
ms365 mail folders --parent Inbox          # Child folders (name or ID)
ms365 mail folder "Sent Items"             # Lookup folder by name
ms365 mail create-folder --name "Projects"
ms365 mail create-folder --name "Clients" --parent Inbox
```

#### Inbox rules

```bash
ms365 mail rules                           # List rules
ms365 mail rule <ruleId>                   # Get rule details
ms365 mail create-rule --name "Auto-archive" --from "news@co.com" --move-to Archive --mark-read
ms365 mail create-rule --name "Disabled draft rule" --subject-contains "[spam]" --delete --no-enabled
ms365 mail update-rule <ruleId> --enabled false   # Disable a rule
```

- `create-rule` needs at least one action (`--move-to <name>`, `--forward-to`, `--delete`, `--mark-read`, `--importance low|normal|high`, `--stop-processing`). `--from` matches exact sender addresses. `--sequence <n>` sets the order (default: after the last rule). Rules are enabled unless you pass `--no-enabled`.
- `--move-to` takes a folder name or ID; well-known names are converted to the real folder ID. Under `--dry-run` the name is shown unchanged.
- A rule with **no conditions** applies to every incoming message; creating one asks for interactive confirmation.
- `update-rule --enabled` takes `true/false`, `yes/no`, `on/off` or `1/0`.

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
ms365 cal view --start 2026-03-17T00:00:00 --end 2026-03-23T23:59:59 --all
ms365 cal get <eventId>

# Create / update / delete
ms365 cal create --subject "Standup" --start "2026-03-17T09:00:00" --end "2026-03-17T09:30:00" --timezone "Europe/Berlin" --attendees "a@b.com"
ms365 cal update <eventId> --subject "Renamed" --location "Room B"
ms365 cal delete <eventId>

# Respond to invites
ms365 cal accept <eventId> --comment "See you there"
ms365 cal decline <eventId>
ms365 cal decline <eventId> --propose-start "2026-03-18T10:00:00" --propose-end "2026-03-18T10:30:00"
ms365 cal tentative <eventId>

# Scheduling
ms365 cal free-busy --emails "alice@co.com,bob@co.com" --start "2026-03-17T08:00:00" --end "2026-03-17T18:00:00"
ms365 cal rooms
ms365 cal room-lists

# Calendar management
ms365 cal list
ms365 cal create-calendar --name "Side Project" --color lightGreen
ms365 cal update-calendar <calendarId> --name "Renamed"
```

- Event reads return times in your local time zone; day and week boundaries are local. If the runtime cannot name the zone (a POSIX `TZ` such as `JST-9`, or an empty `TZ=`), UTC is used.
- Every `--timezone` (`create`, `update`, `free-busy`, `decline`) defaults to the local time zone. A wall-clock `--start`/`--end` is sent with the local IANA zone when Graph documents it; otherwise (e.g. `Europe/Amsterdam`, or a POSIX `TZ` such as `JST-9`) it is converted to the UTC instant and sent as UTC. An explicit `--timezone` is sent as given. A time with `Z` or an offset is sent as UTC.
- `decline --propose-start`/`--propose-end` must be given together and cannot be combined with `--no-notify`.
- `free-busy --interval` is an integer from 5 to 1440 minutes (default 30).
- `create-calendar`/`update-calendar --color` take `auto`, `lightBlue`, `lightGreen`, `lightOrange`, `lightGray`, `lightYellow`, `lightTeal`, `lightPink`, `lightBrown`, `lightRed` or `maxColor`.
- `update` / `update-calendar` with no fields fail with "Nothing to update" (exit 1). `update --timezone` needs `--start` or `--end`.
- `rooms` / `room-lists` use the Graph **beta** endpoints `/me/findRooms` and `/me/findRoomLists` (work/school accounts).
- `list` and `view` take `--top`/`--all`/`--page-limit`/`--page-delay`.

---

### drive

16 commands for OneDrive.

```bash
ms365 drive list -o table                  # List drives
ms365 drive root <driveId>                 # Root folder of a drive
ms365 drive files <driveId> <folderId>     # List files in folder
ms365 drive search "quarterly report"      # Search files (default top 25)
ms365 drive info <driveId> <itemId>        # File/folder metadata
ms365 drive quota                          # Storage usage

# Path-based access to your OneDrive
ms365 drive ls /                           # List the root
ms365 drive ls "/Documents/Reports"
ms365 drive cat "/Documents/notes.md"      # Raw bytes to stdout
ms365 drive cat "/Documents/report.pdf" --out ./report.pdf
ms365 drive cp ./report.pdf /Documents/    # Trailing / = folder; local file name appended
ms365 drive cp ./report.pdf "/Documents/Q3 report#v2.pdf" --conflict fail

# Upload / download by ID
ms365 drive upload <driveId> <parentId> ./report.xlsx   # Auto-detects MIME, streams large files
ms365 drive upload <driveId> <parentId> ./data.bin --content-type "application/gzip" --conflict replace
ms365 drive download <driveId> <itemId> --out ./report.pdf
ms365 drive download-shared <shareId> > shared.bin

# Organize
ms365 drive mkdir <driveId> <parentId> "New Folder"
ms365 drive delete <driveId> <itemId>

# Sharing
ms365 drive share <driveId> <itemId> --type edit --scope organization
ms365 drive share <driveId> <itemId> --type view --password "secret" --expiry "2026-04-01"
ms365 drive permissions <driveId> <itemId>
```

- `download`, `download-shared` and `cat` write to `--out <path>`, or stream raw bytes to stdout without it. An existing `--out` file is never overwritten without `--force` (exit 1).
- `--conflict rename|replace|fail` decides what happens when the target name exists: `upload` defaults to `rename`, `cp` to `replace`. The output and the stderr notice show the name OneDrive actually stored.
- Paths may contain spaces, `#`, `?` and `%`.
- `share --type` is `view|edit|embed` (default `view`); `--scope` is `anonymous|organization|users` (default `organization`).
- `permissions` returns every page as one result.
- Files up to 4 MiB are sent in one request; larger files use a resumable upload session.

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

`--importance` is `low|normal|high`; `--status` is `notStarted|inProgress|completed|waitingOnOthers|deferred`. `update` with no fields fails with "Nothing to update". `--due` is sent as midnight in your local time zone when Graph documents that zone, otherwise as midnight UTC.

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

- `list` and `channels` return every page as one result; `members` takes `--top`/`--all`/`--page-limit`/`--page-delay`.
- `send-channel-message`, `reply-channel-message`, `chat send` and `chat reply` print the created message on stdout (use `-o json` to get its `id`); the confirmation goes to stderr.

---

### chat

5 commands. Requires `--org-mode`.

```bash
ms365 --org-mode chat list
ms365 --org-mode chat get <chatId>
ms365 --org-mode chat messages <chatId> --top 20
ms365 --org-mode chat send <chatId> --body "Quick question..."
ms365 --org-mode chat reply <chatId> <messageId> --body "Got it."
```

Chats have no threaded replies: `chat reply` posts a new message that quotes the original.

---

### sharepoint

11 commands. Alias: `sp`. Requires `--org-mode`.

```bash
ms365 --org-mode sp search "Marketing"
ms365 --org-mode sp site <siteId>
ms365 --org-mode sp site-by-path contoso.sharepoint.com /sites/marketing
ms365 --org-mode sp drives <siteId>
ms365 --org-mode sp drive <siteId> <driveId>
ms365 --org-mode sp lists <siteId> -o table
ms365 --org-mode sp list <siteId> <listId>
ms365 --org-mode sp list-items <siteId> <listId> --expand "fields" --all
ms365 --org-mode sp list-item <siteId> <listId> <itemId>
ms365 --org-mode sp item <siteId> <itemId>  # One item by ID (Graph cannot list site items)
ms365 --org-mode sp delta                  # Track site changes
ms365 --org-mode sp delta --delta-link "<nextLink or deltaLink>"
```

`sp delta` prints `{ value, nextLink }` or `{ value, deltaLink }`. Pass `nextLink` back with `--delta-link` to continue, or a saved `deltaLink` to get later changes; `--all` drains the whole round. Only `https://graph.microsoft.com/…/sites/delta` links (with or without `()`) are accepted.

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

- Updates and deletes fetch the current ETag automatically when `--etag` is omitted.
- `--due` is that day in your local time zone. `--percent` accepts `0`, `50` or `100`.
- `create-plan --group` sets the group that owns the plan.
- Update commands with no field fail with "Nothing to update".
- `plans`, `my-tasks`, `tasks`, `buckets` and `bucket-tasks` return every page as one result.

---

### onenote

5 commands.

```bash
ms365 onenote notebooks -o table
ms365 onenote sections <notebookId>
ms365 onenote pages <sectionId>
ms365 onenote page <pageId>                # Raw HTML to stdout
ms365 onenote page <pageId> --out page.html
ms365 onenote create-page <sectionId> --title "Notes" --content "<p>Action items</p>"
```

---

### excel

5 commands for Excel workbooks stored in OneDrive.

```bash
ms365 excel worksheets <driveId> <itemId>
ms365 excel range <driveId> <itemId> <worksheetId> "A1:D10"
ms365 excel create-chart <driveId> <itemId> <worksheetId> --type ColumnClustered --source "A1:D10" --series-by columns
ms365 excel format-range <driveId> <itemId> <worksheetId> "A1:D1" --font-bold --font-color "#FF0000"
ms365 excel format-range <driveId> <itemId> <worksheetId> "A2:D2" --fill-color "#FFFF00"
ms365 excel sort-range <driveId> <itemId> <worksheetId> "A1:D10" --fields '[{"key":0,"ascending":true}]'
```

- `format-range` needs at least one of `--font-bold [true|false]` (flag alone = true), `--font-color`, `--fill-color`. `sort-range --fields` must be a JSON array.
- `create-chart --type` is one of the 74 `Excel.ChartType` values of ExcelApi 1.1 (e.g. `ColumnClustered`, `3DPie`, `XYScatterSmooth`); `--series-by` is `auto|columns|rows`.
- `worksheets` returns every page as one result.

---

### search

2 commands.

```bash
ms365 search query "quarterly report"
ms365 search query "budget" --entity message --top 5
ms365 search query "alpha" --entity "driveItem,site"
ms365 search people "John"
```

`--entity` is a comma-separated list of `driveItem`, `drive`, `list`, `listItem`, `site`, `externalItem`, `message`, `event`, `chatMessage`, `person`, `acronym`, `bookmark`, `qna`. Other values exit 1. Types can only be combined within one group:

| Group | Types |
|-------|-------|
| Files and sites | `driveItem`, `drive`, `list`, `listItem`, `site`, `externalItem` |
| Mail | `message` |
| Calendar | `event` |
| Teams | `chatMessage` |
| People | `person` |
| Answers | `acronym`, `bookmark`, `qna` |

Other combinations (e.g. `driveItem,message`) are rejected before any request. `--top` is 1–1000.

---

### user

14 commands for profile, presence, OOO, and settings.

```bash
ms365 user me
ms365 --org-mode user list --filter "department eq 'Engineering'" -o table

# Out of Office
ms365 user ooo                             # Get current auto-reply settings
ms365 user set-ooo --status scheduled --internal "I'm OOO until Monday" --start "2026-03-20T00:00:00" --end "2026-03-24T00:00:00"
ms365 user set-ooo --status alwaysEnabled --external "Away" --audience contactsOnly
ms365 user set-ooo --status disabled       # Turn off

# Presence
ms365 user presence                        # Get current status
ms365 user set-presence --availability DoNotDisturb --duration PT8H
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

- `set-ooo --status` is `alwaysEnabled|scheduled|disabled`. `--start`/`--end` are ISO 8601, local time unless they carry `Z` or an offset; `scheduled` needs both, other statuses reject them. `--audience` (`all|contactsOnly|none`) is left unchanged when omitted.
- `set-presence` sets your **preferred** presence, which shows only while you have a Teams (or other app) presence session. `--availability` is one of `Available`, `Busy`, `DoNotDisturb`, `BeRightBack`, `Away`, `Offline`; `--activity` is optional and derived (the only valid pairs are the same name, and `Offline`/`OffWork`). `--duration` is an ISO 8601 duration (default `PT60M`).
- `download-photo --out` defaults to `profile-photo.<ext>` from the content type and never overwrites without `--force`.
- `upload-photo` accepts `.jpg`, `.jpeg`, `.png` and `.gif` files (case-insensitive); other files exit 1 before any upload.

---

### groups

7 commands. Requires `--org-mode`.

```bash
ms365 --org-mode groups list -o table
ms365 --org-mode groups list --search "Marketing"
ms365 --org-mode groups my                 # Groups you belong to
ms365 --org-mode groups get <groupId>
ms365 --org-mode groups members <groupId>
ms365 --org-mode groups owners <groupId>
ms365 --org-mode groups distribution-lists # Mail-enabled DLs (excludes Microsoft 365 groups)
ms365 --org-mode groups teams              # Groups with Teams
```

`owners` takes `--top`/`--all`/`--page-limit`/`--page-delay`. `distribution-lists` drops Microsoft 365 (Unified) groups on the client, so a page can hold fewer than `--top` items.

---

### workflow

6 cross-service productivity commands. Alias: `wf`.

```bash
ms365 wf standup                           # Today's calendar + inbox unread count
ms365 wf meeting-prep                      # Next meeting: attendees, agenda, join link
ms365 wf email-to-task <messageId>         # Convert email to a task in your default To Do list
ms365 wf digest                            # Weekly summary
ms365 wf focus-time --work-start 08:30 --work-end 17:00
ms365 --org-mode wf file-announce <driveId> <itemId> --team <id> --channel <id>
```

- `standup` and `digest` report the inbox `unreadItemCount`; in JSON `unreadEmails` is a number.
- `digest` JSON: `week: { start, end }` are local dates (Monday–Sunday); `tasksDueThisWeek` counts open tasks due today through Sunday, and overdue tasks (`overdueTasks[].due`, `YYYY-MM-DD`) are listed separately. All To Do lists are scanned.
- `meeting-prep` picks the next timed, non-cancelled meeting starting within 24 hours. With none it prints a message on stderr (and `null` with `-o json`).
- `email-to-task` uses the default To Do list unless `--list` is given; `--importance` is `low|normal|high`. `--due` is midnight in the local time zone (UTC when Graph does not document the zone). Under `--dry-run` the preview uses `<default-list>` and the title `[Email] <subject>`.
- `file-announce` posts an HTML message (`--message` is plain text and is escaped) and prints the posted message.
- `focus-time --work-start`/`--work-end` take `HH:MM` (24 h; defaults 09:00/17:00).

---

### schema

Inspect the Graph requests behind a command before calling it. Runs without login. Every command that calls Microsoft Graph has an entry, keyed `<group>.<command>` (aliases work, e.g. `cal.view`); `auth`, `schema`, `generate-skills` and `serve` have none.

```bash
ms365 schema --list                        # JSON array of {endpoint, method, path, description}
ms365 schema --list -o table
ms365 schema mail.send                     # Parameters, scopes, response shape (JSON)
ms365 schema calendar.create -o yaml
ms365 schema drive.upload
```

- `--list` prints one row per Graph request, so a command that makes several requests (lookups, ETag GET then PATCH/DELETE, workflows) has several rows.
- 35 commonly used endpoints include parameters, scopes and response shape. The others show `httpMethod`, `path` and `description`, or `description` plus `requests: [{httpMethod, path}, …]` for multi-request commands.
- An unknown endpoint exits 1 with "Did you mean: …" and "Run: ms365 schema --list".

---

### generate-skills

Generate SKILL.md files for AI agent integration. Runs without login.

```bash
ms365 generate-skills                      # All skills to ./skills/
ms365 generate-skills --output-dir ./my-skills
ms365 generate-skills --filter "mail|calendar"   # Regex on the group name
```

The skills are generated from the registered commands, so new commands appear automatically. It writes `ms365-<group>/SKILL.md` for every top-level group (with every command's usage, arguments and options), `ms365-shared/SKILL.md` and `INDEX.md`. Progress goes to stderr (hidden by `-q`); nothing goes to stdout.

---

## OData Query Parameters

Many list commands support (check `ms365 <group> <command> --help`):

| Flag              | OData Parameter | Example                                      |
|-------------------|-----------------|----------------------------------------------|
| `--top <n>`       | `$top`          | `--top 25`                                   |
| `--filter <expr>` | `$filter`       | `--filter "isRead eq false"`                 |
| `--select <fields>` | `$select`     | `--select "subject,from,receivedDateTime"`   |
| `--orderby <field>` | `$orderby`   | `--orderby "receivedDateTime desc"`          |
| `--search <query>` | `$search`      | `--search "quarterly report"`                |

`--search` values (`mail list`, `contacts list`, `groups list`, `search people`) are sent as one quoted phrase, so quotes in your query are escaped rather than breaking the request.

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

Default page limit: **100 pages**. Default delay: **0ms**. `--top`, `--page-limit` and `--page-delay` must be integers. With `--all -o json` the output is one JSON array.

---

## Dry Run

Preview API requests without executing them:

```bash
ms365 --dry-run mail list --top 5 --filter "isRead eq false"
# stderr: [dry-run] GET https://graph.microsoft.com/v1.0/me/messages?$filter=isRead%20eq%20false&$top=5
# stdout: { "dryRun": true, "method": "GET", "url": "…" }

ms365 --dry-run mail send --to "a@b.com" --subject "Test" --body "Hello"
# stderr: [dry-run] POST https://graph.microsoft.com/v1.0/me/sendMail
# stderr: [dry-run] Body: {"message":{"subject":"Test",...}}
# (no "Email sent successfully." line)
```

- The `[dry-run]` preview lines always go to **stderr**, even with `--quiet`. Commands with `-o` print a `{ dryRun, method, url, body }` object on stdout, so `-o json` output stays parseable.
- Dry run skips authentication: it works without logging in (including Planner updates and deletes, which preview `If-Match: <dry-run>` instead of fetching the ETag).
- Write commands print only the preview under `--dry-run`, never a success line such as `Email sent successfully.`; `mail batch-read`/`batch-move` print no summary.

---

## Read-only Mode

`--read-only` (or `MS365_CLI_READ_ONLY=true`) blocks every request except `GET`/`HEAD` and three read-only `POST` actions: Microsoft Search (`/search/query`), free/busy (`getSchedule`) and `findMeetingTimes`. A blocked command exits 1 with "Read-only mode: … requests are blocked" before anything is sent.

---

## Error Handling and Exit Codes

| Exit code | Meaning |
|-----------|---------|
| `0` | Success (including `--help`, `--version`, and dry runs) |
| `1` | Any failure: Graph/network errors, usage errors (unknown option, missing argument, invalid choice), not logged in, read-only block, declined or non-interactive confirmation ("Nothing changed") |

Errors print one line, `Error: …`, on stderr (no stack traces).

- **401** — the token is refreshed once and the request repeated; if it still fails: "Authentication failed. Run: ms365 auth login".
- **403** — "Permission denied (403)"; suggests `--org-mode` when the message mentions scopes or permissions.
- **429** — retried for every request (including `POST`/`PATCH`), honouring `Retry-After`, otherwise with exponential backoff (1 s, 2 s).
- **5xx and timeouts** (30 s; 120 s for uploads) — retried only for requests that are safe to repeat: `GET`, `HEAD`, `PUT`, `DELETE` (a `drive upload`/`cp` with `--conflict rename` is not retried, because a retry could store a second copy). A `POST`/`PATCH` that times out or gets a 5xx fails with "… may have been applied; not retried" so you can check before re-running — for example, a mail send is never sent twice by the CLI.
- At most **3 attempts** per request (the first try plus 2 retries).

### Destructive operations

Deletes (`mail delete`, `mail delete-attachment`, `drive delete`, `todo delete`, `contacts delete`, `cal delete`, Planner deletes, …) and catch-all inbox rules ask for confirmation on an interactive terminal. There is no `--yes` flag: without a terminal (scripts, agents) they exit 1 and change nothing.

---

## Remote Access (Tailscale)

Agents on other machines in your tailnet can use one Microsoft 365 login without installing a token anywhere else. One host runs `ms365 serve`. Agents run `ms365-remote`, which accepts the same commands as `ms365`.

```
Agent machine                             ms365 host
ms365-remote mail list ──HTTPS──► tailscale serve ──unix socket──► ms365 serve ──► ms365 mail list ──► Microsoft Graph
                                  (adds Tailscale-App-Capabilities)
```

Microsoft login, refresh tokens and the app registration exist only on the host. Agents authenticate to the service with their Tailscale identity, and the tailnet policy decides which commands each agent may run.

### On the host

```bash
ms365 auth login                                   # once, on the host
ms365 serve --app-cap example.com/cap/ms365        # foreground; run it under systemd/launchd
tailscale serve --bg --accept-app-caps example.com/cap/ms365 unix:$HOME/.config/ms365-cli/serve.sock
```

`ms365 [--org-mode] [--read-only] [--api-version <v>] serve --app-cap <name> [--socket <path>] [--port <n> [--host <loopback>]] [--audit-log <path>] [--timeout <seconds>]`

- `--app-cap` is required and must equal the `--accept-app-caps` value. Use a capability name under a domain you control.
- The default socket is `~/.config/ms365-cli/serve.sock`, mode 0600 in a 0700 directory, so other Unix users cannot connect. A stale socket is replaced; a socket that is still served is refused (exit 1).
- `--org-mode`, `--read-only` and `--api-version` given before `serve` apply to every request.
- On SIGTERM/SIGINT the server stops accepting, finishes running requests, removes the socket and exits 0.
- Serving a `unix:` target needs a Tailscale version that supports it. Fallback for older versions: `ms365 serve --app-cap example.com/cap/ms365 --port 8365` with `tailscale serve --bg --accept-app-caps example.com/cap/ms365 http://127.0.0.1:8365`. In TCP mode every local process on the host can connect and forge the capability header, and serve prints a warning. Non-loopback `--host` values are refused.

Serve does not use Tailscale identity headers (`Tailscale-User-Login`, …): Tailscale does not set them for tagged devices. Access comes only from the `Tailscale-App-Capabilities` header, which `tailscale serve` sets and strips from incoming requests.

### Grants

In the tailnet policy, grant each agent tag the capabilities it needs:

```json
{
  "grants": [
    {
      "src": ["tag:agent-chief"],
      "dst": ["tag:ms365-host"],
      "ip":  ["tcp:443"],
      "app": { "example.com/cap/ms365": [{ "agent": "chief", "allow": ["mail.read", "mail.draft", "calendar.read", "todo.write"] }] }
    }
  ]
}
```

The `allow` lists of all grants that match a caller are combined. `agent` is the name written to the audit log. This grant syntax follows Tailscale's app-capability docs and has not been tested against a real tailnet.

Capabilities: `mail.read`, `mail.draft`, `mail.send`, `mail.manage`, `calendar.read`, `calendar.write`, `drive.read`, `drive.write`, `sharepoint.read`, `teams.read`, `teams.send` (teams and chat), `planner.read`, `planner.write`, `todo.read`, `todo.write`, `contacts.read`, `contacts.write`, `excel.read`, `excel.write`, `onenote.read`, `onenote.write`, `user.read`, `user.write`, `groups.read`, `search.read`, `meta` (schema). A workflow needs every capability it uses. For example, `workflow digest` needs `calendar.read`, `mail.read` and `todo.read`. `auth *`, `setup`, `serve` and `generate-skills` never run remotely. Deletes do not work remotely: they need an interactive confirmation, so they exit 1 and change nothing.

### On agent machines

```bash
export MS365_REMOTE_URL=https://ms365-host.<tailnet>.ts.net
ms365-remote mail list --top 5
ms365-remote --dry-run mail send --to a@b.com --subject Hi --body Hello
ms365-remote drive upload <driveId> <folderId> ./report.pdf     # the local file is uploaded
ms365-remote drive download <driveId> <itemId> --out ./copy.pdf  # saved locally
```

- Commands, help, usage errors and `-o` (TTY → table, pipe → json) behave as with `ms365`. Output is formatted locally.
- Local input files are uploaded. Downloads go to the local `--out`, or stream to stdout when `--out` is omitted. Existing files are not overwritten unless you pass `--force`.
- The only global flags are `--dry-run` and `-q`. Org mode, read-only mode and the API version come from the host.
- Errors print a message on stderr and exit 1. If neither transport variable is set, the CLI exits 1.
- No client ID or login is needed on the agent machine. `ms365-remote` reads real environment variables only and does not load `.env`.
- Interim mode without a server: `MS365_REMOTE_SSH=<host>` runs `tailscale ssh <host> -- ms365 <args>` with every argument quoted. Exit codes pass through, but file paths then refer to the host's disk.

### HTTP API

- `GET /v1/health` → `{ok: true, request_id}`
- `GET /v1/capabilities` → `{ok, request_id, agent, capabilities[], commands[]}` (the commands the caller may run)
- `POST /v1/execute` with JSON `{"command": "mail.list", "positionals": [], "args": {"top": 5, "dryRun": true}}`. `args` keys are option names in camelCase, and `dryRun` is the only global flag allowed per request. Send input files as `multipart/form-data`: a `request` JSON part, plus file parts referenced as `{"$file": "<part>"}` in `positionals`.
- Success: `{ok: true, request_id, data, files?, stderr?}`. `data` is the parsed JSON output, the raw string if the output is not JSON, or `null` when there is no output. `files` is `[{name, contentType, size, base64}]` for downloads.
- Errors: `{ok: false, request_id, error: {code, message}}`. Statuses: 400 `invalid_args`, 403 `capability_missing` (missing or bad header, capability not granted, command not remote), 404 `not_found`, 413 `invalid_args` (uploads over 25 MB), 502 `command_failed` (the CLI exited non-zero; the message is its stderr), 504 `timeout`, 500 `command_failed` (server error).
- Limits: 4 commands run at once and further requests queue; 120 s per command (`--timeout`); 50 MB stdout; 25 MB of uploaded files and 25 MB of returned files per request.

Each request runs the CLI as a child process in an empty temporary directory, which is deleted afterwards. The child receives only `PATH`, `HOME`, `TZ`, the log, org-mode and read-only variables, and the client/tenant ID. `MS365_CLI_TOKEN` is never passed on, so children use the host's token cache. The audit log (`~/.config/ms365-cli/serve-audit.jsonl`, 0600) has one line per request: `{ts, request_id, agent, src, method, path, command, args, status, duration_ms, error?}`. `args` lists option names only. Bodies and values are never logged.

---

## Environment Variables

| Variable                  | Description                                          | Default |
|---------------------------|------------------------------------------------------|----------|
| `MS365_CLI_CLIENT_ID`    | Azure AD app client ID                               | Built-in Microsoft Graph Command Line Tools ID |
| `MS365_CLI_TENANT_ID`    | Azure AD tenant ID                                   | `common` |
| `MS365_CLI_TOKEN`        | Pre-obtained OAuth2 access token (bypasses the token cache) | — |
| `MS365_CLI_ORG_MODE`     | Enable organization mode (`true` or `1`)             | off |
| `MS365_CLI_READ_ONLY`    | Enable read-only mode (`true` or `1`)                | off |
| `MS365_CLI_API_VERSION`  | Default for `--api-version` (`v1.0` or `beta`)       | `v1.0` |
| `MS365_CLI_LOG_LEVEL`    | Log level: `debug`, `info`, `warn`, `error`, `none`  | `none` |
| `MS365_CLI_LOG_FILE`     | JSON log file path (a date is appended: daily rotation) | — |
| `MS365_REMOTE_URL`       | `ms365-remote`: URL of the `ms365 serve` endpoint     | — |
| `MS365_REMOTE_SSH`       | `ms365-remote`: Tailscale SSH host (interim mode)     | — |

The token cache lives under `$HOME/.config/ms365-cli/`.

### `.env` files

The CLI loads `.env` from the **current working directory**, not from the install directory. A globally installed `ms365` therefore reads the `.env` of whatever project you run it in. Values already set in the environment win over `.env`.

- Set your real settings (client ID, tenant, org mode, read-only, and agent/remote settings) as **environment variables** in your shell profile or the agent's service environment, not in per-project `.env` files.
- Be aware that a project's `.env` containing `MS365_CLI_TOKEN` or `MS365_CLI_CLIENT_ID` silently switches the token or app for commands run in that directory.

---

## MCP Server Compatibility

If you're also using the [ms-365-mcp-server](https://github.com/softeria/ms-365-mcp-server), the CLI accepts its env var names as fallbacks:

| MCP Server Variable       | CLI Equivalent          |
|---------------------------|-------------------------|
| `MS365_MCP_CLIENT_ID`    | `MS365_CLI_CLIENT_ID`  |
| `MS365_MCP_TENANT_ID`    | `MS365_CLI_TENANT_ID`  |
| `MS365_MCP_OAUTH_TOKEN`  | `MS365_CLI_TOKEN`      |
| `MS365_MCP_ORG_MODE`     | `MS365_CLI_ORG_MODE`   |

The `MS365_CLI_*` name wins when both are set. The CLI keeps its own token cache; it does not share logins with the MCP server.

---

## Architecture

```
ms365-cli/
├── src/
│   ├── index.ts                Entry point — .env loading, global flags, exit code
│   ├── auth.ts                 MSAL device code auth, file token cache, scopes, multi-account
│   ├── graph-client.ts         Graph API client — retry policy, pagination, dry-run, read-only, uploads, binary downloads
│   ├── graph-helpers.ts        OData quoting, drive path encoding, time zone helpers
│   ├── list.ts                 Shared --top/--all/--page-limit/--page-delay handling
│   ├── files.ts                Safe file output (--out/--force, stdout streaming)
│   ├── command-tree.ts         Commander tree introspection (used by generate-skills)
│   ├── formatter.ts            Output formatting (JSON, table, CSV, YAML, text)
│   ├── errors.ts               Error reporting and the single process exit
│   ├── logger.ts               Structured JSON logging with daily rotation
│   ├── mime.ts                 MIME type detection (70+ extensions)
│   ├── remote.ts               ms365-remote: same commands, run on the ms365 serve host
│   ├── serve/                  ms365 serve: HTTP server, capability table, app-caps header, executor, audit log
│   ├── utils.ts                Output option, --quiet, confirmation prompt
│   └── commands/
│       ├── registry.ts         The one list of command registrations
│       ├── auth.ts             login, logout, status, accounts, select, remove (+ setup) (7)
│       ├── mail.ts             list, triage, read, send, reply, attachments, rules, batch (36)
│       ├── calendar.ts         events, view, create, respond, free-busy, rooms (15)
│       ├── drive.ts            files, ls/cat/cp, search, upload, download, share, mkdir, quota (16)
│       ├── todo.ts             lists, tasks, create/update (8)
│       ├── contacts.ts         CRUD + contact folders (9)
│       ├── teams.ts            teams, channels, messages (8) + chat (5)
│       ├── sharepoint.ts       sites, drives, lists, list items, delta (11)
│       ├── planner.ts          plans, tasks, buckets, details (21)
│       ├── onenote.ts          notebooks, sections, pages (5)
│       ├── excel.ts            worksheets, ranges, charts, format, sort (5)
│       ├── search.ts           query, people (2)
│       ├── user.ts             profile, OOO, presence, photo, settings (14)
│       ├── groups.ts           groups, members, DLs (7)
│       ├── workflow.ts         standup, meeting-prep, digest, focus-time (6)
│       ├── schema.ts           API endpoint introspection
│       ├── generate-skills.ts  SKILL.md generation for AI agents
│       ├── serve.ts            ms365 serve command (remote access over Tailscale)
│       └── setup.ts            Interactive auth setup wizard
├── test/                       Vitest unit, command and built-binary (e2e) tests
├── dist/                       Compiled output (ESM)
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
- **dotenv** — auto-loads `.env` from the working directory, compatible with MCP server env vars.
- **Streaming upload** — files <= 4 MiB use a simple PUT; larger files use resumable upload sessions (10 MiB chunks).
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

---

## Development

### Build

```bash
npm run build
```

### Test

```bash
npm test        # tsc --noEmit + vitest
```

### Dev mode

```bash
npm run dev -- mail list --top 5
```

### Adding a new command

1. Create `src/commands/myservice.ts`
2. Export `registerMyServiceCommands(program, getClient)`
3. Call it in `registerAllCommands` in `src/commands/registry.ts` (the CLI and `generate-skills` both use this list)
4. `npm run build`

---

## License

**Creative Commons Attribution-NonCommercial 4.0 International** (CC BY-NC 4.0)

Free to use, share, and adapt for **non-commercial purposes** with attribution.

See [LICENSE](./LICENSE) for full text: https://creativecommons.org/licenses/by-nc/4.0/
