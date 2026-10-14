# ms365-cli

[![CI](https://github.com/Cklaus1/ms365-cli/actions/workflows/ci.yml/badge.svg)](https://github.com/Cklaus1/ms365-cli/actions/workflows/ci.yml)
[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](./LICENSE)
[![Node.js](https://img.shields.io/badge/node-%3E%3D22.12-brightgreen.svg)](https://nodejs.org)
[![TypeScript](https://img.shields.io/badge/TypeScript-5.8-blue.svg)](https://www.typescriptlang.org/)
[![Microsoft Graph](https://img.shields.io/badge/Microsoft%20Graph-v1.0%20%7C%20beta-0078D4.svg)](https://learn.microsoft.com/en-us/graph/overview)

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
- [OData Query Parameters](#odata-query-parameters)
- [Pagination](#pagination)
- [Dry Run](#dry-run)
- [Error Handling](#error-handling)
- [Environment Variables](#environment-variables)
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
- **`.env` file support** — auto-loads from current directory
- **Organization mode** — opt-in access to Teams, SharePoint, shared mailboxes, Groups
- **Graph API version** — switch between `v1.0` and `beta` with `--api-version`
- **Structured logging** — JSON log files with daily rotation
- **Token storage** — a single owner-only (0600) file, `~/.config/ms365-cli/token-cache.json`, written atomically under a lock
- **Zero MCP dependency** — standalone CLI, talks directly to Microsoft Graph API

---

## Requirements

- **Node.js** >= 22.12 (Node 20 reached end-of-life 2026-04-30)
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

Tokens are stored in `~/.config/ms365-cli/token-cache.json` (mode 0600, directory 0700), written atomically under a lock so parallel commands are safe. Override the directory with `MS365_CLI_CONFIG_DIR`. Refreshed tokens are persisted automatically; `ms365 auth logout` deletes the file.

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

Global flags go before the command. The full list is in [All global flags](#all-global-flags).

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
| JSON    | `-o json`     | Pretty-printed JSON (default when piped)        |
| NDJSON  | `-o ndjson`   | One JSON object per line                        |
| Table   | `-o table`    | ASCII table (default on a terminal)             |
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

The reference below is generated from the live command tree (`ms365 describe --json`) by `npm run docs`; CI fails when it is stale. Run `ms365 <command> --help` for full option details, and `ms365 generate-skills` to write the same catalogue as SKILL.md files for AI agents.

<!-- BEGIN GENERATED COMMANDS -->
<!-- Generated by `npm run docs` from `ms365 describe --json`; do not edit by hand. -->

211 commands. Notes list the Graph scopes each command needs; **destructive** commands delete data, and most ask for confirmation, which needs a terminal. Options shared by list commands (`-o`, `--page-limit`, `--page-delay`) are left out of the tables; run `ms365 <command> --help` or `ms365 describe <command> --json` for the full list.

[audit](#audit) (2) · [auth](#auth) (12) · [calendar](#calendar) (16) · [chat](#chat) (6) · [completion](#completion) (1) · [contacts](#contacts) (9) · [describe](#describe) (1) · [drive](#drive) (16) · [excel](#excel) (8) · [generate-skills](#generate-skills) (1) · [groups](#groups) (7) · [mail](#mail) (36) · [mcp](#mcp) (1) · [meetings](#meetings) (7) · [notes](#notes) (5) · [onenote](#onenote) (5) · [planner](#planner) (21) · [resolve](#resolve) (1) · [schema](#schema) (1) · [search](#search) (2) · [sharepoint](#sharepoint) (12) · [teams](#teams) (8) · [todo](#todo) (12) · [user](#user) (14) · [watch](#watch) (1) · [workflow](#workflow) (6)

### All global flags

| Option | Description |
|---|---|
| `--account <upn\|homeAccountId>` | Use this cached account for this call only (env MS365_CLI_ACCOUNT) |
| `--actor <name>` | Name recorded as the actor in the audit log (default: MS365_CLI_ACTOR or "cli") |
| `--api-version <version>` | Graph API version |
| `--cursor <token>` | Fetch the page after a previous list call (the nextCursor it printed on stderr) |
| `--dry-run` | Preview the API request without executing it |
| `--error-format <format>` | Error output on stderr (default: text) |
| `--fields <list>` | Only output these comma-separated fields (dot paths allowed, e.g. id,from.emailAddress.address) |
| `--org-mode` | Enable organization mode (Teams, SharePoint, shared mailboxes) |
| `--profile <name>` | Apply a saved profile: client ID, tenant, scopes, org mode, account (env MS365_CLI_PROFILE) |
| `-q, --quiet` | Suppress non-data output (confirmations, progress) |
| `--read-only` | Restrict to read-only operations |
| `-V, --version` | output the version number |

### audit

| Command | Description | Arguments | Key options | Notes |
|---|---|---|---|---|
| `audit path` | Print the audit log path |  |  | offline |
| `audit tail` | Print the most recent audit entries |  | `-n, --lines <n>` `--follow` | offline |

### auth

| Command | Description | Arguments | Key options | Notes |
|---|---|---|---|---|
| `auth accounts` | List all cached accounts |  |  | offline |
| `auth login` | Login (device code flow by default) |  | `--browser` | not idempotent · User.Read |
| `auth logout` | Log out and clear saved credentials |  |  | **destructive** · offline |
| `auth profile delete` | Delete a saved profile | `<name>` |  | **destructive** · offline |
| `auth profile list` | List saved profiles |  |  | offline |
| `auth profile set` | Create a profile, or update the given fields of an existing one | `<name>` | `--client-id <id>` `--tenant-id <id>` `--scopes <list>` `--org-mode` `--no-org-mode` `--account <upn\|homeAccountId>` | offline |
| `auth profile show` | Show a saved profile | `<name>` |  | offline |
| `auth remove` | Remove a cached account | `<accountId>` |  | **destructive** · offline |
| `auth scopes` | Show the scopes granted to the current access token (plus audience, tenant, user, expiry) |  |  | offline |
| `auth select` | Select an account by ID | `<accountId>` |  | offline |
| `auth setup` | Interactive setup wizard for Azure AD app registration |  |  | not idempotent · User.Read |
| `auth status` | Verify current login status |  |  | User.Read |

### calendar

| Command | Description | Arguments | Key options | Notes |
|---|---|---|---|---|
| `calendar accept` | Accept a meeting invite | `<eventId>` | `--comment <text>` `--no-notify` | not idempotent · Calendars.ReadWrite |
| `calendar create` | Create a calendar event |  | `--subject <subject>` (required) `--start <datetime>` (required) `--end <datetime>` (required) `--timezone <tz>` `--body <body>` `--location <location>` `--attendees <emails>` `--calendar <calendarId>` `--html` `--teams` `--idempotency-key <key>` | not idempotent · Calendars.ReadWrite |
| `calendar create-calendar` | Create a secondary calendar |  | `--name <name>` (required) `--color <color>` | not idempotent · Calendars.ReadWrite |
| `calendar decline` | Decline a meeting invite | `<eventId>` | `--comment <text>` `--no-notify` `--propose-start <datetime>` `--propose-end <datetime>` `--timezone <tz>` | not idempotent · Calendars.ReadWrite |
| `calendar delete` | Delete a calendar event | `<eventId>` | `--calendar <calendarId>` | **destructive** · Calendars.ReadWrite |
| `calendar events` | List calendar events |  | `--top <n>` `--filter <filter>` `--select <fields>` `--orderby <field>` `--calendar <calendarId>` `--all` `--timezone <tz>` | Calendars.Read |
| `calendar find-time` | Suggest meeting times when the attendees are free (findMeetingTimes) |  | `--attendees <emails>` (required) `--duration <d>` (required) `--start <datetime>` `--end <datetime>` `--timezone <tz>` | Calendars.Read.Shared |
| `calendar free-busy` | Get free/busy schedule for one or more users |  | `--emails <addresses>` (required) `--start <datetime>` (required) `--end <datetime>` (required) `--timezone <tz>` `--interval <min>` | Calendars.Read |
| `calendar get` | Get a specific event | `<eventId>` | `--calendar <calendarId>` `--select <fields>` `--timezone <tz>` | Calendars.Read |
| `calendar list` | List calendars |  |  | Calendars.Read |
| `calendar room-lists` | List room lists / buildings (requires --org-mode) |  |  | Place.Read.All |
| `calendar rooms` | List meeting rooms (requires --org-mode) |  |  | Place.Read.All |
| `calendar tentative` | Tentatively accept a meeting invite | `<eventId>` | `--comment <text>` `--no-notify` | not idempotent · Calendars.ReadWrite |
| `calendar update` | Update a calendar event | `<eventId>` | `--subject <subject>` `--start <datetime>` `--end <datetime>` `--timezone <tz>` `--body <body>` `--location <location>` `--calendar <calendarId>` | Calendars.ReadWrite |
| `calendar update-calendar` | Update a calendar | `<calendarId>` | `--name <name>` `--color <color>` | Calendars.ReadWrite |
| `calendar view` | Get calendar view (time range) |  | `--start <datetime>` (required) `--end <datetime>` (required) `--select <fields>` `--all` `--timezone <tz>` | Calendars.Read |

### chat

| Command | Description | Arguments | Key options | Notes |
|---|---|---|---|---|
| `chat create` | Create a 1:1 chat (one member) or group chat (two or more) |  | `--members <emails>` (required) `--topic <topic>` | not idempotent · Chat.Create, User.Read |
| `chat get` | Get chat details | `<chatId>` |  | Chat.Read |
| `chat list` | List chats |  | `--top <n>` `--all` | Chat.Read |
| `chat messages` | List chat messages (marked untrusted; see --body, --wrap-untrusted) | `<chatId>` | `--top <n>` `--all` `--body <mode>` `--max-chars <n>` `--wrap-untrusted` | ChatMessage.Read |
| `chat reply` | Reply to a chat message | `<chatId> <messageId>` | `--body <body>` (required) `--html` | not idempotent · ChatMessage.Read, ChatMessage.Send |
| `chat send` | Send a chat message | `<chatId>` | `--body <body>` (required) `--html` `--mention <upn>` | not idempotent · ChatMessage.Send, User.ReadBasic.All |

### completion

| Command | Description | Arguments | Key options | Notes |
|---|---|---|---|---|
| `completion` | Print a shell completion script (bash, zsh, fish, powershell) | `<shell>` |  | offline |

### contacts

| Command | Description | Arguments | Key options | Notes |
|---|---|---|---|---|
| `contacts create` | Create a contact |  | `--given-name <name>` (required) `--surname <name>` `--email <email>` `--phone <phone>` `--company <company>` `--job-title <title>` | not idempotent · Contacts.ReadWrite |
| `contacts create-folder` | Create a contact folder |  | `--name <name>` (required) `--parent <folderId>` | not idempotent · Contacts.ReadWrite |
| `contacts delete` | Delete a contact | `<contactId>` |  | **destructive** · Contacts.ReadWrite |
| `contacts folder-contacts` | List contacts in a specific folder | `<folderId>` | `--top <n>` `--filter <filter>` `--orderby <field>` `--all` | Contacts.Read |
| `contacts folders` | List contact folders |  |  | Contacts.Read |
| `contacts get` | Get a specific contact | `<contactId>` |  | Contacts.Read |
| `contacts list` | List contacts |  | `--top <n>` `--filter <filter>` `--orderby <field>` `--select <fields>` `--search <query>` `--all` | Contacts.Read |
| `contacts update` | Update a contact | `<contactId>` | `--given-name <name>` `--surname <name>` `--email <email>` `--replace-email` `--phone <phone>` `--company <company>` | Contacts.ReadWrite |
| `contacts update-folder` | Rename a contact folder (use "[ARCHIVED] Name" to archive) | `<folderId>` | `--name <name>` (required) | Contacts.ReadWrite |

### describe

| Command | Description | Arguments | Key options | Notes |
|---|---|---|---|---|
| `describe` | Describe every command (arguments, options, Graph scopes, safety) as a machine-readable catalogue | `[command...]` | `--json` | offline |

### drive

| Command | Description | Arguments | Key options | Notes |
|---|---|---|---|---|
| `drive cat` | Download a file by path (e.g. "/Documents/notes.md") | `<path>` | `--out <file>` `--force` | Files.Read |
| `drive cp` | Upload a file by path (e.g. ./report.pdf /Documents/report.pdf) | `<localFile> <remotePath>` | `--content-type <mime>` | Files.ReadWrite |
| `drive delete` | Delete a drive item | `<driveId> <itemId>` |  | **destructive** · Files.ReadWrite |
| `drive download` | Download a file | `<driveId> <itemId> <childId>` | `--out <path>` `--force` | Files.Read |
| `drive download-shared` | Download a file via share ID or sharing link URL | `<shareId>` | `--out <path>` `--force` | Files.Read.All |
| `drive files` | List files in a folder | `<driveId> <folderId>` | `--top <n>` `--filter <filter>` `--select <fields>` `--all` | Files.Read |
| `drive info` | Get metadata for a file or folder | `<driveId> <itemId>` |  | Files.Read |
| `drive list` | List available drives |  |  | Files.Read |
| `drive ls` | List files by path (e.g. "/Documents", "/Documents/Reports") | `[path]` | `--top <n>` | Files.Read |
| `drive mkdir` | Create a folder | `<driveId> <parentId> <name>` |  | not idempotent · Files.ReadWrite |
| `drive permissions` | List sharing permissions on a file or folder | `<driveId> <itemId>` |  | Files.Read |
| `drive quota` | Show OneDrive storage usage and quota |  |  | Files.Read |
| `drive root` | Get root folder of a drive | `<driveId>` |  | Files.Read |
| `drive search` | Search files across OneDrive | `<query>` | `--drive <driveId>` `--top <n>` | Files.Read |
| `drive share` | Create a sharing link for a file or folder | `<driveId> <itemId>` | `--type <type>` `--scope <scope>` `--password-stdin` `--expiry <datetime>` | Files.ReadWrite |
| `drive upload` | Upload a file (auto-detects MIME type, streams large files) | `<driveId> <parentId> <filePath>` | `--content-type <mime>` | Files.ReadWrite |

### excel

| Command | Description | Arguments | Key options | Notes |
|---|---|---|---|---|
| `excel add-rows` | Append rows to a table (table name or ID) | `<driveId> <itemId> <table>` | `--values <json>` `--values-file <path>` | not idempotent · Files.ReadWrite |
| `excel create-chart` | Create a chart | `<driveId> <itemId> <worksheetId>` | `--type <type>` (required) `--source <range>` (required) `--series-by <by>` (required) | not idempotent · Files.ReadWrite |
| `excel format-range` | Format a cell range (e.g. "A1:D10") | `<driveId> <itemId> <worksheetId> <address>` | `--font-bold [bool]` `--font-color <color>` `--fill-color <color>` | Files.ReadWrite |
| `excel range` | Get a cell range (e.g. "A1:D10") | `<driveId> <itemId> <worksheetId> <address>` |  | Files.Read |
| `excel sort-range` | Sort a cell range (e.g. "A1:D10") | `<driveId> <itemId> <worksheetId> <address>` | `--fields <json>` (required) `--match-case` | Files.ReadWrite |
| `excel table-rows` | List rows of a table (table name or ID) | `<driveId> <itemId> <table>` | `--top <n>` | Files.Read |
| `excel tables` | List tables in a workbook | `<driveId> <itemId>` |  | Files.Read |
| `excel worksheets` | List worksheets in a workbook | `<driveId> <itemId>` |  | Files.Read |

### generate-skills

| Command | Description | Arguments | Key options | Notes |
|---|---|---|---|---|
| `generate-skills` | Generate SKILL.md files for AI agent integration |  | `--output-dir <dir>` `--filter <pattern>` `--force` | offline |

### groups

| Command | Description | Arguments | Key options | Notes |
|---|---|---|---|---|
| `groups distribution-lists` | List mail-enabled distribution groups |  | `--top <n>` `--all` | Group.Read.All |
| `groups get` | Get group details | `<groupId>` |  | Group.Read.All |
| `groups list` | List groups in the organization |  | `--top <n>` `--filter <filter>` `--orderby <field>` `--search <query>` `--select <fields>` `--all` | Group.Read.All |
| `groups members` | List group members | `<groupId>` | `--top <n>` `--all` | Group.Read.All |
| `groups my` | List groups you are a member of |  | `--top <n>` `--select <fields>` `--all` | Group.Read.All |
| `groups owners` | List group owners | `<groupId>` |  | Group.Read.All |
| `groups teams` | List groups that have a Team associated |  | `--top <n>` `--all` | Group.Read.All |

### mail

| Command | Description | Arguments | Key options | Notes |
|---|---|---|---|---|
| `mail add-attachment` | Add an attachment to a draft message | `<messageId>` | `--name <name>` `--content-type <type>` (required) `--content <base64>` `--file <path>` | not idempotent · Mail.ReadWrite |
| `mail attachment` | Get a specific attachment | `<messageId> <attachmentId>` |  | Mail.Read |
| `mail attachments` | List attachments for a message | `<messageId>` |  | Mail.Read |
| `mail batch-move` | Move multiple messages to a folder (comma-separated IDs; sent via JSON batching) | `<messageIds>` | `--folder <folderId>` (required) | not idempotent · Mail.ReadWrite |
| `mail batch-read` | Mark multiple messages as read (comma-separated IDs; sent via JSON batching) | `<messageIds>` |  | Mail.ReadWrite |
| `mail categorize` | Set categories on a message | `<messageId>` | `--categories <names>` (required) | Mail.ReadWrite |
| `mail complete` | Mark a flagged message as complete | `<messageId>` |  | Mail.ReadWrite |
| `mail create-folder` | Create a mail folder |  | `--name <name>` (required) `--parent <parentFolderId>` | not idempotent · Mail.ReadWrite |
| `mail create-rule` | Create an inbox rule |  | `--name <name>` (required) `--from <addresses>` `--subject-contains <words>` `--move-to <folderId>` `--forward-to <addresses>` `--delete` `--mark-read` `--importance <level>` `--stop-processing` `--all-messages` `--disabled` | not idempotent · MailboxSettings.ReadWrite |
| `mail delete` | Delete a message | `<messageId>` |  | **destructive** · Mail.ReadWrite |
| `mail delete-attachment` | Delete an attachment | `<messageId> <attachmentId>` |  | **destructive** · Mail.ReadWrite |
| `mail download-attachment` | Download an attachment to a file | `<messageId> <attachmentId>` | `--out <path>` `--force` | Mail.Read |
| `mail draft` | Create a draft email |  | `--to <addresses>` (required) `--subject <subject>` (required) `--body <body>` (required) `--cc <addresses>` `--html` `--attach <path>` | not idempotent · Mail.ReadWrite |
| `mail drafts` | List draft emails |  | `--top <n>` | Mail.Read |
| `mail flag` | Set follow-up flag on a message | `<messageId>` | `--due <when>` `--timezone <tz>` | Mail.ReadWrite |
| `mail folder` | Find a folder by display name (e.g. "Drafts", "Sent Items", "Archive") | `<name>` |  | Mail.Read |
| `mail folders` | List mail folders |  | `--parent <folderId>` `--all` | Mail.Read |
| `mail forward` | Forward a message | `<messageId>` | `--to <addresses>` (required) `--comment <text>` `--attach <path>` `--idempotency-key <key>` | not idempotent · Mail.Send |
| `mail get` | Get a specific message (marked untrusted; see --body, --unique, --wrap-untrusted) | `<messageId>` | `--select <fields>` `--unique` `--body <mode>` `--max-chars <n>` `--wrap-untrusted` | Mail.Read |
| `mail list` | List inbox messages |  | `--top <n>` `--filter <filter>` `--select <fields>` `--search <query>` `--folder <name>` `--all` `--body <mode>` `--max-chars <n>` `--wrap-untrusted` | Mail.Read |
| `mail mark-read` | Mark a message as read | `<messageId>` |  | Mail.ReadWrite |
| `mail mark-unread` | Mark a message as unread | `<messageId>` |  | Mail.ReadWrite |
| `mail move` | Move a message to a folder | `<messageId>` | `--folder <folderId>` (required) | not idempotent · Mail.ReadWrite |
| `mail read` | Read a message with full body, attachments info, and metadata | `<messageId>` | `--text` | Mail.Read |
| `mail reply` | Reply to a message | `<messageId>` | `--body <body>` (required) `--html` `--reply-all` `--attach <path>` `--idempotency-key <key>` | not idempotent · Mail.Send |
| `mail rule` | Get a specific inbox rule | `<ruleId>` |  | MailboxSettings.Read |
| `mail rules` | List inbox rules |  |  | MailboxSettings.Read |
| `mail send` | Send an email |  | `--to <addresses>` (required) `--cc <addresses>` `--bcc <addresses>` `--subject <subject>` (required) `--body <body>` (required) `--html` `--attach <path>` `--idempotency-key <key>` | not idempotent · Mail.Send |
| `mail send-draft` | Send an existing draft | `<messageId>` | `--idempotency-key <key>` | not idempotent · Mail.Send |
| `mail shared-get` | Get a shared mailbox message (requires --org-mode) | `<userId> <messageId>` | `--select <fields>` | Mail.Read.Shared |
| `mail shared-list` | List shared mailbox messages (requires --org-mode) | `<userId>` | `--top <n>` `--filter <filter>` `--select <fields>` `--folder <folderId>` `--all` | Mail.Read.Shared |
| `mail shared-send` | Send email from a shared mailbox (requires --org-mode) | `<userId>` | `--to <addresses>` (required) `--subject <subject>` (required) `--body <body>` (required) `--html` | not idempotent · Mail.Send.Shared |
| `mail triage` | Show inbox summary — unread messages with key metadata |  | `--top <n>` `--filter <filter>` `--folder <folderId>` `--include-read` `--max-chars <n>` `--wrap-untrusted` | Mail.Read |
| `mail unflag` | Remove follow-up flag from a message | `<messageId>` |  | Mail.ReadWrite |
| `mail update-draft` | Update an existing draft | `<messageId>` | `--to <addresses>` `--cc <addresses>` `--subject <subject>` `--body <body>` `--html` | Mail.ReadWrite |
| `mail update-rule` | Update an inbox rule | `<ruleId>` | `--name <name>` `--enabled <bool>` | MailboxSettings.ReadWrite |

### mcp

| Command | Description | Arguments | Key options | Notes |
|---|---|---|---|---|
| `mcp serve` | Serve every command as an MCP tool over stdio (JSON-RPC; each call runs the CLI) |  | `--allow <patterns>` `--read-only` | offline |

### meetings

| Command | Description | Arguments | Key options | Notes |
|---|---|---|---|---|
| `meetings all-transcripts` | List transcripts of all meetings you organized (beta API; needs --api-version beta) |  | `--start <datetime>` `--end <datetime>` `--timezone <tz>` `--all` | User.Read, OnlineMeetingTranscript.Read.All |
| `meetings get` | Get an online meeting by Teams join URL or id | `<joinUrlOrId>` |  | OnlineMeetings.Read |
| `meetings list` | List your online meetings in a time window (calendar events with a Teams link) |  | `--start <datetime>` `--end <datetime>` `--timezone <tz>` `--all` | Calendars.Read |
| `meetings recording` | Download a recording (default: the latest) to a file (<meeting>: online meeting id or Teams join URL) | `<meeting> [recordingId]` | `--out <path>` (required) `--force` | OnlineMeetings.Read, OnlineMeetingRecording.Read.All |
| `meetings recordings` | List a meeting's recordings (<meeting>: online meeting id or Teams join URL) | `<meeting>` |  | OnlineMeetings.Read, OnlineMeetingRecording.Read.All |
| `meetings transcript` | Download a transcript (default: the latest) as VTT or "Speaker: text" (<meeting>: online meeting id or Teams join URL) | `<meeting> [transcriptId]` | `--format <fmt>` `--out <path>` `--force` | OnlineMeetings.Read, OnlineMeetingTranscript.Read.All |
| `meetings transcripts` | List a meeting's transcripts (<meeting>: online meeting id or Teams join URL) | `<meeting>` |  | OnlineMeetings.Read, OnlineMeetingTranscript.Read.All |

### notes

| Command | Description | Arguments | Key options | Notes |
|---|---|---|---|---|
| `notes create` | Create a note |  | `--body <text>` (required) `--subject <subject>` | not idempotent · ShortNotes.ReadWrite |
| `notes delete` | Delete a note | `<id>` |  | **destructive** · ShortNotes.ReadWrite |
| `notes get` | Get a note | `<id>` |  | ShortNotes.Read |
| `notes list` | List notes |  | `--top <n>` `--all` | ShortNotes.Read |
| `notes update` | Update a note | `<id>` | `--body <text>` `--subject <subject>` | ShortNotes.ReadWrite |

### onenote

| Command | Description | Arguments | Key options | Notes |
|---|---|---|---|---|
| `onenote create-page` | Create a page in a section | `<sectionId>` | `--title <title>` (required) `--content <html>` (required) | not idempotent · Notes.Create |
| `onenote notebooks` | List notebooks |  |  | Notes.Read |
| `onenote page` | Get page content | `<pageId>` |  | Notes.Read |
| `onenote pages` | List pages in a section | `<sectionId>` |  | Notes.Read |
| `onenote sections` | List sections in a notebook | `<notebookId>` |  | Notes.Read |

### planner

| Command | Description | Arguments | Key options | Notes |
|---|---|---|---|---|
| `planner bucket` | Get bucket details | `<bucketId>` | `--plan <planId>` | Tasks.Read |
| `planner bucket-tasks` | List tasks in a bucket | `<bucketId>` | `--plan <planId>` | Tasks.Read |
| `planner buckets` | List plan buckets | `<planId>` |  | Tasks.Read |
| `planner create-bucket` | Create a bucket in a plan |  | `--plan <planId>` (required) `--name <name>` (required) | not idempotent · Tasks.ReadWrite |
| `planner create-plan` | Create a planner plan |  | `--title <title>` (required) `--group <groupId>` (required) | not idempotent · Tasks.ReadWrite |
| `planner create-task` | Create a planner task |  | `--plan <planId>` (required) `--title <title>` (required) `--bucket <bucketId>` `--due <date>` `--timezone <tz>` `--assigned-to <userId>` | not idempotent · Tasks.ReadWrite |
| `planner delete-bucket` | Delete a planner bucket | `<bucketId>` | `--plan <planId>` `--etag <etag>` | **destructive** · Tasks.ReadWrite |
| `planner delete-plan` | Delete a planner plan | `<planId>` | `--etag <etag>` | **destructive** · Tasks.ReadWrite |
| `planner delete-task` | Delete a planner task | `<taskId>` | `--etag <etag>` | **destructive** · Tasks.ReadWrite |
| `planner my-tasks` | List my planner tasks |  |  | Tasks.Read |
| `planner plan` | Get plan details | `<planId>` |  | Tasks.Read |
| `planner plan-details` | Get plan extended details (category labels, etc.) | `<planId>` |  | Tasks.Read |
| `planner plans` | List my planner plans |  |  | Tasks.Read |
| `planner task` | Get task details | `<taskId>` |  | Tasks.Read |
| `planner task-details` | Get task extended details (description, checklist, etc.) | `<taskId>` |  | Tasks.Read |
| `planner tasks` | List tasks in a plan | `<planId>` |  | Tasks.Read |
| `planner update-bucket` | Update a planner bucket | `<bucketId>` | `--plan <planId>` `--name <name>` `--etag <etag>` | Tasks.ReadWrite |
| `planner update-plan` | Update a planner plan | `<planId>` | `--title <title>` `--etag <etag>` | Tasks.ReadWrite |
| `planner update-plan-details` | Update plan details (category labels) | `<planId>` | `--etag <etag>` `--categories <json>` | Tasks.ReadWrite |
| `planner update-task` | Update a planner task | `<taskId>` | `--title <title>` `--due <date>` `--timezone <tz>` `--percent <n>` `--bucket <bucketId>` `--plan <planId>` `--etag <etag>` | Tasks.ReadWrite |
| `planner update-task-details` | Update task details (description, checklist) | `<taskId>` | `--etag <etag>` `--description <text>` | Tasks.ReadWrite |

### resolve

| Command | Description | Arguments | Key options | Notes |
|---|---|---|---|---|
| `resolve` | Resolve a name to its ID (exact, then unique prefix/substring match; mail-folder: exact name or path like Inbox/Projects) | `<type> <name>` | `--team <team>` `--plan <plan>` `--user <user>` | Tasks.Read, Team.ReadBasic.All, Channel.ReadBasic.All |

### schema

| Command | Description | Arguments | Key options | Notes |
|---|---|---|---|---|
| `schema` | Introspect API schema for an endpoint (e.g. mail.list, calendar.create) | `[endpoint]` | `--list` | offline |

### search

| Command | Description | Arguments | Key options | Notes |
|---|---|---|---|---|
| `search people` | Search people | `[query]` | `--top <n>` | People.Read |
| `search query` | Search mail, events, chats, files, sites and list items in one query (hits merged into one list) | `<query>` | `--entity <types>` `--size <n>` `--from <n>` | Files.Read.All, Sites.Read.All, Mail.Read, Calendars.Read, Chat.Read |

### sharepoint

| Command | Description | Arguments | Key options | Notes |
|---|---|---|---|---|
| `sharepoint delta` | Get changed sites (delta sync) |  |  | Sites.Read.All |
| `sharepoint drive` | Get a specific site drive | `<siteId> <driveId>` |  | Sites.Read.All |
| `sharepoint drives` | List site drives (document libraries) | `<siteId>` |  | Sites.Read.All |
| `sharepoint item` | Get a specific site item | `<siteId> <itemId>` |  | Sites.Read.All |
| `sharepoint items` | List site items | `<siteId>` | `--all` | Sites.Read.All |
| `sharepoint list` | Get a specific SharePoint list | `<siteId> <listId>` |  | Sites.Read.All |
| `sharepoint list-item` | Get a specific list item | `<siteId> <listId> <itemId>` |  | Sites.Read.All |
| `sharepoint list-items` | List items in a SharePoint list | `<siteId> <listId>` | `--top <n>` `--filter <filter>` `--expand <fields>` `--all` | Sites.Read.All |
| `sharepoint lists` | List site lists | `<siteId>` |  | Sites.Read.All |
| `sharepoint search` | Search SharePoint sites | `<query>` |  | Sites.Read.All |
| `sharepoint site` | Get site details | `<siteId>` |  | Sites.Read.All |
| `sharepoint site-by-path` | Get site by hostname and path (e.g. contoso.sharepoint.com /sites/marketing) | `<hostname> <path>` |  | Sites.Read.All |

### teams

| Command | Description | Arguments | Key options | Notes |
|---|---|---|---|---|
| `teams channel` | Get channel details | `<teamId> <channelId>` |  | Channel.ReadBasic.All |
| `teams channel-messages` | List channel messages (marked untrusted; see --body, --wrap-untrusted) | `<teamId> <channelId>` | `--top <n>` `--all` `--body <mode>` `--max-chars <n>` `--wrap-untrusted` | ChannelMessage.Read.All |
| `teams channels` | List team channels | `<teamId>` |  | Channel.ReadBasic.All |
| `teams get` | Get team details | `<teamId>` |  | Team.ReadBasic.All |
| `teams list` | List joined teams |  |  | Team.ReadBasic.All |
| `teams members` | List team members | `<teamId>` |  | TeamMember.Read.All |
| `teams reply-channel-message` | Reply to a channel message (threaded) | `<teamId> <channelId> <messageId>` | `--body <body>` (required) `--html` | not idempotent · ChannelMessage.Send |
| `teams send-channel-message` | Send a message to a channel | `<teamId> <channelId>` | `--body <body>` (required) `--html` `--mention <upn>` | not idempotent · ChannelMessage.Send, User.ReadBasic.All, Team.ReadBasic.All, Channel.ReadBasic.All |

### todo

| Command | Description | Arguments | Key options | Notes |
|---|---|---|---|---|
| `todo check-add` | Add a checklist item to a task | `<listId> <taskId>` | `--title <title>` (required) | not idempotent · Tasks.ReadWrite |
| `todo check-delete` | Delete a checklist item | `<listId> <taskId> <itemId>` |  | **destructive** · Tasks.ReadWrite |
| `todo check-done` | Tick a checklist item (--undo to untick) | `<listId> <taskId> <itemId>` | `--undo` | Tasks.ReadWrite |
| `todo checklist` | List checklist items (steps) of a task | `<listId> <taskId>` |  | Tasks.Read |
| `todo create` | Create a task | `<listId>` | `--title <title>` (required) `--body <body>` `--due <date>` `--timezone <tz>` `--importance <level>` | not idempotent · Tasks.ReadWrite |
| `todo create-list` | Create a new task list |  | `--name <name>` (required) | not idempotent · Tasks.ReadWrite |
| `todo delete` | Delete a task | `<listId> <taskId>` |  | **destructive** · Tasks.ReadWrite |
| `todo get` | Get a specific task | `<listId> <taskId>` |  | Tasks.Read |
| `todo lists` | List task lists |  |  | Tasks.Read |
| `todo tasks` | List tasks in a list | `<listId>` | `--top <n>` `--filter <filter>` `--orderby <field>` `--all` | Tasks.Read |
| `todo update` | Update a task | `<listId> <taskId>` | `--title <title>` `--body <body>` `--status <status>` `--importance <level>` `--due <date>` `--timezone <tz>` | Tasks.ReadWrite |
| `todo update-list` | Rename a task list | `<listId>` | `--name <name>` (required) | Tasks.ReadWrite |

### user

| Command | Description | Arguments | Key options | Notes |
|---|---|---|---|---|
| `user clear-presence` | Clear custom presence and revert to automatic |  |  | Presence.ReadWrite |
| `user download-photo` | Download profile photo |  | `--out <path>` `--force` | User.Read |
| `user list` | List users in the organization (requires --org-mode) |  | `--top <n>` `--filter <filter>` `--select <fields>` `--all` | User.Read.All |
| `user mailbox-settings` | Get mailbox settings (timezone, language, working hours) |  |  | MailboxSettings.Read |
| `user me` | Get current user profile |  | `--select <fields>` | User.Read |
| `user ooo` | Get out-of-office / automatic reply settings |  |  | MailboxSettings.Read |
| `user photo` | Get profile photo metadata |  |  | User.Read |
| `user presence` | Get your current presence status (Available, Busy, etc.) |  |  | Presence.Read |
| `user set-language` | Set mailbox language (e.g. "en-US", "fr-FR") | `<locale>` |  | MailboxSettings.ReadWrite |
| `user set-ooo` | Set out-of-office / automatic replies |  | `--status <status>` (required) `--internal <message>` `--external <message>` `--audience <audience>` `--start <datetime>` `--end <datetime>` `--timezone <tz>` | MailboxSettings.ReadWrite |
| `user set-presence` | Set your presence status |  | `--availability <status>` (required) `--activity <activity>` (required) `--duration <minutes>` | Presence.ReadWrite |
| `user set-timezone` | Set mailbox timezone (e.g. "America/New_York", "UTC") | `<timezone>` |  | MailboxSettings.ReadWrite |
| `user upload-photo` | Upload/update profile photo (JPEG, PNG, GIF — max 4MB) | `<filePath>` |  | User.ReadWrite |
| `user working-hours` | Get working hours configuration |  |  | MailboxSettings.Read |

### watch

| Command | Description | Arguments | Key options | Notes |
|---|---|---|---|---|
| `watch` | Print changes since the last run as NDJSON (delta queries; cursor saved per --name) | `<resource>` | `--folder <folder>` `--start <time>` `--end <time>` `--timezone <tz>` `--name <name>` `--reset` `--include-existing` `--strict` `--follow` `--interval <seconds>` `--max-passes <n>` | Mail.Read, Calendars.Read, Files.Read, Contacts.Read |

### workflow

| Command | Description | Arguments | Key options | Notes |
|---|---|---|---|---|
| `workflow digest` | Weekly digest: meetings this week, unread emails, tasks due |  | `--timezone <tz>` | Calendars.Read, Mail.Read, Tasks.Read |
| `workflow email-to-task` | Convert an email into a To Do task | `<messageId>` | `--list <listId>` `--due <date>` `--importance <level>` `--timezone <tz>` | not idempotent · Mail.Read, Tasks.ReadWrite |
| `workflow file-announce` | Get file info and post announcement to a Teams channel (requires --org-mode) | `<driveId> <itemId>` | `--team <teamId>` (required) `--channel <channelId>` (required) `--message <text>` | not idempotent · Files.Read, ChannelMessage.Send |
| `workflow focus-time` | Find gaps in your calendar today for focus time |  | `--work-start <hour>` `--work-end <hour>` `--timezone <tz>` | Calendars.Read |
| `workflow meeting-prep` | Get details for the next upcoming meeting: attendees, agenda, location |  | `--event-id <id>` `--timezone <tz>` | Calendars.Read |
| `workflow standup` | Today's standup: calendar events + open tasks + unread email count |  | `--timezone <tz>` | Calendars.Read, Mail.Read |

<!-- END GENERATED COMMANDS -->

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
- **dotenv** — auto-loads `.env` from working directory.
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
4. Add each leaf command to `src/catalog.ts` (scopes, safety)
5. `npm run docs` to regenerate the command reference, then `npm run build`

---

## License

**MIT** — free to use, modify, and distribute, including commercially, with attribution.

See [LICENSE](./LICENSE) for full text.
