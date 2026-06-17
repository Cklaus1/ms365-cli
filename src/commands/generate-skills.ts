import { Command } from 'commander';
import { handleErrors } from '../errors.js';
import fs from 'fs';
import path from 'path';

interface ServiceDef {
  name: string;
  description: string;
  commands: Array<{ name: string; description: string; args?: string }>;
  orgMode?: boolean;
}

const SERVICES: ServiceDef[] = [
  {
    name: 'mail',
    description: 'Outlook mail: read, send, reply, search, and manage email messages and folders.',
    commands: [
      { name: 'list', description: 'List inbox messages with OData filtering and search' },
      { name: 'get', description: 'Get a specific message by ID (raw API response)' },
      { name: 'read', description: 'Read a message with full body, attachments, categories, and flags' },
      { name: 'triage', description: 'Show inbox summary with sender, subject, date, and flag indicators' },
      { name: 'send', description: 'Send an email with to/cc/bcc, subject, body (text or HTML)' },
      { name: 'reply', description: 'Reply or reply-all to a message' },
      { name: 'draft', description: 'Create a draft email' },
      { name: 'folders', description: 'List mail folders' },
      { name: 'create-folder', description: 'Create a mail folder' },
      { name: 'move', description: 'Move a message to a folder', args: '<messageId>' },
      { name: 'delete', description: 'Delete a message', args: '<messageId>' },
      { name: 'attachments', description: 'List attachments for a message', args: '<messageId>' },
      { name: 'shared-list', description: 'List shared mailbox messages (org mode)', args: '<userId>' },
      { name: 'shared-send', description: 'Send from shared mailbox (org mode)', args: '<userId>' },
    ],
  },
  {
    name: 'calendar',
    description: 'Calendar: manage events, check availability, and view schedules.',
    commands: [
      { name: 'list', description: 'List calendars' },
      { name: 'events', description: 'List calendar events with filtering and sorting' },
      { name: 'get', description: 'Get a specific event by ID', args: '<eventId>' },
      { name: 'create', description: 'Create a calendar event with attendees and location' },
      { name: 'update', description: 'Update an existing event', args: '<eventId>' },
      { name: 'delete', description: 'Delete an event', args: '<eventId>' },
      { name: 'view', description: 'Get calendar view for a date range (expands recurring)' },
    ],
  },
  {
    name: 'drive',
    description: 'OneDrive: upload, download, list, and manage files and folders.',
    commands: [
      { name: 'list', description: 'List available drives' },
      { name: 'root', description: 'Get root folder of a drive', args: '<driveId>' },
      { name: 'files', description: 'List files in a folder', args: '<driveId> <folderId>' },
      { name: 'upload', description: 'Upload a file (auto-detects MIME, streams large files)', args: '<driveId> <parentId> <filePath>' },
      { name: 'download', description: 'Download a file', args: '<driveId> <itemId> <childId>' },
      { name: 'download-shared', description: 'Download via sharing link', args: '<shareId>' },
      { name: 'delete', description: 'Delete a drive item', args: '<driveId> <itemId>' },
    ],
  },
  {
    name: 'todo',
    description: 'Microsoft To Do: manage task lists and tasks.',
    commands: [
      { name: 'lists', description: 'List task lists' },
      { name: 'tasks', description: 'List tasks in a list', args: '<listId>' },
      { name: 'get', description: 'Get a specific task', args: '<listId> <taskId>' },
      { name: 'create', description: 'Create a task with title, body, due date, importance', args: '<listId>' },
      { name: 'update', description: 'Update a task', args: '<listId> <taskId>' },
      { name: 'delete', description: 'Delete a task', args: '<listId> <taskId>' },
    ],
  },
  {
    name: 'contacts',
    description: 'Outlook contacts: create, read, update, and delete contacts.',
    commands: [
      { name: 'list', description: 'List contacts with search and filtering' },
      { name: 'get', description: 'Get a specific contact', args: '<contactId>' },
      { name: 'create', description: 'Create a contact' },
      { name: 'update', description: 'Update a contact', args: '<contactId>' },
      { name: 'delete', description: 'Delete a contact', args: '<contactId>' },
    ],
  },
  {
    name: 'teams',
    description: 'Microsoft Teams: list teams, channels, members, and send channel messages.',
    orgMode: true,
    commands: [
      { name: 'list', description: 'List joined teams' },
      { name: 'get', description: 'Get team details', args: '<teamId>' },
      { name: 'members', description: 'List team members', args: '<teamId>' },
      { name: 'channels', description: 'List team channels', args: '<teamId>' },
      { name: 'channel-messages', description: 'List channel messages', args: '<teamId> <channelId>' },
      { name: 'send-channel-message', description: 'Send message to channel', args: '<teamId> <channelId>' },
    ],
  },
  {
    name: 'chat',
    description: 'Teams chat: list chats, read messages, send messages, and reply.',
    orgMode: true,
    commands: [
      { name: 'list', description: 'List chats' },
      { name: 'messages', description: 'List chat messages', args: '<chatId>' },
      { name: 'send', description: 'Send a chat message', args: '<chatId>' },
      { name: 'reply', description: 'Reply to a message', args: '<chatId> <messageId>' },
    ],
  },
  {
    name: 'sharepoint',
    description: 'SharePoint: search sites, browse lists, and access document libraries.',
    orgMode: true,
    commands: [
      { name: 'search', description: 'Search SharePoint sites', args: '<query>' },
      { name: 'site', description: 'Get site details', args: '<siteId>' },
      { name: 'lists', description: 'List site lists', args: '<siteId>' },
      { name: 'list-items', description: 'List items in a list', args: '<siteId> <listId>' },
      { name: 'drives', description: 'List site document libraries', args: '<siteId>' },
      { name: 'delta', description: 'Track site changes (delta sync)' },
    ],
  },
  {
    name: 'planner',
    description: 'Microsoft Planner: manage plans, tasks, and buckets.',
    commands: [
      { name: 'plans', description: 'List my plans' },
      { name: 'tasks', description: 'List tasks in a plan', args: '<planId>' },
      { name: 'my-tasks', description: 'List my tasks across all plans' },
      { name: 'create-task', description: 'Create a task in a plan' },
      { name: 'update-task', description: 'Update a task (requires ETag)', args: '<taskId>' },
      { name: 'buckets', description: 'List plan buckets', args: '<planId>' },
    ],
  },
  {
    name: 'onenote',
    description: 'OneNote: list notebooks, sections, pages, and create pages.',
    commands: [
      { name: 'notebooks', description: 'List notebooks' },
      { name: 'sections', description: 'List sections', args: '<notebookId>' },
      { name: 'pages', description: 'List pages', args: '<sectionId>' },
      { name: 'page', description: 'Get page content', args: '<pageId>' },
      { name: 'create-page', description: 'Create a page', args: '<sectionId>' },
    ],
  },
  {
    name: 'excel',
    description: 'Excel: read ranges, create charts, format and sort data in workbooks.',
    commands: [
      { name: 'worksheets', description: 'List worksheets', args: '<driveId> <itemId>' },
      { name: 'range', description: 'Get cell range', args: '<driveId> <itemId> <worksheetId> <address>' },
      { name: 'create-chart', description: 'Create a chart', args: '<driveId> <itemId> <worksheetId>' },
      { name: 'format-range', description: 'Format cells (bold, color)', args: '<driveId> <itemId> <worksheetId>' },
      { name: 'sort-range', description: 'Sort a range', args: '<driveId> <itemId> <worksheetId>' },
    ],
  },
  {
    name: 'search',
    description: 'Microsoft 365 search: search across mail, files, sites, events, and people.',
    commands: [
      { name: 'query', description: 'Search across M365 (driveItem, message, site, event, chatMessage)', args: '<query>' },
      { name: 'people', description: 'Search people', args: '[query]' },
    ],
  },
  {
    name: 'workflow',
    description: 'Productivity workflows that combine multiple M365 services.',
    commands: [
      { name: 'standup', description: "Today's standup: calendar events + unread email count" },
      { name: 'meeting-prep', description: 'Next meeting details: attendees, agenda, join link' },
      { name: 'email-to-task', description: 'Convert an email into a To Do task', args: '<messageId>' },
      { name: 'digest', description: 'Weekly digest: meetings, unread emails, tasks due' },
      { name: 'focus-time', description: 'Find gaps in your calendar for focus work' },
      { name: 'file-announce', description: 'Announce a file in a Teams channel', args: '<driveId> <itemId>' },
    ],
  },
];

function generateServiceSkill(service: ServiceDef): string {
  const orgNote = service.orgMode ? '\n> Requires `--org-mode` flag.\n' : '';

  const commandTable = service.commands.map(cmd => {
    const usage = cmd.args ? `ms365 ${service.name} ${cmd.name} ${cmd.args}` : `ms365 ${service.name} ${cmd.name}`;
    return `| \`${cmd.name}\` | ${cmd.description} | \`${usage}\` |`;
  }).join('\n');

  return `---
name: ms365-${service.name}
version: 0.1.0
description: "${service.description}"
metadata:
  category: "productivity"
  requires:
    bins: ["ms365"]
  cliHelp: "ms365 ${service.name} --help"
---

# ms365 ${service.name}

${service.description}
${orgNote}
> **PREREQUISITE:** Read [\`../ms365-shared/SKILL.md\`](../ms365-shared/SKILL.md) for auth, global flags, and output formats.

\`\`\`bash
ms365 ${service.name} <command> [options]
\`\`\`

## Commands

| Command | Description | Usage |
|---------|-------------|-------|
${commandTable}

## Common Options

All list commands support:
- \`--top <n>\` — Limit results
- \`--filter <expr>\` — OData filter
- \`--select <fields>\` — Choose fields
- \`--all\` — Auto-paginate
- \`-o json|table|csv|yaml|text\` — Output format

## Discovering Parameters

\`\`\`bash
ms365 schema ${service.name}.list    # See parameters for list endpoint
ms365 schema --list                  # See all available schemas
\`\`\`
`;
}

function generateSharedSkill(): string {
  return `---
name: ms365-shared
version: 0.1.0
description: "Shared authentication, global flags, output formats, and security rules for ms365-cli."
metadata:
  category: "shared"
  requires:
    bins: ["ms365"]
---

# ms365-cli — Shared Reference

## Authentication

\`\`\`bash
ms365 auth login     # Device code flow (interactive)
ms365 auth status    # Check current session
ms365 auth logout    # Clear credentials
ms365 auth accounts  # List cached accounts
\`\`\`

## Global Flags

| Flag | Description |
|------|-------------|
| \`--org-mode\` | Enable Teams, SharePoint, shared mailboxes |
| \`--dry-run\` | Preview API request without executing |
| \`--api-version <ver>\` | Graph API version: \`v1.0\` (default) or \`beta\` |
| \`--read-only\` | Restrict to GET operations |
| \`-o <format>\` | Output: \`json\`, \`table\`, \`csv\`, \`yaml\`, \`text\` |

## Output Formats

- **json** — Pretty-printed JSON (default)
- **table** — ASCII table with nested object flattening (dot notation)
- **csv** — Comma-separated with headers
- **yaml** — YAML format
- **text** — Key-value pairs

## Pagination

- \`--all\` — Fetch all pages automatically
- \`--page-limit <n>\` — Max pages (default: 100)
- \`--page-delay <ms>\` — Delay between pages

## OData Filters

\`\`\`bash
--filter "isRead eq false"
--filter "importance eq 'high'"
--filter "start/dateTime ge '2026-01-01'"
--select "id,subject,from"
--orderby "receivedDateTime desc"
--search "quarterly report"
\`\`\`

## Environment Variables

| Variable | Description |
|----------|-------------|
| \`MS365_CLI_TOKEN\` | Pre-obtained OAuth2 token |
| \`MS365_CLI_CLIENT_ID\` | Custom Azure AD app client ID |
| \`MS365_CLI_TENANT_ID\` | Azure AD tenant ID |
| \`MS365_CLI_ORG_MODE\` | Enable org mode |
| \`MS365_CLI_API_VERSION\` | API version |
| \`MS365_CLI_LOG_LEVEL\` | Log level: debug, info, warn, error |
| \`MS365_CLI_LOG_FILE\` | JSON log file path |

## Security Rules

- Never expose tokens in output or logs
- Use \`--dry-run\` to verify requests before executing
- Use \`--read-only\` when exploring to prevent accidental writes
- Planner update/delete operations require ETag for concurrency control
`;
}

function generateSkillIndex(services: ServiceDef[]): string {
  const rows = [
    '| Skill | Description | Category |',
    '|-------|-------------|----------|',
    '| [ms365-shared](ms365-shared/SKILL.md) | Auth, global flags, output formats | shared |',
  ];

  for (const service of services) {
    const cat = service.orgMode ? 'organization' : 'personal';
    rows.push(`| [ms365-${service.name}](ms365-${service.name}/SKILL.md) | ${service.description.substring(0, 70)} | ${cat} |`);
  }

  return `# ms365-cli Skills Index

Auto-generated skill definitions for AI agent integration.

${rows.join('\n')}

## Usage with AI Agents

These SKILL.md files describe ms365-cli capabilities in a format that AI agents (Claude Code, etc.) can consume to discover and use Microsoft 365 operations.

Each skill file contains:
- Service description and scope
- Available commands with usage examples
- Common options and parameters
- Links to related skills

Generated by: \`ms365 generate-skills\`
`;
}

export function registerGenerateSkillsCommand(program: Command): void {
  program
    .command('generate-skills')
    .description('Generate SKILL.md files for AI agent integration')
    .option('--output-dir <dir>', 'Output directory', './skills')
    .option('--filter <pattern>', 'Only generate skills matching pattern')
    .action(handleErrors(async (opts) => {
      const outDir = path.resolve(opts.output_dir || opts.outputDir || './skills');

      if (!fs.existsSync(outDir)) {
        fs.mkdirSync(outDir, { recursive: true });
      }

      let services = SERVICES;
      if (opts.filter) {
        const re = new RegExp(opts.filter, 'i');
        services = services.filter(s => re.test(s.name));
      }

      // Generate shared skill
      const sharedDir = path.join(outDir, 'ms365-shared');
      if (!fs.existsSync(sharedDir)) fs.mkdirSync(sharedDir, { recursive: true });
      fs.writeFileSync(path.join(sharedDir, 'SKILL.md'), generateSharedSkill());
      console.log(`  Generated: ms365-shared/SKILL.md`);

      // Generate per-service skills
      for (const service of services) {
        const serviceDir = path.join(outDir, `ms365-${service.name}`);
        if (!fs.existsSync(serviceDir)) fs.mkdirSync(serviceDir, { recursive: true });
        fs.writeFileSync(path.join(serviceDir, 'SKILL.md'), generateServiceSkill(service));
        console.log(`  Generated: ms365-${service.name}/SKILL.md`);
      }

      // Generate index
      fs.writeFileSync(path.join(outDir, 'INDEX.md'), generateSkillIndex(services));
      console.log(`  Generated: INDEX.md`);

      console.log(`\n${services.length + 1} skill files written to ${outDir}/`);
    }));
}
