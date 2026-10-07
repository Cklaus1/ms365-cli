import { Command } from 'commander';
import fs from 'fs';
import path from 'path';
import { handleErrors } from '../errors.js';
import { defaultFormat, info } from '../utils.js';
import { buildCommandTree, type CommandNode, type CommandTree, type GroupNode, type OptionNode } from '../command-tree.js';
import { endpointKeys } from './schema.js';
import { introspectionProgram } from './registry.js';
import { VERSION } from '../version.js';

export interface SkillFile {
  /** Path relative to the output directory. */
  path: string;
  content: string;
}

const isOrgMode = (description: string) => /requires --org-mode/i.test(description);

/** Escape `|` so a value stays inside its Markdown table cell. */
const escapeCell = (s: string) => s.replace(/\|/g, '\\|');

function usage(cmd: CommandNode): string {
  const args = cmd.arguments.map(a => {
    const name = `${a.name}${a.variadic ? '...' : ''}`;
    return a.required ? `<${name}>` : `[${name}]`;
  });
  const required = cmd.options.filter(o => o.required).map(o => o.flags.replace(/^-\w, /, ''));
  return ['ms365', ...cmd.names, ...args, ...required].join(' ');
}

function optionLine(opt: OptionNode): string {
  const notes: string[] = [];
  if (opt.required) notes.push('required');
  if (opt.choices) notes.push(`one of: ${opt.choices.join(', ')}`);
  if (opt.listChoices) notes.push(`comma-separated list of: ${opt.listChoices.join(', ')}`);
  if (opt.long === '--output' && opt.default === defaultFormat) notes.push('default: table on a TTY, json when piped');
  else if (opt.default !== undefined && !opt.negate) notes.push(`default: ${JSON.stringify(opt.default)}`);
  const suffix = notes.length > 0 ? ` (${notes.join('; ')})` : '';
  return `- \`${opt.flags}\` — ${opt.description}${suffix}`;
}

/** Table of commands plus a per-command section listing arguments and options. */
function renderCommands(leaves: CommandNode[]): string {
  const table = [
    '| Command | Description | Usage |',
    '|---------|-------------|-------|',
    ...leaves.map(c => `| \`${c.names.join(' ')}\` | ${escapeCell(c.description)} | \`${escapeCell(usage(c))}\` |`),
  ];
  const details = leaves.map(c => {
    const lines = [`### ${c.names.join(' ')}`, '', c.description, '', '```bash', usage(c), '```'];
    if (c.aliases.length > 0) lines.push('', `Aliases: ${c.aliases.map(a => `\`${a}\``).join(', ')}`);
    if (c.arguments.length > 0) {
      lines.push('', 'Arguments:', ...c.arguments.map(a =>
        `- \`${a.name}\`${a.required ? '' : ' (optional)'}${a.variadic ? ' (repeatable)' : ''}${a.description ? ` — ${a.description}` : ''}`));
    }
    if (c.options.length > 0) lines.push('', 'Options:', ...c.options.map(optionLine));
    return lines.join('\n');
  });
  return `${table.join('\n')}\n\n${details.join('\n\n')}\n`;
}

function generateGroupSkill(group: GroupNode, leaves: CommandNode[], schemaKeys: string[]): string {
  const name = group.names[0];
  const orgNote = isOrgMode(group.description) ? '\n> Requires the global `--org-mode` flag.\n' : '';
  const aliasNote = group.aliases.length > 0
    ? `\nAlias: ${group.aliases.map(a => `\`ms365 ${a}\``).join(', ')}\n`
    : '';
  const schemaHints = schemaKeys.map(k => `ms365 schema ${k}`.padEnd(40) + `# Graph parameters for ${k}`);

  return `---
name: ms365-${name}
version: ${VERSION}
description: ${JSON.stringify(group.description)}
metadata:
  category: "productivity"
  requires:
    bins: ["ms365"]
  cliHelp: "ms365 ${name} --help"
---

# ms365 ${name}

${group.description}
${orgNote}${aliasNote}
> **PREREQUISITE:** Read [\`../ms365-shared/SKILL.md\`](../ms365-shared/SKILL.md) for auth, global flags, and output formats.

\`\`\`bash
ms365 ${name} <command> [options]
\`\`\`

## Commands

${renderCommands(leaves)}
## Discovering Parameters

\`\`\`bash
ms365 ${name} <command> --help
${[...schemaHints, `ms365 schema --list`.padEnd(40) + '# All available endpoint schemas'].join('\n')}
\`\`\`
`;
}

function generateSharedSkill(globals: OptionNode[], topLevelLeaves: CommandNode[]): string {
  const flagRows = globals.map(o => `| \`${o.flags}\` | ${escapeCell(o.description)} |`).join('\n');
  return `---
name: ms365-shared
version: ${VERSION}
description: "Shared authentication, global flags, output formats, and security rules for ms365-cli."
metadata:
  category: "shared"
  requires:
    bins: ["ms365"]
---

# ms365-cli — Shared Reference

## Authentication

See [\`../ms365-auth/SKILL.md\`](../ms365-auth/SKILL.md). Start with \`ms365 auth login\` (device code flow) and check with \`ms365 auth status\`.

## Global Flags

Global flags go **before** the command: \`ms365 --dry-run mail send …\`.

| Flag | Description |
|------|-------------|
${flagRows}

## Output Formats

Commands that return data accept \`-o, --output <format>\`:

- **json** — Pretty-printed JSON (default when stdout is not a TTY)
- **table** — ASCII table with nested object flattening (default on a TTY)
- **csv** — Comma-separated with headers
- **yaml** — YAML format
- **text** — Key-value pairs

Use \`-o json\` in scripts and agents: an empty result is still valid JSON.

## Pagination

- \`--top <n>\` — Number of results (page size with \`--all\`)
- \`--all\` — Fetch all pages automatically
- \`--page-limit <n>\` — Max pages with \`--all\` (default: 100)
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

## Other Commands

${renderCommands(topLevelLeaves)}
## Environment Variables

| Variable | Description |
|----------|-------------|
| \`MS365_CLI_TOKEN\` | Pre-obtained OAuth2 token |
| \`MS365_CLI_CLIENT_ID\` | Custom Azure AD app client ID |
| \`MS365_CLI_TENANT_ID\` | Azure AD tenant ID |
| \`MS365_CLI_ORG_MODE\` | Enable org mode |
| \`MS365_CLI_READ_ONLY\` | Restrict to read-only operations |
| \`MS365_CLI_API_VERSION\` | API version |
| \`MS365_CLI_LOG_LEVEL\` | Log level: debug, info, warn, error |
| \`MS365_CLI_LOG_FILE\` | JSON log file path |

## Security Rules

- Never expose tokens in output or logs
- Use \`--dry-run\` to verify requests before executing
- Use \`--read-only\` when exploring to prevent accidental writes
- Destructive commands ask for confirmation and refuse without a TTY
- Planner update/delete operations require ETag for concurrency control
`;
}

function generateSkillIndex(groups: GroupNode[]): string {
  const rows = [
    '| Skill | Description | Category |',
    '|-------|-------------|----------|',
    '| [ms365-shared](ms365-shared/SKILL.md) | Auth, global flags, output formats | shared |',
    ...groups.map(g => {
      const name = g.names[0];
      const category = isOrgMode(g.description) ? 'organization' : 'personal';
      return `| [ms365-${name}](ms365-${name}/SKILL.md) | ${escapeCell(g.description)} | ${category} |`;
    }),
  ];

  return `# ms365-cli Skills Index

Auto-generated skill definitions for AI agent integration.

${rows.join('\n')}

## Usage with AI Agents

These SKILL.md files describe ms365-cli capabilities in a format that AI agents (Claude Code, etc.) can consume to discover and use Microsoft 365 operations.

Each skill file contains:
- Service description and scope
- Every command with its arguments and options
- Pointers to the endpoint schemas (\`ms365 schema\`)

Generated by: \`ms365 generate-skills\`
`;
}

/**
 * Render the skill catalog from a command tree: one skill per top-level group, a shared skill
 * (global flags + top-level leaf commands) and an index. `filter` selects groups by name.
 */
export function generateSkillFiles(tree: CommandTree, globals: OptionNode[], filter?: RegExp): SkillFile[] {
  const schemaKeys = endpointKeys();
  const groups = tree.groups.filter(g => g.names.length === 1 && (!filter || filter.test(g.path)));
  const files: SkillFile[] = [
    { path: 'ms365-shared/SKILL.md', content: generateSharedSkill(globals, tree.leaves.filter(l => l.groups.length === 0)) },
  ];
  for (const group of groups) {
    const leaves = tree.leaves.filter(l => l.groups[0] === group.path);
    const keys = schemaKeys.filter(k => k.startsWith(`${group.path}.`));
    files.push({ path: `ms365-${group.path}/SKILL.md`, content: generateGroupSkill(group, leaves, keys) });
  }
  files.push({ path: 'INDEX.md', content: generateSkillIndex(groups) });
  return files;
}

export function registerGenerateSkillsCommand(program: Command): void {
  program
    .command('generate-skills')
    .description('Generate SKILL.md files for AI agent integration')
    .option('--output-dir <dir>', 'Output directory', './skills')
    .option('--filter <pattern>', 'Only generate skills whose group name matches this regex')
    .action(handleErrors(async (opts: { outputDir: string; filter?: string }, cmd: Command) => {
      const outDir = path.resolve(opts.outputDir);
      const tree = buildCommandTree(introspectionProgram());
      // Global flags are declared on the running program (index.ts), not by the command modules.
      const globals = cmd.parent ? buildCommandTree(cmd.parent).options : [];
      const files = generateSkillFiles(tree, globals, opts.filter ? new RegExp(opts.filter, 'i') : undefined);

      for (const file of files) {
        const target = path.join(outDir, file.path);
        fs.mkdirSync(path.dirname(target), { recursive: true });
        fs.writeFileSync(target, file.content);
        info(`  Generated: ${file.path}`);
      }
      info(`\n${files.length} skill files written to ${outDir}/`);
    }));
}
