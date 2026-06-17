import { Option } from 'commander';
import readline from 'readline';

const OUTPUT_FORMATS = ['json', 'table', 'csv', 'text', 'yaml'] as const;

/** Default output format: table for interactive terminals, json for pipes/scripts. */
export const defaultFormat = process.stdout.isTTY ? 'table' : 'json';

/** Commander Option for -o/--output with validated choices and TTY-aware default. */
export function outputOption(overrideDefault?: string): Option {
  return new Option('-o, --output <format>', 'Output format: json, table, csv, text, yaml')
    .choices([...OUTPUT_FORMATS])
    .default(overrideDefault ?? defaultFormat);
}

/** Global quiet flag — suppressed confirmation messages when true. */
export let quiet = false;
export function setQuiet(val: boolean): void { quiet = val; }

/** Print a non-data message (e.g. "Email sent.") unless --quiet is active. */
export function info(msg: string): void {
  if (!quiet) console.error(msg);
}

/**
 * Prompt user for confirmation on destructive operations.
 * Always prompts — no --yes override — so AI agents can't bypass.
 * Non-TTY stdin auto-rejects to prevent accidental deletions in scripts.
 */
export async function confirm(message: string): Promise<boolean> {
  if (!process.stdin.isTTY) {
    console.error('Destructive operation requires interactive confirmation. Run in a terminal.');
    return false;
  }
  const rl = readline.createInterface({ input: process.stdin, output: process.stderr });
  return new Promise((resolve) => {
    rl.question(`${message} [y/N] `, (answer) => {
      rl.close();
      resolve(answer.trim().toLowerCase() === 'y');
    });
  });
}

/** Validate and return a YYYY-MM-DD date string. */
export function parseDate(input: string): string {
  const d = new Date(input + 'T00:00:00Z');
  if (isNaN(d.getTime())) {
    throw new Error(`Invalid date: "${input}". Use YYYY-MM-DD format.`);
  }
  return input;
}

/** Truncate a string with an ellipsis if it exceeds max length. */
export function truncate(str: string, max: number): string {
  return str.length > max ? str.substring(0, max - 1) + '…' : str;
}

/** Validate a single email address. */
export function validateEmail(email: string): string {
  const trimmed = email.trim();
  if (!/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(trimmed)) {
    throw new Error(`Invalid email address: "${trimmed}"`);
  }
  return trimmed;
}

/** Parse a comma-separated list of emails into Graph API recipient format. */
export function parseRecipients(csv: string): Array<{ emailAddress: { address: string } }> {
  return csv.split(',').map(e => ({
    emailAddress: { address: validateEmail(e) },
  }));
}

/** Format a byte count as a human-readable size. */
export function formatSize(bytes: number): string {
  if (bytes < 1024) return `${bytes} B`;
  if (bytes < 1024 * 1024) return `${(bytes / 1024).toFixed(1)} KB`;
  if (bytes < 1024 * 1024 * 1024) return `${(bytes / (1024 * 1024)).toFixed(1)} MB`;
  return `${(bytes / (1024 * 1024 * 1024)).toFixed(1)} GB`;
}

/** Build OData query params from CLI options. Only includes non-undefined values. */
export function buildODataParams(
  opts: Record<string, string | undefined>,
  mapping: Record<string, string>,
): Record<string, string> {
  const params: Record<string, string> = {};
  for (const [optKey, paramKey] of Object.entries(mapping)) {
    const val = opts[optKey];
    if (val !== undefined) params[paramKey] = val;
  }
  return params;
}
