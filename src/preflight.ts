import { CliError, ExitCode } from './errors.js';

/**
 * Pre-flight consent check (spec §6 "consent introspection"): fail fast with a clear
 * missing_scope error instead of a mid-command 403.
 *
 * A required scope is satisfied by itself or by a broader grant of the same resource:
 * X.Read ⊂ X.ReadWrite ⊂ X.ReadWrite.All, X.Read ⊂ X.Read.All, and an exact match with
 * an .All/.Shared suffix (e.g. Mail.Send ⊂ Mail.Send.Shared is NOT assumed — only .All).
 */
export function scopeSatisfied(required: string, granted: Set<string>): boolean {
  if (granted.has(required)) return true;
  const m = required.match(/^(.+?)\.(Read|ReadWrite)(?:\.(All|Shared))?$/);
  if (!m) return granted.has(`${required}.All`);
  const [, resource, level, suffix] = m;
  const candidates = new Set<string>();
  const levels = level === 'Read' ? ['Read', 'ReadWrite'] : ['ReadWrite'];
  for (const l of levels) {
    candidates.add(`${resource}.${l}.All`);
    if (suffix) candidates.add(`${resource}.${l}.${suffix}`);
    else candidates.add(`${resource}.${l}`);
  }
  return [...candidates].some(c => granted.has(c));
}

/**
 * Throw missing_scope (exit 4) when NONE of the command's catalog scopes is satisfied.
 * "Any of" keeps false positives near zero for commands whose catalog lists alternatives.
 */
export function assertConsented(command: string, required: string[], granted: string[]): void {
  if (required.length === 0) return;
  const g = new Set(granted);
  if (required.some(r => scopeSatisfied(r, g))) return;
  throw new CliError(
    `"${command}" needs one of these permissions, which your login hasn't consented to: ${required.join(', ')}.`,
    { code: 'missing_scope', exitCode: ExitCode.PERMISSION, fix: 'ms365 auth login (re-consent; add --org-mode for org scopes)' },
  );
}
