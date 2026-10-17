import type { ScopeClause } from './catalog.js';
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
 * Throw missing_scope (exit 4) when any clause has none of its scopes satisfied (CNF, F14).
 * "Any of" inside a clause keeps false positives near zero for clauses that list alternatives.
 */
export function assertConsented(command: string, clauses: ScopeClause[], granted: string[], adminProfile?: string): void {
  const g = new Set(granted);
  const unmet = clauses.find(c => c.scopes.length > 0 && !c.scopes.some(r => scopeSatisfied(r, g)));
  if (!unmet) return;
  if (adminProfile !== undefined) {
    // ID-8: the scopes are whatever the Azure CLI app has in the tenant, so re-consenting cannot help; a user profile can.
    const fix = 'ms365 --org-mode auth login';
    throw new CliError(
      `"${command}"${unmet.for ? ` ${unmet.for}` : ''} needs ${unmet.scopes.join(' or ')}; the Azure CLI token used by admin profile "${adminProfile}" does not include it — use a user profile (${fix}) for this command.`,
      { code: 'missing_scope', exitCode: ExitCode.PERMISSION, fix: `use a user profile: ${fix}` },
    );
  }
  const what = unmet.for
    ? `"${command}" ${unmet.for} needs ${unmet.scopes.join(' or ')}, which your login hasn't consented to.`
    : `"${command}" needs one of these permissions, which your login hasn't consented to: ${unmet.scopes.join(', ')}.`;
  throw new CliError(what, { code: 'missing_scope', exitCode: ExitCode.PERMISSION, fix: 'ms365 auth login (re-consent; add --org-mode for org scopes)' });
}
