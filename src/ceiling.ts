/**
 * Profile scopes as a fail-closed ceiling (AH-1). For `user` and `host` profiles that list `scopes`, every scope a
 * command run may call (`usedScopes`, catalog) must be satisfied by those scopes. Checked in preAction before the
 * offline and dry-run returns, whatever MS365_CLI_PREFLIGHT says and whatever the token holds; `mcp serve` and
 * `describe` show only what passes. Admin profiles (ID-8) and runs without a profile have no ceiling (OQ-3).
 */
import { commandMeta, usedScopes, type ScopeClause } from './catalog.js';
import { CliError, ExitCode } from './errors.js';
import { scopeSatisfied } from './preflight.js';
import { profileType, readProfile, type Profile } from './profiles.js';

/** Sign-in scopes that grant no Graph access, so they never widen or narrow a ceiling. */
const OIDC_SCOPES: Record<string, true> = { offline_access: true, openid: true, profile: true, email: true };
const GRAPH_RESOURCE = 'https://graph.microsoft.com/';

/** The profile's ceiling (its `scopes`, Graph resource prefix and OIDC scopes dropped), or null when none applies. */
export function profileCeiling(profile: Profile): Set<string> | null {
  if (profileType(profile) === 'admin' || !profile.scopes) return null;
  return new Set(profile.scopes
    .map(s => (s.startsWith(GRAPH_RESOURCE) ? s.slice(GRAPH_RESOURCE.length) : s))
    .filter(s => OIDC_SCOPES[s] !== true));
}

/** The ceiling of the profile chosen by `--profile` (else MS365_CLI_PROFILE), with its name; null when none applies. */
export function activeCeiling(profileFlag: string | undefined): { profile: string; ceiling: Set<string> } | null {
  const name = profileFlag || process.env.MS365_CLI_PROFILE;
  if (!name) return null;
  const ceiling = profileCeiling(readProfile(name));
  return ceiling ? { profile: name, ceiling } : null;
}

/** The first clause none of whose scopes the ceiling satisfies, or undefined when the run fits. */
export function unmetClause(clauses: ScopeClause[], ceiling: Set<string>): ScopeClause | undefined {
  return clauses.find(c => c.scopes.length > 0 && !c.scopes.some(s => scopeSatisfied(s, ceiling)));
}

/** Throw `scope_ceiling` (exit 4) when the run needs a scope outside the profile's ceiling. */
export function assertWithinCeiling(command: string, clauses: ScopeClause[], ceiling: Set<string>, profileName: string): void {
  const unmet = unmetClause(clauses, ceiling);
  if (!unmet) return;
  const needs = unmet.scopes.join(' or ');
  throw new CliError(
    `"${command}"${unmet.for ? ` ${unmet.for}` : ''} needs ${needs}, which is outside the scope ceiling of profile "${profileName}" (${[...ceiling].join(', ') || 'no scopes'}).`,
    { code: 'scope_ceiling', exitCode: ExitCode.PERMISSION, fix: `use a profile whose scopes include ${needs}, or widen profile "${profileName}": ms365 auth profile set ${profileName} --scopes <list>` },
  );
}

/**
 * What a command may do under the ceiling, without parsed arguments (MCP tool list, `describe`): `allowed` when some
 * run fits; for a selector command `values` are the selector values that fit (the MCP enum is narrowed to them).
 */
export function ceilingSurface(path: string, ceiling: Set<string>): { allowed: boolean; values?: string[] } {
  const sel = commandMeta(path)?.scopesBy;
  // A delegate read (`base`) fits when the caller's own read does; `--mailbox` is then checked per run.
  if (!sel || 'base' in sel) return { allowed: !unmetClause(sel ? [{ scopes: sel.base }] : usedScopes(path), ceiling) };
  // A boolean-flag selector fits when either mode does; the run's own mode is checked per run.
  if ('flag' in sel) return { allowed: !unmetClause([{ scopes: sel.with }], ceiling) || !unmetClause([{ scopes: sel.without }], ceiling) };
  const values = Object.keys(sel.map).filter(v => !unmetClause([{ scopes: sel.map[v] }], ceiling));
  return { allowed: values.length > 0, values };
}
