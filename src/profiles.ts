import fs from 'fs';
import path from 'path';
import { configDir, writeFileAtomic } from './auth.js';
import { CliError, ExitCode } from './errors.js';

/** Named sign-in settings (PM-8, D21): `${configDir()}/profiles/<name>.json`, applied with --profile. */
export interface Profile {
  clientId?: string;
  tenantId?: string;
  scopes?: string[];
  orgMode?: boolean;
  account?: string;
}

const PROFILE_NAME = /^[a-z0-9_-]+$/;

const profilesDir = () => path.join(configDir(), 'profiles');

/** Profile names become file names, so only [a-z0-9_-] is allowed (no path traversal). */
export function validateProfileName(name: string): void {
  if (!PROFILE_NAME.test(name)) {
    throw new CliError(`Invalid profile name "${name}". Use lowercase letters, digits, "-" and "_".`, { code: 'usage', exitCode: ExitCode.USAGE });
  }
}

function profilePath(name: string): string {
  validateProfileName(name);
  return path.join(profilesDir(), `${name}.json`);
}

function parseProfile(name: string, data: string): Profile {
  const invalid = () => new CliError(`Profile "${name}" is not valid JSON of the form {clientId?, tenantId?, scopes?, orgMode?, account?}.`,
    { code: 'invalid_profile', fix: `ms365 auth profile set ${name}` });
  let raw: unknown;
  try {
    raw = JSON.parse(data);
  } catch {
    throw invalid();
  }
  if (!raw || typeof raw !== 'object' || Array.isArray(raw)) throw invalid();
  const r = raw as Record<string, unknown>;
  const profile: Profile = {};
  for (const key of ['clientId', 'tenantId', 'account'] as const) {
    if (r[key] === undefined) continue;
    if (typeof r[key] !== 'string') throw invalid();
    profile[key] = r[key];
  }
  if (r.scopes !== undefined) {
    if (!Array.isArray(r.scopes) || !r.scopes.every(s => typeof s === 'string')) throw invalid();
    profile.scopes = r.scopes;
  }
  if (r.orgMode !== undefined) {
    if (typeof r.orgMode !== 'boolean') throw invalid();
    profile.orgMode = r.orgMode;
  }
  return profile;
}

export function readProfile(name: string): Profile {
  let data: string;
  try {
    data = fs.readFileSync(profilePath(name), 'utf8');
  } catch (err) {
    if ((err as NodeJS.ErrnoException).code !== 'ENOENT') throw err;
    throw new CliError(`Profile "${name}" not found.`, { code: 'not_found', exitCode: ExitCode.NOT_FOUND, fix: 'ms365 auth profile list' });
  }
  return parseProfile(name, data);
}

export function profileExists(name: string): boolean {
  return fs.existsSync(profilePath(name));
}

export function listProfiles(): string[] {
  const names = fs.existsSync(profilesDir()) ? fs.readdirSync(profilesDir()) : [];
  return names
    .filter(n => n.endsWith('.json') && PROFILE_NAME.test(n.slice(0, -5)))
    .map(n => n.slice(0, -5))
    .sort();
}

export function saveProfile(name: string, profile: Profile): void {
  const file = profilePath(name);
  fs.mkdirSync(profilesDir(), { recursive: true, mode: 0o700 });
  writeFileAtomic(file, JSON.stringify(profile, null, 2));
}

export function deleteProfile(name: string): boolean {
  try {
    fs.unlinkSync(profilePath(name));
    return true;
  } catch (err) {
    if ((err as NodeJS.ErrnoException).code === 'ENOENT') return false;
    throw err;
  }
}
