/**
 * The environment for a child process that is not ms365 (az, keychain tools, the browser opener): the parent's,
 * minus every `MS365_CLI_TOKEN*` variable, so a host-supplied token never reaches another program (AH-2).
 * The MCP runner spawns ms365 itself and keeps the token; every other spawn site takes its env from here
 * (guarded by test/child-env.test.ts).
 */
export function childEnv(base: NodeJS.ProcessEnv = process.env): NodeJS.ProcessEnv {
  const env: NodeJS.ProcessEnv = {};
  for (const [k, v] of Object.entries(base)) if (!k.startsWith('MS365_CLI_TOKEN')) env[k] = v;
  return env;
}
