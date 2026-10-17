import { config } from 'dotenv';

/**
 * Load `.env` from the working directory, except the token variables (AH-2(g)): a host supplies
 * MS365_CLI_TOKEN / MS365_CLI_TOKEN_COMMAND through the real process environment only, so a `.env`
 * dropped into a working directory can never switch on external mode. Imported first by src/index.ts.
 */
const TOKEN_KEYS = ['MS365_CLI_TOKEN', 'MS365_CLI_TOKEN_COMMAND'];
const absent = TOKEN_KEYS.filter(k => process.env[k] === undefined);
config({ quiet: true });
for (const k of absent) delete process.env[k];
