// Runs before every test file. Tests must never touch the developer's real config:
// point HOME (which also covers code that ignores MS365_CLI_CONFIG_DIR, e.g. older
// revisions during fail-first runs) and the config dir at a throwaway directory.
import fs from 'fs';
import os from 'os';
import path from 'path';

const sandbox = fs.mkdtempSync(path.join(os.tmpdir(), 'ms365-test-home-'));
process.env.HOME = sandbox;
process.env.USERPROFILE = sandbox;
process.env.MS365_CLI_CONFIG_DIR = path.join(sandbox, '.config', 'ms365-cli');
process.env.MS365_CLI_AUDIT = '0';
