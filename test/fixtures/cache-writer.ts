// Child process for the concurrency gate: hammer writeFileAtomic with distinct payloads.
import { writeFileAtomic } from '../../src/auth.js';
const [target, id, rounds] = process.argv.slice(2);
const payload = JSON.stringify({ writer: id, filler: 'x'.repeat(200_000 + Number(id) * 1000) });
for (let i = 0; i < Number(rounds); i++) writeFileAtomic(target, payload);
