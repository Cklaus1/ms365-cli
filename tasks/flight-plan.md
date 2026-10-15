# Flight plan: identity profiles

**Mission:** give ms365 two profile types (user; admin borrowed from az, never stored) and a credential-store layer with an opt-in OS keychain, keeping the file store as the default until real-hardware checks pass.

**Decisions:**
- **E1 ⚠ close call:** default stays the 0600 file everywhere; keychain is opt-in (`auth storage use keychain`) until IDT-7 passes — no Mac/Windows available here.
- **E2:** Windows backend = DPAPI blob via PowerShell, no `cmdkey` (argv exposure, size limit).
- **E3:** Windows `auto` → file if PowerShell start > 400 ms.
- **E4:** serial build IDT-4 → 1 → 2 → 3 → 5 → 6 (all touch auth.ts).

**needs-human:** IDT-7 (real macOS + Windows verification; flips E1).

**Step 1 revisions:** 7 revised (incl. macOS `security -i` via stdin instead of `-w <value>`; honest threat model: keychain protects at rest, not against same-user code), 7 added (ID-13..ID-19).

**Critical path:** IDT-2 — multi-account round-trip, 8-process concurrency, corrupt-item handling.

**Shape:** 6 tasks, 1 track, longest chain 6. Test `npm test`+typecheck (6.5 s); smoke `scripts/smoke.sh`; reviewer fable; budget unbounded. Baseline 1535/1535, smoke 12/0/0.

**First:** IDT-4, IDT-1, IDT-2.
