#!/usr/bin/env bash
# Offline smoke suite for ms365-cli (build-loop smoke/integration gate).
# Each check has a concrete signal (exit code + output match). Checks whose
# feature doesn't exist yet report SKIP; the final stop condition requires 0 SKIP.
# Usage: scripts/smoke.sh            (builds first)
#        SMOKE_NO_BUILD=1 scripts/smoke.sh
set -u
cd "$(dirname "$0")/.."
[ -z "${SMOKE_NO_BUILD:-}" ] && npm run -s build >/dev/null 2>&1
CLI="node dist/index.js"
TMP="$(mktemp -d)"; trap 'rm -rf "$TMP"' EXIT
pass=0; fail=0; skip=0

check() { # name expected_exit cmd... ; output in $TMP/out, $TMP/err
  local name="$1" want="$2"; shift 2
  timeout 30 "$@" >"$TMP/out" 2>"$TMP/err"; local got=$?
  if [ "$got" = 124 ]; then echo "FAIL  $name (hang/timeout)"; fail=$((fail+1)); return 1; fi
  if [ "$got" != "$want" ]; then echo "FAIL  $name (exit $got, want $want): $(head -c 200 "$TMP/err" "$TMP/out" | tr '\n' ' ')"; fail=$((fail+1)); return 1; fi
  return 0
}
ok()   { echo "PASS  $1"; pass=$((pass+1)); }
bad()  { echo "FAIL  $1"; fail=$((fail+1)); }
skipc(){ echo "SKIP  $1"; skip=$((skip+1)); }
has_cmd() { $CLI --help 2>/dev/null | grep -qE "^  $1( |\b)"; }

# 1. version
check "version" 0 $CLI --version && { grep -qx "0.1.0" "$TMP/out" && ok "version" || bad "version output"; }

# 2. schema --list works offline (no auth)
check "schema --list" 0 env MS365_CLI_CLIENT_ID=  $CLI schema --list && ok "schema --list"

# 3. dry-run send exits 0 and prints the requests (AH-6: always the draft path, create then /send)
check "dry-run mail send" 0 $CLI --dry-run mail send --to a@b.com --subject s --body b && {
  # F1: the dry-run trace goes to stderr; stdout carries only the command's output
  if grep -q "\[dry-run\]" "$TMP/out"; then bad "dry-run trace leaked to stdout";
  elif grep -q "\[dry-run\] POST https://graph.microsoft.com/v1.0/me/messages$" "$TMP/err" \
    && grep -q "\[dry-run\] POST https://graph.microsoft.com/v1.0/me/messages/DRAFT_ID/send" "$TMP/err"; then ok "dry-run mail send";
  else bad "dry-run mail send trace"; fi; }

# 4. dry-run large upload targets the correct upload-session URL
head -c 12582912 /dev/zero > "$TMP/big.bin"
check "dry-run drive cp 12MB" 0 $CLI --dry-run drive cp "$TMP/big.bin" /smoke/big.bin && {
  if grep -q "content/createUploadSession" "$TMP/err"; then bad "upload session URL (has /content/)";
  elif grep -q ":/createUploadSession" "$TMP/err"; then ok "dry-run drive cp 12MB";
  else bad "upload session URL not shown"; fi; }

# 5. auth failure maps to exit code 3 with an actionable hint (no valid login available offline)
HOME_SAVE="$HOME"; export HOME="$TMP/home"; mkdir -p "$HOME"
check "auth failure exit 3" 3 env MS365_CLI_TOKEN=  MS365_CLI_CLIENT_ID=00000000-0000-0000-0000-000000000000 $CLI --read-only user me && {
  grep -qi "auth login" "$TMP/err" && ok "auth failure exit 3" || bad "auth failure hint"; }

# 6. JSON error format
check "json error" 3 env MS365_CLI_TOKEN= MS365_CLI_CLIENT_ID=00000000-0000-0000-0000-000000000000 $CLI --error-format json --read-only user me && {
  node -e 'const e=JSON.parse(require("fs").readFileSync(process.argv[1],"utf8").trim().split("\n").pop()); if(e.error.exitCode!==3) process.exit(1)' "$TMP/err" && ok "json error" || bad "json error shape"; }
export HOME="$HOME_SAVE"

# 7. describe --json catalog
if has_cmd describe; then
  check "describe --json" 0 $CLI describe --json && {
    node -e 'const c=JSON.parse(require("fs").readFileSync(process.argv[1],"utf8")); if(!Array.isArray(c.commands)||c.commands.length<100) process.exit(1)' "$TMP/out" && ok "describe --json" || bad "describe --json shape"; }
else skipc "describe --json"; fi

# 8. MCP server answers tools/list
if has_cmd mcp; then
  printf '%s\n' '{"jsonrpc":"2.0","id":1,"method":"initialize","params":{"protocolVersion":"2025-06-18","capabilities":{},"clientInfo":{"name":"smoke","version":"0"}}}' '{"jsonrpc":"2.0","id":2,"method":"tools/list"}' > "$TMP/mcp.in"
  timeout 20 $CLI mcp serve < "$TMP/mcp.in" > "$TMP/out" 2>"$TMP/err"
  node -e 'const l=require("fs").readFileSync(process.argv[1],"utf8").trim().split("\n").map(JSON.parse); const r=l.find(m=>m.id===2); if(!r||!r.result||r.result.tools.length<50) process.exit(1)' "$TMP/out" && ok "mcp tools/list" || bad "mcp tools/list"
else skipc "mcp tools/list"; fi

# 9. confirmation refused (non-TTY) exits 7
check "confirm refused exit 7" 7 env MS365_CLI_TOKEN=dummy $CLI todo delete L T < /dev/null && ok "confirm refused exit 7"

# 10. dry-run needs no client ID, no config and no home (DR AC5): exits 0 and writes nothing
mkdir -p "$TMP/empty-home" "$TMP/empty-cfg"
check "dry-run without client ID" 0 env HOME="$TMP/empty-home" MS365_CLI_CONFIG_DIR="$TMP/empty-cfg" MS365_CLI_CLIENT_ID= MS365_CLI_TOKEN= \
  $CLI --dry-run mail send --to a@b.com --subject s --body b && {
  if [ -n "$(find "$TMP/empty-home" "$TMP/empty-cfg" -mindepth 1)" ]; then bad "dry-run without client ID (wrote files)";
  else ok "dry-run without client ID"; fi; }

# 11. the dry-run trace is stderr-only for other families too (DR AC4): cal create -o json stays parseable
check "dry-run trace on stderr" 0 $CLI --dry-run calendar create --subject s --start 2026-10-20T10:00 --end 2026-10-20T11:00 --timezone UTC -o json && {
  if grep -q "\[dry-run\]" "$TMP/out"; then bad "dry-run trace leaked to stdout (cal create)";
  elif ! grep -q "\[dry-run\] POST https://graph.microsoft.com/v1.0/me/events" "$TMP/err"; then bad "cal create trace missing on stderr";
  elif node -e 'const d=JSON.parse(require("fs").readFileSync(process.argv[1],"utf8")); if(d.dryRun!==true||!d.body) process.exit(1)' "$TMP/out"; then ok "dry-run trace on stderr";
  else bad "cal create -o json stdout shape"; fi; }

# 12. a stdin password is redacted from the preview (DR AC3)
check "dry-run password redacted" 0 sh -c "printf 's3cret' | $CLI --dry-run drive share D I --scope anonymous --password-stdin" && {
  if grep -q "s3cret" "$TMP/out" "$TMP/err"; then bad "password leaked in dry-run output";
  elif grep -q "<redacted>" "$TMP/out" "$TMP/err"; then ok "dry-run password redacted";
  else bad "dry-run password not shown as <redacted>"; fi; }

# 13. default credential storage is the 0600 file (ID AC 2, E1): MS365_CLI_STORAGE unset, fresh config dir
mkdir -p "$TMP/store-cfg"
check "storage default is file" 0 env -u MS365_CLI_STORAGE MS365_CLI_CONFIG_DIR="$TMP/store-cfg" $CLI auth storage show -o json && {
  node -e 'const s=JSON.parse(require("fs").readFileSync(process.argv[1],"utf8")); if(s.configured!=="auto"||s.resolved!=="file"||s.backend!=="file") process.exit(1)' "$TMP/out" \
    && ok "storage default is file" || bad "storage default shape (want configured auto, resolved file)"; }

# 14. a hanging keychain tool never hangs the CLI (ID AC 3, lesson 6): stub secret-tool sleeps, the CLI kills it and exits 3
mkdir -p "$TMP/stub-bin" "$TMP/hang-cfg"
printf '#!/bin/sh\nexec sleep 47\n' > "$TMP/stub-bin/secret-tool"; chmod +x "$TMP/stub-bin/secret-tool"
t0=$(date +%s)
check "hanging keychain exits 3" 3 env MS365_CLI_STORAGE=keychain MS365_CLI_KEYCHAIN_BACKEND=libsecret PATH="$TMP/stub-bin:$PATH" \
  DBUS_SESSION_BUS_ADDRESS="unix:path=$TMP/no-bus" MS365_CLI_CONFIG_DIR="$TMP/hang-cfg" $CLI auth logout && {
  dt=$(( $(date +%s) - t0 ))
  if [ "$dt" -gt 15 ]; then bad "hanging keychain exits 3 (took ${dt}s)";
  elif ! grep -q "keychain_unavailable\|keychain is unavailable" "$TMP/err"; then bad "hanging keychain error message";
  elif pgrep -f "^sleep 47$" >/dev/null; then bad "hanging keychain left the stub running";
  else ok "hanging keychain exits 3 (${dt}s)"; fi; }

# 15. audit-trail import is offline and leak-free (AO-31, AC 14/16): no client ID, token or config, exits 0 with >=1
#     aggregated row, and stdout carries none of the fixture's subjects, InternetMessageIds or recipients
mkdir -p "$TMP/imp-home" "$TMP/imp-cfg"
check "audit-trail import offline" 0 env HOME="$TMP/imp-home" MS365_CLI_CONFIG_DIR="$TMP/imp-cfg" MS365_CLI_CLIENT_ID= MS365_CLI_TOKEN= \
  $CLI audit-trail import test/fixtures/purview-export.csv --cross-mailbox -o json && {
  if grep -qiE "SECRET-SUBJECT|SECRET-IMID|secret-recipient" "$TMP/out"; then bad "audit-trail import leaked message content to stdout";
  elif [ -n "$(find "$TMP/imp-home" "$TMP/imp-cfg" -mindepth 1 -name '*token*')" ]; then bad "audit-trail import touched the token cache";
  elif node -e 'const d=JSON.parse(require("fs").readFileSync(process.argv[1],"utf8")); const r=Array.isArray(d)?d:(d.value||d.rows||[]); if(r.length<1) process.exit(1)' "$TMP/out"; then ok "audit-trail import offline";
  else bad "audit-trail import -o json returned no rows"; fi; }

# --- AHT-6a: Gate-2 conformance pack (host mode, offline; ADR 0014 condition 2) ---------------------------------
# Every check below runs the built dist/ with the fetch-stub preload (test/fixtures/fetch-stub.mjs, E5): Graph answers
# come from MS365_TEST_FETCH_ROUTES and every fetch is logged to stderr as "[fetch] …", so "zero fetches" is provable.
# The result is keyed to the dist/ digest a host pins (--version is fixed at 0.1.0, R-29).
DIGEST="$(cd dist && find . -type f | LC_ALL=C sort | xargs sha256sum | sha256sum | cut -d' ' -f1)"
echo "conformance: dist sha256=$DIGEST"
STUB="node --import ./test/fixtures/fetch-stub.mjs"
HCLI="$STUB dist/index.js"
TENANT=11111111-2222-3333-4444-555555555555
jwt() { node -e 'const b=o=>Buffer.from(JSON.stringify(o)).toString("base64url");const c=JSON.parse(process.argv[1]);console.log(b({alg:"RS256",typ:"JWT"})+"."+b(c)+".sig")' "$1"; }
LIVE="$(jwt "{\"aud\":\"https://graph.microsoft.com\",\"tid\":\"$TENANT\",\"oid\":\"22222222-3333-4444-5555-666666666666\",\"upn\":\"ceo@contoso.com\",\"scp\":\"Mail.Read Calendars.Read User.Read\",\"exp\":$(( $(date +%s) + 3600 ))}")"
EXPIRED="$(jwt "{\"aud\":\"https://graph.microsoft.com\",\"tid\":\"$TENANT\",\"scp\":\"Mail.Read\",\"exp\":1}")"
OTHER="$(jwt "{\"aud\":\"https://graph.microsoft.com\",\"tid\":\"99999999-2222-3333-4444-555555555555\",\"scp\":\"Mail.Read\",\"exp\":$(( $(date +%s) + 3600 ))}")"
H="$TMP/host"; mkdir -p "$H/home" "$H/cfg/profiles"
printf '{"type":"host","tenantId":"%s","scopes":["User.Read","Mail.Read"]}\n' "$TENANT" > "$H/cfg/profiles/chief.json"
hrun() { # name expected_exit routes-json cmd-args... (host profile "chief", token $TOKEN)
  local name="$1" want="$2" routes="$3"; shift 3
  check "$name" "$want" env HOME="$H/home" MS365_CLI_CONFIG_DIR="$H/cfg" MS365_CLI_CLIENT_ID= MS365_CLI_TOKEN="${TOKEN-}" \
    MS365_TEST_FETCH_ROUTES="$routes" $HCLI --profile chief --error-format json "$@"
}
fetches() { grep -c '^\[fetch\]' "$TMP/err"; }
errcode() { node -e 'const l=require("fs").readFileSync(process.argv[1],"utf8").split("\n").filter(s=>s.trim()).pop();console.log(JSON.parse(l).error.code)' "$TMP/err" 2>/dev/null; }

# 16. ceiling refusal: a command outside the host profile's scopes exits 4 scope_ceiling with zero fetches, also under --dry-run
TOKEN="$LIVE"
hrun "ceiling refusal" 4 '[]' calendar list && {
  [ "$(errcode)" = scope_ceiling ] && [ "$(fetches)" = 0 ] && ok "ceiling refusal" || bad "ceiling refusal (code $(errcode), $(fetches) fetches)"; }
hrun "ceiling refusal --dry-run" 4 '[]' --dry-run calendar list && {
  [ "$(errcode)" = scope_ceiling ] && ok "ceiling refusal --dry-run" || bad "ceiling refusal --dry-run (code $(errcode))"; }

# 17. token checks: no token, expired, other tenant → exit 3 with the code, zero fetches
for row in "token_missing:" "token_expired:$EXPIRED" "token_identity_mismatch:$OTHER"; do
  want="${row%%:*}"; TOKEN="${row#*:}"
  hrun "token check $want" 3 '[]' mail list && {
    [ "$(errcode)" = "$want" ] && [ "$(fetches)" = 0 ] && ok "token check $want" || bad "token check $want (code $(errcode), $(fetches) fetches)"; }
done
TOKEN="$LIVE"

# 18. a host-mode read: immutable ids, _ref, complete:true; MSAL never loaded; nothing written to the config dir
hrun "host read --refs" 0 '[{"match":"/me/messages","body":{"value":[{"id":"m1","subject":"s","changeKey":"ck1","webLink":"https://outlook.office.com/m1"}]}},{"match":"/me\\?","body":{"id":"22222222-3333-4444-5555-666666666666","userPrincipalName":"ceo@contoso.com","mail":"ceo@contoso.com"}}]' \
  --refs mail list -o json && {
  node -e 'const o=JSON.parse(require("fs").readFileSync(process.argv[1],"utf8"));const r=o.value[0]._ref;if(o.complete!==true||!r||r.id!=="m1"||r.version!=="ck1"||r.tenantId!==process.argv[2])process.exit(1)' "$TMP/out" "$TENANT" \
    && [ -z "$(find "$H/cfg" -type f ! -name chief.json ! -name audit.log)" ] && ok "host read --refs" || bad "host read --refs (_ref/complete or files written)"; }
check "auth status -o json (external, no MSAL)" 0 env HOME="$H/home" MS365_CLI_CONFIG_DIR="$H/cfg" MS365_CLI_TOKEN="$LIVE" \
  node --import ./test/fixtures/msal-probe.mjs dist/index.js --profile chief auth status -o json && {
  node -e 'const s=JSON.parse(require("fs").readFileSync(process.argv[1],"utf8"));if(s.source!=="external"||s.type!=="host"||s.ok!==true||JSON.stringify(s.ceiling)!==JSON.stringify(["User.Read","Mail.Read"]))process.exit(1)' "$TMP/out" \
    && ! grep -q '\[probe\] msal-node loaded' "$TMP/err" && ok "auth status -o json (external, no MSAL)" || bad "auth status -o json shape or MSAL loaded"; }

# 19. feature probe: describe --json (offline) shows the agent-host surface
check "describe feature probe" 0 env HOME="$H/home" MS365_CLI_CONFIG_DIR="$TMP/empty-cfg" $CLI describe --json && {
  node -e 'const d=JSON.parse(require("fs").readFileSync(process.argv[1],"utf8"));const f=o=>o.map(x=>x.flags).join(" ");const c=p=>d.commands.find(x=>x.path===p);
    if(!f(d.globalOptions).includes("--refs")||!/--checkpoint/.test(f(c("watch").options))||!/--select/.test(f(c("watch").options))||!c("health")||!/--output/.test(f(c("auth status").options)))process.exit(1)' "$TMP/out" \
    && ok "describe feature probe" || bad "describe feature probe (missing --refs/--checkpoint/--select/health/auth status -o)"; }

# 20. watch --checkpoint: a baseline ends with one checkpoint line; a checkpoint for events refused on mail before any fetch
hrun "watch --checkpoint new" 0 '[{"match":"/messages/delta","body":{"value":[],"@odata.deltaLink":"https://graph.microsoft.com/v1.0/me/mailFolders/inbox/messages/delta?$deltatoken=t1"}}]' \
  watch mail --checkpoint new && {
  node -e 'const l=require("fs").readFileSync(process.argv[1],"utf8").trim().split("\n").map(s=>JSON.parse(s));const c=l.filter(x=>x.type==="checkpoint");if(c.length!==1||l[l.length-1].type!=="checkpoint"||!c[0].checkpoint)process.exit(1)' "$TMP/out" \
    && ok "watch --checkpoint new" || bad "watch --checkpoint new (no single trailing checkpoint line)"; }
EVCP="$(node -e 'console.log(Buffer.from(JSON.stringify({v:1,resource:"events",endpoint:"/me/calendarView/delta",deltaLink:"https://graph.microsoft.com/v1.0/me/calendarView/delta?$deltatoken=e",syncedAt:"2026-10-17T00:00:00Z",coverageStart:"2026-10-01T00:00:00Z"})).toString("base64url"))')"
hrun "checkpoint_mismatch" 2 '[]' watch mail --checkpoint "$EVCP" && {
  [ "$(errcode)" = checkpoint_mismatch ] && [ "$(fetches)" = 0 ] && ok "checkpoint_mismatch" || bad "checkpoint_mismatch (code $(errcode), $(fetches) fetches)"; }

# 21. health mapping rows through the built CLI: each probe answer maps to its connector-health status
ME='{"match":"/me\\?","body":{"id":"x"}}'
for row in "HEALTHY::200" "STALE:graph_error:503" "AUTH_REQUIRED:access_denied:403" "RATE_LIMITED:throttled:429"; do
  IFS=: read -r st rs code <<<"$row"
  hrun "health $st" 0 "[$ME,{\"match\":\"/me/messages\",\"status\":$code,\"headers\":{\"Retry-After\":\"1\"},\"body\":{\"value\":[]}}]" health mail events && {
    [ "$st" = HEALTHY ] && rs=null
    node -e 'const a=JSON.parse(require("fs").readFileSync(process.argv[1],"utf8"));const m=a.find(x=>x.resource==="mail"),e=a.find(x=>x.resource==="events");const want=process.argv[3]==="null"?null:process.argv[3];
      if(m.status!==process.argv[2]||m.reason!==want||e.status!=="UNSUPPORTED"||e.reason!=="scope_ceiling")process.exit(1)' "$TMP/out" "$st" "$rs" \
      && ! grep -q '/me/events' "$TMP/err" && ok "health $st (+ events UNSUPPORTED scope_ceiling)" || bad "health $st row"; }
done
t0=$(date +%s)
check "health timeout probe" 0 env HOME="$H/home" MS365_CLI_CONFIG_DIR="$H/cfg" MS365_CLI_TOKEN="$LIVE" MS365_CLI_REQUEST_TIMEOUT_MS=1000 \
  MS365_TEST_FETCH_ROUTES="[$ME,{\"match\":\"/me/messages\",\"hang\":true}]" $HCLI --profile chief health mail && {
  dt=$(( $(date +%s) - t0 ))
  node -e 'const m=JSON.parse(require("fs").readFileSync(process.argv[1],"utf8"))[0];if(m.status!=="STALE"||m.reason!=="timeout")process.exit(1)' "$TMP/out" \
    && [ "$(grep -c '/me/messages' "$TMP/err")" = 1 ] && [ "$dt" -le 5 ] && ok "health timeout probe (${dt}s, 1 attempt)" || bad "health timeout probe (${dt}s)"; }

# 22. request budget: MS365_CLI_MAX_RETRIES=0 → one fetch, exit 1 graph_error 503; out of range → exit 2
check "MAX_RETRIES=0 503" 1 env HOME="$H/home" MS365_CLI_CONFIG_DIR="$H/cfg" MS365_CLI_TOKEN="$LIVE" MS365_CLI_MAX_RETRIES=0 \
  MS365_TEST_FETCH_ROUTES='[{"match":"/me","status":503}]' $HCLI --profile chief --error-format json user me && {
  [ "$(errcode)" = graph_error ] && [ "$(fetches)" = 1 ] && ok "MAX_RETRIES=0 503" || bad "MAX_RETRIES=0 503 (code $(errcode), $(fetches) fetches)"; }
check "budget out of range" 2 env MS365_CLI_REQUEST_TIMEOUT_MS=10 $CLI --error-format json describe --json && ok "budget out of range"

# 23. AHT-15 contract fixture: regenerating test/fixtures/contract/ from this dist/ (host mode, fetch-stub, frozen
#     clock) gives the committed bytes; any contract change must regenerate the fixture in the same commit
check "contract fixture" 0 node scripts/contract.mjs --check && ok "contract fixture"

# 24. AHT-9 (OQ-8 b): without a terminal, an approved calendar delete plan is applied under a type: host profile
#     (exactly 1 DELETE; its If-Match is asserted in test/calendar-plans.test.ts); the same plan under a user profile
#     exits 7 confirmation_refused with zero fetches
printf '{"type":"host","tenantId":"%s","scopes":["User.Read","Calendars.ReadWrite"]}\n' "$TENANT" > "$H/cfg/profiles/chief-cal.json"
printf '{"type":"user","scopes":["User.Read","Calendars.ReadWrite"]}\n' > "$H/cfg/profiles/ceo-user.json"
CALTOK="$(jwt "{\"aud\":\"https://graph.microsoft.com\",\"tid\":\"$TENANT\",\"oid\":\"22222222-3333-4444-5555-666666666666\",\"scp\":\"Calendars.ReadWrite User.Read\",\"exp\":$(( $(date +%s) + 3600 ))}")"
EV='[{"method":"GET","match":"/me/events/E1\\?","body":{"id":"E1","type":"singleInstance","isOrganizer":true,"@odata.etag":"W/\"ek1\"","attendees":[]}},{"method":"DELETE","match":"/me/events/E1$","status":204}]'
crun() { # name expected_exit profile cmd-args... (calendar token, stdin not a terminal; MS365_CLI_PROFILE names the
         # host profile, so the user-profile row also proves --profile wins over the environment)
  local name="$1" want="$2" prof="$3"; shift 3
  check "$name" "$want" env HOME="$H/home" MS365_CLI_CONFIG_DIR="$H/cfg" MS365_CLI_CLIENT_ID= MS365_CLI_TOKEN="$CALTOK" \
    MS365_CLI_PROFILE=chief-cal MS365_TEST_FETCH_ROUTES="$EV" $HCLI --profile "$prof" --error-format json "$@" </dev/null
}
crun "plan calendar delete (host)" 0 chief-cal --plan "$TMP/del.json" calendar delete E1 && ok "plan calendar delete (host)"
DIG="$(node -e 'console.log(JSON.parse(require("fs").readFileSync(process.argv[1],"utf8")).digest)' "$TMP/del.json" 2>/dev/null)"
crun "apply delete non-TTY, user profile" 7 ceo-user apply "$TMP/del.json" --expect-digest "$DIG" && {
  [ "$(errcode)" = confirmation_refused ] && [ "$(fetches)" = 0 ] && ok "apply delete non-TTY, user profile (exit 7, 0 fetches)" || bad "apply delete non-TTY, user profile (code $(errcode), $(fetches) fetches)"; }
crun "apply delete non-TTY, host profile" 0 chief-cal apply "$TMP/del.json" --expect-digest "$DIG" && {
  [ "$(fetches)" = 1 ] && grep -q '^\[fetch\] DELETE .*/me/events/E1$' "$TMP/err" && ok "apply delete non-TTY, host profile (1 DELETE)" || bad "apply delete non-TTY, host profile ($(fetches) fetches)"; }

echo "smoke: pass=$pass fail=$fail skip=$skip"
[ "$fail" = 0 ]
