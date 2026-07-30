# Email Classifier (Stage 1 — classification only, runs locally)

Polls a shared Microsoft 365 mailbox via Microsoft Graph, scores each
inbound email with a heuristic rules engine, applies an Outlook category
(`Ticket`, `Needs Review`, or `Ignored`), and generates an HTML report so
you can validate the classifier's decisions before any Accelo
ticket-creation logic is added.

**Runs entirely on your machine.** No GitHub Actions, no cloud scheduler,
no CI. Just Node and a terminal. State and reports are saved to local
files in this folder between runs.

**This stage does not call the Accelo API at all.** It only classifies and
categorizes. Ticket creation is a deliberate stage 2, once the scoring here
is validated against your real inbox traffic.

## Requirements

- Node.js 18 or later (check with `node -v`)
- An Azure AD app registration with Graph API access to the mailbox (see below)

## One-time setup

**If you only want to run the PST test (Option A above), skip straight to
step 2 (install dependencies) — you don't need an Azure AD app or any
credentials for that mode.** Steps 1 and 3 below are only needed once you
move to testing against or running against a live mailbox.

### 1. Register an Azure AD app for Graph API access

In the Azure Portal → Azure Active Directory → App registrations → New registration:

1. Name it something like `email-classifier`.
2. Under **API permissions**, add **Application permission** (not delegated):
   `Mail.ReadWrite`. Grant admin consent.
3. Under **Certificates & secrets**, create a new client secret. Copy the
   value immediately — you can't retrieve it again later.
4. Note down the **Application (client) ID** and **Directory (tenant) ID**
   from the app's Overview page.

**Important — scope the permission down.** `Mail.ReadWrite` as an
application permission grants access to *every* mailbox in your tenant by
default. Restrict it to only the shared mailbox using an Exchange Online
**application access policy**:

```powershell
# Run in Exchange Online PowerShell, replacing the placeholders
New-ApplicationAccessPolicy `
  -AppId "<your-app-client-id>" `
  -PolicyScopeGroupId "shared-mailbox@yourdomain.com" `
  -AccessRight RestrictAccess `
  -Description "Restrict email-classifier app to the shared support mailbox only"
```

This step is not optional hardening — without it, a leaked secret for this
app would expose every mailbox in the organisation, not just the one it's
meant to read. This matters more, not less, for a local setup, since the
secret now lives in a `.env` file on a laptop instead of a managed secrets
store.

### 2. Install dependencies

```bash
npm install
```

### 3. Configure your environment

```bash
cp .env.example .env
```

Open `.env` and fill in:

| Variable | Value |
|---|---|
| `GRAPH_TENANT_ID` | Directory (tenant) ID from step 1 |
| `GRAPH_CLIENT_ID` | Application (client) ID from step 1 |
| `GRAPH_CLIENT_SECRET` | The client secret value from step 1 |
| `MAILBOX_USER_ID` | The shared mailbox's UPN, e.g. `support@yourdomain.com` |

`.env` is gitignored — it never gets committed, and this whole project
doesn't need to be pushed to GitHub at all if you don't want it to be.

## Running a test pass first (recommended before anything else)

There are two ways to test the scoring logic before it ever touches a live
mailbox. Pick whichever fits what you have on hand.

### Option A: Test against a .pst export (no Azure app registration needed)

If you've exported a `.pst` file from Outlook (Settings → Files → Export in
New Outlook for Windows, or the classic Import/Export wizard), you can
score its contents directly — **this mode makes no network calls at all**
and needs none of the `GRAPH_*` credentials:

```bash
npm run test-pst -- path/to/export.pst
```

(or set `PST_FILE_PATH` in `.env` and just run `npm run test-pst`)

This reads the most recent messages (default 100, configurable via
`PST_SAMPLE_SIZE`) out of the `Inbox` folder inside the PST (configurable
via `PST_FOLDER_NAME` if your export uses a different folder name or a
non-English Outlook locale), scores them with the same heuristics used
against a live mailbox, and opens `report/latest-pst.html`.

This mode is always read-only/offline by nature — there's no live mailbox
connection to apply categories to, so nothing about your PST or a real
inbox is ever modified. It's genuinely the lowest-friction way to sanity
check the scoring logic, since it skips the Azure AD app registration
entirely.

**Note on what's available in a PST vs. Graph.** Both give the same
information used for scoring — subject, sender, date, and full raw email
headers (including things like `List-Unsubscribe` and `Auto-Submitted`,
which some of the heuristics check). Outlook categories aren't meaningfully
carried over from a static export, so every message shows as freshly
scored rather than "already handled."

### Option B: Test against your live mailbox via Graph (no export step)

This needs the Azure AD app registration below, but skips the manual
export/import round-trip — it pulls the last 100 emails directly:

```bash
npm run test-run
```

This opens `report/latest.html` in your browser automatically when it
finishes. Review the scores and reasons against emails you already know
should or shouldn't become tickets.

To change the sample size, add to `.env`:

```
TEST_SAMPLE_SIZE=100
```

If you want a test run to actually apply categories to those 100 real
emails (so you can go look at them in Outlook and see the labels), override
the dry-run default explicitly:

```
DRY_RUN=false
```

I'd hold off on doing that until the scoring in `report/latest.html` looks
right on a few dry runs first — it's easy to relabel emails, less easy to
un-mislabel 100 of them by hand if the heuristics are off.

## Running for real (single run)

Once you're happy with the scoring:

```bash
npm run classify
```

This uses Graph's **delta query**, so it only processes messages that are
new or changed since the last run — not a full inbox re-scan — and it
**does** apply categories (no dry run in normal mode). Run this manually
whenever you want a check, or schedule it (see below).

## Keeping it running continuously

Two options, pick whichever fits how you work:

**A. OS-level scheduler calling `npm run classify` (recommended for "set and forget")**

- **macOS/Linux (cron):**
  ```
  */15 * * * * cd /path/to/email-classifier && /usr/local/bin/node scripts/classify-emails.js >> classifier.log 2>&1
  ```
  Edit with `crontab -e`. Use `which node` to get the correct node path — cron
  doesn't inherit your shell's PATH.

- **Windows (Task Scheduler):** create a task that runs
  `node.exe scripts/classify-emails.js` with "Start in" set to this
  project's folder, triggered every 15 minutes.

With either, note `OPEN_REPORT=true` will try to pop a browser window every
run — fine on a machine you're at, but set `OPEN_REPORT=false` in `.env` if
this runs unattended or on a server.

**B. Built-in watch mode (one Node process, runs indefinitely)**

```bash
npm run watch
```

Runs immediately, then re-runs on the schedule set by `WATCH_CRON` in
`.env` (default every 15 minutes), all within one long-lived process. Stop
with `Ctrl+C`. Good for a machine that's already always-on (e.g. a mini PC
or a spare desktop), less good if you want it to survive a reboot without
you remembering to restart it — pair with `pm2` or a similar process
manager if you want that, or just use Option A instead.

## Viewing reports

- `report/latest.html` — most recent live-mailbox or `test-run` result,
  opened automatically after each run unless `OPEN_REPORT=false`.
- `report/latest-pst.html` — most recent `test-pst` result. Kept separate
  from `latest.html` so a PST test run never overwrites your live-mailbox
  report, or vice versa.
- `report/history/` — a timestamped copy from every run of either kind, so
  you can look back at how classification looked over time as you tune it.

All of these are local files. Nothing is uploaded anywhere in either mode.

## Tuning the classifier

The scoring logic lives entirely in `scripts/classify-emails.js`, in the
`scoreEmail()` function. **The keyword lists and weights shipped are
placeholders** — they will not be accurate for your inbox on day one.

To tune it:

1. Run `npm run test-run` against your last 100 emails.
2. Compare `report/latest.html` against what you already know is signal
   vs. noise in that inbox.
3. Adjust `noiseKeywords` / `supportKeywords`, the `KNOWN_CLIENT_DOMAINS`
   env variable, and the header checks (`List-Unsubscribe`,
   `Auto-Submitted`) to match real patterns you're seeing.
4. Adjust `HIGH_CONFIDENCE_THRESHOLD` / `LOW_CONFIDENCE_THRESHOLD` in
   `.env` based on how large your "Needs Review" bucket ends up being —
   too narrow a gap pushes almost everything into manual review, which
   defeats the point; too wide risks false positives/negatives.
5. Re-run `npm run test-run` and repeat until it looks right.

## What's intentionally not built yet

- **Accelo ticket creation.** Once categories are trustworthy, stage 2 adds
  a step that reads messages categorized `Ticket` and creates the Accelo
  ticket via `POST /api/v0/tickets`, including Accelo contact/company
  matching.
- **Reply threading.** Detecting that an inbound message is a reply to an
  already-ticketed conversation (via `conversationId`) and posting it as a
  note on the existing ticket rather than creating a duplicate. This needs
  to exist before ticket creation goes live, or every customer follow-up
  becomes a new ticket.
- **Human review resolution.** A step that watches for a category change
  made manually on a `Needs Review` item (to `Ticket` or `Ignored`) and
  acts on it — this is what makes the daily human clearing of
  `Needs Review` actually flow through to a ticket once stage 2 exists.

## Known limitations

- **Delta link expiry (410 Gone).** Graph delta links can occasionally
  expire and force a full resync. The script handles this automatically,
  but a resync means every message currently in the inbox gets
  re-evaluated — messages already carrying a marker category are skipped,
  so this is safe, just slightly slower on that one run.
- **`npm run watch` doesn't survive a machine restart or crash** on its
  own. Use an OS scheduler (Option A above) or a process manager like
  `pm2` if you need that guarantee.
- **Local secrets.** The Graph client secret lives in a local `.env` file.
  Keep this machine's disk encrypted and don't share the `.env` file —
  it's gitignored specifically so it can't end up in version control by
  accident.
