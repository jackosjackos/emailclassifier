# Email Classifier (Stage 1 — classification only)

Polls a shared Microsoft 365 mailbox every 15 minutes via Microsoft Graph,
scores each new inbound email with a heuristic rules engine, applies an
Outlook category (`Ticket`, `Needs Review`, or `Ignored`), and generates an
HTML report so you can validate the classifier's decisions before any
Accelo ticket-creation logic is added.

**This stage does not call the Accelo API at all.** It only classifies and
categorizes. Ticket creation is a deliberate stage 2, once the scoring here
is validated against your real inbox traffic.

## How it works

- Runs on a schedule via GitHub Actions (`*/15 * * * *`), no server to host.
- Uses Microsoft Graph's **delta query** so each run only sees messages
  that are new or changed since the last run — not a full inbox re-scan.
- Delta state and the HTML report are committed back into this repo by the
  workflow, so state persists between runs without any external database.
- Messages that already carry one of the marker categories (`Ticket`,
  `Needs Review`, `Ignored`) are skipped — safe to re-run without
  reclassifying, and safe if a human manually adjusts a category.

## One-time setup

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

This step is not optional hardening — without it, a compromised secret for
this app would expose every mailbox in the organisation, not just the one
it's meant to read.

### 2. Add repo secrets

In your GitHub repo → Settings → Secrets and variables → Actions → **Secrets** tab:

| Secret name | Value |
|---|---|
| `GRAPH_TENANT_ID` | Directory (tenant) ID from step 1 |
| `GRAPH_CLIENT_ID` | Application (client) ID from step 1 |
| `GRAPH_CLIENT_SECRET` | The client secret value from step 1 |
| `MAILBOX_USER_ID` | The shared mailbox's UPN, e.g. `support@yourdomain.com` |

### 3. Add repo variables (non-sensitive tuning knobs)

Same location, **Variables** tab:

| Variable name | Example value | Purpose |
|---|---|---|
| `INTERNAL_DOMAIN` | `yourcompany.com` | Slight negative weight for internal senders |
| `KNOWN_CLIENT_DOMAINS` | `clienta.com,clientb.com` | Positive weight for known client senders |
| `HIGH_CONFIDENCE_THRESHOLD` | `25` | Score at/above which a message is marked `Ticket` |
| `LOW_CONFIDENCE_THRESHOLD` | `-25` | Score at/below which a message is marked `Ignored` |

If you skip the last two, the script defaults to `25` and `-25`.

### 4. Enable Actions and run it once manually

Push this repo to GitHub, go to the **Actions** tab, select
**Classify Inbox Emails**, and click **Run workflow** to trigger it manually
the first time rather than waiting for the schedule. Check the run logs and
`report/latest.html` afterward.

## Viewing the report

Two ways, pick whichever's easier:

1. **Committed file** — `report/latest.html` in the repo updates every run.
   Open it locally, or enable GitHub Pages pointed at the repo root/`report`
   folder to view it at a URL without downloading anything.
2. **Workflow artifact** — every run also uploads `report/latest.html` as a
   downloadable artifact from the Actions run summary page (kept 14 days).

## Tuning the classifier

The scoring logic lives entirely in `scripts/classify-emails.js`, in the
`scoreEmail()` function. **The keyword lists and weights shipped are
placeholders** — they will not be accurate for your inbox on day one.

To tune it:

1. Let it run for a few days in `Needs Review`-heavy mode (conservative
   thresholds) so you're not silently auto-ignoring real support requests
   or auto-ticketing noise while you're still calibrating.
2. Look at `report/latest.html` (and `report/history/` for past runs)
   against what you know is actually signal vs. noise.
3. Adjust the `noiseKeywords` / `supportKeywords` arrays, the
   `KNOWN_CLIENT_DOMAINS` variable, and the header checks
   (`List-Unsubscribe`, `Auto-Submitted`) to match real patterns you're
   seeing.
4. Widen or narrow `HIGH_CONFIDENCE_THRESHOLD` / `LOW_CONFIDENCE_THRESHOLD`
   based on how large your "Needs Review" bucket ends up being — too narrow
   a gap between the thresholds pushes almost everything into manual review,
   which defeats the point; too wide risks false positives/negatives.

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
- **Human review resolution.** A second job that watches for a category
  change made manually on a `Needs Review` item (to `Ticket` or `Ignored`)
  and acts on it — this is what makes the daily human clearing of
  `Needs Review` actually flow through to a ticket.

## Known limitations

- **Schedule timing isn't exact.** GitHub Actions' `schedule` trigger is
  best-effort and can be delayed under platform load. Fine for this use
  case given daily human review of the ambiguous bucket regardless.
- **Delta link expiry (410 Gone).** Graph delta links can occasionally
  expire and force a full resync. The script handles this automatically
  (falls back to a fresh delta query), but a resync means every message
  currently in the inbox gets re-evaluated — messages already carrying a
  marker category are skipped, so this is safe, just slightly slower on
  that one run.
- **Committing state via git** means every run creates a commit if
  anything changed. That's expected and intentional — it's what gives you
  a free, auditable history of every classification decision in
  `report/history/`.
