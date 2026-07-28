/**
 * Email Classifier — Stage 1 (classification only, no Accelo integration yet)
 *
 * What this does:
 *   1. Authenticates to Microsoft Graph via client-credentials OAuth2.
 *   2. Polls the shared mailbox inbox using a delta query, so each run only
 *      sees messages that are new or changed since the last run.
 *   3. Skips anything that already carries one of our marker categories
 *      (idempotency — safe to re-run without reclassifying).
 *   4. Scores each remaining message with a heuristic rules engine.
 *   5. Applies an Outlook category based on the score: "Ticket",
 *      "Needs Review", or "Ignored".
 *   6. Writes an HTML report of everything seen this run to report/latest.html
 *      (plus a timestamped copy in report/history/) so you can validate the
 *      classifier's decisions before any ticket-creation logic is added.
 *
 * Deliberately NOT included yet:
 *   - Any call to the Accelo API.
 *   - Contact/company matching.
 *   - Reply-threading detection.
 * These come in stage 2, once the scoring here has been validated against
 * real inbox traffic.
 */

const axios = require('axios');
const fs = require('fs');
const path = require('path');

// ---------------------------------------------------------------------------
// Configuration
// ---------------------------------------------------------------------------

const TENANT_ID = requireEnv('GRAPH_TENANT_ID');
const CLIENT_ID = requireEnv('GRAPH_CLIENT_ID');
const CLIENT_SECRET = requireEnv('GRAPH_CLIENT_SECRET');
const MAILBOX_USER_ID = requireEnv('MAILBOX_USER_ID'); // shared mailbox UPN or object ID

// Optional tuning inputs — safe defaults if not set, but you should set these
// via repo Variables (not Secrets, since they're not sensitive) once you know
// your real domains/keywords.
const INTERNAL_DOMAIN = (process.env.INTERNAL_DOMAIN || '').toLowerCase();
const KNOWN_CLIENT_DOMAINS = (process.env.KNOWN_CLIENT_DOMAINS || '')
  .split(',')
  .map(d => d.trim().toLowerCase())
  .filter(Boolean);

const HIGH_CONFIDENCE_THRESHOLD = Number(process.env.HIGH_CONFIDENCE_THRESHOLD || 25);
const LOW_CONFIDENCE_THRESHOLD = Number(process.env.LOW_CONFIDENCE_THRESHOLD || -25);

const STATE_DIR = path.join(__dirname, '..', 'state');
const DELTA_LINK_FILE = path.join(STATE_DIR, 'delta-link.json');
const REPORT_DIR = path.join(__dirname, '..', 'report');
const REPORT_HISTORY_DIR = path.join(REPORT_DIR, 'history');
const LATEST_REPORT_FILE = path.join(REPORT_DIR, 'latest.html');

// Marker categories — presence of ANY of these means "already handled, skip".
const MARKER_CATEGORIES = ['Ticket', 'Needs Review', 'Ignored'];

function requireEnv(name) {
  const value = process.env[name];
  if (!value) {
    console.error(`Missing required environment variable: ${name}`);
    process.exit(1);
  }
  return value;
}

// ---------------------------------------------------------------------------
// Graph authentication (client-credentials flow)
// ---------------------------------------------------------------------------

async function getGraphAccessToken() {
  const url = `https://login.microsoftonline.com/${TENANT_ID}/oauth2/v2.0/token`;
  const body = new URLSearchParams({
    client_id: CLIENT_ID,
    client_secret: CLIENT_SECRET,
    scope: 'https://graph.microsoft.com/.default',
    grant_type: 'client_credentials'
  });

  try {
    const { data } = await axios.post(url, body, {
      headers: { 'Content-Type': 'application/x-www-form-urlencoded' }
    });
    return data.access_token;
  } catch (err) {
    throw new Error(
      `Graph auth failed: ${err.response?.status} ${JSON.stringify(err.response?.data || err.message)}`
    );
  }
}

// ---------------------------------------------------------------------------
// Delta query state (persisted to a committed file between workflow runs)
// ---------------------------------------------------------------------------

function loadDeltaLink() {
  try {
    const raw = fs.readFileSync(DELTA_LINK_FILE, 'utf8');
    return JSON.parse(raw).deltaLink || null;
  } catch {
    return null; // first run, or file doesn't exist yet — that's fine
  }
}

function saveDeltaLink(deltaLink) {
  fs.mkdirSync(STATE_DIR, { recursive: true });
  fs.writeFileSync(DELTA_LINK_FILE, JSON.stringify({ deltaLink, updatedAt: new Date().toISOString() }, null, 2));
}

// ---------------------------------------------------------------------------
// Fetch messages via delta query, following pagination
// ---------------------------------------------------------------------------

async function fetchDeltaMessages(accessToken) {
  const storedDeltaLink = loadDeltaLink();

  // $select keeps the payload light — only fetch what scoring/reporting needs.
  const selectFields = [
    'id', 'subject', 'from', 'receivedDateTime', 'categories',
    'conversationId', 'internetMessageId', 'bodyPreview', 'internetMessageHeaders'
  ].join(',');

  let url = storedDeltaLink ||
    `https://graph.microsoft.com/v1.0/users/${MAILBOX_USER_ID}/mailFolders/inbox/messages/delta?$select=${selectFields}`;

  const allMessages = [];
  let finalDeltaLink = storedDeltaLink;

  while (url) {
    let response;
    try {
      response = await axios.get(url, {
        headers: { Authorization: `Bearer ${accessToken}` }
      });
    } catch (err) {
      // A stale/invalid delta link returns 410 Gone — Graph's documented way
      // of saying "resync from scratch". Handle it rather than crashing.
      if (err.response?.status === 410) {
        console.warn('Delta link expired (410 Gone) — resetting to a fresh full sync.');
        fs.rmSync(DELTA_LINK_FILE, { force: true });
        return fetchDeltaMessages(accessToken); // retry once, from scratch
      }
      throw new Error(`Graph delta query failed: ${err.response?.status} ${JSON.stringify(err.response?.data || err.message)}`);
    }

    allMessages.push(...(response.data.value || []));

    if (response.data['@odata.nextLink']) {
      url = response.data['@odata.nextLink']; // more pages
    } else {
      url = null;
      finalDeltaLink = response.data['@odata.deltaLink']; // save for next run
    }
  }

  return { messages: allMessages, deltaLink: finalDeltaLink };
}

// ---------------------------------------------------------------------------
// Heuristic scoring engine
//
// THIS IS A SCAFFOLD. The keyword lists and weights below are placeholders —
// they need tuning against real examples from your inbox before this is
// trustworthy. Treat the first several runs as calibration, not production.
// ---------------------------------------------------------------------------

function scoreEmail(message) {
  let score = 0;
  const reasons = [];

  const senderAddress = (message.from?.emailAddress?.address || '').toLowerCase();
  const senderDomain = senderAddress.split('@')[1] || '';
  const subject = (message.subject || '').toLowerCase();
  const headers = message.internetMessageHeaders || [];

  // --- Negative signals (likely noise) ---
  if (headers.some(h => h.name?.toLowerCase() === 'list-unsubscribe')) {
    score -= 50;
    reasons.push('has List-Unsubscribe header (likely newsletter/marketing)');
  }

  if (headers.some(h => h.name?.toLowerCase() === 'auto-submitted' && h.value?.toLowerCase() !== 'no')) {
    score -= 40;
    reasons.push('Auto-Submitted header present (likely automated notification)');
  }

  if (INTERNAL_DOMAIN && senderDomain === INTERNAL_DOMAIN) {
    score -= 10;
    reasons.push('sender is internal domain');
  }

  const noiseKeywords = ['out of office', 'automatic reply', 'newsletter', 'unsubscribe', 'calendar invite', 'meeting accepted', 'meeting declined'];
  const matchedNoise = noiseKeywords.filter(k => subject.includes(k));
  if (matchedNoise.length) {
    score -= 40;
    reasons.push(`subject matches noise keyword(s): ${matchedNoise.join(', ')}`);
  }

  // --- Positive signals (likely support) ---
  const supportKeywords = ['issue', 'problem', 'help', 'not working', 'error', 'request', 'urgent', 'broken', 'question'];
  const matchedSupport = supportKeywords.filter(k => subject.includes(k));
  if (matchedSupport.length) {
    score += 20;
    reasons.push(`subject matches support keyword(s): ${matchedSupport.join(', ')}`);
  }

  if (KNOWN_CLIENT_DOMAINS.includes(senderDomain)) {
    score += 30;
    reasons.push('sender domain is a known client');
  }

  return { score, reasons };
}

function classify(score) {
  if (score >= HIGH_CONFIDENCE_THRESHOLD) return 'Ticket';
  if (score <= LOW_CONFIDENCE_THRESHOLD) return 'Ignored';
  return 'Needs Review';
}

// ---------------------------------------------------------------------------
// Apply category back to the message in Outlook
// ---------------------------------------------------------------------------

async function applyCategory(messageId, accessToken, categoryName, existingCategories) {
  // Preserve any categories a human/other process already applied that
  // aren't one of our markers, rather than clobbering them.
  const preserved = (existingCategories || []).filter(c => !MARKER_CATEGORIES.includes(c));
  const newCategories = [...preserved, categoryName];

  await axios.patch(
    `https://graph.microsoft.com/v1.0/users/${MAILBOX_USER_ID}/messages/${messageId}`,
    { categories: newCategories },
    { headers: { Authorization: `Bearer ${accessToken}`, 'Content-Type': 'application/json' } }
  );
}

// ---------------------------------------------------------------------------
// HTML report generation
// ---------------------------------------------------------------------------

function escapeHtml(str) {
  return String(str || '')
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#039;');
}

function decisionBadgeClass(decision) {
  if (decision === 'Ticket') return 'badge-ticket';
  if (decision === 'Ignored') return 'badge-ignored';
  if (decision === 'skipped') return 'badge-skipped';
  return 'badge-review';
}

function buildHtmlReport(results, runMetadata) {
  const rows = results.map(r => `
    <tr>
      <td>${escapeHtml(r.receivedDateTime)}</td>
      <td>${escapeHtml(r.from)}</td>
      <td class="subject-cell">${escapeHtml(r.subject)}</td>
      <td class="score-cell">${r.score !== null ? r.score : '—'}</td>
      <td><span class="badge ${decisionBadgeClass(r.decision)}">${escapeHtml(r.decision)}</span></td>
      <td class="reasons-cell">${escapeHtml((r.reasons || []).join('; ') || '—')}</td>
    </tr>
  `).join('');

  const counts = results.reduce((acc, r) => {
    acc[r.decision] = (acc[r.decision] || 0) + 1;
    return acc;
  }, {});

  return `<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<title>Email Classification Report — ${escapeHtml(runMetadata.runAt)}</title>
<style>
  :root {
    --ink: #1a1a1a;
    --muted: #6b6b6b;
    --border: #e4e4e4;
    --bg: #fafafa;
    --ticket: #1e6b3f;
    --ticket-bg: #e6f4ea;
    --review: #8a5a00;
    --review-bg: #fff4e0;
    --ignored: #6b6b6b;
    --ignored-bg: #ececec;
    --skipped: #4a5a8a;
    --skipped-bg: #e7eaf6;
  }
  * { box-sizing: border-box; }
  body {
    font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Helvetica, Arial, sans-serif;
    background: var(--bg);
    color: var(--ink);
    margin: 0;
    padding: 2rem;
  }
  .container { max-width: 1100px; margin: 0 auto; }
  h1 { font-size: 1.4rem; margin-bottom: 0.25rem; }
  .meta { color: var(--muted); font-size: 0.85rem; margin-bottom: 1.5rem; }
  .summary {
    display: flex;
    gap: 0.75rem;
    margin-bottom: 1.5rem;
    flex-wrap: wrap;
  }
  .summary-card {
    background: white;
    border: 1px solid var(--border);
    border-radius: 8px;
    padding: 0.75rem 1rem;
    min-width: 110px;
  }
  .summary-card .n { font-size: 1.4rem; font-weight: 600; }
  .summary-card .l { font-size: 0.75rem; color: var(--muted); text-transform: uppercase; letter-spacing: 0.03em; }
  table {
    width: 100%;
    border-collapse: collapse;
    background: white;
    border: 1px solid var(--border);
    border-radius: 8px;
    overflow: hidden;
  }
  th, td {
    text-align: left;
    padding: 0.6rem 0.75rem;
    border-bottom: 1px solid var(--border);
    font-size: 0.85rem;
    vertical-align: top;
  }
  th {
    background: #f2f2f2;
    font-size: 0.72rem;
    text-transform: uppercase;
    letter-spacing: 0.03em;
    color: var(--muted);
  }
  tr:last-child td { border-bottom: none; }
  .subject-cell { max-width: 260px; }
  .reasons-cell { max-width: 320px; color: var(--muted); }
  .score-cell { font-variant-numeric: tabular-nums; text-align: right; }
  .badge {
    display: inline-block;
    padding: 0.15rem 0.55rem;
    border-radius: 999px;
    font-size: 0.72rem;
    font-weight: 600;
    white-space: nowrap;
  }
  .badge-ticket { color: var(--ticket); background: var(--ticket-bg); }
  .badge-review { color: var(--review); background: var(--review-bg); }
  .badge-ignored { color: var(--ignored); background: var(--ignored-bg); }
  .badge-skipped { color: var(--skipped); background: var(--skipped-bg); }
  .empty-state { color: var(--muted); padding: 2rem; text-align: center; background: white; border: 1px solid var(--border); border-radius: 8px; }
</style>
</head>
<body>
  <div class="container">
    <h1>Email Classification Report</h1>
    <div class="meta">
      Run at ${escapeHtml(runMetadata.runAt)} &middot;
      Mailbox: ${escapeHtml(runMetadata.mailbox)} &middot;
      ${results.length} message(s) seen this run
    </div>

    <div class="summary">
      <div class="summary-card"><div class="n">${counts['Ticket'] || 0}</div><div class="l">Ticket</div></div>
      <div class="summary-card"><div class="n">${counts['Needs Review'] || 0}</div><div class="l">Needs Review</div></div>
      <div class="summary-card"><div class="n">${counts['Ignored'] || 0}</div><div class="l">Ignored</div></div>
      <div class="summary-card"><div class="n">${counts['skipped'] || 0}</div><div class="l">Already Handled</div></div>
    </div>

    ${results.length ? `
    <table>
      <thead>
        <tr>
          <th>Received</th>
          <th>From</th>
          <th>Subject</th>
          <th>Score</th>
          <th>Decision</th>
          <th>Reasons</th>
        </tr>
      </thead>
      <tbody>${rows}</tbody>
    </table>
    ` : `<div class="empty-state">No new or changed messages since the last run.</div>`}
  </div>
</body>
</html>`;
}

// ---------------------------------------------------------------------------
// Main
// ---------------------------------------------------------------------------

async function main() {
  const runAt = new Date().toISOString();
  console.log(`[${runAt}] Starting classification run for mailbox ${MAILBOX_USER_ID}`);

  const accessToken = await getGraphAccessToken();
  const { messages, deltaLink } = await fetchDeltaMessages(accessToken);
  console.log(`Fetched ${messages.length} new/changed message(s) since last run.`);

  const results = [];

  for (const message of messages) {
    const existingCategories = message.categories || [];
    const alreadyHandled = existingCategories.some(c => MARKER_CATEGORIES.includes(c));

    const baseRow = {
      receivedDateTime: message.receivedDateTime,
      from: message.from?.emailAddress?.address || '(unknown sender)',
      subject: message.subject || '(no subject)'
    };

    if (alreadyHandled) {
      results.push({ ...baseRow, score: null, decision: 'skipped', reasons: [`already categorized: ${existingCategories.join(', ')}`] });
      continue;
    }

    const { score, reasons } = scoreEmail(message);
    const decision = classify(score);

    try {
      await applyCategory(message.id, accessToken, decision, existingCategories);
      results.push({ ...baseRow, score, decision, reasons });
      console.log(`  ${baseRow.subject} — score ${score} → ${decision}`);
    } catch (err) {
      console.error(`  Failed to apply category to message ${message.id}: ${err.response?.status} ${JSON.stringify(err.response?.data || err.message)}`);
      results.push({ ...baseRow, score, decision: 'ERROR applying category', reasons });
    }
  }

  // Persist delta link for next run — do this even if there were errors above,
  // since Graph delta semantics mean re-fetching the same page is safe but
  // we don't want to lose forward progress on a partial failure.
  if (deltaLink) {
    saveDeltaLink(deltaLink);
  }

  // Write reports
  fs.mkdirSync(REPORT_HISTORY_DIR, { recursive: true });
  const html = buildHtmlReport(results, { runAt, mailbox: MAILBOX_USER_ID });
  fs.writeFileSync(LATEST_REPORT_FILE, html);

  const historyFileName = `${runAt.replace(/[:.]/g, '-')}.html`;
  fs.writeFileSync(path.join(REPORT_HISTORY_DIR, historyFileName), html);

  console.log(`Report written to ${LATEST_REPORT_FILE}`);
  console.log(`Run complete. ${results.length} message(s) processed.`);
}

main().catch(err => {
  console.error('Classifier run failed:', err.message);
  process.exit(1);
});
