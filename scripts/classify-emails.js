/**
 * Email Classifier — live mailbox mode (Microsoft Graph).
 *
 * Runs entirely locally via Node — no cloud scheduler, no CI. State and
 * reports persist to local files in this project folder between runs.
 * Scoring logic lives in scripts/lib/scoring.js, shared with import-pst.js.
 *
 * What this does:
 *   1. Authenticates to Microsoft Graph via client-credentials OAuth2.
 *   2. Fetches messages — either via delta query (normal mode, only
 *      new/changed since last run) or a plain top-N fetch (TEST_MODE, for
 *      running against a fixed sample like "the last 100 emails").
 *   3. Skips anything that already carries one of our marker categories
 *      (idempotency — safe to re-run without reclassifying).
 *   4. Scores each remaining message with the shared heuristic engine.
 *   5. Applies an Outlook category based on the score — UNLESS DRY_RUN is
 *      set, in which case nothing in the mailbox is touched at all.
 *   6. Writes an HTML report to report/latest.html (+ a timestamped copy in
 *      report/history/) and opens it in your default browser.
 *
 * For testing against a static export instead of a live mailbox connection,
 * see scripts/import-pst.js — no Graph credentials needed for that path.
 *
 * Deliberately NOT included yet:
 *   - Any call to the Accelo API.
 *   - Contact/company matching.
 *   - Reply-threading detection.
 */

require('dotenv').config();

const axios = require('axios');
const fs = require('fs');
const path = require('path');
const { scoreEmail, classify, buildHtmlReport, MARKER_CATEGORIES } = require('./lib/scoring');

// ---------------------------------------------------------------------------
// Configuration
// ---------------------------------------------------------------------------

const TENANT_ID = requireEnv('GRAPH_TENANT_ID');
const CLIENT_ID = requireEnv('GRAPH_CLIENT_ID');
const CLIENT_SECRET = requireEnv('GRAPH_CLIENT_SECRET');
const MAILBOX_USER_ID = requireEnv('MAILBOX_USER_ID'); // shared mailbox UPN or object ID

// TEST_MODE: fetch a fixed sample of the most recent N emails instead of
// using delta query. Use this for a first-pass validation run.
const TEST_MODE = /^true$/i.test(process.env.TEST_MODE || '');
const TEST_SAMPLE_SIZE = Number(process.env.TEST_SAMPLE_SIZE || 100);

// DRY_RUN: compute scores/decisions and write the report, but never call
// PATCH against the mailbox. Defaults to true whenever TEST_MODE is on.
const DRY_RUN = process.env.DRY_RUN !== undefined
  ? /^true$/i.test(process.env.DRY_RUN)
  : TEST_MODE;

const OPEN_REPORT = process.env.OPEN_REPORT !== 'false'; // opt out with OPEN_REPORT=false

const PROJECT_ROOT = path.join(__dirname, '..');
const STATE_DIR = path.join(PROJECT_ROOT, 'state');
const DELTA_LINK_FILE = path.join(STATE_DIR, 'delta-link.json');
const REPORT_DIR = path.join(PROJECT_ROOT, 'report');
const REPORT_HISTORY_DIR = path.join(REPORT_DIR, 'history');
const LATEST_REPORT_FILE = path.join(REPORT_DIR, 'latest.html');

function requireEnv(name) {
  const value = process.env[name];
  if (!value) {
    console.error(`Missing required environment variable: ${name}`);
    console.error(`Copy .env.example to .env and fill in your values.`);
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
// Delta query state (persisted to a local file between runs)
// ---------------------------------------------------------------------------

function loadDeltaLink() {
  try {
    const raw = fs.readFileSync(DELTA_LINK_FILE, 'utf8');
    return JSON.parse(raw).deltaLink || null;
  } catch {
    return null;
  }
}

function saveDeltaLink(deltaLink) {
  fs.mkdirSync(STATE_DIR, { recursive: true });
  fs.writeFileSync(DELTA_LINK_FILE, JSON.stringify({ deltaLink, updatedAt: new Date().toISOString() }, null, 2));
}

const SELECT_FIELDS = [
  'id', 'subject', 'from', 'receivedDateTime', 'categories',
  'conversationId', 'internetMessageId', 'bodyPreview', 'internetMessageHeaders'
].join(',');

// ---------------------------------------------------------------------------
// TEST_MODE fetch — last N messages, no delta state involved at all
// ---------------------------------------------------------------------------

async function fetchTestSample(accessToken) {
  const url = `https://graph.microsoft.com/v1.0/users/${MAILBOX_USER_ID}/mailFolders/inbox/messages` +
    `?$select=${SELECT_FIELDS}&$orderby=receivedDateTime desc&$top=${TEST_SAMPLE_SIZE}`;

  try {
    const { data } = await axios.get(url, { headers: { Authorization: `Bearer ${accessToken}` } });
    return { messages: data.value || [], deltaLink: null };
  } catch (err) {
    throw new Error(`Graph test-sample fetch failed: ${err.response?.status} ${JSON.stringify(err.response?.data || err.message)}`);
  }
}

// ---------------------------------------------------------------------------
// Normal mode fetch — delta query, following pagination
// ---------------------------------------------------------------------------

async function fetchDeltaMessages(accessToken) {
  const storedDeltaLink = loadDeltaLink();

  let url = storedDeltaLink ||
    `https://graph.microsoft.com/v1.0/users/${MAILBOX_USER_ID}/mailFolders/inbox/messages/delta?$select=${SELECT_FIELDS}`;

  const allMessages = [];
  let finalDeltaLink = storedDeltaLink;

  while (url) {
    let response;
    try {
      response = await axios.get(url, { headers: { Authorization: `Bearer ${accessToken}` } });
    } catch (err) {
      if (err.response?.status === 410) {
        console.warn('Delta link expired (410 Gone) — resetting to a fresh full sync.');
        fs.rmSync(DELTA_LINK_FILE, { force: true });
        return fetchDeltaMessages(accessToken);
      }
      throw new Error(`Graph delta query failed: ${err.response?.status} ${JSON.stringify(err.response?.data || err.message)}`);
    }

    allMessages.push(...(response.data.value || []));

    if (response.data['@odata.nextLink']) {
      url = response.data['@odata.nextLink'];
    } else {
      url = null;
      finalDeltaLink = response.data['@odata.deltaLink'];
    }
  }

  return { messages: allMessages, deltaLink: finalDeltaLink };
}

// ---------------------------------------------------------------------------
// Apply category back to the message in Outlook (skipped entirely in DRY_RUN)
// ---------------------------------------------------------------------------

async function applyCategory(messageId, accessToken, categoryName, existingCategories) {
  const preserved = (existingCategories || []).filter(c => !MARKER_CATEGORIES.includes(c));
  const newCategories = [...preserved, categoryName];

  await axios.patch(
    `https://graph.microsoft.com/v1.0/users/${MAILBOX_USER_ID}/messages/${messageId}`,
    { categories: newCategories },
    { headers: { Authorization: `Bearer ${accessToken}`, 'Content-Type': 'application/json' } }
  );
}

// ---------------------------------------------------------------------------
// Main — exported so watch.js can call it repeatedly without spawning a
// new process each time.
// ---------------------------------------------------------------------------

async function runOnce() {
  const runAt = new Date().toISOString();
  const modeLabel = TEST_MODE ? `TEST MODE (last ${TEST_SAMPLE_SIZE} emails${DRY_RUN ? ', dry run' : ''})` : 'normal (delta)';
  console.log(`[${runAt}] Starting classification run for mailbox ${MAILBOX_USER_ID} — mode: ${modeLabel}`);

  const accessToken = await getGraphAccessToken();
  const { messages, deltaLink } = TEST_MODE
    ? await fetchTestSample(accessToken)
    : await fetchDeltaMessages(accessToken);

  console.log(`Fetched ${messages.length} message(s).`);

  const results = [];

  for (const message of messages) {
    const existingCategories = message.categories || [];
    const alreadyHandled = !TEST_MODE && existingCategories.some(c => MARKER_CATEGORIES.includes(c));

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

    if (DRY_RUN) {
      results.push({ ...baseRow, score, decision, reasons });
      console.log(`  [DRY RUN] ${baseRow.subject} — score ${score} → ${decision} (not applied)`);
      continue;
    }

    try {
      await applyCategory(message.id, accessToken, decision, existingCategories);
      results.push({ ...baseRow, score, decision, reasons });
      console.log(`  ${baseRow.subject} — score ${score} → ${decision}`);
    } catch (err) {
      console.error(`  Failed to apply category to message ${message.id}: ${err.response?.status} ${JSON.stringify(err.response?.data || err.message)}`);
      results.push({ ...baseRow, score, decision: 'ERROR applying category', reasons });
    }
  }

  if (!TEST_MODE && deltaLink) {
    saveDeltaLink(deltaLink);
  }

  fs.mkdirSync(REPORT_HISTORY_DIR, { recursive: true });
  const html = buildHtmlReport(results, {
    runAt,
    mailbox: MAILBOX_USER_ID,
    testMode: TEST_MODE,
    modeBannerText: `TEST MODE — sample of last ${TEST_SAMPLE_SIZE} emails${DRY_RUN ? ', DRY RUN (no categories were applied to your mailbox)' : ''}`
  });
  fs.writeFileSync(LATEST_REPORT_FILE, html);

  const historyFileName = `${runAt.replace(/[:.]/g, '-')}.html`;
  fs.writeFileSync(path.join(REPORT_HISTORY_DIR, historyFileName), html);

  console.log(`Report written to ${LATEST_REPORT_FILE}`);
  console.log(`Run complete. ${results.length} message(s) processed.`);

  if (OPEN_REPORT) {
    try {
      const open = (await import('open')).default;
      await open(LATEST_REPORT_FILE);
    } catch (err) {
      console.warn(`Could not auto-open the report (${err.message}). Open it manually: ${LATEST_REPORT_FILE}`);
    }
  }

  return { results, reportPath: LATEST_REPORT_FILE };
}

module.exports = { runOnce };

if (require.main === module) {
  runOnce().catch(err => {
    console.error('Classifier run failed:', err.message);
    process.exit(1);
  });
}
