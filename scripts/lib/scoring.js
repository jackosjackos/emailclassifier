/**
 * Shared scoring engine and HTML report builder.
 *
 * Used by both:
 *   - classify-emails.js (live mailbox via Microsoft Graph)
 *   - import-pst.js       (static .pst file, no mailbox connection at all)
 *
 * Both callers normalize their source data into the same plain-object shape
 * before calling scoreEmail(), so the actual scoring logic only needs to
 * exist once:
 *
 *   {
 *     subject: string,
 *     receivedDateTime: string (ISO) | null,
 *     from: { emailAddress: { address: string } },
 *     categories: string[],                 // [] for PST (not applicable)
 *     internetMessageHeaders: [{ name, value }]
 *   }
 */

const INTERNAL_DOMAIN = (process.env.INTERNAL_DOMAIN || '').toLowerCase();
const KNOWN_CLIENT_DOMAINS = (process.env.KNOWN_CLIENT_DOMAINS || '')
  .split(',')
  .map(d => d.trim().toLowerCase())
  .filter(Boolean);

const HIGH_CONFIDENCE_THRESHOLD = Number(process.env.HIGH_CONFIDENCE_THRESHOLD || 25);
const LOW_CONFIDENCE_THRESHOLD = Number(process.env.LOW_CONFIDENCE_THRESHOLD || -25);

const MARKER_CATEGORIES = ['Ticket', 'Needs Review', 'Ignored'];

// ---------------------------------------------------------------------------
// Heuristic scoring engine — PLACEHOLDER weights, tune against real mail
// before trusting this. See README "Tuning the classifier".
// ---------------------------------------------------------------------------

function scoreEmail(message) {
  let score = 0;
  const reasons = [];

  const senderAddress = (message.from?.emailAddress?.address || '').toLowerCase();
  const senderDomain = senderAddress.split('@')[1] || '';
  const subject = (message.subject || '').toLowerCase();
  const headers = message.internetMessageHeaders || [];

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
      <td class="score-cell">${r.score !== null && r.score !== undefined ? r.score : '—'}</td>
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
  .meta { color: var(--muted); font-size: 0.85rem; margin-bottom: 1rem; }
  .mode-banner {
    display: inline-block;
    background: #fff4e0;
    color: #8a5a00;
    border: 1px solid #f0d9a8;
    border-radius: 6px;
    padding: 0.4rem 0.75rem;
    font-size: 0.8rem;
    font-weight: 600;
    margin-bottom: 1.5rem;
  }
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
      Source: ${escapeHtml(runMetadata.mailbox)} &middot;
      ${results.length} message(s) seen this run
    </div>
    ${runMetadata.testMode ? `<div class="mode-banner">${escapeHtml(runMetadata.modeBannerText || 'TEST MODE')}</div>` : ''}

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
    ` : `<div class="empty-state">No messages found for this run.</div>`}
  </div>
</body>
</html>`;
}

module.exports = {
  scoreEmail,
  classify,
  buildHtmlReport,
  escapeHtml,
  MARKER_CATEGORIES,
  HIGH_CONFIDENCE_THRESHOLD,
  LOW_CONFIDENCE_THRESHOLD
};
