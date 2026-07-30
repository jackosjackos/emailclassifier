/**
 * Email Classifier — PST import mode.
 *
 * Parses a local .pst file directly (via the pure-JS `pst-extractor`
 * library — no native compilation, no Outlook installation, no Microsoft
 * Graph credentials required) and scores the messages with the same
 * heuristic engine used against a live mailbox in classify-emails.js.
 *
 * This is a READ-ONLY, offline operation:
 *   - Nothing is written back to the .pst file.
 *   - No network calls are made at all.
 *   - No Outlook categories are applied anywhere (there's no live mailbox
 *     to apply them to). This is always effectively a dry run.
 *
 * Usage:
 *   node scripts/import-pst.js path/to/export.pst
 *   PST_FILE_PATH=path/to/export.pst npm run test-pst
 *
 * Env vars (same tuning knobs as the live classifier, read from .env):
 *   PST_SAMPLE_SIZE   — how many of the most recent messages to score
 *                        (default 100)
 *   PST_FOLDER_NAME   — which folder to read from, by display name
 *                        (default "Inbox")
 *   INTERNAL_DOMAIN, KNOWN_CLIENT_DOMAINS, HIGH_CONFIDENCE_THRESHOLD,
 *   LOW_CONFIDENCE_THRESHOLD — same as classify-emails.js
 */

require('dotenv').config();

const fs = require('fs');
const path = require('path');
const { PSTFile } = require('pst-extractor');
const { scoreEmail, classify, buildHtmlReport } = require('./lib/scoring');

const PST_SAMPLE_SIZE = Number(process.env.PST_SAMPLE_SIZE || 100);
const PST_FOLDER_NAME = process.env.PST_FOLDER_NAME || 'Inbox';
const OPEN_REPORT = process.env.OPEN_REPORT !== 'false';

const PROJECT_ROOT = path.join(__dirname, '..');
const REPORT_DIR = path.join(PROJECT_ROOT, 'report');
const REPORT_HISTORY_DIR = path.join(REPORT_DIR, 'history');
const LATEST_PST_REPORT_FILE = path.join(REPORT_DIR, 'latest-pst.html');

// ---------------------------------------------------------------------------
// Header parsing — PST stores raw RFC 822 headers as one text blob, same
// shape as an .eml. We need to turn that into the [{name, value}] array
// shape the shared scoring engine expects (matching what Graph returns).
// ---------------------------------------------------------------------------

function parseRawHeaders(rawHeaderBlock) {
  if (!rawHeaderBlock) return [];

  // Unfold header continuation lines (a line starting with whitespace is a
  // continuation of the previous header, per RFC 822/2822).
  const unfolded = rawHeaderBlock.replace(/\r\n[ \t]+/g, ' ').replace(/\n[ \t]+/g, ' ');
  const lines = unfolded.split(/\r?\n/);

  const headers = [];
  for (const line of lines) {
    const match = line.match(/^([^:\s]+):\s*(.*)$/);
    if (match) {
      headers.push({ name: match[1], value: match[2] });
    }
  }
  return headers;
}

// ---------------------------------------------------------------------------
// Walk the PST folder tree, find the target folder, collect messages
// ---------------------------------------------------------------------------

function findFolderByName(folder, targetName, depth = 0) {
  if (depth > 0 && folder.displayName === targetName) {
    return folder;
  }
  if (folder.hasSubfolders) {
    for (const child of folder.getSubFolders()) {
      const found = findFolderByName(child, targetName, depth + 1);
      if (found) return found;
    }
  }
  return null;
}

function collectMessagesFromFolder(folder) {
  const messages = [];
  if (folder.contentCount > 0) {
    let email = folder.getNextChild();
    while (email !== null) {
      messages.push(email);
      email = folder.getNextChild();
    }
  }
  return messages;
}

function pstMessageToScoringShape(pstMessage) {
  const deliveryDate = pstMessage.messageDeliveryTime || pstMessage.clientSubmitTime || null;

  return {
    subject: pstMessage.subject || '(no subject)',
    receivedDateTime: deliveryDate ? new Date(deliveryDate).toISOString() : null,
    from: {
      emailAddress: {
        address: pstMessage.senderEmailAddress || pstMessage.senderName || '(unknown sender)'
      }
    },
    categories: [], // not meaningfully preserved/applicable from a static PST export
    internetMessageHeaders: parseRawHeaders(pstMessage.transportMessageHeaders)
  };
}

// ---------------------------------------------------------------------------
// Main
// ---------------------------------------------------------------------------

async function runPstImport(pstPath) {
  if (!pstPath) {
    console.error('Usage: node scripts/import-pst.js <path-to-file.pst>');
    console.error('   or: set PST_FILE_PATH in .env and run `npm run test-pst`');
    process.exit(1);
  }

  const resolvedPath = path.resolve(pstPath);
  if (!fs.existsSync(resolvedPath)) {
    console.error(`File not found: ${resolvedPath}`);
    process.exit(1);
  }

  const runAt = new Date().toISOString();
  console.log(`[${runAt}] Reading PST file: ${resolvedPath}`);
  console.log(`Looking for folder "${PST_FOLDER_NAME}", most recent ${PST_SAMPLE_SIZE} message(s).`);

  const pstFile = new PSTFile(fs.readFileSync(resolvedPath));
  const rootFolder = pstFile.getRootFolder();

  const targetFolder = findFolderByName(rootFolder, PST_FOLDER_NAME);
  if (!targetFolder) {
    console.error(`Could not find a folder named "${PST_FOLDER_NAME}" in this PST.`);
    console.error('If your export used a different folder name/language, set PST_FOLDER_NAME in .env.');
    process.exit(1);
  }

  const rawMessages = collectMessagesFromFolder(targetFolder);
  console.log(`Found ${rawMessages.length} message(s) in "${PST_FOLDER_NAME}".`);

  // Sort newest first, take the requested sample size
  const sorted = rawMessages
    .map(pstMessageToScoringShape)
    .sort((a, b) => {
      const dateA = a.receivedDateTime ? new Date(a.receivedDateTime).getTime() : 0;
      const dateB = b.receivedDateTime ? new Date(b.receivedDateTime).getTime() : 0;
      return dateB - dateA;
    })
    .slice(0, PST_SAMPLE_SIZE);

  const results = sorted.map(message => {
    const { score, reasons } = scoreEmail(message);
    const decision = classify(score);
    console.log(`  ${message.subject} — score ${score} → ${decision}`);
    return {
      receivedDateTime: message.receivedDateTime,
      from: message.from.emailAddress.address,
      subject: message.subject,
      score,
      decision,
      reasons
    };
  });

  fs.mkdirSync(REPORT_HISTORY_DIR, { recursive: true });
  const html = buildHtmlReport(results, {
    runAt,
    mailbox: `${path.basename(resolvedPath)} (folder: ${PST_FOLDER_NAME})`,
    testMode: true,
    modeBannerText: `PST IMPORT — read-only, offline. No categories applied, no network calls made. Sample of ${results.length} most recent message(s).`
  });

  fs.writeFileSync(LATEST_PST_REPORT_FILE, html);
  const historyFileName = `pst-${runAt.replace(/[:.]/g, '-')}.html`;
  fs.writeFileSync(path.join(REPORT_HISTORY_DIR, historyFileName), html);

  console.log(`Report written to ${LATEST_PST_REPORT_FILE}`);
  console.log(`Run complete. ${results.length} message(s) processed.`);

  if (OPEN_REPORT) {
    try {
      const open = (await import('open')).default;
      await open(LATEST_PST_REPORT_FILE);
    } catch (err) {
      console.warn(`Could not auto-open the report (${err.message}). Open it manually: ${LATEST_PST_REPORT_FILE}`);
    }
  }

  return { results, reportPath: LATEST_PST_REPORT_FILE };
}

module.exports = { runPstImport };

if (require.main === module) {
  const pstPathArg = process.argv[2] || process.env.PST_FILE_PATH;
  runPstImport(pstPathArg).catch(err => {
    console.error('PST import failed:', err.message);
    process.exit(1);
  });
}
