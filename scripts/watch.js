/**
 * Optional continuous local runner.
 *
 * `npm run classify` runs once and exits — good for testing, and good
 * enough in production if you trigger it yourself via OS-level cron /
 * Windows Task Scheduler (see README).
 *
 * This script is an alternative for people who'd rather leave one Node
 * process running in a terminal/tmux/pm2 than configure an OS scheduler.
 * It re-runs the same classification logic on a cron schedule using
 * node-cron, in-process, indefinitely.
 *
 * Usage: npm run watch
 * Stop with Ctrl+C.
 */

require('dotenv').config();
const cron = require('node-cron');
const { runOnce } = require('./classify-emails');

// Every 15 minutes by default — override with WATCH_CRON in .env using
// standard cron syntax, e.g. "*/5 * * * *" for every 5 minutes.
const CRON_EXPRESSION = process.env.WATCH_CRON || '*/15 * * * *';

if (!cron.validate(CRON_EXPRESSION)) {
  console.error(`Invalid cron expression in WATCH_CRON: "${CRON_EXPRESSION}"`);
  process.exit(1);
}

console.log(`Email classifier watch mode started.`);
console.log(`Schedule: "${CRON_EXPRESSION}" (cron syntax, local time)`);
console.log(`Press Ctrl+C to stop.\n`);

let running = false;

async function tick() {
  if (running) {
    console.log('Previous run still in progress, skipping this tick.');
    return;
  }
  running = true;
  try {
    await runOnce();
  } catch (err) {
    console.error('Scheduled run failed:', err.message);
    // Deliberately don't exit the process — a single failed run (e.g.
    // transient Graph API error) shouldn't kill the whole watcher. It'll
    // just try again next tick.
  } finally {
    running = false;
  }
}

// Run once immediately on startup, then follow the schedule.
tick();
cron.schedule(CRON_EXPRESSION, tick);
