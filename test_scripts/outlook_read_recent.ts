import { chromium, BrowserContext, Page } from 'playwright';
import * as path from 'path';
import * as fs from 'fs';

type OutlookFlavor = 'work' | 'personal';

interface EmailItem {
  index: number;
  sender: string;
  subject: string;
  preview: string;
  received: string;
}

const PROFILE_DIR = path.resolve(__dirname, '..', '.playwright-profile');
const LOGIN_TIMEOUT_MS = 5 * 60 * 1000;
const LIST_TIMEOUT_MS = 60 * 1000;

function parseFlavor(arg: string | undefined): OutlookFlavor {
  if (arg === 'personal') return 'personal';
  return 'work';
}

function urlFor(flavor: OutlookFlavor): string {
  return flavor === 'personal'
    ? 'https://outlook.live.com/mail/0/inbox'
    : 'https://outlook.office.com/mail/inbox';
}

async function waitForInbox(page: Page): Promise<void> {
  const start = Date.now();
  while (Date.now() - start < LOGIN_TIMEOUT_MS) {
    const url = page.url();
    if (/outlook\.(office|live)\.com\/mail/.test(url)) {
      try {
        await page.waitForSelector(
          'div[role="option"], [aria-label="Message list"] [role="option"]',
          {
            timeout: 10_000,
          },
        );
        return;
      } catch {
        // keep waiting — maybe still loading
      }
    }
    await page.waitForTimeout(2000);
  }
  throw new Error('Timed out waiting for Outlook inbox to load.');
}

async function extractTopEmails(page: Page, count: number): Promise<EmailItem[]> {
  await page.waitForSelector('div[role="option"]', { timeout: LIST_TIMEOUT_MS });

  const items = await page.evaluate((n) => {
    const results: Array<{ sender: string; subject: string; preview: string; received: string }> =
      [];
    const nodes = Array.from(document.querySelectorAll('div[role="option"]')) as HTMLElement[];

    for (const node of nodes) {
      if (results.length >= n) break;

      const text = (sel: string): string => {
        const el = node.querySelector(sel) as HTMLElement | null;
        return el?.innerText?.trim() ?? '';
      };

      const aria = node.getAttribute('aria-label') ?? '';

      // Try structured selectors first; fall back to aria-label parsing.
      let sender = text('[class*="SenderName"], [class*="sender"], span[title]');
      let subject = text('[class*="Subject"], span[role="heading"], [class*="subject"]');
      let preview = text('[class*="Preview"], [class*="preview"], [class*="Snippet"]');
      let received = text('[class*="Date"], [class*="timestamp"], time');

      if (!sender && aria) {
        const m = aria.match(/^([^,]+),/);
        if (m) sender = m[1].trim();
      }
      if (!subject && aria) {
        const parts = aria.split(',').map((s) => s.trim());
        if (parts.length > 1) subject = parts[1];
      }
      if (!preview && aria) preview = aria;

      if (sender || subject) {
        results.push({ sender, subject, preview, received });
      }
    }
    return results;
  }, count);

  return items.map((it, i) => ({ index: i + 1, ...it }));
}

async function main() {
  const flavor = parseFlavor(process.argv[2]);
  const count = Number(process.argv[3] ?? 5);
  const outPath = process.argv[4] ?? path.resolve(__dirname, '..', 'outlook_report.json');

  fs.mkdirSync(PROFILE_DIR, { recursive: true });

  console.log(`[outlook-reader] Launching Chrome for ${flavor} Outlook (${urlFor(flavor)})`);
  console.log(`[outlook-reader] Profile: ${PROFILE_DIR}`);

  const context: BrowserContext = await chromium.launchPersistentContext(PROFILE_DIR, {
    channel: 'chrome',
    headless: false,
    viewport: { width: 1280, height: 900 },
  });

  const page = context.pages()[0] ?? (await context.newPage());
  await page.goto(urlFor(flavor), { waitUntil: 'domcontentloaded' });

  console.log('[outlook-reader] If a login screen appears, please log in in the browser window.');
  console.log('[outlook-reader] Waiting for inbox to load (up to 5 minutes)...');

  await waitForInbox(page);

  console.log(`[outlook-reader] Inbox detected. Extracting top ${count} emails...`);
  const emails = await extractTopEmails(page, count);

  fs.writeFileSync(outPath, JSON.stringify(emails, null, 2));
  console.log(`[outlook-reader] Wrote report to ${outPath}`);
  console.log('---BEGIN_REPORT---');
  console.log(JSON.stringify(emails, null, 2));
  console.log('---END_REPORT---');

  await context.close();
}

main().catch((err) => {
  console.error('[outlook-reader] ERROR:', err);
  process.exit(1);
});
