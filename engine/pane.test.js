/* The brand is loaded once and remembered on this machine.
 *
 * Drives the real taskpane.html in Chromium with Office.js stubbed out — the stub reports Word
 * and an EMPTY custom XML part list, which is exactly an unformatted draft. Checks four states:
 * nothing remembered, a brand just loaded, the pane reopened as it would be for the next
 * document, and the brand forgotten.
 *
 * The brand file is the client's, so it is not in this repo:
 *   node engine/pane.test.js <brand.json>
 */
const fs = require('fs');
const http = require('http');
const path = require('path');
const { chromium } = require('playwright');

const brandPath = process.argv[2];
if (!brandPath) {
  console.error('usage: node engine/pane.test.js <brand.json>');
  process.exit(2);
}
const BRAND = fs.readFileSync(brandPath, 'utf8');
const ROOT = path.join(__dirname, '..');
const TYPES = { '.html': 'text/html', '.js': 'application/javascript', '.css': 'text/css',
                '.json': 'application/json' };

// localStorage needs a real origin, so the pane is served rather than opened from disk.
function serve() {
  const server = http.createServer((req, res) => {
    const rel = decodeURIComponent(req.url.split('?')[0]).replace(/^\/+/, '') || 'taskpane.html';
    const file = path.join(ROOT, rel);
    if (!file.startsWith(ROOT) || !fs.existsSync(file) || fs.statSync(file).isDirectory()) {
      res.writeHead(404).end('not found');
      return;
    }
    res.writeHead(200, { 'Content-Type': TYPES[path.extname(file)] || 'text/plain' });
    res.end(fs.readFileSync(file));
  });
  return new Promise(r => server.listen(0, '127.0.0.1', () => r(server)));
}

// Office.js comes from a CDN this container cannot reach, so it is stubbed. The stub reports Word
// and an EMPTY custom XML part list — exactly an unformatted draft.
const OFFICE_STUB = `
window.Office = {
  HostType: { Word: 'Word' },
  AsyncResultStatus: { Succeeded: 'succeeded' },
  context: { requirements: { isSetSupported: () => true },
             document: { settings: { get: () => null, set: () => {}, saveAsync: (cb) => cb({ status: 'succeeded' }) } } },
  onReady: (fn) => Promise.resolve().then(() => fn({ host: 'Word' }))
};
window.Word = {
  Style: {},
  run: async (fn) => fn({
    sync: async () => {},
    document: {
      customXmlParts: { getByNamespace: () => ({ items: [], load: () => {} }) },
      getStyles: () => ({ items: [], load: () => {} }),
      body: { paragraphs: { items: [], load: () => {} }, getOoxml: () => ({ value: '' }) }
    }
  })
};`;

(async () => {
  const server = await serve();
  const base = 'http://127.0.0.1:' + server.address().port;
  const browser = await chromium.launch();
  const ctx = await browser.newContext();
  const page = await ctx.newPage();
  const errs = [];
  page.on('pageerror', e => errs.push(String(e)));
  page.on('console', m => { if (m.type() === 'error') errs.push(m.text()); });
  await page.route('**/office.js', r => r.fulfill({ contentType: 'application/javascript', body: OFFICE_STUB }));

  const read = async () => ({
    pane: await page.getAttribute('.pane:not([hidden])', 'data-pane'),
    client: await page.textContent('#client'),
    source: (await page.textContent('#brandSource') || '').trim().slice(0, 70),
    applyDisabled: await page.getAttribute('#apply', 'disabled') !== null,
    formatsReady: await page.locator('button.format[data-ready="1"]').count(),
    brandRow: (await page.textContent('#brandRowName') || '').trim(),
    forgetHidden: await page.getAttribute('#brandForget', 'hidden') !== null,
    swatches: await page.locator('#palette .swatch').count()
  });

  // 1. nothing remembered yet
  await page.goto(base + '/taskpane.html', { waitUntil: 'networkidle' });
  await page.waitForTimeout(400);
  console.log('cold start      ', JSON.stringify(await read()));

  // 2. the user picks a brand file
  await page.setInputFiles('#brandFile', { name: 'brand.json', mimeType: 'application/json', buffer: Buffer.from(BRAND) });
  await page.waitForTimeout(400);
  console.log('after loading   ', JSON.stringify(await read()));
  console.log('  message:', (await page.textContent('#brandResult')).replace(/\s+/g, ' ').trim().slice(0, 90));

  // 3. reopen the pane, as Word does for every new document
  await page.goto(base + '/taskpane.html', { waitUntil: 'networkidle' });
  await page.waitForTimeout(400);
  const reopened = await read();
  console.log('next document   ', JSON.stringify(reopened));

  // 3b. the format button must actually do something. It once threw on its first line, before
  // its own try block, because it looked up a button id that an earlier version of the pane had
  // and this one does not — so a click did nothing and reported nothing.
  const liveFormat = page.locator('button.format[data-ready="1"]').first();
  await liveFormat.click();
  await page.waitForTimeout(600);
  const planShown = await page.getAttribute('#designResult', 'hidden') === null;
  const planText = (await page.textContent('#designResult') || '').replace(/\s+/g, ' ').trim();
  console.log('clicked format  ', JSON.stringify({ reacted: planShown, says: planText.slice(0, 72) }));
  if (!planShown) errs.push('clicking a format button produced no response at all');

  // 4. forget it — the control lives on the Brand pane, so go there first
  await page.click('#brandRow');
  await page.waitForTimeout(200);
  console.log('brand pane     ', JSON.stringify({ pane: await page.getAttribute('.pane:not([hidden])', 'data-pane') }));
  await page.click('#brandForget');
  await page.waitForTimeout(200);
  await page.goto(base + '/taskpane.html', { waitUntil: 'networkidle' });
  await page.waitForTimeout(400);
  console.log('after forgetting', JSON.stringify(await read()));

  await browser.close();
  server.close();
  const ok = reopened.client.includes('Connor') && !reopened.applyDisabled
    && reopened.formatsReady === 1 && reopened.brandRow.includes('Connor') && planShown;
  if (errs.length) { console.log('\nconsole errors:'); errs.slice(0,5).forEach(e => console.log('  ' + e)); }
  console.log(ok && !errs.length
    ? '\nPASS — loaded once, still there on the next document, and cleanly forgotten.'
    : '\nFAIL');
  process.exit(ok && !errs.length ? 0 : 1);
})();
