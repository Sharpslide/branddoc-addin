/* Replays real document.xml through blocks.js in a browser engine and deep-compares the result
 * with what the engine's extract.py produced for the same files.
 *
 * It runs in Chromium rather than Node because blocks.js uses DOMParser, and so will Word's
 * WebView. The drafts are the client's, so they are not in this repo:
 *
 *   node engine/blocks.test.js <dir of <section>.xml> <content.json>
 */
const fs = require('fs');
const path = require('path');
const { chromium } = require('playwright');

const [xmlDir, contentPath] = process.argv.slice(2);
if (!xmlDir || !contentPath) {
  console.error('usage: node engine/blocks.test.js <dir of <section>.xml> <content.json>');
  process.exit(2);
}
const expected = JSON.parse(fs.readFileSync(contentPath, 'utf8'));
const SECTIONS = Object.keys(expected).sort();
const source = fs.readFileSync(path.join(__dirname, 'blocks.js'), 'utf8');

const diffs = [];
function cmp(a, b, p) {
  if (diffs.length > 12) return;
  if (a === b) return;
  const ta = a === null ? 'null' : Array.isArray(a) ? 'array' : typeof a;
  const tb = b === null ? 'null' : Array.isArray(b) ? 'array' : typeof b;
  if (ta !== tb) { diffs.push(`${p}: expected ${tb} ${JSON.stringify(b)}, got ${ta} ${JSON.stringify(a)}`); return; }
  if (ta === 'array') {
    if (a.length !== b.length) diffs.push(`${p}: ${a.length} items, expected ${b.length}`);
    for (let i = 0; i < Math.max(a.length, b.length); i++) cmp(a[i], b[i], `${p}[${i}]`);
    return;
  }
  if (ta === 'object') {
    for (const k of new Set([...Object.keys(a), ...Object.keys(b)])) {
      if (!(k in a)) { diffs.push(`${p}.${k}: missing, expected ${JSON.stringify(b[k])}`); continue; }
      if (!(k in b)) { diffs.push(`${p}.${k}: unexpected ${JSON.stringify(a[k])}`); continue; }
      cmp(a[k], b[k], `${p}.${k}`);
    }
    return;
  }
  diffs.push(`${p}: got ${JSON.stringify(a)}, expected ${JSON.stringify(b)}`);
}

(async () => {
  const browser = await chromium.launch();
  const page = await browser.newPage();
  await page.setContent('<!doctype html><meta charset="utf-8"><title>blocks</title>');
  await page.addScriptTag({ content: source });

  let totalRuns = 0;
  for (const sec of SECTIONS) {
    const xml = fs.readFileSync(path.join(xmlDir, sec + '.xml'), 'utf8');
    const got = await page.evaluate(x => window.BrandDocBlocks.fromDocumentXml(x), xml);
    const want = expected[sec].blocks;
    const runs = want.reduce((a, b) => a + (b.runs ? b.runs.length : 0), 0);
    totalRuns += runs;
    const mark = got.length === want.length ? ' ' : '!';
    console.log(`${mark} section ${sec}: ${String(got.length).padStart(3)} blocks (expected ${want.length}), ${runs} runs`);
    cmp(got, want, `section ${sec}`);
  }
  await browser.close();

  console.log(`\n${totalRuns} runs compared across ${SECTIONS.length} documents`);
  if (diffs.length) {
    console.log('\nFAIL — ' + diffs.length + (diffs.length > 12 ? '+' : '') + ' divergence(s):');
    diffs.slice(0, 12).forEach(d => console.log('  ' + d));
    process.exit(1);
  }
  console.log('\nPASS — the pane reads a draft exactly as the engine does, run for run.');
})();
