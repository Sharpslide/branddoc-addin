/* Replays the engine's own fixture through classify.js and fails on any divergence.
 *
 * The fixtures are the client's draft, so they are not in this repo. Point at them:
 *   node engine/classify.test.js <content.json> <classified.json> <profile.json>
 */
const fs = require('fs');
const C = require('./classify.js');

const [contentPath, expectedPath, profilePath] = process.argv.slice(2);
if (!contentPath || !expectedPath) {
  console.error('usage: node engine/classify.test.js <content.json> <classified.json> [profile.json]');
  process.exit(2);
}
const content = JSON.parse(fs.readFileSync(contentPath, 'utf8'));
const expected = JSON.parse(fs.readFileSync(expectedPath, 'utf8')).sections;
const profile = profilePath ? JSON.parse(fs.readFileSync(profilePath, 'utf8')) : C.NEUTRAL;

const SECTIONS = ['0', '1', '2', '3', '4'];
const got = {};
for (const s of SECTIONS) got[s] = C.classify(s, content[s].blocks, profile);
for (const s of SECTIONS) C.structure(got[s], profile);
let leads = 0;
for (const s of SECTIONS) leads += C.addLeads(got[s], profile);

// ---- deep compare, reporting the first few differences in full
const diffs = [];
function cmp(a, b, path) {
  if (diffs.length > 12) return;
  if (a === b) return;
  const ta = a === null ? 'null' : Array.isArray(a) ? 'array' : typeof a;
  const tb = b === null ? 'null' : Array.isArray(b) ? 'array' : typeof b;
  if (ta !== tb) { diffs.push(`${path}: type ${tb} expected, got ${ta}`); return; }
  if (ta === 'array') {
    if (a.length !== b.length) diffs.push(`${path}: length ${a.length}, expected ${b.length}`);
    for (let i = 0; i < Math.max(a.length, b.length); i++) cmp(a[i], b[i], `${path}[${i}]`);
    return;
  }
  if (ta === 'object') {
    const keys = new Set([...Object.keys(a), ...Object.keys(b)]);
    for (const k of keys) {
      if (!(k in a)) { diffs.push(`${path}.${k}: missing, expected ${JSON.stringify(b[k])}`); continue; }
      if (!(k in b)) { diffs.push(`${path}.${k}: unexpected ${JSON.stringify(a[k])}`); continue; }
      cmp(a[k], b[k], `${path}.${k}`);
    }
    return;
  }
  diffs.push(`${path}: got ${JSON.stringify(a)}, expected ${JSON.stringify(b)}`);
}
for (const s of SECTIONS) cmp(got[s], expected[s], `section ${s}`);

const roleCount = o => {
  const c = {};
  for (const s of SECTIONS) for (const b of o[s]) c[b.role] = (c[b.role] || 0) + 1;
  return c;
};
console.log('blocks in :', SECTIONS.map(s => content[s].blocks.length).join(' + '));
console.log('roles out :', JSON.stringify(roleCount(got)));
console.log('expected  :', JSON.stringify(roleCount(expected)));
console.log('lead paragraphs chosen:', leads);
if (diffs.length) {
  console.log('\nFAIL — ' + diffs.length + (diffs.length > 12 ? '+' : '') + ' divergence(s):');
  diffs.slice(0, 12).forEach(d => console.log('  ' + d));
  process.exit(1);
}
console.log('\nPASS — JS classifier matches the engine on every block of all five sections.');
