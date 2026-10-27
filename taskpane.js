/* BrandDoc task pane.
 *
 * Three jobs: show the brand, write it into the open document, and let a formatting decision be
 * questioned and corrected.
 *
 * Two things are worth knowing about how this is built.
 *
 * It uses the Style API (getStyles / addStyle, then style.font and style.paragraphFormat) rather
 * than injecting OOXML. Injecting looks neater and reuses the XML the Python engine already
 * produces, but Word documents it: "if a style is already in use in the destination document,
 * your content will take on the resident definition for the style, regardless of what you include
 * in styles.xml". A draft's Heading 1 is always already in use, so injection cannot restyle the
 * one thing this add-in exists to restyle.
 *
 * It carries no client data. The brand and the decision log ride inside the .docx as a custom
 * XML part, because serving them from the add-in's web host would mean publishing a client's
 * unreleased brand system and, since every decision quotes the text it was made about, excerpts
 * of the confidential document itself. The document is already the confidential thing;
 * anyone who can open it can already read it. It also means a second client's document simply
 * brings its own brand.
 *
 * It checks what this build of Word can actually do before claiming to have done it. Letter
 * spacing, all-caps and style borders live in the WordApiDesktop sets, which Word on the web does
 * not have. Anything unsupported is reported, never silently skipped -- a style sheet that
 * quietly applied two thirds of itself would be worse than one that refused.
 */
'use strict';

const NS = 'urn:sharpslide:branddoc:1';
let BRAND = null;
let SUPPORT = {};
let LAST = null;            // the paragraph most recently inspected

const $ = (id) => document.getElementById(id);

/* ------------------------------------------------------------------ startup */
Office.onReady(async (info) => {
  if (info.host !== Office.HostType.Word) {
    document.body.innerHTML = '<p style="padding:24px">BrandDoc runs in Word.</p>';
    return;
  }
  SUPPORT = {
    styles: req('WordApi', '1.5'),      // style.font, style.paragraphFormat
    setBase: req('WordApi', '1.6'),     // baseStyle / nextParagraphStyle are settable
    borders: req('WordApiDesktop', '1.1'),
    typography: req('WordApiDesktop', '1.3'),   // font.spacing, font.allCaps
  };
  reportSupport();
  try {
    BRAND = await loadBrand();
  } catch (e) {
    BRAND = null;
  }
  renderBrand();
  fillRoles();
  if ($('brandFile')) $('brandFile').onchange = onBrandFile;
  if ($('brandForget')) $('brandForget').onclick = onBrandForget;
  $('apply').onclick = applyBrand;
  if ($('design')) $('design').onclick = design;
  $('inspect').onclick = inspect;
  $('record').onclick = record;
  $('copyLog').onclick = copyLog;
  await drawLog();
});

/* The brand and the reasoning come out of the document's own custom XML part. */
/* Three places a brand can come from, in this order:
 *
 *   the document   written by the engine, and the only one that also carries the reasoning
 *   this machine   loaded once by hand and kept in the add-in's own storage, so every document
 *                  after the first opens with it already there
 *   the host       a copy served next to the add-in, if whoever deployed it chose to put one
 *                  there. This host is public, so none is committed.
 *
 * The middle one is why the pane does not ask on every document. localStorage belongs to the
 * add-in on this machine: it never reaches the document, the host, or anybody else.
 */
const MEMORY_KEY = 'branddoc.brand';

async function loadBrand() {
  try {
    const b = await brandFromDocument();
    b.source = 'document';
    return b;
  } catch (e) {
    const remembered = brandFromMemory();
    if (remembered) return remembered;
    const served = await brandFromHost();
    if (served) return served;
    throw e;
  }
}

function brandFromMemory() {
  let raw = null;
  try { raw = localStorage.getItem(MEMORY_KEY); } catch { return null; }
  if (!raw) return null;
  try {
    const b = JSON.parse(raw);
    if (!b || !Array.isArray(b.styles) || !b.styles.length) return null;
    b.source = 'memory';
    return b;
  } catch { return null; }
}

function rememberBrand(b) {
  try {
    localStorage.setItem(MEMORY_KEY, JSON.stringify(b));
    return true;
  } catch {
    return false;      // private window, or storage turned off
  }
}

function forgetBrand() {
  try { localStorage.removeItem(MEMORY_KEY); } catch { /* nothing to do */ }
}

async function brandFromHost() {
  try {
    const res = await fetch('brand.json', { cache: 'no-store' });
    if (!res.ok) return null;
    const b = await res.json();
    if (!b || !Array.isArray(b.styles) || !b.styles.length) return null;
    b.source = 'host';
    return b;
  } catch { return null; }
}

async function brandFromDocument() {
  let xml = null;
  await Word.run(async (context) => {
    const parts = context.document.customXmlParts.getByNamespace(NS);
    parts.load('items');
    await context.sync();
    if (!parts.items.length) throw new Error('No BrandDoc part found in this document.');
    const got = parts.items[0].getXml();
    await context.sync();
    xml = got.value;
  });
  const m = /<payload>([\s\S]*)<\/payload>/.exec(xml);
  if (!m) throw new Error('The BrandDoc part is there but carries no payload.');
  const doc = new DOMParser().parseFromString(xml, 'application/xml');
  const text = doc.getElementsByTagName('payload')[0].textContent;
  return JSON.parse(text);
}

const req = (set, ver) => {
  try { return Office.context.requirements.isSetSupported(set, ver); } catch { return false; }
};

function reportSupport() {
  const missing = [];
  if (!SUPPORT.styles) missing.push('reading and writing styles at all (WordApi 1.5)');
  if (!SUPPORT.typography) missing.push('letter spacing and all-caps (WordApiDesktop 1.3)');
  if (!SUPPORT.borders) missing.push('borders on a style (WordApiDesktop 1.1)');
  if (!missing.length) return;
  $('support').hidden = false;
  $('support').innerHTML = '<strong>This build of Word is missing:</strong><ul>' +
    missing.map((m) => `<li>${esc(m)}</li>`).join('') +
    '</ul><p style="margin:6px 0 0">Anything unsupported is listed after applying, not skipped quietly.</p>';
}

/* ------------------------------------------------------------------ brand */
const SOURCE_TEXT = {
  document: 'From this document. It carries its own brand, and the reasoning behind every '
            + 'paragraph in it.',
  memory:   'Remembered on this Mac. It is not in this document yet \u2014 Apply writes the styles in.',
  host:     'Served with the add-in. It is not in this document yet \u2014 Apply writes the styles in.',
};

function renderBrand() {
  const has = !!(BRAND && BRAND.styles);
  const set = (id, prop, value) => { const el = $(id); if (el) el[prop] = value; };
  set('apply', 'disabled', !has);
  set('design', 'disabled', !has);
  set('brandForget', 'hidden', !(BRAND && BRAND.source === 'memory'));

  if (!has) {
    set('client', 'textContent', 'No brand loaded');
    set('brandSource', 'textContent', 'Nothing to format with yet. Load a brand file below, or '
      + 'open a document the engine has already built.');
    set('palette', 'innerHTML', '');
    set('fonts', 'innerHTML', '');
    return;
  }
  set('client', 'textContent', BRAND.client + (BRAND.profile ? ` \u00b7 ${BRAND.profile}` : ''));
  set('brandSource', 'textContent', SOURCE_TEXT[BRAND.source] || '');
  if (BRAND.profile) {
    const opt = document.querySelector('#scope option[value="profile"]');
    if (opt) opt.textContent = `Always in ${BRAND.profile} documents`;
  }
  drawBrand();
}

async function onBrandFile(ev) {
  const file = ev.target.files && ev.target.files[0];
  if (!file) return;
  try {
    const b = JSON.parse(await file.text());
    if (!b || !Array.isArray(b.styles) || !b.styles.length) {
      throw new Error('That file has no styles in it.');
    }
    b.source = 'memory';
    BRAND = b;
    const kept = rememberBrand(b);
    renderBrand();
    show('brandResult', kept ? 'Loaded and remembered' : 'Loaded, but not remembered',
      kept
        ? `<p><strong>${esc(b.client || 'This brand')}</strong> with ${b.styles.length} styles. `
          + 'Every document you open from now on starts with it already here.</p>'
        : '<p>This Word build would not let the add-in keep it, so it lasts only while the pane '
          + 'is open. Nothing is wrong with the file.</p>');
  } catch (e) {
    show('brandResult', 'Could not read that file', `<p>${esc(e.message)}</p>`);
  } finally {
    ev.target.value = '';
  }
}

function onBrandForget() {
  forgetBrand();
  BRAND = null;
  renderBrand();
  show('brandResult', 'Forgotten', '<p>Nothing of it is left on this Mac.</p>');
}

function drawBrand() {
  $('palette').innerHTML = Object.entries(BRAND.palette).map(([name, hex]) =>
    `<div class="swatch"><div class="chip" style="background:#${hex}"></div><span>${esc(name)}</span></div>`
  ).join('');
  const role = { G: 'Display', SL: 'Text', R: 'Labels', SB: 'Emphasis' };
  $('fonts').innerHTML = Object.entries(BRAND.fonts).map(([k, v]) =>
    `<dt>${esc(role[k] || k)}</dt><dd>${esc(v)}</dd>`).join('');
}

/* ------------------------------------------------------------------ apply */
const ALIGN = { left: 'Left', right: 'Right', center: 'Centered', both: 'Justified' };

async function applyBrand() {
  if (!SUPPORT.styles) {
    return show('applyResult', 'Not supported here',
      '<p>This build of Word cannot write styles through the add-in API (it needs WordApi 1.5).</p>');
  }
  $('apply').disabled = true;
  $('apply').textContent = 'Applying…';
  const made = [], changed = [], failed = [], skipped = new Set();
  try {
    await Word.run(async (context) => {
      const styles = context.document.getStyles();
      styles.load('items/nameLocal,items/type');
      await context.sync();

      const have = new Map(styles.items.map((s) => [s.nameLocal.toLowerCase(), s]));
      for (const rec of BRAND.styles) {
        const key = rec.name.toLowerCase();
        let style = have.get(key);
        try {
          if (!style) {
            style = context.document.addStyle(rec.name, rec.type === 'character' ? 'Character' : 'Paragraph');
            made.push(rec.name);
          } else {
            changed.push(rec.name);
          }
          applyRun(style, rec, skipped);
          if (rec.type !== 'character') applyPara(style, rec, skipped);
        } catch (e) {
          failed.push(`${rec.name}: ${e.message}`);
        }
      }
      await context.sync();
    });
  } catch (e) {
    failed.push(e.message);
  }
  $('apply').disabled = false;
  $('apply').textContent = 'Apply brand to this document';

  let html = `<p><strong>${changed.length} styles redefined, ${made.length} created.</strong></p>`;
  if (skipped.size) {
    html += '<p class="quiet">Not applied, because this Word build or the add-in API cannot set them:</p><ul>' +
      [...skipped].map((s) => `<li>${esc(s)}</li>`).join('') + '</ul>';
  }
  if (failed.length) {
    html += '<p class="quiet">Failed:</p><ul>' + failed.slice(0, 8).map((f) => `<li>${esc(f)}</li>`).join('') + '</ul>';
  }
  show('applyResult', 'Result', html);
}

function applyRun(style, rec, skipped) {
  const r = rec.run || {};
  const f = style.font;
  if (r.font) f.name = BRAND.fonts[r.font] || r.font;
  if (r.size != null) f.size = r.size;
  if (r.color) f.color = '#' + (BRAND.palette[r.color] || r.color);
  if (r.bold != null) f.bold = r.bold;
  if (r.italic != null) f.italic = r.italic;
  if (r.caps != null) {
    if (SUPPORT.typography) f.allCaps = r.caps;
    else skipped.add('all-caps on tracked labels');
  }
  if (r.track != null) {
    if (SUPPORT.typography) f.spacing = r.track;
    else skipped.add('letter spacing (the tracked caps that labels depend on)');
  }
}

function applyPara(style, rec, skipped) {
  const p = rec.paragraph || {};
  const pf = style.paragraphFormat;
  if (p.jc) pf.alignment = ALIGN[p.jc] || 'Left';
  if (p.before != null) pf.spaceBefore = p.before;
  if (p.after != null) pf.spaceAfter = p.after;
  if (p.ind_left != null) pf.leftIndent = p.ind_left;
  if (p.ind_right != null) pf.rightIndent = p.ind_right;
  if (p.hanging != null) pf.firstLineIndent = -p.hanging;
  if (p.keep_next != null) pf.keepWithNext = p.keep_next;
  if (p.keep_lines != null) pf.keepTogether = p.keep_lines;
  if (p.widow != null) pf.widowControl = p.widow;
  if (rec.leading) {
    // The design asks for a Multiple, so the leading stays proportional if a size changes. The
    // add-in API only exposes lineSpacing in points and gives no way to set the rule, so setting
    // it here would pin the document to Exactly and quietly break that. Left to the engine.
    skipped.add('line spacing (the API has no way to set a Multiple, only an exact value)');
  }
  if (rec.borders && rec.borders.length && !SUPPORT.borders) {
    skipped.add('rules above and below the lead paragraph (style borders)');
  }
}

/* ------------------------------------------------------------------ design */
/* Read the draft, decide what each paragraph is, set the style that carries it. The reading is
   engine/classify.js, verified block for block against the Python engine; the mapping from a role
   to a style is engine/apply.js. Shown as a plan first, because a formatting pass over someone
   else's document should never be a surprise. */
let PLAN = null;

async function design() {
  const btn = $('design');
  btn.disabled = true;
  btn.textContent = 'Reading\u2026';
  try {
    await Word.run(async (context) => {
      const blocks = await BrandDocBlocks.fromDocument(context);
      const reading = (BRAND && BRAND.reading) || BrandDocClassify.NEUTRAL;
      const roles = BrandDocClassify.classify('', blocks, reading, { index: true });
      BrandDocClassify.structure(roles, reading);
      BrandDocClassify.addLeads(roles, reading);
      PLAN = BrandDocApply.plan(blocks, roles);
      showPlan(blocks);
    });
  } catch (e) {
    show('designResult', 'Could not read this document', `<p>${esc(e.message)}</p>`);
  }
  btn.disabled = false;
  btn.textContent = 'Design this document';
}

function tally(list, key) {
  const c = {};
  for (const x of list) c[x[key]] = (c[x[key]] || 0) + 1;
  return Object.entries(c).sort((a, b) => b[1] - a[1]);
}

function showPlan(blocks) {
  const byStyle = tally(PLAN.steps, 'style');
  const byRole = tally(PLAN.unplaced, 'role');
  const judged = PLAN.steps.filter((s) => s.confidence === 'judgement').length;

  let html = `<p><strong>${PLAN.steps.length} of ${blocks.length} blocks would be styled.</strong>`
    + ` ${PLAN.paragraphs} paragraphs in the document.</p>`;
  html += '<ul>' + byStyle.map(([k, v]) => `<li>${v} &times; ${esc(k)}</li>`).join('') + '</ul>';
  if (judged) {
    html += `<p class="quiet">${judged} of these are judgement calls rather than rules. `
      + 'Ask "Why this?" on any paragraph afterwards and it will say which.</p>';
  }
  if (byRole.length) {
    html += '<p class="quiet">Left alone, because a paragraph style cannot carry them:</p><ul>'
      + byRole.map(([k, v]) => `<li>${v} &times; ${esc(k)} \u2014 ${esc(BrandDocApply.NEEDS_ENGINE[k] || 'no style mapped')}</li>`).join('')
      + '</ul>';
  }
  if (BRAND && BRAND.source !== 'document') {
    html += '<p class="quiet">These styles are not in this document yet. Apply the brand first, '
      + 'or there will be nothing for this to set.</p>';
  }
  html += '<button id="designGo" class="primary">Set these styles</button>';
  show('designResult', 'The plan', html);
  $('designGo').onclick = designGo;
}

async function designGo() {
  if (!PLAN) return;
  const go = $('designGo');
  go.disabled = true;
  go.textContent = 'Setting\u2026';
  let result = null;
  try {
    await Word.run(async (context) => {
      result = await BrandDocApply.apply(context, PLAN.steps, {
        Word: Word,
        onProgress: (done, total) => { go.textContent = `Setting\u2026 ${done} of ${total}`; },
      });
    });
  } catch (e) {
    show('designResult', 'Stopped', `<p>${esc(e.message)}</p>`);
    return;
  }
  let html = `<p><strong>${result.applied.length} paragraphs styled.</strong></p>`;
  if (result.failed.length) {
    const why = tally(result.failed, 'error');
    html += `<p class="quiet">${result.failed.length} could not be set:</p><ul>`
      + why.map(([k, v]) => `<li>${v} &times; ${esc(k)}</li>`).join('') + '</ul>'
      + '<p class="quiet">A style that is missing has to be created first \u2014 that is what '
      + 'Apply does.</p>';
  }
  show('designResult', 'Done', html);
}

/* ------------------------------------------------------------------ inspect */
async function inspect() {
  try {
    await Word.run(async (context) => {
      const para = context.document.getSelection().paragraphs.getFirst();
      para.load('text,style,styleBuiltIn');
      await context.sync();
      LAST = { text: para.text, style: para.style };
      const hit = matchDecision(para.text);
      let html = `<p><strong>${esc(para.style || '(no style)')}</strong></p>` +
                 `<p class="quiet">“${esc((para.text || '').slice(0, 120))}”</p>`;
      if (hit) {
        html += `<p>${esc(hit.why)}</p>`;
        if (hit.rule) html += `<p class="quiet">Follows rule ${esc(hit.rule)}.</p>`;
        if (hit.alternatives) html += `<p class="quiet">Could have been: ${esc(hit.alternatives.join(', '))}.</p>`;
        if (hit.confidence === 'judgement') {
          html += '<p class="quiet">The engine weighed this one rather than deriving it — it is a judgement call.</p>';
        }
      } else {
        html += '<p class="quiet">No recorded decision for this paragraph. Either the engine did not ' +
                'format this document, or the text has been edited since.</p>';
      }
      show('inspectResult', 'This paragraph', html);
      $('correctSection').hidden = false;
    });
  } catch (e) {
    show('inspectResult', 'Could not read the selection', `<p>${esc(e.message)}</p>`);
  }
}

/* The decision log is keyed by block, which means nothing in an arbitrary document, so a
   paragraph is matched on its opening words instead. Good enough to show the reasoning, and
   honest about being approximate: an edited paragraph simply will not match. */
function matchDecision(text) {
  if (!BRAND || !BRAND.decisions || !text) return null;
  const key = text.trim().slice(0, 40);
  if (key.length < 8) return null;
  return BRAND.decisions.find((d) => d.text && d.text.trim().slice(0, 40) === key) || null;
}

/* ------------------------------------------------------------------ corrections */
const ROLES = ['part', 'h2', 'h3', 'h4', 'lead', 'body', 'tabletitle', 'caption', 'source',
               'bullet', 'listnum', 'exhibit', 'table'];

function fillRoles() {
  $('newRole').innerHTML = ROLES.map((r) => `<option value="${r}">${r}</option>`).join('');
}

/* Which of the four levels a correction belongs to. The client and profile names come from the
   document's payload, so this file names no client -- it is published, and the document is not. */
function scopeFile(scope) {
  const c = (BRAND && BRAND.clientKey) || (BRAND && BRAND.client) || 'Client';
  const p = (BRAND && BRAND.profile) || 'Profile';
  return {
    exception: `02_Clients/${c}/exceptions`,
    profile: `02_Clients/${c}/profiles/${p}`,
    brand: `02_Clients/${c}/brand`,
    core: '01_Core-System/validation/Layout-Checks.md',
  }[scope];
}

async function record() {
  if (!LAST) return;
  const entry = {
    at: new Date().toISOString(),
    text: (LAST.text || '').slice(0, 160),
    was: LAST.style,
    should: $('newRole').value,
    scope: $('scope').value,
    because: $('because').value.trim(),
  };
  const log = await readLog();
  log.push(entry);
  await writeLog(log);
  $('because').value = '';
  show('recordResult', 'Recorded',
    `<p>Kept with this document. It belongs in <strong>${esc(scopeFile(entry.scope))}</strong>.</p>` +
    '<p class="quiet">The add-in cannot write to the repository from inside Word, so this waits ' +
    'here until it is folded into the rules.</p>');
  await drawLog();
}

const settings = () => Office.context.document.settings;

async function readLog() {
  const raw = settings().get('brandDocCorrections');
  return raw ? JSON.parse(raw) : [];
}

function writeLog(log) {
  settings().set('brandDocCorrections', JSON.stringify(log));
  return new Promise((resolve, reject) => {
    settings().saveAsync((r) => (r.status === Office.AsyncResultStatus.Succeeded
      ? resolve() : reject(new Error(r.error.message))));
  });
}

async function drawLog() {
  const log = await readLog();
  $('log').innerHTML = log.length
    ? log.map((e) => `<li><strong>${esc(e.was || '?')} → ${esc(e.should)}</strong>` +
        `<div class="where">${esc(scopeFile(e.scope))}</div>` +
        (e.because ? `<div>${esc(e.because)}</div>` : '') +
        `<div class="where">“${esc(e.text.slice(0, 60))}”</div></li>`).join('')
    : '<li class="empty">None yet.</li>';
}

async function copyLog() {
  const log = await readLog();
  const text = log.map((e) =>
    `- ${e.was || '?'} -> ${e.should} [${scopeFile(e.scope)}]\n  "${e.text}"\n` +
    (e.because ? `  because: ${e.because}\n` : '')).join('\n');
  try {
    await navigator.clipboard.writeText(text || '(nothing recorded)');
    show('recordResult', 'Copied', '<p>Paste it wherever the rules get written.</p>');
  } catch {
    show('recordResult', 'Copy it from here', `<p class="quiet">${esc(text)}</p>`);
  }
}

/* ------------------------------------------------------------------ helpers */
function show(id, title, html) {
  const el = $(id);
  el.hidden = false;
  el.innerHTML = `<h3>${esc(title)}</h3>${html}`;
}

function esc(s) {
  return String(s == null ? '' : s).replace(/[&<>"']/g,
    (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
}
