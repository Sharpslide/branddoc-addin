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
    $('client').textContent = BRAND.client + (BRAND.profile ? ` \u00b7 ${BRAND.profile}` : '');
    if (BRAND.profile) {
      const opt = document.querySelector('#scope option[value="profile"]');
      if (opt) opt.textContent = `Always in ${BRAND.profile} documents`;
    }
    drawBrand();
    fillRoles();
  } catch (e) {
    $('client').textContent = 'No brand in this document';
    $('apply').disabled = true;
    show('applyResult', 'This document carries no brand',
      '<p>The brand travels inside the document, not with the add-in, so a document the engine ' +
      'has not built has nothing to read.</p>' +
      `<p class="quiet">${esc(e.message)}</p>`);
  }
  $('apply').onclick = applyBrand;
  $('inspect').onclick = inspect;
  $('record').onclick = record;
  $('copyLog').onclick = copyLog;
  await drawLog();
});

/* The brand and the reasoning come out of the document's own custom XML part. */
async function loadBrand() {
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
