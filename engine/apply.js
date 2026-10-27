/* Turning roles into Word styles.
 *
 * classify.js says what each paragraph IS. This says which style carries that, and sets it.
 * Built-ins go through styleBuiltIn, which is locale-safe; the brand's own styles go through
 * the display name, because styleBuiltIn has no entry for them. The style must already exist in
 * the document — the pane's Apply step writes the definitions first.
 *
 * Roles this cannot place are reported, never guessed at. A pane that silently did two thirds
 * of the job would be worse than one that says which third it left.
 */
(function (root, factory) {
  if (typeof module === 'object' && module.exports) module.exports = factory();
  else root.BrandDocApply = factory();
})(typeof self !== 'undefined' ? self : this, function () {
  'use strict';

  // role -> the style that carries it. `builtIn` is a Word.Style key; `name` is a display name.
  const STYLES = {
    part:        { name: 'Section Label' },
    h2:          { builtIn: 'heading2', name: 'Heading 2' },
    h3:          { builtIn: 'heading3', name: 'Heading 3' },
    h4:          { builtIn: 'heading4', name: 'Heading 4' },
    lead:        { name: 'Lead Paragraph' },
    body:        { builtIn: 'normal', name: 'Normal' },
    listnum:     { builtIn: 'listNumber', name: 'List Number' },
    bullet:      { builtIn: 'listBullet', name: 'List Bullet' },
    caption:     { builtIn: 'caption', name: 'Caption' },
    source:      { name: 'Source Note' },
    tabletitle:  { name: 'Table Title' },
    bioname:     { name: 'Bio Name' },
    biotext:     { name: 'Bio Text' },
    // The timeline alternates across its spine, so these take a side.
    tlyear:      { sides: ['Timeline Year Left', 'Timeline Year Right'] },
    tlbody:      { sides: ['Timeline Body Left', 'Timeline Body Right'] }
  };

  /* Roles a paragraph style cannot express, with the reason. The engine still owns these. */
  const NEEDS_ENGINE = {
    sectiontitle:   'becomes a title page of its own',
    exhibit:        'is a placed photograph, sized and cropped to the grid',
    table:          'is a table, styled cell by cell',
    principle:      'is set as a label beside its own paragraph, not above it',
    principletext:  'is the paragraph beside a principle label',
    drop:           'is removed from the formatted document'
  };

  /* Where each block starts in Word's flat list of paragraphs. A table block covers every
   * paragraph in every one of its cells, so it advances the count by more than one. */
  function paragraphIndexOfBlock(blocks) {
    const at = [];
    let n = 0;
    for (const b of blocks) {
      at.push(n);
      if (b.kind === 'table') {
        let c = 0;
        for (const row of (b.rows || [])) for (const cell of row) c += cell.length;
        n += c;
      } else n += 1;
    }
    return { at: at, total: n };
  }

  /* What would be done, before anything is. classified must come from classify(..., {index:true}). */
  function plan(blocks, classified) {
    const idx = paragraphIndexOfBlock(blocks);
    const steps = [];
    const unplaced = [];
    const sides = {};
    for (const c of classified) {
      const role = c.role;
      if (NEEDS_ENGINE[role]) {
        unplaced.push({ role: role, why: NEEDS_ENGINE[role], text: (c.text || '').slice(0, 60) });
        continue;
      }
      const spec = STYLES[role];
      if (!spec) {
        unplaced.push({ role: role, why: 'no style is mapped to this role', text: (c.text || '').slice(0, 60) });
        continue;
      }
      if (typeof c._i !== 'number') {
        unplaced.push({ role: role, why: 'classified without a source index', text: (c.text || '').slice(0, 60) });
        continue;
      }
      let name = spec.name, builtIn = spec.builtIn;
      if (spec.sides) {
        const k = role === 'tlbody' ? 'tlyear' : role;      // a year and its text share a side
        if (role === 'tlyear') sides[k] = (sides[k] || 0) + 1;
        name = spec.sides[((sides[k] || 1) - 1) % 2];
        builtIn = null;
      }
      steps.push({
        paragraph: idx.at[c._i], block: c._i, role: role,
        style: name, builtIn: builtIn || null,
        text: (c.text || (c.runs || []).map(r => r.t || '').join('')).slice(0, 60),
        why: c.why, confidence: c.confidence
      });
    }
    return { steps: steps, unplaced: unplaced, paragraphs: idx.total };
  }

  function setStyle(paragraph, step, WordNS) {
    const key = step.builtIn;
    if (key && WordNS && WordNS.Style && WordNS.Style[key]) paragraph.styleBuiltIn = WordNS.Style[key];
    else paragraph.style = step.style;
  }

  /* Apply in chunks. A bad style name only surfaces on sync, so a failed chunk is replayed one
   * paragraph at a time to find which step caused it rather than blaming the whole chunk. */
  async function apply(context, steps, options) {
    const o = options || {};
    const WordNS = o.Word || (typeof Word !== 'undefined' ? Word : null);
    const chunk = o.chunk || 40;
    const paras = context.document.body.paragraphs;
    paras.load('items');
    await context.sync();
    const items = paras.items;

    const applied = [], failed = [];
    for (let i = 0; i < steps.length; i += chunk) {
      const slice = steps.slice(i, i + chunk).filter(s => {
        if (items[s.paragraph]) return true;
        failed.push(Object.assign({}, s, { error: 'no paragraph at that position' }));
        return false;
      });
      if (!slice.length) continue;
      try {
        for (const s of slice) setStyle(items[s.paragraph], s, WordNS);
        await context.sync();
        applied.push.apply(applied, slice);
      } catch (e) {
        for (const s of slice) {
          try {
            setStyle(items[s.paragraph], s, WordNS);
            await context.sync();
            applied.push(s);
          } catch (e2) {
            failed.push(Object.assign({}, s, { error: String((e2 && e2.message) || e2) }));
          }
        }
      }
      if (o.onProgress) o.onProgress(Math.min(i + chunk, steps.length), steps.length);
    }
    return { applied: applied, failed: failed };
  }

  return { plan, apply, STYLES, NEEDS_ENGINE, paragraphIndexOfBlock };
});
