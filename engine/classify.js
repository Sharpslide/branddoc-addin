/* What each paragraph of a draft IS.
 *
 * A port of the engine's classify.py, so the decision can be made inside Word instead of in a
 * Python process the client does not have. The rules are the same ones and must stay the same:
 * engine/classify.test.js replays the engine's own fixture through this file and fails on any
 * divergence.
 *
 * Pure logic. It takes blocks and returns roles, knows nothing about Office.js, and carries no
 * client data — the blocks are handed in.
 */
(function (root, factory) {
  if (typeof module === 'object' && module.exports) module.exports = factory();
  else root.BrandDocClassify = factory();
})(typeof self !== 'undefined' ? self : this, function () {
  'use strict';

  const SMALL = new Set(['a', 'an', 'and', 'as', 'at', 'by', 'for', 'from', 'in', 'of', 'on',
    'or', 'the', 'to', 'with', 'vs', 'via']);
  /* Everything client-specific arrives as a profile. Nothing about one client belongs in
   * here: in Word the profile is read from the document's own custom XML part, and the
   * neutral default below simply does less rather than guessing.
   *
   *   acronyms      words title case must not touch
   *   sectionNames  Heading 1 texts that name their own section, so they open a title page
   *   groupOpeners  Heading 2 texts that parent the headings under them
   *   topics        Heading 2 text -> how to read the Heading 3s beneath it
   */
  const NEUTRAL = { acronyms: [], sectionNames: [], groupOpeners: [], topics: {} };

  function asProfile(p) {
    p = p || NEUTRAL;
    return {
      acronyms: new Set(p.acronyms || []),
      sectionNames: new Set((p.sectionNames || []).map(x => x.toUpperCase())),
      groupOpeners: new Set((p.groupOpeners || []).map(x => x.toLowerCase())),
      topics: p.topics || {}
    };
  }

  // Python's str.strip(chars): both ends, any of these, repeatedly.
  const PUNCT = '“”"(),.:;';
  function strip(s) {
    let a = 0, b = s.length;
    while (a < b && PUNCT.indexOf(s[a]) !== -1) a++;
    while (b > a && PUNCT.indexOf(s[b - 1]) !== -1) b--;
    return s.slice(a, b);
  }

  function titlecase(s, acronyms) {
    acronyms = acronyms || new Set();
    if (!/[A-Z]/.test(s) || s !== s.toUpperCase()) return s;   // already mixed case: leave it
    const words = s.split(' ');
    const out = [];
    for (let i = 0; i < words.length; i++) {
      const w = words[i];
      const core = strip(w);
      if (acronyms.has(core) && core !== 'A') { out.push(w); continue; }
      if (/^\d+\.$/.test(w)) { out.push(w); continue; }
      const low = w.toLowerCase();
      if (i && SMALL.has(strip(low)) && i !== words.length - 1) { out.push(low); continue; }
      out.push(low.split('-').map(p => p.slice(0, 1).toUpperCase() + p.slice(1)).join('-'));
    }
    let t = out.join(' ');
    t = t.replace(/\bU\.s\./g, 'U.S.');
    t = t.replace(/\bE-Commerce\b/g, 'E-commerce');
    t = t.replace(/“a”/g, '“A”');
    return t;
  }

  // Hang the reason on a block, so it travels with the decision (decisions.attach).
  function attach(d, why, rule, alts, confidence) {
    d.why = why;
    d.confidence = confidence || 'certain';
    if (rule) d.rule = rule;
    if (alts && alts.length) d.alternatives = alts;
    return d;
  }

  const txtOf = r => (r && typeof r.t === 'string') ? r.t : '';
  const isCaption = b => /^\s*Exhibit\s+\d+\s*[—–-]/.test(b.text);
  const isSource = b => b.text.trim().indexOf('Source:') === 0;
  const hasImg = b => (b.runs || []).some(r => !!r.img);

  /* A paragraph that opens with a bold lead-in ending in a colon. */
  function runinSplit(b) {
    const runs = b.runs || [];
    const lead = [];
    for (const r of runs) {
      if (txtOf(r) && r.b) lead.push(r);
      else if (txtOf(r).trim() === '') lead.push(r);
      else break;
    }
    const txt = lead.map(txtOf).join('');
    if (lead.length && txt.indexOf(':') !== -1 && txt.length < 90) return lead.length;
    return 0;
  }

  function classify(sec, blocks, profile) {
    const P = asProfile(profile);
    const tc = t => titlecase(t, P.acronyms);
    const out = [];
    let firstStyled = blocks.length;
    for (let k = 0; k < blocks.length; k++) {
      const q = blocks[k];
      if (q.kind === 'table' || q.style) { firstStyled = k; break; }
    }
    for (let i = 0; i < blocks.length; i++) {
      const b = blocks[i];
      if (b.kind === 'table') {
        out.push(attach({ role: 'table', rows: b.rows }, 'A table in the draft stays a table.'));
        continue;
      }
      const t = b.text.trim();
      const st = b.style;
      if (!t && !hasImg(b)) continue;
      // The draft's own title block, replaced by the section title page (D13a). Identified by
      // sitting ahead of all the styled content, not by being inside the first few paragraphs.
      if (i < firstStyled && st == null) {
        attach(b, 'Part of the title block the draft opens with, ahead of any styled paragraph. '
          + 'Dropped, because the design gives the section a title page instead.', 'D13a');
        continue;
      }
      if (st === 'Heading1') {
        if (P.sectionNames.has(t.toUpperCase())) {
          out.push(attach({ role: 'sectiontitle', text: tc(t), num: sec },
            'A Heading 1 carrying the section’s own name, so it becomes the section title '
            + 'page rather than a heading inside it.'));
        } else if (t.toUpperCase().indexOf('(CONT') !== -1) {
          attach(b, 'A “continued” heading the draft repeats at a page break. Dropped: '
            + 'the formatted document repeats nothing across pages.', 'D4');
          continue;
        } else {
          out.push(attach({ role: 'part', text: t },
            'A Heading 1 that is not the section’s own name, so it divides the section '
            + 'rather than opening it.'));
        }
        continue;
      }
      const fns = (b.runs || []).filter(r => 'fn' in r).map(r => r.fn);
      if (st === 'Heading2') {
        const numbered = /^\d+\.\s/.test(t);
        out.push(attach({ role: numbered ? 'h3' : 'h2', text: tc(t), fns: fns },
          numbered
            ? 'Numbered, so it sits one level below the unnumbered Heading 2s around it.'
            : 'A Heading 2 in the draft, kept at that level.',
          null, numbered ? ['h2'] : ['h3']));
        continue;
      }
      if (st === 'Heading3') {
        out.push(attach({ role: 'h4', text: tc(t), fns: fns },
          'A Heading 3 in the draft, which the profile sets as its fourth level.'));
        continue;
      }
      if (isCaption(b)) {
        const m = /^\s*Exhibit\s+(\d+)\s*[—–-]\s*([\s\S]*)$/.exec(t);
        out.push(attach({ role: 'caption', n: parseInt(m[1], 10), text: m[2] },
          'Opens “Exhibit N —”, which is how this author marks a figure caption.'));
        continue;
      }
      if (hasImg(b)) {
        out.push(attach({ role: 'exhibit' },
          'Carries a picture, so this paragraph is the exhibit itself.'));
        continue;
      }
      if (isSource(b)) {
        out.push(attach({ role: 'source', text: t },
          'Opens “Source:”, so it credits the exhibit above it.'));
        continue;
      }
      if (t.indexOf('•') === 0) {
        out.push(attach({ role: 'bullet', runs: b.runs, strip_bullet: true },
          'Typed with a bullet character rather than using a list style. Becomes a real list '
          + 'item and the typed bullet comes off.'));
        continue;
      }
      if (b.num) {
        out.push(attach({ role: 'listnum', runs: b.runs, runin: runinSplit(b) },
          'The draft numbered it, so it becomes a numbered list item.'));
        continue;
      }
      // a bold all-caps line on its own is a table title in the client's drafts
      const inked = (b.runs || []).filter(r => txtOf(r).trim());
      if (t === t.toUpperCase() && t.length < 60 && inked.every(r => !!r.b)) {
        out.push(attach({ role: 'tabletitle', text: tc(t), runs: b.runs },
          'A short bold line in capitals, standing alone. In this author’s drafts that marks '
          + 'the title of the table below it — but nothing says so, and a heading can look '
          + 'exactly the same.',
          null, ['h2', 'h4', 'body'], 'judgement'));
        continue;
      }
      out.push(attach({ role: 'body', runs: b.runs, runin: runinSplit(b) },
        'Nothing marks it as anything else, so it is running text.'));
    }
    return out;
  }

  const YEAR = /^\s*(1[89]\d0s|\d{4})\s*(?:[–—-]\s*\d{4})?\s*$/;

  /* A second pass that reads roles from the heading a block sits under. */
  function structure(blocks, profile) {
    const P = asProfile(profile);
    const T = P.topics;                       // topic -> 'timeline' | 'bio' | 'principle'
    let topic = null;
    const n = { timeline: 0, bio: 0, principle: 0, dropped: 0 };
    for (let i = 0; i < blocks.length; i++) {
      const b = blocks[i];
      if (b.role === 'h2') { topic = b.text.trim().toLowerCase(); continue; }
      if (b.role !== 'h4') continue;
      const t = (b.text || '').trim();

      // only the block immediately after is ever taken
      const next = blocks[i + 1];
      const takeNext = (why, role) => {
        if (next && next.role === 'body') { attach(next, why); next.role = role; }
      };

      if (T[topic] === 'timeline' && YEAR.test(t)) {
        attach(b, 'A year standing alone under History, so it is an entry in the chronology '
          + 'rather than a heading in its own right.', null, ['h4']);
        b.role = 'tlyear';
        takeNext('The paragraph belonging to the year above it.', 'tlbody');
        n.timeline += 1;
      } else if (T[topic] === 'bio' && t.indexOf(',') !== -1) {
        const c = t.indexOf(',');
        b.name = t.slice(0, c).trim();
        b.title = t.slice(c + 1).trim();
        attach(b, 'A name and a job title separated by a comma, under a heading that introduces '
          + 'people. The two are set differently, so they are split here rather than left as '
          + 'one line.', null, ['h4']);
        b.role = 'bioname';
        takeNext('The biography belonging to the name above it.', 'biotext');
        n.bio += 1;
      } else if (T[topic] === 'principle') {
        if (t.toLowerCase() === topic) {
          // The same words as the H2 directly above. The H2 is the one with a display size.
          attach(b, 'Repeats the heading immediately above it word for word.', 'E4');
          b.role = 'drop';
          n.dropped += 1;
        } else {
          attach(b, 'A named principle under Guiding Principles: a label beside its own '
            + 'paragraph, as the client’s prior PPM sets them.', null, ['h4']);
          b.role = 'principle';
          takeNext('The text belonging to the principle beside it.', 'principletext');
          n.principle += 1;
        }
      }
    }
    return n;
  }

  /* A heading whose opening paragraph introduces the headings that follow it rather than the
   * blocks beneath it. The draft marks every one of these at the same level, so the
   * relationship is not in the structure and cannot be read out of it — it is named in the
   * profile, so the choice stays visible and can be corrected.
   */
  function addLeads(blocks, profile) {
    const P = asProfile(profile);
    let n = 0;
    for (let i = 0; i < blocks.length; i++) {
      const b = blocks[i];
      if (b.role !== 'h2') continue;
      let j = i + 1;
      while (j < blocks.length
             && ['h2', 'part', 'sectiontitle'].indexOf(blocks[j].role) === -1) j++;
      const under = blocks.slice(i + 1, j).filter(c => c.role !== 'drop');
      if (!under.length || under[0].role !== 'body' || under[0].runin) continue;
      const rest = under.slice(1);
      let opens = rest.some(c => c.role !== 'body') || rest.length >= 2;
      let why = 'The paragraph that opens the topic, and the heading opens a group: '
        + rest.length + ' more blocks sit under it before the next heading, so the paragraph '
        + 'introduces what follows rather than being all of it.';
      if (!opens && P.groupOpeners.has(b.text.trim().toLowerCase())) {
        opens = true;
        why = 'The paragraph that opens the topic. The heading has only this paragraph of its '
          + 'own, but it is the parent of the headings that follow it, so the paragraph '
          + 'introduces them. The draft marks them all at the same level, so this one is named '
          + 'in the profile rather than read out of the structure.';
      }
      if (opens) {
        attach(under[0], why, 'D3', ['body'], 'judgement');
        under[0].role = 'lead';
        n += 1;
      }
    }
    return n;
  }

  return { titlecase, classify, structure, addLeads, runinSplit, isCaption, isSource,
           hasImg, NEUTRAL };
});
