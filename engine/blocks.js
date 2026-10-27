/* The open document, read into the shape the classifier expects.
 *
 * A port of the engine's extract.py. Office.js has no run-level API — paragraph.font.bold is
 * null the moment a paragraph mixes bold with plain — so this reads the document's own
 * WordprocessingML instead, which is what the Python side already parses. One body.getOoxml()
 * call returns the lot, so there is no per-paragraph round trip.
 *
 * Content is preserved exactly: every run's text, its bold/italic/caps, footnote references in
 * place, table cells, and the position of each image. Nothing is rewritten here.
 */
(function (root, factory) {
  if (typeof module === 'object' && module.exports) module.exports = factory();
  else root.BrandDocBlocks = factory();
})(typeof self !== 'undefined' ? self : this, function () {
  'use strict';

  const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
  const PKG = 'http://schemas.microsoft.com/office/2006/xmlPackage';

  // Direct children only, matching ElementTree's `for x in node` / node.find().
  function kids(node, name) {
    const out = [];
    for (let c = node.firstElementChild; c; c = c.nextElementSibling) {
      if (c.namespaceURI === W && c.localName === name) out.push(c);
    }
    return out;
  }
  const kid = (node, name) => kids(node, name)[0] || null;
  const attr = (el, name) => (el && el.getAttributeNS(W, name));

  /* Walk w:r children so run properties stay attached to their text. */
  function paraRuns(p) {
    const out = [];
    for (let r = p.firstElementChild; r; r = r.nextElementSibling) {
      if (r.namespaceURI !== W || r.localName !== 'r') continue;
      const rPr = kid(r, 'rPr');
      const fmt = {};
      if (rPr) {
        const b = kid(rPr, 'b');
        if (b) { const v = b.getAttributeNS(W, 'val'); if (v !== '0' && v !== 'false') fmt.b = true; }
        const i = kid(rPr, 'i');
        if (i) { const v = i.getAttributeNS(W, 'val'); if (v !== '0' && v !== 'false') fmt.i = true; }
        if (kid(rPr, 'caps')) fmt.caps = true;
        const va = kid(rPr, 'vertAlign');
        if (va) fmt.va = va.getAttributeNS(W, 'val');
        const sz = kid(rPr, 'sz');
        if (sz) fmt.sz = parseInt(sz.getAttributeNS(W, 'val'), 10) / 2;
      }
      for (let el = r.firstElementChild; el; el = el.nextElementSibling) {
        if (el.namespaceURI !== W) continue;
        const t = el.localName;
        if (t === 't') out.push(Object.assign({}, fmt, { t: el.textContent || '' }));
        else if (t === 'footnoteReference') out.push({ fn: el.getAttributeNS(W, 'id') });
        else if (t === 'drawing') out.push({ img: true });
        else if (t === 'br') out.push({ br: true });
        else if (t === 'tab') out.push(Object.assign({}, fmt, { t: '\t' }));
      }
    }
    return out;
  }

  function paraObj(p) {
    const pPr = kid(p, 'pPr');
    let style = null, num = null, jc = null;
    if (pPr) {
      const st = kid(pPr, 'pStyle');
      style = st ? attr(st, 'val') : null;
      const np = kid(pPr, 'numPr');
      if (np) { const n = kid(np, 'numId'); num = n ? attr(n, 'val') : null; }
      const j = kid(pPr, 'jc');
      jc = j ? attr(j, 'val') : null;
    }
    const runs = paraRuns(p);
    return {
      kind: 'p', style: style, num: num, jc: jc, runs: runs,
      text: runs.map(r => (typeof r.t === 'string' ? r.t : '')).join('')
    };
  }

  function tableObj(tbl) {
    const rows = [];
    for (const tr of kids(tbl, 'tr')) {
      const cells = [];
      for (const tc of kids(tr, 'tc')) cells.push(kids(tc, 'p').map(paraObj));
      rows.push(cells);
    }
    return { kind: 'table', rows: rows };
  }

  /* w:body -> blocks, in document order. */
  function fromBody(body) {
    const blocks = [];
    for (let el = body.firstElementChild; el; el = el.nextElementSibling) {
      if (el.namespaceURI !== W) continue;
      if (el.localName === 'p') blocks.push(paraObj(el));
      else if (el.localName === 'tbl') blocks.push(tableObj(el));
    }
    return blocks;
  }

  function parse(xml, parser) {
    const p = parser || (s => new DOMParser().parseFromString(s, 'application/xml'));
    const doc = p(xml);
    const err = doc.getElementsByTagName('parsererror')[0];
    if (err) throw new Error('could not parse the document XML: ' + err.textContent.slice(0, 200));
    return doc;
  }

  /* A bare word/document.xml. */
  function fromDocumentXml(xml, parser) {
    const doc = parse(xml, parser);
    const body = doc.getElementsByTagNameNS(W, 'body')[0];
    if (!body) throw new Error('no w:body in the document XML');
    return fromBody(body);
  }

  /* What Office.js body.getOoxml() actually hands back: a flat-OPC package. */
  function fromPackage(xml, parser) {
    const doc = parse(xml, parser);
    const parts = doc.getElementsByTagNameNS(PKG, 'part');
    for (const part of parts) {
      if (part.getAttributeNS(PKG, 'name') !== '/word/document.xml') continue;
      const body = part.getElementsByTagNameNS(W, 'body')[0];
      if (!body) break;
      return fromBody(body);
    }
    // Word has also been seen to return the document part alone.
    const body = doc.getElementsByTagNameNS(W, 'body')[0];
    if (body) return fromBody(body);
    throw new Error('no /word/document.xml part in the package Word returned');
  }

  /* The one call the pane makes. */
  async function fromDocument(context) {
    const ooxml = context.document.body.getOoxml();
    await context.sync();
    return fromPackage(ooxml.value);
  }

  return { fromDocument, fromPackage, fromDocumentXml, fromBody, paraObj, tableObj, paraRuns };
});
