#!/usr/bin/env python3
"""Rebuild preview.html from taskpane.html.

The preview is the pane with office.js stubbed out, so it can be looked at in a browser. Keeping
it as a separate hand-written file meant it drifted from the pane within a day, twice. It is
generated now: edit taskpane.html, then run this.
"""
import re, sys, pathlib

here = pathlib.Path(__file__).resolve().parent.parent
prev = (here / 'preview.html').read_text(encoding='utf-8')
task = (here / 'taskpane.html').read_text(encoding='utf-8')

m = re.search(r'<script>\n(/\* Design review.*?)</script>', prev, re.S)
if not m:
    sys.exit('preview.html has no Office stub to carry over')

out = task.replace('<title>BrandDoc</title>', '<title>BrandDoc preview</title>')
out = out.replace(
    '<script src="https://appsforoffice.microsoft.com/lib/1/hosted/office.js"></script>',
    '<!-- Generated from taskpane.html. Do not edit: run tools/make-preview.py instead. -->\n'
    '<script>\n' + m.group(1) + '</script>')
(here / 'preview.html').write_text(out, encoding='utf-8')
print('preview.html rebuilt from taskpane.html')
