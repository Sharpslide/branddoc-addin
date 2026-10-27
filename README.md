# BrandDoc task pane

The Word add-in. Early, and it does three things: shows the brand, writes the brand's styles into
the open document, and explains why a paragraph was formatted as it was so the decision can be
corrected.

## Installing on a Mac

1. Finder > Go > **Go to Folder** (`Cmd`+`Shift`+`G`) and enter
   `~/Library/Containers/com.microsoft.Word/Data/Documents/wef` — create the `wef` folder if it
   isn't there.
2. Copy `manifest.xml` into it.
3. Restart Word, open a document, then **Home > Add-ins > BrandDoc**.

One file, one time, no Terminal and no store submission.
([Microsoft's instructions](https://learn.microsoft.com/en-us/office/dev/add-ins/testing/sideload-an-office-add-in-on-mac))

`SourceLocation` in the manifest has to point at wherever these files are served from. It is the
only value that changes between hosts.

## It carries no client data

The brand and the decision log ride **inside each .docx**, as a custom XML part under the
namespace `urn:sharpslide:branddoc:1`. `docparts.brand_part` in the engine writes it.

That is deliberate. Serving them from the add-in's web host would publish a client's unreleased
brand system, and because every decision quotes the text it was made about, excerpts of a document
whose own footer reads *Strictly Confidential*. The document is already the confidential thing —
anyone who can open it can already read every word. It also means a second client's document
simply brings its own brand, which is how this has to work anyway.

So `brand.json` here holds nothing. Open a document the engine has not built and the panel says so
rather than pretending.

## What it can and cannot write

Styles are set through the Style API — `getStyles` / `addStyle`, then `style.font` and
`style.paragraphFormat`. Not by injecting OOXML, which looks neater and would reuse the XML the
engine already produces, but Word documents the catch: *"if a style is already in use in the
destination document, your content will take on the resident definition for the style, regardless
of what you include in styles.xml"*. A draft's Heading 1 is always already in use, so injection
cannot restyle the one thing this add-in exists to restyle.

| | |
|---|---|
| Typeface, size, colour, bold, italic | `WordApi 1.5` |
| Alignment, space before/after, indents, keep-with-next, keep-together, widow control | `WordApi 1.5` |
| Letter spacing, all-caps | `WordApiDesktop 1.3` — desktop only, not Word on the web |
| Borders on a style | `WordApiDesktop 1.1` — desktop only |
| **Line spacing as a Multiple** | **not possible.** `paragraphFormat.lineSpacing` is points only and the API exposes no line rule. Setting it would pin the document to Exactly and break the standing preference, so it is left alone. |

Leading is the one value neither direction handles: `sync_styles.py` reports it rather than reading
it back, and the add-in reports it rather than writing it. In both cases because the design records
a *target* and the engine computes the multiple from measured font metrics, and Word only ever
stores the computed answer.

The panel checks what the running build of Word supports and lists anything it could not apply.
A style sheet that quietly applied two thirds of itself would be worse than one that refused.

## Corrections

Inspect a paragraph, say what it should have been, and choose how far the correction goes — a
one-off here, always in PPMs, always for O'Connor, always. That is the four-level separation, so
each correction knows which file it belongs in.

They are kept in the document's own settings for now, because an add-in cannot write to the
repository from inside Word. **Copy all as text** hands them over to be folded into the rules.

## preview.html

Design review outside Word. It stubs just enough of `office.js` to render the panel, with a
sample payload inlined. Word loads `taskpane.html`; this file is only for looking at.
