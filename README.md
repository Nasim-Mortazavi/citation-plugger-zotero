# zotero_linker — make your typed citations into real Zotero citations

You wrote `(Nasim et al., 2026)` by hand in Word. This agent finds every such
citation, looks it up in your Zotero library, and replaces it with a *live*
Zotero citation field — identical to what the Zotero Word plugin inserts when
you click "Add citation". Then Zotero can re-format everything and build the
bibliography.

## Setup (once)
```bash
pip install -r requirements.txt     # python-docx, requests (+ pywin32 on Windows)
```
Zotero 7 → Settings → Advanced → tick **Allow other applications on this computer to communicate with Zotero**. Keep Zotero open.

Optional but recommended: find your numeric user ID at https://www.zotero.org/settings/keys
and pass `--user-id 123456` (or `export ZOTERO_USER_ID=123456`). Then the fields point
straight at your library items. Without it, Zotero will ask once, on first Refresh,
to link the items — the embedded item data makes that a one-click step.

## Run
```bash
python zotero_linker.py examples/sample.docx --dry-run             # detection demo on the bundled sample (its fake references won't be in your library)
python zotero_linker.py thesis.docx --user-id 123456              # -> thesis_zotero.docx
python zotero_linker.py thesis.docx --style ieee --interactive     # choose when ambiguous
python zotero_linker.py thesis.docx --dry-run                      # report only
```
Then open `thesis_zotero.docx` in Word → **Zotero → Refresh**. Citations re-format in your
style and the bibliography is generated where you wrote `[[bibliography]]`
(or at the end of the document).

## What it recognises
```
(Nasim et al., 2026)   (Nasim, 2026)   (Nasim & Lee, 2026)   (Nasim and Lee 2026)
(Nasim et al., 2026, p. 12)   (see Lee & Park, 2024, pp. 10-12; Moore, 1999)
Nasim et al. (2026)            ← narrative: author text kept, year becomes the field
[[cite: any free-text query]]  ← optional explicit placeholder
```
Matching checks first-author surname + year (+ co-author / "et al." consistency).
Also `(Smith, in press)`, `(Smith, n.d.)` and `(Smith, forthcoming)`. Accents are ignored when
matching (`Muller` finds `Müller`). Authors are matched against the item's *author* creators
(editors if there are none), so translators etc. don't break the match.
Citations it can't find stay as plain text and are listed at the end of the run,
so nothing is lost. Run it on a copy — the output is a new file.

## Using the Zotero web API instead
```bash
export ZOTERO_USER_ID=123456 ZOTERO_API_KEY=xxxx
python zotero_linker.py thesis.docx --web
```

## Live mode — convert just what you have selected
`zotero_linker_selection.py` does the same lookup-and-link, but on whatever
text is currently **selected in the open Word document** — no file path, no
save/reopen. Select a citation (or a whole passage with several) in Word,
then run it; it replaces the selection in place with a real Zotero field.

Setup (once, on top of `zotero_linker.py`'s requirements):
```bash
pip install pywin32
```
Word must already be open with your thesis and Zotero running with the local
API enabled, same as above.

```bash
python zotero_linker_selection.py --user-id 123456
python zotero_linker_selection.py --interactive     # ask when ambiguous
python zotero_linker_selection.py --dry-run          # report only, Word untouched
```
Citations it can't find are left selected as plain text, same as the batch
tool — nothing is guessed or silently mis-linked.

`link_selected_citation.bat` is a Windows shortcut for this; set the `ZOTERO_USER_ID`
environment variable first (it is used if present).

## Make citations clickable
For numbered styles (Nature, IEEE, Vancouver, ...) and author-year styles (APA, Harvard, ...), `zotero_hyperlinks.py` makes each
number jump to its entry in the reference list, in Word and in the exported PDF.
```bash
python zotero_hyperlinks.py paper.docx          # -> paper_linked.docx
python zotero_hyperlinks.py paper.docx --blue   # also show links blue/underlined
```
Zotero → Refresh drops these links, so run it as the **last** step, after the final
Refresh and just before exporting the PDF. Running it again is safe.

To double-check the result, `python verify_links.py paper_linked.docx` confirms every
link points at the matching reference-list entry.

Limitations of the linker: the match is strict about author count (`Lee, 2024` will not match a
two-author item), surnames with lowercase particles such as `van der Berg` are not recognised
in typed citations, `2020a`/`2020b` both resolve to the first matching item, and citations in
footnotes are not converted. Live mode refuses a selection that contains fields or hidden text
(positions would not line up) — select just the plain typed citation.

Limitations of the hyperlinker: only the document body is processed (citations inside footnotes, headers or
text boxes are left alone), and an ambiguous author-year match (e.g. two `Smith 2020a`
entries cited as `Smith 2020`) links to the first one.

## Tests
`python tests/test_hyperlinks.py` and `python tests/test_linker.py` build small synthetic documents
(the latter with a fake Zotero) and check that text stays intact, the right entries are linked
and re-running is safe. No Zotero or Word needed.

## Examples
`examples/` holds a small synthetic Word file with typed citations and its converted
counterpart, for trying the tools without touching your own documents.

## Requirements
Python 3.9+, Zotero 7 (desktop, with local API enabled) and Word. Live mode needs Windows.

## License
MIT — see [LICENSE](LICENSE).
