#!/usr/bin/env python3
"""
zotero_linker_selection.py — plug a live Zotero citation into whatever text
you currently have SELECTED in the open Word document (no file path needed).

Select a typed citation in Word — e.g. (Nasim et al., 2026), a narrative
Nasim et al. (2026), or a whole passage/paragraph containing several — then
run this script. It looks each one up in Zotero (exactly like
zotero_linker.py does for a whole file) and replaces it in place, live, in
the open document. No save/reopen needed.

Setup (once, in addition to zotero_linker.py's requirements)
    pip install pywin32
Word must already be open with your thesis, with some text selected.
Zotero 7 must be running with the local API enabled (Settings -> Advanced).

Usage
    python zotero_linker_selection.py                       # convert current selection
    python zotero_linker_selection.py --user-id 123456
    python zotero_linker_selection.py --interactive          # ask when ambiguous
    python zotero_linker_selection.py --dry-run              # report only, don't touch Word

Tip: bind this to a keyboard shortcut (e.g. an AutoHotkey script that runs
`python zotero_linker_selection.py`) to convert citations as you write,
without leaving Word.
"""
from __future__ import annotations

import argparse
import os
import sys
from xml.sax.saxutils import escape

try:
    import win32com.client as win32
except ImportError:
    sys.exit("This script needs pywin32: pip install pywin32")

from zotero_linker import ZoteroClient, Resolver, find_events, non_overlapping

WD_SELECTION_IP = 1        # Word's "just a blinking cursor, nothing selected" constant

# Word's Range.InsertXML wants the legacy flat WordML schema (the same one
# Range.XML emits) — a single <w:p> containing our field runs, which Word
# splices inline into the surrounding paragraph when the target Range sits
# inside one (verified: no stray paragraph break, no leftover text).
FIELD_XML_TEMPLATE = (
    '<?xml version="1.0" standalone="yes"?>'
    '<w:wordDocument xmlns:w="http://schemas.microsoft.com/office/word/2003/wordml" xml:space="preserve">'
    '<w:body><w:p>'
    '<w:r><w:fldChar w:fldCharType="begin"/></w:r>'
    '<w:r><w:instrText xml:space="preserve">{code}</w:instrText></w:r>'
    '<w:r><w:fldChar w:fldCharType="separate"/></w:r>'
    '<w:r><w:t xml:space="preserve">{display}</w:t></w:r>'
    '<w:r><w:fldChar w:fldCharType="end"/></w:r>'
    '</w:p></w:body></w:wordDocument>'
)


def get_word():
    try:
        return win32.GetActiveObject("Word.Application")
    except Exception:
        sys.exit("Could not find a running Word instance. Open Word with your thesis first.")


def insert_field(doc, start: int, end: int, code: str, display: str):
    rng = doc.Range(Start=start, End=end)
    xml = FIELD_XML_TEMPLATE.format(code=escape(f" {code} "), display=escape(display))
    rng.InsertXML(xml)


def main():
    for stream in (sys.stdout, sys.stderr):
        if stream.encoding and stream.encoding.lower() != "utf-8":
            stream.reconfigure(encoding="utf-8", errors="replace")
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--interactive", action="store_true", help="ask when several items could match")
    ap.add_argument("--dry-run", action="store_true")
    ap.add_argument("--web", action="store_true", help="use api.zotero.org instead of the desktop app")
    ap.add_argument("--user-id", default=os.getenv("ZOTERO_USER_ID"), help="numeric zotero.org user ID (recommended)")
    ap.add_argument("--group-id", default=os.getenv("ZOTERO_GROUP_ID"))
    ap.add_argument("--api-key", default=os.getenv("ZOTERO_API_KEY"))
    ap.add_argument("--port", type=int, default=23119)
    a = ap.parse_args()

    word = get_word()
    try:
        doc = word.ActiveDocument
        sel = word.Selection
    except Exception:
        sys.exit("No active document in Word.")

    if sel.Type == WD_SELECTION_IP or not sel.Text.strip():
        sys.exit("Nothing selected — select a citation (or a passage containing several) in Word first.")

    text = sel.Text
    sel_start = sel.Start

    client = ZoteroClient(a.web, a.user_id, a.group_id, a.api_key, a.port)
    resolver = Resolver(client, a.interactive)
    stats = {"fields": 0, "partial": 0}
    events = non_overlapping(find_events(text, resolver, stats))

    print("Citations:\n" + ("\n".join(resolver.log) if resolver.log else "  (none recognised in the selection)"))

    if not events:
        print("\nNothing to insert — selection left unchanged.")
        return

    if a.dry_run:
        print(f"\n--dry-run: {len(events)} citation(s) would be inserted, Word left unchanged.")
        return

    # Offsets in sel.Text only equal document positions for plain text. Fields (including
    # existing Zotero citations), objects, deleted tracked text, etc. take up positions that
    # Text does not show, which would make us replace the wrong characters. Check every
    # match against the document BEFORE changing anything.
    for s, e, code, disp in events:
        if doc.Range(Start=sel_start + s, End=sel_start + e).Text != text[s:e]:
            sys.exit("The selection contains fields, objects or hidden text, so positions don't line up "
                     "and nothing was changed. Select just the plain-text citation(s) and try again.")

    # Insert fields back-to-front so earlier offsets stay valid as the
    # document's character count shifts with each replacement.
    for s, e, code, disp in reversed(events):
        insert_field(doc, sel_start + s, sel_start + e, code, disp)

    print(f"\n{len(events)} citation(s) inserted in the selection. Zotero -> Refresh to format.")


if __name__ == "__main__":
    main()
