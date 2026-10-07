#!/usr/bin/env python3
"""verify_links.py - check that every author-year hyperlink points at the matching bibliography entry.
    python verify_links.py paper_linked.docx
Compares link text (first author, year, and second author if given) with the bookmarked entry."""
import re, sys
sys.stdout.reconfigure(encoding="utf-8", errors="replace")
from docx import Document
from docx.oxml.ns import qn
from zotero_hyperlinks import BM_PREFIX, ENTRY_RE, scan, ay_entries, norm

if len(sys.argv) != 2:
    sys.exit("usage: python verify_links.py paper_linked.docx")
doc = Document(sys.argv[1]); body = doc.element.body
_, bib = scan(body)
ay_entries(body, bib)                                   # ensures bookmarks exist (no-op if present)
entry = {}
for p in bib:
    for b in p.iter(qn("w:bookmarkStart")):
        if b.get(qn("w:name"), "").startswith(BM_PREFIX):
            entry[b.get(qn("w:name"))] = "".join(t.text or "" for t in p.iter(qn("w:t"))).strip()
bad = ok = 0
links = []                                              # merge adjacent links with the same anchor (one split citation)
for h in body.iter(qn("w:hyperlink")):
    a = h.get(qn("w:anchor"), "")
    if not a.startswith(BM_PREFIX): continue
    t_ = "".join(t.text or "" for t in h.iter(qn("w:t")))
    prev = h.getprevious()
    if links and prev is not None and prev is links[-1][2] and links[-1][0] == a:
        links[-1][1] += t_; links[-1][2] = h
    else:
        links.append([a, t_, h])
for a, txt, _h in links:
    txt = txt.strip(" ();")
    ref = entry.get(a)
    em = ENTRY_RE.match(ref) if ref is not None else None
    if em:                                                  # numbered style: number must match the entry's
        if em.group(1) == txt: ok += 1
        else: bad += 1; print("MISMATCH", ["number"], "| link:", repr(txt), "| entry:", ref[:90])
        continue
    m = re.search(r"((?:19|20)\d\d)[a-z]?", txt)
    if ref is None or not m:
        print("UNCHECKABLE", repr(txt), a); bad += 1; continue
    name = re.sub(r"^(?:[A-Z]\.\s*)+", "", txt[:m.start()].strip(" ,"))
    names = [n for n in re.split(r"\s+et\s+al\.?|\s+&\s+|\s+and\s+|,", name) if n.strip()]
    first = norm(names[0]) if names else ""
    ym = re.search(r"\(((?:19|20)\d\d)[a-z]?\)", ref)
    probs = []
    if not ym or ym.group(1) != m.group(1): probs.append("year")
    if first and not norm(ref.split(",")[0]).startswith(first): probs.append("first author")
    if len(names) > 1 and not re.search(r"\bet\s+al", txt) and norm(names[1]) not in norm(ref[:max(ref.find("("), 0)]): probs.append("second author")
    if probs:
        bad += 1; print("MISMATCH", probs, "| link:", repr(txt), "| entry:", ref[:90])
    else: ok += 1
print(f"{ok} links OK, {bad} problems, {len(entry)} bookmarked entries")
