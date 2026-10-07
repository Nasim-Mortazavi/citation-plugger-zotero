#!/usr/bin/env python3
"""
zotero_hyperlinks.py — make Zotero citations clickable: numbered styles (Nature, IEEE,
Vancouver, ...) and author-year styles (APA, Harvard, ...). Each citation jumps to its
entry in the reference list, in Word and in the exported PDF.

    python zotero_hyperlinks.py paper.docx               # -> paper_linked.docx
    python zotero_hyperlinks.py paper.docx --blue        # also show links blue/underlined

The Zotero fields stay live. But Zotero → Refresh rewrites citations and the
bibliography and drops the links, so run this as the LAST step, after the final
Refresh and just before you export the PDF. Running it again on an
already-linked file is safe.
"""
from __future__ import annotations

import argparse
import copy
import re
import sys

from docx import Document
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

BM_PREFIX = "_ZoteroRef_"          # leading "_" = hidden bookmark (like Word's _Toc)
ENTRY_RE = re.compile(r"^\s*\[?(\d+)[.\]]")   # "12. \t..." or "[12] ..."


def parent_p(el):
    while el is not None and el.tag != qn("w:p"):
        el = el.getparent()
    return el


def strip_links(body) -> int:
    """Remove this script's earlier links and bookmarks so numbering is rebuilt from the
    CURRENT bibliography (a Zotero refresh can renumber entries and leave old links stale)."""
    n = 0
    for h in list(body.iter(qn("w:hyperlink"))):
        if (h.get(qn("w:anchor")) or "").startswith(BM_PREFIX):
            par = h.getparent()
            for child in list(h):
                h.addprevious(child)
            par.remove(h)
            n += 1
    for b in list(body.iter(qn("w:bookmarkStart"))):
        if (b.get(qn("w:name")) or "").startswith(BM_PREFIX):
            bid = b.get(qn("w:id"))
            for e in list(body.iter(qn("w:bookmarkEnd"))):
                if e.get(qn("w:id")) == bid:
                    e.getparent().remove(e)
            b.getparent().remove(b)
    return n


def scan(body):
    """Walk runs in document order, tracking nested fields. Returns
    (citation result runs, bibliography paragraphs)."""
    stack = []                     # [kind, code, in_result, group_id]
    cite_runs, bib_paras = [], []   # cite_runs: [(run, group_id)] - one group per citation field
    gid = 0
    for p in body.iter(qn("w:p")):
        in_bib = any(f[0] == "bib" and f[2] for f in stack)
        for r in p.iter(qn("w:r")):
            if parent_p(r) is not p:           # run of a nested text box paragraph
                continue
            fc = r.find(qn("w:fldChar"))
            it = r.find(qn("w:instrText"))
            if fc is not None:
                kind = fc.get(qn("w:fldCharType"))
                if kind == "begin":
                    gid += 1
                    stack.append(["?", "", False, gid])
                elif kind == "separate" and stack:
                    code = stack[-1][1]
                    stack[-1][0] = "cite" if "ZOTERO_ITEM" in code else "bib" if "ZOTERO_BIBL" in code else "other"
                    stack[-1][2] = True
                elif kind == "end" and stack:
                    stack.pop()
                continue
            if it is not None and stack:
                stack[-1][1] += it.text or ""
                continue
            if stack and stack[-1][2]:
                if stack[-1][0] == "cite" and r.getparent() is p:
                    cite_runs.append((r, stack[-1][3]))
                if stack[-1][0] == "bib":
                    in_bib = True
        if in_bib:
            bib_paras.append(p)
    return cite_runs, bib_paras


def add_bookmarks(body, bib_paras) -> set[int]:
    existing = {b.get(qn("w:name")) for b in body.iter(qn("w:bookmarkStart"))}
    ids = [int(b.get(qn("w:id"))) for b in body.iter(qn("w:bookmarkStart")) if b.get(qn("w:id"), "").isdigit()]
    nid = max(ids, default=0) + 1
    numbers = set()
    for p in bib_paras:
        text = "".join(t.text or "" for t in p.iter(qn("w:t")))
        m = ENTRY_RE.match(text)
        if not m:
            continue
        n = int(m.group(1))
        numbers.add(n)
        name = f"{BM_PREFIX}{n}"
        if name in existing:
            continue
        start = OxmlElement("w:bookmarkStart"); start.set(qn("w:id"), str(nid)); start.set(qn("w:name"), name)
        end = OxmlElement("w:bookmarkEnd"); end.set(qn("w:id"), str(nid))
        ppr = p.find(qn("w:pPr"))
        if ppr is not None:
            ppr.addnext(start)
        else:
            p.insert(0, start)
        p.append(end)
        existing.add(name)
        nid += 1
    return numbers


_AFTER_COLOR = ("spacing", "w", "kern", "position", "sz", "szCs", "highlight", "u", "effect", "bdr", "shd",
                "fitText", "vertAlign", "rtl", "cs", "em", "lang", "eastAsianLayout", "specVanish", "oMath")


def insert_rpr(rp, el, successors):
    """Insert el into rPr at its schema position (Word can reject out-of-order children)."""
    for child in rp:
        if child.tag in {qn(f"w:{n}") for n in successors}:
            child.addprevious(el)
            return
    rp.append(el)


def make_run(rpr, text, blue):
    r = OxmlElement("w:r")
    rp = copy.deepcopy(rpr) if rpr is not None else None
    if blue:
        rp = rp if rp is not None else OxmlElement("w:rPr")
        for tag in ("w:color", "w:u"):
            old = rp.find(qn(tag))
            if old is not None:
                rp.remove(old)
        c = OxmlElement("w:color"); c.set(qn("w:val"), "0563C1"); insert_rpr(rp, c, _AFTER_COLOR)
        u = OxmlElement("w:u"); u.set(qn("w:val"), "single"); insert_rpr(rp, u, _AFTER_COLOR[8:])
    if rp is not None:
        r.append(rp)
    t = OxmlElement("w:t"); t.set(qn("xml:space"), "preserve"); t.text = text
    r.append(t)
    return r


def link_citations(cite_runs, numbers, blue) -> tuple[int, list[int]]:
    linked, missing = 0, []
    for r in cite_runs:
        ts = r.findall(qn("w:t"))
        if not ts or len(r) - (r.find(qn("w:rPr")) is not None) != len(ts):
            continue                                   # tabs/breaks etc. — leave alone
        text = "".join(t.text or "" for t in ts)
        if not re.search(r"\d", text):
            continue
        rpr = r.find(qn("w:rPr"))
        new = []
        for piece in re.split(r"(\d+)", text):
            if not piece:
                continue
            if piece.isdigit() and int(piece) in numbers:
                h = OxmlElement("w:hyperlink")
                h.set(qn("w:anchor"), f"{BM_PREFIX}{int(piece)}")
                h.set(qn("w:history"), "1")
                h.append(make_run(rpr, piece, blue))
                new.append(h)
                linked += 1
            else:
                if piece.isdigit():
                    missing.append(int(piece))
                new.append(make_run(rpr, piece, False))
        for el in reversed(new):
            r.addnext(el)
        r.getparent().remove(r)
    return linked, missing


AY_YEAR_RE = re.compile(r"\(((?:19|20)\d\d[a-z]?)\)")        # bibliography: "(2020)" / "(2020a)"
CITE_YEAR_RE = re.compile(r"(?<!\d)((?:19|20)\d\d[a-z]?)(?![\d])")
LEAD_RE = re.compile(r"^\s*(?:e\.g\.,?|see|cf\.|also|e\.g\.)\s+", re.I)


def norm(s: str) -> str:
    return re.sub(r"[^\w]", "", s.casefold())


def ay_entries(body, bib_paras) -> dict[int, tuple[str, str]]:
    """Bookmark each author-year bibliography entry as _ZoteroRef_<n> (n = order).
    Returns {n: (normalised first-author surname, year)}."""
    existing = {b.get(qn("w:name")) for b in body.iter(qn("w:bookmarkStart"))}
    ids = [int(b.get(qn("w:id"))) for b in body.iter(qn("w:bookmarkStart")) if b.get(qn("w:id"), "").isdigit()]
    nid = max(ids, default=0) + 1
    entries, n = {}, 0
    for p in bib_paras:
        text = "".join(t.text or "" for t in p.iter(qn("w:t"))).strip()
        m = AY_YEAR_RE.search(text)
        if not m:
            continue
        n += 1
        entries[n] = (norm(text[:m.start()].split(",")[0]), m.group(1))
        name = f"{BM_PREFIX}{n}"
        if name in existing:
            continue
        start = OxmlElement("w:bookmarkStart"); start.set(qn("w:id"), str(nid)); start.set(qn("w:name"), name)
        end = OxmlElement("w:bookmarkEnd"); end.set(qn("w:id"), str(nid))
        ppr = p.find(qn("w:pPr"))
        if ppr is not None:
            ppr.addnext(start)
        else:
            p.insert(0, start)
        p.append(end)
        existing.add(name)
        nid += 1
    return entries


def ay_lookup(entries, name_text: str, year: str):
    name = LEAD_RE.sub("", name_text).strip().rstrip(",").strip()
    name = re.sub(r"^(?:[A-Z]\.\s*)+", "", name)             # "B. W. Corrigan" -> "Corrigan"
    name = re.split(r"\s+et\s+al\b|\s+&\s+|\s+and\s+|,", name, maxsplit=1)[0]
    key = norm(name)
    if not key:
        return None
    cands = [n for n, (s, y) in entries.items() if s == key and y == year]
    if not cands and year[-1].isdigit():                     # "2020" cites "2020a"
        cands = [n for n, (s, y) in entries.items() if s == key and y[:4] == year]
    if not cands:                                            # compound surnames ("Berg" vs "van der Berg")
        cands = [n for n, (s, y) in entries.items() if y == year and min(len(s), len(key)) >= 4
                 and (s.startswith(key) or key.startswith(s) or s.endswith(key) or key.endswith(s))]
        if len(cands) != 1:                                  # ambiguous: do not guess
            return None
    return cands[0] if cands else None


NARR_RE = re.compile(r"([A-Z][\w\-'’]+(?:\s+et\s+al\.?|\s+(?:&|and)\s+[A-Z][\w\-'’]+)?)\s*$")


def preceding_text(first_run) -> str:
    """Text of the paragraph before the first run of a citation (for narrative 'Liu et al. (2022)')."""
    p = parent_p(first_run)
    out = []
    for t in p.iter(qn("w:t")):
        if first_run in t.iterancestors():
            break
        out.append(t.text or "")
    return "".join(out)


def link_citations_ay(cite_runs, entries, blue) -> tuple[int, list[str]]:
    """Link per citation field, not per run: Zotero/Word often split one citation
    ('(Karlocai' | ' et al., 2014; ') across runs."""
    groups: dict[int, list] = {}
    for r, g in cite_runs:
        groups.setdefault(g, []).append(r)
    linked, missing = 0, []
    for runs in groups.values():
        items = []                                          # (run, text, offset)
        pos = 0
        for r in runs:
            ts = r.findall(qn("w:t"))
            if not ts or len(r) - (r.find(qn("w:rPr")) is not None) != len(ts):
                continue
            text = "".join(t.text or "" for t in ts)
            items.append((r, text, pos))
            pos += len(text)
        if not items:
            continue
        full = "".join(t for _, t, _ in items)
        spans = []                                          # (start, end, entry number)
        prev_end, last_name = 0, ""
        for m in CITE_YEAR_RE.finditer(full):
            seg = max(full.rfind(";", 0, m.start()), full.rfind("(", 0, m.start())) + 1
            same_author = seg < prev_end                    # "(Smith, 2020, 2021)": 2nd year, same author
            seg = max(seg, prev_end)
            seg_text = full[seg:m.start()]
            lead = len(seg_text) - len(seg_text.lstrip())
            name_text, start = seg_text, seg + lead
            prev_end = m.end()
            if same_author and not re.sub(r"[\s,]", "", seg_text):
                name_text, start = last_name, m.start()
            elif not seg_text.strip() and full[:m.start()].strip() in ("", "("):
                nm = NARR_RE.search(preceding_text(items[0][0]).rstrip())   # narrative: Name et al. (2022)
                if nm:
                    name_text, start = nm.group(1), m.start()
            if name_text.strip():
                last_name = name_text
            n = ay_lookup(entries, name_text, m.group(1))
            if n is None:
                if name_text.strip():
                    missing.append(f"{name_text.strip().rstrip(',')} {m.group(1)}")
                continue
            spans.append((start, m.end(), n))
        if not spans:
            continue
        for r, text, off in items:
            lo, hi = off, off + len(text)
            pieces, cur = [], 0                              # (text, entry number | None)
            for s0, e0, n in spans:
                a0, b0 = max(s0, lo) - lo, min(e0, hi) - lo
                if a0 >= b0:
                    continue
                if a0 > cur:
                    pieces.append((text[cur:a0], None))
                pieces.append((text[a0:b0], n))
                cur = b0
            if not pieces:
                continue
            if cur < len(text):
                pieces.append((text[cur:], None))
            rpr = r.find(qn("w:rPr"))
            new = []
            for t_, n in pieces:
                if n is None:
                    new.append(make_run(rpr, t_, False))
                else:
                    h = OxmlElement("w:hyperlink")
                    h.set(qn("w:anchor"), f"{BM_PREFIX}{n}")
                    h.set(qn("w:history"), "1")
                    h.append(make_run(rpr, t_, blue))
                    new.append(h)
                    linked += 1
            for el in reversed(new):
                r.addnext(el)
            r.getparent().remove(r)
    return linked, missing


def is_numbered(bib_paras) -> bool:
    for p in bib_paras:
        if ENTRY_RE.match("".join(t.text or "" for t in p.iter(qn("w:t")))):
            return True
    return False


def main():
    for stream in (sys.stdout, sys.stderr):
        if stream.encoding and stream.encoding.lower() != "utf-8":
            stream.reconfigure(encoding="utf-8", errors="replace")
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("docx")
    ap.add_argument("-o", "--output")
    ap.add_argument("--blue", action="store_true", help="colour the linked numbers blue + underline")
    a = ap.parse_args()
    out = a.output or re.sub(r"(_linked)?\.docx$", "", a.docx) + "_linked.docx"

    doc = Document(a.docx)
    body = doc.element.body
    strip_links(body)
    cite_runs, bib_paras = scan(body)
    if not bib_paras:
        sys.exit("No Zotero bibliography found — insert one (Zotero → Add/Edit Bibliography) and Refresh first.")
    if is_numbered(bib_paras):
        numbers = add_bookmarks(body, bib_paras)
        linked, missing = link_citations([r for r, _ in cite_runs], numbers, a.blue)
        count, what = len(numbers), "citation numbers"
        missing_msg = "Numbers with no matching reference entry (left unlinked): "
        missing = sorted(set(missing))
    else:
        entries = ay_entries(body, bib_paras)
        if not entries:
            sys.exit("Bibliography entries are neither numbered nor author-year (no '(YYYY)' found).")
        linked, missing = link_citations_ay(cite_runs, entries, a.blue)
        count, what = len(entries), "author-year citations"
        missing_msg = "Citations with no matching reference entry (left unlinked): "
        missing = sorted(set(missing))
    doc.save(out)

    print(f"{count} reference entries bookmarked, {linked} {what} linked"
          + ".")
    if missing:
        print(f"{missing_msg}{missing}")
    print(f"Saved → {out}\nNot linked? Zotero → Refresh drops the links; just re-run this afterwards.")


if __name__ == "__main__":
    main()
