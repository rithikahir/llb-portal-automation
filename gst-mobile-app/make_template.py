#!/usr/bin/env python3
"""
make_template.py — build GST_BILL_TEMPLATE.docx from a real bill.

Reads a bill .docx (default: Shreeganesh_Telecom_GST_BILLS\\HINDUSTAN_BILL.docx)
and replaces every per-customer value with a {{TOKEN}} placeholder, keeping the
exact layout, fonts and tables. Upload the result to Google Drive, open it with
Google Docs, and set its ID with:

    python test_backend.py template <docId>

The mobile app then generates every bill PDF from this template — identical to
the Word bill. READ-ONLY on the source bill (the template is a new file here).
"""
from __future__ import annotations

import sys
from pathlib import Path

import docx  # python-docx

HERE = Path(__file__).resolve().parent
DEFAULT_SRC = HERE.parent / "Shreeganesh_Telecom_GST_BILLS" / "HINDUSTAN_BILL.docx"
OUT = HERE / "GST_BILL_TEMPLATE.docx"

# (old, new) in application order — buyer BEFORE user id (buyer contains it),
# and '00.00' (round-off) LAST because it is a substring of the amounts.
REPLACEMENTS = [
    ("HINDUSTAN CHEAP STORES.", "{{BUYER}}"),
    ("95, Mg Road, Pune - 411011.", "{{ADDRESS}}"),
    ("27AAAFH9069C1ZE", "{{GSTIN}}"),
    ("SEVEN THOUSAND EIGHTY RUPEES ONLY/_", "{{WORDS}}"),
    ("28/11/2026 to 27/11/2027", "{{PERIOD}}"),
    ("ST/053/2026-27", "{{INVOICE_NO}}"),
    ("14-11-26", "{{INVOICE_DATE}}"),
    ("HINDUSTAN", "{{USERID}}"),          # after BUYER (buyer contains it)
    ("HIGH SPEED 50 MBPS", "{{DESC1}}"),
    ("TRULY UNLIMITED 12 MONTHS PLAN.", "{{DESC2}}"),
    ("540.00", "{{CGST}}"),               # CGST and SGST cells share the value
    ("7080.00", "{{TOTAL}}"),
    ("6000.00", "{{AMOUNT}}"),
    ("00.00", "{{ROUND}}"),               # round-off — must be LAST
]


def iter_paragraphs(doc):
    """Every paragraph: body, tables (incl. nested), headers, footers."""
    def from_table(t):
        for row in t.rows:
            seen = set()
            for cell in row.cells:            # merged cells repeat — dedupe
                if id(cell._tc) in seen:
                    continue
                seen.add(id(cell._tc))
                for p in cell.paragraphs:
                    yield p
                for nt in cell.tables:
                    yield from from_table(nt)

    for p in doc.paragraphs:
        yield p
    for t in doc.tables:
        yield from from_table(t)
    for s in doc.sections:
        for hf in (s.header, s.footer, s.first_page_header, s.first_page_footer,
                   s.even_page_header, s.even_page_footer):
            for p in hf.paragraphs:
                yield p
            for t in hf.tables:
                yield from from_table(t)


def replace_in_paragraph(para, old: str, new: str) -> int:
    """Replace `old` with `new` even when it spans several runs (Word splits
    text into runs arbitrarily); keeps each run's formatting."""
    hits = 0
    while True:
        runs = para.runs
        text = "".join(r.text for r in runs)
        idx = text.find(old)
        if idx < 0:
            return hits
        end = idx + len(old)
        pos = 0
        spans = []
        for r in runs:
            spans.append((pos, pos + len(r.text), r))
            pos += len(r.text)
        start = next((s for s in spans if s[0] <= idx < s[1]), spans[-1])
        stop = next((s for s in spans if s[0] < end <= s[1]), spans[-1])
        sr, er = start[2], stop[2]
        pre = sr.text[: idx - start[0]]
        post = er.text[end - stop[0]:]
        if sr is er:
            sr.text = pre + new + post
        else:
            sr.text = pre + new
            for s in spans:
                r = s[2]
                if r is sr or r is er:
                    continue
                if s[0] >= start[0] and s[1] <= stop[1]:   # fully inside the match
                    r.text = ""
            er.text = post
        hits += 1


def inline_anchors(path) -> int:
    """Convert floating (wp:anchor) images to inline (wp:inline) drawings.

    Google Docs mangles floating drawings during docx import (they become
    broken, invisible elements) and its PDF exporter then crashes with a
    server error. Inline images convert and export cleanly. Returns the
    number of drawings converted.
    """
    import shutil
    import zipfile

    from lxml import etree

    WP = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
    A = "http://schemas.openxmlformats.org/drawingml/2006/main"
    src = str(path)
    tmp = src + ".tmp"
    converted = 0

    with zipfile.ZipFile(src) as zin:
        xml = zin.read("word/document.xml")
        root = etree.fromstring(xml)
        for anchor in list(root.iter("{%s}anchor" % WP)):
            inline = etree.Element("{%s}inline" % WP)
            for attr in ("distT", "distB", "distL", "distR"):
                inline.set(attr, anchor.get(attr, "0"))
            for ns, tag in ((WP, "extent"), (WP, "docPr"), (WP, "cNvGraphicFramePr"),
                            (A, "graphic")):
                el = anchor.find("{%s}%s" % (ns, tag))
                if el is not None:
                    inline.append(el)          # moves the element out of anchor
            anchor.getparent().replace(anchor, inline)
            converted += 1
        new_xml = etree.tostring(root, xml_declaration=True,
                                 encoding="UTF-8", standalone=True)

    with zipfile.ZipFile(src) as zin, \
         zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            data = new_xml if item.filename == "word/document.xml" else zin.read(item.filename)
            zout.writestr(item, data)
    shutil.move(tmp, src)
    return converted


def main() -> None:
    src = Path(sys.argv[1]) if len(sys.argv) > 1 else DEFAULT_SRC
    if not src.exists():
        sys.exit(f"source bill not found: {src}")
    doc = docx.Document(str(src))
    report = []
    for old, new in REPLACEMENTS:
        n = 0
        for p in iter_paragraphs(doc):
            n += replace_in_paragraph(p, old, new)
        report.append((old, new, n))
        if n == 0:
            print(f"  WARNING: '{old}' not found — check the source bill", file=sys.stderr)
    doc.save(str(OUT))

    converted = inline_anchors(OUT)

    # verify: no original values left, tokens present
    check = docx.Document(str(OUT))
    all_text = "\n".join(p.text for p in iter_paragraphs(check))
    print(f"template written: {OUT}  (floating images converted to inline: {converted})")
    for old, new, n in report:
        print(f"  {n:>2}x  {old[:40]:42} -> {new}")
    leftover = [t for t, _, _ in report if t in all_text]
    print("leftover original values:", leftover or "none - OK")


if __name__ == "__main__":
    main()
