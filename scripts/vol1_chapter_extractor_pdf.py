"""
vol1_chapter_extractor_pdf.py
Extract all items + SAY rates + resource rows for Ch.01, 03-12 directly from
CivilDAR_2019_Vol_1.pdf and save to data/reference_json/ch{NN}_items.json.
(Ch01 "Carriage of Materials" also contains large lead-vs-distance lookup
tables with no per-item "Say" line -- those are correctly skipped; only its
genuine priced cost-buildup items, e.g. 1.3, 1.4.1-1.4.3, are extracted.)

Replaces vol1_chapter_extractor.py's old path (Converted XLSX -> JSON), which
depended on data/converted_xlsx/CivilDAR_2019_Vol_1_Converted.xlsx: a file with
no in-repo provenance that was never verified against the source PDF (see
reports/audit_architecture_process_2026-09-17.md, finding D2). This mirrors
vol2_chapter_extractor.py's already-verified direct-PDF approach.

Usage:
    python scripts/vol1_chapter_extractor_pdf.py --chapter 5
    python scripts/vol1_chapter_extractor_pdf.py --all
"""

import sys, re, json, argparse
from pathlib import Path
sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

import pdfplumber
from scripts.paths import PDF_CIVIL_DAR_VOL1

# Chapter -> (pdf_start_page_0indexed, pdf_end_page_exclusive)
# Boundaries verified against each chapter's "SUB HEAD : N.0" page.
CHAPTER_PAGES = {
    1:  (73, 85),
    3:  (141, 151),
    4:  (151, 187),
    5:  (187, 279),
    6:  (279, 321),
    7:  (321, 365),
    8:  (365, 425),
    9:  (425, 673),
    10: (673, 719),
    11: (719, 811),
    12: (811, 904),
}

OUT_DIR = Path(__file__).resolve().parent.parent / "data" / "reference_json"

# One item-code segment: digits with an optional trailing letter (e.g. "22A",
# "48X"), repeated across dot levels (e.g. "5.22A.1", "9.147A", "12.52.1").
_CODE_SEG = r"\d+[A-Z]?"

# A line matching \d+\.\d+ that is actually a cross-reference back to another
# item's rate ("2.25 Rate as per item No.2.25 of ...") rather than a genuine
# new item header.
_CROSS_REF_DESC = re.compile(
    r"^\s*\(?\s*(rate\s+(?:same\s+)?as\s+per\s+item|sub\s+analysis\s+no\.?\s+\S+\s*\(annexure\))",
    re.I,
)

# Resource-row continuation-line boundary: a line starting one of these should
# never be merged into the previous resource row's wrapped description.
_BOUNDARY_PAT = re.compile(
    r"^(MATERIAL|LABOUR|MACHINERY|CARRIAGE|TOTAL|Add\b|Deduct\b|GRAND|SAY\b|Say\b|Code\b|Detail|"
    rf"{_CODE_SEG}(?:\.{_CODE_SEG})*\s+\(?\s*Rate\s+(?:same\s+)?as\s+per|"
    rf"{_CODE_SEG}(?:\.{_CODE_SEG})*\s+Sub\s+Analysis\s+no)",
    re.I,
)
_CODE_LINE_PAT = re.compile(r"^\d{4}\b")

# Per-page boilerplate that pdfplumber re-inserts into the middle of a
# wrapped description whenever it spans a page break; strip it out entirely
# rather than treating it as a section boundary.
_PAGE_BOILERPLATE_PAT = re.compile(
    r"^(SUB HEAD\s*:.*|Code\s+Description\s+Unit\s+Quantity\s+Rate\s+Amount)$",
    re.I,
)

_UNITS = (
    r"day|cum|kg|kilogram|cm|nos|litre|ltr|tonne|L\.S\.|Qtl|metre|m|sqm|Rmt|"
    r"no|set|pair|each|job|month|hr|test|bag|quintal|door\s+area|shutter\s+area"
)

# Standard 4-digit DSR resource row (code, desc, unit, qty, rate), applied to
# a single (already line-joined) entry.
_RES_PAT = re.compile(
    rf"^(\d{{4}})\s+(.+?)\s+((?:\d+\s+)?(?:{_UNITS}))\s+([\d\.]+)\s+([\d,\.]+)",
    re.I,
)

# Cross-references embedded in a cost buildup: "2.6.1 Rate as per item no 2.6.1 ... unit qty rate amount"
_REF_PAT = re.compile(
    rf"({_CODE_SEG}(?:\.{_CODE_SEG})*)\s+\(?\s*Rate\s+(?:same\s+)?as\s+per\s+[Ii]tem\s+(?:Number|No\.?)\s*.+?"
    rf"({_UNITS})\s+([\d\.]+)\s+([\d,\.]+)\s+([\d,\.]+)",
    re.S | re.I,
)

# Sub-analysis annexure references embedded in a cost buildup, e.g.
# "5.48X Sub Analysis no 5.48X (Annexure) for item 5.48.1 each 1.00 592482.20 592482.20"
_REF_PAT_SUBANALYSIS = re.compile(
    rf"({_CODE_SEG}(?:\.{_CODE_SEG})*)\s+Sub\s+Analysis\s+no\.?\s+.+?"
    rf"({_UNITS})\s+([\d\.]+)\s+([\d,\.]+)\s+([\d,\.]+)",
    re.S | re.I,
)


def _join_wrapped_lines(block: str) -> list[str]:
    """Join a resource row's wrapped description lines back onto its code line."""
    merged: list[str] = []
    for raw_line in block.splitlines():
        line = raw_line.strip()
        if not line or _PAGE_BOILERPLATE_PAT.match(line):
            continue
        if _CODE_LINE_PAT.match(line):
            merged.append(line)
        elif merged and not _BOUNDARY_PAT.match(line):
            merged[-1] = merged[-1] + " " + line
        else:
            merged.append(line)
    return merged


def _parse_basis(block: str) -> tuple[float, str]:
    m = re.search(r"Details?\s+of\s+costs?\s+(?:of|for)", block, re.I)
    if not m:
        return 1.0, "nos"
    window = block[m.end():m.end() + 150]
    m1 = re.search(r"^\s*([\d\.]+)\s*(\w+)", window)
    if m1:
        return float(m1.group(1)), m1.group(2)
    m2 = re.search(r"=\s*([\d\.]+)\s*(\w+)", window)
    if m2:
        return float(m2.group(1)), m2.group(2)
    return 1.0, "nos"


def extract_chapter(ch: int, pdf_path: Path) -> list:
    start, end = CHAPTER_PAGES[ch]
    with pdfplumber.open(str(pdf_path)) as pdf:
        pages = [pdf.pages[i].extract_text() or "" for i in range(start, end)]
    full = "\n".join(pages)
    full = "\n".join(
        ln for ln in full.splitlines() if not _PAGE_BOILERPLATE_PAT.match(ln.strip())
    )

    ch_s = str(ch)
    seg_pat = re.compile(rf"(?m)^({re.escape(ch_s)}\.{_CODE_SEG}(?:\.{_CODE_SEG})*)\s+(.+?)$")

    # A genuine item header's description usually starts with an upper-case
    # word (DAR phrasing convention: "Providing...", "Extra for...", etc.), and
    # a line that merely *starts* with a chapter-shaped number by coincidence
    # (a wrapped dimension calculation like "7.00 x 24 cm x24 cm...") is always
    # lower-case and must be rejected, or it corrupts real item boundaries
    # around it. But a handful of genuine items ARE lower-case in the source
    # (e.g. "6.12.2 cement mortar 1:4...") - for those, require the item's own
    # "Details of cost..." line to appear shortly after, which a stray
    # mid-sentence fragment never has.
    def _accept(m: re.Match) -> bool:
        if _CROSS_REF_DESC.match(m.group(2)):
            return False
        if not m.group(2).lstrip("( ")[:1].islower():
            return True
        lookahead = full[m.end():m.end() + 200]
        return bool(re.search(r"Details?\s+of\s+costs?\s+(?:of|for)", lookahead, re.I))

    segments = [m for m in seg_pat.finditer(full) if _accept(m)]

    results = []
    for idx, m in enumerate(segments):
        code = m.group(1)
        desc_first = m.group(2).strip()
        start_pos = m.end()
        end_pos = segments[idx + 1].start() if idx + 1 < len(segments) else len(full)
        block = full[start_pos:end_pos]

        say_ms = list(re.finditer(r"\bSay\s+([\d,]+\.?\d*)", block, re.I))
        if not say_ms:
            continue
        say = float(say_ms[-1].group(1).replace(",", ""))

        basis, unit = _parse_basis(block)

        resources = []

        for rm in _REF_PAT.finditer(block):
            resources.append({
                "code": rm.group(1),
                "desc": f"Rate as per SH: {rm.group(1)}",
                "unit": rm.group(2),
                "qty": float(rm.group(3)),
                "rate": float(rm.group(4).replace(",", "")),
            })

        for rm in _REF_PAT_SUBANALYSIS.finditer(block):
            resources.append({
                "code": rm.group(1),
                "desc": f"Sub-analysis annexure: {rm.group(1)}",
                "unit": rm.group(2),
                "qty": float(rm.group(3)),
                "rate": float(rm.group(4).replace(",", "")),
            })

        for entry in _join_wrapped_lines(block):
            rm = _RES_PAT.match(entry)
            if rm:
                resources.append({
                    "code": rm.group(1),
                    "desc": rm.group(2).strip(),
                    "unit": rm.group(3),
                    "qty": float(rm.group(4)),
                    "rate": float(rm.group(5).replace(",", "")),
                })

        desc_extra = []
        for ln in block.splitlines()[:5]:
            ln = ln.strip()
            if re.match(r"^(Code|MATERIAL|LABOUR|Detail|\d{4})", ln) or _CROSS_REF_DESC.match(ln):
                break
            if ln and ln not in (desc_first,):
                desc_extra.append(ln)
        full_desc = (desc_first + " " + " ".join(desc_extra)).strip()

        results.append({
            "code": code,
            "desc": full_desc[:140],
            "unit": unit,
            "basis": basis,
            "say": say,
            "resources": resources,
        })

    return results


def save(ch: int, items: list):
    out = OUT_DIR / f"ch{ch:02d}_items.json"
    out.write_text(json.dumps(items, indent=2, ensure_ascii=False), encoding="utf-8")
    return out


def main(argv=None):
    ap = argparse.ArgumentParser()
    ap.add_argument("--chapter", type=int, help="Single chapter number (1, 3-12)")
    ap.add_argument("--all", action="store_true", help="Extract all chapters (1, 3-12)")
    args = ap.parse_args(argv)

    pdf = Path(PDF_CIVIL_DAR_VOL1)
    chapters = ([1] + list(range(3, 13))) if args.all else ([args.chapter] if args.chapter else [])
    if not chapters:
        ap.print_help(); sys.exit(1)

    for ch in chapters:
        if ch not in CHAPTER_PAGES:
            print(f"  Ch {ch}: not in page map, skipping")
            continue
        print(f"Extracting Ch.{ch} …", end=" ", flush=True)
        items = extract_chapter(ch, pdf)
        zero_resource = [it["code"] for it in items if not it["resources"]]
        if zero_resource:
            print(f"\n  [WARN] Ch.{ch}: {len(zero_resource)} item(s) with zero resources: {zero_resource}")
        out = save(ch, items)
        print(f"{len(items)} items -> {out.name}")


if __name__ == "__main__":
    main()
