"""
vol2_chapter_extractor.py
Extract all items + SAY rates + resource rows from a chapter of
CivilDAR_2019_Vol_2.pdf and save to data/reference_json/ch{NN}_items.json.

Usage:
    python scripts/vol2_chapter_extractor.py --chapter 14
    python scripts/vol2_chapter_extractor.py --all
"""

import sys, re, json, argparse
from pathlib import Path
sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

import pdfplumber
from scripts.paths import PDF_CIVIL_DAR_VOL2

# Chapter -> (pdf_start_page_0indexed, pdf_end_page_exclusive)
CHAPTER_PAGES = {
    13: (9,   109),
    14: (109, 179),
    15: (179, 227),
    16: (227, 371),
    17: (371, 485),
    18: (485, 753),
    19: (753, 855),
    20: (855, 887),
    21: (887, 919),
    22: (919, 945),
    23: (945, 971),
    24: (971, 979),
    25: (979, 991),
    26: (991, 1110),
}

OUT_DIR = Path(__file__).resolve().parent.parent / "data" / "reference_json"

# A line matching \d+\.\d+ that is actually a cross-reference back to another
# item's rate ("18.24 Rate as per Item Number 18.24 Of SH: ...") rather than a
# genuine new item header.
_CROSS_REF_DESC = re.compile(r"^\s*\(?\s*rate\s+(?:same\s+)?as\s+per\s+item", re.I)

# Resource-row continuation-line boundary: a line starting one of these should
# never be merged into the previous resource row's wrapped description.
_BOUNDARY_PAT = re.compile(
    r"^(MATERIAL|LABOUR|MACHINERY|TOTAL|Add\b|Deduct\b|GRAND|SAY\b|Say\b|Code\b|Detail|\d+(?:\.\d+)+\s+\(?\s*Rate\s+(?:same\s+)?as\s+per)",
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

# Standard 4-digit DSR resource row (code, desc, unit, qty, rate), applied to a
# single (already line-joined) entry.
# Shared by both resource-row patterns below - keep this the single source of
# unit vocabulary so the two patterns can never drift apart again (a missing
# "metre" here previously caused _REF_PAT to silently drop cross-reference
# resource rows using the spelled-out unit).
_UNITS = (
    r"day|cum|kg|kilogram|cm|nos|litre|ltr|tonne|L\.S\.|Qtl|quintal|metre|m|sqm|Rmt|"
    r"no|set|pair|each|job|month|hr|test|bag|door\s+area|shutter\s+area"
)

_RES_PAT = re.compile(
    rf"^(\d{{4}})\s+(.+?)\s+"
    rf"((?:\d+\s+)?(?:{_UNITS}))\s+"
    r"([\d\.]+)\s+([\d,\.]+)",
    re.I,
)

# Cross-references embedded in a cost buildup: "3.4 Rate as per Item Number ... unit qty rate amount"
_REF_PAT = re.compile(
    rf"(\d+(?:\.\d+)+[A-Z]?)\s+\(?\s*Rate\s+(?:same\s+)?as\s+per\s+[Ii]tem\s+(?:Number|No\.?)\s*.+?"
    rf"({_UNITS})\s+"
    r"([\d\.]+)\s+([\d,\.]+)\s+([\d,\.]+)",
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


_NUMBER_WORDS = {
    "a": 1.0, "an": 1.0, "one": 1.0, "two": 2.0, "three": 3.0, "four": 4.0,
    "five": 5.0, "six": 6.0, "seven": 7.0, "eight": 8.0, "nine": 9.0, "ten": 10.0,
}


def _parse_basis(block: str) -> tuple[float, str]:
    # Accepts "Detail(s) of cost(s) for/of ..." and the bare "Detail cost of ..."
    # variant (no "of" between "Detail" and "cost"); the trailing for/of is
    # optional too ("Details of cost 10sqm ...").
    m = re.search(r"Details?\s+(?:of\s+)?costs?\s+(?:for|of)?\s*", block, re.I)
    if not m:
        return 1.0, "nos"
    window = block[m.end():m.end() + 150]
    head = window[:100]

    m1 = re.match(r"\s*([\d.]+)\s*([A-Za-z]+)", head)
    if m1:
        # "10m x 10m = 100 sqm": a leading number+unit immediately followed by
        # "x <factor>" is one operand of a multiplication, not the basis
        # itself - the true basis is the product printed after "=".
        after = head[m1.end():m1.end() + 6]
        if re.match(r"\s*[xX]\s", after):
            m_eq = re.search(r"=\s*([\d.]+)\s*([A-Za-z]+)", head)
            if m_eq:
                return float(m_eq.group(1)), m_eq.group(2)
        return float(m1.group(1)), m1.group(2)

    m_word = re.match(
        r"\s*(a|an|one|two|three|four|five|six|seven|eight|nine|ten)\b\s*([A-Za-z]+)",
        head, re.I,
    )
    if m_word:
        return _NUMBER_WORDS[m_word.group(1).lower()], m_word.group(2)

    m2 = re.search(r"=\s*([\d.]+)\s*([A-Za-z]+)", window)
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
    seg_pat = re.compile(rf"(?m)^({re.escape(ch_s)}\.\d+(?:\.\d+)*)\s+(.+?)$")

    segments = [
        m for m in seg_pat.finditer(full)
        if not _CROSS_REF_DESC.match(m.group(2))
    ]

    results = []
    for idx, m in enumerate(segments):
        code = m.group(1)
        desc_first = m.group(2).strip()
        start_pos = m.end()
        end_pos = segments[idx + 1].start() if idx + 1 < len(segments) else len(full)
        block = full[start_pos:end_pos]

        # Use the LAST "Say <amount>" in the block, not the first: an item's
        # cost buildup can print an intermediate quantity-derivation annotation
        # phrased as "= 0.43 cum Say 0.43 cum" well before the true final-rate
        # "Say" line at the end of the block.
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

        # Continuation line: description spanned previous page (code already captured)
        # Build full desc from lines before MATERIAL/LABOUR/Code headers
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

    # Some items whose cost table spans a page break get their header+intro
    # reprinted verbatim on the next page for readability; that produces a
    # second, byte-identical entry for the same code. Collapse those, but
    # keep any same-code entries whose content actually differs (that would
    # be a real problem worth surfacing, not silently hiding).
    deduped: list = []
    seen: dict = {}
    for item in results:
        code = item["code"]
        if code in seen and item == seen[code]:
            continue
        seen[code] = item
        deduped.append(item)

    return deduped


def save(ch: int, items: list):
    out = OUT_DIR / f"ch{ch:02d}_items.json"
    out.write_text(json.dumps(items, indent=2, ensure_ascii=False), encoding="utf-8")
    return out


def main(argv=None):
    ap = argparse.ArgumentParser()
    ap.add_argument("--chapter", type=int, help="Single chapter number (13-26)")
    ap.add_argument("--all", action="store_true", help="Extract all chapters 13-26")
    args = ap.parse_args(argv)

    pdf = Path(PDF_CIVIL_DAR_VOL2)
    chapters = list(range(13, 27)) if args.all else ([args.chapter] if args.chapter else [])
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
