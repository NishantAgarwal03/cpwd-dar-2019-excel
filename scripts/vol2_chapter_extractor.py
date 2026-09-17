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


def extract_chapter(ch: int, pdf_path: Path) -> list:
    start, end = CHAPTER_PAGES[ch]
    with pdfplumber.open(str(pdf_path)) as pdf:
        pages = [pdf.pages[i].extract_text() or "" for i in range(start, end)]
    full = "\n".join(pages)

    ch_s = str(ch)
    seg_pat = re.compile(rf"(?m)^({re.escape(ch_s)}\.\d+(?:\.\d+)?)\s+(.+?)$")
    segments = list(seg_pat.finditer(full))

    results = []
    for idx, m in enumerate(segments):
        code = m.group(1)
        desc_first = m.group(2).strip()
        start_pos = m.end()
        end_pos = segments[idx + 1].start() if idx + 1 < len(segments) else len(full)
        block = full[start_pos:end_pos]

        say_m = re.search(r"\bSay\s+([\d,]+\.?\d*)", block)
        if not say_m:
            continue
        say = float(say_m.group(1).replace(",", ""))

        unit_m = re.search(r"Detail of cost for ([\d\.]+)\s+(\w+)", block)
        basis = float(unit_m.group(1)) if unit_m else 1.0
        unit = unit_m.group(2) if unit_m else "nos"

        resources = []

        # Cross-references: "3.4 Rate as per Item Number ... unit qty rate amount"
        ref_pat = re.compile(
            r"(\d+\.\d+[A-Z]?)\s+Rate as per Item Number\s+.+?"
            r"(cum|kg|m|nos|tonne|litre|ltr|Qtl|Rmt|sqm|L\.S\.|no|set|pair|each)\s+"
            r"([\d\.]+)\s+([\d,\.]+)\s+([\d,\.]+)",
            re.S,
        )
        for rm in ref_pat.finditer(block):
            resources.append({
                "code": rm.group(1),
                "desc": f"Rate as per SH: {rm.group(1)}",
                "unit": rm.group(2),
                "qty": float(rm.group(3)),
                "rate": float(rm.group(4).replace(",", "")),
            })

        # Standard 4-digit DSR codes
        res_pat = re.compile(
            r"^(0\d{3}|[89]\d{3})\s+(.+?)\s+"
            r"(day|cum|kg|nos|litre|ltr|tonne|L\.S\.|Qtl|metre|m|sqm|Rmt|no|set|pair|each|job|month|hr)\s+"
            r"([\d\.]+)\s+([\d,\.]+)",
            re.M,
        )
        for rm in res_pat.finditer(block):
            resources.append({
                "code": rm.group(1),
                "desc": rm.group(2)[:60].strip(),
                "unit": rm.group(3),
                "qty": float(rm.group(4)),
                "rate": float(rm.group(5).replace(",", "")),
            })

        # Continuation line: description spanned previous page (code already captured)
        # Build full desc from lines before MATERIAL/LABOUR/Code headers
        desc_extra = []
        for ln in block.splitlines()[:5]:
            ln = ln.strip()
            if re.match(r"^(Code|MATERIAL|LABOUR|Detail|0\d{3}|[89]\d{3})", ln):
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


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--chapter", type=int, help="Single chapter number (14-26)")
    ap.add_argument("--all", action="store_true", help="Extract all chapters 14-26")
    args = ap.parse_args()

    pdf = Path(PDF_CIVIL_DAR_VOL2)
    chapters = list(range(14, 27)) if args.all else ([args.chapter] if args.chapter else [])
    if not chapters:
        ap.print_help(); sys.exit(1)

    for ch in chapters:
        if ch not in CHAPTER_PAGES:
            print(f"  Ch {ch}: not in page map, skipping")
            continue
        print(f"Extracting Ch.{ch} …", end=" ", flush=True)
        items = extract_chapter(ch, pdf)
        out = save(ch, items)
        print(f"{len(items)} items -> {out.name}")


if __name__ == "__main__":
    main()
