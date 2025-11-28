#!/usr/bin/env python3
import sys
import re
from pathlib import Path
import fitz  # PyMuPDF


def dump_page_spans(pdf_path: Path, page_idx: int) -> None:
    doc = fitz.open(str(pdf_path))
    if page_idx < 0 or page_idx >= len(doc):
        print(f"Page index out of range: {page_idx+1}/{len(doc)}")
        return
    page = doc[page_idx]
    blocks = page.get_text("dict")["blocks"]
    print(f"=== PAGE {page_idx+1} of {len(doc)} ===")
    for b in blocks:
        for l in b.get("lines", []):
            for s in l.get("spans", []):
                text = s.get("text", "")
                # Print all spans; caller can grep. Highlight spans that include replacement char.
                warn = " (HAS_U+FFFD)" if "\uFFFD" in text else ""
                ords = " ".join([hex(ord(c)) for c in text])
                print(f"FONT:{s.get('font')} SIZE:{s.get('size')}{warn}")
                print(f"RAW: {repr(text)}")
                print(f"ORDS: {ords}")
                print("-" * 60)


def main() -> None:
    if len(sys.argv) < 3:
        print("Usage: inspect_glyphs.py <pdf_path> <page_numbers_comma_sep>")
        print("Example: inspect_glyphs.py Payment.pdf 2,9,10,13")
        sys.exit(1)
    pdf_path = Path(sys.argv[1]).expanduser().resolve()
    pages_arg = sys.argv[2]
    page_numbers = []
    for tok in pages_arg.split(","):
        tok = tok.strip()
        if not tok:
            continue
        try:
            # 1-based page numbers
            page_numbers.append(int(tok))
        except ValueError:
            print(f"Invalid page number: {tok}")
            sys.exit(2)
    for pno in page_numbers:
        dump_page_spans(pdf_path, pno - 1)


if __name__ == "__main__":
    main()



