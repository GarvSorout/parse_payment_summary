#!/usr/bin/env python3
from __future__ import annotations

import argparse
import re
from pathlib import Path
from typing import List

from pdf2image import convert_from_path
import pytesseract


def clean_ocr_text(raw: str) -> str:
	"""
	Clean Tesseract OCR output to something nicer:
	- remove control characters (except newline and tab)
	- normalize line endings to '\n'
	- strip trailing spaces per line
	- collapse multiple blank lines
	"""
	if raw is None:
		return ""
	text = raw
	# Remove ASCII control chars except newline (0x0A) and tab (0x09)
	text = re.sub(r"[\x00-\x08\x0b-\x0c\x0e-\x1f]", "", text)
	# Normalize Windows/Mac line endings
	text = text.replace("\r\n", "\n").replace("\r", "\n")
	# Strip trailing spaces
	lines = [line.rstrip() for line in text.split("\n")]
	# Collapse multiple blank lines to a single blank
	collapsed: List[str] = []
	blank = False
	for line in lines:
		if line.strip() == "":
			if not blank:
				collapsed.append("")
			blank = True
		else:
			collapsed.append(line)
			blank = False
	return "\n".join(collapsed).strip() + "\n"


def ocr_pdf_pages(pdf_path: Path, num_pages: int | None, tess_config: str) -> List[str]:
	"""
	Render PDF pages to images and OCR with Tesseract.
	Returns a list of cleaned text strings per page.
	"""
	# Convert pages; default 300 DPI is usually fine for text
	images = convert_from_path(str(pdf_path), dpi=300, fmt="png", first_page=1, last_page=None)
	if num_pages is not None:
		images = images[:num_pages]
	results: List[str] = []
	for img in images:
		raw = pytesseract.image_to_string(img, lang="eng", config=tess_config)
		results.append(clean_ocr_text(raw))
	return results


def main() -> None:
	parser = argparse.ArgumentParser(description="OCR a PDF and dump cleaned plaintext per page.")
	parser.add_argument("pdf", help="Path to the PDF")
	parser.add_argument("--pages", default="2", help="How many pages to OCR: N or 'all' (default: 2)")
	parser.add_argument("--write", action="store_true", help="Write each page to its own .txt file")
	args = parser.parse_args()

	pdf_path = Path(args.pdf).expanduser().resolve()
	if not pdf_path.exists():
		raise SystemExit(f"PDF not found: {pdf_path}")

	if args.pages.lower() == "all":
		num_pages = None
	else:
		try:
			num_pages = int(args.pages)
			if num_pages <= 0:
				raise ValueError()
		except ValueError:
			raise SystemExit("--pages must be an integer > 0 or 'all'")

	# Tesseract config: LSTM engine, assume a block of text, preserve spaces
	tess_config = "--oem 1 --psm 6 -c preserve_interword_spaces=1"

	text_pages = ocr_pdf_pages(pdf_path, num_pages=num_pages, tess_config=tess_config)

	# Print combined to stdout, with clear page separators
	for i, page_text in enumerate(text_pages, start=1):
		print(f"===== Page {i} =====")
		print(page_text, end="")
		if i != len(text_pages):
			print("")

	if args.write:
		for i, page_text in enumerate(text_pages, start=1):
			out = pdf_path.with_name(f"{pdf_path.stem}_ocr_page-{i:02d}.txt")
			out.write_text(page_text, encoding="utf-8")
			print(f"Wrote: {out}")


if __name__ == "__main__":
	main()



