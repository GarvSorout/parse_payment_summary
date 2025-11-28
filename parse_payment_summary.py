#!/usr/bin/env python3
import argparse
import csv
import json
import re
import subprocess
import sys
from pathlib import Path
from typing import List, Tuple
import tempfile
import shlex


def run_command(command_args: list[str]) -> None:
	"""Run a system command, raising on failure."""
	subprocess.run(command_args, check=True)


def _render_with_pdftoppm(pdf_path: Path, dpi: int, work_dir: Path) -> List[Path]:
	"""Render PDF pages to PNGs using pdftoppm."""
	base = work_dir / (pdf_path.stem + "_ppm_page")
	run_command([
		"pdftoppm",
		"-r", str(dpi),
		"-png",
		str(pdf_path),
		str(base)
	])
	return sorted(work_dir.glob(base.name + "-*.png"))


def _render_with_pdf2image(pdf_path: Path, dpi: int, work_dir: Path) -> List[Path]:
	"""Render PDF pages to PNGs using pdf2image (Pillow)."""
	try:
		from pdf2image import convert_from_path  # type: ignore
	except Exception as e:
		raise RuntimeError("pdf2image is not installed. Install with: pip install pdf2image Pillow") from e
	images = convert_from_path(str(pdf_path), dpi=dpi, fmt="png")
	out_paths: List[Path] = []
	for idx, img in enumerate(images, start=1):
		out_p = work_dir / f"{pdf_path.stem}_pdf2img_page-{idx:02d}.png"
		img.save(out_p, "PNG")
		out_paths.append(out_p)
	return out_paths


def ocr_pdf_to_text(pdf_path: Path, out_txt_path: Path, dpi: int = 300, lang: str = "eng", renderer: str = "pdftoppm", tess_args: List[str] | None = None) -> None:
	"""
	OCR a PDF into text:
	- Rasterizes pages to PNG using the chosen renderer (pdftoppm or pdf2image)
	- Runs tesseract on each page
	- Concatenates results into a single .txt
	"""
	with tempfile.TemporaryDirectory(prefix="ocr_pages_") as td:
		work_dir = Path(td)
		# 1) Rasterize
		if renderer.lower() == "pdf2image":
			page_images = _render_with_pdf2image(pdf_path, dpi, work_dir)
		else:
			page_images = _render_with_pdftoppm(pdf_path, dpi, work_dir)
		if not page_images:
			raise RuntimeError("No page images produced; ensure rendering tool is installed and PDF is readable.")
		# 2) OCR each image
		ocr_txts: list[Path] = []
		for img in page_images:
			out_base = img.with_suffix("")  # remove .png
			# Default Tesseract settings: LSTM engine, block-of-text layout, preserve spaces
			default_tess = ["--oem", "1", "--psm", "6", "-c", "preserve_interword_spaces=1"]
			cmd = ["tesseract", str(img), str(out_base), "-l", lang] + (tess_args if tess_args is not None else default_tess)
			run_command(cmd)
			ocr_txts.append(out_base.with_suffix(".txt"))
		# 3) Concatenate
		with out_txt_path.open("w", encoding="utf-8") as w:
			for f in ocr_txts:
				w.write(f.read_text(encoding="utf-8", errors="ignore"))
				w.write("\n\f\n")


def read_text_from_input(input_path: Path, use_ocr: bool, renderer: str, dpi: int, tess_args: List[str] | None) -> tuple[str, Path]:
	"""
	Return text content and a path to a sidecar OCR .txt (created if needed).
	- If input is .txt and looks like OCR output, just read it.
	- Else OCR the PDF into a .txt.
	"""
	if input_path.suffix.lower() == ".txt" and not use_ocr:
		return input_path.read_text(encoding="utf-8", errors="ignore"), input_path
	# Otherwise perform OCR (or re-OCR if explicitly requested)
	out_txt = input_path.with_suffix("").with_name(input_path.stem + "_ocr.txt")
	ocr_pdf_to_text(input_path, out_txt, dpi=dpi, renderer=renderer, tess_args=tess_args)
	return out_txt.read_text(encoding="utf-8", errors="ignore"), out_txt


def parse_metadata(header_text: str) -> dict:
	meta: dict[str, str | None] = {}
	m_period = re.search(r"FOR PERIOD .*?:\s*(\d{4}-\d{2}-\d{2})\s*TO\s*(\d{4}-\d{2}-\d{2})", header_text)
	m_remit = re.search(r"REMITTANCE ADVICE:\s*([A-Za-z]+\s+\d{4})", header_text)
	m_run = re.search(r"Run Date:\s*([0-9\-: ]+[AP]M)", header_text, flags=re.IGNORECASE)
	m_group = re.search(r"GROUP:\s*(.+?)\s+PAYMENT TO:\s*(\w+)", header_text)
	m_group_no = re.search(r"GROUP\s*#:\s*([A-Z0-9]+)", header_text)
	meta["period_from"] = m_period.group(1) if m_period else None
	meta["period_to"] = m_period.group(2) if m_period else None
	meta["remittance_advice"] = m_remit.group(1) if m_remit else None
	meta["run_date"] = m_run.group(1) if m_run else None
	meta["group_name"] = m_group.group(1).strip() if m_group else None
	meta["payment_to"] = m_group.group(2) if m_group else None
	meta["group_no"] = m_group_no.group(1) if m_group_no else None
	return meta


_NUM_RE = re.compile(r"(-?\d{1,3}(?:,\d{3})*(?:\.\d{2})?)")
_ONLY_SC_RE = re.compile(r"^[scSC]{2,}$")
_WORD_RE = re.compile(r"^[A-Za-z][A-Za-z\-']*$")
_META_TOKENS = ("REPORT", "RUN DATE", "PAGE:", "GROUP #", "REMITTANCE", "FOR PERIOD", "OHIP PAYMENT SUMMARY")

# Canonical category names (uppercased), with simple normalization to match common OCR variants
_CATEGORY_CANON = {
	"ACCESS BONUS PAYMENT": "ACCESS BONUS PAYMENT",
	"LTC ACCESS BONUS PAYMENT": "LTC ACCESS BONUS PAYMENT",
	"GROUP MANAGEMENT LEADERSHIP PAYMENT": "GROUP MANAGEMENT LEADERSHIP PAYMENT",
	"OFFICE PRACTICE ADMINISTRATION PAYMENT": "OFFICE PRACTICE ADMINISTRATION PAYMENT",
	"HCP RELATIVITY PAYMENT": "HCP RELATIVITY PAYMENT",
	"RMB RELATIVITY PAYMENT": "RMB RELATIVITY PAYMENT",
	"WSIB RELATIVITY PAYMENT": "WSIB RELATIVITY PAYMENT",
	"YEAR 1 (2024-2025) COMPENSATION INCREASE": "YEAR 1 (2024-2025) COMPENSATION INCREASE",
	"NETWORK BASE RATE PAYMENT": "NETWORK BASE RATE PAYMENT",
	"BASE RATE PAYMENT RECONCILIATION ADJMT": "BASE RATE PAYMENT RECONCILIATION ADJMT",
	"BASE RATE ACUITY PAYMENT": "BASE RATE ACUITY PAYMENT",
	"BASE RATE ACUITY ADJUSTMENT": "BASE RATE ACUITY ADJUSTMENT",
	"COMP CARE CAPITATION": "COMP CARE CAPITATION",
	"COMP CARE RECONCILIATION": "COMP CARE RECONCILIATION",
	"BLENDED FEE-FOR-SERVICE PREMIUM": "BLENDED FEE-FOR-SERVICE PREMIUM",
	"BLENDED FEE FOR SERVICE PREMIUM": "BLENDED FEE-FOR-SERVICE PREMIUM",
	"PREVENTIVE CARE BONUS": "PREVENTIVE CARE BONUS",
	"TOTAL CLAIMS PAYABLE": "TOTAL CLAIMS PAYABLE",
	"AGE PREMIUM PAYMENT": "AGE PREMIUM PAYMENT",
	"SPECIAL PREMIUM PAYMENT": "SPECIAL PREMIUM PAYMENT",
}


def _normalize_spaces(s: str) -> str:
	return re.sub(r"\s+", " ", s).strip()


def normalize_category(cat_raw: str) -> str:
	"""
	Normalize category labels to canonical forms:
	- Collapse spaces, remove stray punctuation.
	- Accept both 'FEE-FOR-SERVICE' and 'FEE FOR SERVICE'.
	- Insert spaces around parentheses for YEAR line if missing.
	"""
	cat = cat_raw.upper()
	cat = cat.replace("—", "-").replace("–", "-")
	cat = cat.replace("  ", " ")
	# Ensure spaces around parentheses in YEAR line
	cat = re.sub(r"\bYEAR\s*(\d)\s*\((\d{4}-\d{4})\)\s*COMPENSATION\s*INCREASE", r"YEAR \1 (\2) COMPENSATION INCREASE", cat)
	# Remove trailing colons/dashes
	cat = cat.strip(" :-")
	# Normalize hyphen/spaces in FFS
	cat = cat.replace("FEE - FOR - SERVICE", "FEE FOR SERVICE")
	cat = cat.replace("FEE-FOR-SERVICE", "FEE-FOR-SERVICE")
	cat = _normalize_spaces(cat)
	# Try exact canonical match
	if cat in _CATEGORY_CANON:
		return _CATEGORY_CANON[cat]
	# Try relaxed match (without hyphens)
	relaxed = cat.replace("-", " ")
	if relaxed in _CATEGORY_CANON:
		return _CATEGORY_CANON[relaxed]
	return cat_raw.strip()


def line_to_category_and_numbers(line: str) -> tuple[str | None, float | None, float | None]:
	"""
	Heuristically parse a line into (category, current_month, year_to_date) if it ends with two currency-like numbers.
	Returns (None, None, None) if the pattern is not matched.
	"""
	line = line.strip()
	if not line:
		return None, None, None
	nums = _NUM_RE.findall(line)
	if len(nums) < 2:
		return None, None, None
	# The category is the text up to the start of the last two numbers
	try:
		cur_val = float(nums[-2].replace(",", ""))
		ytd_val = float(nums[-1].replace(",", ""))
	except ValueError:
		return None, None, None
	# Remove trailing two number tokens from the raw line to isolate category
	# Crude but effective: split by the last occurrence of those numbers
	tail = f"{nums[-2]} {nums[-1]}"
	if line.endswith(tail):
		head = line[: -len(tail)].rstrip()
	else:
		# Fallback: try removing just last number then second last
		tmp = line.rsplit(nums[-1], 1)[0].rstrip()
		head = tmp.rsplit(nums[-2], 1)[0].rstrip()
	# Ensure the head has letters (avoid pure noise)
	if not re.search(r"[A-Za-z]", head):
		return None, None, None
	# Normalize spacing
	head = re.sub(r"\s+", " ", head).strip(":-–— ")
	# Normalize category
	canon = normalize_category(head)
	return canon, cur_val, ytd_val


def parse_group_categories(pages: list[str]) -> list[tuple[str, float, float]]:
	"""
	Parse group-level categories from the first two pages (summary pages).
	Returns list of (category, current_month, ytd).
	"""
	categories: dict[str, tuple[float, float]] = {}
	for pi in (0, 1):
		if pi >= len(pages):
			break
		for raw in pages[pi].splitlines():
			cat, cur, ytd = line_to_category_and_numbers(raw)
			if cat is None:
				continue
			# Filter lines that are obviously meta/noise
			if any(k in cat.upper() for k in ("REPORT:", "GROUP #", "RUN DATE", "PAGE:", "BROOKLIN MEDICAL CENTRE", "OHIP PAYMENT SUMMARY")):
				continue
			categories[cat] = (cur, ytd)
	return [(k, v[0], v[1]) for k, v in categories.items()]


def clean_provider_name_raw(cand: str) -> str:
	"""
	Remove obvious OCR noise from a provider name candidate:
	- Strip non-letter punctuation.
	- Drop tokens that are only S/C sequences like 'SS', 'SCSCS'.
	- Keep 2-4 meaningful tokens, title-case, and fix common 'Mc' capitalization.
	"""
	clean = re.sub(r"[^A-Za-z\s\-']", " ", cand)
	clean = _normalize_spaces(clean)
	if not clean:
		return ""
	tokens = []
	for tok in clean.split(" "):
		if not _WORD_RE.match(tok):
			continue
		if _ONLY_SC_RE.match(tok):
			continue
		tokens.append(tok)
	if len(tokens) < 2:
		return ""
	# Limit to first 4 tokens to avoid trailing noise
	tokens = tokens[:4]
	# Drop trailing single-letter artifacts (e.g., stray 'C' from header bleed)
	while len(tokens) >= 2 and len(tokens[-1]) == 1:
		tokens.pop()
	# Title-case with 'Mc' handling
	def fix_case(t: str) -> str:
		if len(t) >= 3 and t[:2].lower() == "mc":
			return "Mc" + t[2:].title()
		return t.title()
	return " ".join(fix_case(t) for t in tokens)


def _canon_name_for_match(name: str) -> str:
	"""
	Canonicalize a provider name for matching:
	- Uppercase
	- Remove non-letters
	- Collapse whitespace
	"""
	name = name.upper()
	name = re.sub(r"[^A-Z]+", " ", name)
	return re.sub(r"\s+", " ", name).strip()


def extract_provider_name_from_page(page_text: str) -> str:
	"""
	Find provider name by scanning lines above 'NETWORK BASE RATE PAYMENT'.
	Attempt to clean OCR artifacts.
	"""
	# Normalize control characters to spaces to avoid hidden separators
	page_text_norm = re.sub(r"[\x00-\x08\x0b-\x0c\x0e-\x1f]", " ", page_text)
	lines = page_text_norm.splitlines()
	idx = -1
	for i, ln in enumerate(lines):
		if "NETWORK BASE RATE PAYMENT" in ln.upper():
			idx = i
			break
	if idx == -1:
		# fallback: first reasonable uppercase line
		for ln in lines[:10]:
			cand = ln.strip()
			if re.search(r"[A-Za-z]", cand) and len(cand) <= 60:
				idx = 0
				break
	# Scan upwards a few lines
	name = None
	for j in range(max(0, idx - 6), idx):
		cand = lines[j].strip()
		cleaned = clean_provider_name_raw(cand)
		if cleaned:
			name = cleaned
	# Title-case conservatively
	if not name:
		return "Unknown Provider"
	# Keep uppercase acronyms if any, but generally title-case
	return name


def parse_provider_page(page_text: str) -> tuple[str, list[tuple[str, float, float]]]:
	"""
	Parse a single provider page into:
	- provider_name
	- list of (category, current_month, ytd)
	"""
	provider = extract_provider_name_from_page(page_text)
	rows: list[tuple[str, float, float]] = []
	for raw in page_text.splitlines():
		cat, cur, ytd = line_to_category_and_numbers(raw)
		if cat is None:
			continue
		# Skip meta/noise
		if any(k in cat.upper() for k in ("REPORT:", "GROUP #", "RUN DATE", "PAGE:", "BROOKLIN MEDICAL CENTRE", "OHIP PAYMENT SUMMARY", "FOR PERIOD", "REMITTANCE")):
			continue
		rows.append((cat, cur, ytd))
	return provider, rows


def write_group_csvs(group_rows: list[tuple[str, float, float]], out_dir: Path) -> None:
	group_csv = out_dir / "group_categories.csv"
	with group_csv.open("w", newline="", encoding="utf-8") as f:
		w = csv.writer(f)
		w.writerow(["category", "current_month", "year_to_date"])
		for cat, cur, ytd in group_rows:
			w.writerow([cat, f"{cur:.2f}", f"{ytd:.2f}"])


def write_provider_csvs(provider_pages: list[str], out_dir: Path) -> None:
	# Long format of all provider category lines
	prov_cat_csv = out_dir / "provider_categories.csv"
	# Totals per provider (prefer the 'TOTAL CLAIMS PAYABLE' row if present, else sum)
	prov_tot_csv = out_dir / "provider_totals.csv"
	with prov_cat_csv.open("w", newline="", encoding="utf-8") as fcat, prov_tot_csv.open("w", newline="", encoding="utf-8") as ftot:
		wcat = csv.writer(fcat)
		wtot = csv.writer(ftot)
		wcat.writerow(["provider", "category", "current_month", "year_to_date"])
		wtot.writerow(["provider", "current_month_total_claims_payable", "ytd_total_claims_payable"])
		for page in provider_pages:
			provider, rows = parse_provider_page(page)
			total_cur = None
			total_ytd = None
			sum_cur = 0.0
			sum_ytd = 0.0
			for cat, cur, ytd in rows:
				wcat.writerow([provider, cat, f"{cur:.2f}", f"{ytd:.2f}"])
				sum_cur += cur
				sum_ytd += ytd
				if "TOTAL CLAIMS PAYABLE" in cat.upper():
					total_cur = cur
					total_ytd = ytd
			if total_cur is None or total_ytd is None:
				total_cur, total_ytd = sum_cur, sum_ytd
			wtot.writerow([provider, f"{total_cur:.2f}", f"{total_ytd:.2f}"])


def write_provider_csvs_from_entries(provider_entries: List[dict], out_dir: Path) -> None:
	# Long format of all provider category lines
	prov_cat_csv = out_dir / "provider_categories.csv"
	# Totals per provider
	prov_tot_csv = out_dir / "provider_totals.csv"
	with prov_cat_csv.open("w", newline="", encoding="utf-8") as fcat, prov_tot_csv.open("w", newline="", encoding="utf-8") as ftot:
		wcat = csv.writer(fcat)
		wtot = csv.writer(ftot)
		wcat.writerow(["provider", "provider_id", "category", "current_month", "year_to_date"])
		wtot.writerow(["provider", "provider_id", "current_month_total_claims_payable", "ytd_total_claims_payable"])
		for p in provider_entries:
			name = p["name"]
			rows: List[Tuple[str, float, float]] = p["rows"]
			provider_id = p.get("id")
			total_cur = p.get("total_cur")
			total_ytd = p.get("total_ytd")
			if total_cur is None or total_ytd is None:
				total_cur = sum(r[1] for r in rows)
				total_ytd = sum(r[2] for r in rows)
			for cat, cur, ytd in rows:
				# Skip meta categories defensively
				if any(tok in cat.upper() for tok in _META_TOKENS):
					continue
				wcat.writerow([name, provider_id or "", cat, f"{cur:.2f}", f"{ytd:.2f}"])
			wtot.writerow([name, provider_id or "", f"{total_cur:.2f}", f"{total_ytd:.2f}"])


def write_combined_readable_text(meta: dict, group_rows: list[tuple[str, float, float]], provider_entries: List[dict], out_dir: Path) -> None:
	"""
	Write a single human-friendly text file that combines:
	- Metadata
	- Group categories
	- Providers index (name + totals)
	- Detailed provider sections with categories and amounts
	"""
	# Keep document order (do not sort)

	def build_text() -> str:
		lines: list[str] = []
		lines.append("OHIP Payment Summary")
		if meta.get("period_from") and meta.get("period_to"):
			lines.append(f"Period: {meta.get('period_from')} to {meta.get('period_to')}")
		if meta.get("remittance_advice"):
			lines.append(f"Remittance Advice: {meta.get('remittance_advice')}")
		if meta.get("run_date"):
			lines.append(f"Run Date: {meta.get('run_date')}")
		if meta.get("group_name") or meta.get("group_no") or meta.get("payment_to"):
			grp_bits = []
			if meta.get("group_name"):
				grp_bits.append(str(meta.get("group_name")))
			if meta.get("group_no"):
				grp_bits.append(f"#{meta.get('group_no')}")
			if meta.get("payment_to"):
				grp_bits.append(f"(Payment to: {meta.get('payment_to')})")
			lines.append("Group: " + " ".join(grp_bits))
		lines.append("")
		lines.append("Group Summary (Current Month; Year to Date):")
		for cat, cur, ytd in group_rows:
			lines.append(f"- {cat}: {cur:,.2f}; {ytd:,.2f}")
		lines.append("")
		lines.append("Providers Index (TOTAL CLAIMS PAYABLE):")
		for p in provider_entries:
			pi = f" (ID: {p['id']})" if p.get("id") else ""
			lines.append(f"- {p['name']}{pi}: {p['total_cur']:,.2f}; {p['total_ytd']:,.2f}")
		lines.append("")
		lines.append("Provider Details:")
		for p in provider_entries:
			pi = f" (ID: {p['id']})" if p.get("id") else ""
			lines.append(f"Provider: {p['name']}{pi}")
			for cat, cur, ytd in p["rows"]:
				# Skip meta categories in readable output
				if any(tok in cat.upper() for tok in _META_TOKENS):
					continue
				lines.append(f"  - {cat}: {cur:,.2f}; {ytd:,.2f}")
			lines.append("")
		return "\n".join(lines).strip() + "\n"

	def build_markdown() -> str:
		lines: list[str] = []
		lines.append("# OHIP Payment Summary")
		md_meta = []
		if meta.get("period_from") and meta.get("period_to"):
			md_meta.append(f"**Period**: {meta.get('period_from')} to {meta.get('period_to')}")
		if meta.get("remittance_advice"):
			md_meta.append(f"**Remittance Advice**: {meta.get('remittance_advice')}")
		if meta.get("run_date"):
			md_meta.append(f"**Run Date**: {meta.get('run_date')}")
		if meta.get("group_name") or meta.get("group_no") or meta.get("payment_to"):
			grp_bits = []
			if meta.get("group_name"):
				grp_bits.append(str(meta.get("group_name")))
			if meta.get("group_no"):
				grp_bits.append(f"#{meta.get('group_no')}")
			if meta.get("payment_to"):
				grp_bits.append(f"(Payment to: {meta.get('payment_to')})")
			md_meta.append(f"**Group**: " + " ".join(grp_bits))
		if md_meta:
			lines.append("\n".join(md_meta))
		lines.append("\n## Group Summary (Current Month; Year to Date)")
		for cat, cur, ytd in group_rows:
			lines.append(f"- {cat}: {cur:,.2f}; {ytd:,.2f}")
		lines.append("\n## Providers Index (TOTAL CLAIMS PAYABLE)")
		for p in provider_entries:
			pi = f" (ID: {p['id']})" if p.get("id") else ""
			lines.append(f"- {p['name']}{pi}: {p['total_cur']:,.2f}; {p['total_ytd']:,.2f}")
		lines.append("\n## Provider Details")
		for p in provider_entries:
			pi = f" (ID: {p['id']})" if p.get("id") else ""
			lines.append(f"### {p['name']}{pi}")
			for cat, cur, ytd in p["rows"]:
				if any(tok in cat.upper() for tok in _META_TOKENS):
					continue
				lines.append(f"- {cat}: {cur:,.2f}; {ytd:,.2f}")
			lines.append("")
		return "\n".join(lines).strip() + "\n"

	(out_dir / "combined_readable.txt").write_text(build_text(), encoding="utf-8")
	(out_dir / "combined_readable.md").write_text(build_markdown(), encoding="utf-8")


def _normalize_unicode_punct(s: str) -> str:
	# Common OCR unicode punctuation to ASCII
	return (s.replace("—", "-")
	        .replace("–", "-")
	        .replace("‐", "-")
	        .replace("’", "'")
	        .replace("‘", "'")
	        .replace("“", '"')
	        .replace("”", '"')
	        .replace("\u00a0", " "))


def _looks_like_noise(line: str) -> bool:
	"""
	Detect obvious decorative/garble lines:
	- Only punctuation and whitespace
	- Long runs of ~ * - _
	- Lines composed only of S/C (common OCR junk) of length >= 6
	- Very low ratio of letters+digits to total
	"""
	s = line.strip()
	if not s:
		return True
	# Only punctuation-like
	if re.fullmatch(r"[~*_\-\s\[\]\(\)\|=\\/]+", s):
		return True
	# Long runs
	if re.search(r"(~{3,}|\*{3,}|_{3,}|\-{5,})", s):
		return True
	up = s.upper()
	# Only S/C (and spaces)
	if re.fullmatch(r"[SC\s]+", up) and len(up.replace(" ", "")) >= 6:
		return True
	# Very low alnum ratio
	alnum = sum(ch.isalnum() for ch in s)
	if alnum / max(len(s), 1) < 0.15:
		return True
	return False


def clean_ocr_text(raw_text: str) -> str:
	"""
	Produce a cleaned OCR text:
	- Normalize unicode punctuation
	- Drop decorative/garbled lines
	- Collapse excessive whitespace
	- Preserve page breaks
	"""
	pages = raw_text.split("\f")
	clean_pages: list[str] = []
	for pg in pages:
		lines_out: list[str] = []
		for ln in pg.splitlines():
			ln = _normalize_unicode_punct(ln)
			ln = re.sub(r"\s+", " ", ln).strip()
			if not ln:
				continue
			if _looks_like_noise(ln):
				continue
			lines_out.append(ln)
		clean_pages.append("\n".join(lines_out))
	return "\n\f\n".join(clean_pages)


def extract_text_with_poppler(pdf_path: Path, out_txt_path: Path, layout: str = "raw") -> None:
	"""
	Extract text using poppler's pdftotext, either in raw or layout mode.
	"""
	args = ["pdftotext"]
	if layout == "raw":
		args.append("-raw")
	elif layout == "layout":
		args.append("-layout")
	args.extend([str(pdf_path), str(out_txt_path)])
	run_command(args)


def decode_caesar_shift(text: str, shift: int = 29, low: int = 33, high: int = 94) -> str:
	"""
	Decode a constant character shift used by the PDF text layer.
	Empirically, characters in ASCII 33..94 map to +29 to yield real letters.
	We decode by mapping c -> chr(ord(c)+shift) for ord in [low, high]; others unchanged.
	"""
	out_chars: list[str] = []
	for ch in text:
		o = ord(ch)
		if low <= o <= high:
			out_chars.append(chr(o + shift))
		else:
			out_chars.append(ch)
	return "".join(out_chars)


def extract_provider_name_and_id_from_decoded(page_text: str) -> tuple[str, str | None]:
	"""
	From decoded text, find the 'GROUP PAYMENTS TO PROVIDER' section and extract the provider name.
	ID digits are typically encoded in the PDF text layer, so we always return None for provider_id here.
	IDs are filled from OCR summary pages/TSV later, but we now also try to read them
	directly from decoded text near the provider name (works if digits decode properly).
	"""
	# Normalize control characters and spacing
	page_text_norm = re.sub(r"[\x00-\x08\x0b-\x0c\x0e-\x1f]", " ", page_text)
	lines = [re.sub(r"\s+", " ", ln).strip() for ln in page_text_norm.splitlines() if ln.strip()]

	name = None
	provider_id = None

	# Build a normalized uppercase line list for robust header matching
	upper_lines = [" ".join(re.findall(r"[A-Z]+", ln.upper())) for ln in lines]

	name_line_idx = None
	for i, up in enumerate(upper_lines):
		if all(w in up for w in ("GROUP", "PAYMENTS", "TO", "PROVIDER")):
			# Look ahead for a plausible name (skip the CURRENT MONTH header)
			for j in range(i + 1, min(i + 8, len(lines))):
				cand = lines[j]
				upcand = upper_lines[j]
				if "CURRENT MONTH" in upcand and "YEAR TO DATE" in upcand:
					continue
				clean = clean_provider_name_raw(cand)
				if clean:
					name = clean
					name_line_idx = j
					break
			break

	if not name:
		name = clean_provider_name_raw(page_text_norm) or "Unknown Provider"

	# Attempt to extract a provider ID from nearby decoded lines (same or following lines)
	if name_line_idx is not None:
		window_start = max(0, name_line_idx - 1)
		window_end = min(len(lines), name_line_idx + 4)
		for k in range(window_start, window_end):
			ln = lines[k]
			# Ignore obvious amount lines (two large currency-like numbers)
			if len(_NUM_RE.findall(ln)) >= 2:
				continue
			m = re.search(r"\b(\d{5,})\b", ln)
			if m:
				provider_id = m.group(1)
				break

	return name, provider_id


def _run_tesseract_tsv(image_path: Path) -> list[dict]:
	"""
	Run Tesseract in TSV mode on an image and return a list of word dicts with bbox/line info.
	"""
	cmd = ["tesseract", str(image_path), "stdout", "-l", "eng", "--oem", "1", "--psm", "6", "tsv"]
	proc = subprocess.run(cmd, check=True, capture_output=True, text=True)
	lines = proc.stdout.splitlines()
	records: list[dict] = []
	if not lines:
		return records
	header = lines[0].split("\t")
	for ln in lines[1:]:
		parts = ln.split("\t")
		if len(parts) != len(header):
			continue
		row = dict(zip(header, parts))
		if row.get("level") == "5":  # word
			# Normalize fields
			for k in ("left", "top", "width", "height", "conf", "page_num", "block_num", "par_num", "line_num", "word_num"):
				if k in row:
					try:
						row[k] = int(row[k]) if k != "conf" else float(row[k])
					except Exception:
						pass
			records.append(row)
	return records


def _find_provider_id_near_name_tsv(tsv_words: list[dict], provider_name: str) -> str | None:
	"""
	Given TSV words and a provider name, locate the name tokens and search for a 6+ digit number
	on the same line or nearby within a small vertical band to the right.
	"""
	if not tsv_words or not provider_name:
		return None
	name_tokens = [t for t in re.split(r"[^A-Za-z']+", provider_name) if t]
	if not name_tokens:
		return None
	# Find candidate rows that contain name token(s)
	name_rows = [w for w in tsv_words if any(w.get("text", "").strip().lower() == tok.lower() for tok in name_tokens)]
	if not name_rows:
		return None
	# Compute rough band: take median line_num and y center from tokens
	line_nums = [w.get("line_num") for w in name_rows if isinstance(w.get("line_num"), int)]
	if not line_nums:
		return None
	target_line = sorted(line_nums)[len(line_nums) // 2]
	# Name bbox union
	name_row_same_line = [w for w in tsv_words if w.get("line_num") == target_line]
	if not name_row_same_line:
		name_row_same_line = name_rows
	name_left = min(w.get("left", 0) for w in name_row_same_line)
	name_top = min(w.get("top", 0) for w in name_row_same_line)
	name_right = max((w.get("left", 0) + w.get("width", 0)) for w in name_row_same_line)
	name_bottom = max((w.get("top", 0) + w.get("height", 0)) for w in name_row_same_line)
	band_top = max(0, name_top - 40)
	band_bottom = name_bottom + 40
	# Search numeric tokens in band, preferably to the right of name
	candidates: list[tuple[str, int]] = []
	for w in tsv_words:
		txt = str(w.get("text", "")).strip()
		if not txt:
			continue
		if not re.fullmatch(r"[0-9]{6,}", txt):
			continue
		left = w.get("left", 0)
		top = w.get("top", 0)
		height = w.get("height", 0)
		bottom = top + height
		if band_top <= top <= band_bottom or band_top <= bottom <= band_bottom:
			# Prefer to the right; penalize left-of-name
			score = left - name_right
			candidates.append((txt, score))
	# Choose closest to the right; if none, choose any in band
	if not candidates:
		return None
	# sort by absolute proximity, but prefer positive (to the right)
	candidates.sort(key=lambda x: (abs(x[1]) if x[1] >= 0 else abs(x[1]) + 10000))
	return candidates[0][0]


def _get_name_bbox_from_tsv(tsv_words: list[dict], provider_name: str) -> tuple[int, int, int, int] | None:
	"""
	Return (left, top, right, bottom) bbox for the provider name line based on TSV tokens.
	"""
	name_tokens = [t for t in re.split(r"[^A-Za-z']+", provider_name) if t]
	if not name_tokens:
		return None
	name_rows = [w for w in tsv_words if any(w.get("text", "").strip().lower() == tok.lower() for tok in name_tokens)]
	if not name_rows:
		return None
	line_nums = [w.get("line_num") for w in name_rows if isinstance(w.get("line_num"), int)]
	if not line_nums:
		return None
	target_line = sorted(line_nums)[len(line_nums) // 2]
	line_words = [w for w in tsv_words if w.get("line_num") == target_line]
	if not line_words:
		line_words = name_rows
	left = min(w.get("left", 0) for w in line_words)
	top = min(w.get("top", 0) for w in line_words)
	right = max((w.get("left", 0) + w.get("width", 0)) for w in line_words)
	bottom = max((w.get("top", 0) + w.get("height", 0)) for w in line_words)
	return left, top, right, bottom


def _ocr_digits_from_crop(image_path: Path, crop_box: tuple[int, int, int, int]) -> list[str]:
	"""
	Crop region and OCR with digits whitelist; return list of 6+ digit tokens found.
	"""
	from PIL import Image  # lazy import
	img = Image.open(image_path)
	W, H = img.size
	l, t, r, b = crop_box
	l = max(0, min(W - 1, l))
	r = max(0, min(W, r))
	t = max(0, min(H - 1, t))
	b = max(0, min(H, b))
	if r <= l or b <= t:
		return []
	cropped = img.crop((l, t, r, b))
	with tempfile.NamedTemporaryFile(suffix=".png", delete=False) as tmp:
		tmp_path = Path(tmp.name)
		cropped.save(tmp_path, "PNG")
	try:
		cmd = [
			"tesseract",
			str(tmp_path),
			"stdout",
			"-l",
			"eng",
			"--oem",
			"1",
			"--psm",
			"7",
			"-c",
			"tessedit_char_whitelist=0123456789 -",
		]
		proc = subprocess.run(cmd, check=True, capture_output=True, text=True)
		raw = proc.stdout
	finally:
		try:
			tmp_path.unlink(missing_ok=True)
		except Exception:
			pass
	cands: list[str] = []
	for m in re.finditer(r"[0-9][0-9 \-]{4,}[0-9]", raw):
		digits = re.sub(r"\D", "", m.group(0))
		if len(digits) >= 6:
			cands.append(digits)
	return cands


def _extract_page_number(page_text: str) -> int | None:
	"""
	Extract 'Page: X of Y' page number from page text, if present.
	"""
	m = re.search(r"Page:\s*(\d+)\s+of\s+\d+", page_text, flags=re.IGNORECASE)
	if m:
		try:
			return int(m.group(1))
		except Exception:
			return None
	return None


def _find_provider_summary_pages(decoded_pages: list[str], provider_name: str) -> list[int]:
	"""
	Find decoded page indices that contain 'PROVIDER SUMMARY' and the provider name nearby.
	Return list of decoded page indices.
	"""
	indices: list[int] = []
	# Tokenize provider name to robustly match across punctuation like '='
	name_tokens = [t for t in re.split(r"[^A-Z]+", provider_name.upper()) if len(t) >= 2]
	for idx, pg in enumerate(decoded_pages):
		up = re.sub(r"[\x00-\x08\x0b-\x0c\x0e-\x1f]", " ", pg).upper().replace("=", " ")
		if (("PROVIDER SUMMARY" in up) or ("GROUP PAYMENTS TO PROVIDER TOTAL" in up)):
			if all(re.search(rf"\b{re.escape(tok)}\b", up) for tok in name_tokens):
				indices.append(idx)
	return indices


def _extract_id_from_decoded_summary_text(page_text: str, provider_name: str) -> str | None:
	"""
	From a DECODED page's text, extract an ID on the same line as the provider name
	and before 'GROUP PAYMENTS TO PROVIDER TOTAL' or 'PROVIDER SUMMARY TOTAL'.
	Decoded text often shows delimiters like '=' between tokens.
	"""
	if not page_text or not provider_name:
		return None
	pt = re.sub(r"[\x00-\x08\x0b-\x0c\x0e-\x1f]", " ", page_text)
	pt = pt.replace("=", " ")
	# Build tolerant name pattern (tokens in order, flexible spaces)
	toks = [t for t in re.split(r"[^A-Za-z]+", provider_name) if len(t) >= 2]
	if not toks:
		return None
	parts = [re.escape(t) for t in toks]
	name_pat = r"\s+".join(parts)
	name_pat = rf"\b{name_pat}\b"
	patterns = [
		rf"{name_pat}.*?\b([0-9]{{5,}})\b.*?(GROUP\s+PAYMENTS\s+TO\s+PROVIDER\s+TOTAL|PROVIDER\s+SUMMARY\s+TOTAL)",
	]
	for pat in patterns:
		m = re.search(pat, pt, flags=re.IGNORECASE | re.DOTALL)
		if m:
			return m.group(1)
	return None


def _extract_id_from_ocr_summary_text(page_text: str, provider_name: str) -> str | None:
	"""
	From an OCR page's plain text, extract an ID located on the same line as the provider name
	and before 'PROVIDER SUMMARY TOTAL', e.g.:
	'LIBBY, THOMAS              010255 | PROVIDER SUMMARY TOTAL ...'
	Returns cleaned digits or None.
	"""
	if not page_text or not provider_name:
		return None
	pt = page_text
	# Build a tolerant token-based name pattern:
	# - ignore case
	# - allow optional commas, flexible spaces
	# - drop single-letter tokens like stray 'C'
	toks = [t for t in re.split(r"[^A-Za-z]+", provider_name) if len(t) >= 2]
	if not toks:
		return None
	# Construct a pattern that matches tokens in order with optional comma and flexible spaces
	parts = [re.escape(t) for t in toks]
	name_pat = r"\s*,?\s*".join(parts)
	name_pat = rf"\b{name_pat}\b"
	# Two patterns: rich (spaced/dashed) and plain digits, both before 'PROVIDER SUMMARY TOTAL'
	patterns = [
		rf"{name_pat}.*?([0-9][0-9\s:\-\.]{{4,}}[0-9]).*?PROVIDER\s+SUMMARY\s+TOTAL",
		rf"{name_pat}.*?\b([0-9]{{5,}})\b.*?PROVIDER\s+SUMMARY\s+TOTAL",
	]
	for pat in patterns:
		m = re.search(pat, pt, flags=re.IGNORECASE | re.DOTALL)
		if m:
			raw = m.group(1)
			digits = re.sub(r"\D", "", raw)
			if len(digits) >= 5:
				return digits
	return None


def _decoded_num_to_float(token: str) -> float | None:
	"""
	Parse a number token from decoded text, where control chars are used as separators:
	- \x0f behaves like a thousands separator (',')
	- \x11 behaves like a decimal point ('.')
	- \x10 behaves like a leading minus ('-')
	"""
	if not token:
		return None
	orig = token.replace("\x0f", ",").replace("\x11", ".").replace("\x10", "-")
	# Try to reconstruct decimals if we see a trailing 2-digit group separated by whitespace
	orig_trim = orig.strip()
	if "." not in orig_trim:
		# Case: grouped thousands and final 2-digit decimal, e.g. "38 370 25" -> 38370.25
		m_grouped = re.fullmatch(r"(-?\d+(?:\s\d{3})+)\s(\d{2})", orig_trim)
		if m_grouped:
			try:
				int_part = re.sub(r"\s+", "", m_grouped.group(1))
				return float(f"{int_part}.{m_grouped.group(2)}")
			except Exception:
				pass
		m = re.fullmatch(r"(-?\d+)\s+(\d{2})", orig_trim)
		if m:
			try:
				return float(f"{m.group(1)}.{m.group(2)}")
			except Exception:
				pass
	# Remove grouping commas and spaces for general case
	s = orig.replace(",", "").replace(" ", "")
	# Keep only digits, '.', and optional leading '-'
	s = re.sub(r"[^0-9\.\-]", "", s)
	if s in ("", "-", ".", "-."):
		return None
	try:
		return float(s)
	except Exception:
		return None


def _extract_total_payment_ab_from_decoded(first_page_decoded: str) -> tuple[str | None, float | None]:
	"""
	From page 1 decoded text, extract TOTAL PAYMENT A B amount and date.
	Returns (date_str 'YYYY-MM-DD', amount_float)
	"""
	src = first_page_decoded or ""
	up = re.sub(r"[\x00-\x08\x0b-\x0c\x0e-\x1f]", " ", src)
	# Prefer a direct capture of the amount and date from the TOTAL PAYMENT line
	m_direct = re.search(
		r"TOTAL\s+PAYMENT[^=]*=\s*[A-Z][^=]*=\s*([0-9\x0f\x11\s]+)\s*=\s*(\d{4}(?:\x10|-)\d{2}(?:\x10|-)\d{2})",
		src,
		flags=re.IGNORECASE | re.DOTALL,
	)
	if m_direct:
		raw_amt = m_direct.group(1)
		raw_date = m_direct.group(2)
		amt = _decoded_num_to_float(raw_amt)
		date_norm = raw_date.replace("\x10", "-")
		return date_norm, amt
	# Second attempt: split the raw (unsanitized) TOTAL PAYMENT line by '=' and pick last two tokens
	for ln in src.splitlines():
		if "TOTAL" in ln.upper() and "PAYMENT" in ln.upper():
			parts = [p.strip() for p in ln.split("=")]
			if len(parts) >= 3:
				raw_amt = parts[-2]
				raw_date = parts[-1]
				amt = _decoded_num_to_float(raw_amt)
				# Normalize date that may be separated by \x10 or spaces
				if re.fullmatch(r"\d{4}\x10\d{2}\x10\d{2}", raw_date):
					date_norm = raw_date.replace("\x10", "-")
				else:
					msp = re.fullmatch(r"(\d{4})\s+(\d{2})\s+(\d{2})", re.sub(r"[\x00-\x1f]", " ", raw_date).strip())
					date_norm = f"{msp.group(1)}-{msp.group(2)}-{msp.group(3)}" if msp else None
				if date_norm or amt is not None:
					return date_norm, amt
	# Locate the line that includes 'TOTAL PAYMENT'
	for ln in up.splitlines():
		if "TOTAL" in ln.upper() and "PAYMENT" in ln.upper():
			# Find a date like 2025\x10 09 \x10 15 and normalize
			dm = re.search(r"(\d{4}\x10\d{2}\x10\d{2})", ln)
			date_str = None
			if dm:
				date_str = dm.group(1).replace("\x10", "-")
			# Find the largest number-like token on the line as amount
			# Split by '=' first, then scan the parts
			parts = [p.strip() for p in ln.split("=")]
			best = None
			for p in parts:
				val = _decoded_num_to_float(p)
				if val is None:
					continue
				if best is None or val > best:
					best = val
			return date_str, best
	# Fallback: search whole page for the special date pattern
	dm_all = re.findall(r"(\d{4}\x10\d{2}\x10\d{2})", up)
	date_any = dm_all[-1].replace("\x10", "-") if dm_all else None
	if not date_any:
		dm_dash = re.findall(r"(\d{4}-\d{2}-\d{2})", up)
		if dm_dash:
			date_any = dm_dash[-1]
	return date_any, None


def _iter_group_sections_from_decoded(pages_dec: list[str]):
	"""
	Yield tuples of (payment_type, line_item_desc, current_month, year_to_date) for group-level sections.
	We consider four payment types:
	- GROUP PAYMENTS
	- EXCEPTION PAYMENTS
	- GROUP PAYMENTS ALL PROVIDERS
	- SUMMARY ALL PROVIDERS
	"""
	payment_type = None
	expect_matrix = False
	for pg in pages_dec:
		# Normalize line text and convert '=' separators to a token we can split on
		lines = [re.sub(r"[\x00-\x08\x0b-\x0c\x0e-\x1f]", " ", ln) for ln in pg.splitlines()]
		for raw in lines:
			up = raw.upper()
			# Detect headers; ignore provider pages
			if "GROUP PAYMENTS TO PROVIDER" in up:
				payment_type = None
				expect_matrix = False
				continue
			if "GROUP PAYMENTS ALL PROVIDERS" in up:
				payment_type = "GROUP PAYMENTS ALL PROVIDERS"
				expect_matrix = False
				continue
			if "SUMMARY" in up and "ALL PROVIDERS" in up:
				payment_type = "SUMMARY ALL PROVIDERS"
				expect_matrix = False
				continue
			# If header line contains the matrix labels on the same line, enable immediately
			has_matrix_header = ("CURRENT MONTH" in up and "YEAR TO DATE" in up)
			if "GROUP PAYMENTS" in up and "EXCEPTION" not in up:
				payment_type = "GROUP PAYMENTS"
				expect_matrix = has_matrix_header or expect_matrix
				# Do not attempt to parse this header line as a row
				continue
			if "EXCEPTION PAYMENTS" in up:
				payment_type = "EXCEPTION PAYMENTS"
				expect_matrix = has_matrix_header or expect_matrix
				continue
			# Detect the start of the (CURRENT MONTH=YEAR TO DATE) matrix for the active section
			if payment_type and ("CURRENT MONTH" in up and "YEAR TO DATE" in up):
				expect_matrix = True
				continue
			# Collect rows while inside a section matrix
			if not payment_type or not expect_matrix:
				continue
			# Split by '=' to get [label, cur, ytd]
			if "=" not in raw:
				continue
			parts = [p.strip() for p in raw.split("=")]
			if len(parts) < 3:
				continue
			label = normalize_category(parts[0])
			# Skip totals lines for this CSV
			if "TOTAL" in label.upper():
				continue
			cur = _decoded_num_to_float(parts[-2])
			ytd = _decoded_num_to_float(parts[-1])
			if cur is None or ytd is None:
				continue
			yield (payment_type, label, cur, ytd)


def write_group_payments_csv(meta: dict, pages_dec: list[str], out_dir: Path) -> None:
	"""
	Write group_payments.csv with repeated metadata columns and group-level payment sections.
	"""
	group_num = meta.get("group_no") or ""
	period = ""
	if meta.get("period_from") and meta.get("period_to"):
		period = f"{meta['period_from']} to {meta['period_to']}"
	ab_date, ab_amount = _extract_total_payment_ab_from_decoded(pages_dec[0] if pages_dec else "")
	out_csv = out_dir / "group_payments.csv"
	with out_csv.open("w", newline="", encoding="utf-8") as f:
		w = csv.writer(f)
		w.writerow(["group_num", "period", "total_payment_ab_date", "total_payment_ab", "payment_type", "line_item_desc", "current_month_amt", "year_to_date_amt"])
		for payment_type, label, cur, ytd in _iter_group_sections_from_decoded(pages_dec):
			w.writerow([group_num, period, ab_date or "", f"{(ab_amount or 0.0):.2f}", payment_type, label, f"{cur:.2f}", f"{ytd:.2f}"])


def _iter_provider_section_rows_from_decoded(page_text: str, section_name: str) -> list[tuple[str, float, float]]:
	"""
	From a decoded provider page, return list of (label, cur, ytd) rows for the given section.
	section_name is either 'GROUP PAYMENTS TO PROVIDER' or 'PROVIDER SUMMARY'.
	"""
	rows: list[tuple[str, float, float]] = []
	up_lines = [re.sub(r"[\x00-\x08\x0b-\x0c\x0e-\x1f]", " ", ln) for ln in page_text.splitlines()]
	expect_matrix = False
	for raw in up_lines:
		up = raw.upper()
		if section_name == "GROUP PAYMENTS TO PROVIDER" and "GROUP PAYMENTS TO PROVIDER" in up:
			expect_matrix = False
			continue
		if section_name == "PROVIDER SUMMARY" and "PROVIDER SUMMARY" in up:
			expect_matrix = False
			continue
		if "CURRENT MONTH" in up and "YEAR TO DATE" in up:
			expect_matrix = True
			continue
		if not expect_matrix:
			continue
		if "=" not in raw:
			continue
		parts = [p.strip() for p in raw.split("=")]
		if len(parts) < 3:
			continue
		label = normalize_category(parts[0])
		if "TOTAL" in label.upper():
			continue
		cur = _decoded_num_to_float(parts[-2])
		ytd = _decoded_num_to_float(parts[-1])
		if cur is None or ytd is None:
			continue
		rows.append((label, cur, ytd))
	return rows


def write_provider_payments_csv(meta: dict, provider_entries: List[dict], pages_dec: list[str], out_dir: Path) -> None:
	"""
	Write provider_payments.csv with repeated metadata columns and provider detail lines across:
	- GROUP PAYMENTS TO PROVIDER
	- PROVIDER SUMMARY
	We use provider_entries rows (from OCR) for the first section and parse decoded summary pages for the second.
	"""
	group_num = meta.get("group_no") or ""
	period = ""
	if meta.get("period_from") and meta.get("period_to"):
		period = f"{meta['period_from']} to {meta['period_to']}"
	ab_date, ab_amount = _extract_total_payment_ab_from_decoded(pages_dec[0] if pages_dec else "")
	# Build quick lookup of decoded pages per provider for 'PROVIDER SUMMARY'
	out_csv = out_dir / "provider_payments.csv"
	with out_csv.open("w", newline="", encoding="utf-8") as f:
		w = csv.writer(f)
		w.writerow(["group_num", "period", "total_payment_ab_date", "total_payment_ab", "ohip_number", "provider_name", "line_item", "current_month_amt", "year_to_date_amt", "section"])
		# First, rows from provider_entries (GROUP PAYMENTS TO PROVIDER)
		for p in provider_entries:
			name = p["name"]
			ohip = p.get("id") or ""
			for cat, cur, ytd in p["rows"]:
				# Skip meta categories
				if any(tok in cat.upper() for tok in _META_TOKENS):
					continue
				w.writerow([group_num, period, ab_date or "", f"{(ab_amount or 0.0):.2f}", ohip, name, cat, f"{cur:.2f}", f"{ytd:.2f}", "GROUP PAYMENTS TO PROVIDER"])
		# Next, decoded PROVIDER SUMMARY pages
		for p in provider_entries:
			name = p["name"]
			ohip = p.get("id") or ""
			# Find summary pages in decoded that include this provider
			summary_idxs = _find_provider_summary_pages(pages_dec, name)
			for si in summary_idxs:
				rows = _iter_provider_section_rows_from_decoded(pages_dec[si], "PROVIDER SUMMARY")
				for label, cur, ytd in rows:
					w.writerow([group_num, period, ab_date or "", f"{(ab_amount or 0.0):.2f}", ohip, name, label, f"{cur:.2f}", f"{ytd:.2f}", "PROVIDER SUMMARY"])


def _find_id_for_provider_in_ocr_pages(pages_ocr: list[str], provider_name: str) -> str | None:
	"""
	Lenient OCR text scan: for each OCR page line, if it contains the provider name tokens
	(in any punctuation/case), extract the first 5+ digit-like chunk on that line.
	Prefer lines that also contain 'TOTAL' or 'PROVIDER' to reduce false positives.
	"""
	toks = [t for t in re.split(r"[^A-Za-z]+", provider_name) if len(t) >= 2]
	if not toks:
		return None
	for pg in pages_ocr:
		for raw in pg.splitlines():
			line_up = raw.upper()
			# Name tokens must all appear (ignoring punctuation/spacing)
			if not all(re.search(rf"\b{re.escape(t.upper())}\b", line_up) for t in toks):
				continue
			# Prefer lines that suggest identity, not amounts
			prefer = ("PROVIDER" in line_up) or ("TOTAL" in line_up) or ("GROUP" in line_up)
			m = re.search(r"([0-9][0-9\s:\-\.]{4,}[0-9])", raw)
			if not m:
				m = re.search(r"\b([0-9]{5,})\b", raw)
			if m:
				digits = re.sub(r"\D", "", m.group(1))
				if len(digits) >= 5:
					# Return immediately if preferred context, else keep as fallback
					if prefer:
						return digits
					# Fallback: still a usable ID
					return digits
	return None
def main() -> None:
	parser = argparse.ArgumentParser(description="Parse OHIP Payment Summary PDF/TXT into clean CSV tables.")
	parser.add_argument("input_path", help="Path to PDF or OCR text file")
	parser.add_argument("--use-ocr", action="store_true", help="Force OCR even if a text file is provided")
	parser.add_argument("--out-dir", default="tables", help="Directory to write CSV outputs")
	parser.add_argument("--renderer", choices=["pdftoppm", "pdf2image"], default="pdftoppm", help="Page renderer for OCR")
	parser.add_argument("--dpi", type=int, default=300, help="Rendering DPI for OCR (higher can improve accuracy)")
	parser.add_argument("--tesseract-arg", action="append", default=None, help="Extra args for tesseract (repeatable), e.g., --tesseract-arg=-c --tesseract-arg=preserve_interword_spaces=1")
	parser.add_argument("--source", choices=["ocr", "poppler-decode"], default="ocr", help="Text extraction source")
	parser.add_argument("--poppler-layout", choices=["raw", "layout"], default="raw", help="pdftotext layout mode when using poppler-decode")
	parser.add_argument("--hybrid", action="store_true", help="Use decoded text for names and OCR for numbers")
	args = parser.parse_args()

	in_path = Path(args.input_path).expanduser().resolve()
	if not in_path.exists():
		print(f"Input not found: {in_path}", file=sys.stderr)
		sys.exit(1)
	out_dir = Path(args.out_dir).expanduser().resolve()
	out_dir.mkdir(parents=True, exist_ok=True)

	# Collect tess args if provided
	tess_args: List[str] | None = args.tesseract_arg if args.tesseract_arg else None

	# Choose extraction source(s)
	if args.hybrid:
		# OCR text for amounts
		ocr_text, _ = read_text_from_input(in_path, use_ocr=True, renderer=args.renderer, dpi=args.dpi, tess_args=tess_args)
		# Decoded text for names/headings
		poppler_txt = in_path.with_suffix("").with_name(in_path.stem + "_poppler.txt")
		extract_text_with_poppler(in_path, poppler_txt, layout=args.poppler_layout)
		raw_dec = poppler_txt.read_text(encoding="utf-8", errors="ignore")
		# Extend decode range lower to include digits (encoded 0x13..0x1c -> '0'..'9')
		decoded_text = decode_caesar_shift(raw_dec, shift=29, low=19, high=94)
		pages_ocr = ocr_text.split("\f")
		pages_dec = decoded_text.split("\f")
		if not pages_ocr or not pages_dec:
			print("No pages found in OCR/decoded text.", file=sys.stderr)
			sys.exit(2)
		# Build a provider ID map from OCR text (summary/total lines)
		def _build_id_map_from_ocr_pages(pages: list[str]) -> dict[str, str]:
			id_map: dict[str, str] = {}
			for pg in pages:
				for raw in pg.splitlines():
					u = raw.upper()
					if "PROVIDER SUMMARY TOTAL" not in u and "GROUP PAYMENTS TO PROVIDER TOTAL" not in u:
						continue
					m = re.search(r"^(.+?)\s+(\d{5,})\s*\|\s*(GROUP PAYMENTS TO PROVIDER TOTAL|PROVIDER SUMMARY TOTAL)", raw, flags=re.IGNORECASE)
					if not m:
						continue
					raw_name = m.group(1).strip()
					digits = m.group(2).strip()
					canon = _canon_name_for_match(raw_name)
					if canon and digits:
						id_map[canon] = digits
			return id_map
		id_map = _build_id_map_from_ocr_pages(pages_ocr)
		# Use OCR for group rows, decoded for provider names
		grp_rows = parse_group_categories(pages_ocr)
		# Detect provider-section pages USING DECODED TEXT (reliable header)
		provider_entries: List[dict] = []
		provider_indices: List[int] = []
		for i in range(1, min(len(pages_ocr), len(pages_dec))):
			page_dec_norm = re.sub(r"[\x00-\x08\x0b-\x0c\x0e-\x1f]", " ", pages_dec[i])
			if "GROUP PAYMENTS TO PROVIDER" in page_dec_norm.upper():
				provider_indices.append(i)
		# Fallback in case header detection fails entirely
		if not provider_indices:
			provider_indices = list(range(2, min(len(pages_ocr), len(pages_dec))))
		# Prepare page images for TSV probing
		with tempfile.TemporaryDirectory(prefix="ocr_pages_for_id_") as td_images:
			images_dir = Path(td_images)
			if args.renderer.lower() == "pdf2image":
				page_images = _render_with_pdf2image(in_path, args.dpi, images_dir)
			else:
				page_images = _render_with_pdftoppm(in_path, args.dpi, images_dir)
			# Build entries
			for i in provider_indices:
				name, provider_id = extract_provider_name_and_id_from_decoded(pages_dec[i])
				# First try OCR-derived ID map by canonical name (strip trailing single-letter token like 'C')
				name_for_match = re.sub(r"\b[A-Za-z]\b$", "", name).strip()
				canon = _canon_name_for_match(name_for_match)
				if provider_id is None and canon in id_map:
					provider_id = id_map[canon]
				# If still missing, try lenient OCR scan across all pages for this provider
				if provider_id is None:
					found_any = _find_id_for_provider_in_ocr_pages(pages_ocr, name_for_match)
					if found_any:
						provider_id = found_any
				rows: List[Tuple[str, float, float]] = []
				for ln in pages_ocr[i].splitlines():
					cat, cur, ytd = line_to_category_and_numbers(ln)
					if cat is not None:
						rows.append((cat, cur, ytd))
				# If no provider_id yet, try TSV probe on corresponding page image
				if provider_id is None:
					# Determine page number from decoded page footer, map to image index
					page_num = _extract_page_number(pages_dec[i]) or (i + 1)
					img_index = max(0, min(len(page_images) - 1, page_num - 1))
					try:
						tsv_words = _run_tesseract_tsv(page_images[idx := img_index])
						# First try TSV numeric to the right of name
						provider_id = _find_provider_id_near_name_tsv(tsv_words, name)
						# If still none, try cropping a band to the right of name and OCR digits
						if provider_id is None:
							bbox = _get_name_bbox_from_tsv(tsv_words, name)
							if bbox:
								l, t, r, b = bbox
								cand_digits = _ocr_digits_from_crop(page_images[idx], (r + 5, max(0, t - 40), r + 800, b + 200))
								if cand_digits:
                                    # choose the longest token
									provider_id = sorted(cand_digits, key=len, reverse=True)[0]
					except Exception:
						pass
					# If still none, probe provider summary pages for this name
					if provider_id is None:
						summary_idxs = _find_provider_summary_pages(pages_dec, name)
						for si in summary_idxs:
							pn = _extract_page_number(pages_dec[si]) or (si + 1)
							simg = max(0, min(len(page_images) - 1, pn - 1))
							try:
								# First, try the decoded summary page directly (most reliable for IDs)
								if provider_id is None:
									found_id_dec = _extract_id_from_decoded_summary_text(pages_dec[si], name)
									if found_id_dec:
										provider_id = found_id_dec
										break
								# Try direct OCR-text extraction on the summary page first (map OCR by footer page number)
								ocr_idx = pn - 1
								ocr_page_text = pages_ocr[ocr_idx] if 0 <= ocr_idx < len(pages_ocr) else ""
								if not provider_id and ocr_page_text:
									found_id = _extract_id_from_ocr_summary_text(ocr_page_text, name)
									if found_id:
										provider_id = found_id
								tsv_words2 = _run_tesseract_tsv(page_images[simg])
								provider_id = _find_provider_id_near_name_tsv(tsv_words2, name)
								if provider_id is None:
									bbox2 = _get_name_bbox_from_tsv(tsv_words2, name)
									if bbox2:
										l2, t2, r2, b2 = bbox2
										cand2 = _ocr_digits_from_crop(page_images[simg], (r2 + 5, max(0, t2 - 40), r2 + 800, b2 + 200))
										if cand2:
											provider_id = sorted(cand2, key=len, reverse=True)[0]
								if provider_id:
									break
							except Exception:
								continue
				# Compute totals
				total_cur = None
				total_ytd = None
				for cat, cur, ytd in rows:
					if "TOTAL CLAIMS PAYABLE" in cat.upper():
						total_cur = cur
						total_ytd = ytd
						break
				if total_cur is None or total_ytd is None:
					total_cur = sum(r[1] for r in rows) if rows else 0.0
					total_ytd = sum(r[2] for r in rows) if rows else 0.0
				provider_entries.append({
					"name": name if provider_id is None else f"{name}",
					"rows": rows,
					"total_cur": total_cur,
					"total_ytd": total_ytd,
					"id": provider_id
				})
		text = ocr_text  # reference text
		pages = pages_ocr
	else:
		if args.source == "poppler-decode":
			poppler_txt = in_path.with_suffix("").with_name(in_path.stem + "_poppler.txt")
			extract_text_with_poppler(in_path, poppler_txt, layout=args.poppler_layout)
			raw_text = poppler_txt.read_text(encoding="utf-8", errors="ignore")
			text = decode_caesar_shift(raw_text, shift=29, low=19, high=94)
		else:
			text, _ = read_text_from_input(in_path, use_ocr=args.use_ocr, renderer=args.renderer, dpi=args.dpi, tess_args=tess_args)
		pages = text.split("\f")
		if not pages:
			print("No pages found in OCR text.", file=sys.stderr)
			sys.exit(2)
		grp_rows = parse_group_categories(pages)
		# Build provider entries from parsed pages
		provider_entries = []
		raw_provider_pages = pages[2:] if len(pages) > 2 else []
		for page in raw_provider_pages:
			name, rows = parse_provider_page(page)
			total_cur = None
			total_ytd = None
			for cat, cur, ytd in rows:
				if "TOTAL CLAIMS PAYABLE" in cat.upper():
					total_cur = cur
					total_ytd = ytd
					break
			if total_cur is None or total_ytd is None:
				total_cur = sum(r[1] for r in rows)
				total_ytd = sum(r[2] for r in rows)
			provider_entries.append({
				"name": name,
				"rows": rows,
				"total_cur": total_cur,
				"total_ytd": total_ytd
			})

	# Metadata from first page
	meta = parse_metadata(pages[0])
	(out_dir / "metadata.json").write_text(json.dumps(meta, indent=2), encoding="utf-8")

	# Group categories
	write_group_csvs(grp_rows, out_dir)
	# Providers
	write_provider_csvs_from_entries(provider_entries, out_dir)
	# Additional CSVs requested: detailed provider and group payments with repeated metadata columns
	if args.hybrid:
		try:
			write_provider_payments_csv(meta, provider_entries, pages_dec, out_dir)
		except Exception:
			# Be resilient; still produce core outputs
			pass
		try:
			write_group_payments_csv(meta, pages_dec, out_dir)
		except Exception:
			pass

	# Metadata from first page for readable export
	meta = parse_metadata(pages[0])
	(out_dir / "metadata.json").write_text(json.dumps(meta, indent=2), encoding="utf-8")
	# Combined readable text export
	write_combined_readable_text(meta, grp_rows, provider_entries, out_dir)

	# Also write the OCR text alongside for reference
	if args.hybrid:
		(out_dir / "hybrid_reference_ocr.txt").write_text(text, encoding="utf-8")
		(out_dir / "hybrid_reference_decoded.txt").write_text(decoded_text, encoding="utf-8")
	# And a cleaned OCR version
	if args.hybrid:
		(out_dir / "hybrid_reference_ocr_clean.txt").write_text(clean_ocr_text(text), encoding="utf-8")
		(out_dir / "hybrid_reference_decoded_clean.txt").write_text(clean_ocr_text(decoded_text), encoding="utf-8")
	else:
		ref_txt = out_dir / ("ocr_text.txt" if args.source == "ocr" else "decoded_text.txt")
		ref_txt.write_text(text, encoding="utf-8")
		ref_txt_clean = out_dir / ("ocr_text_clean.txt" if args.source == "ocr" else "decoded_text_clean.txt")
		ref_txt_clean.write_text(clean_ocr_text(text), encoding="utf-8")

	print("Parsed tables written to:", out_dir)
	print("Files created:")
	print(" -", out_dir / "metadata.json")
	print(" -", out_dir / "group_categories.csv")
	print(" -", out_dir / "provider_categories.csv")
	print(" -", out_dir / "provider_totals.csv")
	print(" -", out_dir / "combined_readable.txt")
	if args.hybrid:
		print(" -", out_dir / "hybrid_reference_ocr.txt")
		print(" -", out_dir / "hybrid_reference_ocr_clean.txt")
		print(" -", out_dir / "hybrid_reference_decoded.txt")
		print(" -", out_dir / "hybrid_reference_decoded_clean.txt")
	elif args.source == "ocr":
		print(" -", out_dir / "ocr_text.txt")
		print(" -", out_dir / "ocr_text_clean.txt")
	else:
		print(" -", out_dir / "decoded_text.txt")
		print(" -", out_dir / "decoded_text_clean.txt")


if __name__ == "__main__":
	main()


