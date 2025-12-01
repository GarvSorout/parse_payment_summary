#!/usr/bin/env python3
"""
PSP-safe album art embedder for MP3 and M4A.

What it does:
- For MP3: writes ID3v2.3 with a single APIC (Front cover) JPEG (baseline),
  removes existing APICs and APE tags to avoid PSP crashes.
- For M4A/MP4/AAC: writes a single covr atom with JPEG (baseline).
- Cover discovery order: <audio_basename>.jpg/.jpeg → cover.jpg → folder.jpg → any *.jpg in same folder.
- Optionally provide an explicit --cover image to use for all files.

Usage:
  python psp_embed_cover.py --input "/path/to/file/or/folder" [--cover cover.jpg] [--recurse]
"""
from __future__ import annotations

import argparse
from io import BytesIO
from pathlib import Path
from typing import Iterable

from PIL import Image

# Lazy import mutagen modules inside functions to allow --help without deps installed

SUPPORTED_AUDIO_EXTS = {".mp3", ".m4a", ".mp4", ".aac"}
JPEG_EXTS = {".jpg", ".jpeg"}


def make_psp_jpeg(src_bytes: bytes, max_dim: int = 300, quality: int = 85) -> bytes:
	"""Return a baseline JPEG resized to fit within max_dim x max_dim, stripped of metadata."""
	img = Image.open(BytesIO(src_bytes)).convert("RGB")
	# Resize with aspect ratio preserved
	img.thumbnail((max_dim, max_dim), Image.LANCZOS)
	out = BytesIO()
	img.save(out, format="JPEG", quality=quality, optimize=True, progressive=False)
	return out.getvalue()


def find_cover_for_audio(audio_path: Path) -> bytes | None:
	"""Locate a nearby JPEG cover file for a given audio, preferring specific filenames."""
	# Prefer same-name jpg/jpeg
	for ext in JPEG_EXTS:
		cand = audio_path.with_suffix(ext)
		if cand.exists():
			try:
				return cand.read_bytes()
			except Exception:
				pass
	# Then cover.jpg/folder.jpg
	for name in ("cover.jpg", "folder.jpg", "cover.jpeg", "folder.jpeg"):
		cand = audio_path.parent / name
		if cand.exists():
			try:
				return cand.read_bytes()
			except Exception:
				pass
	# Fallback to any jpg in folder
	for cand in audio_path.parent.glob("*.jp*g"):
		try:
			return cand.read_bytes()
		except Exception:
			continue
	return None


def embed_cover_mp3(mp3_path: Path, cover_bytes: bytes, max_dim: int, quality: int) -> None:
	from mutagen.id3 import ID3, ID3NoHeaderError, APIC
	from mutagen.apev2 import APEv2, error as APEError  # type: ignore
	# Remove APE tag if present (PSP can choke on APE)
	try:
		ape = APEv2(mp3_path)
		ape.delete()
	except (APEError, FileNotFoundError, Exception):
		pass
	# Ensure ID3 exists
	try:
		tags = ID3(mp3_path)
	except ID3NoHeaderError:
		tags = ID3()
	# Remove any existing APIC frames
	for key in list(tags.keys()):
		if key.startswith("APIC"):
			del tags[key]
	jpg = make_psp_jpeg(cover_bytes, max_dim=max_dim, quality=quality)
	tags.add(APIC(encoding=3, mime="image/jpeg", type=3, desc="Cover (front)", data=jpg))
	# Save as ID3v2.3
	tags.save(mp3_path, v2_version=3)


def embed_cover_m4a(m4a_path: Path, cover_bytes: bytes, max_dim: int, quality: int) -> None:
	from mutagen.mp4 import MP4, MP4Cover
	mp4 = MP4(m4a_path)
	jpg = make_psp_jpeg(cover_bytes, max_dim=max_dim, quality=quality)
	mp4["covr"] = [MP4Cover(jpg, imageformat=MP4Cover.FORMAT_JPEG)]
	mp4.save()


def iter_audio_files(root: Path, recurse: bool) -> Iterable[Path]:
	if root.is_file():
		if root.suffix.lower() in SUPPORTED_AUDIO_EXTS:
			yield root
		return
	if recurse:
		for p in root.rglob("*"):
			if p.is_file() and p.suffix.lower() in SUPPORTED_AUDIO_EXTS:
				yield p
	else:
		for p in root.glob("*"):
			if p.is_file() and p.suffix.lower() in SUPPORTED_AUDIO_EXTS:
				yield p


def process_audio(audio_path: Path, explicit_cover: bytes | None, max_dim: int, quality: int, quiet: bool) -> bool:
	cover = explicit_cover or find_cover_for_audio(audio_path)
	if cover is None:
		if not quiet:
			print(f"[skip:no-cover] {audio_path}")
		return False
	try:
		ext = audio_path.suffix.lower()
		if ext == ".mp3":
			embed_cover_mp3(audio_path, cover, max_dim=max_dim, quality=quality)
		elif ext in {".m4a", ".mp4", ".aac"}:
			embed_cover_m4a(audio_path, cover, max_dim=max_dim, quality=quality)
		else:
			if not quiet:
				print(f"[skip:unsupported] {audio_path}")
			return False
		if not quiet:
			print(f"[ok] {audio_path}")
		return True
	except Exception as e:
		if not quiet:
			print(f"[error] {audio_path}: {e}")
		return False


def main() -> None:
	parser = argparse.ArgumentParser(description="Embed PSP-safe album art into MP3/M4A files.")
	parser.add_argument("--input", required=True, help="Audio file or directory to process")
	parser.add_argument("--cover", help="Optional cover image (JPEG/PNG). Applied to all files if provided.")
	parser.add_argument("--recurse", action="store_true", help="Recurse into subdirectories when input is a folder")
	parser.add_argument("--max-dim", type=int, default=300, help="Max cover dimension (pixels). Default: 300")
	parser.add_argument("--quality", type=int, default=85, help="JPEG quality (1-95). Default: 85")
	parser.add_argument("--quiet", action="store_true", help="Reduce stdout output")
	args = parser.parse_args()

	root = Path(args.input).expanduser().resolve()
	if not root.exists():
		raise SystemExit(f"Input not found: {root}")

	explicit_cover: bytes | None = None
	if args.cover:
		cp = Path(args.cover).expanduser().resolve()
		if not cp.exists():
			raise SystemExit(f"Cover image not found: {cp}")
		explicit_cover = cp.read_bytes()

	total = 0
	ok = 0
	for audio in iter_audio_files(root, recurse=args.recurse):
		total += 1
		if process_audio(audio, explicit_cover, max_dim=args.max_dim, quality=args.quality, quiet=args.quiet):
			ok += 1
	if not args.quiet:
		print(f"Done. Updated {ok}/{total} files.")


if __name__ == "__main__":
	main()


