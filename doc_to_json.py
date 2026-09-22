#!/usr/bin/env python3
"""
Extract QUESTION CONTENT from .docx files into nested JSON.

Handles two document shapes:
  1. Simple: each question is its own small [Qno, Title] Word table (flat output).
  2. Nested: a digit-level question's table cell contains a full block of text
     with category headers (A/B/C/D/E...), Roman-numeral subsections, and
     numbered sub-questions all mixed together as plain paragraph text.

Category header lines (e.g. "A. LISTENING SKILLS (15 marks)") are dropped.
Embedded data tables (fill-in-the-blank grids, substitution tables) are
skipped - only their preceding instruction line is kept.
Ambiguous Roman-numeral detections are flagged for manual review rather
than guessed silently.

Image URLs are S3 presigned links that expire ~900s (15 min) after being
issued -- baking one into the JSON at conversion time and leaving it there
is what causes broken images in the frontend later on. Re-run this script
in refresh mode any time before the JSON is actually displayed:

    python doc_to_json.py --refresh path/to/output.json
    python doc_to_json.py --refresh path/to/output_dir   # refreshes every *.json in it

This re-signs each image's URL from its durable S3 key ("filePath") without
touching question text, and can be run as often as needed (e.g. on a timer,
or right before serving the JSON to the frontend).
"""

import json
import os
import re
import struct
import sys
import zipfile
import zlib
from pathlib import Path

from dotenv import load_dotenv

try:
    import requests
except ImportError:
    requests = None

try:
    from docx2python import docx2python
except ImportError:
    sys.exit("Missing dependency. Install it with:\n    pip install docx2python --break-system-packages")

try:
    import docx as _pydocx  # python-docx -- used only to detect embedded equation objects
except ImportError:
    _pydocx = None

try:
    from PIL import Image as _PILImage  # used to convert unsupported image formats before upload
except ImportError:
    _PILImage = None

try:
    import imagecodecs as _imagecodecs  # decodes JPEG XR (.wdp) - PIL alone can't read it
except ImportError:
    _imagecodecs = None


# ============================================================================
# SET THESE TWO PATHS AND RUN THE SCRIPT WITH NO ARGUMENTS
# ============================================================================
INPUT_PATH = r"C:\Users\Administrator\Desktop\DocQues\files"
OUTPUT_PATH = r"C:\Users\Administrator\Desktop\DocQues\res"
# ============================================================================

load_dotenv()
BEARER_TOKEN = os.getenv("BEARER_TOKEN")
# File Service (sfm) -- confirmed working, unlike the old QuestionBank
# UploadQuestionImages route which the server rejects with a permission
# error regardless of client (verified directly via Swagger).
SFM_UPLOAD_URL = os.getenv("SFM_UPLOAD_URL", "https://campus.psgtech.ac.in/sfm/api/files/upload")
SFM_VIEW_URL = os.getenv("SFM_VIEW_URL", "https://campus.psgtech.ac.in/sfm/api/files/viewfileurl")
# Must already exist in AWS -- FILE_SERVICE's own docs note this route
# does not create the path, only uploads into an existing one.
SFM_UPLOAD_PATH = os.getenv("SFM_UPLOAD_PATH", "/DOTECOE/atlas-9f3c7a1d")
# The upload API silently rejects any file under this size -- see
# pad_image_to_min_size() below for how small embedded images are handled.
MIN_UPLOAD_IMAGE_BYTES = int(os.getenv("MIN_UPLOAD_IMAGE_BYTES", "2048"))
# Extensions the upload API is known to accept as-is. Anything else gets
# converted to PNG first -- see convert_unsupported_image_format() below.
UPLOAD_SUPPORTED_EXTENSIONS = {".jpg", ".jpeg", ".png", ".gif", ".bmp", ".webp"}


CATEGORY_HEADER_RE = re.compile(
    r'^[A-E]\)?\.?\s*[A-Za-z][A-Za-z ]*SKILLS?\s*\(\s*\d+\s*[Mm]arks?\)?\s*$'
)
# Leftover fragment of a category label with nothing after it, e.g. "D) "
# - clearly a stray artifact, not real content.
STRAY_LETTER_RE = re.compile(r'^[A-E]\)?\.?\s*$')
# Longest tokens first so alternation prefers e.g. "VIII" over "V".
ROMAN_HEADING_RE = re.compile(
    r'^(VIII|III|VII|II|IV|VI|IX|I|V|X)\b[).]?\s*(.*)$'
)
SUBQ_RE = re.compile(r'^(\d+)[).]\s+(.*)$')
# Finds the real course-code token in a filename regardless of separator
# style (space- or underscore-based) and skips long numeric upload-ID
# prefixes (e.g. "1789552055228_MA232431_-_Theory_Question" -> "MA232431",
# not the 13-digit prefix and not the whole descriptive filename).
SUBJ_CODE_RE = re.compile(r'(?<!\d)(?:[A-Za-z]{1,3}\d{5,7}|\d{9,10})(?!\d)')


def extract_subj_code(stem: str) -> str:
    """Pull just the course-code token out of a document's filename stem."""
    m = SUBJ_CODE_RE.search(stem)
    if m:
        return m.group(0)
    # Fallback: previous behavior, in case a filename doesn't match the
    # expected code patterns at all.
    return stem.split(" ", 1)[0]
# docx2python inserts a placeholder like "----media/image1.emf----" (or .png,
# .jpg, etc.) into extracted text wherever an embedded image sits in a cell.
IMAGE_MARKER_RE = re.compile(r'----media/(image\d+\.\w+)----')
# A single-letter sub-answer label ("A.", "B)", etc.) as a whole table cell.
LETTER_LABEL_RE = re.compile(r'^[A-Ea-e]\)?\.?$')
# The marks-allocation grid's header row always starts with this cell,
# even when no preceding "ALLOCATION OF MARKS" paragraph is present.
MARKS_TABLE_HEADER_RE = re.compile(r'^S\.?\s*NO\.?$', re.IGNORECASE)

ROMAN_VALUES = {"I": 1, "II": 2, "III": 3, "IV": 4, "V": 5,
                "VI": 6, "VII": 7, "VIII": 8, "IX": 9, "X": 10}
ROMAN_INT_TO_STR = {v: k for k, v in ROMAN_VALUES.items()}


def _int_to_roman(n):
    return ROMAN_INT_TO_STR.get(n, str(n))


def clean_cell(paragraphs):
    if paragraphs is None:
        return ""
    if isinstance(paragraphs, str):
        text = paragraphs
    else:
        text = "\n".join(str(p) for p in paragraphs if p)
    return text.replace("\t", " ").strip()


def _lineage_says_table(par):
    lineage = getattr(par, "lineage", None)
    return bool(lineage) and len(lineage) > 1 and lineage[1] == "tbl"


def _is_real_table(pars_table):
    for row in pars_table:
        for cell in row:
            for par in cell:
                if _lineage_says_table(par):
                    return True
    return False


def _flatten_entry(text_table):
    """Join every cell of a body entry into one newline-joined string."""
    lines = []
    for row in text_table:
        for cell in row:
            t = clean_cell(cell)
            if t:
                lines.append(t)
    return "\n".join(lines)


def _flatten_table(rows):
    """Convert a Word table into a readable text block while keeping row structure."""
    lines = []
    for row in rows:
        cells = [clean_cell(cell) for cell in row]
        cells = [cell for cell in cells if cell]
        if cells:
            lines.append(" | ".join(cells))
    return "\n".join(lines)


def extract_embedded_media(docx_path: Path, out_dir: Path):
    """Extract word/media/* files from a DOCX into a sidecar folder.

    The extracted filenames keep the same basename as the DOCX media names so
    placeholders like ----media/image1.jpeg---- can be resolved later.
    """
    media_dir = out_dir / f"{docx_path.stem}_media"
    media_dir.mkdir(parents=True, exist_ok=True)

    extracted = []
    try:
        with zipfile.ZipFile(docx_path) as archive:
            for name in archive.namelist():
                if not name.startswith("word/media/") or name.endswith("/"):
                    continue
                target = media_dir / Path(name).name
                with archive.open(name) as src, open(target, "wb") as dst:
                    dst.write(src.read())
                extracted.append(target.name)
    except Exception:
        return None, []

    return media_dir.name, extracted


def _pad_jpeg(data: bytes, target_size: int) -> bytes:
    """
    Insert a standard JPEG COM (comment) marker segment -- spec-compliant
    and silently skipped by every decoder -- right before the End-Of-Image
    marker so the file's total byte size reaches target_size. Not one
    pixel of actual image data is touched.
    """
    needed = target_size - len(data)
    if needed <= 0 or not data.endswith(b"\xff\xd9"):
        return data  # already big enough, or not a well-formed JPEG - don't risk it

    overhead = 4  # FF FE marker + 2-byte length field
    payload_len = max(needed - overhead, 0)
    segments = b""
    remaining = payload_len
    while remaining > 0:
        # COM segment length field is 2 bytes and includes itself, so max
        # payload per segment is 65533; loop (harmless, just belt-and-braces
        # since we're only ever padding up to a couple KB in practice).
        chunk = min(remaining, 65533)
        segments += b"\xff\xfe" + struct.pack(">H", chunk + 2) + (b"\x00" * chunk)
        remaining -= chunk
    return data[:-2] + segments + data[-2:]


def _pad_png(data: bytes, target_size: int) -> bytes:
    """
    Insert a standard PNG tEXt ancillary chunk -- ignored by every PNG
    reader -- right before IEND so the file's total byte size reaches
    target_size. Not one pixel of actual image data is touched.
    """
    needed = target_size - len(data)
    if needed <= 0 or not data.endswith(b"IEND\xae\x42\x60\x82"):
        return data  # already big enough, or not a well-formed PNG - don't risk it

    overhead = 12 + 8  # 4-byte len + 4-byte type + 4-byte CRC, plus "Padding\x00" key
    payload_len = max(needed - overhead, 1)
    chunk_type = b"tEXt"
    chunk_data = b"Padding\x00" + (b"0" * payload_len)
    chunk = (
        struct.pack(">I", len(chunk_data)) + chunk_type + chunk_data
        + struct.pack(">I", zlib.crc32(chunk_type + chunk_data) & 0xFFFFFFFF)
    )
    iend_start = len(data) - 12
    return data[:iend_start] + chunk + data[iend_start:]


def pad_image_to_min_size(path: Path, min_bytes: int = MIN_UPLOAD_IMAGE_BYTES) -> bool:
    """
    The upload API rejects any file under ~2KB outright. Some genuinely
    tiny embedded images (small cropped screenshots, equation snapshots)
    fall under that floor. Rather than upscaling/resampling them - which
    would change actual pixel content for no real reason and isn't
    guaranteed to land above the threshold predictably - this pads the
    file with a standard, universally-ignored metadata segment/chunk so
    its byte size clears the API's floor while every pixel stays
    byte-for-byte identical.

    Mutates the file on disk in place (safe: this only ever runs on our
    own extracted scratch copy under the media dir, never the source
    .docx). Returns True if the file is now >= min_bytes (either it
    already was, or padding succeeded), False if it's a format this
    function doesn't know how to safely pad -- the caller should warn
    rather than upload something the API will just bounce anyway.
    """
    try:
        data = path.read_bytes()
    except Exception:
        return False

    if len(data) >= min_bytes:
        return True

    if data.startswith(b"\xff\xd8\xff"):
        padded = _pad_jpeg(data, min_bytes)
    elif data.startswith(b"\x89PNG\r\n\x1a\n"):
        padded = _pad_png(data, min_bytes)
    else:
        return False

    if len(padded) < min_bytes:
        return False

    path.write_bytes(padded)
    return True


def convert_unsupported_image_format(path: Path):
    """
    The upload API rejects some formats Word happily embeds -- most
    commonly hdphotoN.wdp, a JPEG XR fallback bitmap Word writes
    alongside a picture's primary format for older-Office compatibility
    (the "type '.wdp' is not supported" error). Rather than dropping
    these images, decode them and re-save as PNG, which the API accepts.

    Tries Pillow first (covers the common raster formats it ships with),
    then imagecodecs' generic reader as a fallback (an additional ~89
    raster codecs, including JPEG XR/.wdp, which Pillow alone can't
    read). Neither can help with vector formats (.emf/.wmf/.svg) --
    those aren't pixel data to decode, they're stored drawing
    instructions, so they'd need a rendering engine (e.g. LibreOffice
    headless, Inkscape) rather than a codec; such files get a specific
    warning saying so instead of a generic failure.

    Returns the path to a PNG version of the image (a new file next to
    the original; the original is left alone) on success, or None if the
    format is already supported (no conversion needed) or couldn't be
    decoded (caller should warn and skip rather than upload a file the
    API will just bounce).
    """
    ext = path.suffix.lower()
    if ext in UPLOAD_SUPPORTED_EXTENSIONS:
        return None  # already fine, no conversion needed

    out_path = path.with_suffix(".png")

    def _save_array_as_png(arr):
        img = _PILImage.fromarray(arr)
        if img.mode not in ("RGB", "RGBA", "L"):
            img = img.convert("RGB")
        img.save(out_path, "PNG")
        return out_path

    pil_error = None
    if _PILImage is not None:
        try:
            img = _PILImage.open(path)
            img.load()
            if img.mode not in ("RGB", "RGBA", "L"):
                img = img.convert("RGB")
            img.save(out_path, "PNG")
            return out_path
        except Exception as e:
            pil_error = e  # fall through to imagecodecs below

    if _imagecodecs is not None and _PILImage is not None:
        try:
            arr = _imagecodecs.imread(path.read_bytes())
            return _save_array_as_png(arr)
        except Exception as e:
            if ext in (".emf", ".wmf", ".svg"):
                print(f"    [WARN] {path.name} is a vector format ('{ext}') -- neither Pillow "
                      f"nor imagecodecs can rasterize vector drawings (they only decode pixel "
                      f"data); a rendering engine like LibreOffice or Inkscape would be needed "
                      f"to convert it -- skipping")
                return None
            print(f"    [WARN] couldn't convert unsupported format {path.name} ({ext}): {e}")
            return None

    if _PILImage is None:
        print(f"    [WARN] {path.name} has unsupported format '{ext}' and Pillow isn't "
              f"installed to attempt a conversion -- install with:\n"
              f"        pip install Pillow imagecodecs --break-system-packages")
    else:
        print(f"    [WARN] {path.name} has unsupported format '{ext}'; Pillow couldn't read it "
              f"({pil_error}) and 'imagecodecs' isn't installed to try further -- install with:\n"
              f"        pip install imagecodecs --break-system-packages")
    return None


def upload_images_to_s3(image_paths, docx_stem):
    """
    Uploads each local image individually via the File Service
    (sfm/api/files/upload), then fetches an initial viewable URL for it
    via ViewFileUrl. Returns {local_filename: {"filePath": s3_key,
    "imageUrl": url}}.

    ASSUMPTIONS (please verify / correct if wrong):
    - Auth: FILE_SERVICE's docs say "Send RBAC token in header as
      <bearer RBAC Token>" -- assumed to be the same BEARER_TOKEN used
      for SaveQuestion elsewhere in this project, not a separate
      credential. If uploads fail with an auth-looking error, this is
      the first thing to check.
    - The upload field is named "file" (singular) -- FILE_SERVICE.docx
      says 'file : Send Iform File' for this route.
    - SFM_UPLOAD_PATH must already exist in AWS -- FILE_SERVICE's own
      docs note this route does NOT create the path, only uploads into
      an existing one.
    - SFM_UPLOAD_PATH is ONE shared folder across every document in this
      batch (not per-subject), so two different documents each
      extracting an "image1.png" would silently overwrite each other in
      S3 if uploaded under their raw filename. Each upload is prefixed
      with its source document's filename stem (e.g.
      "1012234420_image1.png") to keep them unique.
    - The returned "imageUrl" here is only an initial, short-lived
      preview (ViewFileUrl issues presigned URLs, same ~900s expiry
      pattern as before) -- run this script with --refresh right before
      actual use, same as with the old endpoint. The durable value that
      matters long-term is "filePath" (the S3 key).
    - The API rejects files under ~2KB, which some genuinely tiny
      embedded images fall under -- see pad_image_to_min_size(), called
      per-image below, which pads such files (without touching pixel
      data) rather than skipping them.
    - The API also rejects formats it doesn't recognize -- most commonly
      Word's hdphotoN.wdp fallback bitmaps (JPEG XR). See
      convert_unsupported_image_format(), called per-image below, which
      converts these to PNG before upload rather than skipping them.
    """
    if not image_paths or requests is None:
        return {}

    headers = {}
    if BEARER_TOKEN:
        headers["Authorization"] = f"Bearer {BEARER_TOKEN}"

    mapping = {}
    for local_path in image_paths:
        upload_path = local_path
        if local_path.suffix.lower() not in UPLOAD_SUPPORTED_EXTENSIONS:
            converted_path = convert_unsupported_image_format(local_path)
            if converted_path is None:
                # convert_unsupported_image_format() already printed a
                # [WARN] explaining why -- don't waste a request on a
                # format the API is guaranteed to reject.
                continue
            upload_path = converted_path

        if not pad_image_to_min_size(upload_path):
            print(f"    [WARN] {upload_path.name} is under {MIN_UPLOAD_IMAGE_BYTES} bytes "
                  f"and couldn't be padded (unrecognized format) -- the upload API will "
                  f"likely reject it")

        unique_name = f"{docx_stem}_{upload_path.name}"
        params = {"Path": SFM_UPLOAD_PATH, "Filename": unique_name}

        try:
            with open(upload_path, "rb") as fh:
                resp = requests.post(
                    SFM_UPLOAD_URL, params=params,
                    files={"file": (unique_name, fh)},
                    headers=headers, timeout=120,
                )
            resp_body = None
            try:
                resp_body = resp.json()
            except ValueError:
                pass

            already_exists = (
                resp.status_code == 400 and isinstance(resp_body, dict)
                and "already exists" in str(resp_body.get("message", "")).lower()
            )
            if already_exists:
                # The image genuinely was already uploaded in a prior run
                # (S3 uploads are permanent even if our local tracking
                # JSON lost track of it, e.g. an intervening failed run
                # overwrote it). Reconstruct the deterministic key
                # ourselves instead of treating this as a failure.
                key = f"{SFM_UPLOAD_PATH.strip('/')}/{unique_name}"
                print(f"    [REUSE] {upload_path.name} already exists in S3 at {key}, reusing it")
            elif not resp.ok:
                print(f"    [WARN] image upload failed for {upload_path.name}: "
                      f"{resp.status_code} -- {resp.text[:300]}")
                continue
            else:
                result = resp_body or {}
                if result.get("isError"):
                    print(f"    [WARN] image upload reported error for {upload_path.name}: "
                          f"{result.get('message')}")
                    continue
                key = (result.get("result") or {}).get("key")
                if not key:
                    print(f"    [WARN] upload succeeded but no 'key' returned for {upload_path.name}")
                    continue
        except Exception as e:
            print(f"    [WARN] image upload failed for {upload_path.name}: {e}")
            continue

        image_url = ""
        try:
            view_resp = requests.get(SFM_VIEW_URL, params={"filepath": key}, headers=headers, timeout=30)
            if view_resp.ok:
                view_result = view_resp.json()
                if not view_result.get("isError"):
                    image_url = (view_result.get("result") or {}).get("url", "")
        except Exception as e:
            print(f"    [WARN] could not fetch initial view URL for {upload_path.name}: {e}")

        # Keyed by the ORIGINAL filename (e.g. "hdphoto1.wdp"), not the
        # converted one, since that's what the "----media/...----"
        # placeholders in the extracted text reference.
        mapping[local_path.name] = {"filePath": key, "imageUrl": image_url}

    return mapping


def get_body_entries(docx_path: Path):
    """Returns list of (is_real, rows) for every body entry in the doc."""
    with docx2python(str(docx_path)) as doc:
        body = doc.body
        try:
            body_pars = doc.body_pars
        except AttributeError:
            body_pars = None

        entries = []
        for idx, text_table in enumerate(body):
            if body_pars is not None:
                is_real = _is_real_table(body_pars[idx])
            else:
                is_real = not (len(text_table) == 1 and len(text_table[0]) == 1)
            entries.append((is_real, text_table))
        return entries


def _digit_heading_info(is_real, rows):
    """
    Returns (is_heading, digit, initial_buffer_lines) for an entry that
    marks the start of a new digit-numbered question. Handles two shapes:
      - classic: [digit, description] pair in one table row (description
        may continue across further rows with a blank first cell)
      - bare: a digit completely alone as its own entry, with no inline
        description -- the description arrives via later entries instead
        (seen in docs that also split "A."/"B." sub-answers out as
        separate table/paragraph entries rather than one flat cell)
    """
    if len(rows) < 1 or len(rows[0]) < 1:
        return False, None, []
    cell0 = clean_cell(rows[0][0]).rstrip(".").strip()
    if not cell0.isdigit():
        return False, None, []

    if len(rows[0]) >= 2:
        for row in rows[1:]:
            if len(row) >= 1 and clean_cell(row[0]).strip():
                return False, None, []
        buffer_lines = [clean_cell(rows[0][1])]
        for extra_row in rows[1:]:
            if len(extra_row) >= 2:
                t = clean_cell(extra_row[1])
                if t:
                    buffer_lines.append(t)
        return True, cell0, buffer_lines

    if len(rows) == 1 and len(rows[0]) == 1:
        return True, cell0, []

    return False, None, []


def _is_digit_heading(is_real, rows):
    """Backwards-compatible bool wrapper around _digit_heading_info."""
    is_heading, _, _ = _digit_heading_info(is_real, rows)
    return is_heading


def _is_lettered_subanswer_table(rows):
    """A small 2-col table like [['A.', 'Draw the ...']] -- a lettered
    sub-answer (A/B/C/D/E choice), not a data/marks grid. These are real
    question content and should be absorbed, unlike other genuine tables."""
    return (
        len(rows) == 1
        and len(rows[0]) == 2
        and bool(LETTER_LABEL_RE.match(clean_cell(rows[0][0]).strip()))
        and bool(clean_cell(rows[0][1]).strip())
    )


def _is_stop_marker(rows):
    """True for the 'ALLOCATION OF MARKS' paragraph or the marks-grid
    table's own header row (which sometimes appears with no preceding
    'ALLOCATION OF MARKS' label) -- either ends the current digit block."""
    if not rows or not rows[0]:
        return False
    first_cell = clean_cell(rows[0][0]).strip()
    if first_cell.upper().startswith("ALLOCATION OF MARKS"):
        return True
    if MARKS_TABLE_HEADER_RE.match(first_cell):
        return True
    return False


def collect_digit_blocks(entries):
    """
    Walks body entries and groups them into (digit, raw_text) blocks.
    A digit block starts at either a [digit, text] heading or a bare
    digit-alone entry, and absorbs every following entry's text --
    loose paragraphs, and lettered A/B/C sub-answer tables -- until
    either the next digit heading or a stop marker (ALLOCATION OF MARKS
    / the marks-grid header) is hit. Other genuine tables (fill-in
    grids, substitution tables) are skipped entirely - their content
    never enters the text stream, matching the original design.
    """
    blocks = []
    i = 0
    n = len(entries)
    while i < n:
        is_real, rows = entries[i]
        is_heading, digit, initial_buffer = _digit_heading_info(is_real, rows)
        if is_heading:
            buffer_lines = list(initial_buffer)
            i += 1
            while i < n:
                is_real2, rows2 = entries[i]
                is_heading2, _, _ = _digit_heading_info(is_real2, rows2)
                if is_heading2:
                    break
                if _is_stop_marker(rows2):
                    i += 1
                    break
                if not is_real2:
                    flat = _flatten_entry(rows2)
                    if flat.strip():
                        buffer_lines.append(flat)
                    i += 1
                elif _is_lettered_subanswer_table(rows2):
                    label = clean_cell(rows2[0][0]).strip()
                    desc = clean_cell(rows2[0][1]).strip()
                    buffer_lines.append(f"{label} {desc}")
                    i += 1
                else:
                    # Keep genuine tables instead of dropping them so output
                    # tabulations survive into the JSON.
                    flat_table = _flatten_table(rows2)
                    if flat_table.strip():
                        buffer_lines.append(flat_table)
                    i += 1
            blocks.append((digit, "\n".join(buffer_lines)))
        else:
            i += 1
    return blocks


def parse_digit_block(digit, raw_text, flags):
    """
    Parses one digit's raw text into a nested question node:
      {"Qno": digit, "Description": "..."}                    (flat, no romans found)
    or
      {"Qno": digit, "SubDivisions": [...]}                    (roman subsections found)
    Category header lines are dropped. Numbered sub-questions within a
    Roman section are split into that section's "SubQuestions" list.
    """
    lines = [l.rstrip() for l in raw_text.split("\n")]

    sections = []          # list of dicts: {"roman":..., "lines":[...]}
    preamble_lines = []    # any text before the first accepted Roman heading
    just_saw_category_header = False
    seen_any_roman = False
    next_position = 1      # expected roman position (1=I, 2=II, ...) for fallback matching

    for line in lines:
        stripped = line.strip()
        if not stripped:
            continue

        if CATEGORY_HEADER_RE.match(stripped):
            just_saw_category_header = True
            continue

        if STRAY_LETTER_RE.match(stripped):
            # Leftover fragment (e.g. "D) ") - skip without disturbing
            # whether we "just saw" a real category header.
            continue

        m = ROMAN_HEADING_RE.match(stripped)
        if m:
            token, rest = m.group(1), m.group(2)
            accept = False
            if token != "I":
                # Non-"I" Roman tokens never collide with real English
                # words, so they're trusted whenever matched.
                accept = True
            else:
                # "I" collides with the pronoun - only trust it as a
                # heading if it's the very first Roman token in this
                # digit, or immediately follows a category header line
                # (a genuine subsection restart, e.g. section D resets
                # numbering to I after "D READING SKILLS...").
                if not seen_any_roman or just_saw_category_header:
                    accept = True

            if accept:
                is_expected_reset = (not seen_any_roman) or just_saw_category_header
                token_value = ROMAN_VALUES[token]
                if not is_expected_reset and token_value != next_position:
                    flags.append({
                        "digit": digit,
                        "line": stripped,
                        "reason": f"Roman numeral '{token}' found where '{_int_to_roman(next_position)}' "
                                  f"was expected next - possible typo in the source document's own "
                                  f"numbering (not a parsing ambiguity). Verify the sequence is correct.",
                    })
                seen_any_roman = True
                just_saw_category_header = False
                next_position = token_value + 1
                sections.append({"roman": token, "lines": [rest] if rest else []})
                continue
            else:
                flags.append({
                    "digit": digit,
                    "line": stripped,
                    "reason": "Line starts with 'I' but is not the first Roman "
                              "heading and doesn't follow a category header - "
                              "treated as body text, not a new subsection. "
                              "Verify this is correct.",
                })
                # fall through - treated as ordinary text below

        else:
            # Fallback: some documents use a plain Arabic number instead of
            # a Roman numeral for a section heading (a real inconsistency
            # found in this template family). Only trust this as a genuine
            # new section - not an ordinary sub-question - when it
            # immediately follows a category header line, mirroring the
            # same safety gate used for the ambiguous "I" case above.
            am = SUBQ_RE.match(stripped)
            if am and just_saw_category_header:
                roman_token = _int_to_roman(next_position)
                seen_any_roman = True
                just_saw_category_header = False
                next_position += 1
                sections.append({"roman": roman_token, "lines": [am.group(2)] if am.group(2) else []})
                flags.append({
                    "digit": digit,
                    "line": stripped,
                    "reason": f"Section heading used Arabic numeral '{am.group(1)}' instead of "
                              f"a Roman numeral in the source document - normalized to "
                              f"'{roman_token}' since it directly follows a category header. "
                              f"Verify this is correct.",
                })
                continue

        just_saw_category_header = False
        if sections:
            sections[-1]["lines"].append(stripped)
        else:
            preamble_lines.append(stripped)

    if not sections:
        # No Roman subsections at all - simple flat question.
        description = "\n".join(preamble_lines).strip()
        return {"Qno": digit, "Description": description}

    sub_divisions = []
    for sec in sections:
        roman = sec["roman"]
        body_lines = sec["lines"]

        sub_questions = []
        description_lines = []
        for bl in body_lines:
            sm = SUBQ_RE.match(bl)
            if sm:
                sub_questions.append({
                    "Qno": sm.group(1),
                    "ParentQuestionId": roman,
                    "Description": sm.group(2).strip(),
                })
            else:
                if not sub_questions:
                    description_lines.append(bl)
                # else: stray non-numbered line after subquestions started -
                # append to the last subquestion's description as continuation
                elif sub_questions:
                    sub_questions[-1]["Description"] += "\n" + bl

        node = {
            "Qno": roman,
            "ParentQuestionId": digit,
            "Description": "\n".join(description_lines).strip(),
        }
        if sub_questions:
            node["SubQuestions"] = sub_questions
        sub_divisions.append(node)

    return {"Qno": digit, "SubDivisions": sub_divisions}


def extract_questions_from_docx(docx_path: Path):
    entries = get_body_entries(docx_path)
    blocks = collect_digit_blocks(entries)
    flags = []
    questions = [parse_digit_block(digit, raw_text, flags) for digit, raw_text in blocks]
    return questions, flags


def has_omml_equations(docx_path: Path) -> bool:
    """
    True if the document contains real embedded Word/MathType equation
    objects (m:oMath). Plain text extraction of these silently produces
    wrong/ambiguous output (e.g. a fraction 1/2 flattens to "12" with the
    fraction bar lost) -- not just ugly, actually incorrect -- so files
    containing them are skipped for now rather than risking bad content.
    Typed Unicode math symbols (x², π, ≤ etc.) are NOT affected by this
    check and are extracted normally.
    """
    if _pydocx is None:
        return False  # can't check; don't block processing over it
    try:
        d = _pydocx.Document(str(docx_path))
        return "<m:oMath" in d.element.xml
    except Exception:
        return False


def has_embedded_images(docx_path: Path) -> bool:
    """
    True if the document contains real embedded pictures (w:drawing / a:blip
    relationships) -- not docx2python's after-the-fact text placeholder.
    Checking the raw XML directly catches cases where an image is the ONLY
    content in a question cell (extracted text ends up blank, which
    otherwise gets misreported as "no questions found" rather than "image").
    """
    if _pydocx is None:
        return False
    try:
        d = _pydocx.Document(str(docx_path))
        xml = d.element.xml
        return "<w:drawing" in xml or "<a:blip" in xml
    except Exception:
        return False


def substitute_image_placeholders(questions, url_by_filename):
    """
    Recursively walks every Description across Questions/SubDivisions/
    SubQuestions and replaces each "----media/imageN.ext----" placeholder
    in-place with a real <img> tag pointing at the uploaded S3 URL, so the
    raw placeholder text never reaches the database. A placeholder whose
    image failed to upload (not in url_by_filename) is left as-is rather
    than silently deleted, so the missing image stays visible for review
    instead of vanishing without a trace.
    """
    def repl(m):
        fname = m.group(1)
        url = url_by_filename.get(fname)
        if not url:
            return m.group(0)
        return f'<img src="{url}" alt="{fname}" />'

    def process_node(node):
        if node.get("Description") is not None:
            node["Description"] = IMAGE_MARKER_RE.sub(repl, node["Description"])
        for child in node.get("SubDivisions", []) or []:
            process_node(child)
        for child in node.get("SubQuestions", []) or []:
            process_node(child)

    for q in questions:
        process_node(q)


def load_existing_media_map(json_path: Path) -> dict:
    """Reads a prior run's output JSON (if present) and returns
    {fileName: {"filePath":..., "imageUrl":...}} for any image that was
    already successfully uploaded, so a re-run can reuse it instead of
    uploading a duplicate."""
    if not json_path.exists():
        return {}
    try:
        with open(json_path, "r", encoding="utf-8") as f:
            data = json.load(f)
    except Exception:
        return {}
    mapping = {}
    for entry in data.get("_embedded_media") or []:
        fname = entry.get("fileName")
        if fname and entry.get("filePath"):
            mapping[fname] = {"filePath": entry["filePath"], "imageUrl": entry.get("imageUrl", "")}
    return mapping


def fetch_view_url(file_path_key: str) -> str:
    """Fetches a fresh viewable URL for an already-uploaded S3 key via
    ViewFileUrl -- used when reusing a prior upload, since the old cached
    imageUrl from a previous run has likely already expired."""
    if requests is None:
        return ""
    headers = {}
    if BEARER_TOKEN:
        headers["Authorization"] = f"Bearer {BEARER_TOKEN}"
    try:
        resp = requests.get(SFM_VIEW_URL, params={"filepath": file_path_key}, headers=headers, timeout=30)
        if resp.ok:
            result = resp.json()
            if not result.get("isError"):
                return (result.get("result") or {}).get("url", "")
    except Exception:
        pass
    return ""


def refresh_baked_urls(questions, url_by_basename: dict):
    """
    Swaps a freshly-issued presigned URL into every already-baked
    <img src="..."> tag across Questions/SubDivisions/SubQuestions, keyed
    by the S3 key's basename (e.g. "1020235541_image1.jpeg"), which stays
    constant across re-signs -- only the query string (signature/expiry)
    changes. This is what makes refreshing a previously generated output
    JSON possible without re-parsing the source .docx.
    """
    if not url_by_basename:
        return

    patterns = [
        (re.compile(r'src="[^"]*' + re.escape(basename) + r'[^"]*"'), f'src="{fresh_url}"')
        for basename, fresh_url in url_by_basename.items()
    ]

    def process_node(node):
        if node.get("Description") is not None:
            text = node["Description"]
            for pattern, replacement in patterns:
                text = pattern.sub(replacement, text)
            node["Description"] = text
        for child in node.get("SubDivisions", []) or []:
            process_node(child)
        for child in node.get("SubQuestions", []) or []:
            process_node(child)

    for q in questions:
        process_node(q)


def refresh_json_file(json_path: Path) -> bool:
    """
    Re-signs every embedded image's URL in an already-generated output
    JSON, using each image's durable S3 key ("filePath" in
    _embedded_media) to fetch a brand-new ViewFileUrl, then rewrites both
    _embedded_media and every baked <img> tag in place.

    Presigned URLs expire ~900s after being issued, so this should be run
    right before the JSON is actually consumed/displayed -- not just once
    right after the initial docx -> json conversion. Safe to run
    repeatedly; it only touches "imageUrl" values, never "filePath" (the
    permanent S3 key) or any question text.
    """
    try:
        with open(json_path, "r", encoding="utf-8") as f:
            data = json.load(f)
    except Exception as e:
        print(f"  [ERROR] could not read {json_path.name}: {e}")
        return False

    media = data.get("_embedded_media") or []
    if not media:
        print(f"  [SKIP] {json_path.name}: no embedded media to refresh")
        return False

    url_by_basename = {}
    refreshed = 0
    for entry in media:
        file_path = entry.get("filePath")
        fname = entry.get("fileName")
        if not file_path:
            continue
        fresh_url = fetch_view_url(file_path)
        if not fresh_url:
            print(f"    [WARN] could not refresh URL for {fname} ({file_path})")
            continue
        entry["imageUrl"] = fresh_url
        url_by_basename[Path(file_path).name] = fresh_url
        refreshed += 1

    if not url_by_basename:
        print(f"  [SKIP] {json_path.name}: no URLs could be refreshed")
        return False

    refresh_baked_urls(data.get("questions", []), url_by_basename)

    with open(json_path, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)

    print(f"  [OK] {json_path.name}: refreshed {refreshed}/{len(media)} image URL(s)")
    return True


def process_file(docx_path: Path, out_dir: Path):
    has_equations = has_omml_equations(docx_path)

    try:
        questions, flags = extract_questions_from_docx(docx_path)
    except Exception as e:
        print(f"  [ERROR] Failed to process {docx_path.name}: {e}")
        return "error"

    if not questions:
        print(f"  [SKIP] {docx_path.name}: no questions found")
        return "no-questions"

    full_text = json.dumps(questions, ensure_ascii=False)
    media_dir_name, embedded_media = (None, [])
    if has_embedded_images(docx_path) or IMAGE_MARKER_RE.search(full_text):
        media_dir_name, embedded_media = extract_embedded_media(docx_path, out_dir)

    uploaded_media = []
    if embedded_media and media_dir_name:
        media_dir = out_dir / media_dir_name
        existing_out_path = out_dir / (docx_path.stem + ".json")
        already_uploaded = load_existing_media_map(existing_out_path)

        url_map = {}
        to_upload = []
        for fname in embedded_media:
            if fname in already_uploaded:
                old = already_uploaded[fname]
                fresh_url = fetch_view_url(old["filePath"]) or old["imageUrl"]
                url_map[fname] = {"filePath": old["filePath"], "imageUrl": fresh_url}
                print(f"    [REUSE] {fname} already uploaded at {old['filePath']}, skipping re-upload")
            else:
                to_upload.append(fname)

        if to_upload:
            image_paths = [media_dir / fname for fname in to_upload]
            new_results = upload_images_to_s3(image_paths, docx_path.stem)
            url_map.update(new_results)

        for fname in embedded_media:
            entry = {"fileName": fname}
            if fname in url_map:
                entry.update(url_map[fname])
            uploaded_media.append(entry)

        # Bake the real uploaded URLs into each question's Description in
        # place of the raw "----media/imageN.ext----" placeholder, so the
        # placeholder text never lands in the database.
        url_by_filename = {fname: info["imageUrl"] for fname, info in url_map.items() if info.get("imageUrl")}
        if url_by_filename:
            substitute_image_placeholders(questions, url_by_filename)

    subj_code = extract_subj_code(docx_path.stem)
    output = {"subj_code": subj_code, "questions": questions}
    if flags:
        output["_review_flags"] = flags
    if has_equations:
        output["_review_flags"] = output.get("_review_flags", []) + [{
            "file": docx_path.name,
            "reason": "document contains embedded equation object(s); preserved instead of skipped",
        }]
    if media_dir_name:
        output["_media_dir"] = media_dir_name
    if uploaded_media:
        output["_embedded_media"] = uploaded_media
    elif embedded_media:
        output["_embedded_media"] = embedded_media

    out_path = out_dir / (docx_path.stem + ".json")
    with open(out_path, "w", encoding="utf-8") as f:
        json.dump(output, f, ensure_ascii=False, indent=2)

    flag_note = f", {len(flags)} line(s) flagged for review" if flags else ""
    print(f"  [OK] {docx_path.name} -> {out_path.name} ({len(questions)} question(s){flag_note})")
    return "ok"


def main():
    args = sys.argv[1:]

    # --refresh <json_file_or_dir>: re-sign expired presigned image URLs in
    # an already-generated output JSON (or every *.json in a directory)
    # without re-parsing the source .docx. Run this right before the JSON
    # is actually displayed -- the URLs baked in at conversion time expire
    # after ~900s, which is what was causing broken images in the frontend.
    if args and args[0] == "--refresh":
        refresh_args = args[1:]
        target = (Path(refresh_args[0]).expanduser().resolve()
                  if refresh_args else Path(OUTPUT_PATH).expanduser().resolve())
        if not target.exists():
            sys.exit(f"Path does not exist: {target}")
        json_files = [target] if target.is_file() else sorted(target.glob("*.json"))
        if not json_files:
            sys.exit(f"No .json file(s) found at: {target}")

        print(f"Refreshing image URLs in {len(json_files)} file(s)...")
        ok_refresh = sum(1 for jf in json_files if refresh_json_file(jf))
        print(f"\nDone. {ok_refresh}/{len(json_files)} file(s) refreshed.")
        return

    input_path = Path(args[0]).expanduser().resolve() if len(args) >= 1 else Path(INPUT_PATH).expanduser().resolve()
    output_path = Path(args[1]).expanduser().resolve() if len(args) >= 2 else Path(OUTPUT_PATH).expanduser().resolve()

    if not input_path.exists():
        sys.exit(f"Path does not exist: {input_path}")

    if input_path.is_file():
        if input_path.suffix.lower() != ".docx":
            sys.exit(f"Not a .docx file: {input_path}")
        docx_files = [input_path]
    else:
        docx_files = sorted(input_path.rglob("*.docx"))
        docx_files = [f for f in docx_files if not f.name.startswith("~$")]

    if not docx_files:
        sys.exit(f"No .docx files found at: {input_path}")

    seen_stems = {}
    for f in docx_files:
        seen_stems.setdefault(f.stem, []).append(f)
    dupes = {stem: paths for stem, paths in seen_stems.items() if len(paths) > 1}
    if dupes:
        print("[WARNING] Duplicate filenames found across subfolders - later files will overwrite earlier JSON output:")
        for stem, paths in dupes.items():
            for p in paths:
                print(f"    {stem}.docx  <-  {p}")

    print(f"Found {len(docx_files)} .docx file(s).")
    output_path.mkdir(parents=True, exist_ok=True)

    ok_count = 0
    no_q_count = 0
    image_count = 0
    eqn_count = 0
    error_count = 0
    for docx_file in docx_files:
        result = process_file(docx_file, output_path)
        if result == "ok":
            ok_count += 1
        elif result == "has-image":
            image_count += 1
        elif result == "has-equations":
            eqn_count += 1
        elif result == "no-questions":
            no_q_count += 1
        else:
            error_count += 1

    print(f"\nDone. {ok_count}/{len(docx_files)} file(s) converted successfully.")
    if image_count:
        print(f"  {image_count} file(s) skipped: contain embedded image(s) (not handled yet)")
    if eqn_count:
        print(f"  {eqn_count} file(s) skipped: contain embedded equation object(s) (not handled yet)")
    if no_q_count:
        print(f"  {no_q_count} file(s) skipped: no questions found (check structure)")
    if error_count:
        print(f"  {error_count} file(s) skipped: processing error")


if __name__ == "__main__":
    main()