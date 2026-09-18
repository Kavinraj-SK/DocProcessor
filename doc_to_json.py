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
"""

import json
import re
import sys
import zipfile
from pathlib import Path

try:
    from docx2python import docx2python
except ImportError:
    sys.exit("Missing dependency. Install it with:\n    pip install docx2python --break-system-packages")

try:
    import docx as _pydocx  # python-docx -- used only to detect embedded equation objects
except ImportError:
    _pydocx = None


# ============================================================================
# SET THESE TWO PATHS AND RUN THE SCRIPT WITH NO ARGUMENTS
# ============================================================================
INPUT_PATH = r"C:\Users\Administrator\Desktop\DocQues\test"
OUTPUT_PATH = r"C:\Users\Administrator\Desktop\DocQues\test"
# ============================================================================


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
IMAGE_MARKER_RE = re.compile(r'----media/image\d+\.\w+----')
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
    if embedded_media:
        output["_embedded_media"] = embedded_media

    out_path = out_dir / (docx_path.stem + ".json")
    with open(out_path, "w", encoding="utf-8") as f:
        json.dump(output, f, ensure_ascii=False, indent=2)

    flag_note = f", {len(flags)} line(s) flagged for review" if flags else ""
    print(f"  [OK] {docx_path.name} -> {out_path.name} ({len(questions)} question(s){flag_note})")
    return "ok"


def main():
    args = sys.argv[1:]
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