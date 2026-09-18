#!/usr/bin/env python3
"""
Extract exam questions (including any data/fill-in tables that belong to a
question) from .docx files into JSON, using the Gemini API for the actual
structural understanding instead of hand-rolled regex parsing.

Why: exam-paper .docx files in this project come in at least three very
different shapes (flat numbered lists, [digit, text] table-per-question,
free-text "Question" markers with practical-output tables attached) and mix
free text, embedded equations and embedded fill-in-the-blank tables in ways
that are brittle to parse with regex. This script instead:

  1. Renders each .docx into one ordered text/markdown blob, with real
     tables kept as inline markdown pipe-tables (docx_to_markdown()).
  2. Sends that blob to Gemini with a JSON schema (ExtractionResult below)
     so Gemini does the structural reasoning -- which table belongs to
     which question, where a new question starts, what's just a grading
     rubric to discard -- and returns validated, schema-matching JSON.

Also handles embedded images: every image embedded in the .docx is saved to a
temp file and handed to your existing UploadQuestionImages procedure (S3
upload), and the resulting imageUrl is attached to whichever question,
subdivision or table it belonged to -- see "5. IMAGE UPLOAD" below for how
to point this at your actual UploadQuestionImages implementation.

Usage:
    pip install google-genai pydantic docx2python python-docx requests
    export GEMINI_API_KEY=...          # https://aistudio.google.com/apikey
    python3 doc_to_json_gemini.py <input_dir_or_file> <output_dir>

If no arguments are given, INPUT_PATH / OUTPUT_PATH below are used instead.
"""

from __future__ import annotations

import os
import re
import sys
import tempfile
from pathlib import Path
from typing import Dict, List, Optional

from pydantic import BaseModel, Field, ValidationError
from dotenv import load_dotenv
load_dotenv()

try:
    from docx2python import docx2python
except ImportError:
    sys.exit("Missing dependency. Install it with:\n    pip install docx2python --break-system-packages")

try:
    from google import genai
    from google.genai import types
except ImportError:
    sys.exit("Missing dependency. Install it with:\n    pip install google-genai --break-system-packages")


# ============================================================================
# SET THESE TWO PATHS AND RUN THE SCRIPT WITH NO ARGUMENTS
# ============================================================================
INPUT_PATH = r"C:\Users\Administrator\Desktop\DocQues\test"
OUTPUT_PATH = r"C:\Users\Administrator\Desktop\DocQues\test"
# ============================================================================
MODEL = "gemini-2.1-flash"


# ============================================================================
# 1. SCHEMA -- the JSON shape Gemini must return for each file
# ============================================================================
#   {
#     "subj_code": "MA232431",
#     "questions": [
#       {
#         "Qno": "1",
#         "Description": "...",
#         "Tables": [ {"caption": "...", "headers": [...], "rows": [[...]]} ],
#         "SubDivisions": [
#           {
#             "Qno": "I",
#             "ParentQuestionId": "1",
#             "Description": "...",
#             "Tables": [...],
#             "SubQuestions": [
#               {"Qno": "1", "ParentQuestionId": "I", "Description": "...", "Tables": [...]}
#             ]
#           }
#         ]
#       }
#     ]
#   }
#
# Every level (question / sub-division / sub-question) can carry its own
# "Tables" list, since a fill-in-the-blank or data table in the source
# document can sit directly under any of those levels. Purely
# administrative tables (an "ALLOCATION OF MARKS" grading rubric, a
# signature block, etc.) are never emitted as a Table anywhere.

class TableData(BaseModel):
    """A single data/fill-in-the-blank table that belongs to a question."""

    caption: Optional[str] = Field(
        default=None,
        description="The short title/label text that introduced this table in the "
                    "source document, e.g. 'Table 2: Output for Ellipse'. Omit if "
                    "the table had no caption line before it.",
    )
    headers: Optional[List[str]] = Field(
        default=None,
        description="Column header labels for the table, in order, if the table "
                    "has a clear header row. Omit if the table has no header row "
                    "(e.g. a vertical label/value grid).",
    )
    rows: List[List[str]] = Field(
        description="Every data row of the table, each as a list of cell strings "
                    "in column order. Preserve blank cells as empty strings (''), "
                    "since blank cells are usually where a student fills in an "
                    "answer. Do not skip rows just because they are blank."
    )


class ImageData(BaseModel):
    """Matches the shape UploadQuestionImages returns per image: {filePath, imageUrl}."""

    filePath: str
    imageUrl: str


# NOTE: Images is filled in AFTER Gemini responds (see "5. IMAGE UPLOAD"
# below) -- it is not something Gemini itself produces, so it's not
# mentioned in SYSTEM_PROMPT. Gemini is instead told to leave any
# "[IMAGE:filename]" token it finds in the source text exactly as-is inside
# Description/table cells; post-processing then finds those tokens, uploads
# the matching image, replaces the token with nothing, and records the
# resulting {filePath, imageUrl} in this Images list on whichever
# question/subdivision/sub-question the token was found in.

class SubQuestion(BaseModel):
    Qno: str = Field(description="The sub-question's own number/label, e.g. '1', '2'.")
    ParentQuestionId: str = Field(description="The roman-numeral SubDivision this belongs to, e.g. 'I'.")
    Description: str
    Tables: Optional[List[TableData]] = None
    Images: Optional[List[ImageData]] = None


class SubDivision(BaseModel):
    Qno: str = Field(description="Roman numeral label for this subsection, e.g. 'I', 'II'.")
    ParentQuestionId: str = Field(description="The digit Qno of the parent question, e.g. '1'.")
    Description: str = Field(
        default="",
        description="Text belonging directly to this subsection, before/outside any SubQuestions.",
    )
    Tables: Optional[List[TableData]] = None
    Images: Optional[List[ImageData]] = None
    SubQuestions: Optional[List[SubQuestion]] = None


class Question(BaseModel):
    Qno: str = Field(description="The question number as it appears in the source document "
                                  "(or, for documents with no explicit numbering, the "
                                  "1-based sequential order in which the question appears).")
    Description: Optional[str] = Field(
        default=None,
        description="Full question text. Use this field when the question has no roman-numeral "
                    "SubDivisions -- i.e. it's a flat question, possibly with (i)/(ii)/(iii) "
                    "style parts left inline as plain text.",
    )
    Tables: Optional[List[TableData]] = Field(
        default=None,
        description="Data/answer tables that sit directly under this question (not under one "
                    "of its SubDivisions). Never include grading-rubric tables here.",
    )
    Images: Optional[List[ImageData]] = Field(
        default=None,
        description="Filled in by post-processing, not by Gemini -- see ImageData.",
    )
    SubDivisions: Optional[List[SubDivision]] = Field(
        default=None,
        description="Use this instead of Description when the question is explicitly split "
                    "into roman-numeral subsections (I, II, III, ...).",
    )


class ExtractionResult(BaseModel):
    subj_code: str = Field(description="Subject code taken from the file name, e.g. 'MA232431'.")
    questions: List[Question]


# ============================================================================
# 2. DOCX -> MARKDOWN -- render the .docx as one ordered text/markdown blob,
#    with real tables kept as inline markdown pipe-tables, for Gemini to read
# ============================================================================

# docx2python inlines each embedded image as a marker like
# "----Image alt text---->C:\...\Parabolic Arch.jpg<----media/image1.jpeg----"
# (sometimes without the "Image alt text" part). We keep the media filename
# -- it's also the key into docx2python's own doc.images dict -- so we can
# later map "this placeholder token" -> "this image's bytes" -> "this
# image's uploaded URL". IMAGE_TOKEN_RE is what post-processing looks for.
IMAGE_FULL_MARKER_RE = re.compile(
    r"-{2,}\s*Image alt text\s*-{2,}\s*>.*?<\s*-{2,}\s*media/(image\d+\.\w+)\s*-{2,}",
    re.DOTALL | re.IGNORECASE,
)
IMAGE_MARKER_RE = re.compile(r"-{2,}\s*media/(image\d+\.\w+)\s*-{2,}")
IMAGE_TOKEN_RE = re.compile(r"\[IMAGE:([^\[\]]+?)\]")


def _clean(text) -> str:
    if text is None:
        return ""
    if not isinstance(text, str):
        text = "\n".join(str(p) for p in text if p)
    text = IMAGE_FULL_MARKER_RE.sub(lambda m: f" [IMAGE:{m.group(1)}] ", text)
    text = IMAGE_MARKER_RE.sub(lambda m: f" [IMAGE:{m.group(1)}] ", text)
    return text.replace("\t", " ").strip()


def _lineage_says_table(par) -> bool:
    lineage = getattr(par, "lineage", None)
    return bool(lineage) and len(lineage) > 1 and lineage[1] == "tbl"


def _is_real_table(pars_table) -> bool:
    for row in pars_table:
        for cell in row:
            for par in cell:
                if _lineage_says_table(par):
                    return True
    return False


def _row_to_md(cells) -> str:
    return "| " + " | ".join(c.replace("|", "/").replace("\n", " ") for c in cells) + " |"


def _table_to_markdown(rows) -> str:
    """rows: list of list of raw cell values (docx2python paragraph lists)."""
    cleaned = [[_clean(c) for c in row] for row in rows]
    if not cleaned:
        return ""
    ncols = max(len(r) for r in cleaned)
    cleaned = [r + [""] * (ncols - len(r)) for r in cleaned]

    # Skip entries that are just a blank spacer row/table.
    if all(all(not c for c in r) for r in cleaned):
        return ""

    lines = [_row_to_md(cleaned[0]), "| " + " | ".join(["---"] * ncols) + " |"]
    for r in cleaned[1:]:
        lines.append(_row_to_md(r))
    return "\n".join(lines)


def docx_to_markdown(docx_path: Path) -> str:
    """Returns the whole document as one ordered text/markdown blob."""
    with docx2python(str(docx_path)) as doc:
        body = doc.body
        try:
            body_pars = doc.body_pars
        except AttributeError:
            body_pars = None

        out_lines = []
        for idx, text_table in enumerate(body):
            is_real = _is_real_table(body_pars[idx]) if body_pars is not None else (
                not (len(text_table) == 1 and len(text_table[0]) == 1)
            )

            # A genuine table with more than one row (or many columns) is real
            # tabular data -- render as a markdown table. A "table" that's
            # really just a 1-row text container is flattened to plain text.
            looks_tabular = is_real and (len(text_table) > 1 or (
                len(text_table) == 1 and len(text_table[0]) > 2
            ))

            if looks_tabular:
                md = _table_to_markdown(text_table)
                if md:
                    out_lines.append("\n" + md + "\n")
                continue

            flat_lines = []
            for row in text_table:
                for cell in row:
                    t = _clean(cell)
                    if t:
                        flat_lines.append(t)
            if flat_lines:
                out_lines.append("\n".join(flat_lines))

        return "\n\n".join(l for l in out_lines if l.strip())


# ============================================================================
# 3. GEMINI EXTRACTION
# ============================================================================

SYSTEM_PROMPT = """\
You extract exam questions from a university exam paper that has been converted to \
plain text/markdown. Tables in the source appear inline as markdown pipe-tables, in \
the same order they occur in the document.

Return data matching the given JSON schema. Follow these rules exactly:

1. QUESTIONS: Identify every real exam question in the document, in order, and give \
each one a "Qno" matching how it's numbered in the source (e.g. "1", "2"...). If the \
document has NO explicit question numbering at all, number the questions sequentially \
starting at "1" in the order they appear.

2. NEW-QUESTION MARKERS: A new question can start in more than one way in these \
documents -- watch for all of these:
   - A line that is just a number, or "<number> <question text>".
   - The literal word "Question" on its own line (the real question text follows it; \
text like "PRACTICALPART" or a part-number just before "Question" is a section header, \
not part of the question -- drop it).
   - A number together with real question prose appearing as the LAST row of what is \
otherwise an "ALLOCATION OF MARKS" grading-rubric table -- that row is the start of the \
next question, not part of the rubric.

3. TABLES THAT BELONG TO A QUESTION: A markdown table that appears within a question's \
text (e.g. a fill-in-the-blank grid the student completes as part of answering, an \
"Output" table, a data table referenced by the question) belongs to that question -- \
put it in that question's (or subdivision's/sub-question's) "Tables" list. Preserve \
every row and every cell exactly, including blank cells (students fill those in later \
-- keep them as ""). If a short caption line like "Table 2: Output for Ellipse" \
appears right before a table, use it as that table's "caption".

4. TABLES TO DISCARD (do NOT include as a Table anywhere): grading/marks-allocation \
rubrics ("ALLOCATION OF MARKS", "SNo / Descriptions / Max. Marks / Marks Awarded" \
grids, "TOTAL", "Internal Examiner / External Examiner" signature rows), and any purely \
administrative table (name/reg.no/date/signature blocks, blank mark-tally grids like a \
1-20 numbered grid with no question content). These are not question content.

5. SUBDIVISIONS: If a question is explicitly split into roman-numeral subsections \
(I, II, III, ...), use "SubDivisions" instead of "Description". If those subsections \
contain their own numbered sub-questions (1), 2), ...), nest them under "SubQuestions". \
Most questions in these documents are flat -- only use SubDivisions/SubQuestions when \
the source document actually has that structure. Inline lowercase parts like "i)", \
"ii)", "iii)" inside a question's prose are normal sentences, not SubDivisions -- leave \
them as plain text inside "Description".

6. Do not invent, summarize, or omit question text. Reproduce it faithfully (equations \
may appear as <latex>...</latex> fragments -- keep them as-is).

7. IMAGE TOKENS: The source text may contain tokens like "[IMAGE:image1.jpeg]" marking \
where an embedded image/diagram/graph appeared. Keep each such token character-for- \
character, in place, inside whichever "Description" or table "caption"/cell it \
appeared in (a question's own text, or a table cell -- wherever it originally sat). \
Do not remove, rename, renumber, paraphrase, or move these tokens, and do not invent \
new ones. They are handled separately after you respond.
"""


def get_client() -> "genai.Client":
    api_key = os.environ.get("GEMINI_API_KEY") or os.environ.get("GOOGLE_API_KEY")
    if not api_key:
        sys.exit(
            "No API key found. Set the GEMINI_API_KEY environment variable "
            "(get one at https://aistudio.google.com/apikey) and try again."
        )
    return genai.Client(api_key=api_key)


def extract_with_gemini(client: "genai.Client", subj_code: str, doc_markdown: str) -> ExtractionResult:
    prompt = f"Subject code: {subj_code}\n\n--- DOCUMENT START ---\n{doc_markdown}\n--- DOCUMENT END ---"

    response = client.models.generate_content(
        model=MODEL,
        contents=prompt,
        config=types.GenerateContentConfig(
            system_instruction=SYSTEM_PROMPT,
            response_mime_type="application/json",
            response_schema=ExtractionResult,
            temperature=0,
        ),
    )

    # response.parsed is the SDK's auto-validated Pydantic instance when
    # response_schema is a Pydantic model; fall back to manual validation
    # of response.text if that didn't populate for any reason.
    parsed = getattr(response, "parsed", None)
    if isinstance(parsed, ExtractionResult):
        return parsed
    return ExtractionResult.model_validate_json(response.text)


# ============================================================================
# 4. IMAGE UPLOAD -- posts embedded images to the UploadQuestionImages
#    endpoint and attaches the returned imageUrl back into the JSON
# ============================================================================

API_URL = os.getenv(
    "API_BASE_URL",
    "https://campus.psgtech.ac.in/dote/api/QuestionBank/UploadQuestionImages",
)

# Expected response shape:
#   {"isError": false, "message": "...", "images": [{"filePath": ..., "imageUrl": ...}]}
#
# ASSUMPTION (please correct if wrong): the endpoint takes the files as
# multipart/form-data under a field named "images" (one part per file). If
# it actually expects a different field name, a single combined field, or a
# JSON body with base64 data instead, adjust the `files=` construction in
# upload_question_images() below -- everything downstream (parsing the
# response and attaching imageUrl back onto each question) stays the same.
# If the endpoint needs auth, set UPLOAD_QUESTION_IMAGES_AUTH_TOKEN and it's
# sent as a Bearer token.

def extract_docx_images(docx_path: Path) -> Dict[str, bytes]:
    """filename (e.g. 'image1.jpeg') -> raw bytes, for every image embedded in the docx."""
    with docx2python(str(docx_path)) as doc:
        return dict(doc.images)


def upload_question_images(file_paths: List[Path]) -> Optional[dict]:
    """POSTs the given local files to UploadQuestionImages and returns the parsed
    response dict, or None if there's nothing to upload or the call failed."""
    if not file_paths:
        return None

    import requests

    headers = {}
    token = os.environ.get("UPLOAD_QUESTION_IMAGES_AUTH_TOKEN")
    if token:
        headers["Authorization"] = f"Bearer {token}"

    files = [("images", (p.name, open(p, "rb"))) for p in file_paths]
    try:
        resp = requests.post(API_URL, files=files, headers=headers, timeout=120)
        resp.raise_for_status()
        result = resp.json()
    except Exception as e:
        print(f"  [WARN] UploadQuestionImages request failed: {e}; "
              f"[IMAGE:...] tokens left as-is.")
        return None
    finally:
        for _, (_, fh) in files:
            fh.close()

    if result.get("isError"):
        print(f"  [WARN] UploadQuestionImages reported an error: {result.get('message')}")
        return None
    return result


def _strip_and_collect_tokens(text: Optional[str], url_by_filename: Dict[str, str],
                               found: List[ImageData]) -> Optional[str]:
    if not text:
        return text

    def repl(m: "re.Match") -> str:
        fname = m.group(1).strip()
        url = url_by_filename.get(fname)
        if url:
            found.append(ImageData(filePath=fname, imageUrl=url))
            return ""
        # No URL for this one (upload failed / missing) -- keep the token
        # visible rather than silently dropping the reference to it.
        return m.group(0)

    new_text = IMAGE_TOKEN_RE.sub(repl, text)
    return re.sub(r"[ \t]{2,}", " ", new_text).strip()


def _attach_images_to_node(node, url_by_filename: Dict[str, str]) -> None:
    """Recursively walks a Question/SubDivision/SubQuestion, replacing
    [IMAGE:filename] tokens in its own Description/Tables with nothing and
    recording {filePath, imageUrl} on that same node's Images list."""
    found: List[ImageData] = []

    if getattr(node, "Description", None):
        node.Description = _strip_and_collect_tokens(node.Description, url_by_filename, found)

    if getattr(node, "Tables", None):
        for table in node.Tables:
            if table.caption:
                table.caption = _strip_and_collect_tokens(table.caption, url_by_filename, found)
            table.rows = [
                [_strip_and_collect_tokens(cell, url_by_filename, found) or "" for cell in row]
                for row in table.rows
            ]

    if found:
        seen = set()
        deduped = []
        for img in found:
            if img.filePath not in seen:
                seen.add(img.filePath)
                deduped.append(img)
        node.Images = deduped

    for sub_division in getattr(node, "SubDivisions", None) or []:
        _attach_images_to_node(sub_division, url_by_filename)
    for sub_question in getattr(node, "SubQuestions", None) or []:
        _attach_images_to_node(sub_question, url_by_filename)


def resolve_images(result: ExtractionResult, docx_path: Path) -> None:
    """Finds every [IMAGE:filename] token Gemini preserved, uploads the matching
    embedded images via UploadQuestionImages, and attaches the returned imageUrl
    back onto whichever question/subdivision/sub-question referenced it. Mutates
    `result` in place. No-op if the document has no embedded images."""
    all_text_chunks = []
    for q in result.questions:
        all_text_chunks.append(q.Description or "")
        for t in q.Tables or []:
            all_text_chunks.append(t.caption or "")
            all_text_chunks.extend(c for row in t.rows for c in row)
        for sd in q.SubDivisions or []:
            all_text_chunks.append(sd.Description or "")
            for sq in sd.SubQuestions or []:
                all_text_chunks.append(sq.Description or "")

    referenced_filenames = set()
    for chunk in all_text_chunks:
        referenced_filenames.update(m.group(1).strip() for m in IMAGE_TOKEN_RE.finditer(chunk))
    if not referenced_filenames:
        return

    image_bytes = extract_docx_images(docx_path)
    with tempfile.TemporaryDirectory(prefix="question_images_") as tmp_dir:
        tmp_paths = []
        for fname in sorted(referenced_filenames):
            data = image_bytes.get(fname)
            if data is None:
                print(f"  [WARN] {docx_path.name}: referenced image '{fname}' not found "
                      f"in docx media -- leaving its token unresolved.")
                continue
            p = Path(tmp_dir) / fname
            p.write_bytes(data)
            tmp_paths.append(p)

        upload_result = upload_question_images(tmp_paths)
        if not upload_result:
            # No upload mechanism configured, or it failed -- leave the
            # [IMAGE:...] tokens in the text untouched rather than silently
            # deleting them, so the loss of the image is still visible in
            # the JSON instead of vanishing into "draw the graph." with no
            # trace it was ever there.
            return

        url_by_filename = {
            Path(img["filePath"]).name: img["imageUrl"]
            for img in upload_result.get("images", [])
        }
        for q in result.questions:
            _attach_images_to_node(q, url_by_filename)


# ============================================================================
# 5. DRIVER
# ============================================================================

def process_file(client: "genai.Client", docx_path: Path, out_dir: Path) -> str:
    try:
        doc_markdown = docx_to_markdown(docx_path)
    except Exception as e:
        print(f"  [ERROR] Failed to read {docx_path.name}: {e}")
        return "error"

    if not doc_markdown.strip():
        print(f"  [SKIP] {docx_path.name}: no extractable content")
        return "no-content"

    subj_code = docx_path.stem.split(" ", 1)[0].split("_", 1)[0]

    try:
        result = extract_with_gemini(client, subj_code, doc_markdown)
    except ValidationError as e:
        print(f"  [ERROR] {docx_path.name}: Gemini output didn't match schema: {e}")
        return "error"
    except Exception as e:
        print(f"  [ERROR] Gemini request failed for {docx_path.name}: {e}")
        return "error"

    if not result.questions:
        print(f"  [SKIP] {docx_path.name}: no questions found")
        return "no-questions"

    try:
        resolve_images(result, docx_path)
    except Exception as e:
        print(f"  [WARN] {docx_path.name}: image upload step failed ({e}); "
              f"continuing with unresolved [IMAGE:...] tokens.")

    out_path = out_dir / (docx_path.stem + ".json")
    with open(out_path, "w", encoding="utf-8") as f:
        f.write(result.model_dump_json(indent=2, exclude_none=True))

    n_tables = sum(
        len(q.Tables or [])
        + sum(len(sd.Tables or []) + sum(len(sq.Tables or []) for sq in (sd.SubQuestions or []))
              for sd in (q.SubDivisions or []))
        for q in result.questions
    )
    print(f"  [OK] {docx_path.name} -> {out_path.name} "
          f"({len(result.questions)} question(s), {n_tables} table(s))")
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

    print(f"Found {len(docx_files)} .docx file(s).")
    output_path.mkdir(parents=True, exist_ok=True)
    client = get_client()

    counts = {"ok": 0, "error": 0, "no-questions": 0, "no-content": 0}
    for docx_file in docx_files:
        result = process_file(client, docx_file, output_path)
        counts[result] = counts.get(result, 0) + 1

    print(f"\nDone. {counts['ok']}/{len(docx_files)} file(s) converted successfully.")
    for key in ("no-questions", "no-content", "error"):
        if counts.get(key):
            print(f"  {counts[key]} file(s): {key}")


if __name__ == "__main__":
    main()