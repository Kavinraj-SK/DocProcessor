import argparse
import base64
import csv
import json
import mimetypes
import os
import sys
import time
import re
from pathlib import Path

import requests
from dotenv import load_dotenv

load_dotenv()

API_URL = os.getenv(
    "API_BASE_URL",
    "https://campus.psgtech.ac.in/dote/api/QuestionBank/SaveQuestion",
)

# Confirmed via prior working requests against this same API: exam_ID must be
# a real exam record ID, not 0, or the server can't resolve Subj_Id and the
# insert fails with a NULL constraint error. Override with EXAM_ID_DEFAULT in
# .env if this batch belongs to a different exam.
EXAM_ID_DEFAULT = int(os.getenv("EXAM_ID_DEFAULT", "1"))


DEFAULT_FOLDER = r"C:\Users\Administrator\Downloads\DocQuesBuild\test"

QUES_TOKEN = os.getenv("ques_TOKEN")

# (connect, read) timeout in seconds for SaveQuestion. The old fixed 60s read
# timeout was too short when the server is slow. Override in .env, e.g.
#   SAVE_CONNECT_TIMEOUT=10
#   SAVE_READ_TIMEOUT=240
SAVE_TIMEOUT = (
    float(os.getenv("SAVE_CONNECT_TIMEOUT", "10")),
    float(os.getenv("SAVE_READ_TIMEOUT", "180")),
)
QUESTION_IMAGE_UPLOAD_URL = os.getenv(
    "QUESTION_IMAGE_UPLOAD_URL",
    API_URL.rsplit("/", 1)[0] + "/UploadQuestionImages.handle",
)
ENABLE_SEPARATE_IMAGE_UPLOAD = os.getenv("ENABLE_SEPARATE_IMAGE_UPLOAD", "").lower() in ("1", "true", "yes")
QUESTION_IMAGE_FILE_FIELD = os.getenv("QUESTION_IMAGE_FILE_FIELD", "file")
QUESTION_IMAGE_ID_FIELD = os.getenv("QUESTION_IMAGE_ID_FIELD", "ques_Id")

IMAGE_MARKER_RE = re.compile(r"----media/([A-Za-z0-9_.-]+)----")


def _inline_media_references(description: str, media_dir: Path | None) -> str:
    if not description or not media_dir:
        return description or ""

    def replace_marker(match):
        image_name = match.group(1)
        image_path = media_dir / image_name
        if not image_path.exists():
            return match.group(0)
        mime_type = mimetypes.guess_type(image_path.name)[0] or "application/octet-stream"
        encoded = base64.b64encode(image_path.read_bytes()).decode("ascii")
        return f'<img alt="{image_name}" src="data:{mime_type};base64,{encoded}">'

    return IMAGE_MARKER_RE.sub(replace_marker, description)


def format_ques_content(description: str, media_dir: Path | None = None) -> str:
    """Wrap in <p> and turn embedded newlines into <br>, since the frontend
    renders this field as HTML -- plain '\\n' characters are otherwise
    silently dropped and lines don't render."""
    text = _inline_media_references((description or "").replace("\r\n", "\n").strip(), media_dir)
    text = text.replace("\n", "<br>")
    return f"<p>{text}</p>"


def build_payload(subj_code: str, description: str, parent_id, disp_order: int, media_dir: Path | None = None) -> dict:
    """Build the SaveQuestion payload with only the fields we've been told to fill."""
    return {
        "ques_Id": 0,
        "qbConfig_ID": 17,
        "exam_ID": EXAM_ID_DEFAULT,
        "subj_Id": 0,
        "subj_Code": subj_code,
        "unit_Id": 0,
        "co_Id": 0,
        "topic_Id": 0,
        "mark": 7,
        "parntQus_Id": parent_id,  # None -> sent as JSON null for top-level questions
        "ques_Lvl": 0,
        "disp_Order": disp_order,
        "quesType_Id": 10,
        "bloomLvl_Id": 0,
        "diffLvl_Id": 0,
        "ques_Content": format_ques_content(description, media_dir),
        "hash_Ques": "",
        "enc_Text_Ques": "",
        "enc_Key_Ques": "",
        "nonce_Ques": "",
        "enc_Tag_Ques": "",
        "orig_Enc_Text": "",
        "orig_Enc_Key": "",
        "orig_Nonce": "",
        "orig_Enc_Tag": "",
        "answers": [],
        "skippedDuplicates": [],
    }


def extract_new_id(response_json):
    """
    Try to pull the newly-created question's ID out of the API response.

    UNKNOWN until verified against a real response -- run with --test
    first and check the printed raw response. Adjust the key names
    below to match reality.
    """
    if response_json is None:
        return None
    if isinstance(response_json, int):
        return response_json
    if isinstance(response_json, dict):
        for key in ("ques_Id", "Ques_Id", "quesId", "id", "Id", "ID", "data"):
            if key in response_json:
                val = response_json[key]
                if isinstance(val, dict):
                    return extract_new_id(val)
                if isinstance(val, int):
                    return val
    return None


def extract_image_references(description: str):
    if not description:
        return []
    return list(dict.fromkeys(IMAGE_MARKER_RE.findall(description)))


def upload_question_images(session, question_id, description, media_dir, dry_run, test_mode):
    image_names = extract_image_references(description)
    if not image_names:
        return True

    if not media_dir or not media_dir.exists():
        print(f"  WARNING: media folder missing for question {question_id}; skipping image upload")
        return False

    headers = {
        "accept": "*/*",
        "Authorization": f"Bearer {QUES_TOKEN}",
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
                      "(KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36",
    }

    all_ok = True
    for image_name in image_names:
        image_path = media_dir / image_name
        if not image_path.exists():
            print(f"  WARNING: image file not found: {image_path}")
            all_ok = False
            continue

        if dry_run:
            print(f"[DRY RUN] Would upload image for question {question_id}: {image_path.name}")
            continue

        mime_type = mimetypes.guess_type(image_path.name)[0] or "application/octet-stream"
        with open(image_path, "rb") as f:
            files = {
                QUESTION_IMAGE_FILE_FIELD: (image_path.name, f, mime_type),
            }
            data = {
                QUESTION_IMAGE_ID_FIELD: str(question_id),
                "questionId": str(question_id),
                "quesId": str(question_id),
            }

            try:
                resp = session.post(QUESTION_IMAGE_UPLOAD_URL, headers=headers, data=data, files=files, timeout=60)
            except requests.RequestException as e:
                print(f"  WARNING: image upload failed for {image_path.name}: {e}")
                all_ok = False
                continue

        if resp.status_code not in (200, 201):
            allow_hdr = resp.headers.get("Allow", "<not sent>")
            print(
                f"  WARNING: image upload returned http {resp.status_code} for {image_path.name}: "
                f"{resp.text[:200]} | Allow={allow_hdr}"
            )
            all_ok = False
        elif test_mode:
            print(f"  -> uploaded image {image_path.name}: {resp.text[:200]}")

    return all_ok


def post_question(session, payload, dry_run, test_mode):
    if dry_run:
        print(f"[DRY RUN] Would POST: subj={payload['subj_Code']} "
              f"parent={payload['parntQus_Id']} content={payload['ques_Content'][:60]!r}")
        return True, None, "dry-run"

    headers = {
        "accept": "*/*",
        "Authorization": f"Bearer {QUES_TOKEN}",
        "Content-Type": "application/json",
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
                      "(KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36",
    }

    payload_kb = len(json.dumps(payload).encode("utf-8")) / 1024
    if test_mode or payload_kb > 500:
        print(f"  payload size: {payload_kb:.0f} KB")

    for attempt in range(3):
        try:
            resp = session.post(API_URL, headers=headers, json=payload, timeout=SAVE_TIMEOUT)
        except requests.RequestException as e:
            if attempt == 2:
                # NOTE: a read timeout does not prove the save failed - the
                # server may have stored the question and only been slow to
                # answer. Check the DB/portal before re-running.
                return False, None, f"request error (payload {payload_kb:.0f} KB): {e}"
            time.sleep(5 * (attempt + 1))
            continue

        if resp.status_code in (200, 201):
            try:
                resp_json = resp.json()
            except ValueError:
                resp_json = None

            if test_mode:
                print(f"  -> status {resp.status_code}, raw response: {resp.text[:500]}")

            new_id = extract_new_id(resp_json)
            return True, new_id, "ok"

        # Duplicate / near-duplicate detection: the server returns a non-2xx
        # status but with a message describing an existing similar question,
        # rather than a genuine failure. Treat that as "duplicate", not error.
        try:
            body = resp.json()
        except ValueError:
            body = None
        if isinstance(body, dict):
            message = str(body.get("message", "")).lower()
            if "duplicate" in message or "similar questions already exist" in message:
                pct = body.get("similarityPercentage")
                existing = body.get("existingQuestion") or body.get("existingQuestions")
                if test_mode:
                    print(f"  -> DUPLICATE ({pct}% similar): {str(existing)[:200]}")
                return False, None, f"duplicate ({pct}% similar): {body.get('message', '')[:150]}"

        if resp.status_code in (401, 403):
            return False, None, f"auth error {resp.status_code}: {resp.text[:200]}"

        if resp.status_code == 405:
            redirected = " -> ".join(h.url for h in resp.history) + (" -> " + resp.url if resp.history else "")
            allow_hdr = resp.headers.get("Allow", "<not sent>")
            detail = (f"http 405 | final_url={resp.url} | redirects={len(resp.history)} "
                      f"({redirected if resp.history else 'none'}) | Allow_header={allow_hdr} "
                      f"| body={resp.text[:150]}")
            if test_mode:
                print(f"  -> {detail}")
            return False, None, detail

        if resp.status_code >= 500 and attempt < 2:
            time.sleep(5 * (attempt + 1))
            continue

        return False, None, f"http {resp.status_code}: {resp.text[:200]}"

    return False, None, "failed after retries"


def load_processed(log_path):
    """Return set of (file, qno_path) already marked 'ok' or 'duplicate' in a prior run.
    Duplicates are also treated as done -- retrying them just re-confirms the same
    near-duplicate match and wastes an API call."""
    done = set()
    if not os.path.exists(log_path):
        return done
    with open(log_path, newline="", encoding="utf-8") as f:
        for row in csv.DictReader(f):
            if row.get("status") in ("ok", "duplicate"):
                done.add((row["file"], row["qno_path"]))
    return done


class SubjectNotFoundError(Exception):
    """Raised when the server can't resolve subj_Code to a real subject
    (NULL Subj_Id constraint error) -- no point retrying every remaining
    question in this file against the same bad subject code."""
    pass


def process_node(node, subj_code, qno_path, parent_id, disp_order,
                  session, writer, done, dry_run, test_mode, delay, media_dir):
    """
    Recursively process one question node (and its SubDivisions/SubQuestions).
    Returns the ID to use as parent for this node's own children:
      - if this node had content and was saved -> its new ID
      - if it had no content (pure container) -> pass through parent_id unchanged
    Raises SubjectNotFoundError if subj_Code can't be resolved at all.
    """
    description = node.get("Description")
    current_id = parent_id

    if description is not None:
        key = (str(node.get("_file")), qno_path)
        if key not in done:
            payload = build_payload(subj_code, description, parent_id, disp_order, media_dir)
            ok, new_id, status = post_question(session, payload, dry_run, test_mode)
            is_duplicate = (not ok) and status.startswith("duplicate")
            is_subject_missing = (
                not ok and not is_duplicate
                and "subj_id" in status.lower() and "null" in status.lower()
            )
            writer.writerow({
                "file": node.get("_file"),
                "qno_path": qno_path,
                "status": "ok" if ok else ("duplicate" if is_duplicate else "error"),
                "detail": status,
                "new_id": new_id if new_id is not None else "",
            })
            if not dry_run:
                time.sleep(delay)
            if ok and new_id is not None:
                current_id = new_id
                if ENABLE_SEPARATE_IMAGE_UPLOAD:
                    upload_question_images(session, new_id, description, media_dir, dry_run, test_mode)
            elif is_subject_missing:
                raise SubjectNotFoundError(subj_code)
            elif is_duplicate:
                print(f"  SKIPPED (duplicate) {qno_path}: {status}")
            elif not ok:
                print(f"  FAILED {qno_path}: {status}")
        else:
            print(f"  skipping already-done {qno_path}")

    children = list(node.get("SubDivisions", [])) + list(node.get("SubQuestions", []))
    for i, child in enumerate(children):
        child["_file"] = node.get("_file")
        process_node(child, subj_code, f"{qno_path}.{child.get('Qno')}",
                     current_id, i, session, writer, done, dry_run, test_mode, delay, media_dir)


BAD_SUBJ_CACHE_PATH = "bad_subj_codes.json"


def load_bad_subj_codes(cache_path=BAD_SUBJ_CACHE_PATH):
    """subj_codes already confirmed missing (NULL Subj_Id) on the server,
    across this run and any previous one. Checked BEFORE calling the API,
    so a known-bad file is skipped in under a second instead of costing
    another ~90s round trip to find out again."""
    if not os.path.exists(cache_path):
        return set()
    try:
        with open(cache_path, "r", encoding="utf-8") as f:
            return set(json.load(f))
    except (json.JSONDecodeError, OSError):
        return set()


def save_bad_subj_code(subj_code, cache_path=BAD_SUBJ_CACHE_PATH):
    codes = load_bad_subj_codes(cache_path)
    if subj_code in codes:
        return
    codes.add(subj_code)
    with open(cache_path, "w", encoding="utf-8") as f:
        json.dump(sorted(codes), f, indent=2)


def process_file(path, session, writer, done, dry_run, test_mode, delay, bad_subj_codes):
    with open(path, "r", encoding="utf-8") as f:
        data = json.load(f)

    subj_code = data.get("subj_code", "")
    questions = data.get("questions", [])
    media_dir = None
    media_dir_name = data.get("_media_dir")
    if media_dir_name:
        media_dir = path.parent / media_dir_name
    print(f"Processing {path.name} ({len(questions)} top-level questions)")

    # Already confirmed missing on the server (this run or a previous one) ->
    # skip instantly, no API call, no ~90s wait.
    if subj_code in bad_subj_codes:
        print(f"  SKIPPING (known bad subj_code): {subj_code!r} -- {path.name}")
        writer.writerow({
            "file": path.name,
            "qno_path": "FILE",
            "status": "skipped-file",
            "detail": f"subj_code {subj_code!r} previously confirmed not found (cached)",
            "new_id": "",
        })
        return

    for i, q in enumerate(questions):
        q["_file"] = path.name
        try:
            process_node(q, subj_code, str(q.get("Qno")), None, i,
                         session, writer, done, dry_run, test_mode, delay, media_dir)
        except SubjectNotFoundError:
            print(f"  SKIPPING REST OF FILE: subj_code {subj_code!r} not found on server "
                  f"(NULL Subj_Id) -- {path.name}")
            writer.writerow({
                "file": path.name,
                "qno_path": "FILE",
                "status": "skipped-file",
                "detail": f"subj_code {subj_code!r} not found",
                "new_id": "",
            })
            bad_subj_codes.add(subj_code)
            save_bad_subj_code(subj_code)
            break


class MissingTokenError(Exception):
    """Raised by run_upload() instead of sys.exit(), so callers other than
    the CLI (e.g. the web app) can catch this and show it in their own UI
    rather than having the whole process killed."""


def run_upload(folder, log="upload_log.csv", dry_run=False, test_mode=False,
                limit=None, delay=0.6, bad_subj_cache_path=None):
    """
    Core upload loop, callable directly (not just via the CLI). Same
    behaviour as main(), but takes plain arguments, returns a summary
    dict instead of printing "Done." and exiting, and raises
    MissingTokenError instead of sys.exit(1) on a missing token so a
    caller like the web app can handle it gracefully.
    """
    if not dry_run and not QUES_TOKEN:
        raise MissingTokenError(
            "ques_TOKEN not found. Put it in a .env file as ques_TOKEN=..."
        )

    folder = Path(folder)
    files = sorted(folder.glob("*.json"))
    if not files:
        raise FileNotFoundError(f"No .json files found in {folder}")

    if test_mode:
        files = files[:1]
        print(f"TEST MODE: processing only {files[0].name}")
    elif limit:
        files = files[:limit]

    done = load_processed(log)
    log_exists = os.path.exists(log)

    session = requests.Session()
    bad_subj_cache_path = bad_subj_cache_path or BAD_SUBJ_CACHE_PATH

    with open(log, "a", newline="", encoding="utf-8") as logf:
        writer = csv.DictWriter(logf, fieldnames=["file", "qno_path", "status", "detail", "new_id"])
        if not log_exists:
            writer.writeheader()

        bad_subj_codes = load_bad_subj_codes(bad_subj_cache_path)
        if bad_subj_codes:
            print(f"Loaded {len(bad_subj_codes)} known-bad subj_code(s) from "
                  f"{bad_subj_cache_path} -- those files will be skipped instantly.")

        ok = errors = 0
        for path in files:
            try:
                process_file(path, session, writer, done, dry_run, test_mode, delay, bad_subj_codes)
                ok += 1
            except Exception as e:
                errors += 1
                # Deliberately stdout, not stderr: callers like the web app
                # only capture/stream stdout, and an error that's silently
                # invisible in the live console is worse than one on the
                # "wrong" stream for CLI purists.
                print(f"[ERROR] processing {path.name}: {e}")
            logf.flush()

    print(f"\nDone. See {log} for full results.")
    return {"total_files": len(files), "processed": ok, "errors": errors, "log": str(log)}


def main():
    parser = argparse.ArgumentParser(description="Upload question-bank JSON files to SaveQuestion API")
    parser.add_argument("--folder", default=DEFAULT_FOLDER,
                         help=f"Folder containing the .json files (default: {DEFAULT_FOLDER})")
    parser.add_argument("--log", default="upload_log.csv", help="CSV log file (also used to resume)")
    parser.add_argument("--dry-run", action="store_true", help="Don't call the API, just print payloads")
    parser.add_argument("--test", action="store_true", help="Process only the first file and print raw responses")
    parser.add_argument("--limit", type=int, default=None, help="Only process this many files")
    parser.add_argument("--delay", type=float, default=0.6, help="Seconds to sleep between real API calls")
    args = parser.parse_args()

    try:
        run_upload(args.folder, log=args.log, dry_run=args.dry_run, test_mode=args.test,
                   limit=args.limit, delay=args.delay)
    except MissingTokenError as e:
        print(f"ERROR: {e}", file=sys.stderr)
        sys.exit(1)
    except FileNotFoundError as e:
        print(str(e), file=sys.stderr)
        sys.exit(1)


if __name__ == "__main__":
    main()