#!/usr/bin/env python3
"""
refresh_image_urls.py

Walks a folder of question JSON files (produced by doc_to_json.py) and
refreshes every embedded image's "imageUrl" via the File Service's
ViewFileUrl route, since presigned S3 URLs only last ~900 seconds (15
min). ViewFileUrl is documented in FILE_SERVICE.docx and confirmed
working, unlike the old QuestionBank RefreshImageUrls guess this script
originally targeted (same broken family as UploadQuestionImages).

Run this right BEFORE you actually need working image links -- e.g.
right before running json_to_db.py -- not right after doc_to_json.py,
since a refreshed URL will itself go stale again after ~15 minutes. It's
safe to run multiple times; it only touches files that have embedded
media with a stored "filePath".

ASSUMPTIONS (please verify / correct if wrong):
- Auth: FILE_SERVICE's docs say "Send RBAC token in header as <bearer
  RBAC Token>" -- assumed to be the same BEARER_TOKEN used elsewhere in
  this project.
- Request: GET with a "filepath" query param, ONE file per call --
  ViewFileUrl's documented example takes a single filepath, not a batch,
  so this makes one request per image rather than one per file.
- Response: {"isError": false, "result": {"url": "..."}}, per
  FILE_SERVICE.docx's documented example.

SETUP
-----
1. pip install requests python-dotenv --break-system-packages
2. .env should already have BEARER_TOKEN set (same one json_to_db.py uses).

USAGE
-----
    python refresh_image_urls.py --folder C:\\Users\\Administrator\\Desktop\\DocQues\\test
"""

import argparse
import json
import os
import sys
from pathlib import Path

from dotenv import load_dotenv

try:
    import requests
except ImportError:
    sys.exit("Missing dependency. Install it with:\n    pip install requests --break-system-packages")

load_dotenv()
BEARER_TOKEN = os.getenv("BEARER_TOKEN")
SFM_VIEW_URL = os.getenv("SFM_VIEW_URL", "https://campus.psgtech.ac.in/sfm/api/files/viewfileurl")


def refresh_urls(file_paths):
    """s3 key (filePath) -> fresh imageUrl, one ViewFileUrl call per key."""
    if not file_paths:
        return {}

    headers = {}
    if BEARER_TOKEN:
        headers["Authorization"] = f"Bearer {BEARER_TOKEN}"

    url_map = {}
    for fp in file_paths:
        try:
            resp = requests.get(SFM_VIEW_URL, params={"filepath": fp}, headers=headers, timeout=30)
            if not resp.ok:
                print(f"  [WARN] ViewFileUrl failed for {fp}: {resp.status_code} -- {resp.text[:200]}")
                continue
            result = resp.json()
        except Exception as e:
            print(f"  [WARN] ViewFileUrl request failed for {fp}: {e}")
            continue

        if result.get("isError"):
            print(f"  [WARN] ViewFileUrl reported error for {fp}: {result.get('message')}")
            continue
        
        url = (result.get("result") or {}).get("url")
        if url:
            url_map[fp] = url

    return url_map


def update_json_file(path: Path, url_map: dict) -> bool:
    with open(path, "r", encoding="utf-8") as f:
        data = json.load(f)

    media = data.get("_embedded_media")
    if not media:
        return False

    changed = False
    for entry in media:
        fp = entry.get("filePath")
        if fp and fp in url_map:
            entry["imageUrl"] = url_map[fp]
            changed = True

    if changed:
        with open(path, "w", encoding="utf-8") as f:
            json.dump(data, f, ensure_ascii=False, indent=2)
    return changed


def main():
    parser = argparse.ArgumentParser(description="Refresh expired S3 image URLs in question JSON files")
    parser.add_argument("--folder", required=True, help="Folder containing the JSON files to refresh")
    args = parser.parse_args()

    folder = Path(args.folder)
    json_files = sorted(folder.glob("*.json"))
    if not json_files:
        sys.exit(f"No .json files found in {folder}")

    # Collect every filePath across every file first so this is ONE batch
    # request, not one request per file/image.
    all_file_paths = set()
    file_media_map = {}
    for jf in json_files:
        try:
            with open(jf, "r", encoding="utf-8") as f:
                data = json.load(f)
        except Exception as e:
            print(f"  [WARN] couldn't read {jf.name}: {e}")
            continue
        media = data.get("_embedded_media") or []
        fps = [m.get("filePath") for m in media if m.get("filePath")]
        if fps:
            file_media_map[jf] = fps
            all_file_paths.update(fps)

    if not all_file_paths:
        print("No embedded images with a filePath found -- nothing to refresh.")
        return

    print(f"Refreshing {len(all_file_paths)} image URL(s) across {len(file_media_map)} file(s)...")
    url_map = refresh_urls(sorted(all_file_paths))
    if not url_map:
        sys.exit("No URLs were refreshed -- aborting without touching any files.")

    updated = 0
    for jf in file_media_map:
        if update_json_file(jf, url_map):
            updated += 1

    print(f"Done. Updated {updated}/{len(file_media_map)} file(s).")
    missing = all_file_paths - set(url_map.keys())
    if missing:
        preview = sorted(missing)[:5]
        print(f"  [WARN] {len(missing)} filePath(s) got no fresh URL back: "
              f"{preview}{'...' if len(missing) > 5 else ''}")


if __name__ == "__main__":
    main()