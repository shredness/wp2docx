#!/usr/bin/env python3
"""Export WordPress posts to .docx as 'YYYY.MM.DD Post Title.docx'.

Incremental: only re-exports posts whose modified time changed, renames the
file if a title/date changed. Runs once, or loops if INTERVAL is set.
"""
import html
import json
import os
import re
import subprocess
import sys
import tempfile
import time
from datetime import datetime
from pathlib import Path

import requests

WP_URL = os.environ.get("WP_URL", "http://wordpress").rstrip("/")      # URL reachable from this container
PUBLIC_URL = os.environ.get("PUBLIC_URL", "").rstrip("/")              # siteurl as stored in WP (e.g. http://localhost:8080)
WP_USER = os.environ.get("WP_USER")
WP_APP_PASSWORD = os.environ.get("WP_APP_PASSWORD")
STATUS = os.environ.get("WP_STATUS", "publish")                        # e.g. publish,draft,private,future (needs auth)
OUT = Path(os.environ.get("OUT_DIR", "/export"))
REFERENCE_DOCX = os.environ.get("REFERENCE_DOCX")                      # optional Word template for styles
INTERVAL = int(os.environ.get("INTERVAL", "0"))                        # seconds; 0 = run once
FORWARDED_PROTO = os.environ.get("WP_FORWARDED_PROTO", "https")        # empty to disable

STATE_FILE = OUT / ".wp2docx_state.json"
ILLEGAL = re.compile(r'[<>:"/\\|?*\x00-\x1f]')
TAGS = re.compile(r"<[^>]+>")


def safe_title(raw: str) -> str:
    t = html.unescape(TAGS.sub("", raw))
    t = ILLEGAL.sub("", t)
    t = re.sub(r"\s+", " ", t).strip().rstrip(". ")
    return t[:180] or "Untitled"


def fetch_posts(session):
    auth = (WP_USER, WP_APP_PASSWORD) if WP_USER and WP_APP_PASSWORD else None
    page = 1
    while True:
        r = session.get(
            f"{WP_URL}/wp-json/wp/v2/posts",
            params={"per_page": 100, "page": page, "status": STATUS,
                    "orderby": "date", "order": "asc",
                    "context": "edit" if auth else "view"},
            auth=auth, timeout=60,
        )
        if not r.ok:
            try:
                err = r.json()
                detail = f"{err.get('code')}: {err.get('message')}"
            except ValueError:
                detail = r.text[:300]
            raise RuntimeError(f"{r.status_code} from {r.url} -> {detail}")
        batch = r.json()
        if not batch:
            return
        yield from batch
        if page >= int(r.headers.get("X-WP-TotalPages", 1)):
            return
        page += 1


def to_docx(title: str, body_html: str, dest: Path):
    if PUBLIC_URL:
        body_html = body_html.replace(PUBLIC_URL, WP_URL)  # so pandoc can fetch images
    with tempfile.NamedTemporaryFile("w", suffix=".html", delete=False, encoding="utf-8") as f:
        f.write(f"<html><head><meta charset='utf-8'></head><body>{body_html}</body></html>")
        src = f.name
    tmp_out = dest.with_suffix(".docx.tmp")
    cmd = ["pandoc", src, "-f", "html", "-t", "docx", "-o", str(tmp_out),
           "--metadata", f"title={title}"]
    if REFERENCE_DOCX:
        cmd += ["--reference-doc", REFERENCE_DOCX]
    try:
        subprocess.run(cmd, check=True, capture_output=True, text=True)
        tmp_out.replace(dest)
    finally:
        os.unlink(src)
        tmp_out.unlink(missing_ok=True)


def run_once():
    OUT.mkdir(parents=True, exist_ok=True)
    state = json.loads(STATE_FILE.read_text()) if STATE_FILE.exists() else {}
    claimed = {}
    exported = skipped = 0

    with requests.Session() as s:
        if FORWARDED_PROTO:
            # WordPress only honours Application Passwords over HTTPS; the official
            # image maps X-Forwarded-Proto to $_SERVER['HTTPS'] so plain-http internal calls still count.
            s.headers["X-Forwarded-Proto"] = FORWARDED_PROTO
        for p in fetch_posts(s):
            pid = str(p["id"])
            title = safe_title(p["title"]["rendered"])
            date = datetime.fromisoformat(p["date"]).strftime("%Y.%m.%d")
            name = f"{date} {title}.docx"
            if name in claimed and claimed[name] != pid:        # same date + title collision
                name = f"{date} {title} ({pid}).docx"
            claimed[name] = pid
            dest = OUT / name

            prev = state.get(pid)
            if prev and prev["modified"] == p["modified_gmt"] and prev["file"] == name and dest.exists():
                skipped += 1
                continue
            if prev and prev["file"] != name:                   # title/date changed: drop old file
                (OUT / prev["file"]).unlink(missing_ok=True)

            try:
                to_docx(html.unescape(TAGS.sub("", p["title"]["rendered"])), p["content"]["rendered"], dest)
            except subprocess.CalledProcessError as e:
                print(f"[fail] {pid} {name}: {e.stderr.strip()}", file=sys.stderr)
                continue
            state[pid] = {"modified": p["modified_gmt"], "file": name}
            exported += 1
            print(f"[ok] {name}")

    STATE_FILE.write_text(json.dumps(state, indent=2))
    print(f"exported={exported} unchanged={skipped}")


if __name__ == "__main__":
    print(f"wp2docx start: WP_URL={WP_URL} STATUS={STATUS} OUT={OUT} "
          f"INTERVAL={INTERVAL or 'run-once'} auth={'yes' if WP_USER and WP_APP_PASSWORD else 'no'}")
    while True:
        try:
            run_once()
        except Exception as e:  # keep the loop alive if WP is restarting
            print(f"[error] {e}", file=sys.stderr)
        if not INTERVAL:
            break
        time.sleep(INTERVAL)
