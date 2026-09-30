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
PUID = os.environ.get("PUID", "99")                                    # Unraid nobody; empty to leave ownership alone
PGID = os.environ.get("PGID", "100")                                   # Unraid users
FILE_MODE = int(os.environ.get("FILE_MODE", "0666"), 8)                # empty not allowed; e.g. 0666 or 0777

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


IMG_TAG = re.compile(r"<img\b[^>]*>", re.I)
IMG_SRC = re.compile(r"""\bsrc\s*=\s*(["'])(.*?)\1""", re.I | re.S)
IMG_ALT = re.compile(r"""\balt\s*=\s*(["'])(.*?)\1""", re.I | re.S)
IMG_EXT = {"image/jpeg": ".jpg", "image/png": ".png", "image/gif": ".gif", "image/webp": ".webp",
           "image/svg+xml": ".svg", "image/bmp": ".bmp", "image/tiff": ".tif"}


def localize_images(body_html: str, session, workdir: Path, label: str) -> str:
    """Download every <img> ourselves; embed real images, replace dead ones with a visible note."""
    cache = {}

    def swap(m):
        tag = m.group(0)
        sm = IMG_SRC.search(tag)
        if not sm or sm.group(2).startswith("data:"):
            return tag
        url = html.unescape(sm.group(2)).strip()
        if PUBLIC_URL and url.startswith(PUBLIC_URL):
            url = WP_URL + url[len(PUBLIC_URL):]
        elif url.startswith("//"):
            url = "https:" + url
        elif url.startswith("/"):
            url = WP_URL + url
        if url not in cache:
            cache[url] = None
            try:
                r = session.get(url, timeout=30)
                ctype = r.headers.get("Content-Type", "").split(";")[0].strip().lower()
                if r.ok and ctype.startswith("image/") and r.content:
                    f = workdir / f"img{len(cache)}{IMG_EXT.get(ctype, '')}"
                    f.write_bytes(r.content)
                    cache[url] = f
                else:
                    print(f"[warn] {label}: image not usable ({r.status_code} {ctype or 'no type'}): {url}",
                          file=sys.stderr)
            except requests.RequestException as e:
                print(f"[warn] {label}: image fetch failed ({e.__class__.__name__}): {url}", file=sys.stderr)
        local = cache[url]
        if local is None:
            am = IMG_ALT.search(tag)
            what = html.escape(html.unescape(am.group(2)).strip()) if am and am.group(2).strip() else html.escape(url)
            return f"<p><em>[image unavailable: {what}]</em></p>"
        tag = IMG_SRC.sub(lambda _: f'src="{local}"', tag, count=1)
        return re.sub(r"""\bsrcset\s*=\s*(["']).*?\1""", "", tag, flags=re.I | re.S)

    return IMG_TAG.sub(swap, body_html)


def to_docx(title: str, body_html: str, dest: Path, session):
    with tempfile.TemporaryDirectory() as work:
        workdir = Path(work)
        body_html = localize_images(body_html, session, workdir, dest.name)
        src = workdir / "post.html"
        src.write_text(f"<html><head><meta charset='utf-8'></head><body>{body_html}</body></html>",
                       encoding="utf-8")
        tmp_out = dest.with_suffix(".docx.tmp")
        cmd = ["pandoc", str(src), "-f", "html", "-t", "docx", "-o", str(tmp_out), "--no-highlight",
               "--metadata", f"title={title}"]
        if REFERENCE_DOCX:
            cmd += ["--reference-doc", REFERENCE_DOCX]
        try:
            subprocess.run(cmd, check=True, capture_output=True, text=True, cwd=work)
            tmp_out.replace(dest)
        finally:
            tmp_out.unlink(missing_ok=True)


def fix_perms(path: Path, mode: int):
    try:
        if PUID and PGID:
            st = path.stat()
            if (st.st_uid, st.st_gid) != (int(PUID), int(PGID)):
                os.chown(path, int(PUID), int(PGID))
        if (path.stat().st_mode & 0o7777) != mode:
            os.chmod(path, mode)
    except OSError as e:
        print(f"[warn] permissions on {path.name}: {e}", file=sys.stderr)


def run_once():
    if REFERENCE_DOCX and not Path(REFERENCE_DOCX).is_file():
        raise RuntimeError(
            f"REFERENCE_DOCX not found at {REFERENCE_DOCX} (path inside the container; "
            f"the export folder is mounted at {OUT}) - skipping this run")
    OUT.mkdir(parents=True, exist_ok=True)
    fix_perms(OUT, 0o777)
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
                fix_perms(dest, FILE_MODE)
                skipped += 1
                continue
            if prev and prev["file"] != name:                   # title/date changed: drop old file
                (OUT / prev["file"]).unlink(missing_ok=True)

            try:
                to_docx(html.unescape(TAGS.sub("", p["title"]["rendered"])), p["content"]["rendered"], dest, s)
            except subprocess.CalledProcessError as e:
                print(f"[fail] {pid} {name}: {e.stderr.strip()}", file=sys.stderr)
                continue
            fix_perms(dest, FILE_MODE)
            state[pid] = {"modified": p["modified_gmt"], "file": name}
            exported += 1
            print(f"[ok] {name}")

    STATE_FILE.write_text(json.dumps(state, indent=2))
    fix_perms(STATE_FILE, FILE_MODE)
    print(f"exported={exported} unchanged={skipped}")


if __name__ == "__main__":
    print(f"wp2docx start: WP_URL={WP_URL} STATUS={STATUS} OUT={OUT} "
          f"INTERVAL={INTERVAL or 'run-once'} owner={PUID or '-'}:{PGID or '-'} mode={oct(FILE_MODE)} auth={'yes' if WP_USER and WP_APP_PASSWORD else 'no'}")
    while True:
        try:
            run_once()
        except Exception as e:  # keep the loop alive if WP is restarting
            print(f"[error] {e}", file=sys.stderr)
        if not INTERVAL:
            break
        time.sleep(INTERVAL)
