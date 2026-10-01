#!/usr/bin/env python3
"""Export WordPress posts to .docx, filed as 'YYYY/YYYY.MM.DD Post Title.docx'.

Incremental: only re-converts posts whose modified time changed; moves files whose
title/date changed; in mirror mode removes files for posts deleted or excluded in
WordPress. Runs once, or loops every INTERVAL seconds.
"""
import html
from html.parser import HTMLParser
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
PUID = os.environ.get("PUID", "65534")                                 # Debian/OMV nobody; empty to leave ownership alone
PGID = os.environ.get("PGID", "100")                                   # users
FILE_MODE = int(os.environ.get("FILE_MODE", "0666"), 8)                # empty not allowed; e.g. 0666 or 0777
EXPORT_MODE = os.environ.get("EXPORT_MODE", "mirror").strip().lower()   # mirror: recreate missing files; inbox: deleted stays deleted
EXCLUDE_CATEGORIES = [c.strip() for c in os.environ.get("EXCLUDE_CATEGORIES", "").split(",") if c.strip()]  # slugs or names
EXCLUDE_TITLE = os.environ.get("EXCLUDE_TITLE_REGEX", "").strip()       # case-insensitive, matched against the post title
if EXPORT_MODE not in ("mirror", "inbox"):
    raise SystemExit(f"EXPORT_MODE must be 'mirror' or 'inbox', got {EXPORT_MODE!r}")
EXCLUDE_TITLE_RE = re.compile(EXCLUDE_TITLE, re.I) if EXCLUDE_TITLE else None

STATE_FILE = OUT / ".wp2docx_state.json"
STATE_MODE = 0o644                                                     # never world-writable: it steers moves/deletes
MAX_IMAGE_BYTES = int(os.environ.get("MAX_IMAGE_MB", "25")) * 1024 * 1024
PRUNE_GUARD = 0.5                     # refuse to prune more than this share of tracked posts in one run
PANDOC_ENV = {"PATH": os.environ.get("PATH", "/usr/local/bin:/usr/bin:/bin"), "HOME": "/tmp", "LANG": "C.UTF-8"}
perm_failures = []
ILLEGAL = re.compile(r'[<>:"/\\|?*\x00-\x1f]')
TAGS = re.compile(r"<[^>]+>")


def safe_title(raw: str) -> str:
    t = html.unescape(TAGS.sub("", raw))
    t = ILLEGAL.sub("", t)
    t = re.sub(r"\s+", " ", t).strip().rstrip(". ")
    return t[:180] or "Untitled"


def wp_error(r):
    try:
        err = r.json()
        detail = f"{err.get('code')}: {err.get('message')}"
    except ValueError:
        detail = r.text[:300]
    return RuntimeError(f"{r.status_code} from {r.url} -> {detail}")


def excluded_category_ids(session):
    """Resolve EXCLUDE_CATEGORIES (slugs or names, any case) to WordPress category IDs."""
    if not EXCLUDE_CATEGORIES:
        return []
    cats, page = [], 1
    while True:
        r = session.get(f"{WP_URL}/wp-json/wp/v2/categories", params={"per_page": 100, "page": page}, timeout=60)
        if not r.ok:
            raise wp_error(r)
        cats += r.json()
        if page >= int(r.headers.get("X-WP-TotalPages", 1)):
            break
        page += 1
    ids = []
    for want in EXCLUDE_CATEGORIES:
        hits = [c["id"] for c in cats
                if want.lower() in (c["slug"].lower(), html.unescape(c["name"]).lower())]
        if hits:
            ids += hits
        else:
            print(f"[warn] EXCLUDE_CATEGORIES: no category named {want!r}; known: "
                  + ", ".join(sorted(c["slug"] for c in cats)), file=sys.stderr)
    return ids


def fetch_posts(session, exclude_ids=()):
    auth = (WP_USER, WP_APP_PASSWORD) if WP_USER and WP_APP_PASSWORD else None
    page = 1
    while True:
        r = session.get(
            f"{WP_URL}/wp-json/wp/v2/posts",
            params={"per_page": 100, "page": page, "status": STATUS,
                    "orderby": "date", "order": "asc",
                    "context": "edit" if auth else "view",
                    **({"categories_exclude": ",".join(map(str, exclude_ids))} if exclude_ids else {})},
            auth=auth, timeout=60,
        )
        if not r.ok:
            raise wp_error(r)
        batch = r.json()
        if not batch:
            return
        yield from batch
        if page >= int(r.headers.get("X-WP-TotalPages", 1)):
            return
        page += 1


IMG_EXT = {"image/jpeg": ".jpg", "image/png": ".png", "image/gif": ".gif", "image/webp": ".webp",
           "image/svg+xml": ".svg", "image/bmp": ".bmp", "image/tiff": ".tif"}


def fetch_image(session, url, label):
    """Download an image with a size cap. Returns (bytes, content_type) or None."""
    try:
        with session.get(url, timeout=30, stream=True) as r:
            ctype = r.headers.get("Content-Type", "").split(";")[0].strip().lower()
            if not (r.ok and ctype.startswith("image/")):
                print(f"[warn] {label}: image not usable ({r.status_code} {ctype or 'no type'}): {url}", file=sys.stderr)
                return None
            buf = bytearray()
            for chunk in r.iter_content(65536):
                buf += chunk
                if len(buf) > MAX_IMAGE_BYTES:
                    print(f"[warn] {label}: image over {MAX_IMAGE_BYTES // 1048576} MB, skipped: {url}", file=sys.stderr)
                    return None
            return (bytes(buf), ctype) if buf else None
    except requests.RequestException as e:
        print(f"[warn] {label}: image fetch failed ({e.__class__.__name__}): {url}", file=sys.stderr)
        return None


class ImageRewriter(HTMLParser):
    """Re-emits the post HTML unchanged except: every <img> is downloaded by us and pointed at a
    local file in the work dir (or replaced with a note), and inline <svg> is dropped. Pandoc then
    never resolves a path or URL from post content itself, so it cannot embed container files."""

    def __init__(self, session, workdir, label):
        super().__init__(convert_charrefs=False)
        self.session, self.workdir, self.label = session, workdir, label
        self.out, self.cache, self.svg_depth = [], {}, 0

    def resolve(self, src):
        url = html.unescape(src).strip()
        if PUBLIC_URL and url.startswith(PUBLIC_URL):
            return WP_URL + url[len(PUBLIC_URL):]
        if url.startswith("//"):
            return "https:" + url
        if url.startswith("/"):
            return WP_URL + url
        return url if re.match(r"(?i)https?://", url) else None

    def image(self, attrs):
        a = {k.lower(): (v or "") for k, v in attrs}
        alt = html.escape(a.get("alt", "").strip(), quote=True)
        url = self.resolve(a.get("src", ""))
        if url and url not in self.cache:
            got = fetch_image(self.session, url, self.label)
            if got:
                f = self.workdir / f"img{len(self.cache) + 1}{IMG_EXT.get(got[1], '')}"
                f.write_bytes(got[0])
                self.cache[url] = f.name
            else:
                self.cache[url] = None
        local = self.cache.get(url) if url else None
        if not local:
            if not url:
                print(f"[warn] {self.label}: image with unsupported source skipped: {a.get('src', '')[:120]}", file=sys.stderr)
            what = alt or html.escape(a.get("src", "")[:200])
            return f"<p><em>[image unavailable: {what}]</em></p>"
        return f'<img src="{local}" alt="{alt}">'

    def handle_starttag(self, tag, attrs):
        if tag == "svg":
            self.svg_depth += 1
        if self.svg_depth:
            return
        self.out.append(self.image(attrs) if tag == "img" else self.get_starttag_text())

    def handle_startendtag(self, tag, attrs):
        if self.svg_depth or tag == "svg":
            return
        self.out.append(self.image(attrs) if tag == "img" else self.get_starttag_text())

    def handle_endtag(self, tag):
        if tag == "svg" and self.svg_depth:
            self.svg_depth -= 1
            return
        if not self.svg_depth and tag != "img":
            self.out.append(f"</{tag}>")

    def _raw(self, text):
        if not self.svg_depth:
            self.out.append(text)

    def handle_data(self, d): self._raw(d)
    def handle_entityref(self, n): self._raw(f"&{n};")
    def handle_charref(self, n): self._raw(f"&#{n};")
    def handle_comment(self, d): pass
    def handle_decl(self, d): pass
    def handle_pi(self, d): pass
    def unknown_decl(self, d): pass


def localize_images(body_html: str, session, workdir: Path, label: str) -> str:
    rw = ImageRewriter(session, workdir, label)
    rw.feed(body_html)
    rw.close()
    return "".join(rw.out)


def to_docx(title: str, body_html: str, dest: Path, session):
    with tempfile.TemporaryDirectory() as work:
        workdir = Path(work)
        body_html = localize_images(body_html, session, workdir, dest.name)
        (workdir / "post.html").write_text(
            f"<html><head><meta charset='utf-8'></head><body>{body_html}</body></html>", encoding="utf-8")
        tmp_out = dest.with_suffix(".docx.tmp")
        cmd = ["pandoc", "post.html", "-f", "html", "-t", "docx", "-o", str(tmp_out), "--no-highlight",
               "--resource-path", ".", "--metadata", f"title={title}"]
        if REFERENCE_DOCX:
            cmd += ["--reference-doc", REFERENCE_DOCX]
        try:
            subprocess.run(cmd, check=True, capture_output=True, text=True, cwd=work, env=PANDOC_ENV)
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
        perm_failures.append(f"{path.name}: {e.strerror}")


def in_out(rel: str):
    """Resolve a state-file path, refusing anything outside the export folder."""
    try:
        p = (OUT / rel).resolve()
        return p if p != OUT.resolve() and p.is_relative_to(OUT.resolve()) else None
    except (OSError, ValueError):
        return None


def save_state(state):
    tmp = STATE_FILE.with_suffix(".json.tmp")
    tmp.write_text(json.dumps(state, indent=2))
    os.chmod(tmp, STATE_MODE)
    tmp.replace(STATE_FILE)
    fix_perms(STATE_FILE, STATE_MODE)


def drop_if_empty(folder: Path):
    if folder != OUT and folder.parent == OUT:
        try:
            folder.rmdir()                                      # only succeeds when empty
        except OSError:
            pass


def run_once():
    if REFERENCE_DOCX and not Path(REFERENCE_DOCX).is_file():
        raise RuntimeError(
            f"REFERENCE_DOCX not found at {REFERENCE_DOCX} (path inside the container; "
            f"the export folder is mounted at {OUT}) - skipping this run")
    OUT.mkdir(parents=True, exist_ok=True)
    fix_perms(OUT, 0o777)
    perm_failures.clear()
    try:
        state = json.loads(STATE_FILE.read_text()) if STATE_FILE.exists() else {}
    except ValueError:
        raise RuntimeError(f"{STATE_FILE} is not valid JSON; fix or delete it (deleting forces a full re-export)")
    claimed, seen, dirs_fixed = {}, set(), set()
    exported = skipped = moved = excluded = pruned = 0

    with requests.Session() as s:
        if FORWARDED_PROTO:
            # WordPress only honours Application Passwords over HTTPS; the official
            # image maps X-Forwarded-Proto to $_SERVER['HTTPS'] so plain-http internal calls still count.
            s.headers["X-Forwarded-Proto"] = FORWARDED_PROTO
        for p in fetch_posts(s, excluded_category_ids(s)):
            pid = str(p["id"])
            if EXCLUDE_TITLE_RE and EXCLUDE_TITLE_RE.search(html.unescape(TAGS.sub("", p["title"]["rendered"]))):
                excluded += 1
                continue
            seen.add(pid)
            title = safe_title(p["title"]["rendered"])
            posted = datetime.fromisoformat(p["date"])
            date = posted.strftime("%Y.%m.%d")
            name = f"{date} {title}.docx"
            rel = f"{posted.year:04d}/{name}"                     # files live in YYYY/ subfolders
            if rel in claimed and claimed[rel] != pid:           # same date + title collision
                name = f"{date} {title} ({pid}).docx"
                rel = f"{posted.year:04d}/{name}"
            claimed[rel] = pid
            dest = OUT / rel

            prev = state.get(pid)
            if prev and prev["file"] != rel:
                old = in_out(prev["file"])
                if old is None:
                    print(f"[warn] state entry for post {pid} points outside the export folder; ignored", file=sys.stderr)
                    prev, old = None, None
            if prev and prev["file"] != rel:
                if prev["modified"] == p["modified_gmt"] and EXPORT_MODE == "inbox" and not old.is_file():
                    state[pid] = {"modified": p["modified_gmt"], "file": rel}  # already exported and cleared; just track
                    skipped += 1
                    continue
                if prev["modified"] == p["modified_gmt"] and old.is_file() and not dest.exists():
                    # unchanged post, new location (flat -> year folder, or date moved years): re-file, don't re-convert
                    dest.parent.mkdir(parents=True, exist_ok=True)
                    fix_perms(dest.parent, 0o777)
                    old.rename(dest)
                    drop_if_empty(old.parent)
                    fix_perms(dest, FILE_MODE)
                    state[pid] = {"modified": p["modified_gmt"], "file": rel}
                    moved += 1
                    continue
                old.unlink(missing_ok=True)                      # title/date changed: drop the old copy
                drop_if_empty(old.parent)
            elif prev and prev["modified"] == p["modified_gmt"] and (dest.exists() or EXPORT_MODE == "inbox"):
                if dest.exists():
                    if dest.parent not in dirs_fixed:
                        fix_perms(dest.parent, 0o777)
                        dirs_fixed.add(dest.parent)
                    fix_perms(dest, FILE_MODE)
                skipped += 1
                continue

            dest.parent.mkdir(parents=True, exist_ok=True)
            fix_perms(dest.parent, 0o777)
            try:
                to_docx(html.unescape(TAGS.sub("", p["title"]["rendered"])), p["content"]["rendered"], dest, s)
            except subprocess.CalledProcessError as e:
                print(f"[fail] {pid} {rel}: {e.stderr.strip()}", file=sys.stderr)
                continue
            fix_perms(dest, FILE_MODE)
            state[pid] = {"modified": p["modified_gmt"], "file": rel}
            exported += 1
            print(f"[ok] {rel}")

    # posts deleted in WordPress or newly excluded: forget them (mirror mode also removes their files)
    gone = [pid for pid in state if pid not in seen]
    if gone and (not seen or len(gone) > PRUNE_GUARD * len(state)):
        print(f"[warn] {len(gone)} of {len(state)} tracked posts missing from this run; not pruning "
              f"(check WP_STATUS, credentials and exclusions)", file=sys.stderr)
    else:
        for pid in gone:
            f = in_out(state[pid].get("file", ""))
            if EXPORT_MODE == "mirror" and f and f.is_file():
                f.unlink()
                drop_if_empty(f.parent)
            del state[pid]
            pruned += 1

    save_state(state)
    if perm_failures:
        print(f"[warn] could not set owner/mode on {len(perm_failures)} item(s), e.g. {perm_failures[0]}", file=sys.stderr)
    print(f"exported={exported} refiled={moved} unchanged={skipped} pruned={pruned}"
          + (f" excluded_by_title={excluded}" if EXCLUDE_TITLE_RE else ""))


if __name__ == "__main__":
    print(f"wp2docx start: WP_URL={WP_URL} STATUS={STATUS} OUT={OUT} "
          f"INTERVAL={INTERVAL or 'run-once'} mode={EXPORT_MODE} exclude_categories={EXCLUDE_CATEGORIES or '-'} "
          f"exclude_title={EXCLUDE_TITLE or '-'} owner={PUID or '-'}:{PGID or '-'} mode={oct(FILE_MODE)} auth={'yes' if WP_USER and WP_APP_PASSWORD else 'no'}")
    while True:
        try:
            run_once()
        except Exception as e:  # keep the loop alive if WP is restarting
            print(f"[error] {e}", file=sys.stderr)
        if not INTERVAL:
            break
        time.sleep(INTERVAL)
