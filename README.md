# wp2docx

Sidecar container that exports every WordPress post to `.docx`, filed as `YYYY/YYYY.MM.DD Post Title.docx`.
Pulls posts from the WordPress REST API, converts with pandoc, and only re-exports posts that changed.
If a post's title or date changes, the old file is replaced; if only its date's year changes, the file is moved.
Exports from older flat-layout versions are moved into year folders on first run, not re-converted.

## How it ships
Every push to `main`, and a weekly scheduled run (for base-image security updates), builds a multi-arch image (amd64/arm64) and publishes it to
`ghcr.io/<owner>/wp2docx:latest` via GitHub Actions.

## Deploy (Dockhand)
Create a stack from `compose.yaml` and set these environment variables:

| Variable | Purpose | Default |
|---|---|---|
| `GHCR_OWNER` | GitHub username (lowercase) | — |
| `WP_NETWORK` | Docker network of the WordPress stack | `wordpress_default` |
| `WP_URL` | WordPress URL as seen from this container | `http://wordpress` |
| `PUBLIC_URL` | WordPress Site Address, rewritten to `WP_URL` so images embed | — |
| `WP_USER` / `WP_APP_PASSWORD` | Application Password (needed for drafts/private) | — |
| `WP_STATUS` | `publish`, or e.g. `publish,draft,private` | `publish` |
| `WP_FORWARDED_PROTO` | Header telling WordPress the call is HTTPS, so Application Passwords work over internal http (empty to disable) | `https` |
| `INTERVAL` | Seconds between sweeps (`0` = run once and exit) | `3600` |
| `PUID` / `PGID` | Numeric owner applied to every export (Debian/OMV nobody:users = 65534:100; Unraid = 99:100; empty to skip) | `65534` / `100` |
| `FILE_MODE` | Octal mode for exported files (folder always gets 0777) | `0666` |
| `EXPORT_MODE` | `mirror`: the folder matches WordPress; missing files are recreated, files for deleted/excluded posts are removed. `inbox`: the tracker decides; cleared files stay gone and only new/edited posts appear | `mirror` |
| `EXCLUDE_CATEGORIES` | Comma-separated category slugs or names never exported (e.g. `workouts`) | — |
| `EXCLUDE_TITLE_REGEX` | Case-insensitive regex; posts whose title matches are never exported (e.g. `^workout`) | — |
| `EXPORT_DIR` | Host folder for the .docx files | `/srv/exports/wordpress` |
| `MAX_IMAGE_MB` | Largest image that will be downloaded and embedded | `25` |
| `REFERENCE_DOCX` | Optional Word template for styles (mount it into the container) | — |

Logs open with the effective settings, then one `[ok] <file>` per export and an
`exported=N refiled=N unchanged=N pruned=N` summary per pass. `[warn]` lines flag dead images,
unknown categories and permission problems.

## Security notes
- Posts are converted with pandoc. Every `<img>` is downloaded by the exporter itself (size-capped) and
  passed to pandoc as a local file; inline SVG is dropped. Pandoc never resolves paths or URLs from post
  content and runs with a minimal environment, so it cannot embed container files or credentials.
- Paths in `.wp2docx_state.json` are confined to the export folder; the file itself is `0644`.
- `compose.yaml` runs read-only with all capabilities dropped except `CHOWN`/`FOWNER` (needed for
  `PUID`/`PGID`/`FILE_MODE`) and `no-new-privileges`.
- Mirror-mode pruning is skipped if a run would remove more than half the tracked posts.

Images are downloaded by the exporter and embedded in the .docx. Anything that isn't a real image
(dead hotlinks, 404 pages) is replaced with `[image unavailable: <alt or URL>]` and logged as `[warn]`.
