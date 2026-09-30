# wp2docx

Sidecar container that exports every WordPress post to `.docx`, named `YYYY.MM.DD Post Title.docx`.
Pulls posts from the WordPress REST API, converts with pandoc, and only re-exports posts that changed.
If a post's title or date changes, the old file is replaced.

## How it ships
Every push to `main` builds a multi-arch image (amd64/arm64) and publishes it to
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
| `INTERVAL` | Seconds between sweeps (`0` = run once and exit) | `3600` |
| `EXPORT_DIR` | Host folder for the .docx files | `/srv/exports/wordpress` |
| `REFERENCE_DOCX` | Optional Word template for styles (mount it into the container) | — |

Logs show one `[ok] <filename>` per export plus an `exported=N unchanged=M` summary.
