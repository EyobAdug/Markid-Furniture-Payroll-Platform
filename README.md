# Markid Furniture Payroll

Contract-worker and main-employee payroll (no tax/pension in V1) with approval workflow, audit log, Excel reports,
**offline desktop apps (Windows + Mac)** and an **online server** that keeps everyone in sync.

| Piece | What it is |
|---|---|
| `app.py` + `schema.sql` | The whole application (Flask + SQLite) |
| online server | Same app run with `MARKID_MODE=server` (Docker / Render). The admin sees everything here |
| desktop app | `desktop.py` packaged by PyInstaller for Windows/Mac. Works offline, syncs when the internet is back |
| `.github/workflows` | Tests on every push; builds the Windows/Mac apps on demand |

**Start here: [DEPLOY.md](DEPLOY.md)** (GitHub -> online server -> desktop apps -> sync, in plain steps).

## Run it on your own computer (development)
    pip install -r requirements.txt
    python app.py            # http://localhost:5000  (admin / admin123)
    python test_smoke.py     # payroll rules, workflow, exports
    python test_sync.py      # offline -> online sync, conflicts, idempotency

## Modes (environment variables)
- `MARKID_MODE=standalone` (default) - single computer, no sync
- `MARKID_MODE=server` - online server. Needs `SYNC_TOKEN`, `ADMIN_PASSWORD`, `SECRET_KEY`; `MARKID_DB=/data/markid.db` on a persistent disk
- `MARKID_MODE=client` - what the desktop app uses; records offline work and syncs it
