# PSC Coupon Management (Standalone)

This folder is a standalone extraction of the **PSC coupon / “coupen” management** feature (admin + public scratch/lottery entry).

## What’s included

- Admin UI: `/admin` (coupon names, masters, validators, recent users)
- Public UI: `/coupon/entry?cn=<barcode_value>` (scan/entry page)
- APIs:
  - `POST /coupon/validate-mobile`
  - `POST /coupon/register`
  - `POST /coupon/reveal-code`

## Local run

```powershell
cd psc_coupens
python -m venv .venv
.\.venv\Scripts\pip install -r requirements.txt
Copy-Item .env.example .env
.\.venv\Scripts\python scripts\init_db.py
.\.venv\Scripts\python app.py
```

Then open:
- Admin: `http://127.0.0.1:5000/admin` (password = `ADMIN_PASSWORD`)
- Public entry: `http://127.0.0.1:5000/coupon/entry`

## Notes on hosting

This project is a **Flask + database** app. GitHub Pages supports **static files only**, so you can’t host this backend on GitHub Pages without redesigning it into:
- Static frontend on GitHub Pages, **plus** a separate backend API (Render/Fly/Railway/Azure/etc.), or
- A serverless backend (Supabase/Firebase/Cloudflare Workers, etc.).

### Free hosting (SQLite)

If you want to start with **SQLite**, pick a host that:
- runs **one instance** of your app (no horizontal scaling), and
- has a **persistent filesystem** (SQLite is a file).

In practice, **PythonAnywhere** is usually the simplest “free tier” fit for SQLite-based Flask apps.

### Later (Postgres)

When you move to Postgres, you can deploy the same Flask app to most PaaS hosts (Render/Railway/Fly/etc.) and just change `DATABASE_URL`.

### Render quick notes

- Start command (this repo layout): `gunicorn --chdir psc_coupens app:app`
- Set env vars: `SECRET_KEY`, `ADMIN_PASSWORD`, `DATABASE_URL`
- One-time init: run `python psc_coupens/scripts/init_db.py` in the Render Shell to create tables.
- Note: Render **free Postgres** is not meant for long-term persistence (can expire). Use dumps below.

## Backup / migration (Postgres, “no owner”)

To make restores portable across providers, export without ownership/ACL statements:

- Dump (plain SQL):
  - `pg_dump --no-owner --no-acl "$DATABASE_URL" > backup.sql`
- Dump (custom format):
  - `pg_dump --format=custom --no-owner --no-acl "$DATABASE_URL" > backup.dump`
- Restore:
  - `psql "$DATABASE_URL" -f backup.sql`
  - or `pg_restore --no-owner --no-acl --clean --if-exists -d "$DATABASE_URL" backup.dump`
