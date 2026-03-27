# PSC Coupons (Scan & Win) — Coupon + Scratch Card System

**Owner:** Manish Sidhenkiwar

PSC Coupons is a **mobile-first Flask web app** that supports:
- **Admin panel** to manage coupon campaigns, coupon codes, validators, prize levels, exports, and QR codes.
- **Public “Scan & Win” flow** where a user validates **mobile number + coupon code**, scratches a card, and reveals a prize with a **reference code** for prize distribution.

> Note: The project name uses the spelling “coupen” in many file names and UI labels (kept as-is for compatibility).

## Key Features

### Public (User) flow
- Entry page at `/coupon/entry` (the site root `/` redirects here).
- Mobile number validation against **validators list**.
- Coupon code validation (coupon code can be used by multiple users overall, but **only once per mobile number**).
- Scratch card reveal with prize image + reference code (copyable), optimized for mobile UX.

### Admin panel
- Admin login at `/admin/login` (main dashboard at `/admin`).
- Manage:
  - **Coupon Names (campaigns)** + QR code generation for each campaign.
  - **Coupon Codes** under a coupon name (generate, update status, delete unused).
  - **Coupon Masters (prizes)**: type/name, level, weight, max limit, status, image upload/view.
  - **Validators**: CRUD, upload validators via Excel, filtering/pagination.
  - **Recent Coupon Users**: filtering, pagination, delete (with audit safety), exports.
- Excel exports:
  - Validators export: `/admin/export/validators`
  - Users export: `/admin/export/users` (supports grouped summary + raw rows)

## Prize Distribution (Backend Rules)

This app supports **7 prize levels** with **max allowed count per master row** (per prize item).

High-level behavior:
- **Levels 4–7** are always available and selected randomly using weighted probability.
- After **first 1000 total scratches** in a campaign, **Levels 1–3 unlock**.
- Between **1001–1500 (inclusive)**:
  - Level 1: allow **1** prize
  - Level 2: allow **2** prizes
  - Level 3: allow **2** prizes
- After **1500+ scratches**:
  - Level 1: allow **1 more** (total 2)
  - Levels 2 & 3: remaining prizes can be awarded (still capped by each master row’s `max_allowed`)
- Selection uses: `master.weight × level_multiplier` (multipliers are editable in Admin → Prize Distribution Settings).
- The backend enforces caps so that **no level/prize exceeds its max allowed**, even when multiple users reveal prizes concurrently.

## Project Structure

- `app.py` — Flask app entry point
- `psc_coupens_app/` — application package
  - `routes.py` — main routes (admin + public APIs)
  - `models.py` — SQLAlchemy models
  - `templates/` — Jinja templates (admin + public)
  - `static/` — CSS/JS/images
- `scripts/init_db.py` — initialize database schema/tables
- `render.yaml` — Render deployment configuration

## Local Setup (Windows / PowerShell)

```powershell
python -m venv venv
.\venv\Scripts\pip install -r requirements.txt
Copy-Item .env.example .env
.\venv\Scripts\python scripts\init_db.py
.\venv\Scripts\python app.py
```

Open:
- Public entry: `http://127.0.0.1:5000/` (redirects to `/coupon/entry`)
- Admin: `http://127.0.0.1:5000/admin` (login via `/admin/login`)

## Environment Variables

Use `.env.example` as reference. Common values:
- `SECRET_KEY` — Flask session secret
- `ADMIN_PASSWORD` — admin login password
- `DATABASE_URL` — DB connection string (SQLite or Postgres)

## Deploy on Render (Free Tier)

This repo includes `render.yaml`. On Render:
- Create a new Web Service from your GitHub repo.
- Set environment variables (`SECRET_KEY`, `ADMIN_PASSWORD`, `DATABASE_URL`).
- Run DB init once in Render Shell:
  - `python scripts/init_db.py`

## Support / Maintenance Notes

- If you rename routes or change templates, also verify QR code links still open the correct deployed URL.
- When deleting users from admin, ensure prize audit rows stay consistent (audit tables may enforce `NOT NULL` foreign keys).

