# Rack & ID — Clothing Catalog

An internal tool for a clothing store: add product photos tagged with the
store's existing ID, then browse or search the catalog by ID or name.
Works on phone and laptop. Two roles — staff enter a shared viewer
password to browse, and a separate admin password unlocks add/edit/delete.
Supports English and Arabic (right-to-left).

**Data storage:** the catalog database lives on **Supabase** (free
Postgres) and photos live on **Cloudflare R2** (free object storage, up to
10GB). Both are separate from wherever you host the app itself, which
means you can update and redeploy the app's code anytime without touching
or losing the store's data — including on a free hosting tier, since
nothing important lives on the app server's own disk.

## What's included

- **Two roles:** viewer (enters a shared password to browse/search) and
  admin (separate password — add, edit, delete). Nobody sees the catalog
  without a password.
- **Catalog page** — responsive, searchable gallery. Search matches ID,
  name, or category. Click any photo to see it full-size. Cards with more
  than one photo show a count badge and preview the other photos on
  hover.
- **Add / Edit item** (admin only) — upload one or more photos with the
  item's ID from the store's existing system, plus optional
  name/category. Editing lets you change name/category, remove photos,
  and add more. The first photo is the one shown on the catalog grid.
- **Item page** — photo gallery with arrows, thumbnails, keyboard and
  swipe. Buttons to **Print** the item (browser "Save as PDF" works too)
  and to **Save as image** — a shareable product card (photo + ID +
  name), generated server-side, Arabic text included.
- **Export** (admin only) — from the catalog page, download the whole
  catalog as **Excel** (`.xlsx`, formatted, with clickable photo links),
  a **ZIP** (a `catalog.csv` plus every photo file, named by item ID),
  plain **CSV**, or **JSON**. The JSON and ZIP double as backups.
- **Language switch** — EN/AR toggle in the nav, with full right-to-left
  layout for Arabic.
- Thumbnails generated automatically for fast loading even with
  thousands of items.

## One-time setup: create your free accounts

You need two free accounts before this will run. Takes about 10 minutes
total.

### 1. Supabase (database)

1. Go to https://supabase.com and sign up (free, no card required).
2. Create a new project — pick any name/region, and set a database
   password (save it somewhere, you'll need it in a second).
3. Once the project is ready, go to **Project Settings → Database →
   Connection string → URI**. Copy it — it looks like:
   `postgresql://postgres:[YOUR-PASSWORD]@db.xxxxxxxx.supabase.co:5432/postgres`
4. Replace `[YOUR-PASSWORD]` in that string with the real password you set
   in step 2.
5. This full string is your `DATABASE_URL`.

### 2. Cloudflare R2 (photo storage)

1. Go to https://dash.cloudflare.com and sign up (free).
2. In the sidebar, go to **R2 Object Storage** → create a bucket (any
   name, e.g. `clothing-catalog`).
3. Open the bucket → **Settings** → under "Public access", enable the
   **public development URL**. Copy that URL — it looks like
   `https://pub-xxxxxxxxxxxx.r2.dev`. This is your `R2_PUBLIC_URL`.
4. Go back to the R2 overview page → **Manage R2 API Tokens** → **Create
   API Token**. Give it **Object Read & Write** permission, scoped to
   your bucket. Create it, then copy the **Access Key ID** and **Secret
   Access Key** it shows you (shown only once — save them now).
5. Your **Account ID** is shown on the same R2 overview page, in the
   right sidebar.

You now have everything needed: `DATABASE_URL`, `R2_ACCOUNT_ID`,
`R2_ACCESS_KEY_ID`, `R2_SECRET_ACCESS_KEY`, `R2_BUCKET_NAME`, and
`R2_PUBLIC_URL`.

## Running it locally

1. Copy `.env.example` to a new file named `.env` in the same folder.
2. Fill in the values you collected above, plus set your own
   `ADMIN_PASSWORD`, a `VIEWER_PASSWORD` (the password everyone needs just
   to view the catalog), and a random `SECRET_KEY`.
3. Then:

```bash
cd clothing-catalog
python -m venv venv
venv\Scripts\activate        # On Mac/Linux: source venv/bin/activate
pip install -r requirements.txt
python app.py
```

Open http://localhost:5000. On your phone (same wifi), use your
computer's local IP instead of localhost.

## Publishing it online

Since the database and photos now live on Supabase/R2 rather than the
app server, you can safely use a **free** web hosting tier for the app
itself (e.g. Render, Railway, PythonAnywhere) — redeploys and restarts
won't touch the store's data.

Rough steps for Render (free tier):

1. Push this project to a GitHub repo (the `.gitignore` already excludes
   your `.env` file, so your secrets won't be uploaded).
2. On Render, create a new **Web Service**, connect the repo.
3. Build command: `pip install -r requirements.txt`
4. Start command: `gunicorn app:app` (add `gunicorn` to
   `requirements.txt` first).
5. Under **Environment**, add each variable from your `.env` file
   (`DATABASE_URL`, `R2_ACCOUNT_ID`, etc.) — this is how the live app
   gets its credentials, since the `.env` file itself never gets
   uploaded.

After that, whenever you want to ship an update: push your code changes
to GitHub, Render redeploys automatically, and the store's catalog is
untouched throughout.

## Updating the app later

Since your data lives on Supabase/R2 and not on the app server, you can
freely edit code, test locally, and redeploy at any time — the owner's
catalog stays intact. The only thing to be careful with: don't change
the `items` table's column names/types directly in Supabase without
also updating the matching code in `app.py`.

The first time this version starts (locally or on the server) it creates
an `item_images` table and copies each existing product's current photo
into it as that product's first image — nothing is lost, and the step is
safe to run repeatedly.

## Backups and restore

**Manual:** on the catalog page (as admin) use **Export → JSON** (catalog
data) or **ZIP + photos** (data plus every image file). Keep those
somewhere safe — Google Drive, etc.

**Automatic:** `backup.py` writes a JSON snapshot to R2 under `backups/`
(`backups/latest.json` plus timestamped copies, keeping the newest 30).
To run it on a schedule, add a **Cron Job** on Render pointing at this
repo:

- Build command: `pip install -r requirements.txt`
- Command: `python backup.py`
- Schedule: e.g. `0 2 * * *` (daily, 02:00 UTC)
- Environment: the same `DATABASE_URL` and `R2_*` variables as the web
  service.

**Restore:** on the catalog page use **Import / restore**. Upload a JSON
or ZIP export, or click **Restore the latest automatic backup** (reads
`backups/latest.json`). Import only *adds* items whose ID isn't already
in the catalog — it never changes or deletes anything that's already
there. A JSON restore recreates the catalog rows and expects the images
to still be in R2; a ZIP restore also re-uploads the image files.

## A note on scale

Supabase's free tier and R2's free tier both comfortably handle a
catalog of several thousand items with photos. Worth revisiting pricing
only if the store grows well beyond that.

## Security note before going live

- Set `ADMIN_PASSWORD` to something only you and the store owner know.
- Set `VIEWER_PASSWORD` to the password staff use just to view the
  catalog. Anyone without it only ever sees the sign-in screen. The code
  ships with a placeholder default — always set a real one via the
  environment (locally in `.env`, on the host in its env vars).
- Set `SECRET_KEY` to a long random string — e.g. output from
  `python -c "import secrets; print(secrets.token_hex(32))"`.
- These are two shared passwords, not individual staff accounts — fine
  for one store owner. If multiple staff need separate logins later,
  that's a bigger change worth revisiting.
