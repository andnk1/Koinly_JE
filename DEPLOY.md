# Deploying to Railway

This app lives in the `webapp/` subfolder of the `Koinly_JE` repo
(`https://github.com/andnk1/Koinly_JE.git`), alongside `koinly_to_je.py` and
your sample data at the repo root. Railway needs to know to build/run from
that subfolder.

## One-time setup

1. Commit and push `webapp/` to the repo (see below).
2. In Railway: **New Project -> Deploy from GitHub repo** -> pick `Koinly_JE`.
3. Open the new service's **Settings** tab:
   - **Root Directory**: `webapp`
   - **Build**: leave on Nixpacks/auto -- it will detect `requirements.txt`
     and install it automatically.
   - **Start Command**: leave blank (the `Procfile` in `webapp/` already says
     `gunicorn app:app --bind 0.0.0.0:$PORT`), or set it explicitly to that
     same command if Railway doesn't pick up the Procfile.
4. Deploy. Railway assigns a public URL under **Settings -> Networking ->
   Generate Domain** if one isn't already there.

## Notes

- No environment variables are required -- there's no password gate on this
  app (by request). If that changes later, add a `APP_PASSWORD` env var and
  a small login check in `app.py`, the same pattern used in the Bank
  Reconciliation app.
- The app is stateless: nothing uploaded is written to disk, and generated
  reports live only in memory for about 15 minutes (just long enough to
  download them) or until the service restarts/redeploys, whichever comes
  first. There's no database and nothing to back up.
- Local test before pushing:
  ```
  cd webapp
  pip install -r requirements.txt
  python app.py        # http://localhost:5000
  ```
