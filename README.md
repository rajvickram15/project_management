# Project Management Tool (Employee + Scrum Master)

This is a lightweight project management tool with **two separate logins** and a **central SQL database** (Flask + SQLite by default):

- **Employee**: daily update form (last workday, current updates, impediments, leave).
- **Scrum Master**: daily notes form (production, test system, follow-ups) + report export.

## Run (Flask backend + SQL)

1) Install dependencies:

```powershell
python -m pip install -r requirements.txt
```

2) Start the server:

```powershell
python app.py
```

3) Open:

- `http://127.0.0.1:5000/`

Important: don’t double-click `index.html`. Always open the app via the Flask URL above.

## Excel Import (Scrum Master)

On the Scrum Master dashboard, open **Employee Records** → **Import Excel (sample)** and upload your `.xlsx`.

Expected columns (header names can vary, the importer matches by keywords):
- Date
- Team Member / Member / Name
- Leave Type (optional)
- Last workday updates
- Current workday updates
- Impediments/obstacles
- Tags/Lifecycle (optional)

### Database

- Default DB is SQLite file: `pmtool.db` (created automatically).
- To use another SQL DB later (Postgres/MySQL), set `DATABASE_URL` before running.

### Demo data

On first run (empty DB), the app seeds sample **Release Calendar** deadlines for projects `NOVA` and `PILOT-X`.

## Demo logins

- Employee: `employee@company.com` / `Employee@123`
- Scrum Master: `scrum@company.com` / `Scrum@123`

You can change demo credentials in `backend/models.py` (`ensure_seed_users()`).
