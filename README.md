# EasyAnalysis by Panta Development and AI

EasyAnalysis is a production-oriented workspace for cleaning operational loan reports and generating RO performance dashboards.

## Report preparation and arrears comparison

Every source report is checked for a single title row before cleaning. A title is skipped only when the first row contains one non-empty title cell and the second contains recognizable column labels. Reports that already start with labels are read normally. Targets workbooks bypass this check.

Loan Arrears Difference accepts two arrears CSV snapshots and the usual Excel targets workbook. Include a different date in each CSV filename, such as `LoanArrears-2026-10-01.csv` and `LoanArrears-2026-10-16.csv`; upload order does not matter. The older snapshot is the beginning balance. The supplied LoanArrearsDiff notebook's officer matching (60% token-sort threshold), targets master list, change calculations, ranking, and three formatted worksheets are preserved. An Executive Summary sheet also includes the notebook's summary figures.

Run regression checks from the repository root with `python -m unittest discover -s backend/tests -v`, and build the frontend with `npm run build --prefix frontend`.

## Run locally

### Backend

```powershell
cd backend
python -m venv .venv
.\.venv\Scripts\Activate.ps1
pip install -r requirements.txt
uvicorn backend.main:app --reload --app-dir ..
```

### Frontend

```powershell
cd frontend
npm install
npm run dev
```

Set `VITE_API_URL` for a deployed backend. The backend is ready for Railway with the included `Dockerfile`; the frontend is ready for Netlify with `npm run build` and publish directory `dist`.

## Deployment

1. Create a Railway PostgreSQL database.
2. Deploy the project backend to Railway using the `backend/Dockerfile`.
3. Set Railway variables from `backend/.env.example`: `DATABASE_URL`, `FRONTEND_URL`, and a strong `AUTH_SECRET`.
4. In Netlify, set `VITE_API_URL` to the Railway backend URL.
5. Use `npm run build` as the Netlify build command and `dist` as the publish directory.
6. After Netlify gives you the final site URL, update Railway's `FRONTEND_URL` to that exact URL.

Netlify cannot run the Python processing API itself. The browser needs the Netlify frontend and Railway backend to be deployed separately.
