# Ashish College Connect Backend

FastAPI backend with an Excel `.xlsx` workbook datastore.

## Requirements
Python 3.10+ and VS Code.

## Setup (Windows PowerShell)
Open this folder in VS Code, then Terminal → New Terminal:

```powershell
py -m venv .venv
.\.venv\Scripts\Activate.ps1
python -m pip install --upgrade pip
pip install -r requirements.txt
python -m uvicorn main:app --reload
```

If PowerShell blocks activation, use Command Prompt instead:

```bat
py -m venv .venv
.venv\Scripts\activate.bat
pip install -r requirements.txt
python -m uvicorn main:app --reload
```

Open http://127.0.0.1:8000/docs and http://127.0.0.1:8000/health. On first run the app creates `data/college_data.xlsx` and starter FAQs.

## Demo accounts (local demo only)
- Admin: `admin@college.edu` / `Admin@123`
- Student: `student@college.edu` / `Student@123`

## API endpoints
- `GET /health`
- `GET /api/faqs`
- `GET /api/announcements`
- `POST /api/login`
- `GET /api/me` (Bearer token)
- `POST /api/queries`
- `GET /api/queries` (Bearer token)
- `GET /api/admin/summary` (admin token)
- `GET /api/admin/users` (admin token)
- `GET /api/admin/queries` (admin token)
- `PATCH /api/admin/queries/{query_id}/status` (admin token)
- `POST /api/admin/faqs` (admin token)
- `POST /api/admin/announcements` (admin token)

## Test
In Swagger at `/docs`, use Try it out. For protected endpoints, first call `POST /api/login`, copy `access_token`, click Authorize, and paste `Bearer YOUR_TOKEN` if prompted.

## Salesforce connection later
Salesforce cannot call your laptop's `127.0.0.1` address. First test locally. Then host the API at an HTTPS address (or use an HTTPS tunnel for development), configure Salesforce Trusted URLs/CSP Trusted Sites, and set `CORS_ORIGINS` to your Salesforce origin. Replace the development token secret and demo accounts before any public deployment.

## Excel limitations
This workbook approach is intended for a small class demo. Excel is not a robust multi-user production database. Avoid opening the workbook in Excel while the API is writing to it.
