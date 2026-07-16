# ContractVault

A full-stack contract management platform for managing contracts with clients and vendors. Upload templates, create contracts from them, track revisions, compare document changes, and generate finalized PDFs — all with a complete audit trail.

## Run & Operate

- `cd artifacts/contract-manager && python app.py` — run ContractVault (port 8000)
- Workflow: **ContractVault** (managed in Replit workflows pane)

## Stack

- **Backend:** Python 3 + Flask, Flask-Login, Flask-SQLAlchemy
- **Database:** PostgreSQL (Replit built-in) via psycopg2
- **PDF Generation:** ReportLab
- **Document Parsing:** python-docx (for .docx uploads)
- **Diff/Compare:** Python difflib (HtmlDiff)
- **Auth:** Flask-Login + Werkzeug password hashing

## Where things live

```
artifacts/contract-manager/
├── app.py              — Flask app, all routes
├── models.py           — SQLAlchemy models
├── requirements.txt    — Python dependencies
├── uploads/            — Uploaded contract/template files
├── static/
│   ├── css/style.css   — All styles
│   └── js/main.js      — Client-side JS
└── templates/          — Jinja2 HTML templates
    ├── base.html        — App shell with sidebar
    ├── auth/            — Login, register
    ├── dashboard.html   — Overview
    ├── clients/         — Client/vendor management
    ├── templates/       — Contract template management
    └── contracts/       — Contract lifecycle
```

## Database Schema

- `users` — accounts (email, password_hash, full_name)
- `clients` — clients & vendors
- `contract_templates` — reusable templates with `{{FIELD_NAME}}` markers
- `contracts` — contracts with status lifecycle (draft → in_review → approved → finalized)
- `contract_revisions` — version history per contract
- `contract_field_values` — per-contract field fills from template
- `audit_logs` — full timestamped action log

## Template Field System

Templates use `{{FIELD_NAME}}` markers (uppercase, e.g. `{{CLIENT_NAME}}`, `{{START_DATE}}`). When creating a contract from a template, users fill in each field and the app substitutes them into the content.

## Product

- **Auth:** Register/login with email + password
- **Clients & Vendors:** CRUD with contract history
- **Templates:** Upload .txt/.docx or paste text; fields auto-detected
- **Contracts:** Create from template or blank, track status
- **Revisions:** Upload multiple versions; compare any two with side-by-side diff
- **Finalize:** Mark a revision as the official final version
- **PDF Export:** ReportLab-generated professional PDFs for any revision
- **Audit Trail:** Every action timestamped and attributed

## User preferences

_Populate as you build — explicit user instructions worth remembering across sessions._

## Gotchas

- Run from the `artifacts/contract-manager/` directory so relative paths (uploads, templates) resolve correctly
- `DATABASE_URL` is injected automatically by Replit; no manual configuration needed
- Template content uses `{{FIELD_NAME}}` syntax — must be ALL CAPS, no spaces
- Port 8000 (workflow `ContractVault`)
