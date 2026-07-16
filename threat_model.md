# Threat Model

## Project Overview

ContractVault is a Python 3 / Flask web application for contract lifecycle management. Users register and log in with email+password, then create clients/vendors, upload contract templates (with `{{FIELD}}` substitution), create contracts from templates, manage revisions, compare document versions side-by-side, generate PDFs, and view a full audit trail. The database is PostgreSQL via SQLAlchemy. There is no external auth provider; credentials are stored locally with Werkzeug password hashing (PBKDF2/scrypt).

## Assets

- **User credentials** — email addresses and Werkzeug-hashed passwords stored in `users`. Compromise allows account takeover.
- **Contract content** — potentially confidential legal and business agreements in `contracts` and `contract_revisions`. Unauthorized access or modification is a core business risk.
- **Client/vendor PII** — names, emails, phone numbers, addresses in `clients`.
- **Flask session signing key** — if weak or leaked, enables forging session cookies for any user.
- **Uploaded files** — stored under `artifacts/contract-manager/uploads/`; could contain sensitive document content.

## Trust Boundaries

- **Browser ↔ Flask server** — all form submissions and page requests cross here. The server must authenticate and authorize every state-changing request. Currently no CSRF tokens are enforced.
- **Flask server ↔ PostgreSQL** — SQLAlchemy ORM is used throughout; no raw SQL string concatenation was found. Parameterized queries via SQLAlchemy protect against SQL injection.
- **Public ↔ Authenticated** — all substantive routes require `@login_required`. Registration is open (no invite gate).
- **User ↔ User (multi-tenancy)** — no ownership isolation exists today. Any authenticated user can read and modify any other user's data.

## Scan Anchors

- **Entry point:** `artifacts/contract-manager/app.py` — all routes defined here.
- **Highest-risk areas:** login/register (authentication), compare view (XSS via `| safe`), finalize/PDF/compare routes (IDOR), CSRF absence across all POST routes.
- **Public surface:** `/login`, `/register` — no auth required.
- **Authenticated surface:** everything else — all under `@login_required`.
- **No admin surface** — `role` field exists in `User` model but is never checked.
- **Dev-only:** `artifacts/mockup-sandbox/` (Canvas/mockup artifact) — not reachable in production Flask app.

## Threat Categories

### Spoofing

Flask sessions are signed with `SECRET_KEY`. The fallback value `'dev-secret-key-change-in-prod'` is hardcoded in source and publicly visible. A deployment without `SESSION_SECRET` set in the environment is immediately vulnerable to session cookie forgery, allowing an attacker to authenticate as any user (including the first-registered admin) without credentials. **Guarantee required:** `SESSION_SECRET` must be a cryptographically random value set in the environment; no fallback should be accepted at startup.

### Tampering

All state-changing routes (create, edit, delete, status change, finalize) lack CSRF tokens. Any page a logged-in user visits can silently trigger arbitrary POST requests to ContractVault using the user's session cookie. **Guarantee required:** Every HTML form must include a per-session CSRF token validated server-side before processing.

Contract revision IDs supplied by users in the compare, finalize, and PDF routes are not checked against the parent contract. An attacker can substitute a revision ID from another contract to read its content or finalize it. **Guarantee required:** All revision lookups must assert `revision.contract_id == contract_id`.

### Information Disclosure

The contract comparison page renders `diff_html | safe`, inserting user-supplied contract text verbatim into HTML without escaping. This enables stored XSS: an authenticated user who embeds a `<script>` tag in a contract revision can steal session cookies from any user who views the diff. **Guarantee required:** Contract content must be HTML-escaped before diff generation, or the diff output must be sanitized before rendering.

The `next` query parameter on `/login` is accepted without origin validation, enabling open redirect attacks. **Guarantee required:** `next` must be validated as a same-origin relative path.

### Elevation of Privilege

No authorization model exists beyond authentication. Any registered user can read, modify, and delete all contracts, clients, templates, and audit logs — regardless of who created them. The `User.role` field is defined but never consulted by any route handler. **Guarantee required:** Implement ownership filtering on all data queries and verify ownership before any mutation, returning 403 on violations.

### Denial of Service

No rate limiting is applied to `/login` or `/register`, allowing credential brute-force and account creation spam. No captcha or lockout mechanism is present. **Guarantee required:** Apply rate limiting (e.g., via Flask-Limiter) to authentication endpoints.
