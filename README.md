# Legal · Rappi · Money

Automation toolkit for Rappi's Legal team. The repository hosts **two
independent Google Apps Script projects** that share utility code:

| Project | Entry point | Purpose |
|---|---|---|
| **RappiMind** — T&C Generator | `WebApp.Html`, `Admin.gs`, `Setupg.gs` | Generate compliant Terms & Conditions documents for marketing campaigns (Cashback, contests, etc.) across LATAM countries. |
| **Legal Team Tracker** | `Codigo.gs` (+ `Dashboard.html` — *deployed separately, not in this repo*) | Track legal team's projects, tasks, SLAs and weekly OKRs in a kanban-style dashboard. |

> Both projects are deployed as separate Apps Script Web Apps. They
> share **only** the helper modules below.

---

## Repository map

```
.
├── Config.gs                       — Constants + Property loaders (RappiMind)
├── Helper.gs                       — Pure utilities, sheet I/O helpers
├── Security.gs                     — Validation & sanitisation library
├── Admin.gs                        — RappiMind admin RPCs + Gemini wizard
├── Setupg.gs                       — One-shot setup / seeding utilities
├── Codigo.gs                       — Legal Team Tracker backend (separate project)
├── WebApp.Html                     — RappiMind front-end (Web App)
├── Propuesta de ajuste del front   — Draft dark-theme UI proposal (not deployed)
└── README.md                       — You are here
```

### Shared modules

* **`Helper.gs`** — date/number formatting, Spanish number-to-letters
  (`numeroALetras`), `_getSheet` / `_getOrCreateSheet`, response builder.
* **`Security.gs`** — `htmlEscape`, email/URL/country/number validation,
  formula-injection guard (`sheetCellSafe`), Script Property lookups.

---

## RappiMind — T&C Generator

A multi-step web form that lets a Rappi KAM (Key Account Manager) fill
in a campaign brief; the backend renders a Google Doc from a country +
campaign-type template, replacing `{{PLACEHOLDERS}}` with both raw and
*derived* fields (see `Config.gs → DERIVED_FIELDS`).

### Architecture

```
┌────────────────────────────────────────────────────────────────┐
│ Front-end · WebApp.Html (Tailwind + vanilla JS)                │
│   • Country picker → dynamic field loader → live preview       │
│   • Embedded admin panel (Admin.gs RPCs)                       │
│   • Floating Gemini chat helper                                │
└──────────────────────────┬─────────────────────────────────────┘
                google.script.run
                           │
┌──────────────────────────▼─────────────────────────────────────┐
│ Apps Script · Admin.gs / processWebPayload (external)          │
│   • Auth via _requireRole() against `Admin_Team` sheet         │
│   • Template lookup in `Template_Registry`                     │
│   • Field schema from `Template_Fields`                        │
│   • Country legal defaults from `Country_Settings`             │
│   • Gemini API for placeholder detection (Wizard)              │
└──────────────────────────┬─────────────────────────────────────┘
                           │
┌──────────────────────────▼─────────────────────────────────────┐
│ Google services                                                │
│   • Sheets — audit logs, templates registry, fields catalogue  │
│   • Drive  — country/type folder hierarchy + generated Docs    │
│   • Docs   — placeholder replacement                           │
└────────────────────────────────────────────────────────────────┘
```

> ⚠️ The `processWebPayload`, `askGemini`, `getCampaignTypesForUser`
> and `getFieldsForUserForm` functions are defined in the deployed
> Apps Script project but **not currently checked into this repo**.
> If you cloned this and the Web App throws "ReferenceError", add
> them back from the Apps Script editor or restore from version
> history.

### Spreadsheet tabs

| Sheet | Owner | Description |
|---|---|---|
| `Template_Registry` | Admin | One row per (country, campaign type, version). Tracks status (`draft`, `pending_review`, `approved`, `rejected`, `active`, `inactive`) and the source Google Doc ID. |
| `Template_Fields` | Admin | Field catalogue rendered by the dynamic form (label, type, validation, options, default). |
| `Campaign_Types` | Admin | Catalog of dinámicas (campaign types) with icon/color/description. |
| `Country_Settings` | Admin | Per-country legal defaults: jurisdiction, currency, applicable law, privacy URLs, etc. — automatically inlined into templates. |
| `Admin_Team` | Owner | RBAC table. Roles: `viewer < editor < admin < owner`. |
| `Approval_Log` | System | Append-only audit trail of admin actions. |
| `Respuestas_Audit_V2` | System | Every generated document is logged here for traceability. |

### Roles & permissions

| Role | View admin | Edit fields | Create templates | Approve | Activate | Delete | Manage team |
|---|---|---|---|---|---|---|---|
| viewer | ✅ | – | – | – | – | – | – |
| editor | ✅ | ✅ | ✅ | – | – | – | – |
| admin  | ✅ | ✅ | ✅ | ✅ | ✅ | ✅ | – |
| owner  | ✅ | ✅ | ✅ | ✅ | ✅ | ✅ | ✅ |

Enforced server-side via `_requireRole()` (Admin.gs).

### Configuration via Script Properties

Set these in the Apps Script editor → **Project Settings → Script
properties**. The legacy hard-coded values still act as fallback, but
configuring properties is the supported path for prod/staging
separation.

| Key | Type | Example | Used by |
|---|---|---|---|
| `AUDIT_SHEET_ID` | string | `1Ki9...qZI` | Helper.gs / Setup |
| `ADMIN_EMAILS` | csv | `alice@rappi.com,bob@rappi.com` | Config.gs |
| `LEGAL_AUDIT_EMAILS` | csv | `legal@rappi.com` | Config.gs |
| `TEMPLATES_FOLDER_ID` | string | Drive folder ID | Admin.gs / Setup |
| `GEMINI_API_KEY` | string | `AIza...` | Admin.gs (wizard) |
| `TRACKER_SHEET_ID` | string | `19eR...6ms` | Codigo.gs (Tracker) |

### Bootstrap a fresh deployment

1. Push the project to Apps Script: `clasp push`.
2. Set the Script Properties above.
3. From the Apps Script editor run, in this order:
   - `setupAdminSystem()` — provisions `Admin_Team`, `Approval_Log`,
     extra columns on the Registry.
   - `setupTemplateEngine()` — seeds `Cashback` and `Concurso`
     baseline templates + `Template_Registry` + `Template_Fields`.
   - `setupTemplateFolders()` — creates Drive folder tree per country.
   - `setupCampaignTypes()`, `setupCountrySettings()` — seed catalogs.
4. Deploy as **Web App** (Execute as: Me, Access: domain).
5. Open the URL and sign in with a `@rappi.com` account.

### Migrations

Setupg.gs ships idempotent V3.3 migrations:

* `migrateV33()` runs 6 steps (corrupt field cleanup, ghost type
  deactivation, canonical IDs, country settings upgrade, junk sheet
  removal, legacy archive).
* `verifyV33()` validates the post-migration state.

Each step logs its progress and is safe to re-run.

---

## Legal Team Tracker (`Codigo.gs`)

Independent dashboard for weekly task tracking with per-country
leaders, SLA computation and Slack integration.

### Spreadsheet tabs (separate spreadsheet)

| Sheet | Description |
|---|---|
| `Tracking Activo` | Active tasks (16 columns; ID + metadata + project) |
| `Historial` | Completed tasks (same shape) |
| `Proyectos` | Projects (15 columns) |
| `Equipos` | Country teams, leaders, members CSV |
| `Config` | Key/value config rows starting at row 3 |

### Public RPCs

* `doGet(e)` — renders the `Dashboard` HTML template.
* `getTrackerData()` — KPI bundle (tasks, projects, team, SLA).
* `addTask(taskObj)`, `updateTaskField(id, field, value)` — task CRUD.
* `addProject(obj)`, `updateProjectField(id, field, value)` — project CRUD.
* `handleCloseTask(params)`, `handleBlockTask(params)` — Slack hooks
  (use ContentService JSON envelope).

---

## Security model

Recent hardening pass (May 2026):

* **Centralised escaping** — `Security.gs` (backend) and inline
  `escapeHtml` / `escapeAttr` / `escapeJsString` (front-end) on every
  dynamic `innerHTML` write.
* **localStorage hardening** — every `JSON.parse` of stored state
  flows through `safeJsonParse`, plus per-field validation
  (`safeUrl`, length caps).
* **URL whitelisting** — history items reject anything that is not
  `http(s)://`; doc-ID interpolation goes through `encodeURIComponent`.
* **Formula-injection guard** — `sheetCellSafe()` prefixes user input
  starting with `=`, `+`, `-`, `@` with `'` before persisting.
* **Role enforcement** — every admin RPC runs through `_requireRole()`
  and returns a sanitised JSON envelope; errors no longer leak stack.
* **Gemini hygiene** — raw API responses are no longer echoed into
  execution logs (potential PII leak through prompt echo).
* **Property-based secrets** — `AUDIT_SHEET_ID`, `ADMIN_EMAILS`,
  `GEMINI_API_KEY`, `TEMPLATES_FOLDER_ID` resolve via Script
  Properties with safe legacy fallback.

### Reporting a security issue

Email the project owner directly (see `ADMIN_EMAILS`). Do **not** open
a public issue.

---

## Local development

This is a pure Google Apps Script project; there is no Node toolchain.

```bash
# 1. Install clasp (one-time)
npm install -g @google/clasp
clasp login

# 2. Pull from Apps Script
clasp pull

# 3. Push local changes
clasp push
```

### Conventions

* **No build step.** Files are uploaded as-is; keep ES5-compatible
  syntax in `.gs` (no `optional?.chaining`, no template literal types).
  Modern arrow functions and `let`/`const` are fine — V8 runtime.
* **No tests yet.** Use `testData()` and the migration verifiers as
  smoke checks from the Apps Script editor.
* **Naming.** Private helpers start with `_`. Files match the Apps
  Script editor display order.

---

## Roadmap

* Re-include `processWebPayload`/`askGemini` in the repo (currently
  source-of-truth lives only in the Apps Script editor).
* Migrate the front-end to the dark-theme proposal
  (`Propuesta de ajuste del front`) once feature parity is reached.
* Add a CI lint step (`eslint --env googleappsscript`) once clasp + GH
  Actions are wired up.
* Split `WebApp.Html` into Apps Script HTML includes (`<?!= include('foo') ?>`).

---

## License

Internal Rappi tooling — not licensed for external use.
