# Org Chart Project

## Overview
Interactive org chart generator for Jayesh Sahasi's multi-org structure. Reads Excel workbooks and produces two standalone HTML files with drilldown navigation.

## Project Structure

### Code repo (`jsahasi/orgchart`) — public
```
orgchart/
  generate_org_html.py        # Default generator (reads master Excel)
  generate_org_html_legacy.py # Legacy generator (reads 4 original source files)
  org_html_shared.py          # Shared functions (redaction, HTML template, utilities)
  streamlit_app.py            # Streamlit web app (viewer + admin)
  requirements.txt            # Python dependencies for Streamlit Cloud
  .streamlit/config.toml      # Streamlit theme + server config
  .streamlit/secrets.toml     # Local secrets (gitignored)
  .github/workflows/regenerate.yml  # Monthly auto-regeneration
```

### Data repo (`jsahasi/orgchart-data`) — private
```
orgchart-data/
  data/
    orgchart_master_data.xlsx  # Master data (single source of truth)
  org_drilldown.html           # Output: named version
  org_drilldown_redacted.html  # Output: redacted version
```

Data files are kept in a separate private repo to prevent exposure of employee data. The Streamlit app fetches them at runtime via GitHub API using `github_token`.

## Running
```bash
# Default: generate from master Excel
python generate_org_html.py

# Legacy: generate from original 4 source files (backward compat)
python generate_org_html_legacy.py
```
Requires: Python 3.8+, openpyxl (auto-installed if missing)

```bash
# Streamlit web app (local)
pip install -r requirements.txt
streamlit run streamlit_app.py
```

## Data Sources

### Master Excel (default)
| Source | What it provides |
|--------|-----------------|
| `data/orgchart_master_data.xlsx` (People sheet) | All people with names, titles, employment, org, reports-to, scrum teams, talent data |
| `data/orgchart_master_data.xlsx` (Scrum Teams sheet) | Team membership with discipline, lead flags, Product Owner, Scrum Master |
| `data/orgchart_master_data.xlsx` (Teams Hierarchy sheet) | Team → Dev Lead, QA Lead, Director mappings |

### Legacy Sources (in `data/legacy/`)

| Source | What it provides | Format |
|--------|-----------------|--------|
| `on24.xlsx` (on24 sheet) | **Definitive hierarchy** — 410 rows, full company | "Last, First" names |
| `JayeshSahasi_QA-Dev Org List.xlsx` (per-org tabs) | Employment status, scrum team assignments, contractor roster | "First Last" names |
| `JayeshSahasi_SCRUMS.xlsx` (Teams Hierachy sheet) | Team → Dev Lead, QA Lead, Director mappings | First names / short names |
| `JayeshSahasi_EngProduct_Jayesh_Talent_Snapshot_*.xlsx` | Titles + Talent Band/Category/Rationale (75 people) | "First Last" |

## 5 Orgs (from dropdown)
| Org | Jayesh DR | Source |
|-----|-----------|--------|
| Product-Design | Stephen Sims | on24 subtree |
| Full QA Org | Oleg Massakovskyy | on24 subtree |
| Full Dev Org | Jaimini Joshi + Kamal Ghosh | on24 subtree |
| Salesforce | Mahesh Kheny | on24 subtree |
| TPM | Jagjit Bhullar | on24 subtree |

Orgs assigned by BFS from each DR's subtree in on24 hierarchy via `DR_ORG_MAP`.

## Business Rules

### Home View
Home page shows Jayesh Sahasi (Executive VP, Products and CTO) with all **6 direct reports** across 5 orgs.

### Kamal Ghosh
Kamal Ghosh, Vice President Engineering. **Reports directly to Jayesh** (per on24.xlsx). Heads part of Full Dev Org alongside Jaimini.

### QA Hierarchy
Automation contractors report to Ashish Oza (contractor) → Oleg. Oleg's real DRs: Rumana, Shefali, Jenny, Ashish.

### Canonical Scrum Team Names (21 teams)
Source: `JayeshSahasi_SCRUMS.xlsx` "Team Reorg" tab. Aliases mapped via `TEAM_ALIASES`:
- Integration → Integrations, GoLive → Go Live, Video → Vids
- Segmentation → dropped entirely (not staffed)
- Console, Forums, Vids listed as separate teams

### Teams Hierarchy
`JayeshSahasi_SCRUMS.xlsx` "Teams Hierachy" sheet maps each team to Dev Lead, QA Lead, Director. Used to:
- Place contractors under the correct lead
- Identify leads in scrum team views (`isLead` flag)

### Scrum Master & Product Owner
Each scrum team has optional SM and PO fields stored in the "Scrum Teams" sheet (`Scrum Master`, `Product Owner` columns). These are team-level metadata (same value per team, not per member). Stored in `DATA.scrumMeta` as a separate top-level key:
```json
{"scrumMeta": {"Analytics": {"scrumMaster": "Kashan Babar", "productOwner": "Jared Chappin"}, ...}}
```
Dual assignments use slash notation (e.g., "Salma Bargach / Kevin Miller"). Redaction splits on `/` and redacts each name independently. Displayed below the team header in scrum view and compactly on all-scrum cards.

### Default Titles by Org
| Org | Default Title |
|-----|---------------|
| Full Dev Org | Senior Software Engineer |
| Full QA Org | Sr. QA Engineer |
| Product-Design | Senior UX Designer |
| TPM | Scrum Master |

Per-person overrides: Sanel Selimovic → Senior UX Designer, Jagjit Singh/Bhullar → Director, Program Management, Kamal Ghosh → Vice President Engineering

### Contractor Handling
Contractors (c-prefix names) only exist in per-org tabs, NOT in on24.xlsx. Manager resolution:
1. Match scrum team → find lead from Teams Hierarchy
2. Fallback: use Reports To from per-org tab (forward-filled)

### Changelog Row Filtering
Skip rows at bottom of sheets that are audit/changelog entries via `CHANGELOG_SKIP_NAMES`.

## Data Parsing Conventions
- **on24 names**: "Last, First" → converted to "First Last" via `convert_last_first()`
- **Forward-fill**: "Reports To" blank cells inherit from row above (per-org tabs only)
- **Team columns**: `Team` (Product-Design, Dev, TPM), `Teams` + `Teams.1` (QA) — merge both for QA
- **Compound teams**: Raw values like `P10Console Forums- Vids-Vibbio-` split via `COMPOUND_TEAM_MAP`
- **Name matching**: Fuzzy via `NICKNAME_MAP` (Steve→Stephen, Dan→Daniel), partial first-name match, last-name match

## Views
- **Home view**: Jayesh + 6 DRs across 5 orgs, deduplicated headcount
- **Org drilldown**: Hierarchical card view with drill-in navigation
- **Scrum team view**: Team composition grouped by discipline, with lead identification, Product Owner, and Scrum Master
- **List view**: Flat sortable table of all people (Name, Title, Type, Manager, Org, Scrum Teams). Clickable names navigate to org cards; clickable scrum team pills navigate to scrum view. Headcount deduplicated across orgs in Home list view.
- **Talent info tooltip**: "i" icon next to names of people who have talent snapshot data. Hover shows Talent Band, Talent Category, and Rationale. Present in all views (home, org drilldown, scrum, list). Rationale text in redacted version has names scrubbed.

## Redacted Version
- All names → "J████ S█████" format (initial visible, remaining letters blacked out with U+2588)
- All node IDs → "node-NNN" (no name fragments in keys)
- No real names anywhere in HTML source — data, IDs, JS variables, comments all clean
- Not discoverable via view-source; real names never embedded in output
- In-browser Redact toggle (named file only) uses same blacked-out format via `displayName()`
- Rationale text has real names scrubbed (longest-first, case-insensitive replacement)
- Verified by automated scan on every generation

## HTML Template
Uses raw string template `r'''...'''` with `__TITLE_SUFFIX__` and `__DATA_JSON__` placeholders (not f-strings — avoids JS quote escaping issues). Includes responsive CSS (`@media` breakpoints at 768px and 480px).

## Shared Module (`org_html_shared.py`)
Extracted functions used by both generators:
- `normalize_name()`, `slugify()`, `is_contractor()`, `title_seniority_score()`
- `redact_data()`, `verify_redaction()`, `generate_html()`
- `_HTML_TEMPLATE` — the complete HTML/CSS/JS template

## Streamlit App (`streamlit_app.py`)
Password-protected web viewer deployed on Streamlit Community Cloud.

### Authentication
- Single password gate via `st.secrets["app_password"]`
- All views require login; session-state based

### Views (sidebar radio)
- **Org Chart (Named)** — embeds `org_drilldown.html` via `st.components.v1.html()` + download button
- **Org Chart (Redacted)** — same for redacted version
- **Admin** — download current Excel, upload new Excel to regenerate + commit to GitHub

### Admin Upload Flow
1. Upload new `.xlsx` → runs generator pipeline in-memory
2. Commits 3 files (Excel + 2 HTMLs) to private data repo (`jsahasi/orgchart-data`) via Contents API

### Secrets (`.streamlit/secrets.toml` locally, Streamlit Cloud dashboard for prod)
- `app_password` — viewer login password
- `github_token` — GitHub PAT with `contents:write` scope on both repos
- `github_repo` — code repo (`jsahasi/orgchart`)
- `data_repo` — private data repo (`jsahasi/orgchart-data`)

### Monthly Regeneration (`.github/workflows/regenerate.yml`)
- Cron: 1st of each month at midnight UTC
- Also supports manual `workflow_dispatch`
- Checks out both code repo and data repo
- Runs `python generate_org_html.py`, commits + pushes HTML to data repo if changed
- Requires `DATA_REPO_TOKEN` repository secret (PAT with access to `orgchart-data`)

## Key Config Constants (legacy generator)
- `ON24_FILE` / `ORG_FILE` / `SCRUMS_FILE` — legacy data file paths
- `JAYESH_DRS` — list of 6 DR name hints
- `DR_ORG_MAP` — maps DR names to org view names
- `KAMAL_FULL_NAME` / `KAMAL_TITLE` — Kamal Ghosh identity
- `NICKNAME_MAP` — first-name aliases
- `MANUAL_TITLE_OVERRIDES` / `MANUAL_TITLE_OVERRIDES_EXTRA` — per-person title fixes
- `TEAM_ALIASES` — raw team → canonical name mapping (~70 entries)
- `COMPOUND_TEAM_MAP` — multi-team raw values
- `DEFAULT_TITLES_BY_ORG` — fallback titles per org
- `QA_OLEG_REAL_DRS` / `QA_AUTOMATION_MANAGER` — QA hierarchy fix
- `CHANGELOG_SKIP_NAMES` — rows to filter out
