# Org Chart Project

## Overview
Interactive org chart generator for Jayesh Sahasi's multi-org structure. Reads Excel workbooks and produces two standalone HTML files with drilldown navigation.

## Project Structure
```
orgchart/
  data/
    on24.xlsx                                  # Definitive hierarchy (410 rows, full ON24)
    JayeshSahasi_QA-Dev Org List.xlsx          # Per-org tabs (employment + teams)
    JayeshSahasi_EngProduct_Jayesh_Talent_Snapshot_*.xlsx  # Titles
    JayeshSahasi_SCRUMS.xlsx                   # Scrum teams + Teams Hierarchy
  generate_org_html.py    # Generator script (Python)
  org_drilldown.html      # Output: named version
  org_drilldown_redacted.html  # Output: redacted version
```

## Running
```bash
python generate_org_html.py
```
Requires: Python 3.8+, openpyxl (auto-installed if missing)

## Data Sources (priority order)

| Source | What it provides | Format |
|--------|-----------------|--------|
| `data/on24.xlsx` (on24 sheet) | **Definitive hierarchy** — 410 rows, full company | "Last, First" names |
| `data/JayeshSahasi_QA-Dev Org List.xlsx` (per-org tabs) | Employment status, scrum team assignments, contractor roster | "First Last" names |
| `data/JayeshSahasi_SCRUMS.xlsx` (Teams Hierachy sheet) | Team → Dev Lead, QA Lead, Director mappings | First names / short names |
| `data/JayeshSahasi_EngProduct_Jayesh_Talent_Snapshot_*.xlsx` | Supplementary title enrichment (75 people) | "First Last" |

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

### Canonical Scrum Team Names (20 teams)
Source: `JayeshSahasi_SCRUMS.xlsx` "Team Reorg" tab. Aliases mapped via `TEAM_ALIASES`:
- Integration → Integrations, GoLive → Go Live, Video → Vids
- Segmentation → dropped entirely (not staffed)
- Console, Forums, Vids listed as separate teams

### Teams Hierarchy
`JayeshSahasi_SCRUMS.xlsx` "Teams Hierachy" sheet maps each team to Dev Lead, QA Lead, Director. Used to:
- Place contractors under the correct lead
- Identify leads in scrum team views (`isLead` flag)

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
- **Scrum team view**: Team composition grouped by discipline, with lead identification
- **List view**: Flat sortable table of all people (Name, Title, Type, Manager, Org, Scrum Teams). Clickable names navigate to org cards; clickable scrum team pills navigate to scrum view. Headcount deduplicated across orgs in Home list view.

## Redacted Version
- All names → "J████ S█████" format (initial visible, remaining letters blacked out with U+2588)
- All node IDs → "node-NNN" (no name fragments in keys)
- No real names anywhere in HTML source — data, IDs, JS variables, comments all clean
- Not discoverable via view-source; real names never embedded in output
- In-browser Redact toggle (named file only) uses same blacked-out format via `displayName()`
- Verified by automated scan on every generation

## HTML Template
Uses raw string template `r'''...'''` with `__TITLE_SUFFIX__` and `__DATA_JSON__` placeholders (not f-strings — avoids JS quote escaping issues). Includes responsive CSS (`@media` breakpoints at 768px and 480px).

## Key Config Constants
- `ON24_FILE` / `ORG_FILE` / `SCRUMS_FILE` — data file paths
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
