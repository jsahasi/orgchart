# Org Chart Project

## Overview
Interactive org chart generator for Jayesh Sahasi's multi-org structure. Reads Excel workbooks and produces two standalone HTML files with drilldown navigation.

## Project Structure
```
orgchart/
  data/
    JayeshSahasi_ON24 QA-Dev Org List.xlsx    # Org roster (4 tabs)
    JayeshSahasi_EngProduct_Jayesh_Talent_Snapshot_*.xlsx  # Titles
    JayeshSahasi_SCRUMS.xlsx                  # Canonical scrum team names
  generate_org_html.py    # Generator script (Python)
  org_drilldown.html      # Output: named version
  org_drilldown_redacted.html  # Output: redacted version
```

## Running
```bash
python generate_org_html.py
```
Requires: Python 3.8+, openpyxl (auto-installed if missing)

## 5 Orgs (from dropdown)
| Org | Jayesh DR | Source Tab |
|-----|-----------|-----------|
| Product-Design | Steve Sims | Product-Design |
| Full QA Org | Oleg Massakovskyy | Full QA Org |
| Full Dev Org | Jaimini Joshi | Full Dev Org |
| Salesforce | Mahesh Kheny | Full Dev Org (split out at runtime) |
| TPM | Jagjit Singh | TPM |

Salesforce is synthesized by extracting Mahesh's subtree from Full Dev Org via `split_salesforce_org()`.

## Business Rules

### Home View
Home page shows Jayesh Sahasi (EVP Product & CTO) with all 5 direct reports across 5 orgs.

### Kamalaksha Ghosh
Full name: Kamalaksha Ghosh, VP Engineering. Appears as "Kamal" in "Reports To" column only (no Name row). Reports to Jaimini with dotted-line to Jayesh.

### QA Hierarchy
Automation contractors report to Ashish Oza (contractor) → Oleg. Oleg's real DRs: Rumana, Shefali, Jenny, Ashish.

### Canonical Scrum Team Names (20 teams)
Source: `JayeshSahasi_SCRUMS.xlsx` "Team Reorg" tab. Aliases mapped via `TEAM_ALIASES`:
- Integration → Integrations, GoLive → Go Live, Video → Vids
- Segmentation → dropped entirely (not staffed)
- Console, Forums, Vids listed as separate teams

### Default Titles by Org
| Org | Default Title |
|-----|---------------|
| Full Dev Org | Senior Software Engineer |
| Full QA Org | Sr. QA Engineer |
| Product-Design | Senior UX Designer |
| TPM | Scrum Master |

Per-person overrides: Sanel Selimovic → Senior UX Designer, Jagjit Singh → Director, Program Management, Kamalaksha Ghosh → VP Engineering

### Placeholder Nodes
People who appear only in "Reports To" (not as Name rows) become placeholder nodes. All default to employment = "Full Time".

### Changelog Row Filtering
Skip rows at bottom of sheets that are audit/changelog entries (dates, "Updated", "Notes", etc.) via `CHANGELOG_SKIP_NAMES`.

## Data Parsing Conventions
- **Forward-fill**: "Reports To" blank cells inherit from row above
- **Team columns**: `Team` (Product-Design, Dev, TPM), `Teams` + `Teams.1` (QA) — merge both for QA
- **Compound teams**: Raw values like `P10Console Forums- Vids-Vibbio-` split via `COMPOUND_TEAM_MAP`
- **Name matching**: Fuzzy via `NICKNAME_MAP` (Steve→Stephen, Dan→Daniel), partial first-name match, last-name match

## Redacted Version
- All names → "Person NNN" (sequential, consistent)
- All node IDs → "node-NNN" (no name fragments in keys)
- `dottedLine` field also redacted
- No real names anywhere in HTML source — data, IDs, JS variables, comments all clean
- Verified by automated scan on every generation

## HTML Template
Uses raw string template `r'''...'''` with `__TITLE_SUFFIX__` and `__DATA_JSON__` placeholders (not f-strings — avoids JS quote escaping issues). Includes responsive CSS (`@media` breakpoints at 768px and 480px).

## Key Config Constants
- `JAYESH_DRS` — list of DR name hints
- `SALESFORCE_ORG_NAME` / `SALESFORCE_DR_HINT` — Salesforce org split config
- `KAMAL_FULL_NAME` / `KAMAL_TITLE` — Kamalaksha Ghosh identity
- `NICKNAME_MAP` — first-name aliases
- `MANUAL_TITLE_OVERRIDES` / `MANUAL_TITLE_OVERRIDES_EXTRA` — per-person title fixes
- `TEAM_ALIASES` — raw team → canonical name mapping (~70 entries)
- `COMPOUND_TEAM_MAP` — multi-team raw values
- `DEFAULT_TITLES_BY_ORG` — fallback titles per org
- `QA_OLEG_REAL_DRS` / `QA_AUTOMATION_MANAGER` — QA hierarchy fix
- `CHANGELOG_SKIP_NAMES` — rows to filter out
