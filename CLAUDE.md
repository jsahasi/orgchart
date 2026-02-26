# Org Chart Project

## Overview
Interactive org chart generator for Jayesh Sahasi's multi-org structure (Product-Design, QA, Dev, TPM). Reads Excel workbooks and produces two standalone HTML files with drilldown navigation.

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

## Business Rules

### Rule 1: Home View
Home page shows Jayesh Sahasi (EVP Product & CTO) with all 5 direct reports from across all orgs, each linking to their respective org drilldown.

### Rule 2: Jayesh DRs (Cross-Org)
- Steve Sims → Product-Design
- Oleg Massakovskyy → Full QA Org
- Jaimini Joshi → Full Dev Org
- Mahesh Kheny → Full Dev Org (Salesforce)
- Jagjit Singh → TPM

### Rule 3: Kamal Reports to Jaimini
Kamal appears as "Reports To" only (no Name row). Routes to Jaimini with dotted-line annotation to Jayesh.

### Rule 4: Hide Unassigned Bucket
No "Unassigned" placeholder in any org view.

### Rule 5: QA Hierarchy
Automation contractors report to Ashish Oza → Oleg. Oleg's real DRs: Rumana, Shefali, Jenny, Ashish.

### Rule 6: Canonical Scrum Team Names (20 teams)
Source: `JayeshSahasi_SCRUMS.xlsx` "Team Reorg" tab. Aliases mapped via `TEAM_ALIASES`:
- Integration → Integrations
- GoLive → Go Live
- Video → Vids
- Segmentation → dropped entirely (not staffed)
- Console, Forums, Vids listed as separate teams

### Rule 7: Default Titles by Org
| Org | Default Title |
|-----|---------------|
| Full Dev Org | Senior Software Engineer |
| Full QA Org | Sr. QA Engineer |
| Product-Design | Senior UX Designer |
| TPM | Scrum Master |

Per-person overrides: Sanel Selimovic → Senior UX Designer, Jagjit Singh → Director, Program Management

### Rule 8: Changelog Row Filtering
Skip rows at bottom of sheets that are audit/changelog entries (dates, "Updated", "Notes", etc.) via `CHANGELOG_SKIP_NAMES`.

## Data Parsing Conventions
- **Forward-fill**: "Reports To" blank cells inherit from row above
- **Team columns**: `Team` (Product-Design, Dev, TPM), `Teams` + `Teams.1` (QA) — merge both for QA
- **Compound teams**: Raw values like `P10Console Forums- Vids-Vibbio-` split via `COMPOUND_TEAM_MAP`
- **Name matching**: Fuzzy via `NICKNAME_MAP` (Steve→Stephen, Dan→Daniel), partial first-name match, last-name match

## Redacted Version
- All names replaced with "Person NNN" (sequential, consistent)
- `dottedLine` field also redacted
- No real names anywhere in HTML source (verified by automated scan)
- Titles, teams, hierarchy preserved

## HTML Template
Uses raw string template `r'''...'''` with `__TITLE_SUFFIX__` and `__DATA_JSON__` placeholders (not f-strings — avoids JS quote escaping issues).

## Key Config Constants in generate_org_html.py
- `JAYESH_DRS` — list of DR name hints
- `NICKNAME_MAP` — first-name aliases
- `MANUAL_TITLE_OVERRIDES` / `MANUAL_TITLE_OVERRIDES_EXTRA` — per-person title fixes
- `TEAM_ALIASES` — raw team → canonical name mapping (~70 entries)
- `COMPOUND_TEAM_MAP` — multi-team raw values
- `DEFAULT_TITLES_BY_ORG` — fallback titles per org
- `QA_OLEG_REAL_DRS` / `QA_AUTOMATION_MANAGER` — QA hierarchy fix
- `CHANGELOG_SKIP_NAMES` — rows to filter out
