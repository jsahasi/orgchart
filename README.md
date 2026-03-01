# Org Chart

Interactive org chart viewer for a multi-org engineering structure. Reads a master Excel workbook and produces standalone HTML files with drilldown navigation, scrum team views, and talent data.

Deployed on [Streamlit Community Cloud](https://share.streamlit.io) as a password-protected web app.

## Quick Start

### Generate HTML locally
```bash
pip install openpyxl
python generate_org_html.py
```
Produces `org_drilldown.html` (named) and `org_drilldown_redacted.html` (redacted). Open either in any browser — fully self-contained.

### Run the Streamlit app locally
```bash
pip install -r requirements.txt
streamlit run streamlit_app.py
```
Password: configured in `.streamlit/secrets.toml`

## Deployment (Streamlit Community Cloud)

1. Push this repo to GitHub (private recommended)
2. Go to [share.streamlit.io](https://share.streamlit.io) and connect the repo
3. Set **Main file path** to `streamlit_app.py`
4. In **Advanced settings > Secrets**, add:
   ```toml
   app_password = "your-password"
   github_token = "ghp_your_token"
   github_repo = "owner/orgchart"
   ```
5. Monthly auto-regeneration runs via GitHub Actions (1st of each month)

## Features

- **Org drilldown** — hierarchical card view with drill-in navigation across 5 orgs
- **Scrum team view** — team composition grouped by discipline with lead identification
- **List view** — flat sortable table with clickable navigation
- **FTE/Contractor toggles** — show/hide employees or contractors across all views
- **Search** — find people by name across all orgs
- **Talent tooltips** — band, category, and rationale on hover
- **Redacted version** — all names replaced with blacked-out initials, verified clean
- **Admin panel** — upload new Excel to regenerate and auto-commit to GitHub
- **Monthly auto-regen** — GitHub Actions workflow runs on the 1st of each month
- **Responsive layout** — works on desktop, tablet, and mobile

## Input Data

Single source of truth: `data/orgchart_master_data.xlsx` with 3 sheets:

| Sheet | Purpose |
|-------|---------|
| People | Names, titles, employment, org, reports-to, scrum teams, talent data |
| Scrum Teams | Team membership with discipline and lead flags |
| Teams Hierarchy | Team to Dev Lead, QA Lead, Director mappings |

Legacy source files preserved in `data/legacy/` for backward compatibility.

## Project Structure

| File | Purpose |
|------|---------|
| `streamlit_app.py` | Streamlit web app (viewer + admin) |
| `generate_org_html.py` | Generator: master Excel to HTML |
| `generate_org_html_legacy.py` | Legacy generator (4 source files) |
| `org_html_shared.py` | Shared utilities + HTML template |
| `requirements.txt` | Python dependencies |
| `.streamlit/config.toml` | Streamlit theme + server config |
| `.github/workflows/regenerate.yml` | Monthly auto-regeneration |

## Redacted Version

The redacted HTML replaces every name with a blacked-out format: initials visible, remaining letters replaced with block characters (e.g., "J████ S█████"). Node IDs anonymized. No real names anywhere in the HTML source. Verified by automated scan on every generation.
