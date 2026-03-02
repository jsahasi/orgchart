# Org Chart

Interactive org chart viewer for a multi-org engineering structure. Reads a master Excel workbook and produces standalone HTML files with drilldown navigation, scrum team views, and talent data.

Deployed on [Streamlit Community Cloud](https://share.streamlit.io) as a password-protected web app.

## Architecture

- **Code repo** (`jsahasi/orgchart`) — public, contains only code (no data)
- **Data repo** (`jsahasi/orgchart-data`) — private, contains Excel + generated HTML files

The Streamlit app fetches data files at runtime from the private repo via GitHub API, keeping employee data out of the public repo.

## Quick Start

### Generate HTML locally
```bash
pip install openpyxl
python generate_org_html.py
```
Requires `data/orgchart_master_data.xlsx` in the working directory. Produces `org_drilldown.html` (named) and `org_drilldown_redacted.html` (redacted).

### Run the Streamlit app locally
```bash
pip install -r requirements.txt
streamlit run streamlit_app.py
```
Configure `.streamlit/secrets.toml` with passwords and GitHub token.

## Deployment (Streamlit Community Cloud)

1. Connect this repo at [share.streamlit.io](https://share.streamlit.io)
2. Set **Main file path** to `streamlit_app.py`
3. In **Advanced settings > Secrets**, add:
   ```toml
   app_password = "your-password"
   github_token = "ghp_your_token"
   github_repo = "jsahasi/orgchart"
   data_repo = "jsahasi/orgchart-data"
   ```
4. The `github_token` needs `contents:write` on both repos
5. Monthly auto-regeneration runs via GitHub Actions (1st of each month)

## Features

- **Org drilldown** — hierarchical card view with drill-in navigation across 5 orgs
- **Scrum team view** — team composition grouped by discipline with lead identification
- **List view** — flat sortable table with clickable navigation
- **FTE/Contractor toggles** — show/hide employees or contractors across all views
- **Search** — find people by name across all orgs
- **Talent tooltips** — band, category, and rationale on hover
- **Redacted version** — all names replaced with blacked-out initials, verified clean
- **Admin panel** — upload new Excel to regenerate and auto-commit to data repo
- **Monthly auto-regen** — GitHub Actions workflow runs on the 1st of each month
- **Responsive layout** — works on desktop, tablet, and mobile

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
