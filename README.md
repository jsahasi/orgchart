# Org Chart Generator

Generates interactive, standalone HTML org charts from Excel workbooks. Produces two files:

1. **`org_drilldown.html`** — Full org chart with names
2. **`org_drilldown_redacted.html`** — Same chart with all names anonymized

## Quick Start

```bash
python generate_org_html.py
```

Opens `org_drilldown.html` in any browser. No server needed — fully self-contained.

## Requirements

- Python 3.8+
- openpyxl (auto-installed on first run)

## Input Data

Place Excel files in the `data/` folder:

| File | Purpose |
|------|---------|
| `JayeshSahasi_ON24 QA-Dev Org List.xlsx` | Org roster with hierarchy, teams, employment |
| `JayeshSahasi_EngProduct_Jayesh_Talent_Snapshot_*.xlsx` | Job titles enrichment |
| `JayeshSahasi_SCRUMS.xlsx` | Canonical scrum team names |

## Features

- **Home view**: Landing page showing Jayesh + all 5 direct reports across orgs
- **Hierarchical drilldown**: Click any person to see their direct reports
- **Org switching**: Dropdown to switch between Product-Design, QA, Dev, TPM
- **FTE-only toggle**: Hide contractors across all views
- **Scrum team view**: Click team pills to see team composition with lead identification
- **Search**: Find people by name across all orgs
- **Breadcrumb navigation**: Always know where you are, navigate up easily
- **Dotted-line indicators**: Shows secondary reporting relationships (e.g., Kamal → Jayesh)

## Org Structure

```
Jayesh Sahasi (EVP Product & CTO)
├── Steve Sims (Product-Design)
│   ├── Kevin Miller, Jared Chappin, Salma Bargach, Felix Biju
│   └── Trupti Telang (UX)
├── Oleg Massakovskyy (QA)
│   ├── Ashish Oza (Automation)
│   ├── Rumana Ilyas
│   ├── Jenny Wai Li
│   └── Shefali Singh
├── Jaimini Joshi (Dev)
│   ├── Kamal (dotted-line to Jayesh), Alberto, Shishir, Angel, ...
│   └── (Full Dev Org)
├── Mahesh Kheny (Salesforce)
└── Jagjit Singh (TPM)
```

## Scrum Teams (20)

Analytics, Appgen, Automation, Cloud Engineering, Console, EER, EHub/Target, Elite Admin, Elite Studio, Eng Tools, Engineering AI, Engineering Support, Forums, Go Live, Integrations, Presenter, Salesforce, TPM, VC, Vids

## Redacted Version

The redacted HTML replaces every person's name with `Person NNN`. No real names appear anywhere in the HTML source — verified by automated scan. Structure, titles, team names, and employment status are preserved.
