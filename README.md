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
| `JayeshSahasi_EngProduct_Jayesh_Talent_Snapshot_*.xlsx` | Job titles, Talent Band, Category, Rationale |
| `JayeshSahasi_SCRUMS.xlsx` | Canonical scrum team names |

## Features

- **Home view**: Landing page showing Jayesh + all 5 direct reports across 5 orgs
- **Hierarchical drilldown**: Click any person to see their direct reports
- **Org switching**: Dropdown for Product-Design, QA, Dev, Salesforce, TPM
- **FTE/Contractor toggles**: Show/hide employees or contractors across all views
- **Scrum team view**: Click team pills to see team composition with lead identification
- **List view**: Flat sortable table of all people — sort by name, title, type, manager, or org. Clickable scrum team pills.
- **Search**: Find people by name across all orgs
- **Breadcrumb navigation**: Always know where you are, navigate up easily
- **Dotted-line indicators**: Shows secondary reporting (e.g., Kamalaksha Ghosh → Jayesh)
- **Talent info tooltip**: Hover "i" icon next to names to see Talent Band, Category, and Rationale (from talent snapshot)
- **Responsive layout**: Works on desktop, tablet, and mobile

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
│   ├── Kamalaksha Ghosh (dotted-line to Jayesh), Alberto, Shishir, Angel, ...
│   └── (Full Dev Org)
├── Mahesh Kheny (Salesforce)
│   └── Homer Santos → James Gomez, Raj Kommera
└── Jagjit Singh (TPM)
    └── Kashan Babar, Ishwinder Walia, C-Bhagyashree More
```

## Scrum Teams (20)

Analytics, Appgen, Automation, Cloud Engineering, Console, EER, EHub/Target, Elite Admin, Elite Studio, Eng Tools, Engineering AI, Engineering Support, Forums, Go Live, Integrations, Presenter, Salesforce, TPM, VC, Vids

## Redacted Version

The redacted HTML replaces every person's name with a blacked-out format: initials are visible, remaining letters are replaced with block characters (e.g., "J████ S█████"). Node IDs are anonymized to `node-NNN`. No real names appear anywhere in the HTML source — not in data, IDs, variables, or comments. Rationale text has real names scrubbed. Not discoverable via view-source. Verified by automated scan on every generation. Structure, titles, talent data, team names, and employment status are preserved.
