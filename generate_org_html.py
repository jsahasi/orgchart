#!/usr/bin/env python3
"""
Org Chart HTML Generator
Reads org roster + talent snapshot Excel files, produces two standalone HTML files:
  1. org_drilldown.html        — full names
  2. org_drilldown_redacted.html — names replaced with "Person NNN"
"""

import json
import re
import copy
import sys
from pathlib import Path
from collections import defaultdict

try:
    import openpyxl
except ImportError:
    print("Installing openpyxl...")
    import subprocess
    subprocess.check_call([sys.executable, "-m", "pip", "install", "openpyxl", "-q"])
    import openpyxl

# ─── Configuration ───────────────────────────────────────────────────────────

DATA_DIR = Path(__file__).parent / "data"
ON24_FILE = DATA_DIR / "on24.xlsx"
ORG_FILE = DATA_DIR / "JayeshSahasi_QA-Dev Org List.xlsx"
SCRUMS_FILE = DATA_DIR / "JayeshSahasi_SCRUMS.xlsx"
TALENT_FILE = DATA_DIR / "JayeshSahasi_EngProduct_Jayesh_Talent_Snapshot_Leader_Input_2026.01.26.xlsx"
OUTPUT_FILE = Path(__file__).parent / "org_drilldown.html"
REDACTED_FILE = Path(__file__).parent / "org_drilldown_redacted.html"

ORG_TABS = ["Product-Design", "Full QA Org", "Full Dev Org", "TPM", "ORG"]
SALESFORCE_ORG_NAME = "Salesforce"
TALENT_TABS = ["Dev", "QA", "Salesforce", "Product Management", "Program Management"]

JAYESH_NAME = "Jayesh Sahasi"
JAYESH_TITLE = "Executive VP, Products and CTO"
JAYESH_DRS = ["Jaimini", "Kamal", "Oleg", "Mahesh", "Steve Sims", "Jagjit"]

NICKNAME_MAP = {
    "steve": "stephen",
    "dan": "daniel",
    "jay": "jawynson",
    "mike": "michael",
    "ben": "benjamin",
    "raj": "raj",
}

# Full-name aliases: alternate full names → on24 canonical name (normalized)
NAME_ALIASES = {
    "jagjit singh": "jagjit bhullar",
    "mahesh khenny": "mahesh kheny",
    "jared chappins": "jared chappin",
    "ishwinder wallia": "ishwinder walia",
    "jose viscasillas": "jose viscasillas reyes",
    "michale lin": "michael lin",
    "rajesh kumar": "kumar rajesh",
    "bhagyashree more": "c-bhagyashree more",
}

# Kamal identity (on24.xlsx is definitive — reports to Jayesh)
KAMAL_FULL_NAME = "Kamal Ghosh"
KAMAL_TITLE = "Vice President Engineering"

# Manual title overrides for people whose names differ between files
MANUAL_TITLE_OVERRIDES = {
    "jagjit singh": "Director, Program Management",
    "jagjit bhullar": "Director, Program Management",
    "kamal ghosh": "Vice President Engineering",
}

# QA Org hierarchy fix: automation contractors report to Ashish, not directly to Oleg.
# Oleg's real direct reports are: Rumana, Shefali, Jenny, Ashish.
QA_OLEG_REAL_DRS = {"rumana", "shefali", "jenny", "ashish"}
QA_AUTOMATION_MANAGER = "ashish"  # Ashish Oza manages the automation team

# Default titles for people with no title found, by org tab
DEFAULT_TITLES_BY_ORG = {
    "Full Dev Org": "Senior Software Engineer",
    "Full QA Org": "Sr. QA Engineer",
    "Product-Design": "Senior UX Designer",
    "TPM": "Scrum Master",
}

# Per-person title overrides (normalized name -> title)
MANUAL_TITLE_OVERRIDES_EXTRA = {
    "sanel selimovic": "Senior UX Designer",
}

# DR-to-org mapping: which Jayesh DR heads which org view
DR_ORG_MAP = {
    "stephen sims": "Product-Design",
    "steve sims": "Product-Design",
    "oleg massakovskyy": "Full QA Org",
    "jaimini joshi": "Full Dev Org",
    "kamal ghosh": "Full Dev Org",
    "mahesh kheny": SALESFORCE_ORG_NAME,
    "jagjit bhullar": "TPM",
    "jagjit singh": "TPM",
}

# Changelog/notes rows that are NOT people — skip during parsing
CHANGELOG_SKIP_NAMES = {
    "issh", "update by", "notes", "created xl", "added teams", "date",
    "updated",
}

# Canonical scrum teams (from JayeshSahasi_SCRUMS.xlsx "Team Reorg" tab):
#  1. Analytics           2. Integration        3. Segmentation
#  4. Video               5. Cloud Engineering   6. Eng Tools
#  7. EHub/Target         8. VC                  9. GoLive
# 10. Engineering Support 11. Forums            12. Appgen
# 13. Console             14. Elite Studio      15. Elite Admin
# 16. Salesforce          17. Automation        18. Engineering AI
# 19. TPM
# Plus EER and Presenter (in org data but not in SCRUMS.xlsx)

TEAM_ALIASES = {
    # Console
    "p10console": "Console",
    "p10-console": "Console",
    "p10console forums": "Console",
    "p10-console forums": "Console",
    "p10console forums-": "Console",
    "console/es": "Console",
    "console": "Console",
    # Vids
    "vids": "Vids",
    "vids-vibbio": "Vids",
    "vids/vibbio": "Vids",
    "vids/forums": "Vids",
    "vids/es": "Vids",
    "video": "Vids",
    # Go Live
    "gl": "Go Live",
    "golive": "Go Live",
    "go live": "Go Live",
    "golife": "Go Live",
    # Elite Studio
    "es": "Elite Studio",
    "elite studio": "Elite Studio",
    "elite studio forums": "Elite Studio",
    "elite studio/forums": "Elite Studio",
    "elite studio/eng tools": "Elite Studio",
    "es forums": "Elite Studio",
    "es/forums": "Elite Studio",
    # Elite Admin
    "elite admin": "Elite Admin",
    "elite admin forums": "Elite Admin",
    "elite admin/elite admin forums": "Elite Admin",
    "elite admin/ai": "Elite Admin",
    "elite admin / wcc": "Elite Admin",
    # EHub/Target
    "ehub": "EHub/Target",
    "ehub/target": "EHub/Target",
    "ehub/target (orion)": "EHub/Target",
    # Integrations
    "integrations": "Integrations",
    "integration": "Integrations",
    # Engineering AI
    "ai": "Engineering AI",
    "engineering ai": "Engineering AI",
    "ai/analytics/integrations": "Engineering AI",
    "ai/integrations": "Engineering AI",
    # Cloud Engineering
    "cloud": "Cloud Engineering",
    "cloud engineering": "Cloud Engineering",
    "cloud engineering\n(ce)": "Cloud Engineering",
    # Eng Tools
    "engg tools": "Eng Tools",
    "eng tools": "Eng Tools",
    "tools": "Eng Tools",
    "eng tools /es forums": "Eng Tools",
    # Engineering Support
    "engg support": "Engineering Support",
    "eng support": "Engineering Support",
    "engineering support": "Engineering Support",
    "engineering support (cogs)": "Engineering Support",
    # Segmentation — NOT STAFFED, drop entirely
    "analytics -segmentation": None,
    "analytics-segmentation": None,
    "segmentation": None,
    # Appgen
    "appgen": "Appgen",
    # EER (not in SCRUMS.xlsx but referenced in Dev Org)
    "eer": "EER",
    "eer/gl": "EER",
    # Presenter (not in SCRUMS.xlsx but referenced in QA)
    "presenter": "Presenter",
    # Automation
    "automation": "Automation",
    # VC
    "vc": "VC",
    # Forums
    "forums": "Forums",
    "forums (kms)": "Forums",
    # Analytics
    "analytics": "Analytics",
    # Salesforce
    "salesforce": "Salesforce",
    # TPM
    "tpm": "TPM",
    # Non-scrum labels (skip these)
    "product mgmt": None,
    "ux": None,
    "ui/ux": None,
    "ui": None,
    "dir": None,
    "ft": None,
}

# Compound raw values that map to MULTIPLE teams (after % removal)
COMPOUND_TEAM_MAP = {
    "p10console forums- vids-vibbio-": ["Console", "Vids"],
    "p10console forums-  vids-vibbio-": ["Console", "Vids"],
}

# ─── Utilities ───────────────────────────────────────────────────────────────

def normalize_name(name):
    """Lowercase, strip, collapse spaces, remove periods/commas."""
    if not name or not isinstance(name, str):
        return ""
    name = name.strip()
    name = re.sub(r'[.,]', '', name)
    name = re.sub(r'\s+', ' ', name)
    return name.lower()


def slugify(name):
    """Convert name to URL-safe slug."""
    s = normalize_name(name)
    s = re.sub(r'[^a-z0-9\s-]', '', s)
    s = re.sub(r'\s+', '-', s).strip('-')
    return s or 'unknown'


def is_contractor(employment):
    """Check if employment string indicates contractor."""
    if not employment or not isinstance(employment, str):
        return False
    emp_lower = employment.lower()
    return any(kw in emp_lower for kw in ["contract", "vendor", "consult"])


def normalize_team_name(name):
    """Normalize a team name using aliases. Returns None to skip non-scrum labels."""
    key = name.strip().lower()
    key = re.sub(r'\s+', ' ', key)
    if key in TEAM_ALIASES:
        return TEAM_ALIASES[key]  # May be None for non-scrum labels
    # Keep as-is if not in aliases
    return name.strip()


def parse_scrum_teams(raw):
    """Parse team string into list of team names."""
    if not raw or not isinstance(raw, str):
        return []
    raw = raw.strip()
    if not raw or raw.lower() in ('nan', 'none', 'n/a', '-', '', 'dir'):
        return []
    # Remove percentage annotations
    raw = re.sub(r'\d+\s*%', '', raw)
    # Check compound team map first
    raw_key = raw.strip().lower()
    raw_key = re.sub(r'\s+', ' ', raw_key)
    if raw_key in COMPOUND_TEAM_MAP:
        return COMPOUND_TEAM_MAP[raw_key]
    # Split by common delimiters
    tokens = re.split(r'[,;/|&\n]+', raw)
    teams = []
    seen = set()
    for t in tokens:
        t = t.strip()
        t = re.sub(r'\s+', ' ', t)
        if not t or len(t) < 2:
            continue
        # Skip numeric-only tokens (changelog artifacts)
        if re.match(r'^\d+$', t):
            continue
        t = normalize_team_name(t)
        if t is None:
            continue  # Non-scrum label (Product Mgmt, UX, Dir, etc.)
        key = t.lower()
        if key not in seen:
            seen.add(key)
            teams.append(t)
    return teams


def title_seniority_score(title):
    """Score a title for seniority ranking."""
    if not title:
        return 0
    t = title.lower()
    scores = [
        ("evp", 100), ("svp", 95), ("sr. vp", 95), ("senior vice president", 95),
        ("vp", 90), ("vice president", 90),
        ("senior director", 82), ("sr. director", 82),
        ("director", 80),
        ("head", 78),
        ("senior manager", 70), ("sr. manager", 70),
        ("manager", 65), ("mgr", 65),
        ("lead", 60),
        ("principal", 58),
        ("staff", 55),
        ("senior", 45), ("sr.", 45), ("sr ", 45),
        ("engineer", 40),
        ("analyst", 35),
        ("associate", 30),
        ("intern", 10),
    ]
    best = 0
    for keyword, score in scores:
        if keyword in t:
            best = max(best, score)
    return best


def is_changelog_row(reports_to, name):
    """Filter out changelog/audit rows at bottom of sheets."""
    if not name or not isinstance(name, str):
        return True
    name_lower = name.strip().lower()
    rt_lower = (reports_to or "").strip().lower() if isinstance(reports_to, str) else ""

    if name_lower in CHANGELOG_SKIP_NAMES or name_lower == "":
        return True
    # Date-like Reports To
    if rt_lower and re.match(r'^(date|dec|feb|jan|mar|apr|may|jun|jul|aug|sep|oct|nov)\b', rt_lower):
        return True
    # If name is just a number or very short
    if re.match(r'^\d+$', name.strip()):
        return True
    return False


# ─── Step 0a: Parse on24.xlsx — Definitive Hierarchy ────────────────────────

def convert_last_first(name_str):
    """Convert 'Last, First' to 'First Last'. Handles edge cases."""
    if not name_str or not isinstance(name_str, str):
        return ""
    name_str = name_str.strip()
    if ',' not in name_str:
        return name_str
    parts = name_str.split(',', 1)
    last = parts[0].strip()
    first = parts[1].strip()
    if not first or not last:
        return name_str
    return f"{first} {last}"


def parse_on24(filepath):
    """Parse on24.xlsx and return Jayesh's org subtree as a dict.

    Returns: {normalized_name -> {name, reports_to, title, department, location, num_drs}}
    Only includes people in Jayesh's subtree (BFS from Jayesh down).
    """
    wb = openpyxl.load_workbook(filepath, read_only=True, data_only=True)
    ws = wb["on24"]
    rows = list(ws.iter_rows(min_row=1, values_only=True))
    wb.close()

    if not rows:
        return {}

    header = [str(c).strip() if c else "" for c in rows[0]]
    header_lower = [h.lower() for h in header]

    # Find column indices
    col_map = {}
    for i, h in enumerate(header_lower):
        if h == "name":
            col_map["name"] = i
        elif h == "reports to":
            col_map["reports_to"] = i
        elif h == "job title":
            col_map["title"] = i
        elif h == "number of direct reports":
            col_map["num_drs"] = i
        elif h == "department":
            col_map["dept"] = i
        elif h == "location":
            col_map["loc"] = i

    if "name" not in col_map or "reports_to" not in col_map:
        print("  [WARN] on24.xlsx missing Name or Reports To column")
        return {}

    # Parse all rows — build name→record and reports_to index
    all_people = {}  # normalized_name -> record
    children_of = defaultdict(list)  # normalized_manager -> [normalized_name, ...]

    for row in rows[1:]:
        def cell(key):
            idx = col_map.get(key)
            if idx is not None and len(row) > idx and row[idx]:
                return str(row[idx]).strip()
            return ""

        raw_name = cell("name")
        raw_rt = cell("reports_to")
        if not raw_name:
            continue

        name = convert_last_first(raw_name)
        reports_to = convert_last_first(raw_rt)
        norm = normalize_name(name)
        norm_rt = normalize_name(reports_to)

        num_drs_raw = cell("num_drs")
        try:
            num_drs = int(num_drs_raw)
        except (ValueError, TypeError):
            num_drs = 0

        all_people[norm] = {
            "name": name,
            "reports_to": reports_to,
            "title": cell("title"),
            "department": cell("dept"),
            "location": cell("loc"),
            "num_drs": num_drs,
        }
        if norm_rt:
            children_of[norm_rt].append(norm)

    # BFS from Jayesh to collect his subtree
    jayesh_norm = normalize_name(JAYESH_NAME)
    if jayesh_norm not in all_people:
        print(f"  [WARN] '{JAYESH_NAME}' not found in on24.xlsx")
        return {}

    subtree = {}
    queue = [jayesh_norm]
    while queue:
        current = queue.pop(0)
        if current in subtree:
            continue
        if current not in all_people:
            continue
        subtree[current] = all_people[current]
        for child_norm in children_of.get(current, []):
            if child_norm not in subtree:
                queue.append(child_norm)

    print(f"  on24.xlsx: {len(all_people)} total people, {len(subtree)} in Jayesh's org")
    return subtree


# ─── Step 0b: Parse Teams Hierarchy ────────────────────────────────────────

def parse_teams_hierarchy(filepath):
    """Parse SCRUMS.xlsx 'Teams Hierachy' sheet for team→leader mappings.

    Returns: {canonical_team_name -> {dev_leads: [names], qa_leads: [names],
              dev_director: name, qa_director: name, vp: name}}
    """
    wb = openpyxl.load_workbook(filepath, read_only=True, data_only=True)
    sheet_name = "Teams Hierachy"  # Known typo in file
    if sheet_name not in wb.sheetnames:
        print(f"  [WARN] Sheet '{sheet_name}' not found in SCRUMS.xlsx")
        wb.close()
        return {}

    ws = wb[sheet_name]
    rows = list(ws.iter_rows(min_row=1, values_only=True))
    wb.close()

    if len(rows) < 2:
        return {}

    header = [str(c).strip().lower() if c else "" for c in rows[0]]

    # Find columns
    col_team = col_vp = col_dir = col_leads = col_qa_vp = col_qa_dir = col_qa_leads = None
    for i, h in enumerate(header):
        if h == "team":
            col_team = i
        elif h == "vp":
            col_vp = i
        elif h == "director":
            col_dir = i
        elif h in ("mgrs/leads", "mgrs / leads"):
            col_leads = i
        elif h == "qa vp":
            col_qa_vp = i
        elif h == "qa director":
            col_qa_dir = i
        elif h in ("qa leads", "qa lead"):
            col_qa_leads = i

    if col_team is None:
        print("  [WARN] Teams Hierarchy sheet missing 'Team' column")
        return {}

    def split_names(cell_val):
        """Split a cell value like 'Name1 / Name2' into list of names."""
        if not cell_val or not isinstance(cell_val, str):
            return []
        # Split on / and strip
        return [n.strip() for n in re.split(r'[/]', cell_val) if n.strip()]

    def cell_str(row, idx):
        if idx is not None and len(row) > idx and row[idx]:
            return str(row[idx]).strip()
        return ""

    teams_hier = {}
    for row in rows[1:]:
        team_raw = cell_str(row, col_team)
        if not team_raw:
            continue

        # Normalize team name
        team_name = normalize_team_name(team_raw)
        if team_name is None:
            continue

        teams_hier[team_name] = {
            "vp": cell_str(row, col_vp),
            "dev_director": cell_str(row, col_dir),
            "dev_leads": split_names(cell_str(row, col_leads)),
            "qa_director": cell_str(row, col_qa_dir),
            "qa_leads": split_names(cell_str(row, col_qa_leads)),
        }

    print(f"  Teams hierarchy: {len(teams_hier)} teams loaded")
    return teams_hier


# ─── Step 1: Parse Talent Snapshot ───────────────────────────────────────────

def parse_talent_snapshot(filepath):
    """Build title_map: normalize_name -> title."""
    wb = openpyxl.load_workbook(filepath, read_only=True, data_only=True)
    title_map = {}

    for tab_name in TALENT_TABS:
        if tab_name not in wb.sheetnames:
            print(f"  [WARN] Talent tab '{tab_name}' not found, skipping")
            continue
        ws = wb[tab_name]
        rows = list(ws.iter_rows(min_row=1, values_only=True))
        if not rows:
            continue

        # Find header row
        header = [str(c).strip() if c else "" for c in rows[0]]

        # Find column indices
        first_name_idx = None
        last_name_idx = None
        title_idx = None
        for i, h in enumerate(header):
            hl = h.lower()
            if hl == "first name":
                first_name_idx = i
            elif hl == "last name":
                last_name_idx = i
            elif hl == "title":
                title_idx = i

        if first_name_idx is None or last_name_idx is None or title_idx is None:
            print(f"  [WARN] Talent tab '{tab_name}' missing required columns, skipping")
            continue

        for row in rows[1:]:
            if len(row) <= max(first_name_idx, last_name_idx, title_idx):
                continue
            first = str(row[first_name_idx]).strip() if row[first_name_idx] else ""
            last = str(row[last_name_idx]).strip() if row[last_name_idx] else ""
            title = str(row[title_idx]).strip() if row[title_idx] else ""

            if not first and not last:
                continue

            full = f"{first} {last}".strip()
            key = normalize_name(full)
            if key and title:
                # Clean up title
                title = re.sub(r'\s+', ' ', title).strip()
                title_map[key] = title

    wb.close()
    print(f"  Talent snapshot: {len(title_map)} titles loaded")
    return title_map


# ─── Step 2: Parse Org Roster ────────────────────────────────────────────────

def parse_org_tab(ws, tab_name):
    """Parse a single org tab for employment status and scrum teams only.

    Returns: dict of normalized_name -> {name, employment, scrumTeams, teamRaw,
             reportsToRaw, org_tab}
    """
    rows = list(ws.iter_rows(min_row=1, values_only=True))
    if not rows:
        return {}

    header = [str(c).strip() if c else "" for c in rows[0]]
    header_lower = [h.lower() for h in header]

    # Find column indices
    def find_col(candidates):
        for c in candidates:
            for i, h in enumerate(header_lower):
                if h == c.lower():
                    return i
        return None

    reports_to_idx = find_col(["Reports To", "reports to", "Managers", "managers"])
    name_idx = find_col(["Name", "name", "Employee Name", "employee name"])
    employment_idx = find_col(["Employment", "employment"])
    title_idx = find_col(["Title", "title"])

    # Team column: try Team, Teams, Teams.1 in priority
    team_idx = find_col(["Team", "team"])
    teams_idx = find_col(["Teams", "teams"])
    teams1_idx = find_col(["Teams.1", "teams.1"])

    if name_idx is None:
        print(f"  [WARN] Tab '{tab_name}' missing Name column")
        return {}

    people = {}
    last_reports_to = None

    for row in rows[1:]:
        if len(row) <= name_idx:
            continue

        raw_name = str(row[name_idx]).strip() if row[name_idx] else ""

        # Forward-fill Reports To (still needed for contractor manager resolution)
        raw_reports_to = ""
        if reports_to_idx is not None and len(row) > reports_to_idx and row[reports_to_idx]:
            raw_reports_to = str(row[reports_to_idx]).strip()
        if raw_reports_to and raw_reports_to.lower() not in ('none', 'nan', ''):
            last_reports_to = raw_reports_to
        else:
            raw_reports_to = last_reports_to or ""

        if not raw_name or raw_name.lower() in ('none', 'nan'):
            continue

        # Skip changelog rows
        if is_changelog_row(raw_reports_to, raw_name):
            continue

        # Employment
        employment = ""
        if employment_idx is not None and len(row) > employment_idx and row[employment_idx]:
            employment = str(row[employment_idx]).strip()
        if not employment:
            # Infer from c-prefix or Title column
            raw_title_val = ""
            if title_idx is not None and len(row) > title_idx and row[title_idx]:
                raw_title_val = str(row[title_idx]).strip()
            if raw_name.lower().startswith('c-'):
                employment = "Contractor"
            elif 'contract' in raw_title_val.lower():
                employment = "Contractor"
            else:
                employment = "Full Time"

        # Team(s)
        team_raw_parts = []
        for tidx in [team_idx, teams_idx, teams1_idx]:
            if tidx is not None and len(row) > tidx and row[tidx]:
                val = str(row[tidx]).strip()
                if val.lower() not in ('none', 'nan', ''):
                    team_raw_parts.append(val)
        team_raw = " / ".join(team_raw_parts) if team_raw_parts else ""
        scrum_teams = parse_scrum_teams(team_raw)

        norm = normalize_name(raw_name)
        people[norm] = {
            "name": raw_name,
            "employment": employment,
            "scrumTeams": scrum_teams,
            "teamRaw": team_raw,
            "reportsToRaw": raw_reports_to,
            "org_tab": tab_name,
        }

    return people


def fuzzy_title_match(name, title_map):
    """Try fuzzy matching for title lookup."""
    norm = normalize_name(name)
    parts = norm.split()
    if len(parts) < 2:
        return None

    first = parts[0]
    last = parts[-1]

    # Try nickname
    first_alt = NICKNAME_MAP.get(first, first)

    for key, title in title_map.items():
        key_parts = key.split()
        if len(key_parts) < 2:
            continue
        k_first = key_parts[0]
        k_last = key_parts[-1]

        # Same last name + first name starts with or vice versa
        if k_last == last:
            if k_first.startswith(first) or first.startswith(k_first):
                return title
            if k_first.startswith(first_alt) or first_alt.startswith(k_first):
                return title

        # Reversed name
        if k_first == last and k_last == first:
            return title

    return None


def parse_org_roster(filepath):
    """Parse all org tabs for employment + teams. Returns merged dict of norm_name -> info."""
    wb = openpyxl.load_workbook(filepath, data_only=True)
    all_people = {}

    for tab_name in ORG_TABS:
        if tab_name not in wb.sheetnames:
            print(f"  [WARN] Org tab '{tab_name}' not found")
            continue
        ws = wb[tab_name]
        people = parse_org_tab(ws, tab_name)
        print(f"  {tab_name}: {len(people)} people")
        # Merge — later tabs don't overwrite earlier ones (first occurrence wins)
        for norm, info in people.items():
            if norm not in all_people:
                all_people[norm] = info
            else:
                # Merge missing fields from later tabs
                existing = all_people[norm]
                if info["scrumTeams"] and not existing["scrumTeams"]:
                    existing["scrumTeams"] = info["scrumTeams"]
                    existing["teamRaw"] = info["teamRaw"]
                if info["reportsToRaw"] and not existing["reportsToRaw"]:
                    existing["reportsToRaw"] = info["reportsToRaw"]

    wb.close()
    print(f"  Total from per-org tabs: {len(all_people)} unique people")
    return all_people


# ─── Step 3: Build Tree Per Org ──────────────────────────────────────────────

def resolve_name_match(target, candidates_by_norm):
    """Resolve a name reference to a node. Returns node or None."""
    norm_target = normalize_name(target)

    # Exact match
    if norm_target in candidates_by_norm:
        return candidates_by_norm[norm_target]

    # Nickname substitution
    parts = norm_target.split()
    if parts:
        alt_first = NICKNAME_MAP.get(parts[0], parts[0])
        alt_name = " ".join([alt_first] + parts[1:])
        if alt_name in candidates_by_norm:
            return candidates_by_norm[alt_name]

    # Partial match: if target is single word, find by first-name contains
    if len(parts) == 1:
        matches = []
        for key, node in candidates_by_norm.items():
            if key.startswith(norm_target) or norm_target in key.split()[0]:
                matches.append(node)
        if len(matches) == 1:
            return matches[0]
        if len(matches) > 1:
            # Pick highest seniority
            matches.sort(key=lambda n: title_seniority_score(n.get("title", "")), reverse=True)
            return matches[0]

    # Last name match + first name prefix
    if len(parts) >= 2:
        first, last = parts[0], parts[-1]
        alt_first = NICKNAME_MAP.get(first, first)
        for key, node in candidates_by_norm.items():
            key_parts = key.split()
            if len(key_parts) >= 2:
                k_first, k_last = key_parts[0], key_parts[-1]
                if k_last == last and (k_first.startswith(first) or first.startswith(k_first) or
                                       k_first.startswith(alt_first) or alt_first.startswith(k_first)):
                    return node

    return None


def resolve_on24_name(target_name, on24_people):
    """Resolve a short/nickname reference to an on24 person. Returns norm key or None."""
    norm = normalize_name(target_name)
    if norm in on24_people:
        return norm

    # Full-name alias lookup
    alias = NAME_ALIASES.get(norm)
    if alias and alias in on24_people:
        return alias

    # Nickname substitution
    parts = norm.split()
    if parts:
        alt_first = NICKNAME_MAP.get(parts[0], parts[0])
        alt_name = " ".join([alt_first] + parts[1:])
        if alt_name in on24_people:
            return alt_name

    # Single word: partial first-name match
    if len(parts) == 1:
        matches = [k for k in on24_people if k.startswith(norm) or norm in k.split()[0]]
        if len(matches) == 1:
            return matches[0]

    # Last name + first name prefix
    if len(parts) >= 2:
        first, last = parts[0], parts[-1]
        alt_first = NICKNAME_MAP.get(first, first)
        for key in on24_people:
            kp = key.split()
            if len(kp) >= 2:
                kf, kl = kp[0], kp[-1]
                if kl == last and (kf.startswith(first) or first.startswith(kf) or
                                   kf.startswith(alt_first) or alt_first.startswith(kf)):
                    return key
    return None


def build_from_on24(on24_people, org_tab_people, teams_hier, title_map):
    """Build all org datasets from on24 hierarchy + per-org enrichment.

    Returns: dict of org_name -> {top, nodes, children}
    """
    jayesh_norm = normalize_name(JAYESH_NAME)
    jayesh_id = slugify(JAYESH_NAME)

    # ── Phase 1: Create nodes from on24 data ──
    all_nodes = {}  # id -> node dict
    norm_to_id = {}  # normalized name -> node id
    children = defaultdict(list)

    for norm, person in on24_people.items():
        node_id = slugify(person["name"])
        # Avoid slug collisions
        if node_id in all_nodes:
            node_id = f"{node_id}-2"

        # Look up employment & scrum teams from per-org tabs
        tab_info = _match_org_tab(norm, person["name"], org_tab_people)

        node = {
            "id": node_id,
            "name": person["name"],
            "title": person["title"],
            "employment": tab_info.get("employment", "") if tab_info else "Full Time",
            "teamRaw": tab_info.get("teamRaw", "") if tab_info else "",
            "scrumTeams": tab_info.get("scrumTeams", []) if tab_info else [],
            "managerId": None,
            "placeholder": False,
            "org": "",  # assigned later
        }
        # Default employment for on24 FTEs
        if not node["employment"]:
            node["employment"] = "Full Time"

        all_nodes[node_id] = node
        norm_to_id[norm] = node_id

    # Override Jayesh title
    if jayesh_norm in norm_to_id:
        all_nodes[norm_to_id[jayesh_norm]]["title"] = JAYESH_TITLE
        jayesh_id = norm_to_id[jayesh_norm]

    # ── Phase 2: Resolve manager IDs from on24 hierarchy ──
    for norm, person in on24_people.items():
        node_id = norm_to_id[norm]
        if node_id == jayesh_id:
            all_nodes[node_id]["managerId"] = None
            continue
        rt_norm = normalize_name(person["reports_to"])
        if rt_norm in norm_to_id:
            mgr_id = norm_to_id[rt_norm]
            all_nodes[node_id]["managerId"] = mgr_id
            children[mgr_id].append(node_id)
        else:
            # Orphan in on24 — assign to Jayesh
            all_nodes[node_id]["managerId"] = jayesh_id
            children[jayesh_id].append(node_id)

    # ── Phase 3: Add people from per-org tabs not already in on24 ──
    slug_counter = defaultdict(int)
    on24_norms = set(norm_to_id.keys())

    for norm, info in org_tab_people.items():
        if norm in on24_norms:
            continue  # Already in on24 tree
        # Alias check: known alternate spellings
        alias = NAME_ALIASES.get(norm)
        if alias and (alias in norm_to_id or alias in on24_norms):
            continue
        # Fuzzy check: short org-tab names may match longer on24 names
        resolved = resolve_on24_name(info["name"], on24_people)
        if resolved and resolved in norm_to_id:
            continue  # Same person under a different name in on24

        name = info["name"]
        base_slug = slugify(name)
        slug_counter[base_slug] += 1
        node_id = base_slug if slug_counter[base_slug] == 1 else f"{base_slug}-{slug_counter[base_slug]}"
        # Avoid collision with on24 nodes
        while node_id in all_nodes:
            slug_counter[base_slug] += 1
            node_id = f"{base_slug}-{slug_counter[base_slug]}"

        # Resolve title: manual overrides > talent snapshot > fuzzy > org default
        title = ""
        if norm in MANUAL_TITLE_OVERRIDES:
            title = MANUAL_TITLE_OVERRIDES[norm]
        elif norm in MANUAL_TITLE_OVERRIDES_EXTRA:
            title = MANUAL_TITLE_OVERRIDES_EXTRA[norm]
        elif norm in title_map:
            title = title_map[norm]
        else:
            t = fuzzy_title_match(name, title_map)
            if t:
                title = t
            else:
                title = DEFAULT_TITLES_BY_ORG.get(info.get("org_tab", ""), "")

        node = {
            "id": node_id,
            "name": name,
            "title": title,
            "employment": info.get("employment", ""),
            "teamRaw": info.get("teamRaw", ""),
            "scrumTeams": info.get("scrumTeams", []),
            "managerId": None,
            "placeholder": False,
            "org": "",
        }

        # Resolve manager: try team lead from Teams Hierarchy, then per-org Reports To
        mgr_id = _resolve_contractor_manager(
            info, norm_to_id, on24_people, org_tab_people, teams_hier, all_nodes
        )
        node["managerId"] = mgr_id or jayesh_id
        all_nodes[node_id] = node
        norm_to_id[norm] = node_id
        if node["managerId"]:
            children[node["managerId"]].append(node_id)

    # ── Phase 4: QA org fix — automation contractors under Ashish ──
    oleg_norm = resolve_on24_name("Oleg", on24_people)
    ashish_norm = _match_org_tab_norm("ashish", org_tab_people)
    if oleg_norm and ashish_norm:
        oleg_id = norm_to_id.get(oleg_norm)
        ashish_id = norm_to_id.get(ashish_norm)
        if oleg_id and not ashish_id:
            # Ashish is a contractor not in on24 — create node if needed
            ashish_info = org_tab_people.get(ashish_norm)
            if ashish_info and ashish_norm in norm_to_id:
                ashish_id = norm_to_id[ashish_norm]
        if oleg_id and ashish_id:
            # Ensure Ashish reports to Oleg
            all_nodes[ashish_id]["managerId"] = oleg_id
            if ashish_id not in children.get(oleg_id, []):
                children[oleg_id].append(ashish_id)
            # Reroute non-real DRs of Oleg to Ashish
            oleg_children = list(children.get(oleg_id, []))
            for cid in oleg_children:
                if cid == ashish_id:
                    continue
                child_norm = normalize_name(all_nodes[cid]["name"])
                first_name = child_norm.split()[0] if child_norm else ""
                if first_name not in QA_OLEG_REAL_DRS:
                    all_nodes[cid]["managerId"] = ashish_id
                    children[oleg_id].remove(cid)
                    children[ashish_id].append(cid)

    # ── Phase 5: Assign orgs based on DR subtree ──
    # Build DR -> org mapping
    dr_org = {}  # node_id -> org_name
    for dr_hint in JAYESH_DRS:
        dr_norm = resolve_on24_name(dr_hint, on24_people)
        if not dr_norm:
            dr_norm = _match_org_tab_norm(normalize_name(dr_hint), org_tab_people)
        if dr_norm and dr_norm in norm_to_id:
            dr_id = norm_to_id[dr_norm]
            # Determine org from DR_ORG_MAP
            for name_key, org_name in DR_ORG_MAP.items():
                if name_key == dr_norm or dr_norm.startswith(name_key.split()[0]):
                    dr_org[dr_id] = org_name
                    break

    # BFS from each DR to assign org to entire subtree
    node_orgs = {}  # node_id -> org_name
    node_orgs[jayesh_id] = "Home"

    for dr_id, org_name in dr_org.items():
        queue = [dr_id]
        while queue:
            nid = queue.pop(0)
            if nid in node_orgs:
                continue
            node_orgs[nid] = org_name
            for child_id in children.get(nid, []):
                queue.append(child_id)

    # Assign org to nodes
    for nid, node in all_nodes.items():
        node["org"] = node_orgs.get(nid, "Full Dev Org")

    # ── Phase 6: Split into per-org datasets ──
    org_datasets = {}
    for org_name in list(dict.fromkeys(
        ["Product-Design", "Full QA Org", "Full Dev Org", SALESFORCE_ORG_NAME, "TPM"]
    )):
        org_node_ids = {nid for nid, o in node_orgs.items() if o == org_name}
        if not org_node_ids:
            continue

        # Always include Jayesh at top
        org_nodes = {jayesh_id: all_nodes[jayesh_id].copy()}
        org_nodes[jayesh_id]["org"] = org_name
        org_children = defaultdict(list)

        # Jayesh's DRs for this org
        for dr_id, dr_org_name in dr_org.items():
            if dr_org_name == org_name:
                org_children[jayesh_id].append(dr_id)

        # Add all nodes in this org
        for nid in org_node_ids:
            org_nodes[nid] = all_nodes[nid]
            # Add children that are also in this org
            for cid in children.get(nid, []):
                if cid in org_node_ids:
                    org_children[nid].append(cid)

        # Sort children alphabetically
        for pid in org_children:
            org_children[pid].sort(
                key=lambda cid: org_nodes.get(cid, all_nodes.get(cid, {})).get("name", "").lower()
            )

        org_datasets[org_name] = {
            "top": jayesh_id,
            "nodes": org_nodes,
            "children": dict(org_children),
        }

        # Report
        dr_names = [org_nodes.get(d, {}).get("name", "?") for d in org_children.get(jayesh_id, [])]
        print(f"  {org_name}: {len(org_nodes)} nodes, DRs: {dr_names}")

    return org_datasets


def _match_org_tab(norm, name, org_tab_people):
    """Find a person in org_tab_people by normalized name, with fuzzy fallback."""
    if norm in org_tab_people:
        return org_tab_people[norm]

    # Try nickname substitution
    parts = norm.split()
    if parts:
        alt_first = NICKNAME_MAP.get(parts[0], parts[0])
        alt_name = " ".join([alt_first] + parts[1:])
        if alt_name in org_tab_people:
            return org_tab_people[alt_name]

    # Partial: last name + first name prefix
    if len(parts) >= 2:
        first, last = parts[0], parts[-1]
        alt_first = NICKNAME_MAP.get(first, first)
        for key, info in org_tab_people.items():
            kp = key.split()
            if len(kp) >= 2:
                kf, kl = kp[0], kp[-1]
                if kl == last and (kf.startswith(first) or first.startswith(kf) or
                                   kf.startswith(alt_first) or alt_first.startswith(kf)):
                    return info

    return None


def _match_org_tab_norm(target_norm, org_tab_people):
    """Find normalized key in org_tab_people, with fuzzy matching. Returns norm key."""
    if target_norm in org_tab_people:
        return target_norm
    parts = target_norm.split()
    if parts:
        alt_first = NICKNAME_MAP.get(parts[0], parts[0])
        alt_name = " ".join([alt_first] + parts[1:])
        if alt_name in org_tab_people:
            return alt_name
    # Partial first-name match
    if len(parts) == 1:
        matches = [k for k in org_tab_people if k.startswith(target_norm)]
        if len(matches) == 1:
            return matches[0]
    return None


def _resolve_contractor_manager(info, norm_to_id, on24_people, org_tab_people, teams_hier, all_nodes):
    """Resolve which on24 person a contractor should report to.

    Priority: team lead from Teams Hierarchy > per-org Reports To.
    """
    # Try Teams Hierarchy: find lead for contractor's scrum team
    for team in info.get("scrumTeams", []):
        th = teams_hier.get(team)
        if not th:
            continue
        # Determine discipline from org_tab
        org_tab = info.get("org_tab", "")
        if "QA" in org_tab:
            lead_names = th.get("qa_leads", [])
            if not lead_names:
                lead_names = [th.get("qa_director", "")]
        else:
            lead_names = th.get("dev_leads", [])
            if not lead_names:
                lead_names = [th.get("dev_director", "")]

        for lead_name in lead_names:
            if not lead_name:
                continue
            lead_norm = resolve_on24_name(lead_name, on24_people)
            if lead_norm and lead_norm in norm_to_id:
                return norm_to_id[lead_norm]

    # Fallback: use Reports To from per-org tab
    rt = info.get("reportsToRaw", "")
    if rt:
        rt_norm = normalize_name(rt)
        # Try direct match in norm_to_id
        if rt_norm in norm_to_id:
            return norm_to_id[rt_norm]
        # Try on24 resolve
        resolved = resolve_on24_name(rt, on24_people)
        if resolved and resolved in norm_to_id:
            return norm_to_id[resolved]
        # Try org_tab resolve (for contractor→contractor chains)
        tab_norm = _match_org_tab_norm(rt_norm, org_tab_people)
        if tab_norm and tab_norm in norm_to_id:
            return norm_to_id[tab_norm]

    return None


# ─── Step 4: Build Scrum Index ───────────────────────────────────────────────

def build_scrum_index(org_datasets, teams_hier=None):
    """Build global scrum index: team_name -> list of members.

    Uses teams_hier (from SCRUMS.xlsx) to explicitly identify leads.
    """
    teams_hier = teams_hier or {}

    # Build lead name sets per team for quick lookup
    team_lead_names = {}  # team_name_lower -> set of normalized lead names
    for team_name, th in teams_hier.items():
        lead_norms = set()
        for lead in th.get("dev_leads", []) + th.get("qa_leads", []):
            if lead:
                lead_norms.add(normalize_name(lead))
        team_lead_names[team_name.lower()] = lead_norms

    # First pass: collect all teams with case-insensitive dedup
    canonical_names = {}
    raw_index = defaultdict(list)

    for tab_name, dataset in org_datasets.items():
        for node_id, node in dataset["nodes"].items():
            if node.get("placeholder"):
                continue
            for team in node.get("scrumTeams", []):
                if team.lower() in ("dir", "management", ""):
                    continue
                key = team.lower().strip()
                if key not in canonical_names:
                    canonical_names[key] = team

                # Check if this person is a lead for this team
                person_norm = normalize_name(node["name"])
                person_first = person_norm.split()[0] if person_norm else ""
                lead_set = team_lead_names.get(key, set())
                is_lead = (person_norm in lead_set or
                           any(person_first and ln.startswith(person_first) for ln in lead_set))

                raw_index[key].append({
                    "org": tab_name,
                    "id": node_id,
                    "name": node["name"],
                    "title": node.get("title", ""),
                    "employment": node.get("employment", ""),
                    "isLead": is_lead,
                })

    # Build canonical scrum_index
    scrum_index = {}
    for key, members in raw_index.items():
        display = canonical_names[key]
        scrum_index[display] = members

    # Determine discipline and sort (leads first)
    scrum_teams = {}
    for team_name, members in scrum_index.items():
        grouped = {"Dev": [], "QA": [], "Product": [], "TPM": [], "Salesforce": [], "Other": []}
        for m in members:
            org = m["org"]
            if org == SALESFORCE_ORG_NAME:
                grouped["Salesforce"].append(m)
            elif "Dev" in org:
                grouped["Dev"].append(m)
            elif "QA" in org:
                grouped["QA"].append(m)
            elif "Product" in org or "Design" in org:
                grouped["Product"].append(m)
            elif "TPM" in org:
                grouped["TPM"].append(m)
            else:
                grouped["Other"].append(m)

        # Sort: leads first, then by seniority
        for discipline in grouped:
            group = grouped[discipline]
            if group:
                group.sort(key=lambda m: (
                    0 if m.get("isLead") else 1,
                    -title_seniority_score(m["title"]),
                    m["name"].lower()
                ))

        scrum_teams[team_name] = grouped

    return scrum_teams


# ─── Step 5: HTML Generation ─────────────────────────────────────────────────

def make_serializable(org_datasets, scrum_teams):
    """Convert data to JSON-serializable format."""
    orgs = {}
    # Build Home view: Jayesh + his 6 DRs across all orgs (deduplicated)
    home_drs = []
    seen_dr_names = set()

    for tab_name, dataset in org_datasets.items():
        nodes_ser = {}
        top_id = dataset["top"]
        top_children = dataset["children"].get(top_id, [])

        for nid, node in dataset["nodes"].items():
            ser = {
                "id": node["id"],
                "name": node["name"],
                "title": node.get("title", ""),
                "employment": node.get("employment", ""),
                "scrumTeams": node.get("scrumTeams", []),
                "placeholder": node.get("placeholder", False),
                "org": node.get("org", tab_name),
            }
            if node.get("dottedLine"):
                ser["dottedLine"] = node["dottedLine"]
            nodes_ser[nid] = ser

        orgs[tab_name] = {
            "top": top_id,
            "nodes": nodes_ser,
            "children": dataset["children"],
        }

        # Collect Jayesh's DRs for the Home view (dedup across orgs)
        for cid in top_children:
            c = nodes_ser.get(cid)
            if c:
                dr_norm = normalize_name(c["name"])
                if dr_norm not in seen_dr_names:
                    seen_dr_names.add(dr_norm)
                    home_drs.append({
                        "name": c["name"],
                        "title": c.get("title", ""),
                        "employment": c.get("employment", ""),
                        "org": tab_name,
                        "nodeId": cid,
                    })

    return {
        "orgs": orgs,
        "scrum": scrum_teams,
        "missing": {},
        "homeDrs": home_drs,
    }


def redact_data(data, all_names):
    """Deep copy and replace all names with Person NNN + anonymize node IDs."""
    data = copy.deepcopy(data)

    # Build consistent name->redacted mapping (initial + blacked-out letters)
    name_to_redacted = {}
    used_redacted = set()
    collision_counter = {}

    def get_redacted(name):
        norm = normalize_name(name)
        if norm not in name_to_redacted:
            parts = norm.split()
            redacted_parts = []
            for p in parts:
                if p:
                    redacted_parts.append(p[0].upper() + "\u2588" * (len(p) - 1))
            result = " ".join(redacted_parts) if redacted_parts else "X"
            if result in used_redacted:
                collision_counter[result] = collision_counter.get(result, 1) + 1
                result = result + str(collision_counter[result])
            used_redacted.add(result)
            name_to_redacted[norm] = result
        return name_to_redacted[norm]

    # Build consistent id->anonymized id mapping
    id_map = {}
    id_counter = [0]

    def get_anon_id(old_id):
        if old_id not in id_map:
            id_counter[0] += 1
            id_map[old_id] = f"node-{id_counter[0]:03d}"
        return id_map[old_id]

    # First pass: collect all node IDs to build the mapping
    for tab_name, org in data["orgs"].items():
        get_anon_id(org["top"])
        for nid in org["nodes"]:
            get_anon_id(nid)
        for parent_id, child_ids in org["children"].items():
            get_anon_id(parent_id)
            for cid in child_ids:
                get_anon_id(cid)

    # Redact org nodes: names, dottedLine, and IDs
    for tab_name, org in data["orgs"].items():
        # Remap top
        org["top"] = get_anon_id(org["top"])

        # Remap nodes dict
        new_nodes = {}
        for nid, node in org["nodes"].items():
            node["name"] = get_redacted(node["name"])
            if node.get("dottedLine"):
                node["dottedLine"] = get_redacted(node["dottedLine"])
            new_id = get_anon_id(nid)
            node["id"] = new_id
            new_nodes[new_id] = node
        org["nodes"] = new_nodes

        # Remap children dict
        new_children = {}
        for parent_id, child_ids in org["children"].items():
            new_parent = get_anon_id(parent_id)
            new_children[new_parent] = [get_anon_id(c) for c in child_ids]
        org["children"] = new_children

    # Redact scrum members: names and IDs
    for team_name, groups in data["scrum"].items():
        for discipline, members in groups.items():
            for m in members:
                m["name"] = get_redacted(m["name"])
                if "id" in m:
                    m["id"] = get_anon_id(m["id"])

    # Redact missing titles lists
    for tab_name, names in data["missing"].items():
        data["missing"][tab_name] = [get_redacted(n) for n in names]

    # Redact homeDrs: names and nodeIds
    for dr in data.get("homeDrs", []):
        dr["name"] = get_redacted(dr["name"])
        if "nodeId" in dr:
            dr["nodeId"] = get_anon_id(dr["nodeId"])

    return data


def verify_redaction(html_content, all_names):
    """Verify no real names appear in the HTML."""
    html_lower = html_content.lower()
    leaked = []
    for name in all_names:
        if len(name) < 4:
            continue  # Skip very short names to avoid false positives
        if normalize_name(name) in html_lower:
            # Check it's not a substring of something else
            parts = normalize_name(name).split()
            if len(parts) >= 2:
                leaked.append(name)
    return leaked


def generate_html(data, redacted=False):
    """Generate the complete standalone HTML file."""

    title_suffix = " (Redacted)" if redacted else ""
    data_json = json.dumps(data, ensure_ascii=False, indent=None)

    # Use a raw template with placeholder tokens instead of f-string
    # to avoid all the {{/}} and \' escaping issues with embedded JS
    html = _HTML_TEMPLATE.replace("__TITLE_SUFFIX__", title_suffix).replace("__DATA_JSON__", data_json)
    return html


# The entire HTML template as a plain string (no f-string).
# Only __TITLE_SUFFIX__ and __DATA_JSON__ get replaced.
_HTML_TEMPLATE = r'''<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>Org Chart__TITLE_SUFFIX__</title>
<style>
* { margin: 0; padding: 0; box-sizing: border-box; }

body {
    font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, 'Helvetica Neue', Arial, sans-serif;
    background: #f5f7fa;
    color: #1a1a2e;
    min-height: 100vh;
}

/* Header */
.header {
    background: linear-gradient(135deg, #1a1a2e 0%, #16213e 50%, #0f3460 100%);
    color: white;
    padding: 16px 24px;
    display: flex;
    align-items: center;
    gap: 16px;
    flex-wrap: wrap;
    box-shadow: 0 2px 12px rgba(0,0,0,0.15);
    position: sticky;
    top: 0;
    z-index: 100;
}

.header h1 {
    font-size: 20px;
    font-weight: 600;
    white-space: nowrap;
}

.header-controls {
    display: flex;
    align-items: center;
    gap: 12px;
    flex-wrap: wrap;
    flex: 1;
}

select, input[type="text"] {
    padding: 8px 12px;
    border: 1px solid rgba(255,255,255,0.2);
    border-radius: 8px;
    background: rgba(255,255,255,0.1);
    color: white;
    font-size: 14px;
    outline: none;
    transition: border-color 0.2s;
}
select:focus, input[type="text"]:focus {
    border-color: rgba(255,255,255,0.5);
}
select option {
    background: #1a1a2e;
    color: white;
}
input[type="text"]::placeholder {
    color: rgba(255,255,255,0.5);
}

.switch-group {
    display: flex;
    align-items: center;
    gap: 14px;
}
.switch-label {
    display: flex;
    align-items: center;
    gap: 6px;
    cursor: pointer;
    font-size: 13px;
    color: rgba(255,255,255,0.85);
    white-space: nowrap;
    user-select: none;
}
.switch {
    position: relative;
    width: 44px;
    height: 24px;
    background: rgba(255,255,255,0.25);
    border-radius: 12px;
    cursor: pointer;
    transition: background 0.3s;
    flex-shrink: 0;
}
.switch.active {
    background: #34c759;
}
.switch-knob {
    position: absolute;
    top: 2px;
    left: 2px;
    width: 20px;
    height: 20px;
    background: #fff;
    border-radius: 50%;
    transition: transform 0.3s;
    box-shadow: 0 1px 3px rgba(0,0,0,0.3);
}
.switch.active .switch-knob {
    transform: translateX(20px);
}

/* Navigation */
.nav-bar {
    background: white;
    padding: 12px 24px;
    display: flex;
    align-items: center;
    gap: 8px;
    border-bottom: 1px solid #e8ecf1;
    flex-wrap: wrap;
}

.nav-btn {
    padding: 6px 14px;
    border: 1px solid #d1d9e6;
    border-radius: 6px;
    background: white;
    color: #4a5568;
    cursor: pointer;
    font-size: 13px;
    transition: all 0.2s;
}
.nav-btn:hover {
    background: #f0f4f8;
    border-color: #a0aec0;
}

.breadcrumb {
    display: flex;
    align-items: center;
    gap: 4px;
    flex-wrap: wrap;
    flex: 1;
}
.breadcrumb span {
    color: #718096;
    font-size: 13px;
}
.breadcrumb a {
    color: #3182ce;
    text-decoration: none;
    font-size: 13px;
    cursor: pointer;
}
.breadcrumb a:hover {
    text-decoration: underline;
}
.breadcrumb .current {
    color: #1a1a2e;
    font-weight: 600;
    font-size: 13px;
}

/* Main content */
.main {
    padding: 32px 24px;
    max-width: 1400px;
    margin: 0 auto;
}

/* Manager card */
.manager-section {
    display: flex;
    flex-direction: column;
    align-items: center;
    margin-bottom: 32px;
}

.manager-card {
    background: linear-gradient(135deg, #1a1a2e 0%, #16213e 100%);
    color: white;
    border-radius: 16px;
    padding: 24px 32px;
    text-align: center;
    box-shadow: 0 8px 32px rgba(26,26,46,0.2);
    min-width: 280px;
    max-width: 400px;
}
.manager-card .name {
    font-size: 20px;
    font-weight: 700;
    margin-bottom: 4px;
}
.manager-card .title {
    font-size: 14px;
    opacity: 0.85;
    margin-bottom: 8px;
}
.manager-card .badge {
    display: inline-block;
    padding: 2px 10px;
    border-radius: 12px;
    font-size: 11px;
    font-weight: 600;
    text-transform: uppercase;
    letter-spacing: 0.5px;
}
.badge-fte {
    background: rgba(72,187,120,0.2);
    color: #68d391;
}
.badge-contractor {
    background: rgba(237,137,54,0.2);
    color: #f6ad55;
}
.manager-card .teams {
    margin-top: 10px;
    display: flex;
    flex-wrap: wrap;
    justify-content: center;
    gap: 6px;
}

/* Connector */
.connector {
    width: 2px;
    height: 32px;
    background: #cbd5e0;
    margin: 0 auto;
}

.connector-h {
    height: 2px;
    background: #cbd5e0;
    margin: 0 auto 24px;
}

/* Reports grid */
.reports-label {
    text-align: center;
    font-size: 13px;
    color: #718096;
    margin-bottom: 16px;
    font-weight: 500;
}

.reports-grid {
    display: grid;
    grid-template-columns: repeat(auto-fill, minmax(220px, 1fr));
    gap: 16px;
    max-width: 1200px;
    margin: 0 auto;
}

/* Person card */
.person-card {
    background: white;
    border-radius: 12px;
    padding: 20px;
    box-shadow: 0 2px 8px rgba(0,0,0,0.06);
    border: 1px solid #e8ecf1;
    cursor: pointer;
    transition: all 0.2s;
    position: relative;
}
.person-card:hover {
    transform: translateY(-2px);
    box-shadow: 0 6px 20px rgba(0,0,0,0.1);
    border-color: #3182ce;
}
.person-card .name {
    font-size: 15px;
    font-weight: 600;
    color: #1a1a2e;
    margin-bottom: 4px;
}
.person-card .title {
    font-size: 12px;
    color: #718096;
    margin-bottom: 8px;
    line-height: 1.4;
}
.person-card .badge {
    display: inline-block;
    padding: 2px 8px;
    border-radius: 10px;
    font-size: 10px;
    font-weight: 600;
    text-transform: uppercase;
    letter-spacing: 0.5px;
}
.person-card .dr-count {
    font-size: 11px;
    color: #a0aec0;
    margin-top: 8px;
}
.person-card .teams {
    margin-top: 8px;
    display: flex;
    flex-wrap: wrap;
    gap: 4px;
}
.person-card.placeholder {
    border-style: dashed;
    opacity: 0.7;
}

/* Team pill */
.team-pill {
    display: inline-block;
    padding: 2px 8px;
    border-radius: 10px;
    font-size: 10px;
    font-weight: 500;
    cursor: pointer;
    transition: opacity 0.2s;
    text-decoration: none;
}
.team-pill:hover {
    opacity: 0.8;
}

/* Scrum view */
.scrum-view {
    max-width: 900px;
    margin: 0 auto;
}
.scrum-header {
    font-size: 24px;
    font-weight: 700;
    color: #1a1a2e;
    margin-bottom: 24px;
}
.discipline-section {
    margin-bottom: 24px;
    background: white;
    border-radius: 12px;
    padding: 20px;
    box-shadow: 0 2px 8px rgba(0,0,0,0.06);
    border: 1px solid #e8ecf1;
}
.discipline-title {
    font-size: 14px;
    font-weight: 700;
    color: #4a5568;
    text-transform: uppercase;
    letter-spacing: 1px;
    margin-bottom: 12px;
    padding-bottom: 8px;
    border-bottom: 2px solid #e8ecf1;
}
.scrum-member {
    display: flex;
    align-items: center;
    gap: 12px;
    padding: 8px 0;
    border-bottom: 1px solid #f0f4f8;
}
.scrum-member:last-child {
    border-bottom: none;
}
.scrum-member.is-lead {
    font-weight: 700;
}
.scrum-member a {
    color: #3182ce;
    text-decoration: none;
}
.scrum-member a:hover {
    text-decoration: underline;
}
.scrum-member .member-title {
    color: #718096;
    font-size: 12px;
    font-weight: 400;
}

/* Missing titles */
.missing-section {
    margin-top: 40px;
    background: white;
    border-radius: 12px;
    padding: 20px;
    box-shadow: 0 2px 8px rgba(0,0,0,0.06);
    border: 1px solid #e8ecf1;
}
.missing-toggle {
    cursor: pointer;
    color: #718096;
    font-size: 13px;
}
.missing-toggle:hover {
    color: #4a5568;
}
.missing-list {
    display: none;
    margin-top: 12px;
}
.missing-list.open {
    display: block;
}
.missing-list ul {
    list-style: none;
    padding: 0;
}
.missing-list li {
    padding: 4px 0;
    font-size: 13px;
    color: #718096;
}

/* Empty state */
.empty-state {
    text-align: center;
    color: #a0aec0;
    padding: 48px;
    font-size: 15px;
}

/* List view table */
.list-view { overflow-x: auto; background: #fff; border-radius: 12px; box-shadow: 0 2px 8px rgba(0,0,0,0.08); margin: 4px; }
.list-view table { width: 100%; border-collapse: collapse; font-size: 13.5px; }
.list-view th { cursor: pointer; user-select: none; }
.list-view th, .list-view td { text-align: left; padding: 11px 14px; border-bottom: 1px solid #eef1f6; }
.list-view th { background: linear-gradient(180deg, #f5f7fa 0%, #edf0f7 100%); color: #475569; font-weight: 600; font-size: 11.5px; text-transform: uppercase; letter-spacing: 0.6px; position: sticky; top: 0; z-index: 1; border-bottom: 2px solid #dde3ed; }
.list-view th .sort-arrow { font-size: 10px; margin-left: 4px; opacity: 0.3; }
.list-view th .sort-arrow.active { opacity: 1; color: #3b82f6; }
.list-view tbody tr:nth-child(even) { background: #f9fafb; }
.list-view tbody tr { transition: background 0.15s ease; }
.list-view tbody tr:hover { background: #eef3ff; }
.list-view td a { color: #2563eb; cursor: pointer; text-decoration: none; font-weight: 500; }
.list-view td a:hover { color: #1d4ed8; text-decoration: underline; }
.list-view .team-pill { display: inline-block; padding: 2px 10px; border-radius: 12px; font-size: 11px; margin: 1px 3px; text-decoration: none; font-weight: 500; transition: opacity 0.15s ease; }
.list-view .team-pill:hover { opacity: 0.75; }
.list-view .badge { font-size: 11px; padding: 2px 10px; border-radius: 10px; font-weight: 600; }
.list-view .badge-fte { background: #dbeafe; color: #1e40af; }
.list-view .badge-contractor { background: #fef3c7; color: #92400e; }

/* Headcount bar */
.headcount {
    text-align: center;
    font-size: 12px;
    color: rgba(255,255,255,0.7);
    white-space: nowrap;
}

/* Responsive */
@media (max-width: 768px) {
    .header { flex-direction: column; align-items: flex-start; padding: 12px 16px; gap: 10px; }
    .header h1 { font-size: 16px; }
    .header-controls { width: 100%; gap: 8px; }
    .header-controls select,
    .header-controls input { font-size: 13px; min-width: 0; flex: 1; }
    .breadcrumb { padding: 8px 16px; font-size: 12px; }
    .main-content { padding: 16px; }
    .manager-card { padding: 20px; }
    .manager-card .name { font-size: 18px; }
    .reports-grid { grid-template-columns: 1fr; gap: 12px; }
    .person-card { padding: 14px; }
    .person-card .name { font-size: 14px; }
    .scrum-view { padding: 16px; }
}

@media (max-width: 480px) {
    .header { padding: 10px 12px; }
    .header h1 { font-size: 14px; }
    .header-controls { flex-direction: column; }
    .manager-card { padding: 16px; }
    .reports-grid { grid-template-columns: 1fr; }
}
</style>
</head>
<body>

<div class="header">
    <h1>Org Chart__TITLE_SUFFIX__</h1>
    <div class="header-controls">
        <select id="orgSelect" onchange="switchOrg(this.value)"></select>
        <input type="text" id="searchBox" placeholder="Search by name..." onkeydown="if(event.key==='Enter')doSearch()">
        <div class="switch-group">
            <label class="switch-label">
                <div class="switch active" id="empToggle" onclick="toggleFilter('emp')"><div class="switch-knob"></div></div>
                <span>Employees</span>
            </label>
            <label class="switch-label">
                <div class="switch active" id="conToggle" onclick="toggleFilter('con')"><div class="switch-knob"></div></div>
                <span>Contractors</span>
            </label>
            <label class="switch-label" id="redactGroup" style="display:none">
                <div class="switch" id="redactToggle" onclick="toggleFilter('redact')"><div class="switch-knob"></div></div>
                <span>Redact</span>
            </label>
        </div>
        <span class="headcount" id="headcount"></span>
    </div>
</div>

<div class="nav-bar">
    <button class="nav-btn" id="homeBtn" onclick="goHome()">Home</button>
    <button class="nav-btn" id="topBtn" onclick="goTop()">Top</button>
    <button class="nav-btn" id="upBtn" onclick="goUp()">Up</button>
    <button class="nav-btn" id="backToOrg" onclick="backToOrg()" style="display:none">Back to Org</button>
    <button class="nav-btn" id="listBtn" onclick="showListView()">List</button>
    <div class="breadcrumb" id="breadcrumb"></div>
</div>

<div class="main" id="mainContent"></div>

<script>
const DATA = __DATA_JSON__;

let state = {
    currentOrg: null,
    currentNodeId: null,
    breadcrumb: [],
    showEmp: true,
    showCon: true,
    redacted: false,
    scrumView: null,
    lastOrgNodeId: null,
    isHome: true,
    listView: false,
    listSortCol: 'name',
    listSortAsc: true,
};

// ── Color palette for team pills ──
const TEAM_COLORS = [
    '#3182ce','#e94560','#38a169','#d69e2e','#805ad5','#dd6b20',
    '#319795','#b83280','#2b6cb0','#c05621','#2f855a','#6b46c1'
];

function teamColor(name) {
    let hash = 0;
    for (let i = 0; i < name.length; i++) hash = ((hash << 5) - hash) + name.charCodeAt(i);
    return TEAM_COLORS[Math.abs(hash) % TEAM_COLORS.length];
}

function escHtml(s) {
    if (!s) return '';
    return s.replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;');
}

// ── Init ──
function init() {
    const sel = document.getElementById('orgSelect');
    // Add Home option first
    const homeOpt = document.createElement('option');
    homeOpt.value = '__HOME__';
    homeOpt.textContent = 'Home (All Orgs)';
    sel.appendChild(homeOpt);
    const orgNames = Object.keys(DATA.orgs);
    orgNames.forEach(name => {
        const opt = document.createElement('option');
        opt.value = name;
        opt.textContent = name;
        sel.appendChild(opt);
    });
    goHome();
}

function switchOrg(orgName) {
    if (orgName === '__HOME__') { goHome(); return; }
    state.currentOrg = orgName;
    state.scrumView = null;
    state.listView = false;
    state.isHome = false;
    navigateTo(DATA.orgs[orgName].top);
}

function goHome() {
    state.isHome = true;
    state.scrumView = null;
    state.listView = false;
    state.currentOrg = null;
    document.getElementById('orgSelect').value = '__HOME__';
    document.getElementById('backToOrg').style.display = 'none';
    renderHome();
}

function renderHome() {
    // Get top person's info from the first org dataset
    const firstOrg = Object.values(DATA.orgs)[0];
    const topPerson = firstOrg ? firstOrg.nodes[firstOrg.top] : null;
    const topName = topPerson ? topPerson.name : 'Leader';
    const topTitle = topPerson ? topPerson.title : '';

    const bc = document.getElementById('breadcrumb');
    bc.innerHTML = '<span class="current">' + escHtml(displayName(topName)) + (topTitle ? ' — ' + escHtml(topTitle) : '') + '</span>';

    let html = '<div class="manager-section">';
    html += '<div class="manager-card">';
    html += '<div class="name">' + escHtml(displayName(topName)) + '</div>';
    if (topTitle) html += '<div class="title">' + escHtml(topTitle) + '</div>';
    html += '<span class="badge badge-fte">FTE</span>';
    html += '</div>';
    html += '<div class="connector"></div>';
    html += '</div>';

    const drs = DATA.homeDrs || [];
    html += '<div class="reports-label">' + drs.length + ' direct reports (across all orgs)</div>';
    html += '<div class="reports-grid">';
    drs.forEach(dr => {
        const badge = isContractor(dr.employment)
            ? '<span class="badge badge-contractor">Contractor</span>'
            : '<span class="badge badge-fte">FTE</span>';
        html += '<div class="person-card" onclick="switchToOrgDr(\'' + escHtml(dr.org) + '\',\'' + escHtml(dr.nodeId) + '\')">';
        html += '<div class="name">' + escHtml(displayName(dr.name)) + '</div>';
        if (dr.title) html += '<div class="title">' + escHtml(dr.title) + '</div>';
        html += badge;
        html += '<div class="dr-count" style="color:#3182ce">' + escHtml(dr.org) + '</div>';
        html += '</div>';
    });
    html += '</div>';

    document.getElementById('mainContent').innerHTML = html;
    // Show total headcount across all orgs (deduplicated)
    let total = 0, fte = 0, contractors = 0;
    const seenNames = {};
    for (const [orgName, org] of Object.entries(DATA.orgs)) {
        for (const [id, node] of Object.entries(org.nodes)) {
            if (node.placeholder) continue;
            const nk = node.name.toLowerCase();
            if (seenNames[nk]) continue;
            seenNames[nk] = true;
            total++; if (isContractor(node.employment)) contractors++; else fte++;
        }
    }
    const el = document.getElementById('headcount');
    el.textContent = total + ' people (' + fte + ' FTE, ' + contractors + ' contractors) across all orgs';
}

function switchToOrgDr(orgName, nodeId) {
    state.isHome = false;
    state.currentOrg = orgName;
    document.getElementById('orgSelect').value = orgName;
    navigateTo(nodeId);
}

function toggleFilter(which) {
    if (which === 'emp') state.showEmp = !state.showEmp;
    if (which === 'con') state.showCon = !state.showCon;
    if (which === 'redact') state.redacted = !state.redacted;
    document.getElementById('empToggle').classList.toggle('active', state.showEmp);
    document.getElementById('conToggle').classList.toggle('active', state.showCon);
    document.getElementById('redactToggle').classList.toggle('active', state.redacted);
    if (state.listView) renderList();
    else if (state.isHome) renderHome();
    else if (state.scrumView) showScrumView(state.scrumView);
    else render();
}

function displayName(name) {
    if (!state.redacted) return name;
    return name.split(/\s+/).map(function(w) {
        if (!w) return '';
        return w[0].toUpperCase() + '\u2588'.repeat(w.length - 1);
    }).join(' ');
}

function isContractor(emp) {
    if (!emp) return false;
    const e = emp.toLowerCase();
    return e.includes('contract') || e.includes('vendor') || e.includes('consult');
}

// ── Navigation ──
function navigateTo(nodeId) {
    if (!state.currentOrg) return;
    const org = DATA.orgs[state.currentOrg];
    if (!org || !org.nodes[nodeId]) return;
    state.isHome = false;

    state.scrumView = null;
    state.listView = false;
    document.getElementById('backToOrg').style.display = 'none';

    // Build breadcrumb
    const trail = [];
    let cur = nodeId;
    const visited = new Set();
    while (cur && org.nodes[cur] && !visited.has(cur)) {
        visited.add(cur);
        trail.unshift(cur);
        const node = org.nodes[cur];
        // Find parent
        let parent = null;
        for (const [pid, children] of Object.entries(org.children)) {
            if (children.includes(cur)) { parent = pid; break; }
        }
        cur = parent;
    }
    state.breadcrumb = trail;
    state.currentNodeId = nodeId;
    state.lastOrgNodeId = nodeId;
    render();
}

function goTop() {
    if (!state.currentOrg) { goHome(); return; }
    navigateTo(DATA.orgs[state.currentOrg].top);
}

function goUp() {
    if (state.breadcrumb.length > 1) {
        navigateTo(state.breadcrumb[state.breadcrumb.length - 2]);
    }
}

function backToOrg() {
    if (state.lastOrgNodeId) {
        navigateTo(state.lastOrgNodeId);
    }
}

function doSearch() {
    const q = document.getElementById('searchBox').value.trim().toLowerCase();
    if (!q) return;
    const org = DATA.orgs[state.currentOrg];
    // Search across all orgs
    for (const [orgName, orgData] of Object.entries(DATA.orgs)) {
        for (const [nid, node] of Object.entries(orgData.nodes)) {
            const name = node.name.toLowerCase();
            if (name === q || name.startsWith(q) || name.includes(q)) {
                state.currentOrg = orgName;
                document.getElementById('orgSelect').value = orgName;
                navigateTo(nid);
                return;
            }
        }
    }
    alert('No match found for "' + q + '"');
}

// ── Scrum View ──
function showScrumView(teamName) {
    state.scrumView = teamName;
    state.lastOrgNodeId = state.currentNodeId;
    document.getElementById('backToOrg').style.display = '';
    renderScrum(teamName);
}

function navigateToOrgCard(orgName, nodeId) {
    state.currentOrg = orgName;
    document.getElementById('orgSelect').value = orgName;
    navigateTo(nodeId);
}

function renderScrum(teamName) {
    const groups = DATA.scrum[teamName];
    if (!groups) {
        document.getElementById('mainContent').innerHTML = '<div class="empty-state">No data for team: ' + escHtml(teamName) + '</div>';
        return;
    }

    // Update breadcrumb
    const bc = document.getElementById('breadcrumb');
    bc.innerHTML = '<span class="current">#' + escHtml(teamName) + '</span>';

    let html = '<div class="scrum-view">';
    html += '<div class="scrum-header">#' + escHtml(teamName) + '</div>';

    const disciplineOrder = ['Dev', 'Product', 'QA', 'TPM', 'Other'];
    const disciplineLabels = {
        'Dev': 'Development',
        'Product': 'Product & Design',
        'QA': 'Quality Assurance',
        'TPM': 'Technical Program Management',
        'Other': 'Other',
    };

    for (const disc of disciplineOrder) {
        let members = groups[disc] || [];
        members = members.filter(function(m) {
            var con = isContractor(m.employment);
            return con ? state.showCon : state.showEmp;
        });
        if (members.length === 0) continue;

        html += '<div class="discipline-section">';
        html += '<div class="discipline-title">' + escHtml(disciplineLabels[disc] || disc) + '</div>';

        members.forEach((m, idx) => {
            const isLead = idx === 0;
            const leadLabel = isLead ? ' (Lead)' : '';
            const badge = isContractor(m.employment)
                ? '<span class="badge badge-contractor">Contractor</span>'
                : '<span class="badge badge-fte">FTE</span>';
            html += '<div class="scrum-member' + (isLead ? ' is-lead' : '') + '">';
            html += '<a onclick="navigateToOrgCard(\'' + escHtml(m.org) + "','" + escHtml(m.id) + '\')">' + escHtml(displayName(m.name)) + leadLabel + '</a>';
            html += ' ' + badge;
            if (m.title) html += ' <span class="member-title">' + escHtml(m.title) + '</span>';
            html += '</div>';
        });

        html += '</div>';
    }

    html += '</div>';
    document.getElementById('mainContent').innerHTML = html;
    updateHeadcount();
}

// ── List View ──
function collectListRows() {
    var rows = [];
    var seen = {};
    var orgEntries = [];
    if (state.isHome || !state.currentOrg) {
        for (var orgName in DATA.orgs) orgEntries.push([orgName, DATA.orgs[orgName]]);
    } else {
        orgEntries.push([state.currentOrg, DATA.orgs[state.currentOrg]]);
    }
    orgEntries.forEach(function(entry) {
        var orgName = entry[0], org = entry[1];
        for (var nid in org.nodes) {
            var node = org.nodes[nid];
            var nameKey = node.name.toLowerCase();
            if (seen[nameKey]) continue;
            seen[nameKey] = true;
            if (node.placeholder) continue;
            var con = isContractor(node.employment);
            if (con && !state.showCon) continue;
            if (!con && !state.showEmp) continue;
            // Find parent name
            var parentName = '';
            for (var pid in org.children) {
                if (org.children[pid].indexOf(nid) !== -1) {
                    parentName = org.nodes[pid] ? org.nodes[pid].name : '';
                    break;
                }
            }
            rows.push({
                name: node.name,
                title: node.title || '',
                type: con ? 'Contractor' : 'FTE',
                manager: parentName,
                org: orgName,
                scrumTeams: node.scrumTeams || [],
                nodeId: nid,
                orgKey: orgName,
            });
        }
    });
    return rows;
}

function sortListRows(rows) {
    var col = state.listSortCol;
    var asc = state.listSortAsc;
    rows.sort(function(a, b) {
        var va = (a[col] || '').toLowerCase();
        var vb = (b[col] || '').toLowerCase();
        if (va < vb) return asc ? -1 : 1;
        if (va > vb) return asc ? 1 : -1;
        return 0;
    });
    return rows;
}

function renderListTable(rows) {
    var cols = [
        {key: 'name', label: 'Name'},
        {key: 'title', label: 'Title'},
        {key: 'type', label: 'Type'},
        {key: 'manager', label: 'Manager'},
        {key: 'org', label: 'Org'},
        {key: 'scrumTeams', label: 'Scrum Teams'},
    ];
    var html = '<div class="list-view"><table><thead><tr>';
    cols.forEach(function(c) {
        var sortable = c.key !== 'scrumTeams';
        var arrow = '';
        if (sortable) {
            var isActive = state.listSortCol === c.key;
            var dir = isActive ? (state.listSortAsc ? '&#9650;' : '&#9660;') : '&#9650;';
            arrow = ' <span class="sort-arrow' + (isActive ? ' active' : '') + '">' + dir + '</span>';
        }
        if (sortable) {
            html += '<th onclick="sortListBy(\'' + c.key + '\')">' + c.label + arrow + '</th>';
        } else {
            html += '<th>' + c.label + '</th>';
        }
    });
    html += '</tr></thead><tbody>';
    rows.forEach(function(r) {
        html += '<tr>';
        html += '<td><a onclick="navigateToOrgCard(\'' + escHtml(r.orgKey) + "','" + escHtml(r.nodeId) + '\')">' + escHtml(displayName(r.name)) + '</a></td>';
        html += '<td>' + escHtml(r.title) + '</td>';
        var badgeCls = r.type === 'Contractor' ? 'badge-contractor' : 'badge-fte';
        html += '<td><span class="badge ' + badgeCls + '">' + r.type + '</span></td>';
        html += '<td>' + escHtml(displayName(r.manager)) + '</td>';
        html += '<td>' + escHtml(r.org) + '</td>';
        html += '<td>';
        if (r.scrumTeams.length) {
            r.scrumTeams.forEach(function(t) {
                html += '<a class="team-pill" style="background:' + teamColor(t) + '22;color:' + teamColor(t) + ';cursor:pointer" onclick="showScrumView(\'' + escHtml(t).replace(/'/g,"\\\\'") + '\')">#' + escHtml(t) + '</a>';
            });
        }
        html += '</td>';
        html += '</tr>';
    });
    html += '</tbody></table></div>';
    return html;
}

function showListView() {
    state.listView = true;
    state.scrumView = null;
    document.getElementById('backToOrg').style.display = state.isHome ? 'none' : '';
    renderList();
}

function renderList() {
    var rows = collectListRows();
    rows = sortListRows(rows);

    var label = state.isHome ? 'All Orgs' : state.currentOrg;
    var bc = document.getElementById('breadcrumb');
    bc.innerHTML = '<span class="current">List — ' + escHtml(label) + ' (' + rows.length + ')</span>';

    document.getElementById('mainContent').innerHTML = renderListTable(rows);

    // Update headcount
    var el = document.getElementById('headcount');
    var fte = 0, con = 0;
    rows.forEach(function(r) { if (r.type === 'Contractor') con++; else fte++; });
    var total = fte + con;
    if (!state.showEmp || !state.showCon) {
        el.textContent = total + ' shown (' + fte + ' FTE, ' + con + ' contractors)';
    } else {
        el.textContent = total + ' people (' + fte + ' FTE, ' + con + ' contractors)';
    }
}

function sortListBy(col) {
    if (state.listSortCol === col) {
        state.listSortAsc = !state.listSortAsc;
    } else {
        state.listSortCol = col;
        state.listSortAsc = true;
    }
    renderList();
}

// ── Render Org View ──
function render() {
    const org = DATA.orgs[state.currentOrg];
    const nodeId = state.currentNodeId;
    const node = org.nodes[nodeId];
    if (!node) return;

    // Breadcrumb
    const bc = document.getElementById('breadcrumb');
    let bcHtml = '';
    state.breadcrumb.forEach((id, i) => {
        const n = org.nodes[id];
        if (!n) return;
        if (i > 0) bcHtml += '<span>&rsaquo;</span> ';
        if (i === state.breadcrumb.length - 1) {
            bcHtml += '<span class="current">' + escHtml(displayName(n.name)) + '</span>';
        } else {
            bcHtml += '<a onclick="navigateTo(\'' + escHtml(id) + '\')">' + escHtml(displayName(n.name)) + '</a> ';
        }
    });
    bc.innerHTML = bcHtml;

    // Manager card
    const empBadge = isContractor(node.employment)
        ? '<span class="badge badge-contractor">Contractor</span>'
        : '<span class="badge badge-fte">FTE</span>';

    let teamPills = '';
    if (node.scrumTeams && node.scrumTeams.length) {
        teamPills = '<div class="teams">' + node.scrumTeams.map(t =>
            '<a class="team-pill" style="background:' + teamColor(t) + '22;color:' + teamColor(t) + '" onclick="showScrumView(\'' + escHtml(t).replace(/'/g,"\\\\'") + '\')">#' + escHtml(t) + '</a>'
        ).join('') + '</div>';
    }

    let html = '<div class="manager-section">';
    html += '<div class="manager-card">';
    html += '<div class="name">' + escHtml(displayName(node.name)) + '</div>';
    if (node.title) html += '<div class="title">' + escHtml(node.title) + '</div>';
    html += empBadge;
    html += teamPills;
    html += '</div>';

    // Get children
    let children = (org.children[nodeId] || []).map(id => org.nodes[id]).filter(Boolean);
    children = children.filter(function(c) {
        var con = isContractor(c.employment);
        return con ? state.showCon : state.showEmp;
    });

    if (children.length > 0) {
        html += '<div class="connector"></div>';
        html += '</div>';

        html += '<div class="reports-label">' + children.length + ' direct report' + (children.length > 1 ? 's' : '') + '</div>';
        html += '<div class="reports-grid">';

        children.forEach(child => {
            const cBadge = isContractor(child.employment)
                ? '<span class="badge badge-contractor">Contractor</span>'
                : '<span class="badge badge-fte">FTE</span>';

            let cTeams = '';
            if (child.scrumTeams && child.scrumTeams.length) {
                cTeams = '<div class="teams">' + child.scrumTeams.map(t =>
                    '<a class="team-pill" style="background:' + teamColor(t) + '22;color:' + teamColor(t) + '" onclick="event.stopPropagation();showScrumView(\'' + escHtml(t).replace(/'/g,"\\\\'") + '\')">#' + escHtml(t) + '</a>'
                ).join('') + '</div>';
            }

            // Count all descendants (respecting FTE filter)
            let drCount = countReports(org, child.id);
            let drLabel = drCount > 0 ? '<div class="dr-count">' + drCount + ' report' + (drCount > 1 ? 's' : '') + '</div>' : '';

            const phClass = child.placeholder ? ' placeholder' : '';
            html += '<div class="person-card' + phClass + '" onclick="navigateTo(\'' + escHtml(child.id) + '\')">';
            html += '<div class="name">' + escHtml(displayName(child.name)) + '</div>';
            if (child.title) html += '<div class="title">' + escHtml(child.title) + '</div>';
            html += cBadge;
            if (child.dottedLine) html += '<div class="dr-count" style="color:#805ad5;font-style:italic">Dotted-line: ' + escHtml(displayName(child.dottedLine)) + '</div>';
            html += cTeams;
            html += drLabel;
            html += '</div>';
        });

        html += '</div>';
    } else {
        html += '</div>';
        html += '<div class="empty-state">No direct reports</div>';
    }

    // Missing titles section
    const missing = DATA.missing[state.currentOrg] || [];
    if (missing.length > 0) {
        html += '<div class="missing-section">';
        html += '<div class="missing-toggle" onclick="this.nextElementSibling.classList.toggle(\'open\')">';
        html += '&#9888; ' + missing.length + ' people without titles in ' + escHtml(state.currentOrg) + ' (click to expand)';
        html += '</div>';
        html += '<div class="missing-list"><ul>';
        missing.forEach(name => {
            html += '<li>' + escHtml(displayName(name)) + '</li>';
        });
        html += '</ul></div></div>';
    }

    document.getElementById('mainContent').innerHTML = html;
    updateHeadcount();
}

function countReports(org, nodeId) {
    let children = org.children[nodeId] || [];
    children = children.filter(function(id) {
        var n = org.nodes[id];
        if (!n) return false;
        var con = isContractor(n.employment);
        return con ? state.showCon : state.showEmp;
    });
    let count = children.length;
    children.forEach(cid => {
        count += countReports(org, cid);
    });
    return count;
}

function updateHeadcount() {
    const org = DATA.orgs[state.currentOrg];
    let total = 0, fte = 0, contractors = 0;
    for (const [id, node] of Object.entries(org.nodes)) {
        if (node.placeholder) continue;
        total++;
        if (isContractor(node.employment)) contractors++;
        else fte++;
    }
    const shown = (state.showEmp ? fte : 0) + (state.showCon ? contractors : 0);
    const el = document.getElementById('headcount');
    if (!state.showEmp || !state.showCon) {
        el.textContent = shown + ' shown (of ' + total + ' total: ' + fte + ' FTE, ' + contractors + ' contractors)';
    } else {
        el.textContent = total + ' people (' + fte + ' FTE, ' + contractors + ' contractors)';
    }
}

if (!document.title.includes('Redacted')) {
    document.getElementById('redactGroup').style.display = '';
}
init();
</script>
</body>
</html>'''


# ─── Main ────────────────────────────────────────────────────────────────────

def main():
    print("=" * 60)
    print("Org Chart HTML Generator (on24.xlsx hierarchy)")
    print("=" * 60)

    # Step 1: Parse on24.xlsx — definitive FTE hierarchy
    print("\n[1] Parsing on24.xlsx (definitive hierarchy)...")
    on24_people = parse_on24(ON24_FILE)

    # Step 2: Parse talent snapshot — supplementary titles
    print("\n[2] Parsing talent snapshot...")
    title_map = parse_talent_snapshot(TALENT_FILE)

    # Step 3: Parse per-org tabs — employment + teams (no hierarchy)
    print("\n[3] Parsing per-org tabs (employment + scrum teams)...")
    org_tab_people = parse_org_roster(ORG_FILE)

    # Step 4: Parse SCRUMS.xlsx Teams Hierarchy — lead assignments
    print("\n[4] Parsing Teams Hierarchy...")
    teams_hier = parse_teams_hierarchy(SCRUMS_FILE)

    # Step 5: Build unified org datasets from on24 hierarchy + per-org enrichment
    print("\n[5] Building org datasets from on24 hierarchy...")
    org_datasets = build_from_on24(on24_people, org_tab_people, teams_hier, title_map)

    # Step 6: Build scrum index with lead identification
    print("\n[6] Building scrum team index...")
    scrum_teams = build_scrum_index(org_datasets, teams_hier)
    print(f"  {len(scrum_teams)} scrum teams found")
    for team_name in sorted(scrum_teams.keys()):
        groups = scrum_teams[team_name]
        total = sum(len(g) for g in groups.values())
        print(f"    {team_name}: {total} members")

    # Step 7: Serialize + generate HTML (named + redacted)
    print("\n[7] Generating HTML files...")
    data = make_serializable(org_datasets, scrum_teams)

    # Named version
    html_full = generate_html(data, redacted=False)
    OUTPUT_FILE.write_text(html_full, encoding='utf-8')
    print(f"  Written: {OUTPUT_FILE} ({len(html_full):,} bytes)")

    # Redacted version
    all_names = set()
    for tab_name, dataset in org_datasets.items():
        for nid, node in dataset["nodes"].items():
            if not node.get("placeholder"):
                all_names.add(node["name"])

    redacted_data = redact_data(data, all_names)
    html_redacted = generate_html(redacted_data, redacted=True)

    # Verify redaction
    leaked = verify_redaction(html_redacted, all_names)
    if leaked:
        print(f"\n  [WARN] Redaction verification: {len(leaked)} names may have leaked:")
        for name in leaked[:10]:
            print(f"    - {name}")
    else:
        print("  Redaction verification: PASSED (no names found in HTML)")

    REDACTED_FILE.write_text(html_redacted, encoding='utf-8')
    print(f"  Written: {REDACTED_FILE} ({len(html_redacted):,} bytes)")

    # Summary
    print("\n" + "=" * 60)
    print("Summary")
    print("=" * 60)
    for org_name in ["Product-Design", "Full QA Org", "Full Dev Org", SALESFORCE_ORG_NAME, "TPM"]:
        if org_name in org_datasets:
            ds = org_datasets[org_name]
            total = len(ds["nodes"])
            print(f"  {org_name}: {total} nodes")
    print(f"\n  Scrum teams: {len(scrum_teams)}")
    print(f"\n  Output files:")
    print(f"    {OUTPUT_FILE}")
    print(f"    {REDACTED_FILE}")
    print("\nDone!")


if __name__ == "__main__":
    main()
