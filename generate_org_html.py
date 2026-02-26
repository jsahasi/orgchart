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
ORG_FILE = DATA_DIR / "JayeshSahasi_ON24 QA-Dev Org List.xlsx"
TALENT_FILE = DATA_DIR / "JayeshSahasi_EngProduct_Jayesh_Talent_Snapshot_Leader_Input_2026.01.26.xlsx"
OUTPUT_FILE = Path(__file__).parent / "org_drilldown.html"
REDACTED_FILE = Path(__file__).parent / "org_drilldown_redacted.html"

ORG_TABS = ["Product-Design", "Full QA Org", "Full Dev Org", "TPM"]
TALENT_TABS = ["Dev", "QA", "Salesforce", "Product Management", "Program Management"]

JAYESH_NAME = "Jayesh Sahasi"
JAYESH_TITLE = "EVP Product & CTO"
JAYESH_DRS = ["Jaimini", "Steve Sims", "Oleg", "Jagjit", "Mahesh"]

NICKNAME_MAP = {
    "steve": "stephen",
    "dan": "daniel",
    "jay": "jawynson",
    "mike": "michael",
    "ben": "benjamin",
    "raj": "raj",
}

# Known people who appear as "Reports To" but not as "Name" rows — manual overrides
KAMAL_REPORTS_TO = "jaimini"  # Kamal reports to Jaimini, not Jayesh

# Manual title overrides for people whose names differ between files
MANUAL_TITLE_OVERRIDES = {
    "jagjit singh": "Director, Program Management",
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

# Kamal dotted-line reporting
KAMAL_DOTTED_LINE_TO = "Jayesh Sahasi"

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

def parse_org_tab(ws, tab_name, title_map):
    """Parse a single org tab. Returns list of node dicts."""
    rows = list(ws.iter_rows(min_row=1, values_only=True))
    if not rows:
        return []

    header = [str(c).strip() if c else "" for c in rows[0]]
    header_lower = [h.lower() for h in header]

    # Find column indices
    def find_col(candidates):
        for c in candidates:
            for i, h in enumerate(header_lower):
                if h == c.lower():
                    return i
        return None

    reports_to_idx = find_col(["Reports To", "reports to"])
    name_idx = find_col(["Name", "name"])
    employment_idx = find_col(["Employment", "employment"])
    title_idx = find_col(["Title", "title"])

    # Team column: try Team, Teams, Teams.1 in priority
    team_idx = find_col(["Team", "team"])
    teams_idx = find_col(["Teams", "teams"])
    teams1_idx = find_col(["Teams.1", "teams.1"])

    if name_idx is None or reports_to_idx is None:
        print(f"  [WARN] Tab '{tab_name}' missing Name or Reports To column")
        return []

    nodes = []
    last_reports_to = None
    slug_counter = defaultdict(int)

    for row in rows[1:]:
        if len(row) <= max(name_idx, reports_to_idx):
            continue

        raw_name = str(row[name_idx]).strip() if row[name_idx] else ""
        raw_reports_to = str(row[reports_to_idx]).strip() if row[reports_to_idx] else ""

        # Forward-fill Reports To
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

        # Title from org tab
        org_title = ""
        if title_idx is not None and len(row) > title_idx and row[title_idx]:
            org_title = str(row[title_idx]).strip()
            if org_title.lower() in ('none', 'nan'):
                org_title = ""

        # Team(s)
        team_raw_parts = []
        for tidx in [team_idx, teams_idx, teams1_idx]:
            if tidx is not None and len(row) > tidx and row[tidx]:
                val = str(row[tidx]).strip()
                if val.lower() not in ('none', 'nan', ''):
                    team_raw_parts.append(val)
        team_raw = " / ".join(team_raw_parts) if team_raw_parts else ""
        scrum_teams = parse_scrum_teams(team_raw)

        # Generate ID
        base_slug = slugify(raw_name)
        slug_counter[base_slug] += 1
        node_id = base_slug if slug_counter[base_slug] == 1 else f"{base_slug}-{slug_counter[base_slug]}"

        nodes.append({
            "id": node_id,
            "name": raw_name,
            "title": org_title,
            "employment": employment,
            "teamRaw": team_raw,
            "scrumTeams": scrum_teams,
            "reportsToRaw": raw_reports_to,
            "managerId": None,
            "placeholder": False,
            "org": tab_name,
        })

    # Enrich titles from talent snapshot
    missing_titles = []
    default_title = DEFAULT_TITLES_BY_ORG.get(tab_name, "")
    for node in nodes:
        key = normalize_name(node["name"])
        # Check per-person overrides first (both maps)
        if key in MANUAL_TITLE_OVERRIDES:
            node["title"] = MANUAL_TITLE_OVERRIDES[key]
        elif key in MANUAL_TITLE_OVERRIDES_EXTRA:
            node["title"] = MANUAL_TITLE_OVERRIDES_EXTRA[key]
        elif key in title_map:
            node["title"] = title_map[key]
        elif not node["title"]:
            # Try fuzzy: last name match + first name prefix
            matched = fuzzy_title_match(node["name"], title_map)
            if matched:
                node["title"] = matched
            elif default_title:
                node["title"] = default_title
            else:
                missing_titles.append(node["name"])

    return nodes, missing_titles


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


def parse_org_roster(filepath, title_map):
    """Parse all org tabs. Returns dict of tab_name -> (nodes_list, missing_titles)."""
    wb = openpyxl.load_workbook(filepath, read_only=True, data_only=True)
    org_data = {}

    for tab_name in ORG_TABS:
        if tab_name not in wb.sheetnames:
            print(f"  [WARN] Org tab '{tab_name}' not found")
            continue
        ws = wb[tab_name]
        result = parse_org_tab(ws, tab_name, title_map)
        if result:
            nodes, missing = result
            org_data[tab_name] = (nodes, missing)
            print(f"  {tab_name}: {len(nodes)} people, {len(missing)} missing titles")

    wb.close()
    return org_data


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


def build_org_dataset(tab_name, nodes, title_map):
    """Build a full org dataset with Jayesh at top."""

    # Create name->node index
    by_norm = {}
    by_id = {}
    for node in nodes:
        norm = normalize_name(node["name"])
        by_norm[norm] = node
        by_id[node["id"]] = node

    # Create/ensure Jayesh node
    jayesh_norm = normalize_name(JAYESH_NAME)
    if jayesh_norm in by_norm:
        jayesh = by_norm[jayesh_norm]
        jayesh["title"] = JAYESH_TITLE
        if not jayesh["employment"]:
            jayesh["employment"] = "Full Time"
        jayesh["managerId"] = None
    else:
        jayesh_id = "jayesh-sahasi"
        jayesh = {
            "id": jayesh_id,
            "name": JAYESH_NAME,
            "title": JAYESH_TITLE,
            "employment": "Full Time",
            "teamRaw": "",
            "scrumTeams": [],
            "reportsToRaw": "",
            "managerId": None,
            "placeholder": False,
            "org": tab_name,
        }
        nodes.insert(0, jayesh)
        by_norm[jayesh_norm] = jayesh
        by_id[jayesh["id"]] = jayesh

    jayesh_id = jayesh["id"]

    # Resolve all manager references first, before DR overrides
    placeholder_counter = [0]
    placeholders = {}

    for node in nodes:
        if node["id"] == jayesh_id:
            continue

        rt = node.get("reportsToRaw", "")
        if not rt:
            continue

        # Try to find manager
        manager = resolve_name_match(rt, by_norm)
        if manager:
            node["managerId"] = manager["id"]
        else:
            # Create placeholder
            ph_norm = normalize_name(rt)
            if ph_norm in placeholders:
                node["managerId"] = placeholders[ph_norm]["id"]
            else:
                placeholder_counter[0] += 1
                ph_id = f"placeholder-{slugify(rt)}"
                if ph_id in by_id:
                    ph_id = f"{ph_id}-{placeholder_counter[0]}"
                ph_node = {
                    "id": ph_id,
                    "name": rt,
                    "title": "",
                    "employment": "",
                    "teamRaw": "",
                    "scrumTeams": [],
                    "reportsToRaw": "",
                    "managerId": None,
                    "placeholder": True,
                    "org": tab_name,
                }
                # Enrich placeholder title: manual overrides > talent snapshot > fuzzy
                ph_key = normalize_name(rt)
                if ph_key in MANUAL_TITLE_OVERRIDES:
                    ph_node["title"] = MANUAL_TITLE_OVERRIDES[ph_key]
                elif ph_key in title_map:
                    ph_node["title"] = title_map[ph_key]
                else:
                    t = fuzzy_title_match(rt, title_map)
                    if t:
                        ph_node["title"] = t

                placeholders[ph_norm] = ph_node
                nodes.append(ph_node)
                by_norm[ph_norm] = ph_node
                by_id[ph_id] = ph_node
                node["managerId"] = ph_id

    # Now resolve Jayesh's DRs — force their managerId to Jayesh
    dr_resolved = {}
    for dr_hint in JAYESH_DRS:
        match = resolve_name_match(dr_hint, by_norm)
        if match and match["id"] != jayesh_id:
            dr_resolved[match["id"]] = dr_hint
            match["managerId"] = jayesh_id

    # Handle placeholder merging: if a placeholder IS a Jayesh DR, merge
    placeholders_to_remove = set()
    for ph_norm, ph_node in placeholders.items():
        if ph_node["id"] in dr_resolved:
            # This placeholder is already resolved as a DR — keep it
            continue
        if not ph_node["managerId"]:
            # Check if this placeholder matches a real Jayesh DR node
            dr_nodes = {normalize_name(n["name"]): n for n in nodes
                        if n["id"] in dr_resolved and not n.get("placeholder")}
            match = resolve_name_match(ph_node["name"], dr_nodes)
            if match:
                # Reroute children of this placeholder to the real DR
                for node in nodes:
                    if node["managerId"] == ph_node["id"]:
                        node["managerId"] = match["id"]
                placeholders_to_remove.add(ph_node["id"])
                continue

            # Special case: Kamal reports to Jaimini, not Jayesh (dotted-line to Jayesh)
            if normalize_name(ph_node["name"]).startswith("kamal"):
                jaimini = resolve_name_match("Jaimini", by_norm)
                if jaimini:
                    ph_node["managerId"] = jaimini["id"]
                    ph_node["dottedLine"] = KAMAL_DOTTED_LINE_TO
                    continue

            # Otherwise assign under Jayesh
            ph_node["managerId"] = jayesh_id

    # Remove merged placeholders from nodes list
    nodes = [n for n in nodes if n["id"] not in placeholders_to_remove]

    # Handle orphans: any node without managerId (except Jayesh)
    for node in nodes:
        if node["managerId"] is None and node["id"] != jayesh_id:
            # Try matching to a Jayesh DR
            matched_dr = False
            for dr_hint in JAYESH_DRS:
                if normalize_name(dr_hint) in normalize_name(node["name"]):
                    node["managerId"] = jayesh_id
                    matched_dr = True
                    break
            if not matched_dr:
                node["managerId"] = jayesh_id

    # QA Org fix: reroute automation contractors from Oleg to Ashish
    if tab_name == "Full QA Org":
        # Find Oleg and Ashish
        oleg_node = resolve_name_match("Oleg", by_norm)
        ashish_node = resolve_name_match("Ashish", by_norm)
        if oleg_node and ashish_node:
            oleg_id = oleg_node["id"]
            ashish_id = ashish_node["id"]
            # Ashish reports to Oleg
            ashish_node["managerId"] = oleg_id
            # Reroute: anyone reporting to Oleg who is NOT one of his real DRs
            # should report to Ashish instead
            for node in nodes:
                if node["managerId"] == oleg_id and node["id"] != oleg_id:
                    name_lower = normalize_name(node["name"])
                    first_name = name_lower.split()[0] if name_lower else ""
                    if first_name not in QA_OLEG_REAL_DRS:
                        node["managerId"] = ashish_id

    # Build children adjacency
    children = defaultdict(list)
    node_map = {}
    for node in nodes:
        node_map[node["id"]] = node
        if node["managerId"] and node["managerId"] in by_id:
            children[node["managerId"]].append(node["id"])
        elif node["managerId"] and node["id"] != jayesh_id:
            # Manager ID doesn't exist — reparent to Jayesh
            node["managerId"] = jayesh_id
            children[jayesh_id].append(node["id"])

    # Remove duplicates in children lists
    for pid in children:
        children[pid] = list(dict.fromkeys(children[pid]))

    # Sort children alphabetically by name
    for parent_id in children:
        children[parent_id].sort(key=lambda cid: node_map.get(cid, {}).get("name", "").lower())

    # Report Jayesh DRs found in this tab
    jayesh_children = children.get(jayesh_id, [])
    if jayesh_children:
        print(f"    Jayesh DRs in {tab_name}: {[node_map[c]['name'] for c in jayesh_children]}")

    return {
        "top": jayesh_id,
        "nodes": node_map,
        "children": dict(children),
    }


# ─── Step 4: Build Scrum Index ───────────────────────────────────────────────

def build_scrum_index(org_datasets):
    """Build global scrum index: team_name -> list of members."""
    # First pass: collect all teams with case-insensitive dedup
    # canonical_key (lowercase) -> preferred display name (first seen casing)
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
                    # Prefer title-cased version
                    canonical_names[key] = team
                raw_index[key].append({
                    "org": tab_name,
                    "id": node_id,
                    "name": node["name"],
                    "title": node.get("title", ""),
                    "employment": node.get("employment", ""),
                })

    # Merge similar team names (e.g., "p10console forums" and "console forums")
    # Build canonical scrum_index
    scrum_index = {}
    for key, members in raw_index.items():
        display = canonical_names[key]
        scrum_index[display] = members

    # Determine discipline and pick leads
    scrum_teams = {}
    for team_name, members in scrum_index.items():
        grouped = {"Dev": [], "QA": [], "Product": [], "TPM": [], "Other": []}
        for m in members:
            org = m["org"]
            if "Dev" in org:
                grouped["Dev"].append(m)
            elif "QA" in org:
                grouped["QA"].append(m)
            elif "Product" in org or "Design" in org:
                grouped["Product"].append(m)
            elif "TPM" in org:
                grouped["TPM"].append(m)
            else:
                grouped["Other"].append(m)

        # Pick lead per discipline (highest seniority first)
        for discipline in grouped:
            group = grouped[discipline]
            if group:
                group.sort(key=lambda m: (-title_seniority_score(m["title"]), m["name"].lower()))

        scrum_teams[team_name] = grouped

    return scrum_teams


# ─── Step 5: HTML Generation ─────────────────────────────────────────────────

def make_serializable(org_datasets, scrum_teams, missing_titles_map):
    """Convert data to JSON-serializable format."""
    orgs = {}
    # Build Home view: Jayesh + his 5 DRs across all orgs
    home_drs = []  # [{name, title, employment, org, nodeId}]

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

        # Collect Jayesh's DRs for the Home view
        for cid in top_children:
            c = nodes_ser.get(cid)
            if c:
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
        "missing": missing_titles_map,
        "homeDrs": home_drs,
    }


def redact_data(data, all_names):
    """Deep copy and replace all names with Person NNN."""
    data = copy.deepcopy(data)

    # Build consistent name->redacted mapping
    name_to_redacted = {}
    counter = [0]

    def get_redacted(name):
        norm = normalize_name(name)
        if norm not in name_to_redacted:
            counter[0] += 1
            name_to_redacted[norm] = f"Person {counter[0]:03d}"
        return name_to_redacted[norm]

    # Redact org nodes
    for tab_name, org in data["orgs"].items():
        for nid, node in org["nodes"].items():
            node["name"] = get_redacted(node["name"])
            if node.get("dottedLine"):
                node["dottedLine"] = get_redacted(node["dottedLine"])

    # Redact scrum members
    for team_name, groups in data["scrum"].items():
        for discipline, members in groups.items():
            for m in members:
                m["name"] = get_redacted(m["name"])

    # Redact missing titles lists
    for tab_name, names in data["missing"].items():
        data["missing"][tab_name] = [get_redacted(n) for n in names]

    # Redact homeDrs
    for dr in data.get("homeDrs", []):
        dr["name"] = get_redacted(dr["name"])

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

.toggle-btn {
    padding: 8px 16px;
    border: 1px solid rgba(255,255,255,0.3);
    border-radius: 8px;
    background: transparent;
    color: white;
    cursor: pointer;
    font-size: 13px;
    transition: all 0.2s;
    white-space: nowrap;
}
.toggle-btn:hover {
    background: rgba(255,255,255,0.1);
}
.toggle-btn.active {
    background: #e94560;
    border-color: #e94560;
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

/* Headcount bar */
.headcount {
    text-align: center;
    font-size: 12px;
    color: rgba(255,255,255,0.7);
    white-space: nowrap;
}
</style>
</head>
<body>

<div class="header">
    <h1>Org Chart__TITLE_SUFFIX__</h1>
    <div class="header-controls">
        <select id="orgSelect" onchange="switchOrg(this.value)"></select>
        <input type="text" id="searchBox" placeholder="Search by name..." onkeydown="if(event.key==='Enter')doSearch()">
        <button class="toggle-btn" id="fteToggle" onclick="toggleFTE()">FTE Only</button>
        <span class="headcount" id="headcount"></span>
    </div>
</div>

<div class="nav-bar">
    <button class="nav-btn" id="homeBtn" onclick="goHome()">Home</button>
    <button class="nav-btn" id="topBtn" onclick="goTop()">Top</button>
    <button class="nav-btn" id="upBtn" onclick="goUp()">Up</button>
    <button class="nav-btn" id="backToOrg" onclick="backToOrg()" style="display:none">Back to Org</button>
    <div class="breadcrumb" id="breadcrumb"></div>
</div>

<div class="main" id="mainContent"></div>

<script>
const DATA = __DATA_JSON__;

let state = {
    currentOrg: null,
    currentNodeId: null,
    breadcrumb: [],
    fteOnly: false,
    scrumView: null,
    lastOrgNodeId: null,
    isHome: true,
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
    state.isHome = false;
    navigateTo(DATA.orgs[orgName].top);
}

function goHome() {
    state.isHome = true;
    state.scrumView = null;
    state.currentOrg = null;
    document.getElementById('orgSelect').value = '__HOME__';
    document.getElementById('backToOrg').style.display = 'none';
    renderHome();
}

function renderHome() {
    // Get Jayesh's info from the first org dataset
    const firstOrg = Object.values(DATA.orgs)[0];
    const jayesh = firstOrg ? firstOrg.nodes[firstOrg.top] : null;
    const jayeshName = jayesh ? jayesh.name : 'Leader';
    const jayeshTitle = jayesh ? jayesh.title : '';

    const bc = document.getElementById('breadcrumb');
    bc.innerHTML = '<span class="current">' + escHtml(jayeshName) + (jayeshTitle ? ' — ' + escHtml(jayeshTitle) : '') + '</span>';

    let html = '<div class="manager-section">';
    html += '<div class="manager-card">';
    html += '<div class="name">' + escHtml(jayeshName) + '</div>';
    if (jayeshTitle) html += '<div class="title">' + escHtml(jayeshTitle) + '</div>';
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
        html += '<div class="name">' + escHtml(dr.name) + '</div>';
        if (dr.title) html += '<div class="title">' + escHtml(dr.title) + '</div>';
        html += badge;
        html += '<div class="dr-count" style="color:#3182ce">' + escHtml(dr.org) + '</div>';
        html += '</div>';
    });
    html += '</div>';

    document.getElementById('mainContent').innerHTML = html;
    // Show total headcount across all orgs
    let total = 0, fte = 0, contractors = 0;
    for (const [orgName, org] of Object.entries(DATA.orgs)) {
        for (const [id, node] of Object.entries(org.nodes)) {
            if (node.placeholder) continue;
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

function toggleFTE() {
    state.fteOnly = !state.fteOnly;
    const btn = document.getElementById('fteToggle');
    btn.classList.toggle('active', state.fteOnly);
    if (state.scrumView) {
        showScrumView(state.scrumView);
    } else {
        render();
    }
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
        if (state.fteOnly) {
            members = members.filter(m => !isContractor(m.employment));
        }
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
            html += '<a onclick="navigateToOrgCard(\'' + escHtml(m.org) + "','" + escHtml(m.id) + '\')">' + escHtml(m.name) + leadLabel + '</a>';
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
            bcHtml += '<span class="current">' + escHtml(n.name) + '</span>';
        } else {
            bcHtml += '<a onclick="navigateTo(\'' + escHtml(id) + '\')">' + escHtml(n.name) + '</a> ';
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
    html += '<div class="name">' + escHtml(node.name) + '</div>';
    if (node.title) html += '<div class="title">' + escHtml(node.title) + '</div>';
    html += empBadge;
    html += teamPills;
    html += '</div>';

    // Get children
    let children = (org.children[nodeId] || []).map(id => org.nodes[id]).filter(Boolean);
    if (state.fteOnly) {
        children = children.filter(c => !isContractor(c.employment));
    }

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
            html += '<div class="name">' + escHtml(child.name) + '</div>';
            if (child.title) html += '<div class="title">' + escHtml(child.title) + '</div>';
            html += cBadge;
            if (child.dottedLine) html += '<div class="dr-count" style="color:#805ad5;font-style:italic">Dotted-line: ' + escHtml(child.dottedLine) + '</div>';
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
            html += '<li>' + escHtml(name) + '</li>';
        });
        html += '</ul></div></div>';
    }

    document.getElementById('mainContent').innerHTML = html;
    updateHeadcount();
}

function countReports(org, nodeId) {
    let children = org.children[nodeId] || [];
    if (state.fteOnly) {
        children = children.filter(id => {
            const n = org.nodes[id];
            return n && !isContractor(n.employment);
        });
    }
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
    const el = document.getElementById('headcount');
    if (state.fteOnly) {
        el.textContent = fte + ' FTE (of ' + total + ' total)';
    } else {
        el.textContent = total + ' people (' + fte + ' FTE, ' + contractors + ' contractors)';
    }
}

init();
</script>
</body>
</html>'''


# ─── Main ────────────────────────────────────────────────────────────────────

def main():
    print("=" * 60)
    print("Org Chart HTML Generator")
    print("=" * 60)

    # Step 1: Parse talent snapshot
    print("\n[1] Parsing talent snapshot...")
    title_map = parse_talent_snapshot(TALENT_FILE)

    # Step 2: Parse org roster
    print("\n[2] Parsing org roster...")
    org_raw = parse_org_roster(ORG_FILE, title_map)

    # Step 3: Build org datasets
    print("\n[3] Building org datasets...")
    org_datasets = {}
    missing_titles_map = {}
    for tab_name, (nodes, missing) in org_raw.items():
        print(f"  Building {tab_name}...")
        dataset = build_org_dataset(tab_name, nodes, title_map)
        org_datasets[tab_name] = dataset
        missing_titles_map[tab_name] = missing

        # Print DR info
        top = dataset["top"]
        children = dataset["children"].get(top, [])
        print(f"    Top: {dataset['nodes'][top]['name']}")
        print(f"    Direct reports: {len(children)}")
        for cid in children:
            c = dataset["nodes"][cid]
            print(f"      - {c['name']} ({c.get('title', 'no title')})")

    # Step 4: Build scrum index
    print("\n[4] Building scrum team index...")
    scrum_teams = build_scrum_index(org_datasets)
    print(f"  {len(scrum_teams)} scrum teams found")
    for team_name in sorted(scrum_teams.keys()):
        groups = scrum_teams[team_name]
        total = sum(len(g) for g in groups.values())
        print(f"    {team_name}: {total} members")

    # Step 5: Serialize
    print("\n[5] Generating HTML files...")
    data = make_serializable(org_datasets, scrum_teams, missing_titles_map)

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
    for tab_name in ORG_TABS:
        if tab_name in org_datasets:
            ds = org_datasets[tab_name]
            total = len(ds["nodes"])
            missing = len(missing_titles_map.get(tab_name, []))
            print(f"  {tab_name}: {total} nodes, {missing} missing titles")
    print(f"\n  Scrum teams: {len(scrum_teams)}")
    print(f"\n  Output files:")
    print(f"    {OUTPUT_FILE}")
    print(f"    {REDACTED_FILE}")
    print("\nDone!")


if __name__ == "__main__":
    main()
