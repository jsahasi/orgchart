"""
Shared utilities and HTML template for org chart generators.

Contains common functions used by both the main org chart generator
and any other tools that need to work with org chart data:
  - normalize_name / slugify / is_contractor — name/employment helpers
  - title_seniority_score — title ranking
  - redact_data / verify_redaction — redaction pipeline
  - generate_html — HTML output from data dict
  - _HTML_TEMPLATE — the full standalone HTML template
"""

import json
import re
import copy


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


# ─── Redaction ───────────────────────────────────────────────────────────────

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

    # Build sorted names list (longest first) for rationale scrubbing
    sorted_names = sorted(all_names, key=len, reverse=True)

    def redact_rationale(text):
        """Replace any real names found in rationale text with redacted versions."""
        if not text:
            return text
        result = text
        for real_name in sorted_names:
            if len(real_name) < 4:
                continue
            # Case-insensitive replacement of full names
            pattern = re.compile(re.escape(real_name), re.IGNORECASE)
            if pattern.search(result):
                result = pattern.sub(get_redacted(real_name), result)
            # Also try first names only (>= 4 chars) for partial matches
            parts = real_name.split()
            for part in parts:
                if len(part) >= 4:
                    part_pattern = re.compile(r'\b' + re.escape(part) + r'\b', re.IGNORECASE)
                    if part_pattern.search(result):
                        result = part_pattern.sub(get_redacted(part), result)
        return result

    # Redact org nodes: names, dottedLine, rationale, email, and IDs
    for tab_name, org in data["orgs"].items():
        # Remap top
        org["top"] = get_anon_id(org["top"])

        # Remap nodes dict
        new_nodes = {}
        for nid, node in org["nodes"].items():
            node["name"] = get_redacted(node["name"])
            if node.get("dottedLine"):
                node["dottedLine"] = get_redacted(node["dottedLine"])
            if node.get("rationale"):
                node["rationale"] = redact_rationale(node["rationale"])
            if "email" in node:
                node["email"] = ""
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

    # Redact scrum members: names, IDs, rationale, email
    for team_name, groups in data["scrum"].items():
        for discipline, members in groups.items():
            for m in members:
                m["name"] = get_redacted(m["name"])
                if "id" in m:
                    m["id"] = get_anon_id(m["id"])
                if m.get("rationale"):
                    m["rationale"] = redact_rationale(m["rationale"])
                if "email" in m:
                    m["email"] = ""

    # Redact missing titles lists
    for tab_name, names in data["missing"].items():
        data["missing"][tab_name] = [get_redacted(n) for n in names]

    # Redact homeDrs: names, nodeIds, rationale, email
    for dr in data.get("homeDrs", []):
        dr["name"] = get_redacted(dr["name"])
        if "nodeId" in dr:
            dr["nodeId"] = get_anon_id(dr["nodeId"])
        if dr.get("rationale"):
            dr["rationale"] = redact_rationale(dr["rationale"])
        if "email" in dr:
            dr["email"] = ""

    # Redact scrumMeta: SM/PO names (may contain "Name1 / Name2")
    for team_name, meta in data.get("scrumMeta", {}).items():
        for key in ("scrumMaster", "productOwner"):
            if meta.get(key):
                meta[key] = " / ".join(
                    get_redacted(n.strip()) for n in meta[key].split("/")
                )

    return data


RATING_FIELDS = ("talentBand", "talentCategory", "rationale", "stackRank")


def strip_ratings(data):
    """Deep copy and clear per-person rating/talent fields.

    Produces an unredacted dataset (real names intact) with the individual
    talent/performance data removed: talentBand, talentCategory, rationale,
    stackRank on every org node, scrum member, and homeDr entry.
    """
    data = copy.deepcopy(data)

    def clear(obj):
        for field in RATING_FIELDS:
            if field in obj:
                obj[field] = ""

    for org in data.get("orgs", {}).values():
        for node in org.get("nodes", {}).values():
            clear(node)

    for groups in data.get("scrum", {}).values():
        for members in groups.values():
            for m in members:
                clear(m)

    for dr in data.get("homeDrs", []):
        clear(dr)

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


# ─── HTML Generation ─────────────────────────────────────────────────────────

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
<link rel="preconnect" href="https://fonts.googleapis.com">
<link href="https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700&display=swap" rel="stylesheet">
<style>
* { margin: 0; padding: 0; box-sizing: border-box; }

body {
    font-family: 'Inter', -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, 'Helvetica Neue', Arial, sans-serif;
    background: #f8fafc;
    color: #0f172a;
    min-height: 100vh;
    -webkit-font-smoothing: antialiased;
}
:root {
    --org-grad-start: #0f172a;
    --org-grad-end: #1e293b;
    --org-accent: #3b82f6;
    --org-tint: #eff6ff;
    --org-bar-end: #8b5cf6;
}
.skip-link {
    position: absolute;
    left: -9999px;
    top: 0;
    z-index: 1000;
    padding: 8px 16px;
    background: #fff;
    color: #0f172a;
    font-weight: 600;
    font-size: 14px;
    text-decoration: none;
    border-radius: 0 0 8px 0;
    box-shadow: 0 2px 8px rgba(0,0,0,0.1);
}
.skip-link:focus {
    left: 0;
}

/* Header */
.header {
    background: linear-gradient(135deg, var(--org-grad-start) 0%, var(--org-grad-end) 100%);
    color: white;
    padding: 16px 24px;
    display: flex;
    align-items: center;
    gap: 16px;
    flex-wrap: wrap;
    box-shadow: 0 1px 3px rgba(0,0,0,0.1), 0 4px 20px rgba(0,0,0,0.08);
    backdrop-filter: blur(8px);
    border-bottom: 1px solid rgba(255,255,255,0.05);
    position: sticky;
    top: 0;
    z-index: 100;
    transition: background 0.4s ease;
}

.header h1 {
    font-size: 18px;
    font-weight: 600;
    white-space: nowrap;
    letter-spacing: -0.3px;
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
    border: 1px solid rgba(255,255,255,0.4);
    border-radius: 8px;
    background: rgba(255,255,255,0.1);
    color: white;
    font-size: 14px;
    outline: none;
    transition: border-color 0.2s;
}
select:focus, input[type="text"]:focus {
    box-shadow: 0 0 0 3px rgba(59, 130, 246, 0.5);
    outline: 2px solid #3b82f6;
    outline-offset: 1px;
    border-color: rgba(59, 130, 246, 0.6);
    background: rgba(255,255,255,0.12);
}
select option {
    background: #1a1a2e;
    color: white;
}
input[type="text"]::placeholder {
    color: rgba(255,255,255,0.7);
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
    background: rgba(255,255,255,0.4);
    border: 1px solid rgba(255,255,255,0.6);
    border-radius: 12px;
    cursor: pointer;
    transition: background 0.3s cubic-bezier(0.4, 0, 0.2, 1);
    flex-shrink: 0;
}
.switch.active {
    background: #10b981;
}
.switch-knob {
    position: absolute;
    top: 2px;
    left: 2px;
    width: 20px;
    height: 20px;
    background: #fff;
    border-radius: 50%;
    transition: transform 0.3s cubic-bezier(0.4, 0, 0.2, 1);
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
    color: #475569;
    cursor: pointer;
    font-size: 13px;
    font-weight: 500;
    transition: all 0.2s;
}
.nav-btn:hover {
    background: #f1f5f9;
    border-color: #94a3b8;
    color: #0f172a;
}
.nav-btn:active {
    transform: scale(0.97);
}
.nav-btn:focus-visible,
.person-card:focus-visible,
.team-pill:focus-visible,
.all-scrum-card:focus-visible,
.zoom-controls button:focus-visible,
.switch:focus-visible {
    outline: 2px solid #3b82f6;
    outline-offset: 2px;
}

.breadcrumb {
    display: flex;
    align-items: center;
    gap: 4px;
    flex-wrap: wrap;
    flex: 1;
}
.breadcrumb span {
    color: #64748b;
    font-size: 13px;
}
.breadcrumb a {
    color: var(--org-accent);
    text-decoration: none;
    font-size: 13px;
    cursor: pointer;
}
.breadcrumb a:hover {
    color: var(--org-accent);
    filter: brightness(0.85);
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
    background: linear-gradient(135deg, var(--org-grad-start) 0%, var(--org-grad-end) 100%);
    color: white;
    border-radius: 16px;
    padding: 24px 32px;
    text-align: center;
    box-shadow: 0 12px 36px rgba(15, 23, 42, 0.25);
    min-width: 280px;
    max-width: 400px;
    position: relative;
    transition: background 0.4s ease;
}
.manager-card::after {
    content: '';
    position: absolute;
    top: 0;
    right: 0;
    width: 100%;
    height: 100%;
    background: radial-gradient(circle at top right, rgba(255,255,255,0.06) 0%, transparent 70%);
    pointer-events: none;
    border-radius: 16px;
}
.manager-card .name {
    font-size: 22px;
    letter-spacing: -0.3px;
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
    background: rgba(16, 185, 129, 0.15);
    color: #059669;
}
.badge-contractor {
    background: rgba(245, 158, 11, 0.15);
    color: #d97706;
}
.manager-card .badge-fte {
    background: rgba(16, 185, 129, 0.25);
    color: #34d399;
}
.manager-card .badge-contractor {
    background: rgba(245, 158, 11, 0.25);
    color: #fbbf24;
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
    color: #4a5568;
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
    box-shadow: 0 1px 3px rgba(0,0,0,0.06), 0 1px 2px rgba(0,0,0,0.04);
    border: 1px solid #e2e8f0;
    cursor: pointer;
    transition: all 0.2s cubic-bezier(0.4, 0, 0.2, 1);
    position: relative;
}
.person-card::before {
    content: '';
    position: absolute;
    top: 0;
    left: 0;
    right: 0;
    height: 3px;
    background: linear-gradient(90deg, var(--org-accent), var(--org-bar-end));
    opacity: 0;
    transition: opacity 0.2s;
}
.person-card:hover::before {
    opacity: 1;
}
.person-card:hover {
    transform: translateY(-3px);
    box-shadow: 0 12px 28px rgba(0,0,0,0.08), 0 4px 10px rgba(0,0,0,0.04);
    border-color: var(--org-accent);
}
.person-card .name {
    font-size: 15px;
    font-weight: 600;
    color: #0f172a;
    letter-spacing: -0.2px;
    margin-bottom: 4px;
}
.person-card .title {
    font-size: 12px;
    color: #64748b;
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
    color: #64748b;
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
    padding: 3px 10px;
    border-radius: 20px;
    font-size: 11px;
    font-weight: 600;
    cursor: pointer;
    transition: opacity 0.2s, transform 0.2s;
    text-decoration: none;
    letter-spacing: 0.2px;
}
.team-pill:hover {
    opacity: 0.8;
    transform: translateY(-1px);
}

/* Scrum view */
.scrum-view {
    max-width: 900px;
    margin: 0 auto;
}
.scrum-header {
    font-size: 24px;
    font-weight: 700;
    color: #0f172a;
    margin-bottom: 16px;
    letter-spacing: -0.5px;
}
.scrum-meta {
    display: flex;
    gap: 24px;
    flex-wrap: wrap;
    margin-bottom: 20px;
    padding: 12px 16px;
    background: var(--org-tint, #f8fafc);
    border-radius: 10px;
    border-left: 3px solid var(--org-accent, #3b82f6);
}
.scrum-meta-item { font-size: 13.5px; color: #334155; }
.scrum-meta-label {
    font-weight: 600;
    color: #64748b;
    font-size: 11px;
    text-transform: uppercase;
    letter-spacing: 0.5px;
    display: block;
    margin-bottom: 2px;
}
.all-scrum-card-meta {
    padding: 6px 16px 8px;
    font-size: 11.5px;
    color: #64748b;
    border-bottom: 1px solid #e8ecf1;
    display: flex;
    gap: 12px;
    flex-wrap: wrap;
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
    color: #64748b;
    text-transform: uppercase;
    letter-spacing: 1px;
    margin-bottom: 12px;
    padding-bottom: 8px;
    border-bottom: 2px solid var(--org-accent, #3b82f6);
    opacity: 0.85;
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
    text-decoration: none;
}
.scrum-member a:hover {
    text-decoration: underline;
    filter: brightness(0.85);
}
.scrum-member .member-title {
    color: #718096;
    font-size: 12px;
    font-weight: 400;
}

/* All-Scrum View */
.all-scrum-viewport {
    position: relative;
    overflow: auto;
    width: 100%;
    height: calc(100vh - 140px);
    background: #f7fafc;
}
.all-scrum-canvas {
    transform-origin: 0 0;
    padding: 24px;
    min-width: min-content;
}
.all-scrum-grid {
    display: flex;
    flex-wrap: wrap;
    gap: 20px;
    align-items: flex-start;
}
.all-scrum-card {
    background: white;
    border-radius: 12px;
    box-shadow: 0 2px 8px rgba(0,0,0,0.08);
    border: 1px solid #e2e8f0;
    width: 280px;
    cursor: pointer;
    transition: box-shadow 0.2s, border-color 0.2s;
    flex-shrink: 0;
}
.all-scrum-card:hover {
    box-shadow: 0 4px 16px rgba(0,0,0,0.14);
    border-color: var(--org-accent, #3b82f6);
}
.all-scrum-card-header {
    padding: 14px 16px;
    font-weight: 700;
    font-size: 15px;
    border-bottom: 1px solid #e8ecf1;
    display: flex;
    justify-content: space-between;
    align-items: center;
    color: #2d3748;
}
.all-scrum-card-body {
    padding: 10px 16px 14px;
}
.all-scrum-card-body .disc-label {
    font-size: 11px;
    font-weight: 600;
    color: #64748b;
    text-transform: uppercase;
    letter-spacing: 0.5px;
    margin-top: 8px;
    margin-bottom: 4px;
}
.all-scrum-card-body .disc-label:first-child {
    margin-top: 0;
}
.all-scrum-card-body .member-row {
    font-size: 13px;
    padding: 3px 0;
    color: #4a5568;
    display: flex;
    align-items: center;
    gap: 6px;
}
.all-scrum-card-body .member-row.is-lead {
    font-weight: 700;
}
.zoom-controls {
    position: fixed;
    bottom: 24px;
    right: 24px;
    display: flex;
    flex-direction: column;
    gap: 6px;
    z-index: 100;
}
.zoom-controls button {
    width: 44px;
    height: 44px;
    border-radius: 8px;
    border: 1px solid #e2e8f0;
    background: white;
    font-size: 18px;
    cursor: pointer;
    box-shadow: 0 2px 6px rgba(0,0,0,0.1);
    display: flex;
    align-items: center;
    justify-content: center;
    color: #4a5568;
}
.zoom-controls button:hover {
    background: #edf2f7;
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
    color: #4a5568;
}

/* Empty state */
.empty-state {
    text-align: center;
    color: #64748b;
    padding: 48px;
    font-size: 15px;
}

/* List view table */
.list-view { overflow-x: visible; background: #fff; border-radius: 12px; box-shadow: 0 2px 8px rgba(0,0,0,0.08); margin: 4px; }
.list-view table { width: 100%; border-collapse: collapse; font-size: 13.5px; }
.list-view th { cursor: pointer; user-select: none; }
.list-view th, .list-view td { text-align: left; padding: 11px 14px; border-bottom: 1px solid #eef1f6; }
.list-view th { background: linear-gradient(180deg, #f5f7fa 0%, #edf0f7 100%); color: #475569; font-weight: 600; font-size: 11.5px; text-transform: uppercase; letter-spacing: 0.6px; position: sticky; top: 0; z-index: 1; border-bottom: 2px solid #dde3ed; }
.list-view th .sort-arrow { font-size: 10px; margin-left: 4px; opacity: 0.3; }
.list-view th .sort-arrow.active { opacity: 1; color: #3b82f6; }
th button {
    background: none;
    border: none;
    font: inherit;
    color: inherit;
    cursor: pointer;
    padding: 0;
    text-transform: inherit;
    letter-spacing: inherit;
    font-weight: inherit;
    width: 100%;
    text-align: left;
}
th button:focus-visible {
    outline: 2px solid #3b82f6;
    outline-offset: 2px;
}
.list-view tbody tr:nth-child(even) { background: #f9fafb; }
.list-view tbody tr { transition: background 0.15s ease; }
.list-view tbody tr:hover { background: #eff6ff; }
.list-view td a { cursor: pointer; text-decoration: none; font-weight: 500; }
.list-view td a:hover { text-decoration: underline; filter: brightness(0.85); }
.list-view .team-pill { display: inline-block; padding: 2px 10px; border-radius: 12px; font-size: 11px; margin: 1px 3px; text-decoration: none; font-weight: 500; transition: opacity 0.15s ease; }
.list-view .team-pill:hover { opacity: 0.75; }
.list-view .badge { font-size: 11px; padding: 2px 10px; border-radius: 10px; font-weight: 600; }
.list-view .badge-fte { background: #dbeafe; color: #1e40af; }
.list-view .badge-contractor { background: #fef3c7; color: #92400e; }
.list-view .rank-high { background: #dcfce7; color: #166534; }
.list-view .rank-med { background: #fef9c3; color: #854d0e; }
.list-view .rank-low { background: #fee2e2; color: #991b1b; }

/* Headcount bar */
.headcount {
    text-align: center;
    font-size: 12px;
    color: rgba(255,255,255,0.85);
    white-space: nowrap;
}

/* Talent info tooltip */
.talent-info { display: inline-block; position: relative; cursor: pointer; margin-left: 6px; vertical-align: middle; }
.talent-info .info-icon { width: 24px; height: 24px; border-radius: 50%; background: #e2e8f0; color: #64748b; font-size: 12px; font-weight: 700; text-align: center; line-height: 24px; font-style: italic; display: inline-block; cursor: pointer; }
.talent-info .info-icon:hover { background: #3b82f6; color: #fff; }
.talent-info .talent-tip { display: none; position: absolute; bottom: calc(100% + 8px); left: 50%; transform: translateX(-50%); background: #0f172a; color: #f1f5f9; padding: 12px 16px; border-radius: 8px; font-size: 12px; line-height: 1.5; width: 300px; z-index: 100; box-shadow: 0 4px 12px rgba(0,0,0,0.15); border: 1px solid rgba(255,255,255,0.1); pointer-events: auto; }
.talent-info:hover .talent-tip, .talent-info:focus-within .talent-tip { display: block; }
.talent-tip .tip-label { color: #94a3b8; font-size: 10px; text-transform: uppercase; letter-spacing: 0.5px; }
.talent-tip .tip-value { margin-bottom: 8px; }
.talent-tip .tip-value:last-child { margin-bottom: 0; }
.talent-tip a { color: #7dd3fc; text-decoration: underline; }
.talent-tip a:hover { color: #bae6fd; }

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
    .person-card { animation-delay: 0s !important; }
}

@keyframes fadeSlideUp {
    from { opacity: 0; transform: translateY(12px); }
    to { opacity: 1; transform: translateY(0); }
}
.person-card {
    animation: fadeSlideUp 0.3s cubic-bezier(0.4, 0, 0.2, 1) both;
}
.reports-grid .person-card:nth-child(1) { animation-delay: 0.02s; }
.reports-grid .person-card:nth-child(2) { animation-delay: 0.04s; }
.reports-grid .person-card:nth-child(3) { animation-delay: 0.06s; }
.reports-grid .person-card:nth-child(4) { animation-delay: 0.08s; }
.reports-grid .person-card:nth-child(5) { animation-delay: 0.10s; }
.reports-grid .person-card:nth-child(6) { animation-delay: 0.12s; }
.reports-grid .person-card:nth-child(7) { animation-delay: 0.14s; }
.reports-grid .person-card:nth-child(8) { animation-delay: 0.16s; }
.reports-grid .person-card:nth-child(n+9) { animation-delay: 0.18s; }
@media (prefers-reduced-motion: reduce) {
    .person-card,
    .nav-btn,
    .team-pill,
    .switch,
    .switch-knob,
    .all-scrum-card,
    .info-icon {
        animation: none !important;
        transition: none !important;
    }
}
</style>
</head>
<body>
<a href="#mainContent" class="skip-link">Skip to main content</a>

<header class="header" role="banner">
    <h1>Org Chart__TITLE_SUFFIX__</h1>
    <div class="header-controls">
        <select id="orgSelect" aria-label="Select organization" onchange="switchOrg(this.value)"></select>
        <input type="text" id="searchBox" aria-label="Search by name" placeholder="Search by name..." onkeydown="if(event.key==='Enter')doSearch()">
        <div class="switch-group">
            <label class="switch-label">
                <div class="switch active" id="empToggle" role="switch" aria-checked="true" aria-label="Show employees" tabindex="0" onclick="toggleFilter('emp')" onkeydown="if(event.key===' '||event.key==='Enter'){event.preventDefault();toggleFilter('emp')}"><div class="switch-knob"></div></div>
                <span>Employees</span>
            </label>
            <label class="switch-label">
                <div class="switch active" id="conToggle" role="switch" aria-checked="true" aria-label="Show contractors" tabindex="0" onclick="toggleFilter('con')" onkeydown="if(event.key===' '||event.key==='Enter'){event.preventDefault();toggleFilter('con')}"><div class="switch-knob"></div></div>
                <span>Contractors</span>
            </label>
            <label class="switch-label" id="redactGroup" style="display:none">
                <div class="switch" id="redactToggle" role="switch" aria-checked="false" aria-label="Redact names" tabindex="0" onclick="toggleFilter('redact')" onkeydown="if(event.key===' '||event.key==='Enter'){event.preventDefault();toggleFilter('redact')}"><div class="switch-knob"></div></div>
                <span>Redact</span>
            </label>
        </div>
        <span class="headcount" id="headcount" aria-live="polite" aria-atomic="true"></span>
    </div>
</header>

<nav class="nav-bar" aria-label="Org chart navigation">
    <button class="nav-btn" id="homeBtn" onclick="goHome()">Home</button>
    <button class="nav-btn" id="topBtn" onclick="goTop()">Top</button>
    <button class="nav-btn" id="upBtn" onclick="goUp()">Up</button>
    <button class="nav-btn" id="backToOrg" onclick="backToOrg()" style="display:none">Back to Org</button>
    <button class="nav-btn" id="listBtn" onclick="showListView()">List</button>
    <button class="nav-btn" id="scrumAllBtn" onclick="showAllScrumView()">Scrum</button>
    <nav class="breadcrumb" id="breadcrumb" aria-label="Breadcrumb"></nav>
</nav>

<main class="main" id="mainContent" tabindex="-1"></main>

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
    allScrumView: false,
    listSortCol: 'name',
    listSortAsc: true,
};

// ── Color palette for team pills ──
const TEAM_COLORS = ['#3b82f6','#ef4444','#10b981','#f59e0b','#8b5cf6','#f97316','#14b8a6','#ec4899','#6366f1','#ea580c','#059669','#7c3aed'];

function teamColor(name) {
    let hash = 0;
    for (let i = 0; i < name.length; i++) hash = ((hash << 5) - hash) + name.charCodeAt(i);
    return TEAM_COLORS[Math.abs(hash) % TEAM_COLORS.length];
}

function teamPillTextColor(name) {
    var darkMap = {'#f59e0b':'#92400e','#f97316':'#9a3412','#ea580c':'#9a3412','#ec4899':'#9d174d','#10b981':'#065f46','#14b8a6':'#134e4a'};
    var base = teamColor(name);
    return darkMap[base] || base;
}

const ORG_THEMES = {
    '__HOME__':       {start:'#0f172a',end:'#1e293b',accent:'#3b82f6',tint:'#eff6ff',barEnd:'#8b5cf6'},
    'Product-Design': {start:'#7c2d12',end:'#c2410c',accent:'#f97316',tint:'#fff7ed',barEnd:'#fbbf24'},
    'Full QA Org':    {start:'#134e4a',end:'#0f766e',accent:'#14b8a6',tint:'#f0fdfa',barEnd:'#06b6d4'},
    'Full Dev Org':   {start:'#312e81',end:'#4338ca',accent:'#6366f1',tint:'#eef2ff',barEnd:'#a78bfa'},
    'Salesforce':     {start:'#14532d',end:'#15803d',accent:'#22c55e',tint:'#f0fdf4',barEnd:'#4ade80'},
    'TPM':            {start:'#881337',end:'#be123c',accent:'#f43f5e',tint:'#fff1f2',barEnd:'#fb923c'},
};

function applyOrgTheme(orgName) {
    var theme = ORG_THEMES[orgName] || ORG_THEMES['__HOME__'];
    var s = document.documentElement.style;
    s.setProperty('--org-grad-start', theme.start);
    s.setProperty('--org-grad-end', theme.end);
    s.setProperty('--org-accent', theme.accent);
    s.setProperty('--org-tint', theme.tint);
    s.setProperty('--org-bar-end', theme.barEnd);
}

function escHtml(s) {
    if (!s) return '';
    return s.replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;');
}

function talentTooltip(node) {
    if (!node.talentBand && !node.talentCategory && !node.rationale && !node.email) return '';
    var html = '<span class="talent-info" tabindex="0">';
    html += '<span class="info-icon" role="img" aria-label="Talent information">i</span>';
    html += '<span class="talent-tip">';
    if (node.talentBand) html += '<div class="tip-label">Band</div><div class="tip-value">' + escHtml(node.talentBand) + '</div>';
    if (node.talentCategory) html += '<div class="tip-label">Category</div><div class="tip-value">' + escHtml(node.talentCategory) + '</div>';
    if (node.rationale) html += '<div class="tip-label">Rationale</div><div class="tip-value">' + escHtml(node.rationale) + '</div>';
    if (node.email) html += '<div class="tip-label">Email</div><div class="tip-value"><a href="mailto:' + escHtml(node.email) + '" onclick="event.stopPropagation()">' + escHtml(node.email) + '</a></div>';
    html += '</span></span>';
    return html;
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
    applyOrgTheme('__HOME__');
    goHome();
}

function switchOrg(orgName) {
    if (orgName === '__HOME__') { goHome(); return; }
    state.currentOrg = orgName;
    applyOrgTheme(orgName);
    state.scrumView = null;
    state.listView = false;
    state.allScrumView = false;
    state.isHome = false;
    navigateTo(DATA.orgs[orgName].top);
}

function goHome() {
    state.isHome = true;
    applyOrgTheme('__HOME__');
    state.scrumView = null;
    state.listView = false;
    state.allScrumView = false;
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
    html += '<div class="name">' + escHtml(displayName(topName)) + (topPerson ? talentTooltip(topPerson) : '') + '</div>';
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
        var drTheme = ORG_THEMES[dr.org] || ORG_THEMES['__HOME__'];
        html += '<div class="person-card" style="border-left: 4px solid ' + drTheme.accent + ';" onclick="switchToOrgDr(\'' + escHtml(dr.org) + '\',\'' + escHtml(dr.nodeId) + '\')" role="button" tabindex="0" onkeydown="if(event.key===\'Enter\'||event.key===\' \'){event.preventDefault();switchToOrgDr(\'' + escHtml(dr.org) + '\',\'' + escHtml(dr.nodeId) + '\')}">';
        html += '<div class="name">' + escHtml(displayName(dr.name)) + talentTooltip(dr) + '</div>';
        if (dr.title) html += '<div class="title">' + escHtml(dr.title) + '</div>';
        html += badge;
        html += '<div class="dr-count" style="color:' + drTheme.accent + '">' + escHtml(dr.org) + '</div>';
        html += '</div>';
    });
    html += '</div>';

    document.getElementById('mainContent').innerHTML = html;
    document.getElementById('mainContent').focus();
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
    applyOrgTheme(orgName);
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
    document.getElementById('empToggle').setAttribute('aria-checked', state.showEmp);
    document.getElementById('conToggle').setAttribute('aria-checked', state.showCon);
    document.getElementById('redactToggle').setAttribute('aria-checked', state.redacted);
    if (state.allScrumView) showAllScrumView();
    else if (state.listView) renderList();
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
    applyOrgTheme(state.currentOrg || '__HOME__');
    state.allScrumView = false;
    state.listView = false;
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
        document.getElementById('mainContent').focus();
        return;
    }

    // Update breadcrumb
    const bc = document.getElementById('breadcrumb');
    bc.innerHTML = '<span class="current">#' + escHtml(teamName) + '</span>';

    let html = '<div class="scrum-view">';
    html += '<div class="scrum-header">#' + escHtml(teamName) + '</div>';

    var meta = DATA.scrumMeta && DATA.scrumMeta[teamName];
    if (meta && (meta.productOwner || meta.scrumMaster)) {
        html += '<div class="scrum-meta">';
        if (meta.productOwner) html += '<span class="scrum-meta-item"><span class="scrum-meta-label">Product Owner</span>' + escHtml(displayName(meta.productOwner)) + '</span>';
        if (meta.scrumMaster) html += '<span class="scrum-meta-item"><span class="scrum-meta-label">Scrum Master</span>' + escHtml(displayName(meta.scrumMaster)) + '</span>';
        html += '</div>';
    }

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
            var mTheme = ORG_THEMES[m.org] || ORG_THEMES['__HOME__'];
            html += '<a href="#" style="color:' + mTheme.accent + '" onclick="event.preventDefault();navigateToOrgCard(\'' + escHtml(m.org) + "','" + escHtml(m.id) + '\')">' + escHtml(displayName(m.name)) + leadLabel + '</a>';
            html += talentTooltip(m);
            html += ' ' + badge;
            if (m.title) html += ' <span class="member-title">' + escHtml(m.title) + '</span>';
            html += '</div>';
        });

        html += '</div>';
    }

    html += '</div>';
    document.getElementById('mainContent').innerHTML = html;
    document.getElementById('mainContent').focus();
    updateHeadcount();
}

// ── All-Scrum View ──
function showAllScrumView() {
    applyOrgTheme('__HOME__');
    state.allScrumView = true;
    state.scrumView = null;
    state.listView = false;
    state.isHome = false;
    document.getElementById('backToOrg').style.display = 'none';

    const bc = document.getElementById('breadcrumb');
    bc.innerHTML = '<span class="current">All Scrum Teams</span>';

    // Sort teams by total member count descending
    const teamNames = Object.keys(DATA.scrum).sort(function(a, b) {
        function countMembers(t) {
            let n = 0;
            for (const disc in DATA.scrum[t]) n += DATA.scrum[t][disc].length;
            return n;
        }
        return countMembers(b) - countMembers(a);
    });

    const disciplineOrder = ['Dev', 'Product', 'QA', 'TPM', 'Other'];
    const discShort = { 'Dev': 'Dev', 'Product': 'Product', 'QA': 'QA', 'TPM': 'TPM', 'Other': 'Other' };

    let gridHtml = '';
    teamNames.forEach(function(teamName) {
        const groups = DATA.scrum[teamName];
        let totalMembers = 0;
        for (const d in groups) totalMembers += groups[d].length;

        let bodyHtml = '';
        disciplineOrder.forEach(function(disc) {
            let members = groups[disc] || [];
            members = members.filter(function(m) {
                var con = isContractor(m.employment);
                return con ? state.showCon : state.showEmp;
            });
            if (members.length === 0) return;

            bodyHtml += '<div class="disc-label">' + escHtml(discShort[disc] || disc) + '</div>';
            members.forEach(function(m, idx) {
                const isLead = idx === 0;
                const leadTag = isLead ? ' <span style="color:#38a169;font-size:11px">(Lead)</span>' : '';
                const badge = isContractor(m.employment)
                    ? '<span class="badge badge-contractor" style="font-size:10px;padding:1px 5px">C</span>'
                    : '';
                var memTheme = ORG_THEMES[m.org] || ORG_THEMES['__HOME__'];
                bodyHtml += '<div class="member-row' + (isLead ? ' is-lead' : '') + '">'
                    + '<span style="color:' + memTheme.accent + '">' + escHtml(displayName(m.name)) + leadTag + '</span> ' + badge
                    + '</div>';
            });
        });

        var cardMeta = DATA.scrumMeta && DATA.scrumMeta[teamName];
        var cardMetaHtml = '';
        if (cardMeta && (cardMeta.productOwner || cardMeta.scrumMaster)) {
            cardMetaHtml = '<div class="all-scrum-card-meta">';
            if (cardMeta.productOwner) cardMetaHtml += '<span>PO: ' + escHtml(displayName(cardMeta.productOwner)) + '</span>';
            if (cardMeta.scrumMaster) cardMetaHtml += '<span>SM: ' + escHtml(displayName(cardMeta.scrumMaster)) + '</span>';
            cardMetaHtml += '</div>';
        }

        gridHtml += '<div class="all-scrum-card" role="button" tabindex="0" onclick="showScrumView(\'' + escHtml(teamName) + '\')" onkeydown="if(event.key===\'Enter\'||event.key===\' \'){event.preventDefault();showScrumView(\'' + escHtml(teamName) + '\')}">'
            + '<div class="all-scrum-card-header">' + escHtml(teamName)
            + ' <span class="count-badge">' + totalMembers + '</span></div>'
            + cardMetaHtml
            + '<div class="all-scrum-card-body">' + bodyHtml + '</div>'
            + '</div>';
    });

    let html = '<div class="all-scrum-viewport" id="allScrumViewport">'
        + '<div class="all-scrum-canvas" id="allScrumCanvas">'
        + '<div class="all-scrum-grid">' + gridHtml + '</div>'
        + '</div></div>'
        + '<div class="zoom-controls">'
        + '<button onclick="allScrumZoom(1)" title="Zoom In">+</button>'
        + '<button onclick="allScrumZoom(-1)" title="Zoom Out">&minus;</button>'
        + '<button onclick="allScrumReset()" title="Reset">&#8634;</button>'
        + '</div>';

    document.getElementById('mainContent').innerHTML = html;
    document.getElementById('mainContent').focus();
    initAllScrumPanZoom();

    // Update headcount: unique members across all scrum teams
    var seen = {}, fte = 0, con = 0;
    for (var tn in DATA.scrum) {
        for (var d in DATA.scrum[tn]) {
            DATA.scrum[tn][d].forEach(function(m) {
                var nk = m.name.toLowerCase();
                if (seen[nk]) return;
                seen[nk] = true;
                var isCon = isContractor(m.employment);
                if (isCon) { if (state.showCon) con++; }
                else { if (state.showEmp) fte++; }
            });
        }
    }
    var total = fte + con;
    var el = document.getElementById('headcount');
    el.textContent = total + ' people in scrum teams (' + fte + ' FTE, ' + con + ' contractors)';
}

var _panZoom = { scale: 1 };

function applyPanZoomTransform() {
    var canvas = document.getElementById('allScrumCanvas');
    if (canvas) canvas.style.transform = 'scale(' + _panZoom.scale + ')';
}

function allScrumZoom(dir) {
    var factor = dir > 0 ? 1.2 : 1/1.2;
    _panZoom.scale = Math.min(3, Math.max(0.3, _panZoom.scale * factor));
    applyPanZoomTransform();
}

function allScrumReset() {
    _panZoom.scale = 1;
    applyPanZoomTransform();
    var vp = document.getElementById('allScrumViewport');
    if (vp) { vp.scrollTop = 0; vp.scrollLeft = 0; }
}

function initAllScrumPanZoom() {
    _panZoom = { scale: 1 };
    var vp = document.getElementById('allScrumViewport');
    if (!vp) return;

    vp.addEventListener('wheel', function(e) {
        if (!e.ctrlKey && !e.metaKey) return; // normal scroll without Ctrl/Cmd
        e.preventDefault();
        var dir = e.deltaY < 0 ? 1 : -1;
        allScrumZoom(dir);
    }, { passive: false });

    // Touch: pinch-to-zoom only (scrolling handled natively)
    var lastTouchDist = 0;
    vp.addEventListener('touchstart', function(e) {
        if (e.touches.length === 2) {
            lastTouchDist = Math.hypot(e.touches[1].clientX - e.touches[0].clientX, e.touches[1].clientY - e.touches[0].clientY);
        }
    }, { passive: true });

    vp.addEventListener('touchmove', function(e) {
        if (e.touches.length === 2) {
            var dist = Math.hypot(e.touches[1].clientX - e.touches[0].clientX, e.touches[1].clientY - e.touches[0].clientY);
            if (lastTouchDist > 0) {
                var ratio = dist / lastTouchDist;
                _panZoom.scale = Math.min(3, Math.max(0.3, _panZoom.scale * ratio));
                applyPanZoomTransform();
            }
            lastTouchDist = dist;
        }
    }, { passive: true });

    vp.addEventListener('touchend', function() {
        lastTouchDist = 0;
    }, { passive: true });
}

// ── List View ──
function getStackRank(talentCategory) {
    if (!talentCategory) return '';
    var match = talentCategory.match(/^([\d.]+)/);
    if (!match) return '';
    var rating = parseFloat(match[1]);
    if (rating >= 4) return 'H';
    if (rating >= 3) return 'M';
    return 'L';
}

function exportListToExcel() {
    var rows = collectListRows();
    rows = sortListRows(rows);
    var headers = ['Name', 'Title', 'Type', 'Manager', 'Org', 'Scrum Teams', 'Rating', 'Stack Rank', 'Supplier', '#', 'Start Date', 'Email'];
    var csvRows = [headers.join(',')];
    rows.forEach(function(r) {
        var name = displayName(r.name).replace(/"/g, '""');
        var title = (r.title || '').replace(/"/g, '""');
        var manager = displayName(r.manager).replace(/"/g, '""');
        var teams = (r.scrumTeams || []).join('; ');
        var rating = (r.talentCategory || '').replace(/"/g, '""');
        var stackRank = r.stackRank || getStackRank(r.talentCategory);
        var supplier = (r.supplier || '').replace(/"/g, '""');
        var num = r.contractorNumber ? String(r.contractorNumber) : '';
        var startDate = r.startDate || '';
        var email = (r.email || '').replace(/"/g, '""');
        csvRows.push('"' + name + '","' + title + '","' + r.type + '","' + manager + '","' + r.org + '","' + teams + '","' + rating + '","' + stackRank + '","' + supplier + '","' + num + '","' + startDate + '","' + email + '"');
    });
    var csv = csvRows.join('\n');
    var blob = new Blob(['\uFEFF' + csv], {type: 'text/csv;charset=utf-8;'});
    var url = URL.createObjectURL(blob);
    var a = document.createElement('a');
    a.href = url;
    var label = state.isHome ? 'All_Orgs' : state.currentOrg.replace(/\s+/g, '_');
    a.download = 'orgchart_list_' + label + '.csv';
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    URL.revokeObjectURL(url);
}

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
                talentBand: node.talentBand || '',
                talentCategory: node.talentCategory || '',
                rationale: node.rationale || '',
                stackRank: node.stackRank || getStackRank(node.talentCategory || ''),
                supplier: node.supplier || '',
                contractorNumber: node.contractorNumber || '',
                startDate: node.startDate || '',
                email: node.email || '',
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
        {key: 'talentCategory', label: 'Rating'},
        {key: 'stackRank', label: 'Stack Rank'},
        {key: 'supplier', label: 'Supplier'},
        {key: 'contractorNumber', label: '#'},
        {key: 'startDate', label: 'Start Date'},
        {key: 'email', label: 'Email'},
    ];
    var html = '<div class="list-view"><table><thead><tr>';
    cols.forEach(function(c) {
        var sortable = c.key !== 'scrumTeams';
        var arrow = '';
        if (sortable) {
            var isActive = state.listSortCol === c.key;
            var dir = isActive ? (state.listSortAsc ? '&#9650;' : '&#9660;') : '&#9650;';
            arrow = ' <span class="sort-arrow' + (isActive ? ' active' : '') + '" aria-hidden="true">' + dir + '</span>';
        }
        if (sortable) {
            var sortAttr = isActive ? (state.listSortAsc ? 'ascending' : 'descending') : 'none';
            html += '<th aria-sort="' + sortAttr + '"><button type="button" onclick="sortListBy(\'' + c.key + '\')">' + c.label + arrow + '</button></th>';
        } else {
            html += '<th>' + c.label + '</th>';
        }
    });
    html += '</tr></thead><tbody>';
    rows.forEach(function(r) {
        html += '<tr>';
        var lnkTheme = ORG_THEMES[r.orgKey] || ORG_THEMES['__HOME__'];
        html += '<td><a href="#" style="color:' + lnkTheme.accent + '" onclick="event.preventDefault();navigateToOrgCard(\'' + escHtml(r.orgKey) + "','" + escHtml(r.nodeId) + '\')">' + escHtml(displayName(r.name)) + '</a>' + talentTooltip(r) + '</td>';
        html += '<td>' + escHtml(r.title) + '</td>';
        var badgeCls = r.type === 'Contractor' ? 'badge-contractor' : 'badge-fte';
        html += '<td><span class="badge ' + badgeCls + '">' + r.type + '</span></td>';
        html += '<td>' + escHtml(displayName(r.manager)) + '</td>';
        html += '<td>' + escHtml(r.org) + '</td>';
        html += '<td>';
        if (r.scrumTeams.length) {
            r.scrumTeams.forEach(function(t) {
                html += '<a class="team-pill" href="#" style="background:' + teamColor(t) + '22;color:' + teamPillTextColor(t) + ';cursor:pointer" onclick="event.preventDefault();showScrumView(\'' + escHtml(t).replace(/'/g,"\\\\'") + '\')">#' + escHtml(t) + '</a>';
            });
        }
        html += '</td>';
        html += '<td>' + escHtml(r.talentCategory) + '</td>';
        var rankCls = r.stackRank === 'H' ? 'rank-high' : r.stackRank === 'M' ? 'rank-med' : r.stackRank === 'L' ? 'rank-low' : '';
        html += '<td>' + (r.stackRank ? '<span class="badge ' + rankCls + '">' + r.stackRank + '</span>' : '') + '</td>';
        html += '<td>' + escHtml(r.supplier) + '</td>';
        html += '<td>' + (r.contractorNumber ? escHtml(String(r.contractorNumber)) : '') + '</td>';
        html += '<td>' + escHtml(r.startDate) + '</td>';
        html += '<td>' + (r.email ? '<a href="mailto:' + escHtml(r.email) + '">' + escHtml(r.email) + '</a>' : '') + '</td>';
        html += '</tr>';
    });
    html += '</tbody></table></div>';
    return html;
}

function showListView() {
    state.listView = true;
    state.scrumView = null;
    state.allScrumView = false;
    document.getElementById('backToOrg').style.display = state.isHome ? 'none' : '';
    renderList();
}

function renderList() {
    var rows = collectListRows();
    rows = sortListRows(rows);

    var label = state.isHome ? 'All Orgs' : state.currentOrg;
    var bc = document.getElementById('breadcrumb');
    bc.innerHTML = '<span class="current">List — ' + escHtml(label) + ' (' + rows.length + ')</span>';

    var exportBtn = '<div style="text-align:right;margin-bottom:10px"><button onclick="exportListToExcel()" style="padding:6px 16px;border:1px solid #d1d5db;border-radius:6px;background:#fff;cursor:pointer;font-size:14px;color:#374151">&#128229; Export to Excel</button></div>';
    document.getElementById('mainContent').innerHTML = exportBtn + renderListTable(rows);
    document.getElementById('mainContent').focus();

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
        if (i > 0) bcHtml += '<span aria-hidden="true">&rsaquo;</span> ';
        if (i === state.breadcrumb.length - 1) {
            bcHtml += '<span class="current">' + escHtml(displayName(n.name)) + '</span>';
        } else {
            bcHtml += '<a href="#" onclick="event.preventDefault();navigateTo(\'' + escHtml(id) + '\')">' + escHtml(displayName(n.name)) + '</a> ';
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
            '<a class="team-pill" href="#" style="background:' + teamColor(t) + '22;color:' + teamPillTextColor(t) + '" onclick="event.preventDefault();showScrumView(\'' + escHtml(t).replace(/'/g,"\\\\'") + '\')">#' + escHtml(t) + '</a>'
        ).join('') + '</div>';
    }

    let html = '<div class="manager-section">';
    html += '<div class="manager-card">';
    html += '<div class="name">' + escHtml(displayName(node.name)) + talentTooltip(node) + '</div>';
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
                    '<a class="team-pill" href="#" style="background:' + teamColor(t) + '22;color:' + teamPillTextColor(t) + '" onclick="event.preventDefault();event.stopPropagation();showScrumView(\'' + escHtml(t).replace(/'/g,"\\\\'") + '\')">#' + escHtml(t) + '</a>'
                ).join('') + '</div>';
            }

            // Count all descendants (respecting FTE filter)
            let drCount = countReports(org, child.id);
            let drLabel = drCount > 0 ? '<div class="dr-count">' + drCount + ' report' + (drCount > 1 ? 's' : '') + '</div>' : '';

            const phClass = child.placeholder ? ' placeholder' : '';
            html += '<div class="person-card' + phClass + '" role="button" tabindex="0" onclick="navigateTo(\'' + escHtml(child.id) + '\')" onkeydown="if(event.key===\'Enter\'||event.key===\' \'){event.preventDefault();navigateTo(\'' + escHtml(child.id) + '\')}">';
            html += '<div class="name">' + escHtml(displayName(child.name)) + talentTooltip(child) + '</div>';
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
        html += '<button class="missing-toggle" aria-expanded="false" onclick="var list=this.nextElementSibling;list.classList.toggle(\'open\');this.setAttribute(\'aria-expanded\',list.classList.contains(\'open\'))">';
        html += '<span aria-hidden="true">&#9888;</span> ' + missing.length + ' people without titles in ' + escHtml(state.currentOrg) + ' (click to expand)';
        html += '</button>';
        html += '<div class="missing-list"><ul role="list">';
        missing.forEach(name => {
            html += '<li>' + escHtml(displayName(name)) + '</li>';
        });
        html += '</ul></div></div>';
    }

    document.getElementById('mainContent').innerHTML = html;
    document.getElementById('mainContent').focus();
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
document.addEventListener('keydown', function(e) {
    if (e.key === 'Escape') {
        var active = document.activeElement;
        if (active && active.classList.contains('talent-info')) {
            active.blur();
        }
    }
});
</script>
</body>
</html>'''
