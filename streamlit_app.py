"""
Streamlit app for the Org Chart viewer.

Password-protected app with:
- Named org chart view (embed + download)
- Redacted org chart view (embed + download)
- Admin panel (upload/download Excel, regenerate HTMLs, commit to GitHub)
"""

import sys
import base64
import tempfile
from pathlib import Path

import streamlit as st
import requests

# Ensure project root is on sys.path so we can import generators
PROJECT_ROOT = Path(__file__).parent
if str(PROJECT_ROOT) not in sys.path:
    sys.path.insert(0, str(PROJECT_ROOT))

from generate_org_html import (
    parse_people,
    parse_scrum_teams,
    build_org_datasets,
    build_scrum_data,
    build_home_drs,
)
from org_html_shared import generate_html, redact_data, verify_redaction, normalize_name

# ─── Config ──────────────────────────────────────────────────────────────────

NAMED_HTML = PROJECT_ROOT / "org_drilldown.html"
REDACTED_HTML = PROJECT_ROOT / "org_drilldown_redacted.html"
MASTER_EXCEL = PROJECT_ROOT / "data" / "orgchart_master_data.xlsx"

st.set_page_config(
    page_title="Org Chart",
    page_icon="🏢",
    layout="wide",
)


# ─── Authentication ──────────────────────────────────────────────────────────

def check_auth():
    """Return True if the user is authenticated."""
    return st.session_state.get("authenticated", False)


def login_form():
    """Render the login form and handle authentication."""
    st.markdown("## Org Chart Viewer")
    st.markdown("Please enter the password to continue.")

    password = st.text_input("Password", type="password", key="login_password")
    if st.button("Login", type="primary"):
        if password == st.secrets["app_password"]:
            st.session_state["authenticated"] = True
            st.rerun()
        else:
            st.error("Incorrect password.")


# ─── Views ───────────────────────────────────────────────────────────────────

def view_org_chart(html_path: Path, label: str):
    """Embed an HTML org chart with a download button."""
    if not html_path.exists():
        st.warning(f"{html_path.name} not found. Run the generator first.")
        return

    html_content = html_path.read_text(encoding="utf-8")

    col1, col2 = st.columns([6, 1])
    with col1:
        st.subheader(label)
    with col2:
        st.download_button(
            label="Download HTML",
            data=html_content,
            file_name=html_path.name,
            mime="text/html",
        )

    st.components.v1.html(html_content, height=800, scrolling=True)


# ─── Admin ───────────────────────────────────────────────────────────────────

def _github_headers():
    """Return auth headers for GitHub API."""
    return {
        "Authorization": f"token {st.secrets['github_token']}",
        "Accept": "application/vnd.github.v3+json",
    }


def _get_file_sha(repo: str, path: str) -> str | None:
    """Get the current SHA of a file in the repo (needed for updates)."""
    url = f"https://api.github.com/repos/{repo}/contents/{path}"
    resp = requests.get(url, headers=_github_headers())
    if resp.status_code == 200:
        return resp.json()["sha"]
    return None


def _commit_file(repo: str, path: str, content_bytes: bytes, message: str):
    """Create or update a file in GitHub via the Contents API."""
    url = f"https://api.github.com/repos/{repo}/contents/{path}"
    encoded = base64.b64encode(content_bytes).decode("ascii")
    payload = {
        "message": message,
        "content": encoded,
    }
    sha = _get_file_sha(repo, path)
    if sha:
        payload["sha"] = sha
    resp = requests.put(url, json=payload, headers=_github_headers())
    resp.raise_for_status()
    return resp.json()


def regenerate_from_excel(excel_path: Path):
    """Run the generation pipeline on the given Excel file.

    Returns (named_html: str, redacted_html: str, error: str|None).
    """
    try:
        people = parse_people(excel_path)
        scrum_sheet = parse_scrum_teams(excel_path)
        org_datasets = build_org_datasets(people)
        scrum_teams = build_scrum_data(people, scrum_sheet, org_datasets)
        home_drs = build_home_drs(org_datasets)

        data = {
            "orgs": {
                org_name: {
                    "top": ds["top"],
                    "nodes": ds["nodes"],
                    "children": ds["children"],
                }
                for org_name, ds in org_datasets.items()
            },
            "scrum": scrum_teams,
            "homeDrs": home_drs,
            "missing": {},
        }

        named_html = generate_html(data, redacted=False)

        all_names = set()
        for ds in org_datasets.values():
            for nid, node in ds["nodes"].items():
                if not node.get("placeholder"):
                    all_names.add(node["name"])

        redacted_data = redact_data(data, all_names)
        redacted_html = generate_html(redacted_data, redacted=True)

        leaked = verify_redaction(redacted_html, all_names)
        if leaked:
            return named_html, redacted_html, f"Redaction warning: {len(leaked)} names may have leaked"

        return named_html, redacted_html, None

    except Exception as e:
        return None, None, str(e)


def admin_panel():
    """Admin panel: download/upload Excel, regenerate, commit to GitHub."""
    st.subheader("Admin Panel")

    # ── Download current Excel ──
    st.markdown("### Download Current Excel")
    if MASTER_EXCEL.exists():
        with open(MASTER_EXCEL, "rb") as f:
            st.download_button(
                label="Download orgchart_master_data.xlsx",
                data=f.read(),
                file_name="orgchart_master_data.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )
    else:
        st.warning("Master Excel file not found.")

    st.divider()

    # ── Upload new Excel ──
    st.markdown("### Upload New Excel & Regenerate")
    uploaded = st.file_uploader(
        "Upload a new orgchart_master_data.xlsx",
        type=["xlsx"],
        key="excel_upload",
    )

    if uploaded and st.button("Regenerate & Commit to GitHub", type="primary"):
        with st.spinner("Regenerating org charts..."):
            # Write uploaded file to temp location
            with tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False) as tmp:
                tmp.write(uploaded.getvalue())
                tmp_path = Path(tmp.name)

            named_html, redacted_html, error = regenerate_from_excel(tmp_path)
            tmp_path.unlink(missing_ok=True)

            if error and named_html is None:
                st.error(f"Generation failed: {error}")
                return

            if error:
                st.warning(error)

        st.success("HTML files regenerated successfully.")

        # Commit to GitHub
        repo = st.secrets["github_repo"]
        with st.spinner("Committing to GitHub..."):
            try:
                _commit_file(
                    repo,
                    "data/orgchart_master_data.xlsx",
                    uploaded.getvalue(),
                    "Update master Excel via Streamlit admin",
                )
                _commit_file(
                    repo,
                    "org_drilldown.html",
                    named_html.encode("utf-8"),
                    "Regenerate named org chart via Streamlit admin",
                )
                _commit_file(
                    repo,
                    "org_drilldown_redacted.html",
                    redacted_html.encode("utf-8"),
                    "Regenerate redacted org chart via Streamlit admin",
                )
                st.success("All 3 files committed to GitHub. Streamlit Cloud will auto-redeploy.")
            except requests.HTTPError as e:
                st.error(f"GitHub commit failed: {e}")
                st.code(e.response.text if e.response else "No response body")


# ─── Main ────────────────────────────────────────────────────────────────────

def main():
    if not check_auth():
        login_form()
        return

    # Sidebar navigation
    st.sidebar.title("Org Chart")
    view = st.sidebar.radio(
        "View",
        ["Org Chart (Named)", "Org Chart (Redacted)", "Admin"],
        label_visibility="collapsed",
    )

    if st.sidebar.button("Logout"):
        st.session_state["authenticated"] = False
        st.rerun()

    if view == "Org Chart (Named)":
        view_org_chart(NAMED_HTML, "Org Chart")
    elif view == "Org Chart (Redacted)":
        view_org_chart(REDACTED_HTML, "Org Chart (Redacted)")
    elif view == "Admin":
        admin_panel()


if __name__ == "__main__":
    main()
