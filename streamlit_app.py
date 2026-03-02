"""
Streamlit app for the Org Chart viewer.

Password-protected app with:
- Named org chart view (embed + download)
- Redacted org chart view (embed + download)
- Admin panel (upload/download Excel, regenerate HTMLs, commit to GitHub)

Data files (Excel + HTML) live in a separate private repo (orgchart-data)
and are fetched at runtime via the GitHub API.
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

st.set_page_config(
    page_title="Org Chart",
    page_icon="🏢",
    layout="wide",
)


# ─── Custom Styling ──────────────────────────────────────────────────────────

def inject_custom_css():
    """Inject custom CSS to override default Streamlit styles."""
    st.markdown("""
    <style>
    @import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700&display=swap');

    /* Hide default Streamlit branding */
    #MainMenu {visibility: hidden;}
    footer {visibility: hidden;}
    header[data-testid="stHeader"] {display: none;}

    /* Global typography */
    html, body, [class*="css"] {
        font-family: 'Inter', -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, sans-serif;
    }

    /* Sidebar: dark gradient */
    section[data-testid="stSidebar"] {
        background: linear-gradient(180deg, #0f172a 0%, #1e293b 100%) !important;
    }
    section[data-testid="stSidebar"] * {
        color: #f1f5f9 !important;
    }
    section[data-testid="stSidebar"] hr {
        border-color: rgba(255,255,255,0.1) !important;
    }
    section[data-testid="stSidebar"] .stButton > button {
        background: rgba(255,255,255,0.06) !important;
        border: 1px solid rgba(255,255,255,0.12) !important;
        color: #f1f5f9 !important;
        border-radius: 8px;
        font-weight: 500;
        transition: all 0.2s;
    }
    section[data-testid="stSidebar"] .stButton > button:hover {
        background: rgba(255,255,255,0.12) !important;
        border-color: rgba(255,255,255,0.2) !important;
    }

    /* Primary button styling */
    .stButton > button[kind="primary"],
    .stButton > button[data-testid="stBaseButton-primary"] {
        background: linear-gradient(135deg, #3b82f6 0%, #2563eb 100%) !important;
        border: none !important;
        border-radius: 8px;
        padding: 0.6rem 2rem;
        font-weight: 600;
        letter-spacing: 0.3px;
        transition: all 0.2s;
        box-shadow: 0 2px 8px rgba(59, 130, 246, 0.3);
    }
    .stButton > button[kind="primary"]:hover,
    .stButton > button[data-testid="stBaseButton-primary"]:hover {
        box-shadow: 0 4px 16px rgba(59, 130, 246, 0.4);
        transform: translateY(-1px);
    }

    /* Download button styling */
    .stDownloadButton > button {
        background: #f8fafc !important;
        border: 1px solid #e2e8f0 !important;
        border-radius: 8px;
        color: #0f172a !important;
        font-weight: 500;
    }
    .stDownloadButton > button:hover {
        background: #f1f5f9 !important;
        border-color: #cbd5e1 !important;
    }

    /* File uploader */
    [data-testid="stFileUploader"] {
        border: 2px dashed #cbd5e1 !important;
        border-radius: 12px;
        padding: 1.5rem;
        background: #f8fafc;
    }
    [data-testid="stFileUploader"]:hover {
        border-color: #3b82f6 !important;
    }
    </style>
    """, unsafe_allow_html=True)


# ─── GitHub Data Access ──────────────────────────────────────────────────────

def _github_headers():
    """Return auth headers for GitHub API."""
    return {
        "Authorization": f"token {st.secrets['github_token']}",
        "Accept": "application/vnd.github.v3+json",
    }


def _data_repo():
    """Return the private data repo name."""
    return st.secrets["data_repo"]


@st.cache_data(ttl=300)
def _fetch_file(path: str) -> bytes | None:
    """Fetch a file from the private data repo. Returns raw bytes or None."""
    url = f"https://api.github.com/repos/{_data_repo()}/contents/{path}"
    resp = requests.get(url, headers=_github_headers())
    if resp.status_code != 200:
        return None
    content_b64 = resp.json()["content"]
    return base64.b64decode(content_b64)


def _fetch_html(path: str) -> str | None:
    """Fetch an HTML file from the private data repo as a string."""
    data = _fetch_file(path)
    if data is None:
        return None
    return data.decode("utf-8")


def _get_file_sha(repo: str, path: str) -> str | None:
    """Get the current SHA of a file in a repo (needed for updates)."""
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


# ─── Authentication ──────────────────────────────────────────────────────────

def check_auth():
    """Return True if the user is authenticated."""
    return st.session_state.get("authenticated", False)


def login_form():
    """Render a branded, centered login form."""
    inject_custom_css()

    col1, col2, col3 = st.columns([1, 1.5, 1])
    with col2:
        st.markdown("""
        <div style="text-align: center; padding: 4rem 0 1.5rem 0;">
            <div style="
                width: 64px; height: 64px;
                background: linear-gradient(135deg, #0f172a, #1e293b);
                border-radius: 16px; margin: 0 auto 1.5rem;
                display: flex; align-items: center; justify-content: center;
                box-shadow: 0 8px 24px rgba(15, 23, 42, 0.2);
                font-size: 28px; line-height: 1;
            ">&#127970;</div>
            <h1 style="
                font-family: 'Inter', sans-serif;
                font-size: 28px; font-weight: 700;
                color: #0f172a; margin-bottom: 8px;
            ">Org Chart</h1>
            <p style="color: #64748b; font-size: 15px; margin-bottom: 2rem;">
                Enter your password to access the org chart viewer
            </p>
        </div>
        """, unsafe_allow_html=True)

        password = st.text_input(
            "Password", type="password", key="login_password",
            label_visibility="collapsed", placeholder="Enter password...",
        )
        if st.button("Sign In", type="primary", use_container_width=True):
            if password == st.secrets["app_password"]:
                st.session_state["authenticated"] = True
                st.rerun()
            else:
                st.error("Incorrect password. Please try again.")


# ─── Views ───────────────────────────────────────────────────────────────────

def view_org_chart(remote_path: str, filename: str, label: str):
    """Fetch and embed an HTML org chart with a download button."""
    html_content = _fetch_html(remote_path)
    if html_content is None:
        st.warning(f"{filename} not found in data repo. Upload via Admin to generate.")
        return

    col1, col2 = st.columns([6, 1])
    with col1:
        st.subheader(label)
    with col2:
        st.download_button(
            label="Download HTML",
            data=html_content,
            file_name=filename,
            mime="text/html",
        )

    st.components.v1.html(html_content, height=900, scrolling=True)


# ─── Admin ───────────────────────────────────────────────────────────────────

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
    st.markdown("""
    <div style="
        background: linear-gradient(135deg, #0f172a 0%, #1e293b 100%);
        padding: 1.5rem 2rem; border-radius: 12px; margin-bottom: 1.5rem;
        color: white;
    ">
        <h2 style="font-family: 'Inter', sans-serif; font-size: 22px;
            font-weight: 700; margin: 0 0 4px 0; color: white;">Admin Panel</h2>
        <p style="color: #94a3b8; font-size: 14px; margin: 0;">
            Manage org chart data and regenerate HTML files</p>
    </div>
    """, unsafe_allow_html=True)

    data_repo = _data_repo()

    # ── Download current Excel ──
    st.markdown("### Download Current Excel")
    excel_bytes = _fetch_file("data/orgchart_master_data.xlsx")
    if excel_bytes:
        st.download_button(
            label="Download orgchart_master_data.xlsx",
            data=excel_bytes,
            file_name="orgchart_master_data.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
    else:
        st.warning("Master Excel file not found in data repo.")

    st.divider()

    # ── Upload new Excel ──
    st.markdown("### Upload New Excel & Regenerate")
    uploaded = st.file_uploader(
        "Upload a new orgchart_master_data.xlsx",
        type=["xlsx"],
        key="excel_upload",
    )

    if uploaded and st.button("Regenerate & Commit", type="primary"):
        with st.spinner("Regenerating org charts..."):
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

        # Commit all 3 files to the private data repo
        with st.spinner("Committing to data repo..."):
            try:
                _commit_file(
                    data_repo,
                    "data/orgchart_master_data.xlsx",
                    uploaded.getvalue(),
                    "Update master Excel via Streamlit admin",
                )
                _commit_file(
                    data_repo,
                    "org_drilldown.html",
                    named_html.encode("utf-8"),
                    "Regenerate named org chart via Streamlit admin",
                )
                _commit_file(
                    data_repo,
                    "org_drilldown_redacted.html",
                    redacted_html.encode("utf-8"),
                    "Regenerate redacted org chart via Streamlit admin",
                )
                # Clear the cached data so next view loads fresh files
                _fetch_file.clear()
                st.success("All 3 files committed to data repo.")
            except requests.HTTPError as e:
                st.error(f"GitHub commit failed: {e}")
                st.code(e.response.text if e.response else "No response body")


# ─── Main ────────────────────────────────────────────────────────────────────

def main():
    if not check_auth():
        login_form()
        return

    inject_custom_css()

    # Sidebar branding + navigation
    st.sidebar.markdown("""
    <div style="text-align: center; padding: 0.5rem 0 1.5rem 0;
        border-bottom: 1px solid rgba(255,255,255,0.1); margin-bottom: 1rem;">
        <div style="font-size: 20px; font-weight: 700; color: #f8fafc;
            letter-spacing: -0.5px;">
            &#127970; Org Chart
        </div>
        <div style="font-size: 12px; color: #94a3b8; margin-top: 4px;">
            Organization Viewer</div>
    </div>
    """, unsafe_allow_html=True)

    view = st.sidebar.radio(
        "Navigation",
        ["Org Chart (Named)", "Org Chart (Redacted)", "Admin"],
        label_visibility="collapsed",
    )

    st.sidebar.divider()
    if st.sidebar.button("Sign Out", use_container_width=True):
        st.session_state["authenticated"] = False
        st.rerun()

    if view == "Org Chart (Named)":
        view_org_chart("org_drilldown.html", "org_drilldown.html", "Org Chart")
    elif view == "Org Chart (Redacted)":
        view_org_chart("org_drilldown_redacted.html", "org_drilldown_redacted.html", "Org Chart (Redacted)")
    elif view == "Admin":
        admin_panel()


if __name__ == "__main__":
    main()
