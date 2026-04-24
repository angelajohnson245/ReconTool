"""
Financing Line Reconciliation Tool — Streamlit UI
Run with: streamlit run recon_streamlit_app.py
"""
import io
import os
import re
import tempfile
from datetime import datetime
 
import pandas as pd
import streamlit as st
 
from recon_enhanced_output import (
    PRIMARY_TYPE_FUND_CONFIG,
    STREAMLIT_PRIMARY_FILE_TYPES,
    PrimaryFileSchemaError,
    build_output_filename,
    build_workbook,
    filter_recon_to_selected_fund,
    get_primary_config,
    normalise_facility,
    normalise_text,
    normalize_recon_fund_for_output,
    reconcile,
    scope_label_for_primary_type,
)
 
# --------------------------------------------------
# PAGE CONFIG
# --------------------------------------------------
st.set_page_config(
    page_title="Financing Line Recon",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded",
)
 
# --------------------------------------------------
# STYLES
# --------------------------------------------------
st.markdown("""
<style>
  @import url('https://fonts.googleapis.com/css2?family=IBM+Plex+Mono:wght@400;600&family=IBM+Plex+Sans:wght@300;400;600;700&display=swap');
 
  html, body, [class*="css"] {
    font-family: 'IBM Plex Sans', sans-serif;
  }
 
  /* Sidebar */
  [data-testid="stSidebar"] {
    background: #0d1b2a;
    border-right: 1px solid #1f3450;
  }
  [data-testid="stSidebar"] * {
    color: #c8d8e8 !important;
  }
  [data-testid="stSidebar"] h1,
  [data-testid="stSidebar"] h2,
  [data-testid="stSidebar"] h3 {
    color: #e8f4fd !important;
    font-weight: 700;
    letter-spacing: 0.04em;
    font-size: 0.78rem;
    text-transform: uppercase;
  }
 
  /* Main background */
  .main .block-container {
    background: #f4f7fb;
    padding-top: 1.5rem;
  }
 
  /* Top header */
  .recon-header {
    background: linear-gradient(135deg, #0d1b2a 0%, #1a3a5c 60%, #1f5280 100%);
    border-radius: 12px;
    padding: 28px 36px;
    margin-bottom: 24px;
    display: flex;
    align-items: center;
    justify-content: space-between;
  }
  .recon-header h1 {
    color: #ffffff;
    font-size: 1.55rem;
    font-weight: 700;
    margin: 0;
    letter-spacing: -0.02em;
  }
  .recon-header p {
    color: #8fb8d8;
    font-size: 0.82rem;
    margin: 4px 0 0;
    font-family: 'IBM Plex Mono', monospace;
  }
  .badge {
    background: rgba(255,255,255,0.1);
    border: 1px solid rgba(255,255,255,0.2);
    border-radius: 6px;
    padding: 6px 14px;
    color: #c8e6ff;
    font-size: 0.75rem;
    font-family: 'IBM Plex Mono', monospace;
    white-space: nowrap;
  }
 
  /* Metric cards */
  .metric-row {
    display: flex;
    gap: 14px;
    margin-bottom: 24px;
  }
  .metric-card {
    flex: 1;
    border-radius: 10px;
    padding: 18px 22px;
    border: 1px solid;
  }
  .metric-card .label {
    font-size: 0.7rem;
    font-weight: 600;
    text-transform: uppercase;
    letter-spacing: 0.08em;
    opacity: 0.7;
    margin-bottom: 6px;
  }
  .metric-card .value {
    font-size: 2.2rem;
    font-weight: 700;
    font-family: 'IBM Plex Mono', monospace;
    line-height: 1;
  }
  .metric-card .sub {
    font-size: 0.72rem;
    margin-top: 4px;
    opacity: 0.6;
  }
  .mc-total   { background: #e8f0fe; border-color: #b8cdf8; color: #1a3a6c; }
  .mc-match   { background: #e6f4ea; border-color: #a8d5b5; color: #1e5c35; }
  .mc-missing { background: #fef9e7; border-color: #f5d878; color: #7d5a00; }
  .mc-mismatch{ background: #fdecea; border-color: #f4b8b5; color: #8c1c15; }
 
  /* Status pills */
  .pill {
    display: inline-block;
    border-radius: 12px;
    padding: 2px 10px;
    font-size: 0.72rem;
    font-weight: 600;
    font-family: 'IBM Plex Mono', monospace;
    text-transform: uppercase;
    letter-spacing: 0.04em;
  }
  .pill-match   { background: #c6efce; color: #375623; }
  .pill-missing { background: #ffeb9c; color: #7d6608; }
  .pill-mismatch{ background: #ffc7ce; color: #9c0006; }
  .pill-na      { background: #e0e0e0; color: #555; }
 
  /* Section headers */
  .section-label {
    font-size: 0.68rem;
    font-weight: 700;
    text-transform: uppercase;
    letter-spacing: 0.12em;
    color: #5a7a99;
    margin: 20px 0 10px;
    padding-bottom: 6px;
    border-bottom: 1px solid #d8e4f0;
  }
 
  /* Download button */
  .stDownloadButton > button {
    background: linear-gradient(135deg, #1a3a5c, #1f5280) !important;
    color: #ffffff !important;
    border: none !important;
    border-radius: 8px !important;
    font-weight: 600 !important;
    letter-spacing: 0.03em !important;
    padding: 10px 24px !important;
    font-size: 0.88rem !important;
    transition: all 0.2s ease !important;
  }
  .stDownloadButton > button:hover {
    background: linear-gradient(135deg, #1f5280, #2874a6) !important;
    transform: translateY(-1px) !important;
    box-shadow: 0 4px 12px rgba(31,82,128,0.35) !important;
  }
 
  /* File uploader */
  [data-testid="stFileUploader"] {
    background: #f8fafd;
    border-radius: 10px;
    border: 1.5px dashed #b8cce4;
    padding: 6px;
  }
 
  /* Tabs */
  [data-testid="stTabs"] [role="tab"] {
    font-weight: 600;
    font-size: 0.82rem;
    letter-spacing: 0.03em;
  }
 
  /* Dataframe */
  [data-testid="stDataFrame"] {
    border-radius: 8px;
    overflow: hidden;
    border: 1px solid #dce8f4;
  }
  /* Status column visuals are driven by pandas Styler (see ``status_cols`` / ``set_properties``).
     Keep only light container rules here so non-status columns stay Streamlit-default. */
 
  /* Info box */
  .info-box {
    background: #eaf2ff;
    border-left: 4px solid #2874a6;
    border-radius: 6px;
    padding: 12px 16px;
    font-size: 0.82rem;
    color: #1a3a6c;
    margin-bottom: 16px;
  }
 
  /* Filter chips */
  .filter-row {
    display: flex;
    gap: 8px;
    flex-wrap: wrap;
    margin-bottom: 12px;
  }

  /* Sidebar: stack checkboxes vertically (Recon Status + other sidebar toggles) */
  [data-testid="stSidebar"] div[data-testid="stCheckbox"] {
    display: block !important;
    width: 100% !important;
    margin-bottom: 0.35rem;
  }
</style>
""", unsafe_allow_html=True)
 

def ui_primary_label(primary_type: str) -> str:
    cfg = get_primary_config(primary_type)
    return cfg.get("ui_display_label", cfg.get("display_name", primary_type))


def primary_scope_label_for_missing_banner(
    uploaded_filename: str | None, primary_file_type: str
) -> str:
    """Short fund/file token for UI-only missing-side labels (matches filename heuristics elsewhere)."""
    if uploaded_filename:
        name = uploaded_filename.upper()
        if re.search(r"\bACP\s+III\b", name):
            return PRIMARY_TYPE_FUND_CONFIG["ACORE"]["scope_label"]
        if re.search(r"\bACP\s+II\b", name):
            return PRIMARY_TYPE_FUND_CONFIG["ACP II"]["scope_label"]
        if re.search(r"\bAOC\s+II\b", name):
            return PRIMARY_TYPE_FUND_CONFIG["AOC II"]["scope_label"]
    cfg = PRIMARY_TYPE_FUND_CONFIG.get(primary_file_type) or {}
    return str(cfg.get("scope_label") or primary_file_type).strip()


def format_missing_status_display(
    raw,
    *,
    primary_scope_label: str,
    m61_label: str = "M61",
) -> str:
    """UI-only: normalize reconciliation missing strings to ``MISSING FROM … FILE`` wording."""
    if raw is None:
        return ""
    try:
        if isinstance(raw, float) and pd.isna(raw):
            return ""
    except (TypeError, ValueError):
        pass
    s = str(raw).strip()
    if not s:
        return ""

    su = s.upper()
    # Backend bug-shield: effective date path could double-prefix; collapse for display.
    dup_in = "MISSING IN MISSING IN "
    dup_from = "MISSING FROM MISSING FROM "
    if su.startswith(dup_in):
        s = "MISSING IN " + s[len(dup_in) :].lstrip()
        su = s.upper()
    elif su.startswith(dup_from):
        s = "MISSING FROM " + s[len(dup_from) :].lstrip()
        su = s.upper()

    def _is_m61_missing(u: str) -> bool:
        return (
            u.startswith("MISSING IN M61")
            or u.startswith("MISSING FROM M61")
            or (
                (u.startswith("MISSING IN ") or u.startswith("MISSING FROM "))
                and "M61" in u
            )
        )

    if _is_m61_missing(su):
        return f"MISSING FROM {m61_label} FILE"

    scope = str(primary_scope_label).strip()
    if not scope:
        return s

    if su.startswith("MISSING IN ") or su.startswith("MISSING FROM "):
        if not (su.startswith("MISSING IN M61") or su.startswith("MISSING FROM M61")):
            # Replace any prior token with the current upload/scope token.
            return f"MISSING FROM {scope} FILE"

    return s


def infer_primary_type_from_filename(uploaded_filename: str | None) -> str | None:
    """Best-effort type hint from uploaded business filename (warning only)."""
    if not uploaded_filename:
        return None
    name = uploaded_filename.upper()
    # Check ACP III before ACP II and use full-term matching to avoid substring overlap.
    if re.search(r"\bACP\s+III\b", name):
        return "ACORE"
    if re.search(r"\bACP\s+II\b", name):
        return "ACP II"
    if re.search(r"\bAOC\s+II\b", name):
        return "AOC II"
    if "ACORE" in name:
        return "ACORE"
    return None


def looks_like_m61_liability_relationship(filename: str | None) -> bool:
    """M61 comparison file validator (Liability_Relationship exports only)."""
    if not filename:
        return False
    name = filename.lower().replace("-", "_").replace(" ", "_")
    if "liability" not in name or "relationship" not in name:
        return False
    if "liabilitynote" in name or "assetnote" in name:
        return False
    return True


# --------------------------------------------------
# SIDEBAR
# --------------------------------------------------
with st.sidebar:
    st.markdown("## 📁 Upload Files")
    st.markdown("---")

    drag_drop_friendly = st.checkbox(
        "Drag & drop friendly upload",
        value=True,
        key="drag_drop_friendly_upload",
        help="Larger dashed drop zones and visible labels. You can always drag files onto the uploader or click to browse.",
    )
    if drag_drop_friendly:
        st.caption(
            "Drag a file from your computer onto a box below, or click the box to browse."
        )
        st.markdown(
            """
            <style>
            [data-testid="stSidebar"] [data-testid="stFileUploader"] {
                min-height: 108px !important;
                padding: 12px !important;
                border-width: 2px !important;
                border-color: #7eb8e8 !important;
                background: #f0f6fc !important;
            }
            [data-testid="stSidebar"] [data-testid="stFileUploader"] section {
                min-height: 72px !important;
            }
            </style>
            """,
            unsafe_allow_html=True,
        )
    else:
        st.markdown(
            """
            <style>
            [data-testid="stSidebar"] [data-testid="stFileUploader"] {
                min-height: unset !important;
                padding: 6px !important;
                border-width: 1.5px !important;
                border-color: #b8cce4 !important;
                background: #f8fafd !important;
            }
            [data-testid="stSidebar"] [data-testid="stFileUploader"] section {
                min-height: unset !important;
            }
            </style>
            """,
            unsafe_allow_html=True,
        )

    primary_file_type = st.selectbox(
        "Primary file type",
        STREAMLIT_PRIMARY_FILE_TYPES,
        index=0,
        format_func=ui_primary_label,
        help="Select which model template to use for column mapping.",
    )
    _pc = get_primary_config(primary_file_type)
    selected_ui_label = ui_primary_label(primary_file_type)

    st.markdown(f"**{selected_ui_label} business file** *(Liquidity & Earnings Model)*")
    file_b_upload = st.file_uploader(
        "Drag & drop or browse — business model (.xlsm, .xlsx)"
        if drag_drop_friendly
        else "Upload .xlsm or .xlsx",
        type=["xlsm", "xlsx"],
        key="file_b",
        label_visibility="visible" if drag_drop_friendly else "collapsed",
    )
    inferred_primary_type = infer_primary_type_from_filename(
        file_b_upload.name if file_b_upload else None
    )
    if inferred_primary_type and inferred_primary_type != primary_file_type:
        st.warning(
            f"Selected type is **{selected_ui_label}**, but uploaded business file name "
            f"looks like **{ui_primary_label(inferred_primary_type)}**. "
            "You can still run reconciliation."
        )

    st.markdown("**M61 file** *(In M61 export — comparison source)*")
    file_a_upload = st.file_uploader(
        "Drag & drop or browse — M61 Liability Relationship export (.xlsx)"
        if drag_drop_friendly
        else "Upload .xlsx",
        type=["xlsx"],
        key="file_a",
        label_visibility="visible" if drag_drop_friendly else "collapsed",
    )
    m61_file_valid = looks_like_m61_liability_relationship(
        file_a_upload.name if file_a_upload else None
    )
    if file_a_upload and not m61_file_valid:
        st.warning(
            "Uploaded comparison file does not look like a Liability Relationship export. "
            "Expected filename to include both `liability` and `relationship` "
            "(and not mapping names like `LiabilityNote` / `AssetNote`)."
        )

    # Optional mapping workbook is still supported by reconcile(); not exposed in UI.
    mapping_upload = None

    st.markdown("---")
    st.markdown("## ⚙️ Options")
    run_on_upload = st.checkbox("Auto-run on upload", value=True)
 
    st.markdown("---")
    st.markdown("## 🔍 Filters")
    st.caption("Recon status")


    if st.session_state.pop("reset_status_filters", False):
        st.session_state["filter_status_match"] = True
        st.session_state["filter_status_missing"] = True
        st.session_state["filter_status_mismatch"] = True
    if "filter_status_match" not in st.session_state:
        st.session_state["filter_status_match"] = True
    if "filter_status_missing" not in st.session_state:
        st.session_state["filter_status_missing"] = True
    if "filter_status_mismatch" not in st.session_state:
        st.session_state["filter_status_mismatch"] = True

    st.checkbox("Match", key="filter_status_match")
    st.checkbox("Missing", key="filter_status_missing")
    st.checkbox("Mismatch", key="filter_status_mismatch")
    scope_mode = st.radio(
        "Scope",
        ["All Results", "Selected Fund Only"],
        index=0,
        help=(
            "**All Results:** every reconciliation row. "
            "**Selected Fund Only:** only rows whose `Fund` belongs to "
            f"{scope_label_for_primary_type(primary_file_type)} scope "
            "(including **Both**, primary-only, and M61-only rows for that fund)."
        ),
    )
 
    st.markdown("---")
    st.markdown(
        f"""
    <div style='font-size:0.7rem; color:#5a7890; line-height:1.6'>
    <strong style='color:#8fb8d8'>Primary Source</strong><br>{_pc["model_descriptor"]}<br><br>
    <strong style='color:#8fb8d8'>Comparison Source</strong><br>(In) M61 Relationship Export<br><br>
    <strong style='color:#8fb8d8'>Target Advance Rate</strong><br>From M61 file only
    </div>
    """,
        unsafe_allow_html=True,
    )
 
 
# --------------------------------------------------
# HEADER
# --------------------------------------------------
st.markdown(
    f"""
<div class="recon-header">
  <div>
    <h1>📊 Financing Line Reconciliation</h1>
    <p>Primary: <strong>{selected_ui_label}</strong> · Comparison: M61 export</p>
  </div>
  <div class="badge">Run: {datetime.now().strftime("%b %d, %Y  %H:%M")}</div>
</div>
""",
    unsafe_allow_html=True,
)
 
 
# --------------------------------------------------
# HELPERS
# --------------------------------------------------
def pill(status):
    s = str(status).strip()
    su = s.upper()
    if su in ("", "N/A", "NAN", "NONE"):
        return '<span class="pill pill-na">N/A</span>'
    if "MATCH" in su and "MIS" not in su and "DIFFERENT" not in su:
        return f'<span class="pill pill-match">✓ {status}</span>'
    if "MISMATCH" in su or "DIFFERENT" in su or "NO MATCH" in su:
        return f'<span class="pill pill-mismatch">✗ {status}</span>'
    if "MISSING" in su:
        return f'<span class="pill pill-missing">⚠ {status}</span>'
    return f'<span class="pill pill-na">{status}</span>'
 
 
def pct(v):
    try:
        return f"{float(v):.2%}"
    except Exception:
        return "—"


def fmt_fraction_as_pct(v, *, ndigits: int = 3):
    """Format a stored fraction (e.g. 0.02275) for display as a percent; — when missing."""
    if v is None or (isinstance(v, float) and pd.isna(v)):
        return "—"
    try:
        return f"{float(v):.{ndigits}%}"
    except (TypeError, ValueError):
        return "—"
 
 
def fmt_date(v):
    try:
        return pd.to_datetime(v).strftime("%m/%d/%y")
    except Exception:
        return "—"


def _col(row, *keys):
    """First non-null column (backend renames)."""
    for k in keys:
        if k in row.index and pd.notna(row.get(k)):
            return row.get(k)
    return None


def fmt_num_plain(v):
    try:
        return f"{float(v):,.0f}"
    except Exception:
        return "—"


def fmt_opt_text(v):
    """Optional index-style fields: show em dash when blank (not empty string)."""
    if v is None or pd.isna(v):
        return "—"
    s = str(v).strip()
    return s if s else "—"


def _acore_source_type_family(raw) -> str:
    """Normalize ACORE `Source` / Source Type display to a single type family token.

    Examples: ``Subline | Bank of America`` → ``Subline``; ``Repo`` → ``Repo``.
    Blank / NaN → empty string (caller excludes from distinct-family logic).
    """
    if raw is None or (isinstance(raw, float) and pd.isna(raw)):
        return ""
    s = str(raw).strip()
    if not s or s.upper() in ("NAN", "NONE", "<NA>", "NAT"):
        return ""
    if "|" in s:
        s = s.split("|", 1)[0].strip()
    return s


def derive_liability_type_for_filter(row: pd.Series) -> str:
    """Best-effort liability type for UI filtering from available row fields."""
    # Prefer explicit M61 type when present in schema.
    explicit_keys = ("Liability Type (M61)", "Liability Type")
    explicit_present = any(k in row.index for k in explicit_keys)
    for k in explicit_keys:
        if k in row.index and pd.notna(row.get(k)):
            s = str(row.get(k)).strip()
            if s:
                return s

    # If explicit type fields are present but blank, keep blank (do not infer from Source).
    if explicit_present:
        return ""

    # Fallback only when explicit type fields are not present in this dataset version.
    return _acore_source_type_family(row.get("Source"))


def to_excel_bytes(df_recon, primary_file_type: str):
    df_recon = normalize_recon_fund_for_output(df_recon)
    wb = build_workbook(df_recon, primary_file_type=primary_file_type)
    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf.read()


def filter_recon_to_primary_file_rows(df: pd.DataFrame) -> pd.DataFrame:
    """Scoped UI rows for selected-fund view.

    Keep full outer-merge visibility (Both / primary-only / M61-only).
    """
    if df is None or df.empty:
        return df.copy() if df is not None else pd.DataFrame()
    return df.copy()


def filter_recon_scoped_to_business_lines(df: pd.DataFrame, run_primary: str) -> pd.DataFrame:
    """Selected Fund view: only rows in the chosen fund scope.

    Keeps side-by-side behavior for that fund (Both / primary-only / M61-only),
    and excludes rows from other funds.
    """
    base = filter_recon_to_primary_file_rows(df)
    if base.empty:
        return base
    return filter_recon_to_selected_fund(base, run_primary)


def _current_upload_signature(file_a_upload, file_b_upload, mapping_upload):
    """Track uploads only (not dropdown) so changing type does not auto-rerun."""
    if not file_a_upload or not file_b_upload:
        return None
    return (
        file_a_upload.name,
        getattr(file_a_upload, "size", None),
        file_b_upload.name,
        getattr(file_b_upload, "size", None),
        mapping_upload.name if mapping_upload else None,
        getattr(mapping_upload, "size", None) if mapping_upload else None,
    )


def _reset_table_filter_state() -> None:
    """Clear persisted table/grid filter widget state after each successful run."""
    prefixes = (
        "recon_tbl_primary_ms_",
        "recon_tbl_adv_ms_",
        "recon_tbl_sort_",
    )
    for k in list(st.session_state.keys()):
        if isinstance(k, str) and k.startswith(prefixes):
            del st.session_state[k]
    st.session_state["recon_hide_blank_cols"] = False
    st.session_state["recon_deal_pick"] = "All deals"


def run_reconciliation_for_selection(
    file_a_upload,
    file_b_upload,
    primary_file_type: str,
    mapping_upload=None,
):
    with st.spinner("Running reconciliation…"):
        with tempfile.TemporaryDirectory() as tmpdir:
            path_a = os.path.join(tmpdir, "liability.xlsx")
            path_b = os.path.join(tmpdir, "primary_model.xlsm")
            path_map = None
            with open(path_a, "wb") as f:
                f.write(file_a_upload.getbuffer())
            with open(path_b, "wb") as f:
                f.write(file_b_upload.getbuffer())
            if primary_file_type == "AOC II" and mapping_upload:
                path_map = os.path.join(tmpdir, "liability_to_cre_mapping.xlsx")
                with open(path_map, "wb") as f:
                    f.write(mapping_upload.getbuffer())
            df_recon, _, df_excluded_type = reconcile(
                path_a,
                path_b,
                primary_file_type=primary_file_type,
                mapping_path=path_map,
                uploaded_primary_filename=file_b_upload.name,
            )
            df_recon = normalize_recon_fund_for_output(df_recon)
            st.session_state["df_recon"] = df_recon
            st.session_state["df_excluded_by_liability_type"] = (
                df_excluded_type.copy() if df_excluded_type is not None else pd.DataFrame()
            )
            st.session_state["primary_file_type"] = primary_file_type
            st.session_state["primary_upload_name"] = file_b_upload.name
            st.session_state["excel_bytes"] = to_excel_bytes(df_recon, primary_file_type)
            # Persist exact download names from this successful run context.
            st.session_state["last_run_excel_name"] = build_output_filename(
                primary_file_type, "xlsx", uploaded_filename=file_b_upload.name
            )
            st.session_state["last_run_csv_name"] = build_output_filename(
                primary_file_type, "csv", uploaded_filename=file_b_upload.name
            )
            # Last successful run context for stale-state checks.
            st.session_state["last_successful_upload_signature"] = _current_upload_signature(
                file_a_upload, file_b_upload, mapping_upload
            )
            # Defer status reset to pre-widget stage (Streamlit-safe session_state mutation).
            st.session_state["reset_status_filters"] = True
            _reset_table_filter_state()
 
 
# --------------------------------------------------
# MAIN CONTENT
# --------------------------------------------------
has_required_uploads = file_a_upload and file_b_upload and m61_file_valid

manual_run_requested = st.button("▶  Run Reconciliation", type="primary")
upload_signature = _current_upload_signature(file_a_upload, file_b_upload, mapping_upload)
last_success_sig = st.session_state.get("last_successful_upload_signature")
auto_run_requested = bool(
    has_required_uploads and run_on_upload and upload_signature and upload_signature != last_success_sig
)

if has_required_uploads and (manual_run_requested or auto_run_requested):
    try:
        run_reconciliation_for_selection(
            file_a_upload=file_a_upload,
            file_b_upload=file_b_upload,
            primary_file_type=primary_file_type,
            mapping_upload=mapping_upload,
        )
    except PrimaryFileSchemaError as e:
        st.error(
            f"The **{e.primary_type}** workbook is missing required columns for this template."
        )
        st.markdown("**Missing or unmapped:**")
        for line in e.missing:
            st.markdown(f"- `{line}`")
        st.stop()
    except Exception as e:
        st.error(f"Reconciliation failed: {e}")
        st.stop()
 
# ---- Display results if available ----
if "df_recon" in st.session_state:
    df_recon = st.session_state["df_recon"]
    df_excluded_by_type = st.session_state.get("df_excluded_by_liability_type", pd.DataFrame())
    run_primary = st.session_state.get("primary_file_type", primary_file_type)
    run_primary_upload_name = st.session_state.get("primary_upload_name")
    run_excel_name = st.session_state.get(
        "last_run_excel_name",
        build_output_filename(run_primary, "xlsx", uploaded_filename=run_primary_upload_name),
    )
    run_csv_name = st.session_state.get(
        "last_run_csv_name",
        build_output_filename(run_primary, "csv", uploaded_filename=run_primary_upload_name),
    )
    pl_run = get_primary_config(run_primary)
    col_tag = pl_run["excel_primary_column_suffix"]
    run_primary_label = ui_primary_label(run_primary)
    primary_missing_scope_lbl = primary_scope_label_for_missing_banner(
        run_primary_upload_name, run_primary
    )
    is_stale_selection = run_primary != primary_file_type

    if is_stale_selection:
        st.warning(
            f"These results were generated using **{run_primary_label}**. "
            f"Your current selection is **{selected_ui_label}**. "
            "Please rerun reconciliation to refresh the results."
        )
        rerun_col, _ = st.columns([2, 5])
        with rerun_col:
            rerun_requested = st.button("🔄 Re-run Reconciliation")
        if rerun_requested:
            if not has_required_uploads:
                st.warning("Please upload both files before rerunning.")
            else:
                try:
                    run_reconciliation_for_selection(
                        file_a_upload=file_a_upload,
                        file_b_upload=file_b_upload,
                        primary_file_type=primary_file_type,
                        mapping_upload=mapping_upload,
                    )
                    st.rerun()
                except PrimaryFileSchemaError as e:
                    st.error(
                        f"The **{e.primary_type}** workbook is missing required columns for this template."
                    )
                    st.markdown("**Missing or unmapped:**")
                    for line in e.missing:
                        st.markdown(f"- `{line}`")
                    st.stop()
                except Exception as e:
                    st.error(f"Reconciliation failed: {e}")
                    st.stop()

    st.markdown('<div class="section-label">Deal filter</div>', unsafe_allow_html=True)
    df_all = df_recon.copy()
    deal_names = (
        sorted(df_all["Deal Name"].dropna().astype(str).unique().tolist())
        if "Deal Name" in df_all.columns
        else []
    )
    deal_options = ["All deals"] + deal_names
    deal_pick = st.selectbox(
        "Deal name",
        options=deal_options,
        index=0,
        key="recon_deal_pick",
        help="Type in the box to jump to a deal (Streamlit search). Choose **All deals** to clear.",
    )

    # Scoped subset: not In M61-only, then Fin Inpt–anchored (deal + ACP effective date + note + facility).
    df_scoped = filter_recon_scoped_to_business_lines(df_all, run_primary)
    in_scope_ix = set(df_scoped.index)
    if scope_mode == "Selected Fund Only":
        df_view = df_all.loc[df_all.index.isin(in_scope_ix)].copy()
        st.info(
            f"Scoped rows: **{run_primary_label}** fund scope only "
        )
    else:
        df_view = df_all.copy()
        st.caption(
            "**All Results:** full reconciliation output. **Selected Fund Only:** "
            f"subset to **{run_primary_label}** fund scope (see Scope info)."
        )

    # Apply status + deal filters on the current view.
    status_filter = []
    if st.session_state.get("filter_status_match", True):
        status_filter.append("MATCH")
    if st.session_state.get("filter_status_missing", True):
        status_filter.append("MISSING")
    if st.session_state.get("filter_status_mismatch", True):
        status_filter.append("MISMATCH")
    if status_filter:
        df_view = df_view[df_view["recon_status"].isin(status_filter)]
    else:
        df_view = df_view.iloc[0:0]
    if deal_pick and deal_pick != "All deals":
        df_view = df_view[df_view["Deal Name"] == deal_pick]

    # TEMP validation only (remove when done debugging fund-scope vs primary-key-scope).
    _val_cols = [
        c
        for c in (
            "Fund",
            "Deal Name",
            "Facility",
            "Financial Line",
            "Source Indicator",
            "recon_status",
        )
        if c in df_recon.columns
    ]

    def _sample_df(df: pd.DataFrame) -> pd.DataFrame:
        if df is None or df.empty:
            return df
        use = [c for c in _val_cols if c in df.columns]
        return df.head(10)[use] if use else df.head(10)

    with st.expander("TEMP: validation — df_recon / df_scoped / df_view", expanded=False):
        st.caption(
            "Compare **df_recon** (full recon) vs **df_scoped** (Fin Inpt–anchored scoped subset) "
            "vs **df_view** (after Scope + status + deal filters)."
        )
        v1, v2, v3 = st.columns(3)
        with v1:
            st.metric("df_recon rows", int(len(df_recon)))
        with v2:
            st.metric("df_scoped rows", int(len(df_scoped)))
        with v3:
            st.metric("df_view rows (displayed)", int(len(df_view)))
        st.markdown("**Sample: df_recon (10)**")
        st.dataframe(_sample_df(df_recon), use_container_width=True, height=180)
        st.markdown("**Sample: df_scoped (10)**")
        st.dataframe(_sample_df(df_scoped), use_container_width=True, height=180)
        if "Source Indicator" in df_recon.columns:
            c1, c2, c3 = st.columns(3)
            with c1:
                st.caption("Source Indicator mix: df_recon")
                st.dataframe(
                    df_recon["Source Indicator"]
                    .fillna("<NA>")
                    .astype(str)
                    .value_counts(dropna=False)
                    .rename("rows")
                    .to_frame(),
                    use_container_width=True,
                    height=160,
                )
            with c2:
                st.caption("Source Indicator mix: df_scoped")
                st.dataframe(
                    df_scoped["Source Indicator"]
                    .fillna("<NA>")
                    .astype(str)
                    .value_counts(dropna=False)
                    .rename("rows")
                    .to_frame(),
                    use_container_width=True,
                    height=160,
                )
            with c3:
                st.caption("Source Indicator mix: df_view")
                st.dataframe(
                    df_view["Source Indicator"]
                    .fillna("<NA>")
                    .astype(str)
                    .value_counts(dropna=False)
                    .rename("rows")
                    .to_frame(),
                    use_container_width=True,
                    height=160,
                )

    with st.expander("Excluded by Liability Type (validation)", expanded=False):
        st.caption(
            "Rows present in uploaded M61 file but excluded before financing-line reconciliation "
            "because `Liability Type` is not one of `Repo` / `Non` / `Subline`."
        )
        ex_count = int(len(df_excluded_by_type)) if df_excluded_by_type is not None else 0
        st.metric("Excluded rows", ex_count)
        if ex_count == 0:
            st.info("No rows were excluded by the Liability Type filter for this run.")
        else:
            if "Liability Type" in df_excluded_by_type.columns:
                st.markdown("**Excluded count by Liability Type**")
                st.dataframe(
                    df_excluded_by_type["Liability Type"]
                    .fillna("<NA>")
                    .astype(str)
                    .value_counts(dropna=False)
                    .rename("rows")
                    .to_frame(),
                    use_container_width=True,
                    height=160,
                )
            show_cols = [
                c
                for c in (
                    "Deal Name",
                    "Liability Type",
                    "Liability Name",
                    "Liability Note",
                    "Effective Date",
                    "Exclusion Reason",
                )
                if c in df_excluded_by_type.columns
            ]
            st.markdown("**Sample excluded rows (first 200)**")
            st.dataframe(
                df_excluded_by_type.loc[:, show_cols].head(200),
                use_container_width=True,
                height=320,
            )
 
    # ---- Metric cards ----
    total    = len(df_view)
    n_match  = (df_view["recon_status"] == "MATCH").sum()
    n_miss   = (df_view["recon_status"] == "MISSING").sum()
    n_mismatch = (df_view["recon_status"] == "MISMATCH").sum()
    match_rate = (n_match / total * 100) if total else 0
 
    c1, c2, c3, c4 = st.columns(4)
    with c1:
        st.markdown(f"""
        <div class="metric-card mc-total">
          <div class="label">Total Records</div>
          <div class="value">{total}</div>
          <div class="sub">Rows in current view</div>
        </div>""", unsafe_allow_html=True)
    with c2:
        st.markdown(f"""
        <div class="metric-card mc-match">
          <div class="label">✓ Match</div>
          <div class="value">{n_match}</div>
          <div class="sub">{match_rate:.0f}% match rate</div>
        </div>""", unsafe_allow_html=True)
    with c3:
        st.markdown(f"""
        <div class="metric-card mc-missing">
          <div class="label">⚠ Missing</div>
          <div class="value">{n_miss}</div>
          <div class="sub">Not in M61 file</div>
        </div>""", unsafe_allow_html=True)
    with c4:
        st.markdown(f"""
        <div class="metric-card mc-mismatch">
          <div class="label">✗ Mismatch</div>
          <div class="value">{n_mismatch}</div>
          <div class="sub">Requires review</div>
        </div>""", unsafe_allow_html=True)
 
# --- TEMP DIAGNOSTICS (HIDDEN FOR CLEAN UI) ---
# diag = st.session_state.get("recon_diagnostics", {}) or {}
# with st.expander("Temporary diagnostics (M61 row flow)", expanded=False):
#     d1, d2, d3, d4 = st.columns(4)
#     with d1:
#         st.metric("M61 raw rows", int(diag.get("m61_raw_rows", 0)))
#     with d2:
#         st.metric("After fund filter", int(diag.get("m61_after_fund_filter_rows", 0)))
#     with d3:
#         st.metric(
#             "After Liability Type filter",
#             int(diag.get("m61_after_liability_type_filter_rows", 0)),
#         )
#     with d4:
#         st.metric(
#             "Final reconciliation rows",
#             int(diag.get("recon_output_rows", len(df_recon)))
#         )

#     lt_counts = diag.get("m61_liability_type_counts_after_filter", {}) or {}
#     st.caption("Liability Type counts after M61 filter")
#     lt1, lt2, lt3 = st.columns(3)
#     with lt1:
#         st.metric("Repo", int(lt_counts.get("Repo", 0)))
#     with lt2:
#         st.metric("Non", int(lt_counts.get("Non", 0)))
#     with lt3:
#         st.metric("Subline", int(lt_counts.get("Subline", 0)))

    st.markdown("<br>", unsafe_allow_html=True)
 
    # ---- Download + count ----
    col_dl1, col_dl2, col_count = st.columns([2, 2, 5])

    with col_dl1:
        st.download_button(
            label="⬇️ Download Excel",
            data=st.session_state["excel_bytes"],
            file_name=run_excel_name,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            disabled=is_stale_selection,
        )

    with col_dl2:
        csv_data = df_view.to_csv(index=False).encode("utf-8")

        st.download_button(
            label="⬇️ Download CSV",
            data=csv_data,
            file_name=run_csv_name,
            mime="text/csv",
            disabled=is_stale_selection,
        )
    if is_stale_selection:
        st.caption("Downloads are disabled until you rerun with the current selection.")
    
    # ---- Tabs ----
    tab1, tab2 = st.tabs(["  All Results  ", "  Deal Drilldown  "])
 
    with tab1:
        st.markdown('<div class="section-label">Record-by-Record Reconciliation</div>', unsafe_allow_html=True)
 
        if df_view.empty:
            st.info("No records match the current filters.")
        else:
            # Build display table (aligned with RECON_ORDERED_COLS / Excel export)
            display_rows = []
            for _, row in df_view.iterrows():
                ed_acp = _col(row, "Effective Date (ACP)", "Effective Date")
                adv_acp = _col(row, "Advance Rate (ACP)", "Advance Rate")
                sp_acp = _col(row, "Spread (ACP)", "Spread")
                und_acp = _col(row, "Undrawn Capacity (ACP)", "Current Undrawn Capacity")
                und_liab = _col(
                    row, "Undrawn Capacity (M61)", "Current Undrawn Capacity (M61)"
                )

                def _status_display(v):
                    if v is None:
                        return ""
                    try:
                        if isinstance(v, float) and pd.isna(v):
                            return ""
                    except (TypeError, ValueError):
                        pass
                    s = str(v).strip()
                    if "MISSING" not in s.upper():
                        return s
                    return format_missing_status_display(
                        s, primary_scope_label=primary_missing_scope_lbl
                    )

                rec = {
                    "Fund": "" if pd.isna(row.get("Fund")) else str(row.get("Fund")),
                    "Deal Name": row.get("Deal Name", ""),
                    "Facility": row.get("Facility", ""),
                    "Financial Line": row.get("Financial Line", ""),
                    "Source Type (ACORE)": row.get("Source", ""),
                    "Liability Type (M61)": derive_liability_type_for_filter(row),
                    f"Eff Date ({col_tag})": fmt_date(ed_acp),
                    "Eff Date (M61)": fmt_date(row.get("Effective Date (M61)")),
                    f"Pledge Date ({col_tag})": fmt_date(
                        _col(row, "Pledge Date (ACP)", "Pledge Date")
                    ),
                    "Pledge Date (M61)": fmt_date(row.get("Pledge Date (M61)")),
                    f"Adv Rate ({col_tag})": pct(adv_acp),
                    "Adv Rate (M61)": pct(row.get("Advance Rate (M61)")),
                    "Target Adv Rate": pct(row.get("Target Advance Rate (M61)")),
                    f"Spread ({col_tag})": pct(sp_acp),
                    "Spread (M61)": pct(row.get("Spread (M61)")),
                    f"Undrawn ({col_tag})": fmt_num_plain(und_acp),
                    "Undrawn (M61)": fmt_num_plain(und_liab),
                    f"Index Floor ({col_tag})": fmt_fraction_as_pct(
                        row.get("Index Floor (ACP)"), ndigits=3
                    ),
                    "Index Floor (M61)": fmt_fraction_as_pct(
                        row.get("Index Floor (M61)"), ndigits=3
                    ),
                    f"Index Name ({col_tag})": fmt_opt_text(row.get("Index Name (ACP)")),
                    "Index Name (M61)": fmt_opt_text(row.get("Index Name (M61)")),
                    f"Recourse % ({col_tag})": pct(row.get("Recourse % (ACP)")),
                    "Recourse % (M61)": pct(row.get("Recourse % (M61)")),
                    "Source Status": row.get("Source Indicator", ""),
                    "Adv Rate Status": _status_display(row.get("Advance Rate Status", "")),
                    "Spread Status": _status_display(row.get("Spread Status", "")),
                    "Eff Date Status": _status_display(row.get("Effective Date Status", "")),
                    "Pledge Date Status": _status_display(row.get("Pledge Date Status", "")),
                    "Undrawn Capacity Status": _status_display(
                        row.get("Undrawn Capacity Status", "")
                    ),
                    "Index Floor Status": row.get("Index Floor Status", ""),
                    "Index Name Status": row.get("Index Name Status", ""),
                    "Recourse % Status": row.get("Recourse % Status", ""),
                    "Recon Status": _status_display(row.get("recon_status", "")),
                }
                display_rows.append(rec)

            df_display = pd.DataFrame(display_rows)

            # Option pools for **Source Type (ACORE)** / **Liability Type (M61)** table filters:
            # - Selected Fund Only → rows in ``df_all`` whose index is in ``in_scope_ix``
            #   (same fund scope as the initial ``df_view`` *before* status/deal narrowing).
            # - All Results → full ``df_all``.
            # Intentionally excludes sidebar recon-status checkboxes and the Deal picker so
            # type dropdowns are scope-aware but not over-reduced vs. the displayed row set.
            if scope_mode == "Selected Fund Only":
                df_type_opts_base = df_all.loc[df_all.index.isin(in_scope_ix)].copy()
            else:
                df_type_opts_base = df_all.copy()
            src_type_opt_series = (
                df_type_opts_base["Source"].map(_acore_source_type_family)
                if not df_type_opts_base.empty and "Source" in df_type_opts_base.columns
                else pd.Series(dtype="object")
            )
            liability_type_opt_series = (
                df_type_opts_base.apply(derive_liability_type_for_filter, axis=1)
                if not df_type_opts_base.empty
                else pd.Series(dtype="object")
            )

            # --- Column visibility (display-only; does not modify df_recon / df_view) ---
            all_display_cols = list(df_display.columns)
            col_sig = (tuple(all_display_cols), str(col_tag))
            if st.session_state.get("_recon_display_col_sig") != col_sig:
                # Reset per-column checkbox keys when display schema changes.
                for k in list(st.session_state.keys()):
                    if isinstance(k, str) and k.startswith("rcv_col_"):
                        del st.session_state[k]
                st.session_state["_recon_display_col_sig"] = col_sig

            def _col_vis_key(i: int, col_name: str) -> str:
                safe = re.sub(r"\W+", "_", col_name)[:60]
                return f"rcv_col_{i}_{safe}"

            for i, cname in enumerate(all_display_cols):
                k = _col_vis_key(i, cname)
                if k not in st.session_state:
                    st.session_state[k] = True

            st.markdown('<div class="section-label">Column Visibility</div>', unsafe_allow_html=True)

            btn_all, btn_clear = st.columns(2)
            with btn_all:
                # Label reflects prior run state; caption below uses fresh selection after expander.
                pre_sel = sum(
                    1
                    for i, cname in enumerate(all_display_cols)
                    if st.session_state.get(_col_vis_key(i, cname), True)
                )
                n_total_btn = len(all_display_cols)
                all_sel_btn = n_total_btn > 0 and pre_sel == n_total_btn
                none_sel_btn = n_total_btn > 0 and pre_sel == 0
                sel_lbl = "✔ All columns selected" if all_sel_btn else "Select All columns"
                if st.button(
                    sel_lbl,
                    key="recon_cols_select_all",
                    type="primary" if all_sel_btn else "secondary",
                ):
                    for i, cname in enumerate(all_display_cols):
                        st.session_state[_col_vis_key(i, cname)] = True
                    st.rerun()
            with btn_clear:
                if st.button(
                    "Clear All columns",
                    key="recon_cols_clear_all",
                    type="primary" if none_sel_btn else "secondary",
                ):
                    for i, cname in enumerate(all_display_cols):
                        st.session_state[_col_vis_key(i, cname)] = False
                    st.rerun()

            hide_blank_cols = st.checkbox(
                "Hide fully blank columns in current view",
                value=st.session_state.get("recon_hide_blank_cols", False),
                key="recon_hide_blank_cols",
                help="After filters/scoping, hide columns where every cell is blank, em dash, or N/A.",
            )

            with st.expander("Choose visible columns", expanded=False):
                n_chk_cols = 3
                n_rows = (len(all_display_cols) + n_chk_cols - 1) // n_chk_cols
                for r in range(n_rows):
                    chk_cols = st.columns(n_chk_cols)
                    for c in range(n_chk_cols):
                        idx = r * n_chk_cols + c
                        if idx >= len(all_display_cols):
                            break
                        cname = all_display_cols[idx]
                        k = _col_vis_key(idx, cname)
                        with chk_cols[c]:
                            st.checkbox(cname, key=k)

            visible_cols = [
                cname
                for i, cname in enumerate(all_display_cols)
                if st.session_state.get(_col_vis_key(i, cname), True)
            ]
            n_total = len(all_display_cols)
            n_sel = len(visible_cols)
            all_selected = n_total > 0 and n_sel == n_total
            if n_total:
                if all_selected:
                    st.caption(f"Showing all columns ({n_total}).")
                else:
                    st.caption(f"Showing {n_sel} / {n_total} columns.")

            def _cell_is_blankish(v) -> bool:
                if v is None or (isinstance(v, float) and pd.isna(v)):
                    return True
                s = str(v).strip()
                if not s:
                    return True
                su = s.upper()
                if su in ("N/A", "NAN", "NONE", "—", "-", "<NA>"):
                    return True
                return False

            df_table = df_display.loc[:, visible_cols] if visible_cols else df_display.iloc[:, 0:0]

            if hide_blank_cols and not df_table.empty:
                non_blank = []
                for c in df_table.columns:
                    if not df_table[c].map(_cell_is_blankish).all():
                        non_blank.append(c)
                df_table = df_table.loc[:, non_blank] if non_blank else df_table.iloc[:, 0:0]

            df_table_view = df_table.copy()
            if not df_table_view.empty:
                st.markdown(
                    '<div class="section-label">Table filters</div>',
                    unsafe_allow_html=True,
                )

                # Primary/basic filters (single horizontal row)
                primary_filter_cols = [
                    "Fund",
                    "Deal Name",
                    "Facility",
                    "Financial Line",
                    "Source Type (ACORE)",
                    "Liability Type (M61)",
                    "Source Status",
                ]
                auto_fund_value = ""
                auto_source_type_filter_vals = None  # list[str] | None
                auto_liability_type_filter_val = None  # str | None
                if scope_mode == "Selected Fund Only" and not df_display.empty:
                    if "Fund" in df_display.columns:
                        fund_vals = (
                            df_display["Fund"]
                            .fillna("")
                            .astype(str)
                            .str.strip()
                        )
                        fund_vals = fund_vals[fund_vals.ne("")]
                        if not fund_vals.empty:
                            # Most-common visible scoped fund for this run.
                            auto_fund_value = str(fund_vals.value_counts().index[0]).strip()
                    # Single ACORE source/type family (normalized before ``|``): prefill Source
                    # filter with that family token. Use the same scope basis as type dropdowns
                    # (``df_type_opts_base``), not deal/status-narrowed ``df_display``.
                    if "Source" in df_type_opts_base.columns and not df_type_opts_base.empty:
                        fam_series = df_type_opts_base["Source"].map(_acore_source_type_family)
                        fam_nonempty = fam_series[fam_series.ne("")]
                        if not fam_nonempty.empty:
                            distinct_families = set(fam_nonempty.unique().tolist())
                            if len(distinct_families) == 1:
                                sole_family = next(iter(distinct_families))
                                auto_source_type_filter_vals = [sole_family]
                                lt = (
                                    liability_type_opt_series.fillna("")
                                    .astype(str)
                                    .str.strip()
                                )
                                lt = lt[lt.ne("")]
                                if not lt.empty and lt.nunique(dropna=False) == 1:
                                    auto_liability_type_filter_val = str(lt.iloc[0]).strip()
                pf_ui_cols = st.columns(len(primary_filter_cols))
                for i, fc in enumerate(primary_filter_cols):
                    with pf_ui_cols[i]:
                        # Source/Liability: options from scope-mode basis (fund subset or full
                        # ``df_all``), not from ``df_display`` / post-deal / post-status ``df_view``.
                        if fc == "Source Type (ACORE)":
                            opts_src = src_type_opt_series
                        elif fc == "Liability Type (M61)":
                            opts_src = liability_type_opt_series
                        else:
                            opts_src = (
                                df_display[fc]
                                if fc in df_display.columns
                                else pd.Series(dtype="object")
                            )
                        opts = sorted(
                            {
                                str(v).strip()
                                for v in opts_src.fillna("").astype(str).tolist()
                                if str(v).strip()
                            }
                        )
                        base_key = f"recon_tbl_primary_ms_{re.sub(r'\\W+', '_', fc)}_{col_tag}"
                        # Selected Fund Only: Fund = most common fund; Source/Liability = safe
                        # prefill when scoped ACORE-side data implies one type family (Source
                        # normalized before ``|``) and, for Liability, one distinct resolved type.
                        if fc == "Fund" and scope_mode == "Selected Fund Only":
                            ms_key = f"{base_key}_auto_scope"
                            # Only prefill when empty/missing; allow manual edits in this mode.
                            if auto_fund_value and auto_fund_value in opts:
                                if ms_key not in st.session_state or not st.session_state.get(ms_key):
                                    st.session_state[ms_key] = [auto_fund_value]
                        elif fc == "Source Type (ACORE)":
                            # Dedicated key so legacy raw-string widget state does not clash
                            # with family-token options.
                            ms_key = f"{base_key}_dispopts"
                            if scope_mode == "Selected Fund Only" and auto_source_type_filter_vals:
                                if all(v in opts for v in auto_source_type_filter_vals):
                                    if ms_key not in st.session_state or not st.session_state.get(
                                        ms_key
                                    ):
                                        st.session_state[ms_key] = list(auto_source_type_filter_vals)
                        elif fc == "Liability Type (M61)":
                            ms_key = f"{base_key}_dispopts"
                            if (
                                scope_mode == "Selected Fund Only"
                                and auto_liability_type_filter_val
                                and auto_liability_type_filter_val in opts
                            ):
                                if ms_key not in st.session_state or not st.session_state.get(ms_key):
                                    st.session_state[ms_key] = [auto_liability_type_filter_val]
                        else:
                            ms_key = base_key
                        selected_vals = st.multiselect(
                            fc,
                            options=opts,
                            default=[],
                            key=ms_key,
                            help="Empty = show all values.",
                        )
                        if selected_vals:
                            allow = set(selected_vals)
                            if fc == "Source Type (ACORE)" and "Source Type (ACORE)" in df_display.columns:
                                fam = df_display["Source Type (ACORE)"].map(_acore_source_type_family)
                                keep_idx = df_display.loc[fam.isin(allow)].index
                            elif fc in df_display.columns:
                                keep_idx = df_display[
                                    df_display[fc].fillna("").astype(str).str.strip().isin(allow)
                                ].index
                            else:
                                keep_idx = df_display.index[:0]
                            df_table_view = df_table_view.loc[
                                df_table_view.index.intersection(keep_idx)
                            ]

                # TEMP: Advanced Filters expander disabled — restore by uncommenting the block below
                # and removing the standalone ``sort_cols_list`` block that follows.
                # # Advanced filters retain extra capabilities.
                # with st.expander("Advanced Filters", expanded=False):
                #     advanced_filter_cols = [
                #         "Adv Rate Status",
                #         "Spread Status",
                #         "Eff Date Status",
                #         "Undrawn Capacity Status",
                #         "Index Floor Status",
                #         "Index Name Status",
                #         "Recourse % Status",
                #         "Pledge Date Status",
                #     ]
                #     adv_cols_present = [c for c in advanced_filter_cols if c in df_display.columns]
                #     if adv_cols_present:
                #         adv_ui_cols = st.columns(min(len(adv_cols_present), 3))
                #         for i, fc in enumerate(adv_cols_present):
                #             with adv_ui_cols[i % len(adv_ui_cols)]:
                #                 adv_opts = sorted(
                #                     {
                #                         str(v).strip()
                #                         for v in df_display[fc].fillna("").astype(str).tolist()
                #                         if str(v).strip()
                #                     }
                #                 )
                #                 adv_sel = st.multiselect(
                #                     fc,
                #                     options=adv_opts,
                #                     default=[],
                #                     key=f"recon_tbl_adv_ms_{re.sub(r'\\W+', '_', fc)}_{col_tag}",
                #                     help="Optional additional filter; empty = all.",
                #                 )
                #                 if adv_sel:
                #                     keep_idx = df_display[
                #                         df_display[fc]
                #                         .fillna("")
                #                         .astype(str)
                #                         .str.strip()
                #                         .isin(set(adv_sel))
                #                     ].index
                #                     df_table_view = df_table_view.loc[
                #                         df_table_view.index.intersection(keep_idx)
                #                     ]

                sort_cols_list = [c for c in df_table_view.columns]
                if sort_cols_list:
                    s1, s2 = st.columns(2)
                    with s1:
                        sort_by = st.selectbox(
                            "Sort by",
                            options=["(none)"] + sort_cols_list,
                            index=0,
                            key=f"recon_tbl_sort_{col_tag}",
                            help="Applies after filters; uses string order if types differ.",
                        )
                    with s2:
                        sort_asc = st.checkbox(
                            "Ascending",
                            value=True,
                            key=f"recon_tbl_sort_asc_{col_tag}",
                        )
                    if sort_by != "(none)":
                        try:
                            df_table_view = df_table_view.sort_values(
                                by=sort_by, ascending=sort_asc, na_position="last"
                            )
                        except TypeError:
                            t = df_table_view.copy()
                            t["_sort_tmp"] = t[sort_by].astype(str)
                            df_table_view = t.sort_values(
                                "_sort_tmp", ascending=sort_asc
                            ).drop(columns="_sort_tmp")

                n_after = len(df_table_view)
                n_before = len(df_table)
                if n_after != n_before:
                    st.caption(f"Table filters: showing **{n_after}** of **{n_before}** row(s).")

            if not visible_cols and not df_display.empty:
                st.info(
                    "No columns selected for display. Use the **Select All** button "
                    "(shows **✔ All columns selected** when active) or open **Choose visible columns**."
                )
            elif hide_blank_cols and df_table.empty and visible_cols and not df_display.empty:
                st.caption(
                    "No table to show: every selected column is blank for the current view "
                    "(try turning off **Hide fully blank columns** or select different columns)."
                )
            elif not df_table.empty and df_table_view.empty:
                st.caption(
                    "No rows match the **Table filters** above. Clear one or more multiselect filters to see data."
                )

            # Status columns: color only (layout / padding / alignment shared via set_properties).
            def color_status(val):
                v = str(val).strip().upper()
                if v in ("N/A", "", "—", "-", "NAN", "NONE"):
                    return ""
                if "DIFFERENT" in v or "MISMATCH" in v or "NO MATCH" in v:
                    return "background-color: #ffc7ce; color: #9c0006; font-weight: 600;"
                if "MATCH" in v and "MIS" not in v:
                    return "background-color: #c6efce; color: #375623; font-weight: 600;"
                if "MISSING" in v:
                    return "background-color: #ffeb9c; color: #7d6608; font-weight: 600;"
                return ""

            # Single ordered list: Spread Status is the reference column; all share identical UI.
            status_cols = [
                "Source Status",
                "Adv Rate Status",
                "Spread Status",
                "Eff Date Status",
                "Pledge Date Status",
                "Undrawn Capacity Status",
                "Index Floor Status",
                "Index Name Status",
                "Recourse % Status",
                "Recon Status",
            ]
            status_cols_visible = [c for c in status_cols if c in df_table_view.columns]

            # Same Streamlit width + Styler box for every status column; shorter header label
            # for Undrawn only so the column is not widened by a long title vs. peers.
            def _status_column_config(col_name: str):
                if col_name == "Undrawn Capacity Status":
                    return st.column_config.TextColumn(
                        "Undrawn status",
                        width="medium",
                        help="Undrawn Capacity Status",
                    )
                return st.column_config.TextColumn(width="medium")

            _status_col_cfg = {
                c: _status_column_config(c)
                for c in df_table_view.columns
                if isinstance(c, str)
                and (c.endswith(" Status") or c == "Recon Status")
            }

            if df_table_view.empty:
                styled = df_table_view.style
            else:
                styled = df_table_view.style
                if status_cols_visible:
                    styled = styled.map(color_status, subset=status_cols_visible)
                    # Uniform cell box for all status columns (padding, wrap, center — Spread baseline).
                    styled = styled.set_properties(
                        subset=status_cols_visible,
                        **{
                            "text-align": "center",
                            "vertical-align": "middle",
                            "white-space": "normal",
                            "word-wrap": "break-word",
                            "overflow-wrap": "break-word",
                            "box-sizing": "border-box",
                            "padding": "10px 10px",
                            "min-height": "3.25rem",
                            "line-height": "1.35",
                            "background-clip": "padding-box",
                        },
                    )
            _df_kwargs: dict = {
                "use_container_width": True,
                "height": 580,
            }
            if _status_col_cfg:
                _df_kwargs["column_config"] = _status_col_cfg
            st.dataframe(styled, **_df_kwargs)
 
    with tab2:
        st.markdown('<div class="section-label">Deal Drilldown</div>', unsafe_allow_html=True)
 
        deal_names = df_view["Deal Name"].dropna().unique().tolist()
        if not deal_names:
            st.info("No deals available for the current filters.")
            selected_deal = None
        else:
            selected_deal = st.selectbox("Select a deal", sorted(deal_names))
 
        if selected_deal:
            deal_rows = df_view[df_view["Deal Name"] == selected_deal]
 
            st.markdown(f"""
            <div class="info-box">
              <strong>{selected_deal}</strong> — {len(deal_rows)} effective date record(s) found
            </div>
            """, unsafe_allow_html=True)
 
            for _, row in deal_rows.iterrows():
                recon = str(row.get("recon_status", "")).upper()
                border_color = "#4caf50" if "MATCH" in recon and "MIS" not in recon else ("#f44336" if "MISMATCH" in recon else "#ffc107")
                ed_acp = _col(row, "Effective Date (ACP)", "Effective Date")
                adv_acp = _col(row, "Advance Rate (ACP)", "Advance Rate")
                sp_acp = _col(row, "Spread (ACP)", "Spread")
                und_acp = _col(row, "Undrawn Capacity (ACP)", "Current Undrawn Capacity")
                und_liab = _col(
                    row, "Undrawn Capacity (M61)", "Current Undrawn Capacity (M61)"
                )
                fund_lbl = "" if pd.isna(row.get("Fund")) else str(row.get("Fund"))

                def _deal_status_display(v):
                    if v is None:
                        return ""
                    try:
                        if isinstance(v, float) and pd.isna(v):
                            return ""
                    except (TypeError, ValueError):
                        pass
                    s = str(v).strip()
                    if "MISSING" not in s.upper():
                        return s
                    return format_missing_status_display(
                        s, primary_scope_label=primary_missing_scope_lbl
                    )

                st.markdown(f"""
                <div style='background:#fff; border-left:4px solid {border_color}; border-radius:8px;
                             padding:18px 22px; margin-bottom:14px;
                             box-shadow: 0 2px 6px rgba(0,0,0,0.06); max-width:100%;'>
                  <div style='display:flex; justify-content:space-between; align-items:flex-start; gap:16px; margin-bottom:12px; flex-wrap:wrap;'>
                    <div style='min-width:220px;'>
                      <span style='font-size:0.78rem; color:#666; font-family:monospace'>EFFECTIVE DATE ({col_tag})</span><br>
                      <span style='font-size:1.08rem; font-weight:700; color:#1a3a6c'>{fmt_date(ed_acp)}</span>
                      <div style='font-size:0.78rem; color:#5a7a99; margin-top:6px;'>Fund: <strong>{fund_lbl or "—"}</strong></div>
                    </div>
                    <div style='flex-shrink:0;'>{pill(_deal_status_display(row.get("recon_status", "")))}</div>
                  </div>
                  <div style='display:grid; grid-template-columns:repeat(2, minmax(280px, 1fr)); gap:12px;'>
                    <div style='background:#f8fafd; border-radius:6px; padding:12px 14px;'>
                      <div style='font-size:0.65rem; color:#888; text-transform:uppercase; letter-spacing:.06em'>Facility</div>
                      <div style='font-size:0.92rem; font-weight:600; color:#1a3a6c'>{row.get("Facility", "—")}</div>
                    </div>
                    <div style='background:#f8fafd; border-radius:6px; padding:12px 14px;'>
                      <div style='font-size:0.65rem; color:#888; text-transform:uppercase; letter-spacing:.06em'>Financial Line</div>
                      <div style='font-size:0.92rem; font-weight:600; color:#1a3a6c'>{row.get("Financial Line", "—")}</div>
                    </div>
                    <div style='background:#f8fafd; border-radius:6px; padding:12px 14px;'>
                      <div style='font-size:0.65rem; color:#888; text-transform:uppercase; letter-spacing:.06em'>Pledge Date ({col_tag})</div>
                      <div style='font-size:0.92rem; font-weight:600; color:#1a3a6c'>{fmt_date(_col(row, "Pledge Date (ACP)", "Pledge Date"))}</div>
                    </div>
                    <div style='background:#f8fafd; border-radius:6px; padding:12px 14px;'>
                      <div style='font-size:0.65rem; color:#888; text-transform:uppercase; letter-spacing:.06em'>Pledge Date (M61)</div>
                      <div style='font-size:0.92rem; font-weight:600; color:#1a3a6c'>{fmt_date(row.get("Pledge Date (M61)"))}</div>
                    </div>
                    <div style='background:#f8fafd; border-radius:6px; padding:12px 14px;'>
                      <div style='font-size:0.65rem; color:#888; text-transform:uppercase; letter-spacing:.06em'>Adv Rate ({col_tag})</div>
                      <div style='font-size:0.92rem; font-weight:600; color:#1a3a6c'>{pct(adv_acp)}</div>
                    </div>
                    <div style='background:#f8fafd; border-radius:6px; padding:12px 14px;'>
                      <div style='font-size:0.65rem; color:#888; text-transform:uppercase; letter-spacing:.06em'>Adv Rate (M61)</div>
                      <div style='font-size:0.92rem; font-weight:600; color:#1a3a6c'>{pct(row.get("Advance Rate (M61)"))}</div>
                    </div>
                    <div style='background:#f8fafd; border-radius:6px; padding:12px 14px;'>
                      <div style='font-size:0.65rem; color:#888; text-transform:uppercase; letter-spacing:.06em'>Target Adv Rate</div>
                      <div style='font-size:0.92rem; font-weight:600; color:#1a3a6c'>{pct(row.get("Target Advance Rate (M61)"))}</div>
                    </div>
                    <div style='background:#f8fafd; border-radius:6px; padding:12px 14px;'>
                      <div style='font-size:0.65rem; color:#888; text-transform:uppercase; letter-spacing:.06em'>Spread ({col_tag})</div>
                      <div style='font-size:0.92rem; font-weight:600; color:#1a3a6c'>{pct(sp_acp)}</div>
                    </div>
                    <div style='background:#f8fafd; border-radius:6px; padding:12px 14px;'>
                      <div style='font-size:0.65rem; color:#888; text-transform:uppercase; letter-spacing:.06em'>Spread (M61)</div>
                      <div style='font-size:0.92rem; font-weight:600; color:#1a3a6c'>{pct(row.get("Spread (M61)"))}</div>
                    </div>
                    <div style='background:#f8fafd; border-radius:6px; padding:12px 14px;'>
                      <div style='font-size:0.65rem; color:#888; text-transform:uppercase; letter-spacing:.06em'>Undrawn ({col_tag})</div>
                      <div style='font-size:0.92rem; font-weight:600; color:#1a3a6c'>{fmt_num_plain(und_acp)}</div>
                    </div>
                    <div style='background:#f8fafd; border-radius:6px; padding:12px 14px;'>
                      <div style='font-size:0.65rem; color:#888; text-transform:uppercase; letter-spacing:.06em'>Undrawn (M61)</div>
                      <div style='font-size:0.92rem; font-weight:600; color:#1a3a6c'>{fmt_num_plain(und_liab)}</div>
                    </div>
                    <div style='background:#f8fafd; border-radius:6px; padding:12px 14px;'>
                      <div style='font-size:0.65rem; color:#888; text-transform:uppercase; letter-spacing:.06em'>Index Floor ({col_tag})</div>
                      <div style='font-size:0.92rem; font-weight:600; color:#1a3a6c'>{fmt_fraction_as_pct(row.get("Index Floor (ACP)"), ndigits=3)}</div>
                    </div>
                    <div style='background:#f8fafd; border-radius:6px; padding:12px 14px;'>
                      <div style='font-size:0.65rem; color:#888; text-transform:uppercase; letter-spacing:.06em'>Index Floor (M61)</div>
                      <div style='font-size:0.92rem; font-weight:600; color:#1a3a6c'>{fmt_fraction_as_pct(row.get("Index Floor (M61)"), ndigits=3)}</div>
                    </div>
                    <div style='background:#f8fafd; border-radius:6px; padding:12px 14px;'>
                      <div style='font-size:0.65rem; color:#888; text-transform:uppercase; letter-spacing:.06em'>Index Name ({col_tag})</div>
                      <div style='font-size:0.92rem; font-weight:600; color:#1a3a6c'>{fmt_opt_text(row.get("Index Name (ACP)"))}</div>
                    </div>
                    <div style='background:#f8fafd; border-radius:6px; padding:12px 14px;'>
                      <div style='font-size:0.65rem; color:#888; text-transform:uppercase; letter-spacing:.06em'>Index Name (M61)</div>
                      <div style='font-size:0.92rem; font-weight:600; color:#1a3a6c'>{fmt_opt_text(row.get("Index Name (M61)"))}</div>
                    </div>
                    <div style='background:#f8fafd; border-radius:6px; padding:12px 14px;'>
                      <div style='font-size:0.65rem; color:#888; text-transform:uppercase; letter-spacing:.06em'>Recourse % ({col_tag})</div>
                      <div style='font-size:0.92rem; font-weight:600; color:#1a3a6c'>{pct(row.get("Recourse % (ACP)"))}</div>
                    </div>
                    <div style='background:#f8fafd; border-radius:6px; padding:12px 14px;'>
                      <div style='font-size:0.65rem; color:#888; text-transform:uppercase; letter-spacing:.06em'>Recourse % (M61)</div>
                      <div style='font-size:0.92rem; font-weight:600; color:#1a3a6c'>{pct(row.get("Recourse % (M61)"))}</div>
                    </div>
                  </div>
                  <div style='display:grid; grid-template-columns:repeat(auto-fit, minmax(200px, 1fr)); gap:8px; margin-top:10px;'>
                    <div style='font-size:0.74rem;'>Adv Rate: {pill(_deal_status_display(row.get("Advance Rate Status","")))}</div>
                    <div style='font-size:0.74rem;'>Spread: {pill(_deal_status_display(row.get("Spread Status","")))}</div>
                    <div style='font-size:0.74rem;'>Eff Date: {pill(_deal_status_display(row.get("Effective Date Status","")))}</div>
                    <div style='font-size:0.74rem;'>Undrawn: {pill(_deal_status_display(row.get("Undrawn Capacity Status","")))}</div>
                    <div style='font-size:0.74rem;'>Index Floor: {pill(row.get("Index Floor Status",""))}</div>
                    <div style='font-size:0.74rem;'>Index Name: {pill(row.get("Index Name Status",""))}</div>
                    <div style='font-size:0.74rem;'>Recourse %: {pill(row.get("Recourse % Status",""))}</div>
                    <div style='font-size:0.74rem;'>Pledge Date: {pill(_deal_status_display(row.get("Pledge Date Status","")))}</div>
                  </div>
                </div>
                """, unsafe_allow_html=True)
 
else:
    # Empty state
    st.markdown(
        f"""
    <div style='text-align:center; padding:60px 40px; background:#fff; border-radius:14px;
                border: 1.5px dashed #ccd9ea; color:#5a7a99;'>
      <div style='font-size:3rem; margin-bottom:16px;'>📂</div>
      <h3 style='color:#1a3a6c; font-weight:700; margin-bottom:8px'>Upload both files to begin</h3>
      <p style='font-size:0.88rem; max-width:420px; margin:0 auto; line-height:1.6;'>
        Choose the <strong>{selected_ui_label}</strong> primary fund template in the sidebar,
        then upload the matching <strong>Liquidity &amp; Earnings Model</strong> and the
        <strong>(In) M61 Relationship</strong> export.
      </p>
      <div style='display:flex; justify-content:center; gap:16px; margin-top:24px; flex-wrap:wrap;'>
        <div style='background:#e8f0fe; border-radius:8px; padding:10px 18px; font-size:0.78rem; color:#1a3a6c; font-weight:600;'>
          📊 Primary .xlsm / .xlsx
        </div>
        <div style='background:#e8f0fe; border-radius:8px; padding:10px 18px; font-size:0.78rem; color:#1a3a6c; font-weight:600;'>
          📄 (In) M61 .xlsx
        </div>
      </div>
    </div>
    """,
        unsafe_allow_html=True,
    )