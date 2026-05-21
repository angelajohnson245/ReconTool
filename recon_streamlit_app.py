"""
Financing Line Reconciliation Tool — Streamlit UI
Run with: streamlit run recon_streamlit_app.py
"""
import calendar
import gc
import hashlib
import html
import io
import os
import re
import sys
import tempfile
import unicodedata
from datetime import date, datetime
 
import pandas as pd
import streamlit as st
 
from recon_enhanced_output import (
    FILE_SOURCE_ACORE_ONLY,
    FILE_SOURCE_BOTH,
    FILE_SOURCE_M61_ONLY,
    PRIMARY_TYPE_FUND_CONFIG,
    STREAMLIT_PRIMARY_FILE_TYPES,
    M61_NOTE_CATEGORIES,
    PrimaryFileSchemaError,
    build_output_filename,
    build_workbook,
    canonical_primary_file_type,
    categorize_m61_note_category,
    categorize_m61_note_type,
    filter_recon_to_selected_fund,
    get_primary_config,
    get_last_recon_context,
    normalise_facility,
    normalise_text,
    normalize_recon_fund_for_output,
    reconcile,
    safe_str_strip,
    scope_label_for_primary_type,
    _coerce_numeric_value,
    _is_blank_for_compare,
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
    font-weight: 800;
    letter-spacing: 0.06em;
    font-size: 0.92rem;
    text-transform: uppercase;
    margin-top: 0.55rem;
    margin-bottom: 0.5rem;
    line-height: 1.3;
  }

  /* Subheaders: Show Records, Scope, M61 Note Category (below FILTERS / section titles) */
  .sidebar-subheader {
    color: #e8f4fd !important;
    font-size: 0.94rem !important;
    font-weight: 700 !important;
    letter-spacing: 0.03em !important;
    line-height: 1.4;
    margin: 0.55rem 0 0.5rem 0 !important;
  }
  .sidebar-subheader-row {
    display: flex;
    align-items: baseline;
    gap: 0.35rem;
    margin: 0.55rem 0 0.5rem 0 !important;
  }
  .sidebar-subheader-row .sidebar-subheader-inline {
    color: #e8f4fd !important;
    font-size: 0.94rem !important;
    font-weight: 700 !important;
    letter-spacing: 0.03em !important;
    line-height: 1.4;
    margin: 0 !important;
  }
  .sidebar-subheader-hint {
    font-size: 0.8rem !important;
    opacity: 0.82;
    cursor: help;
    color: #8fb8d8 !important;
  }

  /* Extra vertical separation between Show Records / Scope / M61 blocks */
  .sidebar-section-gap {
    display: block;
    height: 0;
    margin: 0.9rem 0 0.35rem 0;
  }
  .sidebar-subheader--after-gap {
    margin-top: 0.15rem !important;
  }

  /* Scope radio: tuck under subheader with same vertical rhythm as other groups */
  [data-testid="stSidebar"] .stRadio {
    margin-top: 0 !important;
  }

  /* M61 Note Category (and other sidebar selects): clearer control surface */
  [data-testid="stSidebar"] div[data-baseweb="select"] > div {
    background-color: rgba(22, 48, 74, 0.92) !important;
    border: 1px solid rgba(143, 184, 216, 0.65) !important;
    border-radius: 8px !important;
    min-height: 2.65rem !important;
    padding-left: 0.65rem !important;
    padding-right: 0.5rem !important;
    box-shadow: inset 0 1px 0 rgba(255, 255, 255, 0.06);
  }
  [data-testid="stSidebar"] div[data-baseweb="select"] > div:hover {
    border-color: rgba(168, 205, 232, 0.85) !important;
    background-color: rgba(28, 58, 90, 0.95) !important;
  }
  [data-testid="stSidebar"] div[data-baseweb="select"] [aria-selected="true"],
  [data-testid="stSidebar"] div[data-baseweb="select"] span {
    color: #e8f4fd !important;
  }

  /* Sidebar: make “Advanced filters” expander easier to notice */
  [data-testid="stSidebar"] [data-testid="stExpander"] {
    background: rgba(26, 58, 92, 0.55) !important;
    border: 1px solid rgba(143, 184, 216, 0.5) !important;
    border-radius: 8px;
    padding: 0.2rem 0.35rem 0.45rem;
    margin-top: 0.35rem;
    margin-bottom: 0.35rem;
    box-shadow: 0 1px 4px rgba(0, 0, 0, 0.22);
  }
  [data-testid="stSidebar"] [data-testid="stExpander"] summary,
  [data-testid="stSidebar"] [data-testid="stExpander"] summary * {
    color: #e8f4fd !important;
    font-weight: 600 !important;
    letter-spacing: 0.03em !important;
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
  .mc-needs   { background: #fff3e0; border-color: #ffcc80; color: #bf360c; }
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
  .pill-match      { background: #c6efce; color: #375623; }
  .pill-missing    { background: #ffeb9c; color: #7d6608; }
  .pill-incomplete { background: #ffeb9c; color: #7d6608; }
  .pill-mismatch   { background: #ffc7ce; color: #9c0006; }
  .pill-na         { background: #e0e0e0; color: #555; }
 
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
            return PRIMARY_TYPE_FUND_CONFIG["ACP III"]["scope_label"]
        if re.search(r"\bACP\s+II\b", name):
            return PRIMARY_TYPE_FUND_CONFIG["ACP II"]["scope_label"]
        if re.search(r"\bACP\s+I\b(?!\s*I)", name):
            return PRIMARY_TYPE_FUND_CONFIG["ACP I"]["scope_label"]
        if re.search(r"\bAOC\s+II\b", name):
            return PRIMARY_TYPE_FUND_CONFIG["AOC II"]["scope_label"]
        if re.search(r"\bAOC\s+I\b(?!\s*I)", name):
            return PRIMARY_TYPE_FUND_CONFIG["AOC I"]["scope_label"]
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
    if su in ("MISSING FROM M61", "MISSING FROM ACORE", "MISSING FROM BOTH"):
        return s
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
        return "ACP III"
    if re.search(r"\bACP\s+II\b", name):
        return "ACP II"
    if re.search(r"\bACP\s+I\b(?!\s*I)", name):
        return "ACP I"
    if re.search(r"\bAOC\s+II\b", name):
        return "AOC II"
    if re.search(r"\bAOC\s+I\b(?!\s*I)", name):
        return "AOC I"
    return None


def primary_filename_incompatible_acp_ii_vs_iii(
    uploaded_filename: str | None, primary_file_type: str
) -> tuple[bool, str]:
    """ACP II vs ACP III workbooks use different Fin Inpt layouts — block run if filename disagrees with sidebar."""
    inf = infer_primary_type_from_filename(uploaded_filename)
    if not inf:
        return False, ""
    sel = canonical_primary_file_type(primary_file_type)
    if inf == "ACP II" and sel == "ACP III":
        return (
            True,
            "**Primary type mismatch:** The uploaded file name looks like **ACP II**, but the sidebar is set to "
            "**ACP III** (different Fin Inpt sheet/columns). Choose **ACP II** in the sidebar or upload an "
            "**ACP III** Liquidity & Earnings model, then run again.",
        )
    if inf == "ACP III" and sel == "ACP II":
        return (
            True,
            "**Primary type mismatch:** The uploaded file name looks like **ACP III**, but the sidebar is set to "
            "**ACP II**. Choose **ACP III** in the sidebar or upload an **ACP II** workbook, then run again.",
        )
    return False, ""


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


def normalize_m61_note_category_label(v: object) -> str:
    """Normalize category labels for consistent sidebar option/filter matching."""
    try:
        if v is pd.NA:
            v = ""
    except (TypeError, ValueError):
        pass
    s = "" if v is None else str(v)
    s = unicodedata.normalize("NFKC", s)
    s = s.replace("\ufeff", "").strip()
    s = re.sub(r"\s+", " ", s).strip().lower()
    if s == "all":
        return "all"
    if s in ("", "nan", "none", "<na>", "n/a", "na", "unknown", "other"):
        return "other"
    # Subline: match "subline", "sub line", unicode variants (after NFKC), etc.
    if re.search(r"\bsub[\s-]*line\b", s) or re.fullmatch(r"sub[\s-]*line", s):
        return "subline"
    if s in ("financing", "repo", "sale", "non", "clo", "sub debt", "tbd"):
        return "financing"
    if s in ("equity/fund", "eq/fund", "equity", "fund"):
        return "other"
    if "whole loan" in s or "wholeloan" in s or "wl-cpace" in s or "wlcpace" in s:
        return "other"
    return "other"


def m61_note_category_series_for_ui(
    df: pd.DataFrame, *, primary_file_type: str | None = None
) -> pd.Series:
    """Row-level normalized M61 note categories for sidebar + table filtering.

    Uses ``categorize_m61_note_category`` (same as ``reconcile``) when ``primary_file_type`` and
    ``Liability Type (M61 Raw)`` are available, so the filter matches displayed M61 context even if
    ``M61 Note Category`` were stale (e.g. before related-M61 surfacing refreshed liability type).
    Falls back to the stored ``M61 Note Category`` column only when inputs are insufficient.
    """
    if df is None or df.empty:
        return pd.Series(dtype="object")
    pft = str(primary_file_type or "").strip()
    if pft and "Liability Type (M61 Raw)" in df.columns:
        note_s = (
            df["Liability Note (M61)"]
            if "Liability Note (M61)" in df.columns
            else pd.Series(pd.NA, index=df.index)
        )
        lt_s = df["Liability Type (M61 Raw)"]
        src_s = df["Source"] if "Source" in df.columns else pd.Series(pd.NA, index=df.index)
        raw_labels = [
            categorize_m61_note_category(n, t, s, primary_file_type=pft)
            for n, t, s in zip(note_s.tolist(), lt_s.tolist(), src_s.tolist())
        ]
        return (
            pd.Series(raw_labels, index=df.index, dtype="object")
            .astype(str)
            .map(normalize_m61_note_category_label)
        )
    if "M61 Note Category" in df.columns:
        return df["M61 Note Category"].fillna("Other").astype(str).map(
            normalize_m61_note_category_label
        )
    for col in ("Liability Type (M61 Raw)", "Liability Type (M61)", "Liability Type"):
        if col in df.columns:
            return (
                df[col]
                .map(categorize_m61_note_type)
                .fillna("Other")
                .astype(str)
                .map(normalize_m61_note_category_label)
            )
    return pd.Series(["other"] * len(df), index=df.index, dtype="object")


def _date_key_ui(v) -> str:
    if v is None:
        return ""
    try:
        if pd.isna(v):
            return ""
    except (TypeError, ValueError):
        pass
    dt = pd.to_datetime(v, errors="coerce")
    if pd.isna(dt):
        return ""
    return dt.strftime("%Y-%m-%d")


def _target_22203_mask_ui(df: pd.DataFrame) -> pd.Series:
    if df is None or df.empty:
        return pd.Series(dtype=bool)
    m = pd.Series(True, index=df.index)
    if "Deal Name" in df.columns:
        m &= df["Deal Name"].fillna("").astype(str).str.strip().str.lower().eq("block 21 san mateo")
    if "Facility" in df.columns:
        m &= df["Facility"].fillna("").astype(str).str.strip().str.lower().eq("tbk bank")
    if "Deal ID Match Key (ACP)" in df.columns:
        did = df["Deal ID Match Key (ACP)"].fillna("").astype(str).str.strip()
        if did.ne("").any():
            m &= did.eq("222203")
    elif "Deal ID (ACP)" in df.columns:
        did2 = (
            df["Deal ID (ACP)"]
            .fillna("")
            .astype(str)
            .str.replace(r"[^A-Za-z0-9]", "", regex=True)
            .str.lower()
        )
        if did2.ne("").any():
            m &= did2.eq("222203")
    eff_acp_col = next((c for c in ("Effective Date (ACP)", "Eff Date (AOC I)", "Eff Date (ACP)") if c in df.columns), None)
    if eff_acp_col:
        m &= pd.to_datetime(df[eff_acp_col], errors="coerce").dt.strftime("%Y-%m-%d").eq("2022-08-22")
    pl_acp_col = next((c for c in ("Pledge Date (ACP)", "Pledge Date (AOC I)", "Pledge Date (ACP)") if c in df.columns), None)
    if pl_acp_col:
        m &= pd.to_datetime(df[pl_acp_col], errors="coerce").dt.strftime("%Y-%m-%d").eq("2023-08-31")
    src_col = next((c for c in ("Source Type (ACORE)", "Source") if c in df.columns), None)
    if src_col:
        m &= df[src_col].fillna("").astype(str).str.lower().str.contains(r"\bsale\b", regex=True, na=False)
    return m


def _target_22203_stage_rows(stage: str, df: pd.DataFrame) -> list[dict[str, object]]:
    out_cols = [
        "Stage",
        "Row Count Found",
        "Deal ID",
        "Deal Name",
        "Source Type (ACORE)",
        "File Source",
        "Eff Date (AOC I)",
        "Eff Date (M61)",
        "Pledge Date (AOC I)",
        "Pledge Date (M61)",
        "Adv Rate (M61)",
        "Spread (M61)",
        "Index Floor (M61)",
    ]
    if df is None or df.empty:
        return [{c: (stage if c == "Stage" else (0 if c == "Row Count Found" else "")) for c in out_cols}]
    mask = _target_22203_mask_ui(df)
    hit = df.loc[mask].copy()
    if hit.empty:
        return [{c: (stage if c == "Stage" else (0 if c == "Row Count Found" else "")) for c in out_cols}]

    def _pick(rr: pd.Series, *keys: str):
        for k in keys:
            if k in rr.index and pd.notna(rr.get(k)):
                return rr.get(k)
        return ""

    rows = []
    cnt = int(len(hit))
    for _, rr in hit.iterrows():
        rows.append(
            {
                "Stage": stage,
                "Row Count Found": cnt,
                "Deal ID": _pick(rr, "Deal ID (ACP)", "Deal ID Match Key (ACP)"),
                "Deal Name": _pick(rr, "Deal Name"),
                "Source Type (ACORE)": _pick(rr, "Source Type (ACORE)", "Source"),
                "File Source": _pick(rr, "File Source"),
                "Eff Date (AOC I)": _pick(rr, "Eff Date (AOC I)", "Eff Date (ACP)", "Effective Date (ACP)"),
                "Eff Date (M61)": _pick(rr, "Eff Date (M61)", "Effective Date (M61)"),
                "Pledge Date (AOC I)": _pick(rr, "Pledge Date (AOC I)", "Pledge Date (ACP)"),
                "Pledge Date (M61)": _pick(rr, "Pledge Date (M61)"),
                "Adv Rate (M61)": _pick(rr, "Adv Rate (M61)", "Advance Rate (M61)"),
                "Spread (M61)": _pick(rr, "Spread (M61)"),
                "Index Floor (M61)": _pick(rr, "Index Floor (M61)"),
            }
        )
    return rows


def _trace_1551_ui_enabled() -> bool:
    return os.environ.get("RECON_UI_TRACE_1551", "").strip() in ("1", "true", "TRUE", "yes", "YES")


def _trace_1551_eff_feb9_series(s: pd.Series) -> pd.Series:
    dt = pd.to_datetime(s, errors="coerce")
    norm_ok = dt.dt.normalize().eq(pd.Timestamp("2026-02-09").normalize())
    st = s.astype(str)
    str_ok = (
        st.str.contains("2/9/26", regex=False, na=False)
        | st.str.contains("02/09/26", regex=False, na=False)
        | st.str.contains("2/9/2026", regex=False, na=False)
        | st.str.contains("2026-02-09", regex=False, na=False)
    )
    return norm_ok | str_ok


def _trace_1551_broadway_ui_mask(df: pd.DataFrame) -> pd.Series:
    """UI pipeline debug: 1551 Broadway + Deal ID 25-2852 + Effective Date (ACP) 2026-02-09."""
    if df is None or df.empty:
        return pd.Series(dtype=bool)
    if "Deal Name" not in df.columns:
        return pd.Series(False, index=df.index)
    m = df["Deal Name"].astype(str).str.contains("1551 Broadway", case=False, na=False)
    id_cols = [c for c in ("Deal ID (ACP)", "Deal ID Match Key (ACP)", "Deal ID") if c in df.columns]
    if id_cols:
        idm = pd.Series(False, index=df.index)
        for c in id_cols:
            idm |= df[c].astype(str).str.contains("25-2852", na=False)
        m &= idm
    eff_col = next(
        (
            c
            for c in (
                "Effective Date (ACP)",
                "Effective Date (ACORE)",
                "Eff Date (ACP)",
                "Eff Date (AOC I)",
                "Effective Date",
            )
            if c in df.columns
        ),
        None,
    )
    if eff_col is not None:
        m &= _trace_1551_eff_feb9_series(df[eff_col])
    return m


def _trace_1551_broadway_ui_stderr(stage: str, df: pd.DataFrame | None) -> None:
    if not _trace_1551_ui_enabled():
        return
    n = 0 if df is None else len(df)
    if df is None or df.empty:
        print(f"[RECON_UI_TRACE_1551] {stage} | total_rows={n} | target_present=False", file=sys.stderr)
        return
    mask = _trace_1551_broadway_ui_mask(df)
    nh = int(mask.sum())
    print(
        f"[RECON_UI_TRACE_1551] {stage} | total_rows={n} | target_row_count={nh} | target_present={nh > 0}",
        file=sys.stderr,
    )
    if nh <= 0:
        return
    if nh > 1:
        print(f"    (note: {nh} target_matches; showing first row only)", file=sys.stderr)
    rr = df.loc[mask].iloc[0]
    rs = rr.get("recon_status", "")
    fs = rr.get("File Source", "")
    nc = rr.get("M61 Note Category", "")
    ed_acp = ""
    for k in ("Effective Date (ACP)", "Effective Date (ACORE)", "Eff Date (ACP)", "Effective Date"):
        if k in rr.index:
            v = rr.get(k)
            if v is not None and not (isinstance(v, float) and pd.isna(v)) and str(v).strip():
                ed_acp = v
                break
    ed_m61 = rr.get("Effective Date (M61)", rr.get("Eff Date (M61)", ""))
    print(
        f"    recon_status={rs!r} | File Source={fs!r} | M61 Note Category={nc!r} | "
        f"Effective Date (ACP)={ed_acp!r} | Effective Date (M61)={ed_m61!r}",
        file=sys.stderr,
    )


# Columns consulted for display-only effective date range filtering (any match → row matches).
EFF_DATE_DISPLAY_COLUMNS = (
    "Effective Date (ACORE)",
    "Effective Date (ACP)",
    "Effective Date (M61)",
    "Effective Date",
    "effective_date_key",
)


def _coerce_to_date(v) -> date | None:
    if v is None:
        return None
    if isinstance(v, datetime):
        return v.date()
    if isinstance(v, date):
        return v
    try:
        if pd.isna(v):
            return None
    except (TypeError, ValueError):
        pass
    try:
        if isinstance(v, pd.Timestamp):
            return v.date()
    except Exception:
        pass
    ts = pd.to_datetime(v, errors="coerce")
    if pd.isna(ts):
        return None
    return ts.date()


def resolve_effective_date_range_bounds(
    preset: str,
    custom_start: date | datetime | None,
    custom_end: date | datetime | None,
) -> tuple[date | None, date | None]:
    """Return inclusive (start, end) calendar bounds, or (None, None) when filtering is off."""
    p = (preset or "All dates").strip()
    if p == "All dates":
        return None, None
    today = date.today()
    if p == "This month":
        start = today.replace(day=1)
        last_d = calendar.monthrange(today.year, today.month)[1]
        end = today.replace(day=last_d)
        return start, end
    if p == "This year":
        return date(today.year, 1, 1), date(today.year, 12, 31)
    if p == "2024":
        return date(2024, 1, 1), date(2024, 12, 31)
    if p == "2025":
        return date(2025, 1, 1), date(2025, 12, 31)
    if p == "Custom range":
        a = _coerce_to_date(custom_start)
        b = _coerce_to_date(custom_end)
        if a is None or b is None:
            return None, None
        if a > b:
            a, b = b, a
        return a, b
    return None, None


def filter_display_dataframe_by_effective_dates(
    df: pd.DataFrame,
    start: date | None,
    end: date | None,
) -> pd.DataFrame:
    """Subset rows for UI/export: keep if any known effective-date column falls in [start, end].

    Rows with ``recon_status`` mismatch signals are always kept when that column is present, so
    negative-test cases (wrong dates or rates on one side) stay visible even when every
    parseable effective date falls outside the selected display window.
    """
    if df is None:
        return df
    if df.empty:
        return df.copy()
    if start is None or end is None:
        return df.copy()
    cols = [c for c in EFF_DATE_DISPLAY_COLUMNS if c in df.columns]
    if not cols:
        return df.copy()

    low = pd.Timestamp(start)
    high = pd.Timestamp(end)
    any_in_range = pd.Series(False, index=df.index)
    has_any_parsed = pd.Series(False, index=df.index)
    for c in cols:
        ts = pd.to_datetime(df[c], errors="coerce")
        valid = ts.notna()
        has_any_parsed = has_any_parsed | valid
        day = ts.dt.normalize()
        in_r = valid & (day >= low) & (day <= high)
        any_in_range = any_in_range | in_r
    # Rows with no parseable dates are kept so blanks never hide data unexpectedly.
    keep = any_in_range | (~has_any_parsed)
    if "recon_status" in df.columns:
        rs_kind = df["recon_status"].map(_recon_status_bucket)
        keep = keep | rs_kind.eq("MISMATCH")
    return df.loc[keep].copy()


def _recon_status_bucket(v: object) -> str:
    """Map business-facing recon status text to MATCH / MISSING / MISMATCH buckets."""
    s = safe_str_strip(v).upper()
    if not s:
        return ""
    # One-sided row — entire record absent from one source.
    if s.startswith("MISSING IN "):
        return "MISSING"
    if "MATCH WITH DIFFERENCES" in s or "DIFFERENCE" in s or "NO MATCH" in s:
        return "MISMATCH"
    if "MATCH WITH MISSING FIELDS" in s or "MISSING FIELDS" in s:
        return "MISSING"
    # Legacy labels (kept for backward compatibility with older cached results).
    if s.startswith("MISMATCH"):
        return "MISMATCH"
    if s.startswith("INCOMPLETE"):
        return "MISSING"
    if s == "MATCH":
        return "MATCH"
    if "MISSING" in s:
        return "MISSING"
    if "MATCH" in s:
        return "MATCH"
    return ""


def _card_contextual_insight(
    card: str,
    n_total: int,
    n_match: int,
    n_needs_review: int,
    n_missing_m61: int,
    df_needs_review: "pd.DataFrame",
) -> str:
    """Return a short plain-English insight paragraph for a clicked dashboard card."""
    if card == "Total ACORE Records":
        if n_total == 0:
            return "No ACORE records are loaded yet. Upload an ACORE file to begin."
        pct_match = round(n_match / n_total * 100) if n_total else 0
        pct_review = round(n_needs_review / n_total * 100) if n_total else 0
        return (
            f"You have **{n_total}** ACORE records in the current view. "
            f"Of these, **{n_match}** ({pct_match}%) are clean matches and "
            f"**{n_needs_review}** ({pct_review}%) need review. "
            f"Use the table below to drill into specific deals."
        )

    if card == "Matches":
        if n_match == 0:
            return (
                "No clean matches found in the current view. "
                "Try adjusting your filters, or check whether the advance rates "
                "and effective dates align between ACORE and M61."
            )
        pct = round(n_match / n_total * 100) if n_total else 0
        return (
            f"**{n_match}** record(s) — {pct}% of the view — reconcile perfectly: "
            f"advance rate, effective date, and all tracked fields agree between ACORE and M61. "
            f"No action needed for these rows."
        )

    if card == "Needs Review":
        if n_needs_review == 0:
            return "Everything in the current view either matches or is missing from one system — no partial mismatches to review right now."
        # Identify the most common mismatch field if data is available
        top_issue = ""
        if not df_needs_review.empty:
            mismatch_cols = [c for c in df_needs_review.columns if c.endswith(" Status")]
            hits: dict[str, int] = {}
            for col in mismatch_cols:
                count = int(df_needs_review[col].astype(str).str.upper().str.contains("MISMATCH|DIFFERENCE|NO MATCH", na=False).sum())
                if count:
                    label = col.replace(" Status", "")
                    hits[label] = count
            if hits:
                top_label, top_count = max(hits.items(), key=lambda x: x[1])
                top_issue = f" The most common discrepancy is **{top_label}** ({top_count} row(s))."
        return (
            f"**{n_needs_review}** record(s) have at least one field that doesn't agree between ACORE and M61.{top_issue} "
            f"Filter the table to 'Needs Review' rows and use **Explain This Row** to see a plain-English summary of each issue."
        )

    if card == "Missing in M61":
        if n_missing_m61 == 0:
            return "All ACORE records in the current view have a matching entry in M61 — nothing is missing."
        return (
            f"**{n_missing_m61}** ACORE record(s) have no corresponding row in M61. "
            f"This usually means the loan hasn't been booked in M61 yet, the deal name or facility "
            f"label differs between systems, or the effective date doesn't align. "
            f"Filter to 'Missing in M61' in the table and confirm with your operations team."
        )

    return ""


def _mismatch_detail_html(row: pd.Series) -> str:
    """Short drilldown hint for negative testing: why recon_status is MISMATCH (from status columns)."""
    if _recon_status_bucket(row.get("recon_status", "")) != "MISMATCH":
        return ""
    parts: list[str] = []
    ed = safe_str_strip(row.get("Effective Date Status", "")).upper()
    if "NO MATCH" in ed or ("MISMATCH" in ed and "MISSING" not in ed):
        parts.append("effective date differs between ACORE and M61")
    ar = safe_str_strip(row.get("Advance Rate Status", "")).upper()
    if "MISMATCH" in ar:
        parts.append("advance rate differs")
    sp = safe_str_strip(row.get("Spread Status", "")).upper()
    if "MISMATCH" in sp:
        parts.append("spread differs")
    if not parts:
        return (
            "<div style='font-size:0.78rem;color:#7b1fa2;margin-top:10px;'>"
            "<strong>Why mismatch:</strong> see status pills above (key field differs)."
            "</div>"
        )
    return (
        "<div style='font-size:0.78rem;color:#7b1fa2;margin-top:10px;'>"
        f"<strong>Why mismatch:</strong> {'; '.join(parts)}."
        "</div>"
    )


def explain_reconciliation_row(row: pd.Series, *, display_row: pd.Series | None = None) -> str:
    """Short business-friendly explanation for the selected grid row (rules-based text only).

    ``display_row`` is the post–Same-as-Above grid row when available; used so explanations match
    what the user sees (consolidated M61 vs. truly missing).
    """
    status_cols: tuple[tuple[str, str], ...] = (
        ("Effective Date Status", "Effective Date"),
        ("Pledge Date Status", "Pledge Date"),
        ("Advance Rate Status", "Advance Rate"),
        ("Spread Status", "Spread"),
        ("Undrawn Capacity Status", "Undrawn Capacity"),
        ("Index Floor Status", "Index Floor"),
        ("Index Name Status", "Index Name"),
        ("Recourse % Status", "Recourse %"),
    )

    def _field_kind(val: object) -> str:
        u = safe_str_strip(val).upper()
        if not u or u in ("N/A", "NAN", "NONE", "-", "—"):
            return "blank"
        if "MISSING FROM BOTH" in u or ("MISSING" in u and "BOTH" in u):
            return "both_missing"
        if "MISSING" in u or "INCOMPLETE" in u:
            return "missing"
        if any(k in u for k in ("MISMATCH", "NO MATCH", "DIFFERENCE", "DIFFERENT")):
            return "mismatch"
        if u == "MATCH" or (u.startswith("MATCH") and "DIFF" not in u):
            return "missing" if "MISSING" in u else "match"
        return "other"

    deal = safe_str_strip(row.get("Deal Name", "")) or "this deal"
    facility = safe_str_strip(row.get("Facility", "")) or "this facility"
    overall = safe_str_strip(row.get("recon_status", ""))
    file_src = safe_str_strip(row.get("File Source", ""))
    bucket = _recon_status_bucket(row.get("recon_status", ""))
    o = overall.upper()

    if display_row is not None and _display_row_shows_same_as_above(display_row):
        return _explain_same_as_above_consolidated_copy(deal, facility)

    mismatch_labels: list[str] = []
    missing_labels: list[str] = []
    both_missing_labels: list[str] = []

    for col, friendly in status_cols:
        if col not in row.index:
            continue
        k = _field_kind(row.get(col))
        if k == "mismatch":
            mismatch_labels.append(friendly)
        elif k == "missing":
            missing_labels.append(friendly)
        elif k == "both_missing":
            both_missing_labels.append(friendly)

    parts: list[str] = []

    if bucket == "MATCH" and not mismatch_labels and not missing_labels:
        parts.append(
            "This row is a clean match between ACORE and M61 for the fields we compare. "
            f"The deal {deal} — facility {facility} — lines up on both sides with no open issues "
            "in the tracked columns."
        )
        parts.append(
            "Next step: No follow-up is required unless you are checking something for audit purposes."
        )
        return "\n\n".join(parts)

    if file_src == FILE_SOURCE_ACORE_ONLY or "MISSING IN M61" in o or "MISSING FROM M61" in o:
        parts.append(
            "This line is on the ACORE side but does not have a matching row in M61, so there is "
            "nothing to compare yet."
        )
        parts.append(
            f"The deal {deal} — facility {facility} — should be checked against your M61 export "
            "and your usual filters (fund, note category, effective date) to see whether the "
            "liability simply has not been booked in M61 or is named differently."
        )
        parts.append(
            "Next step: Confirm with your M61 data owner whether this obligation should appear, "
            "then refresh the export if needed."
        )
        return "\n\n".join(parts)

    if file_src == FILE_SOURCE_M61_ONLY or "MISSING IN ACORE" in o or "MISSING FROM ACORE" in o:
        parts.append(
            "This line appears in M61 but not in your ACORE model for this pairing, so the "
            "primary-side values are not available here."
        )
        parts.append(
            f"The deal {deal} — facility {facility} — may be missing from the model you uploaded, "
            "or the deal or facility label may not match what ACORE expects."
        )
        parts.append(
            "Next step: Check the ACORE workbook export for this deal and line, then re-run "
            "reconciliation after any corrections."
        )
        return "\n\n".join(parts)

    if mismatch_labels and missing_labels:
        parts.append("This row has both value differences and missing data.")
    elif mismatch_labels:
        parts.append("This row has value differences between ACORE and M61.")
    elif missing_labels:
        parts.append(
            "This row matches where we have values, but some fields are missing on one side."
        )
    else:
        parts.append(
            f"For this row the reconciliation status reads: {overall or 'see the status columns in the table'}."
        )

    if mismatch_labels or missing_labels:
        parts.append(
            f"The deal {deal} — facility {facility} — was found in both ACORE and M61, "
            + (
                "but some fields do not agree and others are blank on one side."
                if (mismatch_labels and missing_labels)
                else (
                    "but some fields do not agree between the two files."
                    if mismatch_labels
                    else "and the items below are worth a quick look because one side has no value."
                )
            )
        )

    if mismatch_labels:
        parts.append(
            "These fields have different values between ACORE and M61:\n"
            + "\n".join(f"- {name}" for name in mismatch_labels)
        )

    if missing_labels:
        parts.append(
            "However, these fields need review because one side is missing data:\n"
            + "\n".join(f"- {name}" for name in missing_labels)
        )

    if both_missing_labels:
        parts.append(
            "These fields are empty on both sides for this row, which is often normal if neither "
            "system tracks them:\n"
            + "\n".join(f"- {name}" for name in both_missing_labels)
        )

    if mismatch_labels or missing_labels:
        parts.append(
            "Next step: Open the Deal Drilldown tab to compare the source values side by side."
        )
    else:
        parts.append(
            "Next step: Use the Deal Drilldown tab or the status columns in the table to see how "
            "the two files line up for this deal."
        )

    return "\n\n".join(parts)


def render_full_app_shell(selected_ui_label: str, m61_export_line_html: str) -> None:
    """Main-area header (sidebar is built separately in its own ``with st.sidebar`` block)."""
    st.markdown(
        f"""
<div class="recon-header">
  <div>
    <h1>📊 Financing Line Reconciliation</h1>
    <p>Primary: <strong>{html.escape(str(selected_ui_label), quote=True)}</strong> · {m61_export_line_html}</p>
  </div>
  <div class="badge">Run: {datetime.now().strftime("%b %d, %Y  %H:%M")}</div>
</div>
""",
        unsafe_allow_html=True,
    )


def render_original_landing_page_if_no_results(selected_ui_label: str) -> None:
    """Original polished holding page when no reconciliation results are in session."""
    st.markdown(
        f"""
    <div style='text-align:center; padding:60px 40px; background:#fff; border-radius:14px;
                border: 1.5px dashed #ccd9ea; color:#5a7a99;'>
      <div style='font-size:3rem; margin-bottom:16px;'>📂</div>
      <h3 style='color:#1a3a6c; font-weight:700; margin-bottom:8px'>Upload both files to begin</h3>
      <p style='font-size:0.88rem; max-width:420px; margin:0 auto; line-height:1.6;'>
        Choose the <strong>{html.escape(str(selected_ui_label), quote=True)}</strong> primary fund template in the sidebar,
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
    _fn_bad, _fn_msg = primary_filename_incompatible_acp_ii_vs_iii(
        file_b_upload.name if file_b_upload else None, primary_file_type
    )
    if _fn_bad:
        st.error(_fn_msg)
    elif inferred_primary_type and inferred_primary_type != primary_file_type:
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
    # Optional mapping workbook is still supported by reconcile(); not exposed in UI.
    mapping_upload = None

    st.markdown("---")
    st.markdown("## ⚙️ Options")
    run_on_upload = st.checkbox("Auto-run on upload", value=True)

    st.markdown("---")
    st.markdown("## 🔍 Filters")
    st.markdown(
        '<div class="sidebar-subheader-row">'
        '<span class="sidebar-subheader-inline">Show Records</span>'
        '<span class="sidebar-subheader-hint" title="Choose which rows appear in the results table. '
        "Clean Matches are fully aligned records. Records Needing Review groups mismatches, one-sided missing "
        'records, and incomplete matches — refine with Advanced Review Filters.">ⓘ</span>'
        "</div>",
        unsafe_allow_html=True,
    )

    # Apply M61 note-category default *before* the selectbox is instantiated (same-run safe).
    if st.session_state.pop("recon_pending_m61_note_category_reset", False):
        st.session_state["recon_m61_note_category"] = "Financing"

    if st.session_state.pop("reset_status_filters", False):
        st.session_state["filter_needs_review_bundle"] = True
        st.session_state["filter_review_clean_matches"] = False
        st.session_state["filter_review_missing_fields"] = True
        st.session_state["filter_review_differences"] = True
        st.session_state["filter_review_missing_m61"] = True
        st.session_state["filter_review_missing_acore"] = True
    if "filter_needs_review_bundle" not in st.session_state:
        st.session_state["filter_needs_review_bundle"] = True
    if "filter_review_clean_matches" not in st.session_state:
        st.session_state["filter_review_clean_matches"] = False
    if "filter_review_missing_fields" not in st.session_state:
        st.session_state["filter_review_missing_fields"] = True
    if "filter_review_differences" not in st.session_state:
        st.session_state["filter_review_differences"] = True
    if "filter_review_missing_m61" not in st.session_state:
        st.session_state["filter_review_missing_m61"] = True
    if "filter_review_missing_acore" not in st.session_state:
        st.session_state["filter_review_missing_acore"] = True
    st.checkbox(
        "Clean Matches",
        key="filter_review_clean_matches",
        help="Include rows that are a clean **MATCH** (fully aligned between ACORE and M61).",
    )
    st.checkbox(
        "Records Needing Review",
        key="filter_needs_review_bundle",
        help=(
            "Include rows that typically need follow-up: mismatches, missing on M61 or ACORE, "
            "or incomplete / one-sided missing fields. Use **Advanced Review Filters** to narrow."
        ),
    )
    with st.expander("Advanced Review Filters", expanded=False):
        st.caption(
            "Applies when **Records Needing Review** is checked. Same row categories as before — "
            "only organized here for clarity."
        )
        st.checkbox(
            "Mismatches",
            key="filter_review_differences",
            help="Value differences and field-level mismatches.",
        )
        st.checkbox(
            "Missing in M61",
            key="filter_review_missing_m61",
            help="ACORE-side rows with no matching M61 record.",
        )
        st.checkbox(
            "Missing in ACORE",
            key="filter_review_missing_acore",
            help="Rows present in M61 but not in ACORE (M61-only lines).",
        )
        st.checkbox(
            "Incomplete (one-sided missing)",
            key="filter_review_missing_fields",
            help="Matched rows with missing fields on one side.",
        )

    st.markdown(
        '<div class="sidebar-section-gap" aria-hidden="true"></div>',
        unsafe_allow_html=True,
    )

    # Developer / debug UI hidden from finance users — uncomment to expose full-M61 reconciliation toggle.
    # with st.expander("Developer / debug", expanded=False):
    #     st.checkbox(
    #         "Show full reconciliation (all M61 funds)",
    #         key="recon_debug_full_m61",
    #         help=(
    #             "Entire reconciliation output, including liabilities for other funds present in the "
    #             "M61 export. For troubleshooting only — finance views stay scoped to the uploaded primary fund."
    #         ),
    #     )

    # Session key still honored if set elsewhere; without the widget it defaults to False.
    _debug_full_sidebar = bool(st.session_state.get("recon_debug_full_m61", False))

    # Scope: both finance modes use the uploaded primary fund only; "Selected" also shows M61-only
    # rows that passed the same fund filter (hides other funds, not all M61-only).
    scope_mode = "Selected Fund Only"
    scope_toggle_needed = True
    if _debug_full_sidebar:
        scope_toggle_needed = False
    elif "df_recon" in st.session_state:
        _scope_df = st.session_state.get("df_recon", pd.DataFrame())
        _scope_primary = canonical_primary_file_type(
            st.session_state.get("primary_file_type", primary_file_type)
        )
        if isinstance(_scope_df, pd.DataFrame) and not _scope_df.empty:
            _fund = filter_recon_to_selected_fund(_scope_df, _scope_primary)
            _has_other_funds = not _scope_df.sort_index().equals(_fund.sort_index())
            _has_m61_only = False
            if "File Source" in _fund.columns:
                _fs = _fund["File Source"].fillna("").astype(str).str.strip()
                _has_m61_only = bool(_fs.eq(FILE_SOURCE_M61_ONLY).any())
            scope_toggle_needed = _has_other_funds or _has_m61_only

    if scope_toggle_needed:
        _scope_label_all = "All Results for Uploaded Primary Fund"
        _scope_label_fund = "Selected Primary Fund Only"
        st.markdown(
            '<div class="sidebar-subheader sidebar-subheader--after-gap">Scope</div>',
            unsafe_allow_html=True,
        )
        _scope_choice = st.radio(
            "Scope",
            [_scope_label_all, _scope_label_fund],
            index=1,
            label_visibility="collapsed",
            help=(
                f"**{_scope_label_all}:** All records for **{scope_label_for_primary_type(primary_file_type)}** "
                "in this run, including lines that appear only on the comparison export. "
                f"**{_scope_label_fund}:** ACORE rows for this fund, their M61 matches, and M61-only lines "
                "that match this fund’s scope (other funds remain excluded)."
            ),
        )
        st.caption(
            "All Results shows the full reconciliation output for the uploaded primary fund, "
            "including related M61-only records. "
            f"Selected Primary Fund Only shows the same fund filter as All Results, but focuses the table on "
            f"**{scope_label_for_primary_type(primary_file_type)}**-tied ACORE rows, their M61 matches, and "
            "M61-only export lines that match this fund (not other funds)."
        )
        scope_mode = (
            "Selected Fund Only"
            if _scope_choice == _scope_label_fund
            else "All Results"
        )
    else:
        scope_mode = "Selected Fund Only"
        if _debug_full_sidebar:
            st.caption(
                "**Developer:** Scope choices are hidden while full-export view is on. Turn it off to use Scope again."
            )
        else:
            st.caption(
                f"Everything here is already for **{scope_label_for_primary_type(primary_file_type)}** only."
            )

    st.markdown(
        '<div class="sidebar-section-gap" aria-hidden="true"></div>',
        unsafe_allow_html=True,
    )

    _note_options = ["All", "Financing", "Subline", "Other"]
    if "recon_m61_note_category" not in st.session_state:
        st.session_state["recon_m61_note_category"] = "Financing"
    if st.session_state.get("recon_m61_note_category") not in _note_options:
        st.session_state["recon_m61_note_category"] = "All"
    st.markdown(
        '<div class="sidebar-subheader sidebar-subheader--after-gap">M61 Note Category</div>',
        unsafe_allow_html=True,
    )
    st.selectbox(
        "M61 Note Category",
        options=_note_options,
        key="recon_m61_note_category",
        label_visibility="collapsed",
        help=(
            "Filters the **displayed** table, drilldown, and downloads together with **Scope**, "
            "**Show Records**, and **More filters / Advanced filters → Effective date**. "
            "Values come from row-level **M61 Note Category**. "
            "Choose **All** to show every category."
        ),
    )

    with st.expander("Advanced filters", expanded=False):
        st.caption("Display only — does not change reconciliation. Applies with Scope and other sidebar filters.")
        st.selectbox(
            "Effective date range",
            options=[
                "All dates",
                "2024",
                "2025",
                "2026",
                "2027",
                "2028",
                "2029",
                "2030",
                "Custom range",
            ],
            index=0,
            key="recon_eff_date_preset",
            help=(
                "Uses columns present on the output: Effective Date (ACORE), Effective Date (ACP), "
                "Effective Date (M61), Effective Date, effective_date_key. "
                "A row matches if any non-blank date falls in the range; rows with no parseable dates stay visible. "
                "Rows with reconciliation status **MISMATCH** are always shown, even if all dates are outside the range "
                "(supports negative testing)."
            ),
        )
        if st.session_state.get("recon_eff_date_preset") == "Custom range":
            _ed_c1, _ed_c2 = st.columns(2)
            with _ed_c1:
                st.date_input(
                    "Start",
                    value=date.today().replace(day=1),
                    key="recon_eff_date_custom_start",
                    format="MM/DD/YYYY",
                )
            with _ed_c2:
                st.date_input(
                    "End",
                    value=date.today(),
                    key="recon_eff_date_custom_end",
                    format="MM/DD/YYYY",
                )
 
    st.markdown("---")
    _m61_display_name = (
        file_a_upload.name
        if file_a_upload
        else st.session_state.get("persisted_m61_upload_name")
    )
    _m61_export_line = (
        f"M61 Export: {html.escape(_m61_display_name, quote=True)}"
        if _m61_display_name
        else "M61 Export: —"
    )
    st.markdown(
        f"""
    <div style='font-size:0.7rem; color:#5a7890; line-height:1.6'>
    <strong style='color:#8fb8d8'>Primary Source</strong><br>{html.escape(str(_pc["model_descriptor"]), quote=True)}<br><br>
    <strong style='color:#8fb8d8'>Comparison Source</strong><br>{_m61_export_line}<br><br>
    
    </div>
    """,
        unsafe_allow_html=True,
    )
 
# <strong style='color:#8fb8d8'>Target Advance Rate</strong><br>From M61 file only


render_full_app_shell(selected_ui_label, _m61_export_line)


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
    if su.startswith("INCOMPLETE"):
        # One-sided missing fields only — amber, same weight as MISSING.
        return f'<span class="pill pill-incomplete">⚠ {status}</span>'
    # MISSING FROM BOTH = absent on both sides, not a real issue — render muted, not amber.
    # Must be checked before the general MISSING branch.
    if "MISSING FROM BOTH" in su:
        return f'<span class="pill pill-na">— {status}</span>'
    if "MISSING" in su:
        return f'<span class="pill pill-missing">⚠ {status}</span>'
    return f'<span class="pill pill-na">{status}</span>'
 
 
SAME_AS_ABOVE_LABEL = "Same as Above"


def _is_same_as_above_label(v) -> bool:
    return isinstance(v, str) and str(v).strip() == SAME_AS_ABOVE_LABEL


def pct(v, *, ndigits: int = 2, missing: str = "—"):
    if _is_same_as_above_label(v):
        return SAME_AS_ABOVE_LABEL
    if v is None:
        return missing
    try:
        if isinstance(v, str) and str(v).strip() in ("-", "—", ""):
            return missing
    except Exception:
        pass
    try:
        s = str(v).strip()
        if s in ("-", "—"):
            return missing
        if s.endswith("%"):
            fv = float(s.replace("%", "").replace(",", "").strip()) / 100.0
        else:
            fv = float(v)
        if pd.isna(fv):
            return missing
        return f"{fv:.{ndigits}%}"
    except Exception:
        return missing


def pct_spread(v):
    """Spread-only: percent with 3 decimal places; missing → ``-`` (aligned with display/export dashes)."""
    if _is_same_as_above_label(v):
        return SAME_AS_ABOVE_LABEL
    return pct(v, ndigits=3, missing="-")


def _is_spread_value_column(name: object) -> bool:
    cl = str(name).lower()
    return "spread" in cl and "status" not in cl


def fmt_fraction_as_pct(v, *, ndigits: int = 3):
    """Format a stored fraction (e.g. 0.02275) for display as a percent; — when missing."""
    if _is_same_as_above_label(v):
        return SAME_AS_ABOVE_LABEL
    if v is None or (isinstance(v, float) and pd.isna(v)):
        return "—"
    try:
        fv = float(v)
        if pd.isna(fv):
            return "—"
        return f"{fv:.{ndigits}%}"
    except (TypeError, ValueError):
        return "—"
 
 
def fmt_date(v):
    if _is_same_as_above_label(v):
        return SAME_AS_ABOVE_LABEL
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
    # Treat None, NA, and numeric zero as blank for display purposes.
    # Undrawn Capacity = 0 in M61 exports is an Excel artifact for blank cells,
    # not a meaningful "fully-drawn" value distinct from missing.
    if _is_same_as_above_label(v):
        return SAME_AS_ABOVE_LABEL
    if v is None:
        return "—"
    try:
        if pd.isna(v):
            return "—"
    except (TypeError, ValueError):
        pass
    try:
        f = float(v)
        return "—" if f == 0 else f"{f:,.0f}"
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


# Text columns where a lone ``NA`` from Fin Inpt is treated like explicit ``N/A`` for display.
_DISPLAY_PRESERVE_NA_TEXT_COLS = frozenset({"Facility", "Note Name", "Source", "Financial Line"})


def _display_missing_dash(v, col: str | None = None):
    """Display-only missing-value normalizer: missing -> '-', keep real zeros.

    Literal ``N/A`` is never treated as missing. ``Facility`` (and related text fields) may use
    ``NA`` as an explicit placeholder — show as ``N/A`` rather than dash.
    """
    if v is None:
        return "-"
    try:
        if pd.isna(v):
            return "-"
    except (TypeError, ValueError):
        pass
    s = str(v).strip()
    if not s:
        return "-"
    su = s.upper()
    if su == "N/A":
        return "N/A"
    if col in _DISPLAY_PRESERVE_NA_TEXT_COLS and su == "NA":
        return "N/A"
    # String sentinels from ``astype(str)`` on missing floats — still missing for display.
    if su in ("NAN", "<NA>", "NAT", "NONE"):
        return "-"
    if s in ("—", "–", "−"):
        return "-"
    return v


def _normalize_display_missing_df(df: pd.DataFrame) -> pd.DataFrame:
    """Apply business-facing display normalization to all table/export cells."""
    if df is None or df.empty:
        return df.copy() if isinstance(df, pd.DataFrame) else pd.DataFrame()
    out = df.copy()
    for c in out.columns:
        out[c] = out[c].map(lambda x, _c=c: _display_missing_dash(x, col=_c))
    return out


def derive_liability_type_for_filter(row: pd.Series) -> str:
    """Best-effort liability type for UI filtering from available row fields."""
    # Prefer explicit M61 type when present in schema.
    explicit_keys = ("Liability Type (M61)", "Liability Type (M61 Raw)", "Liability Type")
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


def categorize_m61_note_for_filter(raw_liability_type) -> str:
    """Delegate to the canonical backend function — single source of truth."""
    return categorize_m61_note_type(raw_liability_type)


def _series_m61_note_category(df: pd.DataFrame) -> pd.Series:
    """Per-row M61 note category for filtering (matches ``categorize_m61_note_type``)."""
    if df is None or df.empty:
        return pd.Series(dtype="object")
    if "M61 Note Category" in df.columns:
        return df["M61 Note Category"].fillna("Other").astype(str).str.strip()
    for col in ("Liability Type (M61 Raw)", "Liability Type (M61)", "Liability Type"):
        if col in df.columns:
            return (
                df[col]
                .map(categorize_m61_note_type)
                .fillna("Other")
                .astype(str)
                .str.strip()
            )
    return pd.Series(["Other"] * len(df), index=df.index, dtype="object")


def to_excel_bytes(df_recon, primary_file_type: str):
    df_recon = normalize_recon_fund_for_output(df_recon).reset_index(drop=True)
    # User-facing export hides internal Target Advance Rate column.
    df_recon = df_recon.drop(columns=["Target Advance Rate (M61)"], errors="ignore")
    wb = build_workbook(df_recon, primary_file_type=primary_file_type)
    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf.read()


def _display_file_source_cell(row: pd.Series) -> str:
    """Non-blank File Source for the All Results grid (handles NA and legacy ``ID Match Result`` fallback)."""
    v = row.get("File Source")
    try:
        if v is not None and not pd.isna(v):
            s = str(v).strip()
            if s and s.lower() not in ("nan", "none", "<na>"):
                return s
    except (TypeError, ValueError):
        pass
    imr = row.get("ID Match Result")
    try:
        if imr is None or pd.isna(imr):
            return ""
    except (TypeError, ValueError):
        return ""
    k = str(imr).strip().lower()
    return {
        "both": FILE_SOURCE_BOTH,
        "left_only": FILE_SOURCE_ACORE_ONLY,
        "right_only": FILE_SOURCE_M61_ONLY,
    }.get(k, "")


# ---------------------------------------------------------------------------
# "Same as Above" display helpers
#
# Business context: ACORE holds individual rows for each tranche / sub-line.
# M61 may consolidate those rows into a single liability record.  When that
# happens the extra ACORE-Only siblings should NOT look like they are missing
# from M61 — they are simply covered by the matched row above.
#
# Rules:
#   • Applied at the FULL ROW level — every M61-side value column AND every
#     status column on that sibling row is replaced with "Same as Above".
#   • **File Source** is set to **Both** for those sibling rows (M61 data exists for the
#     group; display-only — not a new M61 match row).
#   • Sibling rows are sorted to appear directly after the first Both row of their group.
#   • Rows genuinely absent from M61 (no Both sibling in the same group) are
#     never touched.
#   • Styling (light-blue highlight) is applied separately in the render block.
# ---------------------------------------------------------------------------

# Display-side M61 value columns (formatted strings in df_display)
_SAA_DISPLAY_M61_VALUE_COLS = [
    "Eff Date (M61)",
    "Pledge Date (M61)",
    "Adv Rate (M61)",
    "Spread (M61)",
    "Undrawn (M61)",
    "Index Floor (M61)",
    "Index Name (M61)",
    "Recourse % (M61)",
    "Liability Type (M61)",
    "M61 Note Category",
]

# Display-side status columns that also get replaced
_SAA_DISPLAY_STATUS_COLS = [
    "Effective Date Status",
    "Pledge Date Status",
    "Advance Rate Status",
    "Spread Status",
    "Undrawn Capacity Status",
    "Index Floor Status",
    "Index Name Status",
    "Recourse % Status",
    "Overall Recon Status",
]

# Raw M61 value columns in the export / df_view dataframe
_SAA_EXPORT_M61_VALUE_COLS = [
    "Effective Date (M61)",
    "Pledge Date (M61)",
    "Advance Rate (M61)",
    "Spread (M61)",
    "Undrawn Capacity (M61)",
    "Index Floor (M61)",
    "Index Name (M61)",
    "Recourse % (M61)",
    "Liability Type (M61)",
    "Liability Type (M61 Raw)",
    "M61 Note Category",
]

# Raw status columns in the export dataframe
_SAA_EXPORT_STATUS_COLS = [
    "Effective Date Status",
    "Pledge Date Status",
    "Advance Rate Status",
    "Spread Status",
    "Undrawn Capacity Status",
    "Index Floor Status",
    "Index Name Status",
    "Recourse % Status",
    "recon_status",
    "Overall Recon Status",
]


_SAA_CELL_LABEL = SAME_AS_ABOVE_LABEL
_SAA_UNDRAWN_DISPLAY_COL = "Undrawn (M61)"
_SAA_UNDRAWN_EXPORT_COL = "Undrawn Capacity (M61)"


def _raw_m61_undrawn_present(v) -> bool:
    """True when the source M61 undrawn cell has a real value (not blank; 0 treated as blank)."""
    if _is_blank_for_compare(v):
        return False
    n = _coerce_numeric_value(v)
    return n is not None and n != 0


def _raw_m61_undrawn_values_equal(cur_raw, anchor_raw) -> bool:
    if not _raw_m61_undrawn_present(cur_raw) or not _raw_m61_undrawn_present(anchor_raw):
        return False
    return _coerce_numeric_value(cur_raw) == _coerce_numeric_value(anchor_raw)


def _same_as_above_undrawn_cell(cur_raw, anchor_raw) -> str | None:
    """Undrawn (M61) for consolidated siblings; ``None`` = keep existing formatted value."""
    if _raw_m61_undrawn_present(cur_raw) and _raw_m61_undrawn_values_equal(cur_raw, anchor_raw):
        return SAME_AS_ABOVE_LABEL
    if not _raw_m61_undrawn_present(cur_raw):
        return "-"
    return None


def _display_row_shows_same_as_above(disp: pd.Series | None) -> bool:
    """True when the results grid row uses Same-as-Above (consolidated M61) display."""
    if disp is None or len(disp) == 0:
        return False
    for c in _SAA_DISPLAY_M61_VALUE_COLS:
        if c in disp.index and str(disp.get(c, "")).strip() == _SAA_CELL_LABEL:
            return True
    for c in _SAA_DISPLAY_STATUS_COLS:
        if c in disp.index and str(disp.get(c, "")).strip() == _SAA_CELL_LABEL:
            return True
    return False


def _explain_same_as_above_consolidated_copy(deal: str, facility: str) -> str:
    """Explain text for rows where M61 values are intentionally not repeated (Same as Above)."""
    return (
        f"This ACORE line is included in the same M61 liability record as the "
        f"matched **Both** row above (**{deal}** – **{facility}**). "
        "M61 combines some related ACORE lines into a single liability entry, "
        "so the M61 values are not repeated on this row."

        "\n\n"

        "This row is **not missing** from **M61** — refer to the matched row above "
        "for the related M61 details."

        "\n\n"

        "Next step: Review the shared M61 row above if you need the underlying "
        "M61 values for this ACORE line."
    )


def _apply_same_as_above_display(
    df: pd.DataFrame,
    col_tag: str,
    *,
    raw_undrawn_m61: pd.Series | None = None,
) -> pd.DataFrame:
    """Replace M61 value + status columns with 'Same as Above' for ACORE Only rows
    that are tied to the same consolidated M61 record as a matched 'Both' row in the
    same (Deal Name, Source Type (ACORE), Eff Date) group. Sets **File Source** to **Both**
    for those sibling rows (display-only).

    The sibling rows are also re-sorted to sit directly underneath their matched row
    so the visual relationship is immediately clear.

    Rows that are genuinely missing from M61 (no Both sibling in their group) are
    never touched.
    """
    if df.empty:
        return df

    eff_col = f"Eff Date ({col_tag})"
    m61_cols = [c for c in _SAA_DISPLAY_M61_VALUE_COLS if c in df.columns]
    status_cols = [c for c in _SAA_DISPLAY_STATUS_COLS if c in df.columns]
    all_saa_cols = m61_cols + status_cols

    required = ["Deal Name", "Source Type (ACORE)", eff_col, "File Source"]
    if not all_saa_cols or any(c not in df.columns for c in required):
        return df

    df = df.copy()
    df["_orig_pos"] = range(len(df))

    gkey = (
        df["Deal Name"].fillna("").astype(str)
        + "||"
        + df["Source Type (ACORE)"].fillna("").astype(str)
        + "||"
        + df[eff_col].fillna("").astype(str)
    )
    df["_gkey"] = gkey

    both_mask = df["File Source"] == FILE_SOURCE_BOTH
    acore_only_mask = df["File Source"] == FILE_SOURCE_ACORE_ONLY

    eligible_groups = set(gkey[both_mask]) & set(gkey[acore_only_mask])
    if not eligible_groups:
        df.drop(columns=["_orig_pos", "_gkey"], inplace=True)
        return df

    # Rows to mark "Same as Above": ACORE Only siblings whose group has a Both match
    saa_mask = acore_only_mask & gkey.isin(eligible_groups)
    if not saa_mask.any():
        df.drop(columns=["_orig_pos", "_gkey"], inplace=True)
        return df

    first_both_pos = (
        df.loc[both_mask & df["_gkey"].isin(eligible_groups)]
        .groupby("_gkey")["_orig_pos"]
        .min()
    )

    undrawn_display_col = _SAA_UNDRAWN_DISPLAY_COL
    saa_value_cols = [c for c in m61_cols if c != undrawn_display_col]
    for col in saa_value_cols + status_cols:
        df.loc[saa_mask, col] = _SAA_CELL_LABEL
    df.loc[saa_mask, "File Source"] = FILE_SOURCE_BOTH

    if undrawn_display_col in df.columns and saa_mask.any():
        for idx in df.index[saa_mask]:
            gk = df.at[idx, "_gkey"]
            anchor_pos = first_both_pos.get(gk)
            if anchor_pos is None or raw_undrawn_m61 is None:
                df.at[idx, undrawn_display_col] = "-"
                continue
            anchor_ixs = df.index[df["_orig_pos"] == anchor_pos]
            if len(anchor_ixs) == 0:
                df.at[idx, undrawn_display_col] = "-"
                continue
            anchor_idx = anchor_ixs[0]
            cur_raw = raw_undrawn_m61.loc[idx] if idx in raw_undrawn_m61.index else pd.NA
            anchor_raw = (
                raw_undrawn_m61.loc[anchor_idx] if anchor_idx in raw_undrawn_m61.index else pd.NA
            )
            undrawn_disp = _same_as_above_undrawn_cell(cur_raw, anchor_raw)
            if undrawn_disp is not None:
                df.at[idx, undrawn_display_col] = undrawn_disp

    # --- Reorder: siblings go directly after the first Both row of their group ---
    # Sort key for each row:
    #   Non-eligible rows         → (_orig_pos, 0, _orig_pos)
    #   Eligible Both rows        → (first_both_pos_in_group, 0, _orig_pos)
    #   Eligible ACORE Only rows  → (first_both_pos_in_group, 1, _orig_pos)
    #
    # This keeps non-eligible rows at their natural positions while pulling ACORE
    # Only siblings to sit immediately after the Both row of their group.

    df["_group_anchor"] = df["_gkey"].map(first_both_pos)

    in_eligible = df["_gkey"].isin(eligible_groups) & (both_mask | saa_mask)
    df["_sort_anchor"] = df["_orig_pos"]
    df.loc[in_eligible, "_sort_anchor"] = df.loc[in_eligible, "_group_anchor"]

    df["_sort_type"] = 0
    df.loc[saa_mask, "_sort_type"] = 1   # siblings sort after Both rows

    df = (
        df.sort_values(["_sort_anchor", "_sort_type", "_orig_pos"], kind="stable")
        .reset_index(drop=True)
    )
    df.drop(
        columns=["_orig_pos", "_gkey", "_group_anchor", "_sort_anchor", "_sort_type"],
        inplace=True,
    )
    return df


def _apply_same_as_above_export(df: pd.DataFrame) -> pd.DataFrame:
    """Export-side equivalent of ``_apply_same_as_above_display``.

    Replaces both M61 value columns and status columns with 'Same as Above' for
    sibling ACORE Only rows, using raw reconciliation-engine column names.
    Sets **File Source** to **Both** for those rows (display/export parity with the grid).
    Grouping key: (Deal Name, Source, Effective Date (ACP)).
    """
    if df.empty:
        return df

    m61_cols = [c for c in _SAA_EXPORT_M61_VALUE_COLS if c in df.columns]
    status_cols = [c for c in _SAA_EXPORT_STATUS_COLS if c in df.columns]
    all_saa_cols = m61_cols + status_cols

    eff_col = "Effective Date (ACP)" if "Effective Date (ACP)" in df.columns else "Effective Date"
    required = ["Deal Name", "Source", eff_col, "File Source"]
    if not all_saa_cols or any(c not in df.columns for c in required):
        return df

    df = df.copy()
    gkey = (
        df["Deal Name"].fillna("").astype(str)
        + "||"
        + df["Source"].fillna("").astype(str)
        + "||"
        + df[eff_col].fillna("").astype(str)
    )
    both_groups = set(gkey[df["File Source"] == FILE_SOURCE_BOTH])
    if not both_groups:
        return df

    mask = (df["File Source"] == FILE_SOURCE_ACORE_ONLY) & gkey.isin(both_groups)
    if not mask.any():
        return df

    both_mask_orig = df["File Source"] == FILE_SOURCE_BOTH
    undrawn_export_col = _SAA_UNDRAWN_EXPORT_COL
    saa_value_cols = [c for c in m61_cols if c != undrawn_export_col]
    for col in saa_value_cols + status_cols:
        df.loc[mask, col] = _SAA_CELL_LABEL
    df.loc[mask, "File Source"] = FILE_SOURCE_BOTH

    if undrawn_export_col in df.columns and mask.any():
        for idx in df.index[mask]:
            gk = gkey.loc[idx]
            anchor_ixs = df.index[both_mask_orig & gkey.eq(gk)]
            if len(anchor_ixs) == 0:
                df.at[idx, undrawn_export_col] = "-"
                continue
            anchor_idx = anchor_ixs[0]
            cur_raw = df.at[idx, undrawn_export_col]
            anchor_raw = df.at[anchor_idx, undrawn_export_col]
            undrawn_val = _same_as_above_undrawn_cell(cur_raw, anchor_raw)
            if undrawn_val is not None:
                df.at[idx, undrawn_export_col] = undrawn_val

    return df


def _scope_mode_display(scope_mode: str, debug_full: bool) -> str:
    """Short labels for debug readouts (internal scope_mode values unchanged)."""
    if debug_full:
        return "Full export (developer)"
    if scope_mode == "Selected Fund Only":
        return "Selected primary fund"
    if scope_mode == "All Results":
        return "All results for primary fund"
    return scope_mode


_UPLOAD_EXTS_ALLOWED = frozenset({".xlsx", ".xlsm", ".xls", ".xlsb"})


def _upload_has_payload(uploaded_file) -> bool:
    """True when the Streamlit upload object contains non-empty file bytes."""
    if uploaded_file is None:
        return False
    try:
        return len(_bytes_from_streamlit_upload(uploaded_file)) > 0
    except (TypeError, ValueError, OSError):
        return False


def _both_uploads_ready(file_a_upload, file_b_upload) -> bool:
    """Both workbooks must be present in memory before reconcile runs (cloud-safe)."""
    return bool(file_a_upload and file_b_upload and _upload_has_payload(file_a_upload) and _upload_has_payload(file_b_upload))


def _file_hash_hex(data: bytes) -> str:
    return hashlib.sha256(data).hexdigest()


def _upload_fingerprint(uploaded_file):
    """Stable workbook identity: ``(filename, byte length, sha256 hex)``."""
    if not uploaded_file:
        return None
    try:
        raw = _bytes_from_streamlit_upload(uploaded_file)
    except (TypeError, ValueError, OSError):
        return None
    if not raw:
        return None
    name = str(getattr(uploaded_file, "name", "") or "")
    return (name, len(raw), _file_hash_hex(raw))


def _sanitize_multiselect_state(ms_key: str, opts: list) -> None:
    """Drop stale multiselect values so filter reruns do not crash when options shrink."""
    if ms_key not in st.session_state:
        return
    prior = st.session_state.get(ms_key)
    if not isinstance(prior, list):
        return
    allowed = set(opts)
    cleaned = [v for v in prior if v in allowed]
    if cleaned != prior:
        st.session_state[ms_key] = cleaned


def _update_persisted_upload_fingerprints(
    file_a_upload,
    file_b_upload,
    mapping_upload,
    primary_file_type: str,
) -> None:
    """Remember last in-memory uploads so a transient ``None`` widget does not drop the workspace."""
    fa = _upload_fingerprint(file_a_upload)
    fb = _upload_fingerprint(file_b_upload)
    fm = _upload_fingerprint(mapping_upload) if mapping_upload else None
    if fa is not None:
        st.session_state["persisted_m61_fingerprint"] = fa
        st.session_state["persisted_m61_upload_name"] = fa[0]
    if fb is not None:
        st.session_state["persisted_primary_fingerprint"] = fb
        st.session_state["persisted_primary_upload_name"] = fb[0]
    if fm is not None:
        st.session_state["persisted_mapping_fingerprint"] = fm
    if fa is not None or fb is not None:
        st.session_state["persisted_primary_file_type_at_upload"] = str(
            primary_file_type
        ).strip()


def _upload_conflicts_with_persisted(uploaded_file, persisted_key: str) -> bool:
    """True when the user replaced a workbook with different bytes (not a transient None)."""
    if uploaded_file is None:
        return False
    fp = _upload_fingerprint(uploaded_file)
    if fp is None or fp[1] <= 0:
        return False
    persisted = st.session_state.get(persisted_key)
    if persisted is None:
        return False
    return fp != persisted


def _has_restored_reconciliation_workspace() -> bool:
    return bool(
        "df_recon" in st.session_state
        and st.session_state.get("last_successful_upload_signature")
    )


def _workspace_should_be_cleared(
    file_a_upload,
    file_b_upload,
    mapping_upload,
    primary_file_type: str,
) -> bool:
    """Discard cached reconciliation only when inputs materially changed or user requested reset."""
    if st.session_state.pop("recon_user_requested_clear", False):
        return True
    if "df_recon" not in st.session_state:
        return False
    # Table/filter reruns often leave upload widgets empty; that is not a file change.
    if not _both_uploads_ready(file_a_upload, file_b_upload):
        return False
    if _upload_conflicts_with_persisted(file_a_upload, "persisted_m61_fingerprint"):
        return True
    if _upload_conflicts_with_persisted(file_b_upload, "persisted_primary_fingerprint"):
        return True
    if _upload_conflicts_with_persisted(mapping_upload, "persisted_mapping_fingerprint"):
        return True
    return False


def _clear_recon_session_results() -> None:
    """Drop cached reconciliation output after file changes or an explicit user reset."""
    for key in (
        "df_recon",
        "df_excluded_by_liability_type",
        "recon_row_counts",
        "recon_context",
        "emergency_pri_rowcount",
        "emergency_pri_probe_df",
        "last_run_excel_name",
        "last_run_csv_name",
        "last_successful_upload_signature",
        "primary_upload_name",
        "persisted_m61_fingerprint",
        "persisted_primary_fingerprint",
        "persisted_mapping_fingerprint",
        "persisted_m61_upload_name",
        "persisted_primary_upload_name",
        "persisted_primary_file_type_at_upload",
    ):
        st.session_state.pop(key, None)
    # Release memory held by now-deleted DataFrames so Render doesn't OOM.
    gc.collect()


def _bytes_from_streamlit_upload(uploaded_file) -> bytes:
    """Read all bytes from a Streamlit ``UploadedFile`` / file-like object.

    Never uses ``uploaded_file.name`` as a path — cloud hosts have no client file on disk.
    """
    if uploaded_file is None:
        return b""
    try:
        return uploaded_file.getvalue()
    except Exception:
        pass
    try:
        return uploaded_file.getbuffer().tobytes()
    except Exception:
        pass
    pos = 0
    try:
        pos = uploaded_file.tell()
    except Exception:
        pass
    try:
        uploaded_file.seek(0)
        raw = uploaded_file.read()
        try:
            uploaded_file.seek(pos)
        except Exception:
            pass
        if isinstance(raw, (bytes, bytearray, memoryview)):
            return bytes(raw)
        return bytes(raw or b"")
    except Exception as exc:
        raise TypeError(f"Could not read bytes from upload: {exc!r}") from exc


def _temp_workbook_path_inside_dir(
    tmpdir: str,
    *,
    stem: str,
    client_filename: str | None,
    default_ext: str,
) -> str:
    """Build ``tmpdir/{stem}{ext}`` using only a *safe* extension parsed from ``client_filename``."""
    ext = default_ext if str(default_ext).startswith(".") else f".{default_ext}"
    if client_filename:
        cand = os.path.splitext(client_filename)[1].strip().lower()
        if cand in _UPLOAD_EXTS_ALLOWED:
            ext = cand
    return os.path.join(tmpdir, f"{stem}{ext}")


def _write_upload_to_disk(path: str, uploaded_file) -> None:
    """Write upload bytes to ``path`` (real temp file) and flush for cloud filesystems."""
    data = _bytes_from_streamlit_upload(uploaded_file)
    with open(path, "wb") as out:
        out.write(data)
        out.flush()
        try:
            os.fsync(out.fileno())
        except OSError:
            pass
    # Free the in-memory bytes buffer immediately — the file on disk is the only
    # copy needed for reconcile().  On Render, this matters for large workbooks.
    del data


def _current_upload_signature(
    file_a_upload, file_b_upload, mapping_upload, primary_file_type: str
):
    """Uniquely identify a run: primary template + stable file fingerprints + optional mapping.

    Including ``primary_file_type`` ensures that after e.g. an AOC II run, switching the sidebar
    to AOC I with the same uploads still invalidates the last successful signature so
    **Auto-run on upload** can refresh results (Fund column and scope follow the new template).
    """
    fa = _upload_fingerprint(file_a_upload)
    fb = _upload_fingerprint(file_b_upload)
    if not fa or not fb:
        return None
    fm = _upload_fingerprint(mapping_upload) if mapping_upload else None
    return (
        str(primary_file_type).strip(),
        fa,
        fb,
        fm,
    )


def _reset_table_filter_state() -> None:
    """Clear persisted table/grid filter widget state after each successful run."""
    prefixes = (
        "recon_tbl_primary_ms_",
        "recon_tbl_adv_ms_",
        "recon_tbl_sort_",
        "recon_tbl_notecat_",
    )
    for k in list(st.session_state.keys()):
        if isinstance(k, str) and k.startswith(prefixes):
            del st.session_state[k]
    st.session_state["recon_hide_blank_cols"] = False
    st.session_state["recon_deal_pick"] = "All deals"
    # Cannot set ``recon_m61_note_category`` here: sidebar selectbox may already exist this run.
    # Apply default on the *next* run before the widget is created (see sidebar Filters block).
    st.session_state["recon_pending_m61_note_category_reset"] = True


def run_reconciliation_for_selection(
    file_a_upload,
    file_b_upload,
    primary_file_type: str,
    mapping_upload=None,
):
    with st.spinner("Running reconciliation…"):
        tmp_parent = tempfile.gettempdir()
        with tempfile.TemporaryDirectory(prefix="recon_upload_", dir=tmp_parent) as tmpdir:
            path_a = _temp_workbook_path_inside_dir(
                tmpdir,
                stem="recon_m61",
                client_filename=getattr(file_a_upload, "name", None),
                default_ext=".xlsx",
            )
            path_b = _temp_workbook_path_inside_dir(
                tmpdir,
                stem="recon_primary",
                client_filename=getattr(file_b_upload, "name", None),
                default_ext=".xlsm",
            )
            path_map = None
            _write_upload_to_disk(path_a, file_a_upload)
            _write_upload_to_disk(path_b, file_b_upload)
            for _label, _path in (("M61", path_a), ("ACORE", path_b)):
                if not os.path.isfile(_path) or os.path.getsize(_path) <= 0:
                    raise OSError(
                        f"{_label} upload could not be written to a temporary workbook "
                        f"(path={_path!r}). Try uploading again."
                    )
            if primary_file_type in ("AOC II", "AOC I") and mapping_upload:
                path_map = _temp_workbook_path_inside_dir(
                    tmpdir,
                    stem="recon_mapping",
                    client_filename=getattr(mapping_upload, "name", None),
                    default_ext=".xlsx",
                )
                _write_upload_to_disk(path_map, mapping_upload)
            _bad_pri_fn, _msg_pri_fn = primary_filename_incompatible_acp_ii_vs_iii(
                getattr(file_b_upload, "name", None), primary_file_type
            )
            if _bad_pri_fn:
                st.error(_msg_pri_fn)
                st.stop()
            df_recon, _, df_excluded_type, recon_diag = reconcile(
                path_a,
                path_b,
                primary_file_type=primary_file_type,
                mapping_path=path_map,
                uploaded_primary_filename=file_b_upload.name,
                return_diagnostics=True,
                match_diagnostics=bool(
                    st.session_state.get("recon_match_diagnostics", False)
                ),
            )
            df_recon = normalize_recon_fund_for_output(df_recon)
            # Drop the previous run's large DataFrames before storing new ones so
            # Python's allocator can reclaim the memory on Render.
            for _stale_key in ("df_recon", "df_excluded_by_liability_type"):
                st.session_state.pop(_stale_key, None)
            gc.collect()
            st.session_state["df_recon"] = df_recon
            st.session_state["df_excluded_by_liability_type"] = (
                df_excluded_type if df_excluded_type is not None else pd.DataFrame()
            )
            st.session_state["recon_row_counts"] = dict(recon_diag or {})
            st.session_state["recon_context"] = get_last_recon_context()
            st.session_state["primary_file_type"] = primary_file_type
            st.session_state["primary_upload_name"] = file_b_upload.name
            # Excel is built at download time from the same filtered view as the table.
            # Persist exact download names from this successful run context.
            st.session_state["last_run_excel_name"] = build_output_filename(
                primary_file_type, "xlsx", uploaded_filename=file_b_upload.name
            )
            st.session_state["last_run_csv_name"] = build_output_filename(
                primary_file_type, "csv", uploaded_filename=file_b_upload.name
            )
            # Last successful run context for stale-state checks.
            st.session_state["last_successful_upload_signature"] = _current_upload_signature(
                file_a_upload, file_b_upload, mapping_upload, primary_file_type
            )
            _update_persisted_upload_fingerprints(
                file_a_upload, file_b_upload, mapping_upload, primary_file_type
            )
            # Defer status reset to pre-widget stage (Streamlit-safe session_state mutation).
            st.session_state["reset_status_filters"] = True
            _reset_table_filter_state()
 
 
# --------------------------------------------------
# MAIN CONTENT
# --------------------------------------------------
both_uploads_ready = _both_uploads_ready(file_a_upload, file_b_upload)

_update_persisted_upload_fingerprints(
    file_a_upload, file_b_upload, mapping_upload, primary_file_type
)
if _workspace_should_be_cleared(
    file_a_upload, file_b_upload, mapping_upload, primary_file_type
):
    _clear_recon_session_results()
elif file_a_upload and not m61_file_valid:
    st.warning(
        "Uploaded comparison file does not look like a Liability Relationship export. "
        "Expected filename to include both `liability` and `relationship` "
        "(and not mapping names like `LiabilityNote` / `AssetNote`). "
        "You can still run reconciliation."
    )

has_required_uploads = both_uploads_ready

manual_run_requested = st.button("▶  Run Reconciliation", type="primary")
upload_signature = _current_upload_signature(
    file_a_upload, file_b_upload, mapping_upload, primary_file_type
)
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
    _run_primary_raw = st.session_state.get("primary_file_type", primary_file_type)
    run_primary = canonical_primary_file_type(_run_primary_raw)
    if _run_primary_raw != run_primary:
        st.session_state["primary_file_type"] = run_primary
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
    # ---- EMERGENCY DEBUG (disabled) — Primary Fin Inpt load vs df_recon (1551 Broadway / Deal ID 25-2852) ----
    #     with st.expander(
    #         "EMERGENCY DEBUG — Primary Fin Inpt load vs `df_recon` (1551 Broadway / Deal ID 25-2852)",
    #         expanded=True,
    #     ):
    #         st.caption(
    #             "Compare **load_primary_file** (same temp path as reconcile) to reconciliation output. "
    #             "If loaded rows > ACORE-side `df_recon` rows, the loss is in **reconcile**. If counts match here "
    #             "but **Total ACORE Records** is lower, the loss is in **this page’s filters** (`df_all` / scope / "
    #             "Show Records / note category)."
    #         )
    #         c_a, c_b, c_c = st.columns(3)
    #         with c_a:
    #             st.markdown(f"**Sidebar primary type (widget):** `{primary_file_type}`")
    #             st.markdown(f"**Last run primary:** `{_run_primary_raw}` → canonical **`{run_primary}`**")
    #         with c_b:
    #             st.markdown(f"**Uploaded primary file:** `{run_primary_upload_name or '—'}`")
    #         with c_c:
    #             _n_pri_em = st.session_state.get("emergency_pri_rowcount")
    #             _rc = st.session_state.get("recon_row_counts") or {}
    #             _n_raw_diag = _rc.get("raw_primary_rows_loaded")
    #             st.markdown(f"**`load_primary_file` row count:** `{_n_pri_em if _n_pri_em is not None else '—'}`")
    #             st.markdown(
    #                 f"**`reconcile` diag `raw_primary_rows_loaded`:** `{_n_raw_diag if _n_raw_diag is not None else '—'}`"
    #             )
    #         st.markdown(
    #             "**Rows from `load_primary_file` where Deal Name contains `1551 Broadway` OR Deal ID contains `25-2852`:**"
    #         )
    #         _em_df = st.session_state.get("emergency_pri_probe_df")
    #         if _em_df is not None and not getattr(_em_df, "empty", True):
    #             st.dataframe(_em_df, use_container_width=True, height=min(320, 40 + 28 * len(_em_df)))
    #         else:
    #             st.info("No probe rows in session (re-run reconciliation after upload), or no rows matched the probe.")
    #         st.markdown("**`df_recon` rows where Deal Name contains `1551 Broadway`:**")
    #         if "Deal Name" in df_recon.columns:
    #             _m1551 = df_recon["Deal Name"].astype(str).str.contains("1551 Broadway", case=False, na=False)
    #             _rec_cols = [
    #                 c
    #                 for c in (
    #                     "Deal Name",
    #                     "Deal ID (ACP)",
    #                     "Effective Date (ACP)",
    #                     "Facility",
    #                     "Source",
    #                     "File Source",
    #                     "recon_status",
    #                     "Match Stage",
    #                     "_row_id_b",
    #                 )
    #                 if c in df_recon.columns
    #             ]
    #             _rec_sub = df_recon.loc[_m1551, _rec_cols] if _rec_cols else df_recon.loc[_m1551]
    #             st.dataframe(
    #                 _rec_sub, use_container_width=True, height=min(360, 40 + 28 * max(len(_rec_sub), 1))
    #             )
    #             st.caption(f"**{int(_m1551.sum())}** row(s) on full `df_recon` for this deal name filter.")
    #         else:
    #             st.warning("`df_recon` has no `Deal Name` column.")
    #         if "File Source" in df_recon.columns:
    #             _n_acore_recon = int(
    #                 df_recon["File Source"]
    #                 .fillna("")
    #                 .astype(str)
    #                 .str.strip()
    #                 .isin([FILE_SOURCE_BOTH, FILE_SOURCE_ACORE_ONLY])
    #                 .sum()
    #             )
    #         else:
    #             _n_acore_recon = int(len(df_recon))
    #         st.markdown(
    #             f"**ACORE-side rows on full `df_recon` (Both + ACORE Only):** **`{_n_acore_recon}`** — compare to "
    #             f"**`load_primary_file`** count **`{_n_pri_em if _n_pri_em is not None else '—'}`**."
    #         )
    #         if _n_pri_em is not None and _n_pri_em != _n_acore_recon:
    #             st.error(
    #                 f"**Row count mismatch:** `load_primary_file` returned **`{_n_pri_em}`** rows but **`df_recon`** "
    #                 f"only has **`{_n_acore_recon}`** ACORE-side rows (Both + ACORE Only). The gap is in **reconcile** "
    #                 "(or diagnostics skew if `raw_primary_rows_loaded` disagrees). If counts match here but the "
    #                 "dashboard **Total ACORE Records** is still lower, the loss is in **filters on this page**."
    #             )
    #         if _n_raw_diag is not None and _n_pri_em is not None and int(_n_raw_diag) != int(_n_pri_em):
    #             st.warning(
    #                 f"`reconcile` reported `raw_primary_rows_loaded={_n_raw_diag}` but this page’s "
    #                 f"`load_primary_file` probe counted **`{_n_pri_em}`** — check for version/path skew."
    #             )

    st.markdown('<div class="section-label">Deal filter</div>', unsafe_allow_html=True)

    _debug_full = bool(st.session_state.get("recon_debug_full_m61", False))
    # filter_recon_to_selected_fund receives a copy so the session-state df_recon
    # is never mutated by the engine's internal filtering.
    df_scoped = filter_recon_to_selected_fund(df_recon.copy(), run_primary)
    if _debug_full:
        df_all = df_recon   # read-only reference in debug mode; no copy needed
    else:
        df_all = df_scoped  # df_scoped is already a fresh filtered DataFrame

    if "File Source" in df_all.columns:
        _deal_source_mask = df_all["File Source"].fillna("").astype(str).str.strip().isin(
            [FILE_SOURCE_BOTH, FILE_SOURCE_ACORE_ONLY]
        )
        _deal_pool = df_all.loc[_deal_source_mask]   # view; only used for .unique()
    else:
        _deal_pool = df_all
    deal_names = (
        sorted(_deal_pool["Deal Name"].dropna().astype(str).unique().tolist())
        if "Deal Name" in _deal_pool.columns
        else []
    )
    deal_options = ["All deals"] + deal_names
    if st.session_state.get("recon_deal_pick") not in deal_options:
        st.session_state["recon_deal_pick"] = "All deals"
    deal_pick = st.selectbox(
        "Deal name",
        options=deal_options,
        index=0,
        key="recon_deal_pick",
        help="Type in the box to jump to a deal (Streamlit search). Choose **All deals** to clear.",
    )

    # Scope (finance): df_all is fund-scoped. Selected view = primary-tied rows + M61-only rows
    # that already passed fund scope (same df_all), not "all M61-only" from other funds.
    if "File Source" in df_all.columns:
        _fs_scope = df_all["File Source"].fillna("").astype(str).str.strip()
        in_scope_ix_primary_tied = set(
            df_all.index[
                _fs_scope.isin([FILE_SOURCE_BOTH, FILE_SOURCE_ACORE_ONLY])
            ].tolist()
        )
        in_scope_ix_m61_only_fund_scoped = set(
            df_all.index[_fs_scope.eq(FILE_SOURCE_M61_ONLY)].tolist()
        )
    else:
        in_scope_ix_primary_tied = set(df_all.index)
        in_scope_ix_m61_only_fund_scoped = set()
    in_scope_ix_selected_fund_view = in_scope_ix_primary_tied | in_scope_ix_m61_only_fund_scoped

    if _debug_full:
        df_view = df_recon.copy()
        st.warning(
            "**Developer view:** Showing every fund in this run. Use standard Scope when you return to the main workflow."
        )
    elif scope_mode == "Selected Fund Only":
        # ACORE is the row driver: show only rows where ACORE data is present (Both or ACORE Only).
        # M61-only rows are intentionally excluded — the ACORE file defines the row universe.
        df_view = df_all.loc[df_all.index.isin(in_scope_ix_primary_tied)].copy()
        st.info(
            "**Selected Primary Fund Only:** Shows every ACORE row for this fund with M61 comparison "
            "columns filled in where a match exists. M61-only rows are hidden — ACORE drives the row count."
        )
    else:
        # No copy — df_view is immediately reassigned by every filter step below
        # (deal filter, date filter, note category, status filter all use .loc[] or pd.concat
        # which produce new DataFrames, leaving df_all untouched).
        df_view = df_all
        st.caption(
            f"**All results for {run_primary_label}:** Every record for this fund—matches, gaps, and export-only lines. "
            "Other funds are not shown."
        )
    # _td_active snapshots disabled — set to None so no DataFrame copies are made.
    _td_active = False
    _td_after_scope = None

    # Deal filter is applied after scope (on the scoped/unscoped base view).
    if deal_pick and deal_pick != "All deals":
        df_view = df_view[df_view["Deal Name"] == deal_pick]
    _td_after_deal = None

    # Effective date range (display-only; sidebar Advanced filters). Applied before note category / metrics.
    _eff_preset = str(st.session_state.get("recon_eff_date_preset") or "All dates").strip()
    _eff_bounds = resolve_effective_date_range_bounds(
        _eff_preset,
        st.session_state.get("recon_eff_date_custom_start"),
        st.session_state.get("recon_eff_date_custom_end"),
    )
    _td_after_eff = None
    # Always keep rows where ACORE/primary data is present (File Source = Both or ACORE Only).
    # The effective-date range filter is for display convenience; it must never hide primary model rows.
    _primary_row_ix = (
        set(
            df_view.index[
                df_view["File Source"].fillna("").astype(str).str.strip()
                .isin([FILE_SOURCE_BOTH, FILE_SOURCE_ACORE_ONLY])
            ].tolist()
        )
        if "File Source" in df_view.columns
        else set()
    )
    df_view = filter_display_dataframe_by_effective_dates(df_view, *_eff_bounds)
    # Restore any primary rows the date filter dropped (preserves all ACORE Fin Inpt rows).
    if _primary_row_ix:
        _dropped_primary = _primary_row_ix - set(df_view.index.tolist())
        if _dropped_primary:
            _pre_filter_view = (
                df_all.loc[df_all.index.isin(in_scope_ix_primary_tied)]
                if "File Source" in df_all.columns
                else df_all
            )
            if deal_pick and deal_pick != "All deals":
                _pre_filter_view = _pre_filter_view[_pre_filter_view["Deal Name"] == deal_pick]
            _restore = _pre_filter_view.loc[
                _pre_filter_view.index.isin(_dropped_primary)
            ]
            if not _restore.empty:
                df_view = pd.concat([df_view, _restore]).sort_index()
    _td_after_eff = None

    # Read note-category selection before source-type narrowing.
    _note_pick = str(st.session_state.get("recon_m61_note_category", "Financing") or "Financing").strip()

    # M61 Note Category (sidebar): same series as m61_note_category_series_for_ui (single normalize pass).
    _before_note_filter = len(df_view)
    _raw_m61_note_cat = (
        df_view["M61 Note Category"].fillna("Other").astype(str)
        if "M61 Note Category" in df_view.columns
        else pd.Series(dtype=str)
    )
    _note_unique_raw = sorted(
        {unicodedata.normalize("NFKC", str(x)).strip() for x in _raw_m61_note_cat.tolist() if str(x).strip()}
    )
    _selected_note_norm = normalize_m61_note_category_label(_note_pick)
    _note_series_pre = m61_note_category_series_for_ui(df_view, primary_file_type=run_primary)
    _note_unique_norm = sorted({_ for _ in _note_series_pre.unique().tolist() if str(_).strip()})
    if _selected_note_norm != "all":
        _keep = _note_series_pre.eq(_selected_note_norm)
        df_view = df_view.loc[_keep].copy()
    _after_note_filter = len(df_view)

    if False:  # TEMP DEBUG — M61 Note Category filter (diagnostic); set True to re-enable
        with st.expander("TEMP DEBUG — M61 Note Category filter (diagnostic)", expanded=False):
            st.caption(f"selected_raw={_note_pick!r} | selected_norm={_selected_note_norm!r}")
            st.caption(f"unique_raw_M61_Note_Category (before note filter)={_note_unique_raw!r}")
            st.caption(f"unique_normalized (before note filter)={_note_unique_norm!r}")
            st.caption(
                f"rows_before_note_filter={_before_note_filter} | rows_after_note_filter={_after_note_filter}"
            )
            if not df_view.empty and "M61 Note Category" in df_view.columns:
                st.caption("value_counts M61 Note Category (after note filter, before status filter)")
                _vc_note = (
                    df_view["M61 Note Category"]
                    .fillna("<NA>")
                    .astype(str)
                    .value_counts(dropna=False)
                    .rename_axis("M61 Note Category")
                    .reset_index(name="rows")
                )
                st.dataframe(
                    _vc_note,
                    use_container_width=True,
                    hide_index=True,
                    height=min(240, 40 + 28 * max(len(_vc_note), 1)),
                )
    _td_after_note_cat = None

    # Apply status filters on the current view (same categories as before; gated by Show Records UI).
    review_filter = []
    if st.session_state.get("filter_review_clean_matches", False):
        review_filter.append("Clean Matches")
    _needs_bundle_on = st.session_state.get("filter_needs_review_bundle", True)
    if _needs_bundle_on:
        if st.session_state.get("filter_review_missing_fields", True):
            review_filter.append("Matches with Missing Fields")
        if st.session_state.get("filter_review_differences", True):
            review_filter.append("Differences / Mismatches")
        if st.session_state.get("filter_review_missing_m61", True):
            review_filter.append("Missing in M61")
        if st.session_state.get("filter_review_missing_acore", True):
            review_filter.append("Missing in ACORE")
    if review_filter:
        rs_upper = (
            df_view["recon_status"].fillna("").astype(str).str.upper().str.strip()
            if "recon_status" in df_view.columns
            else pd.Series([""] * len(df_view), index=df_view.index, dtype="object")
        )
        fs = (
            df_view["File Source"].fillna("").astype(str).str.strip()
            if "File Source" in df_view.columns
            else pd.Series([""] * len(df_view), index=df_view.index, dtype="object")
        )
        mask = pd.Series(False, index=df_view.index)
        # Field-level status columns used for "Differences / Mismatches" grouping.
        _field_status_cols = [
            "Advance Rate Status",
            "Spread Status",
            "Effective Date Status",
            "Undrawn Capacity Status",
            "Index Floor Status",
            "Index Name Status",
            "Recourse % Status",
            "Pledge Date Status",
        ]
        _has_field_mismatch = pd.Series(False, index=df_view.index)
        for _c in _field_status_cols:
            if _c in df_view.columns:
                _has_field_mismatch |= (
                    df_view[_c].fillna("").astype(str).str.upper().str.contains("MISMATCH", regex=False, na=False)
                )
        if "Clean Matches" in review_filter:
            mask |= rs_upper.eq("MATCH")
        if "Matches with Missing Fields" in review_filter:
            mask |= rs_upper.str.contains("MATCH WITH MISSING FIELDS", regex=False, na=False)
        if "Differences / Mismatches" in review_filter:
            mask |= rs_upper.str.contains("MATCH WITH DIFFERENCES", regex=False, na=False) | _has_field_mismatch
        if "Missing in M61" in review_filter:
            mask |= rs_upper.str.contains("MISSING IN M61", regex=False, na=False) | fs.eq(FILE_SOURCE_ACORE_ONLY)
        if "Missing in ACORE" in review_filter:
            mask |= rs_upper.str.contains("MISSING IN ACORE", regex=False, na=False) | fs.eq(FILE_SOURCE_M61_ONLY)
        df_view = df_view.loc[mask].copy()
    else:
        df_view = df_view.iloc[0:0]
    _after_status_filter = len(df_view)
    _td_after_status = None

    # For selected primary types: when Note Category is not All, drop M61-only rows (same as prior default
    # with the removed "Show M61-only exceptions" checkbox always off).
    _td_after_m61_hide = None  # TEMP DEBUG default
    if run_primary in ("ACP III", "AOC II", "AOC I"):
        # When Note Category = All, show full universe (including M61-only rows).
        hide_m61_only = _note_pick != "All"
        if hide_m61_only and "File Source" in df_view.columns:
            _before = len(df_view)
            df_view = df_view.loc[df_view["File Source"].fillna("").astype(str).str.strip().ne("M61 Only")].copy()
            _hidden = _before - len(df_view)
            if _hidden > 0:
                st.caption(f"Hidden M61-only exceptions: {_hidden}")
        _td_after_m61_hide = None

    _displayed_rows_final = len(df_view)
    _note_cat_m61_only_hidden_hint = (
        run_primary in ("ACP III", "AOC II", "AOC I")
        and _after_note_filter > 0
        and _after_status_filter > 0
        and _displayed_rows_final == 0
        and _note_pick != "All"
    )

    # Avoid showing non-contiguous / upstream row positions in the index column (confused with deal IDs).
    df_view = df_view.reset_index(drop=True)
    if _trace_1551_ui_enabled():
        _trace_1551_broadway_ui_stderr("1. df_recon", df_recon)
        _trace_1551_broadway_ui_stderr("2. df_scoped", df_scoped)
        _trace_1551_broadway_ui_stderr("3. df_all", df_all)
        _trace_1551_broadway_ui_stderr("4. df_view after scope", _td_after_scope)
        _trace_1551_broadway_ui_stderr("5. df_view after deal filter", _td_after_deal)
        _trace_1551_broadway_ui_stderr("6. df_view after effective date filter", _td_after_eff)
        _trace_1551_broadway_ui_stderr("7. df_view after M61 Note Category filter", _td_after_note_cat)
        _trace_1551_broadway_ui_stderr("8. df_view after Show Records / status filter", _td_after_status)
        _td_m61_hide_or_noop = _td_after_m61_hide if _td_after_m61_hide is not None else _td_after_status
        _trace_1551_broadway_ui_stderr("9. df_view after M61-only hide (or no-op if not ACP III/AOC)", _td_m61_hide_or_noop)
        _trace_1551_broadway_ui_stderr("10. final df_view after reset_index", df_view)
    _target_tracking_rows: list[dict[str, object]] = []
    _diag_target = st.session_state.get("recon_row_counts", {}) or {}
    _s1 = pd.DataFrame(_diag_target.get("target_22203_stage1_rows", []) or [])
    _s2 = pd.DataFrame(_diag_target.get("target_22203_stage2_rows", []) or [])
    _target_tracking_rows.extend(_target_22203_stage_rows("1. after reconciliation output", _s1))
    _target_tracking_rows.extend(_target_22203_stage_rows("2. after related-M61 enhancement", _s2))
    _target_tracking_rows.extend(_target_22203_stage_rows("3. after scope filter", _td_after_scope if _td_after_scope is not None else df_view))
    _target_tracking_rows.extend(_target_22203_stage_rows("4. after Show Records filter", _td_after_status if _td_after_status is not None else df_view))

    if False and run_primary == "ACORE":
        diag_counts = st.session_state.get("recon_row_counts", {}) or {}
        with st.expander("ACP III row count validation", expanded=False):
            st.caption(
                "ACP III baseline pipeline counters to explain why raw M61 rows can differ from final reconciliation rows."
            )
            d1, d2, d3, d4 = st.columns(4)
            with d1:
                st.metric(
                    "Raw ACP III rows loaded",
                    int(diag_counts.get("raw_primary_rows_loaded", 0)),
                )
            with d2:
                st.metric(
                    "Raw M61 rows loaded",
                    int(diag_counts.get("raw_m61_rows_loaded", 0)),
                )
            with d3:
                st.metric(
                    "M61 rows after ACP III fund filter",
                    int(
                        diag_counts.get(
                            "m61_rows_after_fund_filter_for_primary",
                            diag_counts.get("m61_rows_after_filters", 0),
                        )
                    ),
                )
            with d4:
                st.metric(
                    "M61 rows after liability-type filter",
                    int(diag_counts.get("m61_rows_after_filters", 0)),
                )

            d5, d6, d7 = st.columns(3)
            with d5:
                st.metric(
                    "ACP III rows after exclusions",
                    int(
                        diag_counts.get(
                            "primary_rows_after_exclusions",
                            diag_counts.get("raw_primary_rows_loaded", 0),
                        )
                    ),
                )
            with d6:
                st.metric(
                    "Final reconciliation rows",
                    int(diag_counts.get("final_reconciliation_rows", len(df_all))),
                )
            with d7:
                st.metric("Displayed rows", int(len(df_view)))

            excl = int(diag_counts.get("m61_rows_excluded_by_type_filter", 0))
            st.caption(
                f"M61 rows excluded by liability-type filter: **{excl}** "
                "(current in-scope types: Repo / Non / Subline)."
            )
            st.caption(
                "Displayed rows include current scope + sidebar status filters + Deal filter. "
                "Use this to compare with final reconciliation output."
            )
            basis = str(
                diag_counts.get("reconciliation_basis", "outer_merge_preserving_both_files")
            )
            if basis == "outer_merge_preserving_both_files":
                st.info(
                    "Final reconciliation rows come from an **outer merge preserving both ACP III and M61 sides** "
                    "(matched rows + ACP-only rows + M61-only rows)."
                )
            st.write("Temporary diagnostics payload", diag_counts)

    # TEMP validation only (remove when done debugging fund-scope vs primary-key-scope).
    _val_cols = [
        c
        for c in (
            "Fund",
            "Deal Name",
            "Facility",
            "Financial Line",
            "File Source",
            "recon_status",
        )
        if c in df_recon.columns
    ]

    def _sample_df(df: pd.DataFrame) -> pd.DataFrame:
        if df is None or df.empty:
            return df
        use = [c for c in _val_cols if c in df.columns]
        return df.head(10)[use] if use else df.head(10)

    if False:  # TEMP diagnostics block hidden
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
        if "File Source" in df_recon.columns:
            c1, c2, c3 = st.columns(3)
            with c1:
                st.caption("File Source mix: df_recon")
                st.dataframe(
                    df_recon["File Source"]
                    .fillna("<NA>")
                    .astype(str)
                    .value_counts(dropna=False)
                    .rename("rows")
                    .to_frame(),
                    use_container_width=True,
                    height=160,
                )
            with c2:
                st.caption("File Source mix: df_scoped")
                st.dataframe(
                    df_scoped["File Source"]
                    .fillna("<NA>")
                    .astype(str)
                    .value_counts(dropna=False)
                    .rename("rows")
                    .to_frame(),
                    use_container_width=True,
                    height=160,
                )
            with c3:
                st.caption("File Source mix: df_view")
                st.dataframe(
                    df_view["File Source"]
                    .fillna("<NA>")
                    .astype(str)
                    .value_counts(dropna=False)
                    .rename("rows")
                    .to_frame(),
                    use_container_width=True,
                    height=160,
                )

    if False:  # validation block hidden
        pass
 
    # ---- Metric cards (ACORE-backed rows in the *current table view*, not df_all) ----
    # Must match visible Fin Inpt rows: same scope/deal/date/note/status filters as df_view, including
    # primary rows restored after the effective-date display filter. Do not dedupe by deal or facility.
    if "File Source" in df_view.columns:
        _fs_dash = df_view["File Source"].fillna("").astype(str).str.strip()
        df_view_acore_backed = df_view.loc[
            _fs_dash.isin([FILE_SOURCE_BOTH, FILE_SOURCE_ACORE_ONLY])
        ].copy()
    else:
        df_view_acore_backed = df_view.copy()

    n_total_acore = int(len(df_view_acore_backed))
    if n_total_acore and "recon_status" in df_view_acore_backed.columns:
        _fs_ac = df_view_acore_backed["File Source"].fillna("").astype(str).str.strip()
        _rs_ac = df_view_acore_backed["recon_status"].fillna("").astype(str).str.strip().str.upper()
        _m_missing_m61 = _fs_ac.eq(FILE_SOURCE_ACORE_ONLY) | _rs_ac.str.contains(
            "MISSING IN M61", regex=False, na=False
        )
        _m_clean_match = _rs_ac.eq("MATCH")
        n_missing_m61_dash = int(_m_missing_m61.sum())
        n_match_dash = int((~_m_missing_m61 & _m_clean_match).sum())
        n_needs_review_dash = int((~_m_missing_m61 & ~_m_clean_match).sum())
        match_rate_acore = (n_match_dash / n_total_acore * 100) if n_total_acore else 0.0
        _df_needs_review_rows = df_view_acore_backed.loc[~_m_missing_m61 & ~_m_clean_match].copy()
    else:
        n_missing_m61_dash = 0
        n_match_dash = 0
        n_needs_review_dash = 0
        match_rate_acore = 0.0
        _df_needs_review_rows = pd.DataFrame()

    match_subtitle = f"{match_rate_acore:.0f}% of ACORE records fully match"
    needs_subtitle = "Mismatches, incomplete fields, or other review (paired rows)"
    miss_m61_subtitle = "ACORE rows with no M61 match"

    st.caption(
        "Table reflects your **Show Records** filter, deal, dates, and note category. "
        f"**{len(df_view)}** row(s) in the current view."
    )

    c1, c2, c3, c4 = st.columns(4)
    with c1:
        st.markdown(
            f"""
        <div class="metric-card mc-total">
          <div class="label">Total ACORE Records</div>
          <div class="value">{n_total_acore}</div>
          <div class="sub">ACORE-backed rows in current view (Both + ACORE Only)</div>
        </div>""",
            unsafe_allow_html=True,
        )
    with c2:
        st.markdown(
            f"""
        <div class="metric-card mc-match">
          <div class="label">✓ Matches</div>
          <div class="value">{n_match_dash}</div>
          <div class="sub">{match_subtitle}</div>
        </div>""",
            unsafe_allow_html=True,
        )
    with c3:
        st.markdown(
            f"""
        <div class="metric-card mc-needs">
          <div class="label">⚠ Needs Review</div>
          <div class="value">{n_needs_review_dash}</div>
          <div class="sub">{needs_subtitle}</div>
        </div>""",
            unsafe_allow_html=True,
        )
    with c4:
        st.markdown(
            f"""
        <div class="metric-card mc-missing">
          <div class="label">○ Missing in M61</div>
          <div class="value">{n_missing_m61_dash}</div>
          <div class="sub">{miss_m61_subtitle}</div>
        </div>""",
            unsafe_allow_html=True,
        )

    st.caption(
        "Metrics count **ACORE-backed rows** in the **same filtered view** as the table."
    )

    st.session_state.pop("selected_metric", None)

    _card_explain_options = [
        "Total ACORE Records",
        "Matches",
        "Needs Review",
        "Missing in M61",
    ]
    _selected_explain = st.radio(
        "What would you like explained?",
        options=_card_explain_options,
        index=None,
        horizontal=True,
        key="dashboard_metric_explain_radio",
    )
    if _selected_explain:
        _insight = _card_contextual_insight(
            _selected_explain,
            n_total_acore,
            n_match_dash,
            n_needs_review_dash,
            n_missing_m61_dash,
            _df_needs_review_rows,
        )
        if _insight:
            st.markdown(_insight)

    # Debug expanders intentionally hidden in normal UI.

    # Validation/debug expanders hidden from submission UI.

    if False:  # Temporary Adv Rate debug hidden
        with st.expander("Temporary Adv Rate row debug", expanded=False):
            dbg_cols = [
                "Deal Name",
                "Liability Note (M61)",
                "Liability Type (M61 Raw)",
                "Target Advance Rate (M61)",
                "Current Advance Rate (M61 Raw)",
                "Deal Level Advance Rate (M61 Raw)",
                "Advance Rate (M61)",
                "Advance Rate Source (M61)",
            ]
            dbg_cols_present = [c for c in dbg_cols if c in df_view.columns]
            if not dbg_cols_present:
                st.caption("No advance-rate debug columns available in current view.")
            else:
                st.dataframe(
                    df_view.loc[:, dbg_cols_present],
                    use_container_width=True,
                    height=260,
                )
 
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

    # ---- TEMP DEBUG: Backend vs UI row counts (ACP III / AOC II only) ----
    # Remove this entire block when debugging is complete.
    if _td_active:
        def _td_vc(df, col):
            """Return value_counts DataFrame for col, or empty frame if missing."""
            if df is None or df.empty or col not in df.columns:
                return pd.DataFrame(columns=[col, "rows"])
            vc = (
                df[col].fillna("<NA>").astype(str).value_counts(dropna=False)
                .rename("rows").reset_index().rename(columns={"index": col})
            )
            return vc

        def _td_stage_summary(label, df):
            """Return a one-row dict with row count for a pipeline stage."""
            return {"Stage": label, "Row Count": 0 if df is None else len(df)}

        _td_stages = [
            ("1. df_all (raw backend)", df_all),
            ("2. df_scoped (fund/BL scope filter)", df_scoped),
            ("3. df_view after scope mode", _td_after_scope),
            ("4. df_view after deal filter", _td_after_deal),
            ("5. df_view after note category filter", _td_after_note_cat),
            ("6. df_view after status filter", _td_after_status),
            ("7. df_view after M61-only hide", _td_after_m61_hide),
            ("8. df_view final (displayed)", df_view),
        ]

        with st.expander(f"⚙️ TEMP DEBUG — Backend vs UI Row Counts ({run_primary})", expanded=False):
            st.caption(
                "TEMP DEBUG: Counts at each pipeline stage to identify where rows are gained or lost. "
                "Remove this block when debugging is complete."
            )

            # --- Row count waterfall ---
            st.markdown("**Row count at each pipeline stage**")
            _td_count_df = pd.DataFrame([_td_stage_summary(lbl, df) for lbl, df in _td_stages])
            st.dataframe(_td_count_df, use_container_width=True, hide_index=True)

            # --- Breakdowns at key stages (df_all, df_scoped, df_view final) ---
            _td_key_stages = [
                ("df_all (backend)", df_all),
                ("df_scoped (BL filter)", df_scoped),
                ("df_view after scope mode", _td_after_scope),
                ("df_view final (displayed)", df_view),
            ]

            st.markdown("**By Fund Name**")
            _td_fund_cols = []
            for lbl, sdf in _td_key_stages:
                _vc = _td_vc(sdf, "Fund Name" if "Fund Name" in (sdf.columns if sdf is not None else []) else "Fund")
                _col_name = "Fund Name" if (sdf is not None and "Fund Name" in sdf.columns) else "Fund"
                _vc = _td_vc(sdf, _col_name).rename(columns={"rows": lbl})
                if _col_name in _vc.columns:
                    _vc = _vc.set_index(_col_name)
                _td_fund_cols.append(_vc)
            if _td_fund_cols:
                try:
                    _td_fund_joined = pd.concat(_td_fund_cols, axis=1).fillna(0).astype(int)
                    st.dataframe(_td_fund_joined, use_container_width=True)
                except Exception:
                    for lbl, sdf in _td_key_stages:
                        _col_name = "Fund Name" if (sdf is not None and "Fund Name" in sdf.columns) else "Fund"
                        st.caption(lbl)
                        st.dataframe(_td_vc(sdf, _col_name), use_container_width=True, hide_index=True, height=120)

            st.markdown("**By File Source**")
            _td_src_cols = []
            for lbl, sdf in _td_key_stages:
                _vc = _td_vc(sdf, "File Source").rename(columns={"rows": lbl})
                if "File Source" in _vc.columns:
                    _vc = _vc.set_index("File Source")
                _td_src_cols.append(_vc)
            if _td_src_cols:
                try:
                    _td_src_joined = pd.concat(_td_src_cols, axis=1).fillna(0).astype(int)
                    st.dataframe(_td_src_joined, use_container_width=True)
                except Exception:
                    for lbl, sdf in _td_key_stages:
                        st.caption(lbl)
                        st.dataframe(_td_vc(sdf, "File Source"), use_container_width=True, hide_index=True, height=120)

            st.markdown("**By M61 Note Category**")
            _td_note_cols = []
            for lbl, sdf in _td_key_stages:
                _vc = _td_vc(sdf, "M61 Note Category").rename(columns={"rows": lbl})
                if "M61 Note Category" in _vc.columns:
                    _vc = _vc.set_index("M61 Note Category")
                _td_note_cols.append(_vc)
            if _td_note_cols:
                try:
                    _td_note_joined = pd.concat(_td_note_cols, axis=1).fillna(0).astype(int)
                    st.dataframe(_td_note_joined, use_container_width=True)
                except Exception:
                    for lbl, sdf in _td_key_stages:
                        st.caption(lbl)
                        st.dataframe(_td_vc(sdf, "M61 Note Category"), use_container_width=True, hide_index=True, height=120)

            st.markdown("**By Liability Type (M61 Raw)**")
            _td_lt_cols = []
            for lbl, sdf in _td_key_stages:
                _vc = _td_vc(sdf, "Liability Type (M61 Raw)").rename(columns={"rows": lbl})
                if "Liability Type (M61 Raw)" in _vc.columns:
                    _vc = _vc.set_index("Liability Type (M61 Raw)")
                _td_lt_cols.append(_vc)
            if _td_lt_cols:
                try:
                    _td_lt_joined = pd.concat(_td_lt_cols, axis=1).fillna(0).astype(int)
                    st.dataframe(_td_lt_joined, use_container_width=True)
                except Exception:
                    for lbl, sdf in _td_key_stages:
                        st.caption(lbl)
                        st.dataframe(_td_vc(sdf, "Liability Type (M61 Raw)"), use_container_width=True, hide_index=True, height=120)

            # --- Download row count ---
            st.markdown(f"**Download (df_export_ui) row count:** {len(df_view)} "
                        f"*(will equal df_view final — export is set just below)*")

            st.caption(
                "Active UI filters — "
                f"Scope: **{_scope_mode_display(scope_mode, _debug_full)}** | "
                f"Deal: **{deal_pick}** | "
                f"Note Category: **{_note_pick}** | "
                f"Show Records: **{review_filter}**"
            )
    # ---- END TEMP DEBUG ----

    # Default export matches df_view; tab1 updates this to the table row set + sort order when applicable.
    df_export_ui = df_view.copy()

    # ---- Tabs ----
    tab1, tab2 = st.tabs(["  All Results  ", "  Deal Drilldown  "])
 
    with tab1:
        st.markdown('<div class="section-label">Record-by-Record Reconciliation</div>', unsafe_allow_html=True)
        st.caption(
            "Rows are matched by Deal ID and Effective Date. If the same deal has a different "
            "effective date, it is shown separately."
        )

        if _note_cat_m61_only_hidden_hint:
            st.info(
                f"{_after_note_filter} rows match your selected M61 Note Category, but none appear here because "
                "M61-only lines are hidden when a specific category is selected. Try **All** in M61 Note Category "
                "or use **All Results for Uploaded Primary Fund** in Scope."
            )
        elif df_view.empty:
            st.info("No records match the selected filters.")
        else:
            # Build display table (aligned with RECON_ORDERED_COLS / Excel export)
            display_rows = []
            raw_undrawn_for_saa: list[object] = []
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

                def _m61_missing_by_status(status_col: str) -> bool:
                    if status_col not in row.index:
                        return False
                    s = safe_str_strip(row.get(status_col, "")).upper()
                    return ("MISSING FROM M61" in s) or ("MISSING FROM BOTH" in s)

                rec = {
                    "Fund": "" if pd.isna(row.get("Fund")) else str(row.get("Fund")),
                    "Deal Name": row.get("Deal Name", ""),
                    "Facility": row.get("Facility", ""),
                    "Financial Line": row.get("Financial Line", ""),
                    "Source Type (ACORE)": row.get("Source", ""),
                    "Liability Type (M61)": (
                        (
                            safe_str_strip(row.get("Liability Type (M61 Raw)"))
                            if pd.notna(row.get("Liability Type (M61 Raw)"))
                            else ""
                        )
                        or (
                            safe_str_strip(row.get("Liability Type (M61)"))
                            if pd.notna(row.get("Liability Type (M61)"))
                            else ""
                        )
                        or (
                            safe_str_strip(row.get("Liability Type"))
                            if pd.notna(row.get("Liability Type"))
                            else ""
                        )
                    ),
                    "M61 Note Category": (safe_str_strip(row.get("M61 Note Category")) or "Other"),
                    f"Eff Date ({col_tag})": fmt_date(ed_acp),
                    "Eff Date (M61)": fmt_date(row.get("Effective Date (M61)")),
                    f"Pledge Date ({col_tag})": fmt_date(
                        _col(row, "Pledge Date (ACP)", "Pledge Date")
                    ),
                    "Pledge Date (M61)": fmt_date(row.get("Pledge Date (M61)")),
                    f"Adv Rate ({col_tag})": pct(adv_acp),
                    "Adv Rate (M61)": (
                        "—" if _m61_missing_by_status("Advance Rate Status") else pct(row.get("Advance Rate (M61)"))
                    ),
                    f"Spread ({col_tag})": pct_spread(sp_acp),
                    "Spread (M61)": (
                        "-" if _m61_missing_by_status("Spread Status") else pct_spread(row.get("Spread (M61)"))
                    ),
                    f"Undrawn ({col_tag})": fmt_num_plain(und_acp),
                    "Undrawn (M61)": (
                        "—" if _m61_missing_by_status("Undrawn Capacity Status") else fmt_num_plain(und_liab)
                    ),
                    f"Index Floor ({col_tag})": fmt_fraction_as_pct(
                        row.get("Index Floor (ACP)"), ndigits=3
                    ),
                    "Index Floor (M61)": (
                        "—"
                        if _m61_missing_by_status("Index Floor Status")
                        else fmt_fraction_as_pct(row.get("Index Floor (M61)"), ndigits=3)
                    ),
                    f"Index Name ({col_tag})": fmt_opt_text(row.get("Index Name (ACP)")),
                    "Index Name (M61)": fmt_opt_text(row.get("Index Name (M61)")),
                    f"Recourse % ({col_tag})": pct(row.get("Recourse % (ACP)")),
                    "Recourse % (M61)": (
                        "—" if _m61_missing_by_status("Recourse % Status") else pct(row.get("Recourse % (M61)"))
                    ),
                    "File Source": _display_file_source_cell(row),
                    "Effective Date Status": _status_display(row.get("Effective Date Status", "")),
                    "Pledge Date Status": _status_display(row.get("Pledge Date Status", "")),
                    "Advance Rate Status": _status_display(row.get("Advance Rate Status", "")),
                    "Spread Status": _status_display(row.get("Spread Status", "")),
                    "Undrawn Capacity Status": _status_display(
                        row.get("Undrawn Capacity Status", "")
                    ),
                    "Index Floor Status": _status_display(row.get("Index Floor Status", "")),
                    "Index Name Status": _status_display(row.get("Index Name Status", "")),
                    "Recourse % Status": _status_display(row.get("Recourse % Status", "")),
                    # Keep full business-friendly recon summary text (no bucket/missing remap here).
                    "Overall Recon Status": (
                        ""
                        if pd.isna(row.get("recon_status"))
                        else str(row.get("recon_status", "")).strip()
                    ),
                }
                _is_target_ui = (
                    safe_str_strip(row.get("Deal Name")).lower() == "block 21 san mateo"
                    and safe_str_strip(row.get("Facility")).lower() == "tbk bank"
                    and safe_str_strip(row.get("Source")).lower() == "sale"
                    and _date_key_ui(row.get("Effective Date (ACP)")) == "2022-08-22"
                )
                if _is_target_ui:
                    print(
                        "UNDRAWN TRACE 5 (after UI formatting): "
                        f"raw_undrawn_m61={und_liab!r} "
                        f"formatted_undrawn_m61={rec.get('Undrawn (M61)')!r}"
                    )
                raw_undrawn_for_saa.append(und_liab)
                display_rows.append(rec)

            df_display = pd.DataFrame(display_rows).reset_index(drop=True)
            # "Same as Above": replace M61 value columns for ACORE Only rows that share
            # the same Deal / Source Type / Eff Date group as a matched Both row.
            df_display = _apply_same_as_above_display(
                df_display,
                col_tag,
                raw_undrawn_m61=pd.Series(raw_undrawn_for_saa, index=df_display.index),
            )

            # Option pools for **Source Type (ACORE)** / **Liability Type (M61)** table filters:
            # - Selected Fund Only → same row universe as the main table (primary-tied + fund-scoped M61-only).
            # - All Results for uploaded primary fund → all rows for that fund (same df_all as Selected for fund).
            # - Developer full-M61 → full ``df_recon`` when debug is enabled.
            # Intentionally excludes sidebar recon-status checkboxes and the Deal picker so
            # type dropdowns are scope-aware but not over-reduced vs. the displayed row set.
            # df_type_opts_base is read-only (only .map() / .apply() / .unique() called on it).
            # No copy needed — views/references are safe here.
            if _debug_full:
                df_type_opts_base = df_recon
            elif scope_mode == "Selected Fund Only":
                df_type_opts_base = df_all.loc[df_all.index.isin(in_scope_ix_primary_tied)]
            else:
                df_type_opts_base = df_all
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
            _diag = st.session_state.get("recon_row_counts", {}) or {}
            m61_type_opts_from_diag = _diag.get("m61_liability_type_values_found", []) or []

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
                # Literal ``N/A`` is substantive for Facility / source text — not "blank".
                if su == "N/A":
                    return False
                if su in ("NAN", "NONE", "—", "-", "<NA>"):
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
                enable_table_filters = st.checkbox(
                    "Enable table filters",
                    value=False,
                    key=f"recon_enable_table_filters_{col_tag}",
                    help="Off by default. Enable to open Fund/Deal/Facility/Type table filters.",
                )
                if enable_table_filters:
                    st.markdown(
                        '<div class="section-label">Table filters</div>',
                        unsafe_allow_html=True,
                    )
                    st.caption(
                        "**M61 Note Category** is set in the sidebar (with **Scope**); it already applies to this table."
                    )

                    # ── Rows 1–2: secondary filters, 3 columns each ───────────────────────────
                    # Compute option pools once (scope-aware, not narrowed by status/deal filters).
                    auto_fund_value = ""
                    auto_source_type_filter_vals = None
                    auto_liability_type_filter_val = None
                    if scope_mode == "Selected Fund Only" and not df_display.empty:
                        if "Fund" in df_display.columns:
                            fund_vals = (
                                df_display["Fund"].fillna("").astype(str).str.strip()
                            )
                            fund_vals = fund_vals[fund_vals.ne("")]
                            if not fund_vals.empty:
                                auto_fund_value = str(fund_vals.value_counts().index[0]).strip()
                        if "Source" in df_type_opts_base.columns and not df_type_opts_base.empty:
                            fam_series = df_type_opts_base["Source"].map(_acore_source_type_family)
                            fam_nonempty = fam_series[fam_series.ne("")]
                            if not fam_nonempty.empty:
                                distinct_families = set(fam_nonempty.unique().tolist())
                                if len(distinct_families) == 1:
                                    sole_family = next(iter(distinct_families))
                                    auto_source_type_filter_vals = [sole_family]
                                    lt = (
                                        liability_type_opt_series.fillna("").astype(str).str.strip()
                                    )
                                    lt = lt[lt.ne("")]
                                    if not lt.empty and lt.nunique(dropna=False) == 1:
                                        auto_liability_type_filter_val = str(lt.iloc[0]).strip()

                    # Row 1: identity — Fund, Deal Name, Facility
                    # Row 2: source — Source Type (ACORE), Liability Type (M61), File Source
                    secondary_filter_rows = [
                        ["Fund", "Deal Name", "Facility"],
                        ["Source Type (ACORE)", "Liability Type (M61)", "File Source"],
                    ]

                    for row_filters in secondary_filter_rows:
                        pf_ui_cols = st.columns(len(row_filters))
                        for i, fc in enumerate(row_filters):
                            with pf_ui_cols[i]:
                                if fc == "Source Type (ACORE)":
                                    opts_src = src_type_opt_series
                                elif fc == "Liability Type (M61)":
                                    opts_src = (
                                        pd.Series(m61_type_opts_from_diag, dtype="object")
                                        if m61_type_opts_from_diag
                                        else liability_type_opt_series
                                    )
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
                                _fc_safe = re.sub(r'\W+', '_', fc)
                                base_key = f"recon_tbl_primary_ms_{_fc_safe}_{col_tag}"
                                if fc == "Fund" and scope_mode == "Selected Fund Only":
                                    ms_key = f"{base_key}_auto_scope"
                                    if auto_fund_value and auto_fund_value in opts:
                                        if ms_key not in st.session_state or not st.session_state.get(ms_key):
                                            st.session_state[ms_key] = [auto_fund_value]
                                elif fc == "Source Type (ACORE)":
                                    ms_key = f"{base_key}_dispopts"
                                    if scope_mode == "Selected Fund Only" and auto_source_type_filter_vals:
                                        if all(v in opts for v in auto_source_type_filter_vals):
                                            if ms_key not in st.session_state or not st.session_state.get(ms_key):
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

                                _sanitize_multiselect_state(ms_key, opts)
                                selected_vals = st.multiselect(
                                    fc,
                                    options=opts,
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

                    # Debug breakdown expander intentionally hidden in normal UI.

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
            _target_tracking_rows.extend(
                _target_22203_stage_rows("5. after table filters", df_table_view)
            )

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
                st.info(
                    "No records match the selected filters. "
                    "Clear one or more table filters to see data."
                )

            # Status columns: color only (layout / padding / alignment shared via set_properties).
            def color_status(val):
                v = str(val).strip()
                # "Same as Above" rows — light blue, communicates M61 consolidation
                if v == "Same as Above":
                    return "background-color: #daeeff; color: #1a4f7a; font-style: italic;"
                v = v.upper()
                if v in ("N/A", "", "—", "-", "NAN", "NONE"):
                    return ""
                # MISSING FROM BOTH = absent on both sides, not a real issue — render muted/gray.
                # Must be checked before the general MISSING branch.
                if "MISSING FROM BOTH" in v:
                    return "background-color: #f0f0f0; color: #aaaaaa; font-style: italic;"
                if "DIFFERENCE" in v or "DIFFERENT" in v or "MISMATCH" in v or "NO MATCH" in v:
                    return "background-color: #ffc7ce; color: #9c0006; font-weight: 600;"
                if "MISSING" in v:
                    return "background-color: #ffeb9c; color: #7d6608; font-weight: 600;"
                if "MATCH" in v and "MIS" not in v:
                    return "background-color: #c6efce; color: #375623; font-weight: 600;"
                return ""

            # Single ordered list driving column visibility + styling.
            # Order: File Source → field-level statuses → Overall Recon Status (summary last).
            status_cols = [
                "File Source",
                "Effective Date Status",
                "Pledge Date Status",
                "Advance Rate Status",
                "Spread Status",
                "Undrawn Capacity Status",
                "Index Floor Status",
                "Index Name Status",
                "Recourse % Status",
                "Overall Recon Status",
            ]
            status_cols_visible = [c for c in status_cols if c in df_table_view.columns]

            # Same Streamlit width + Styler box for every status column; shorter header label
            # for Undrawn only so the column is not widened by a long title vs. peers.
            def _status_column_config(col_name: str):
                if col_name == "Overall Recon Status":
                    return st.column_config.TextColumn(
                        width="large",
                        help="Business summary of matched fields, differences, and missing fields.",
                    )
                if col_name == "File Source":
                    return st.column_config.TextColumn(width="medium")
                if col_name in (
                    "Effective Date Status",
                    "Pledge Date Status",
                    "Advance Rate Status",
                    "Spread Status",
                    "Undrawn Capacity Status",
                    "Index Floor Status",
                    "Index Name Status",
                    "Recourse % Status",
                ):
                    return st.column_config.TextColumn(width="medium")
                return st.column_config.TextColumn(width="medium")

            _status_col_cfg = {
                c: _status_column_config(c)
                for c in df_table_view.columns
                if isinstance(c, str)
                and (c.endswith(" Status") or c in ("Overall Recon Status", "File Source"))
            }

            df_table_view_display = _normalize_display_missing_df(df_table_view)
            if df_table_view_display.empty:
                styled = df_table_view_display.style
            else:
                styled = df_table_view_display.style
                _spread_cols = [c for c in df_table_view_display.columns if _is_spread_value_column(c)]
                if _spread_cols:

                    def _fmt_spread_styler(v):
                        if _is_same_as_above_label(v):
                            return SAME_AS_ABOVE_LABEL
                        if isinstance(v, str) and str(v).strip() in ("-", "—", ""):
                            return "-"
                        try:
                            if pd.isna(v):
                                return "-"
                        except (TypeError, ValueError):
                            pass
                        return pct_spread(v)

                    styled = styled.format(_fmt_spread_styler, subset=_spread_cols)
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
                # "Same as Above" highlight on M61 value columns (light blue).
                # Status columns are already handled by color_status above.
                _saa_m61_visible = [
                    c for c in _SAA_DISPLAY_M61_VALUE_COLS if c in df_table_view_display.columns
                ]
                if _saa_m61_visible:
                    def _saa_cell_style(val):
                        if str(val).strip() == "Same as Above":
                            return "background-color: #daeeff; color: #1a4f7a; font-style: italic;"
                        return ""
                    styled = styled.map(_saa_cell_style, subset=_saa_m61_visible)
            _df_kwargs: dict = {
                "use_container_width": True,
                "height": 580,
            }
            if _status_col_cfg:
                _df_kwargs["column_config"] = _status_col_cfg
            _target_tracking_rows.extend(
                _target_22203_stage_rows("6. final displayed dataframe", df_table_view)
            )
            st.caption(
                "ℹ️ M61 can include multiple liability lines per deal. The table lists every row "
                "returned by reconciliation for your current filters—nothing is collapsed or deduplicated here."
            )
            _table_sel = st.dataframe(
                styled,
                on_select="rerun",
                selection_mode="single-row",
                **_df_kwargs,
            )
            _sel_rows = []
            if isinstance(_table_sel, dict):
                _sel_rows = (
                    (_table_sel.get("selection") or {}).get("rows")
                    or []
                )
            else:
                _sel_rows = (
                    getattr(getattr(_table_sel, "selection", None), "rows", None)
                    or []
                )
            if _sel_rows:
                _sel_pos = int(_sel_rows[0])
                if 0 <= _sel_pos < len(df_table_view):
                    _sel_deal = str(df_table_view.iloc[_sel_pos].get("Deal Name", "")).strip()
                    if _sel_deal:
                        st.session_state["drilldown_deal_pick"] = _sel_deal
                    _ix = df_table_view.index[_sel_pos]
                    _explain_row = df_view.loc[_ix]
                    _disp_row = df_display.loc[_ix] if _ix in df_display.index else None
                    with st.expander("Explain this row", expanded=True):
                        st.markdown(
                            explain_reconciliation_row(_explain_row, display_row=_disp_row)
                        )

            if len(df_table_view) == 0:
                df_export_ui = df_view.iloc[0:0].copy()
            else:
                df_export_ui = df_view.loc[df_table_view.index].copy()

            # Spread MISMATCH diagnostic expander removed after percent-normalized Spread Status compare.

    with tab2:
        st.markdown('<div class="section-label">Deal Drilldown</div>', unsafe_allow_html=True)
 
        deal_names = df_view["Deal Name"].dropna().unique().tolist()
        if not deal_names:
            st.info("No deals available for the current filters.")
            selected_deal = None
        else:
            _deal_sorted = sorted(deal_names)
            _prefill_deal = (
                st.session_state.get("drilldown_deal_pick")
                or (deal_pick if (deal_pick and deal_pick != "All deals") else None)
            )
            if _prefill_deal and _prefill_deal in _deal_sorted:
                st.session_state["deal_drilldown_select"] = _prefill_deal
            _deal_idx = (
                _deal_sorted.index(_prefill_deal)
                if (_prefill_deal and _prefill_deal in _deal_sorted)
                else 0
            )
            selected_deal = st.selectbox(
                "Select a deal",
                _deal_sorted,
                index=_deal_idx,
                key="deal_drilldown_select",
                help="Prefilled from sidebar Deal filter when available.",
            )
 
        if selected_deal:
            deal_rows = df_view[df_view["Deal Name"] == selected_deal]
 
            st.markdown(f"""
            <div class="info-box">
              <strong>{selected_deal}</strong> — {len(deal_rows)} effective date record(s) found
            </div>
            """, unsafe_allow_html=True)
 
            for _, row in deal_rows.iterrows():
                recon_bucket = _recon_status_bucket(row.get("recon_status", ""))
                border_color = (
                    "#4caf50"
                    if recon_bucket == "MATCH"
                    else ("#f44336" if recon_bucket == "MISMATCH" else "#ffc107")
                )
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
                      <div style='font-size:0.92rem; font-weight:600; color:#1a3a6c'>{_display_missing_dash(row.get("Facility"), "Facility")}</div>
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
                      <div style='font-size:0.65rem; color:#888; text-transform:uppercase; letter-spacing:.06em'>Spread ({col_tag})</div>
                      <div style='font-size:0.92rem; font-weight:600; color:#1a3a6c'>{pct_spread(sp_acp)}</div>
                    </div>
                    <div style='background:#f8fafd; border-radius:6px; padding:12px 14px;'>
                      <div style='font-size:0.65rem; color:#888; text-transform:uppercase; letter-spacing:.06em'>Spread (M61)</div>
                      <div style='font-size:0.92rem; font-weight:600; color:#1a3a6c'>{pct_spread(row.get("Spread (M61)"))}</div>
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
                  {_mismatch_detail_html(row)}
                </div>
                """, unsafe_allow_html=True)

                with st.expander(
                    f"Show Details — {row.get('Deal Name', 'Deal')} | {fmt_date(ed_acp)}",
                    expanded=False,
                ):
                    ctx = st.session_state.get("recon_context", {}) or {}
                    df_b_ctx = ctx.get("df_primary_matchable", pd.DataFrame())
                    df_a_ctx = ctx.get("df_m61_matchable", pd.DataFrame())

                    did = safe_str_strip(row.get("Deal ID Match Key (ACP)"))
                    mid = safe_str_strip(row.get("M61 Extracted Deal ID"))
                    dkey = did or mid
                    eff_acp_key = _date_key_ui(row.get("Effective Date (ACP)"))
                    eff_m61_key = _date_key_ui(row.get("Effective Date (M61)"))
                    note = safe_str_strip(row.get("Liability Note (M61)"))

                    st.caption(
                        "Underlying source rows: "
                        f"deal_id_key={dkey or '—'}, "
                        f"ACORE effective_date_key={eff_acp_key or '—'}, "
                        f"M61 effective_date_key (from this recon row)={eff_m61_key or '—'}, "
                        f"liability_note={note or '—'}"
                    )
                    st.caption(
                        "ACORE table: filtered by deal ID + **ACORE** effective date. "
                        "M61 table: **all** liability lines for this deal ID (each row keeps its own M61 Effective Date / "
                        "effective_date_key — not overwritten by ACORE)."
                    )
                    st.caption(
                        "Linked Match groups rows that share the same deal id + effective-date key on each side."
                    )

                    b_hit = pd.DataFrame()
                    a_hit = pd.DataFrame()
                    if isinstance(df_b_ctx, pd.DataFrame) and not df_b_ctx.empty:
                        b = df_b_ctx.copy()
                        mask_b = pd.Series(False, index=b.index)
                        if dkey and "acp_match_key" in b.columns and eff_acp_key and "effective_date_key" in b.columns:
                            mask_b = b["acp_match_key"].fillna("").astype(str).eq(dkey) & b[
                                "effective_date_key"
                            ].fillna("").astype(str).eq(eff_acp_key)
                        elif dkey and "acp_match_key" in b.columns:
                            mask_b = b["acp_match_key"].fillna("").astype(str).eq(dkey)
                        b_hit = b.loc[mask_b].copy()
                        if b_hit.empty and dkey and "acp_match_key" in b.columns:
                            b_hit = b.loc[b["acp_match_key"].fillna("").astype(str).eq(dkey)].copy()

                    if isinstance(df_a_ctx, pd.DataFrame) and not df_a_ctx.empty:
                        a = df_a_ctx.copy()
                        if dkey and "m61_match_key" in a.columns:
                            # Deal cohort only — do not require M61 effective_date_key to match ACORE (avoids wrong filter).
                            mask_a = a["m61_match_key"].fillna("").astype(str).eq(dkey)
                        elif note and "Liability Note" in a.columns:
                            mask_a = a["Liability Note"].fillna("").astype(str).str.strip().eq(note)
                        else:
                            mask_a = pd.Series(False, index=a.index)
                        a_hit = a.loc[mask_a].copy()
                        if a_hit.empty and dkey and "m61_match_key" in a.columns:
                            a_hit = a.loc[a["m61_match_key"].fillna("").astype(str).eq(dkey)].copy()

                    def _row_group_key(rr: pd.Series, side: str) -> str:
                        dk = ""
                        if side == "acore" and "acp_match_key" in rr.index:
                            dk = safe_str_strip(rr.get("acp_match_key"))
                        if side == "m61" and "m61_match_key" in rr.index:
                            dk = safe_str_strip(rr.get("m61_match_key"))
                        ek = safe_str_strip(rr.get("effective_date_key"))
                        if not ek:
                            ek = _date_key_ui(rr.get("Effective Date"))
                        nk = safe_str_strip(rr.get("Liability Note"))
                        if dk or ek:
                            return f"{dk}|{ek}"
                        if nk:
                            return f"note|{nk}"
                        return "unkeyed"

                    _all_group_keys = []
                    if not b_hit.empty:
                        _all_group_keys.extend(
                            b_hit.apply(lambda rr: _row_group_key(rr, "acore"), axis=1).tolist()
                        )
                    if not a_hit.empty:
                        _all_group_keys.extend(
                            a_hit.apply(lambda rr: _row_group_key(rr, "m61"), axis=1).tolist()
                        )
                    _group_rank = {}
                    for _g in _all_group_keys:
                        if _g not in _group_rank:
                            _group_rank[_g] = len(_group_rank) + 1

                    st.markdown("**Underlying ACORE rows**")
                    if b_hit.empty:
                        st.caption("No ACORE source rows found for this key.")
                    else:
                        b_cols = [
                            c
                            for c in [
                                "Deal ID",
                                "Deal Name",
                                "Facility",
                                "Source",
                                "Effective Date",
                                "Pledge Date",
                                "Maturity Date",
                                "Advance Rate",
                                "Spread",
                                "Current Undrawn Capacity",
                                "acp_match_key",
                                "effective_date_key",
                            ]
                            if c in b_hit.columns
                        ]
                        b_disp = b_hit.loc[:, b_cols].copy()
                        _bg = b_hit.apply(lambda rr: _row_group_key(rr, "acore"), axis=1)
                        b_disp.insert(
                            0,
                            "Linked Match",
                            _bg.map(lambda g: f"Match {_group_rank.get(g, 1)}"),
                        )
                        for dc in ("Effective Date", "Pledge Date", "Maturity Date"):
                            if dc in b_disp.columns:
                                b_disp[dc] = b_disp[dc].map(fmt_date)
                        for dc in [c for c in b_disp.columns if c.endswith("_date_key")]:
                            b_disp[dc] = b_disp[dc].map(fmt_date)
                        for pc in ("Advance Rate", "Spread", "Floor", "Recourse", "Recourse %"):
                            if pc in b_disp.columns:
                                b_disp[pc] = b_disp[pc].map(pct_spread if pc == "Spread" else pct)
                        st.dataframe(
                            b_disp.style.set_properties(
                                subset=["Linked Match"],
                                **{"background-color": "#eef4ff", "color": "#294a7a", "font-weight": "600"},
                            ),
                            use_container_width=True,
                            height=180,
                        )

                        st.markdown("**Underlying M61 rows**")
                        st.caption(
                            "Rows are matched by Deal ID and Effective Date. If the deal is the same "
                            "but the date is different, it will show as 'Same deal – different date'."
                        )
                        if a_hit.empty:
                            st.caption("No M61 source rows found for this key.")
                        else:
                            a_cols = [
                                c
                                for c in [
                                    "Fund Name",
                                    "Deal Name",
                                    "Liability Name",
                                    "Liability Type",
                                    "Financial Institution",
                                    "Status",
                                    "Pledge Date",
                                    "Effective Date",
                                    "Maturity Date",
                                    "Liability Note",
                                    "Target Advance Rate",
                                    "DealLevelAdvanceRate",
                                    "Deal Level Advance Rate",
                                    "Spread",
                                    "m61_match_key",
                                    "effective_date_key",
                                ]
                                if c in a_hit.columns
                            ]
                            a_disp = a_hit.loc[:, a_cols].copy()
                            # Visual-only drilldown aid: ACORE vs M61 effective date, plus rate/spread mismatch hints
                            # when dates align (does not affect reconciliation/matching).
                            _acore_eff_key = _date_key_ui(ed_acp)
                            _ar_mismatch = "MISMATCH" in safe_str_strip(
                                row.get("Advance Rate Status", "")
                            ).upper()
                            _sp_mismatch = "MISMATCH" in safe_str_strip(
                                row.get("Spread Status", "")
                            ).upper()

                            def _match_explain_for_m61_date(m61_k: str) -> str:
                                if _acore_eff_key and m61_k and m61_k == _acore_eff_key:
                                    parts: list[str] = []
                                    if _ar_mismatch:
                                        parts.append("advance rate differs vs ACORE")
                                    if _sp_mismatch:
                                        parts.append("spread differs vs ACORE")
                                    if parts:
                                        return "Same date – " + "; ".join(parts)
                                    return "Matches (same date)"
                                return "Same deal – different date"

                            if "Effective Date" in a_hit.columns:
                                _m61_eff_key = a_hit["Effective Date"].map(_date_key_ui)
                                _match_explanation = _m61_eff_key.map(_match_explain_for_m61_date)
                            else:
                                _match_explanation = pd.Series(
                                    [_match_explain_for_m61_date("") for _ in range(len(a_hit))],
                                    index=a_hit.index,
                                    dtype="object",
                                )
                            if "Effective Date" in a_disp.columns:
                                _eff_idx = a_disp.columns.get_loc("Effective Date") + 1
                            else:
                                _eff_idx = len(a_disp.columns)
                            a_disp.insert(
                                _eff_idx,
                                "Match Explanation",
                                _match_explanation.reindex(a_disp.index),
                            )
                            # Display-only mirror of reconciliation basis:
                            # follow the already-computed source on the current reconciliation row.
                            _adv_src = safe_str_strip(row.get("Advance Rate Source (M61)")).lower()
                            _use_deal_level = _adv_src == "deal level advance rate"
                            _tgt_adv = (
                                a_hit["Target Advance Rate"]
                                if "Target Advance Rate" in a_hit.columns
                                else pd.Series(pd.NA, index=a_hit.index)
                            )
                            if "DealLevelAdvanceRate" in a_hit.columns:
                                _dl_adv = a_hit["DealLevelAdvanceRate"]
                            elif "Deal Level Advance Rate" in a_hit.columns:
                                _dl_adv = a_hit["Deal Level Advance Rate"]
                            else:
                                _dl_adv = pd.Series(pd.NA, index=a_hit.index)
                            _eff_adv = _dl_adv.copy() if _use_deal_level else _tgt_adv.copy()
                            a_disp.insert(
                                len(a_disp.columns),
                                "Advance Rate Used (M61)",
                                _eff_adv.reindex(a_disp.index),
                            )
                            _ag = a_hit.apply(lambda rr: _row_group_key(rr, "m61"), axis=1)
                            a_disp.insert(
                                0,
                                "Linked Match",
                                _ag.map(lambda g: f"Match {_group_rank.get(g, 1)}"),
                            )
                            for dc in ("Effective Date", "Pledge Date", "Maturity Date"):
                                if dc in a_disp.columns:
                                    a_disp[dc] = a_disp[dc].map(fmt_date)
                            for dc in [c for c in a_disp.columns if c.endswith("_date_key")]:
                                a_disp[dc] = a_disp[dc].map(fmt_date)
                            for pc in (
                                "Target Advance Rate",
                                "Current Advance Rate",
                                "DealLevelAdvanceRate",
                                "Deal Level Advance Rate",
                                "Advance Rate Used (M61)",
                                "Spread",
                                "IndexFloor",
                                "Floor",
                                "Recourse",
                                "Recourse %",
                            ):
                                if pc in a_disp.columns:
                                    a_disp[pc] = a_disp[pc].map(pct_spread if pc == "Spread" else pct)
                            st.dataframe(
                                a_disp.style.set_properties(
                                    subset=["Linked Match"],
                                    **{"background-color": "#eef4ff", "color": "#294a7a", "font-weight": "600"},
                                ),
                                use_container_width=True,
                                height=200,
                            )

    st.markdown("<br>", unsafe_allow_html=True)

    # ---- Download ----
    df_export_ready = df_export_ui.copy()
    # Export parity with displayed table: table view falls back Spread (ACORE) to raw primary "Spread".
    if "Spread (ACP)" in df_export_ready.columns and "Spread" in df_export_ready.columns:
        _sp_acp_blank = df_export_ready["Spread (ACP)"].isna()
        if _sp_acp_blank.any():
            df_export_ready.loc[_sp_acp_blank, "Spread (ACP)"] = df_export_ready.loc[_sp_acp_blank, "Spread"]
    df_export_ready = _normalize_display_missing_df(df_export_ready)
    # "Same as Above": mirror the display transformation in the export output.
    df_export_ready = _apply_same_as_above_export(df_export_ready)

    col_dl1, col_dl2 = st.columns(2)

    with col_dl1:
        _excel_payload = b""
        if not is_stale_selection:
            _excel_payload = to_excel_bytes(df_export_ready, run_primary)
        st.download_button(
            label="⬇️ Download Excel",
            data=_excel_payload,
            file_name=run_excel_name,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            disabled=is_stale_selection,
        )

    with col_dl2:
        df_csv_export = df_export_ready.drop(
            columns=["Target Advance Rate (M61)"], errors="ignore"
        ).copy()
        if "recon_status" in df_csv_export.columns and "Overall Recon Status" not in df_csv_export.columns:
            df_csv_export = df_csv_export.rename(columns={"recon_status": "Overall Recon Status"})

        csv_status_cols = [
            "File Source",
            "Effective Date Status",
            "Pledge Date Status",
            "Advance Rate Status",
            "Spread Status",
            "Undrawn Capacity Status",
            "Index Floor Status",
            "Index Name Status",
            "Recourse % Status",
            "Overall Recon Status",
        ]
        for col in csv_status_cols:
            if col not in df_csv_export.columns:
                df_csv_export[col] = ""

        # Keep existing column order, but ensure status columns appear in the canonical sequence.
        base_cols = [c for c in df_csv_export.columns if c not in csv_status_cols]
        df_csv_export = df_csv_export.loc[:, base_cols + csv_status_cols]

        for c in df_csv_export.columns:
            cl = str(c).lower()
            if "date" in cl:
                df_csv_export[c] = df_csv_export[c].map(
                    lambda v: "-" if pd.isna(v) else fmt_date(v)
                )
            elif _is_spread_value_column(c):
                df_csv_export[c] = df_csv_export[c].map(
                    lambda v: "-" if pd.isna(v) else pct_spread(v)
                )
            elif any(tok in cl for tok in ("rate", "recourse", "index floor", " floor")):
                df_csv_export[c] = df_csv_export[c].map(
                    lambda v: "-" if pd.isna(v) else pct(v)
                )
        df_csv_export = _normalize_display_missing_df(df_csv_export)
        csv_data = df_csv_export.to_csv(index=False).encode("utf-8")

        st.download_button(
            label="⬇️ Download CSV",
            data=csv_data,
            file_name=run_csv_name,
            mime="text/csv",
            disabled=is_stale_selection,
        )
    st.caption(
        "Excel and CSV match the **Results** table: same rows and order after **Scope**, **M61 Note Category**, "
        "status and deal filters, **table filters**, and **sort**."
    )
    if is_stale_selection:
        st.caption("Downloads are disabled until you rerun with the current selection.")

else:
    render_original_landing_page_if_no_results(selected_ui_label)