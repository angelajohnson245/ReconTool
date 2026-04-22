from __future__ import annotations

import argparse
import os
import re
import sys
from datetime import date, datetime

import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter


# --------------------------------------------------
# 1. FILE PATHS
# --------------------------------------------------
# ACP II: ACP II - Liquidity & Earnings Model - 2026-03-20 - v1.xlsm
# M61 ^: (In)  Liability_Relationship_4132026_1776102909429.xlsx

# ACP III: ACP III - Liquidity & Earnings Model - 2026-03-20 - v1.xlsm
# M61^: (In)  Liability_Relationship_4102026_1775858455341.xlsx

BASE_DIR = os.path.dirname(os.path.abspath(__file__))


# m61
DEFAULT_FILE_A_PATH = os.path.join(
    BASE_DIR, "(In)  Liability_Relationship_4102026_1775858455341.xlsx"
)

DEFAULT_FILE_B_PATH = os.path.join(
    BASE_DIR, "ACP II - Liquidity & Earnings Model - 2026-03-20 - v1.xlsm"
)

# Fund behavior by selected primary type:
# - export_label: file naming label
# - scope_label: short label shown in scoped UI messages
# - fund_token: case-insensitive token to scope rows by Fund column
PRIMARY_TYPE_FUND_CONFIG = {
    "ACORE": {
        "export_label": "ACORE - ACP III",
        "scope_label": "ACP III",
        "fund_token": "credit partners iii",
        "fund_regex": r"\bcredit partners iii\b",
        # M61-style Fund column when Liability export has no Fund Name for this row.
        "recon_fund_display": "ACORE Credit Partners III",
    },
    "ACP II": {
        "export_label": "ACORE - ACP II",
        "scope_label": "ACP II",
        "fund_token": "credit partners ii",
        "fund_regex": r"\bcredit partners ii\b",
        "recon_fund_display": "ACORE Credit Partners II",
    },
    "AOC II": {
        "export_label": "ACORE - AOC II",
        "scope_label": "AOC II",
        # Scoped rows: Fund must spell out Roman II (do not match Opportunistic Credit I).
        "fund_token": "opportunistic credit ii",
        "fund_regex": r"\bopportunistic\s+credit\s+ii\b",
        "recon_fund_display": "Opportunistic Credit II",
    },
}


def _fund_shorthand_to_canonical_map() -> dict[str, str]:
    """Shorthand / internal labels -> M61-style Fund column (single source of truth with PRIMARY_TYPE_FUND_CONFIG)."""
    out: dict[str, str] = {}
    for _ptype, cfg in PRIMARY_TYPE_FUND_CONFIG.items():
        canon = cfg.get("recon_fund_display")
        if not canon:
            continue
        for key in (cfg.get("export_label"), cfg.get("scope_label")):
            if key and str(key).strip():
                out[str(key).strip()] = str(canon).strip()
    return out


# Used by ``normalize_recon_fund_for_output`` (case-insensitive match on keys).
FUND_SHORTHAND_TO_CANONICAL = _fund_shorthand_to_canonical_map()


def normalize_recon_fund_for_output(df: pd.DataFrame) -> pd.DataFrame:
    """
    Replace shorthand Fund labels (e.g. ``ACORE - ACP III``) with canonical M61-style names
    for display and Excel export. Does not alter ``reconcile()``; call on the recon dataframe
    after reconciliation.
    """
    if df is None or "Fund" not in getattr(df, "columns", ()):
        return df
    if df.empty:
        return df
    lc = {k.lower(): v for k, v in FUND_SHORTHAND_TO_CANONICAL.items()}

    def _norm_one(v):
        if v is None:
            return v
        try:
            if pd.isna(v):
                return v
        except (TypeError, ValueError):
            return v
        s = str(v).strip()
        if not s or s.lower() in ("nan", "<na>", "nat", "none"):
            return v
        return lc.get(s.lower(), v)

    out = df.copy()
    out["Fund"] = out["Fund"].map(_norm_one)
    return out


def _fund_cfg(primary_file_type: str) -> dict:
    return PRIMARY_TYPE_FUND_CONFIG.get(primary_file_type, {})


def _debug_rows(msg: str) -> None:
    """Temporary row-count diagnostics for reconciliation pipeline."""
    print(f"[RECON DEBUG] {msg}")


def _debug_m61_load_preview(df: pd.DataFrame, n: int = 5) -> None:
    """TEMP: first n M61 rows after load_file_a (remove after verification)."""
    want = ("Deal Name", "Liability Note", "Effective Date", "Liability Type")
    cols = [c for c in want if c in df.columns]
    if not cols:
        _debug_rows("TEMP DEBUG: M61 load preview skipped (expected columns missing)")
        return
    _debug_rows(f"TEMP DEBUG: M61 load preview (head {n}) cols={cols}")
    snippet = df[cols].head(n).to_string(index=True, max_colwidth=60)
    for ln in snippet.splitlines():
        _debug_rows(f"TEMP DEBUG:   {ln}")


def _debug_match_key_sample_rows(
    label: str,
    df: pd.DataFrame,
    *,
    deal_col: str,
    facility_col: str,
    note_col: str,
    eff_col: str,
    n: int = 10,
) -> None:
    """TEMP: print raw key components + match_key for head(n) rows (remove after diagnosis)."""
    need = (deal_col, facility_col, note_col, eff_col, "match_key")
    missing = [c for c in need if c not in df.columns]
    _debug_rows(
        f"TEMP DEBUG: match_key components — {label} (head {n}); "
        f"source cols Deal={deal_col!r} Facility={facility_col!r} "
        f"Note={note_col!r} Date={eff_col!r}"
    )
    if missing:
        _debug_rows(f"TEMP DEBUG:   skip sample (missing columns: {missing})")
        return
    sub = df.loc[:, list(need)].head(n)
    for i, (_, row) in enumerate(sub.iterrows(), start=1):
        d, fac, nte, eff, mk = (
            row[deal_col],
            row[facility_col],
            row[note_col],
            row[eff_col],
            row["match_key"],
        )
        mk_s = "" if pd.isna(mk) else str(mk)
        if len(mk_s) > 220:
            mk_s = mk_s[:220] + "..."
        _debug_rows(
            f"TEMP DEBUG:   #{i} deal={d!r} facility={fac!r} note={nte!r} "
            f"eff={eff!r} match_key={mk_s!r}"
        )


def _debug_match_key_overlap_diagnosis(df_b: pd.DataFrame, df_a: pd.DataFrame, *, n: int = 10) -> None:
    """TEMP: for head(n) primary rows, show whether any M61 row shares deal+eff date and if facility_norm aligns."""
    req = {"deal_norm", "facility_norm", "effective_date_key", "match_key"}
    if not req.issubset(df_b.columns) or not req.issubset(df_a.columns):
        _debug_rows("TEMP DEBUG: match_key overlap diagnosis skipped (norm columns missing)")
        return
    _debug_rows(
        "TEMP DEBUG: match_key diagnosis — join key is deal + facility + eff date only; "
        "note is shown above for context but is NOT in match_key. "
        "Below: same deal_norm+date_key as primary? If yes, compare facility_norm lists."
    )
    for _, r in df_b.head(n).iterrows():
        d = r["deal_norm"]
        ed = r["effective_date_key"]
        fk = r["facility_norm"]
        mk = r["match_key"]
        cand = df_a[(df_a["deal_norm"] == d) & (df_a["effective_date_key"] == ed)]
        if cand.empty:
            _debug_rows(
                "TEMP DEBUG:   "
                f"no M61 row with same deal_norm+date_key | "
                f"deal_norm={d!r} date_key={ed!r} primary_facility_norm={fk!r} match_key={mk!r}"
            )
        else:
            m61_facs = sorted({str(x) for x in cand["facility_norm"].tolist()})
            hit = mk in set(cand["match_key"].tolist())
            _debug_rows(
                "TEMP DEBUG:   "
                f"deal_norm={d!r} date_key={ed!r} | primary_facility_norm={fk!r} | "
                f"m61_facility_norm_candidates={m61_facs} | full_key_match={hit}"
            )


def _debug_unmatched_after_merge(merged: pd.DataFrame, *, label_a: str, n: int = 10) -> None:
    """TEMP: isolate rows that survive filtering but still do not align on match_key."""
    if merged.empty or "_merge" not in merged.columns:
        return
    left = merged[merged["_merge"] == "left_only"].copy()
    right = merged[merged["_merge"] == "right_only"].copy()
    _debug_rows(
        "TEMP DEBUG: unmatched survivors after outer merge — "
        f"left_only_primary={len(left)} right_only_m61={len(right)}"
    )
    left_deals = left["Deal Name"].dropna().astype(str).str.strip()
    right_deals = right[f"{label_a}_Deal Name"].dropna().astype(str).str.strip()
    overlap_deals = sorted(set(left_deals).intersection(set(right_deals)))
    _debug_rows(
        "TEMP DEBUG: unmatched deal-name overlap (indicates key-component mismatch, not filter loss) — "
        f"overlap_deals_count={len(overlap_deals)} sample={overlap_deals[:n]}"
    )
    if not left.empty:
        _debug_rows("TEMP DEBUG: sample unmatched PRIMARY rows (head 10)")
        show_cols_l = [c for c in ("Deal Name", "Facility", "Note Name", "Effective Date", "match_key") if c in left.columns]
        for i, (_, r) in enumerate(left[show_cols_l].head(n).iterrows(), start=1):
            _debug_rows(
                "TEMP DEBUG:   "
                f"#{i} deal={r.get('Deal Name')!r} facility={r.get('Facility')!r} "
                f"note={r.get('Note Name')!r} eff={r.get('Effective Date')!r} key={r.get('match_key')!r}"
            )
    if not right.empty:
        _debug_rows("TEMP DEBUG: sample unmatched M61 rows (head 10)")
        show_cols_r = [
            c
            for c in (
                f"{label_a}_Deal Name",
                f"{label_a}_Liability Name",
                f"{label_a}_Liability Note",
                f"{label_a}_Effective Date",
                f"{label_a}_match_key",
            )
            if c in right.columns
        ]
        for i, (_, r) in enumerate(right[show_cols_r].head(n).iterrows(), start=1):
            _debug_rows(
                "TEMP DEBUG:   "
                f"#{i} deal={r.get(f'{label_a}_Deal Name')!r} facility={r.get(f'{label_a}_Liability Name')!r} "
                f"note={r.get(f'{label_a}_Liability Note')!r} eff={r.get(f'{label_a}_Effective Date')!r} "
                f"key={r.get(f'{label_a}_match_key')!r}"
            )


def detect_fund_label(uploaded_filename: str | None, primary_file_type: str) -> str:
    """Infer fund label from uploaded filename, else use selected primary fund label."""
    if uploaded_filename:
        name = uploaded_filename.upper()
        if re.search(r"\bACP\s+III\b", name):
            return PRIMARY_TYPE_FUND_CONFIG["ACORE"]["export_label"]
        if re.search(r"\bACP\s+II\b", name):
            return PRIMARY_TYPE_FUND_CONFIG["ACP II"]["export_label"]
        if re.search(r"\bAOC\s+II\b", name):
            return PRIMARY_TYPE_FUND_CONFIG["AOC II"]["export_label"]

    cfg = _fund_cfg(primary_file_type)
    return cfg.get("export_label", primary_file_type)


def scope_label_for_primary_type(primary_file_type: str) -> str:
    cfg = _fund_cfg(primary_file_type)
    return cfg.get("scope_label", primary_file_type)


def filter_recon_to_selected_fund(df_recon: pd.DataFrame, primary_file_type: str) -> pd.DataFrame:
    """Filter reconciliation rows by selected fund scope (display/export helper).

    Returns a row subset only; does not change any column values (including ``Fund``).
    """
    cfg = _fund_cfg(primary_file_type)
    token = cfg.get("fund_token")
    if not token:
        return df_recon.copy()
    if "Fund" not in df_recon.columns:
        return df_recon.copy()
    pattern = cfg.get("fund_regex") or re.escape(token)
    fund_series = df_recon["Fund"].fillna("").astype(str)
    mask_m61 = fund_series.str.contains(pattern, case=False, regex=True, na=False)
    # Rows tagged with the business-file fund label (ACORE-only / primary-side Fund fill).
    mask_primary = pd.Series(False, index=df_recon.index)
    for lab in (cfg.get("export_label"), cfg.get("scope_label"), cfg.get("recon_fund_display")):
        if not lab:
            continue
        mask_primary |= fund_series.str.contains(
            re.escape(str(lab).strip()), case=False, regex=True, na=False
        )
    out = df_recon[(mask_m61 | mask_primary)].copy()
    _debug_rows(
        f"Scoped filter ({primary_file_type}) rows: in={len(df_recon)} out={len(out)} "
        f"mask_m61={int(mask_m61.sum())} mask_primary={int(mask_primary.sum())}"
    )
    return out


def build_output_filename(
    primary_file_type: str,
    ext: str,
    uploaded_filename: str | None = None,
) -> str:
    """Download/save name, e.g. 'ACP II Finance Recon - 2026-04-13.xlsx'."""
    date_str = datetime.today().strftime("%Y-%m-%d")
    e = ext[1:] if ext.startswith(".") else ext
    fund_label = detect_fund_label(uploaded_filename, primary_file_type)
    return f"{fund_label} Finance Recon - {date_str}.{e}"


def default_recon_output_path(
    primary_file_type: str = "ACORE",
    uploaded_filename: str | None = None,
) -> str:
    return os.path.join(
        BASE_DIR,
        build_output_filename(primary_file_type, "xlsx", uploaded_filename),
    )
# --------------------------------------------------
# 2. CONSTANTS
# --------------------------------------------------
M61_FINANCING_TYPES = ["Repo", "Non", "Subline"]

TARGET_FUNDS = {
    "acore credit partners iii",
    "acore credit partners ii",
    "aoc ii",
    "mcp",
    "api",
    "acore",
}

# M61 "Fund Name" for AOC: short "aoc ii" or full "… Opportunistic Credit II" (Roman II required;
# do not match Opportunistic Credit I). ACP rows stay on exact TARGET_FUNDS membership only.
AOC_M61_FUND_NAME_RE = re.compile(
    r"\b(?:aoc\s+ii|opportunistic\s+credit\s+ii)\b",
    re.IGNORECASE,
)

FLOAT_TOLERANCE = 1e-6
ENABLE_DEAL_ID_SUFFIX_FALLBACK = False

ACP_SHEET_COLS = [
    "Deal Name",
    "Note Name",
    "Source",
    "Facility",
    "Advance Rate",
    "Spread",
    "Pledge",
    "Pledge Date",
    "Effective Date",
    "Current Undrawn Capacity",
    "Maturity Date",
    "Floor",
    "Recourse %",
]

# Pledge / Pledge Date handled separately (explicit statuses); not in this list.
COMPARE_FIELDS = [
    ("Advance Rate", "Current Advance Rate", "numeric"),
    ("Spread", "Spread", "numeric"),
    ("Effective Date", "Effective Date", "date"),
    ("Current Undrawn Capacity", "Undrawn Capacity", "numeric"),
]

RECON_STATUS_FIELDS = frozenset({"Advance Rate", "Spread", "Effective Date"})

LIABILITY_ADVANCE_RATE_COLUMNS = ("Current Advance Rate", "Advance Rate", "Advance")

LIABILITY_VALUE_LABELS = {
    "Current Advance Rate": "Advance Rate (M61)",
    "Spread": "Spread (M61)",
    "Pledge": "Pledge (M61)",
    "Pledge Date": "Pledge Date (M61)",
    "Effective Date": "Effective Date (M61)",
    "Undrawn Capacity": "Current Undrawn Capacity (M61)",
    "Current Balance": "Current Balance (M61)",
    "Liability Note": "Liability Note (M61)",
    "Financial Institution": "Financial Institution (M61)",
    "Maturity Date": "Maturity Date (M61)",
    "Fund Name": "Fund (M61)",
    "Liability Name": "Liability Name (M61)",
    "target": "Target (M61)",
    "Status": "Status (M61)",
}

# Extra Liability_Relationship columns to load (NA if absent).
LIABILITY_SOURCE_EXTRA_COLS = [
    "IndexFloor",
    "IndexName",
    "Recourse %",
    "Recourse",
    "RecoursePct",
]

RECON_ORDERED_COLS = [
    "Fund",
    "Deal Name",
    "Facility",
    "Financial Line",
    "Note Name",
    "Liability Note (M61)",
    "Deal ID (ACP)",
    "Liability Note Suffix (M61)",
    "Source",
    "Source Indicator",
    "Effective Date (ACP)",
    "Effective Date (M61)",
    "Pledge Date (ACP)",
    "Pledge Date (M61)",
    "Advance Rate (ACP)",
    "Advance Rate (M61)",
    "Target Advance Rate (M61)",
    "Spread (ACP)",
    "Spread (M61)",
    "Undrawn Capacity (ACP)",
    "Undrawn Capacity (M61)",
    "Index Floor (ACP)",
    "Index Floor (M61)",
    "Index Name (ACP)",
    "Index Name (M61)",
    "Recourse % (ACP)",
    "Recourse % (M61)",
    "Advance Rate Status",
    "Spread Status",
    "Effective Date Status",
    "Undrawn Capacity Status",
    "Index Floor Status",
    "Index Name Status",
    "Recourse % Status",
    "Pledge Date Status",
    "recon_status",
]

FACILITY_NORM_MAP = {
    "jpm repo": "jpm",
    "gs repo": "gs",
    "ms repo": "ms",
    "boa repo": "boa",
    "tbd": "tbd",
    "acpiii-jpm-repo": "jpm",
    "acpiii-gs-repo": "gs",
    "acpiii-ms-repo": "ms",
    "acpiii-boa-repo": "boa",
    "acpii-jpm-repo": "jpm",
    "acpii-gs-repo": "gs",
    "acpii-ms-repo": "ms",
    "acpii-boa-repo": "boa",
}

PRIMARY_INTERNAL_FIELDS = (
    "deal_name",
    "note_name",
    "source",
    "facility",
    "advance_rate",
    "spread",
    "pledge",
    "pledge_date",
    "effective_date",
    "undrawn_capacity",
    "maturity_date",
    "floor",
    "recourse_pct",
)

INTERNAL_FIELD_TO_OUTPUT_COL = {
    "deal_name": "Deal Name",
    "note_name": "Note Name",
    "source": "Source",
    "facility": "Facility",
    "advance_rate": "Advance Rate",
    "spread": "Spread",
    "pledge": "Pledge",
    "pledge_date": "Pledge Date",
    "effective_date": "Effective Date",
    "undrawn_capacity": "Current Undrawn Capacity",
    "maturity_date": "Maturity Date",
    "floor": "Floor",
    "recourse_pct": "Recourse %",
}


class PrimaryFileSchemaError(ValueError):
    def __init__(self, primary_type: str, missing: list[str]):
        self.primary_type = primary_type
        self.missing = missing
        detail = "; ".join(missing) if missing else "unknown"
        super().__init__(f"[{primary_type}] Missing required column(s): {detail}")


PRIMARY_FILE_CONFIG: dict[str, dict] = {
    "ACORE": {
        "sheet_name": "10) Fin Inpt",
        "header_row": 6,
        "column_map": {
            "deal_name": "Deal Name",
            "note_name": "Note Name",
            "source": "Source",
            "facility": "Facility",
            "advance_rate": "Advance Rate",
            "spread": "Spread",
            "pledge": "Pledge",
            "pledge_date": "Pledge Date",
            "effective_date": "Effective Date",
            "undrawn_capacity": "Current Undrawn Capacity",
            "maturity_date": "Maturity Date",
            "floor": "Floor",
            "recourse_pct": "Recourse %",
        },
        "display_name": "ACORE",
        "ui_display_label": "ACORE - ACP III",
        "model_descriptor": "ACORE Liquidity & Earnings Model",
        "source_indicator_primary_only": "ACORE",
        "missing_in_primary_label": "ACORE",
        "excel_primary_column_suffix": "ACORE",
        "primary_only_legend_label": "ACORE Only",
        "primary_group_header": "ACORE — PRIMARY DATA",
    },
    "AOC II": {
        # Same tab/layout family as ACP II / ACORE; advance is typically labeled "Advance"
        # (load_primary_file synthesizes "Advance Rate" when needed).
        "sheet_name": "10) Fin Inpt",
        "header_row": 6,
        "sanitize_fin_inpt_headers": True,
        "column_map": {
            "deal_name": "Deal Name",
            "note_name": "Note Name",
            "source": "Source",
            "facility": "Facility",
            "advance_rate": "Advance",
            "spread": "Spread",
            "pledge": "Pledge",
            "pledge_date": "Pledge Date",
            "effective_date": "Effective Date",
            "undrawn_capacity": "Current Undrawn Capacity",
            "maturity_date": "Maturity Date",
            "floor": "Floor",
            "recourse_pct": "Recourse %",
        },
        "display_name": "AOC II",
        "ui_display_label": "ACORE - AOC II",
        "model_descriptor": "AOC II Liquidity & Earnings Model",
        "source_indicator_primary_only": "AOC II",
        "missing_in_primary_label": "AOC II",
        "excel_primary_column_suffix": "AOC II",
        "primary_only_legend_label": "AOC II Only",
        "primary_group_header": "AOC II — PRIMARY DATA",
    },
    "ACP II": {
        # Layout verified on "ACP II - Liquidity & Earnings Model - 2026-03-20 - v1.xlsm".
        # Data tab: "10) Fin Inpt" (same name as ACORE/ACP III). pandas header_row=6 = Excel row 7.
        # Present on that sheet: Deal Name, Note Name, Effective Date, Pledge Date, Advance,
        # Source, Facility, Spread, Floor, Recourse % (plus fees / IDs). Advance → Advance Rate
        # is handled in load_primary_file. Not on this sheet: Pledge, Current Undrawn Capacity,
        # Maturity Date — those are added as NA via ensure_columns (recon still runs).
        "sheet_name": "10) Fin Inpt",
        "header_row": 6,
        "sanitize_fin_inpt_headers": True,
        "column_map": {
            "deal_name": "Deal Name",
            "note_name": "Note Name",
            "source": "Source",
            "facility": "Facility",
            "advance_rate": "Advance",
            "spread": "Spread",
            "pledge": "Pledge",
            "pledge_date": "Pledge Date",
            "effective_date": "Effective Date",
            "undrawn_capacity": "Current Undrawn Capacity",
            "maturity_date": "Maturity Date",
            "floor": "Floor",
            "recourse_pct": "Recourse %",
        },
        "display_name": "ACP II",
        "ui_display_label": "ACORE - ACP II",
        "model_descriptor": "ACP II Liquidity & Earnings Model",
        "source_indicator_primary_only": "ACP II",
        "missing_in_primary_label": "ACP II",
        "excel_primary_column_suffix": "ACP II",
        "primary_only_legend_label": "ACP II Only",
        "primary_group_header": "ACP II — PRIMARY DATA",
    },
}

STREAMLIT_PRIMARY_FILE_TYPES = ("ACORE", "AOC II", "ACP II")

PRIMARY_REQUIRED_IN_SHEET = ("deal_name", "note_name", "facility", "effective_date")


def get_primary_config(primary_file_type: str) -> dict:
    if primary_file_type not in PRIMARY_FILE_CONFIG:
        known = ", ".join(sorted(PRIMARY_FILE_CONFIG))
        raise ValueError(
            f"Unknown primary file type {primary_file_type!r}. Expected one of: {known}"
        )
    return PRIMARY_FILE_CONFIG[primary_file_type]


def ensure_columns(df: pd.DataFrame, columns) -> pd.DataFrame:
    for col in columns:
        if col not in df.columns:
            df[col] = pd.NA
    return df


# ACP II "10) Fin Inpt" can pick up merged-area junk columns (float headers, "0.0025.1", etc.).
_FIN_INPT_NUMERIC_LIKE_HEADER_RE = re.compile(r"^[\d.]+$")


def _sanitize_fin_inpt_raw_df(df: pd.DataFrame, cfg: dict) -> pd.DataFrame:
    if not cfg.get("sanitize_fin_inpt_headers"):
        return df
    keep = []
    for c in df.columns:
        if not isinstance(c, str):
            continue
        s = str(c).strip()
        if _FIN_INPT_NUMERIC_LIKE_HEADER_RE.fullmatch(s):
            continue
        keep.append(c)
    return df.loc[:, keep].copy()


def inspect_primary_workbook(path: str, primary_file_type: str) -> None:
    """
    Print the configured sheet name, header row, and raw column headers as read by pandas.
    Use: python recon_enhanced_output.py --inspect-primary --file-b <workbook.xlsm> --primary-type ACP II
    """
    cfg = get_primary_config(primary_file_type)
    sheet = cfg["sheet_name"]
    hdr = cfg["header_row"]
    xl = pd.ExcelFile(path)
    print(f"primary_file_type={primary_file_type!r}")
    print(f"workbook={path!r}")
    print(f"all_sheet_names ({len(xl.sheet_names)}):")
    for i, name in enumerate(xl.sheet_names):
        mark = " <-- configured" if name == sheet else ""
        print(f"  [{i:3d}] {name!r}{mark}")
    if sheet not in xl.sheet_names:
        print(f"ERROR: configured sheet_name {sheet!r} not found.", file=sys.stderr)
        return
    df0 = pd.read_excel(path, sheet_name=sheet, header=hdr, nrows=0)
    print(f"\nUsing sheet_name={sheet!r}  header_row={hdr} (pandas read_excel header=)")
    print(f"raw column count: {len(df0.columns)}")
    for i, c in enumerate(df0.columns):
        print(f"  [{i:2d}] {c!r}  ({type(c).__name__})")
    if cfg.get("sanitize_fin_inpt_headers"):
        df1 = _sanitize_fin_inpt_raw_df(df0, cfg)
        print(f"\nAfter sanitize_fin_inpt_headers: {len(df1.columns)} columns")
        for i, c in enumerate(df1.columns):
            print(f"  [{i:2d}] {c!r}")


def _missing_primary_columns(df: pd.DataFrame, cfg: dict) -> list[str]:
    missing = []
    cmap = cfg["column_map"]
    cols = set(df.columns.astype(str))
    for key in PRIMARY_REQUIRED_IN_SHEET:
        src = cmap.get(key)
        if not src:
            missing.append(f"{key}: no mapping configured")
        elif src not in cols:
            missing.append(f"{key}: expected column {src!r} (not found in sheet)")
    return missing


def safe_parse_date(series):
    s = pd.Series(series)
    out = pd.to_datetime(s, errors="coerce", dayfirst=False).dt.normalize()

    missing = out.isna() & s.notna()
    if missing.any():
        sub = s[missing].astype(str).str.strip()
        for fmt in ("%m/%d/%y", "%m-%d-%y", "%m/%d/%Y", "%m-%d-%Y"):
            retry = pd.to_datetime(sub, format=fmt, errors="coerce").dt.normalize()
            take = retry.notna()
            if take.any():
                out.loc[sub.index[take]] = retry.loc[sub.index[take]]
            missing = out.isna() & s.notna()
            if not missing.any():
                break

    missing = out.isna() & s.notna()
    if missing.any():
        num = pd.to_numeric(s[missing], errors="coerce")
        ok = num.notna() & (num > 20000) & (num < 1_000_000)
        if ok.any():
            ex = pd.to_datetime(num[ok], unit="D", origin="1899-12-30", errors="coerce").dt.normalize()
            out.loc[ex.index] = ex

    return out


def load_primary_file(path: str, primary_file_type: str) -> pd.DataFrame:
    cfg = get_primary_config(primary_file_type)
    df_raw = pd.read_excel(path, sheet_name=cfg["sheet_name"], header=cfg["header_row"])
    df_raw = _sanitize_fin_inpt_raw_df(df_raw, cfg)

    if "Advance" in df_raw.columns and "Advance Rate" not in df_raw.columns:
        df_raw["Advance Rate"] = df_raw["Advance"]
    if "Undrawn Capacity" in df_raw.columns and "Current Undrawn Capacity" not in df_raw.columns:
        df_raw["Current Undrawn Capacity"] = df_raw["Undrawn Capacity"]

    miss = _missing_primary_columns(df_raw, cfg)
    if miss:
        raise PrimaryFileSchemaError(primary_file_type, miss)

    cmap = cfg["column_map"]
    all_src = list(dict.fromkeys(cmap[k] for k in PRIMARY_INTERNAL_FIELDS))
    df_raw = ensure_columns(df_raw, all_src)

    use_cols = [cmap[k] for k in PRIMARY_INTERNAL_FIELDS]
    renamer = {cmap[k]: INTERNAL_FIELD_TO_OUTPUT_COL[k] for k in PRIMARY_INTERNAL_FIELDS}
    df = df_raw.loc[:, use_cols].rename(columns=renamer)
    if "Deal ID" in df_raw.columns:
        df["Deal ID"] = df_raw["Deal ID"]

    df = ensure_columns(df, ACP_SHEET_COLS)
    keep_cols = list(ACP_SHEET_COLS) + (["Deal ID"] if "Deal ID" in df.columns else [])
    df = df[keep_cols].copy().dropna(subset=["Deal Name"])

    for col in ["Deal Name", "Note Name", "Source", "Facility", "Pledge"]:
        df[col] = df[col].astype(str).str.strip()
    if "Deal ID" in df.columns:
        df["Deal ID"] = df["Deal ID"].apply(
            lambda v: pd.NA if pd.isna(v) or not str(v).strip() else str(v).strip()
        )

    df["Pledge Date"] = safe_parse_date(df["Pledge Date"])
    df["Effective Date"] = safe_parse_date(df["Effective Date"])
    df["Maturity Date"] = safe_parse_date(df["Maturity Date"])
    return df.reset_index(drop=True)


def load_file_a(path: str, *, return_excluded: bool = False):
    sheet_name = "Liability_Relationship"
    df = pd.read_excel(path, sheet_name=sheet_name)
    _debug_rows(
        f"M61 read_excel: sheet={sheet_name!r} rows={len(df)} cols={len(df.columns)} "
        f"(full sheet read; no usecols/nrows subset)"
    )
    keep = [
        "Fund Name",
        "Liability Name",
        "Liability Type",
        "Financial Institution",
        "Deal Name",
        "Liability Note",
        "Status",
        "Pledge",
        "Pledge Date",
        "Effective Date",
        "Maturity Date",
        "Current Advance Rate",
        "Target Advance Rate",
        "Current Balance",
        "Undrawn Capacity",
        "Spread",
        "target",
        "in_liability",
    ]
    keep = list(dict.fromkeys(keep + list(LIABILITY_ADVANCE_RATE_COLUMNS) + LIABILITY_SOURCE_EXTRA_COLS))
    df = ensure_columns(df, keep)
    _debug_rows(f"M61 after ensure_columns(keep): rows={len(df)}")

    df["Fund Name"] = df["Fund Name"].astype(str).str.strip()
    fund_lower = df["Fund Name"].str.lower()
    in_target_set = fund_lower.isin(TARGET_FUNDS)
    aoc_style = df["Fund Name"].str.contains(AOC_M61_FUND_NAME_RE, regex=True, na=False)
    fund_mask = in_target_set | aoc_style
    liab_type_mask = df["Liability Type"].isin(M61_FINANCING_TYPES)
    _debug_rows(
        "M61 pre-filter diagnostics (fund labels are informational; rows are NOT dropped by fund here): "
        f"rows_matching_target_fund_labels={int(fund_mask.sum())}/{len(df)} "
        f"rows_matching_liability_types={int(liab_type_mask.sum())}/{len(df)} "
        f"blank_deal_name={int(df['Deal Name'].isna().sum())} "
        f"blank_liability_note={int(df['Liability Note'].isna().sum())}"
    )
    excluded_by_type = df.loc[
        ~liab_type_mask,
        [c for c in ("Deal Name", "Liability Type", "Liability Name", "Liability Note", "Effective Date") if c in df.columns],
    ].copy()
    excluded_by_type["Exclusion Reason"] = "Excluded by Liability Type filter"
    ex_type_counts = (
        excluded_by_type["Liability Type"].fillna("<NA>").astype(str).value_counts().to_dict()
        if "Liability Type" in excluded_by_type.columns
        else {}
    )
    _debug_rows(
        "M61 rows excluded by Liability Type filter: "
        f"excluded_rows={len(excluded_by_type)} excluded_type_counts={ex_type_counts}"
    )
    if not excluded_by_type.empty:
        _debug_rows("TEMP DEBUG: sample rows excluded by Liability Type filter (head 10)")
        for i, (_, er) in enumerate(excluded_by_type.head(10).iterrows(), start=1):
            _debug_rows(
                "TEMP DEBUG:   "
                f"#{i} deal={er.get('Deal Name')!r} type={er.get('Liability Type')!r} "
                f"liability_name={er.get('Liability Name')!r} note={er.get('Liability Note')!r} "
                f"eff={er.get('Effective Date')!r}"
            )
    df = df[df["Liability Type"].isin(M61_FINANCING_TYPES)].copy()
    _debug_rows(f"M61 after Liability Type filter ({M61_FINANCING_TYPES}): rows={len(df)}")
    df = df[keep].copy()

    for col in ["Deal Name", "Liability Note", "Financial Institution", "Liability Name", "Fund Name", "Pledge"]:
        df[col] = df[col].astype(str).str.strip()
    _debug_rows(
        "M61 post-cleaning diagnostics: "
        f"rows={len(df)} blank_deal_name={(df['Deal Name'].str.strip() == '').sum()} "
        f"blank_liability_note={(df['Liability Note'].str.strip() == '').sum()}"
    )

    df["Pledge Date"] = safe_parse_date(df["Pledge Date"])
    df["Effective Date"] = safe_parse_date(df["Effective Date"])
    df["Maturity Date"] = safe_parse_date(df["Maturity Date"])
    df = df.reset_index(drop=True)
    excluded_by_type = excluded_by_type.reset_index(drop=True)
    if return_excluded:
        return df, excluded_by_type
    return df


def load_file_b(path: str) -> pd.DataFrame:
    return load_primary_file(path, "ACORE")


def normalise_text(value):
    if pd.isna(value):
        return ""
    return str(value).strip().lower()


def normalise_facility(raw):
    return FACILITY_NORM_MAP.get(normalise_text(raw), normalise_text(raw))


def extract_liability_note_suffix(value) -> str:
    """Extract trailing deal-id-like token from Liability Note (e.g., LN_Eq_ACPIII_25-2852 -> 25-2852)."""
    if pd.isna(value):
        return ""
    s = str(value).strip()
    if not s:
        return ""
    m = re.search(r"([A-Za-z0-9]+-\d+)$", s)
    if m:
        return m.group(1).strip()
    parts = [p.strip() for p in s.split("_") if p and p.strip()]
    if not parts:
        return ""
    tail = parts[-1]
    return tail if re.search(r"\d+-\d+", tail) else ""


def normalise_deal_id_key(value) -> str:
    if pd.isna(value):
        return ""
    return str(value).strip().lower()


def _source_bucket(value) -> str:
    s = normalise_text(value)
    if not s:
        return ""
    if "subline" in s:
        return "subline"
    if "repo" in s:
        return "repo"
    if s in ("non", "non-repo", "non repo") or " non" in f" {s}":
        return "non"
    return s


def date_key(series):
    return safe_parse_date(series).dt.strftime("%Y-%m-%d").fillna("")


def add_m61_facility_for_match_key(df: pd.DataFrame) -> pd.DataFrame:
    """Build a facility-like string on M61 rows for ``build_match_key`` (not display).

    Primary Fin Inpt uses **Facility** (e.g. ``GS Repo``). M61 **Liability Name** is not a stable join
    key (repeats across rows). Combine **Financial Institution** + **Liability Type** so
    ``normalise_facility`` can align with the model (e.g. ``JPM`` + ``Repo`` → ``jpm repo``).
    Falls back to **Liability Name** only when FI + type would be blank.
    """
    out = df.copy()
    fi = out["Financial Institution"].fillna("").astype(str).str.strip()
    lt = out["Liability Type"].fillna("").astype(str).str.strip()
    combo = (fi + " " + lt).str.strip()
    liab_name = out["Liability Name"].fillna("").astype(str).str.strip()
    out["_m61_facility_for_match"] = combo.where(combo.ne(""), liab_name)
    return out


def _has_compare_value(v):
    if pd.isna(v):
        return False
    return str(v).strip().lower() not in ("", "nan", "none")


def compare_values(val_a, val_b, comparison_type):
    if pd.isna(val_a) and pd.isna(val_b):
        return "MATCH"
    if pd.isna(val_a) or pd.isna(val_b):
        return "MISMATCH"

    if comparison_type == "numeric":
        try:
            return "MATCH" if abs(float(val_a) - float(val_b)) <= FLOAT_TOLERANCE else "MISMATCH"
        except Exception:
            return "MISMATCH"

    if comparison_type == "date":
        s_a = safe_parse_date(pd.Series([val_a])).iloc[0]
        s_b = safe_parse_date(pd.Series([val_b])).iloc[0]
        if pd.isna(s_a) and pd.isna(s_b):
            return "MATCH"
        if pd.isna(s_a) or pd.isna(s_b):
            return "MISMATCH"
        return "MATCH" if s_a.normalize().date() == s_b.normalize().date() else "MISMATCH"

    return "MATCH" if normalise_text(val_a) == normalise_text(val_b) else "MISMATCH"


def compare_liability_primary_status(val_liab, val_acp, comparison_type, *, missing_in_primary: str):
    """
    Liability (export) vs primary model: explicit missing-side labels instead of generic MISSING/MISMATCH
    when a value is absent on one side.
    """
    l_ok = _has_compare_value(val_liab)
    p_ok = _has_compare_value(val_acp)
    if not l_ok and not p_ok:
        return "BOTH MISSING"
    if not p_ok and l_ok:
        return f"MISSING IN {missing_in_primary}"
    if p_ok and not l_ok:
        return "MISSING IN M61"
    return compare_values(val_liab, val_acp, comparison_type)


def compare_optional(val_liab, val_acp, kind="text"):
    """
    Optional cross-file fields (index floor/name, recourse).
    - Both sides blank → N/A
    - Only one side populated → N/A (values still shown; not a mismatch)
    - Both populated → MATCH / MISMATCH (never affects recon_status)
    """
    l_ok = not _is_blank_for_compare(val_liab)
    a_ok = not _is_blank_for_compare(val_acp)
    if not l_ok and not a_ok:
        return "N/A"
    if not l_ok or not a_ok:
        return "N/A"
    eff_kind = kind
    if kind == "numeric":
        try:
            float(val_liab)
            float(val_acp)
        except (TypeError, ValueError):
            eff_kind = "text"
    return compare_values(val_liab, val_acp, eff_kind)


def _effective_date_cell_populated(v):
    """
    True when the cell resolves to a real calendar date (datetime, string date, Excel serial, etc.).
    Uses the same parsing path as the rest of the tool so primary pledge dates are not treated as blank.
    """
    if v is None or isinstance(v, bool):
        return False
    try:
        if pd.isna(v):
            return False
    except (TypeError, ValueError):
        return False

    if isinstance(v, str):
        s = v.strip()
        if not s or s.lower() in ("nan", "none", "nat", "<na>", "-", "—", "n/a"):
            return False

    dt = safe_parse_date(pd.Series([v])).iloc[0]
    if pd.isna(dt):
        return False

    # Small integers often mis-parse as 1970-01-01 via ns-epoch heuristics in to_datetime.
    try:
        fv = float(v)
    except (TypeError, ValueError):
        fv = None
    if fv is not None and abs(fv) < 1000:
        if dt.normalize() == pd.Timestamp("1970-01-01").normalize():
            return False

    return True


def compare_effective_date_status(val_liability, val_acp, *, missing_in_primary: str = "ACORE"):
    has_acp = _effective_date_cell_populated(val_acp)
    has_liab = _effective_date_cell_populated(val_liability)
    miss_pri = f"MISSING IN {missing_in_primary}"

    if not has_acp and not has_liab:
        return "BOTH MISSING"
    if not has_acp:
        return miss_pri
    if not has_liab:
        return "MISSING IN M61"

    dt_liab = safe_parse_date(pd.Series([val_liability])).iloc[0]
    dt_acp = safe_parse_date(pd.Series([val_acp])).iloc[0]
    if pd.isna(dt_acp):
        return miss_pri
    if pd.isna(dt_liab):
        return "MISSING IN M61"
    return "MATCH" if dt_liab.normalize().date() == dt_acp.normalize().date() else "NO MATCH"


def _recon_token_for_effective_date_status(display):
    if display == "MATCH":
        return "MATCH"
    if display == "NO MATCH":
        return "MISMATCH"
    return "MISSING"


def _recon_token_for_compare_status(display):
    if display == "MATCH":
        return "MATCH"
    if display == "MISMATCH":
        return "MISMATCH"
    return "MISSING"


def _is_blank_for_compare(v):
    if v is None or pd.isna(v):
        return True
    s = str(v).strip().lower()
    return s in ("", "nan", "none", "nat", "<na>")


def compare_pledge_date_status(
    *, val_liability, val_acp, missing_in_primary: str = "ACORE"
):
    has_acp = _effective_date_cell_populated(val_acp)
    has_liab = _effective_date_cell_populated(val_liability)
    miss_pri = f"MISSING IN {missing_in_primary}"

    if not has_acp and not has_liab:
        return "BOTH MISSING"
    if not has_acp and has_liab:
        return miss_pri
    if has_acp and not has_liab:
        return "MISSING IN M61"

    dt_l = safe_parse_date(pd.Series([val_liability])).iloc[0]
    dt_a = safe_parse_date(pd.Series([val_acp])).iloc[0]
    if pd.isna(dt_a):
        return miss_pri
    if pd.isna(dt_l):
        return "MISSING IN M61"
    if dt_l.normalize().date() == dt_a.normalize().date():
        return "MATCH"
    return "MISMATCH"


def build_match_key(df, deal_col, facility_col, note_col, effective_date_col):
    """Join key = normalized deal + facility + effective date (``note_col`` is unused; kept for call-site clarity)."""
    out = df.copy()
    out["deal_norm"] = out[deal_col].apply(normalise_text)
    out["facility_norm"] = out[facility_col].apply(normalise_facility)
    out["effective_date_key"] = date_key(out[effective_date_col])
    out["match_key"] = (
        out["deal_norm"]
        + " | "
        + out["facility_norm"]
        + " | "
        + out["effective_date_key"]
    )
    return out


def _apply_deal_id_suffix_fallback_matches(
    merged: pd.DataFrame, df_b: pd.DataFrame, df_a: pd.DataFrame, *, label_a: str
) -> tuple[pd.DataFrame, int]:
    """
    Fallback aid only for currently-unmatched rows:
    pair one-to-one rows where primary Deal ID == extracted Liability Note suffix.
    """
    need_b = {"match_key", "deal_id_key"}
    need_a = {"match_key", "liability_note_suffix_key"}
    if not need_b.issubset(df_b.columns) or not need_a.issubset(df_a.columns):
        return merged, 0
    if "_merge" not in merged.columns:
        return merged, 0

    mk_a = f"{label_a}_match_key"
    right_only_keys = set(merged.loc[merged["_merge"] == "right_only", mk_a].dropna().tolist())
    left_only_keys = set(merged.loc[merged["_merge"] == "left_only", "match_key"].dropna().tolist())
    if not right_only_keys or not left_only_keys:
        return merged, 0

    b_pool = df_b[df_b["match_key"].isin(left_only_keys)].copy()
    a_pool = df_a[df_a["match_key"].isin(right_only_keys)].copy()
    b_pool = b_pool[b_pool["deal_id_key"].ne("")].copy()
    a_pool = a_pool[a_pool["liability_note_suffix_key"].ne("")].copy()
    if b_pool.empty or a_pool.empty:
        return merged, 0

    b_counts = b_pool["deal_id_key"].value_counts(dropna=False)
    a_counts = a_pool["liability_note_suffix_key"].value_counts(dropna=False)
    unique_keys = set(b_counts[b_counts == 1].index).intersection(set(a_counts[a_counts == 1].index))
    if not unique_keys:
        return merged, 0

    b_pick = b_pool[b_pool["deal_id_key"].isin(unique_keys)].copy()
    a_pick = a_pool[a_pool["liability_note_suffix_key"].isin(unique_keys)].copy()
    fb = pd.merge(
        b_pick,
        a_pick.add_prefix(f"{label_a}_"),
        left_on="deal_id_key",
        right_on=f"{label_a}_liability_note_suffix_key",
        how="inner",
    )
    if fb.empty:
        return merged, 0

    fb["_merge"] = "both"
    fb["_fallback_match_method"] = "deal_id_suffix"

    used_b = set(fb["match_key"].dropna().tolist())
    used_a = set(fb[mk_a].dropna().tolist())
    rm_left = (merged["_merge"] == "left_only") & merged["match_key"].isin(used_b)
    rm_right = (merged["_merge"] == "right_only") & merged[mk_a].isin(used_a)
    merged_keep = merged.loc[~(rm_left | rm_right)].copy()
    merged_out = pd.concat([merged_keep, fb], ignore_index=True, sort=False)
    return merged_out, int(len(fb))


def _merged_liab_col(row: pd.Series, label_a: str, source_col: str):
    key = f"{label_a}_{source_col}"
    if key not in row.index:
        return pd.NA
    return row[key]


def _first_liab_recourse_value(row: pd.Series, label_a: str):
    # Liability export uses "Recourse" in practice; keep fallbacks for other templates.
    for src in ("Recourse", "Recourse %", "RecoursePct"):
        key = f"{label_a}_{src}"
        if key not in row.index:
            continue
        v = row[key]
        if pd.isna(v):
            continue
        if isinstance(v, str) and not str(v).strip():
            continue
        return v
    return pd.NA


def _stripped_nonempty_str(v) -> str | None:
    """Return a stripped display string, or None if absent/blank (for coalescing recon columns)."""
    if v is None:
        return None
    try:
        if pd.isna(v):
            return None
    except (TypeError, ValueError):
        return None
    s = str(v).strip()
    if not s or s.lower() in ("nan", "<na>", "nat", "none"):
        return None
    return s


def load_liability_cre_mapping(path: str) -> pd.DataFrame:
    """
    Loads LiabilityNoteID -> CRENoteID mapping workbook.
    Update sheet_name/column names if the template differs.
    """
    df = pd.read_excel(path, sheet_name=0)
    return df


def reconcile(
    file_a_path,
    file_b_path,
    primary_file_type: str = "ACORE",
    mapping_path: str | None = None,
    uploaded_primary_filename: str | None = None,
):
    label_a = "in_liability"
    p_cfg = get_primary_config(primary_file_type)
    miss_pri = p_cfg["missing_in_primary_label"]
    src_ind_primary = p_cfg.get("primary_only_legend_label") or p_cfg["source_indicator_primary_only"]
    # One fund identity per primary workbook template (no row-level Fund on Fin Inpt).
    business_fund_label = detect_fund_label(uploaded_primary_filename, primary_file_type)
    fund_fallback_display = _fund_cfg(primary_file_type).get("recon_fund_display") or business_fund_label

    raw_m61, excluded_by_type_df = load_file_a(file_a_path, return_excluded=True)
    m61_row_count_after_load = len(raw_m61)
    _debug_rows(f"Reconcile input: M61 rows after load_file_a={len(raw_m61)}")
    _debug_rows(f"TEMP DEBUG: M61 row count immediately after load_file_a = {len(raw_m61)}")
    _debug_m61_load_preview(raw_m61)

    if primary_file_type == "AOC II" and mapping_path:
        df_map = load_liability_cre_mapping(mapping_path)
        raw_m61 = raw_m61.merge(
            df_map,
            how="left",
            left_on="Liability Note",
            right_on="LiabilityNoteID",
        )
        df_a = build_match_key(
            raw_m61,
            "Deal Name",
            "CRENoteID",
            "Liability Note",
            "Effective Date",
        )
        df_pri_raw = load_primary_file(file_b_path, primary_file_type)
        _debug_rows(
            f"TEMP DEBUG: Primary workbook row count immediately after load_primary_file = {len(df_pri_raw)}"
        )
        df_b = build_match_key(
            df_pri_raw,
            "Deal Name",
            "Note Name",
            "Note Name",
            "Effective Date",
        )
        _debug_match_key_sample_rows(
            "primary (Fin Inpt)",
            df_b,
            deal_col="Deal Name",
            facility_col="Note Name",
            note_col="Note Name",
            eff_col="Effective Date",
        )
        _debug_match_key_sample_rows(
            "M61",
            df_a,
            deal_col="Deal Name",
            facility_col="CRENoteID",
            note_col="Liability Note",
            eff_col="Effective Date",
        )
    else:
        # Match key: Deal Name + facility token + Effective Date (legacy V7 parity). Primary uses
        # ``Facility``; M61 uses ``Liability Name`` in that slot — same pairing as Financing Line
        # Reconciliation Tool V7 (FI + Liability Type was not aligning to Fin Inpt Facility strings).
        df_a = build_match_key(
            raw_m61,
            "Deal Name",
            "Liability Name",
            "Liability Note",
            "Effective Date",
        )
        df_pri_raw = load_primary_file(file_b_path, primary_file_type)
        _debug_rows(
            f"TEMP DEBUG: Primary workbook row count immediately after load_primary_file = {len(df_pri_raw)}"
        )
        df_b = build_match_key(
            df_pri_raw,
            "Deal Name",
            "Facility",
            "Note Name",
            "Effective Date",
        )
        _debug_match_key_sample_rows(
            "primary (Fin Inpt)",
            df_b,
            deal_col="Deal Name",
            facility_col="Facility",
            note_col="Note Name",
            eff_col="Effective Date",
        )
        _debug_match_key_sample_rows(
            "M61",
            df_a,
            deal_col="Deal Name",
            facility_col="Liability Name",
            note_col="Liability Note",
            eff_col="Effective Date",
        )

    # Helper IDs for validation / fallback alignment aid.
    if "Deal ID" in df_b.columns:
        df_b["deal_id_value"] = df_b["Deal ID"].apply(lambda v: "" if pd.isna(v) else str(v).strip())
    else:
        df_b["deal_id_value"] = ""
    df_b["deal_id_key"] = df_b["deal_id_value"].apply(normalise_deal_id_key)
    df_a["liability_note_suffix"] = df_a["Liability Note"].apply(extract_liability_note_suffix)
    df_a["liability_note_suffix_key"] = df_a["liability_note_suffix"].apply(normalise_deal_id_key)
    _debug_rows(
        "TEMP DEBUG: Deal ID helper readiness — "
        f"primary_nonblank_deal_id={int(df_b['deal_id_key'].ne('').sum())} "
        f"m61_nonblank_note_suffix={int(df_a['liability_note_suffix_key'].ne('').sum())}"
    )
    _debug_rows("TEMP DEBUG: Deal ID helper sample — primary head 10")
    for i, (_, r) in enumerate(
        df_b.loc[:, [c for c in ["Deal Name", "deal_id_value", "effective_date_key", "match_key"] if c in df_b.columns]]
        .head(10)
        .iterrows(),
        start=1,
    ):
        _debug_rows(
            "TEMP DEBUG:   "
            f"#{i} deal={r.get('Deal Name')!r} deal_id={r.get('deal_id_value')!r} "
            f"eff_key={r.get('effective_date_key')!r} match_key={r.get('match_key')!r}"
        )
    _debug_rows("TEMP DEBUG: Deal ID helper sample — M61 head 10")
    for i, (_, r) in enumerate(
        df_a.loc[
            :,
            [c for c in ["Deal Name", "Liability Note", "liability_note_suffix", "effective_date_key", "match_key"] if c in df_a.columns],
        ]
        .head(10)
        .iterrows(),
        start=1,
    ):
        _debug_rows(
            "TEMP DEBUG:   "
            f"#{i} deal={r.get('Deal Name')!r} liability_note={r.get('Liability Note')!r} "
            f"suffix={r.get('liability_note_suffix')!r} eff_key={r.get('effective_date_key')!r} "
            f"match_key={r.get('match_key')!r}"
        )

    _debug_match_key_overlap_diagnosis(df_b, df_a, n=10)

    # Informational only: do not drop M61 rows here — outer merge below already carries
    # M61-only / primary-only sides; pre-filtering df_a hid M61 export rows absent from the model.
    primary_match_keys = set(df_b["match_key"].dropna().tolist())
    m61_in_primary = int(df_a["match_key"].isin(primary_match_keys).sum())
    _debug_rows(
        "M61 vs primary match_key overlap (no row drop): "
        f"m61_rows={len(df_a)} primary_rows={len(df_b)} "
        f"m61_rows_with_key_in_primary={m61_in_primary} distinct_primary_keys={len(primary_match_keys)}"
    )

    # ---- Row-level staged matching (strict -> source-aware -> fallback), then outer-preserving assembly ----
    b = df_b.copy()
    a = df_a.copy()
    b["_row_id_b"] = range(len(b))
    a["_row_id_a"] = range(len(a))
    b["strict_note_norm"] = b["Note Name"].apply(normalise_text) if "Note Name" in b.columns else ""
    a["strict_note_norm"] = a["Liability Note"].apply(normalise_text) if "Liability Note" in a.columns else ""
    b["strict_key"] = (
        b["deal_norm"] + " | " + b["facility_norm"] + " | " + b["strict_note_norm"] + " | " + b["effective_date_key"]
    )
    a["strict_key"] = (
        a["deal_norm"] + " | " + a["facility_norm"] + " | " + a["strict_note_norm"] + " | " + a["effective_date_key"]
    )
    b["deal_date_key"] = b["deal_norm"] + " | " + b["effective_date_key"]
    a["deal_date_key"] = a["deal_norm"] + " | " + a["effective_date_key"]
    b["source_bucket"] = b["Source"].apply(_source_bucket) if "Source" in b.columns else ""
    a["source_bucket"] = a["Liability Type"].apply(_source_bucket) if "Liability Type" in a.columns else ""
    b["source_aware_key"] = b["deal_date_key"] + " | " + b["source_bucket"]
    a["source_aware_key"] = a["deal_date_key"] + " | " + a["source_bucket"]
    fcfg = _fund_cfg(primary_file_type)
    fpattern = fcfg.get("fund_regex") or (
        re.escape(fcfg.get("fund_token")) if fcfg.get("fund_token") else None
    )
    if fpattern:
        a["_m61_in_scope"] = (
            a["Fund Name"].fillna("").astype(str).str.contains(fpattern, case=False, regex=True, na=False)
        )
    else:
        a["_m61_in_scope"] = True
    _debug_rows(
        "TEMP DEBUG: staged matcher fund-scope guard on M61 candidates — "
        f"in_scope={int(a['_m61_in_scope'].sum())}/{len(a)} "
        f"pattern={fpattern!r} primary_type={primary_file_type}"
    )

    unmatched_b = set(b["_row_id_b"].tolist())
    unmatched_a = set(a["_row_id_a"].tolist())
    matchable_a = set(a.loc[a["_m61_in_scope"], "_row_id_a"].tolist())
    pair_rows: list[dict] = []
    b_by_id = b.set_index("_row_id_b", drop=False)
    a_by_id = a.set_index("_row_id_a", drop=False)

    def _pair_by_key(key_b: str, key_a: str, stage: str) -> int:
        lb = b[b["_row_id_b"].isin(unmatched_b)].copy()
        la = a[a["_row_id_a"].isin(unmatched_a.intersection(matchable_a))].copy()
        lb = lb[lb[key_b].astype(str).str.strip().ne("")]
        la = la[la[key_a].astype(str).str.strip().ne("")]
        if lb.empty or la.empty:
            return 0
        lb = lb.sort_values(["Deal Name", "effective_date_key", "_row_id_b"]).copy()
        la = la.sort_values(["Deal Name", "effective_date_key", "_row_id_a"]).copy()
        lb["_rk"] = lb.groupby(key_b).cumcount()
        la["_rk"] = la.groupby(key_a).cumcount()
        pairs = lb[["_row_id_b", key_b, "_rk"]].merge(
            la[["_row_id_a", key_a, "_rk"]],
            left_on=[key_b, "_rk"],
            right_on=[key_a, "_rk"],
            how="inner",
        )
        if pairs.empty:
            return 0
        for _, pr in pairs.iterrows():
            rid_b = int(pr["_row_id_b"])
            rid_a = int(pr["_row_id_a"])
            if rid_b not in unmatched_b or rid_a not in unmatched_a:
                continue
            unmatched_b.remove(rid_b)
            unmatched_a.remove(rid_a)
            pair_rows.append({"_row_id_b": rid_b, "_row_id_a": rid_a, "_match_stage": stage, "_merge": "both"})
            br = b_by_id.loc[rid_b]
            ar = a_by_id.loc[rid_a]
            _debug_rows(
                "TEMP DEBUG: selected pair "
                f"stage={stage} acp_id={rid_b} m61_id={rid_a} "
                f"deal={br.get('Deal Name')!r} eff_key={br.get('effective_date_key')!r} "
                f"acp_source={br.get('Source')!r} m61_type={ar.get('Liability Type')!r}"
            )
        return len(pairs)

    strict_n = _pair_by_key("strict_key", "strict_key", "strict")
    _debug_rows(f"TEMP DEBUG: staged matcher strict matches={strict_n}")

    # Required debugging: candidate ACP rows, source-filter match, selected row.
    cand_logged = 0
    lb_dbg = b[b["_row_id_b"].isin(unmatched_b)].copy()
    la_dbg = a[a["_row_id_a"].isin(unmatched_a.intersection(matchable_a))].copy()
    common_dd = sorted(set(lb_dbg["deal_date_key"]).intersection(set(la_dbg["deal_date_key"])))
    for dd in common_dd:
        p_cand = lb_dbg[lb_dbg["deal_date_key"] == dd]
        m_cand = la_dbg[la_dbg["deal_date_key"] == dd]
        if p_cand.empty or m_cand.empty:
            continue
        deal_name_dbg = str(p_cand.iloc[0].get("Deal Name", ""))
        if cand_logged < 40 or "geoffrey drive" in normalise_text(deal_name_dbg):
            _debug_rows(
                "TEMP DEBUG: candidate set (source-aware) "
                f"deal_date={dd!r} acp_candidates={len(p_cand)} m61_candidates={len(m_cand)}"
            )
            for _, pr in p_cand.iterrows():
                _debug_rows(
                    "TEMP DEBUG:   ACP cand "
                    f"id={int(pr['_row_id_b'])} deal={pr.get('Deal Name')!r} src={pr.get('Source')!r} "
                    f"source_bucket={pr.get('source_bucket')!r} eff_key={pr.get('effective_date_key')!r}"
                )
            for _, ar in m_cand.iterrows():
                _debug_rows(
                    "TEMP DEBUG:   M61 cand "
                    f"id={int(ar['_row_id_a'])} deal={ar.get('Deal Name')!r} liab_type={ar.get('Liability Type')!r} "
                    f"source_bucket={ar.get('source_bucket')!r} eff_key={ar.get('effective_date_key')!r}"
                )
            src_match_n = len(
                p_cand[["_row_id_b", "source_bucket"]]
                .merge(
                    m_cand[["_row_id_a", "source_bucket"]],
                    on="source_bucket",
                    how="inner",
                )
            )
            _debug_rows(f"TEMP DEBUG:   source-aware candidate links={src_match_n}")
            cand_logged += 1

    source_n = _pair_by_key("source_aware_key", "source_aware_key", "source_aware")
    _debug_rows(f"TEMP DEBUG: staged matcher source-aware matches={source_n}")

    fallback_n = _pair_by_key("deal_date_key", "deal_date_key", "fallback")
    _debug_rows(f"TEMP DEBUG: staged matcher fallback matches={fallback_n}")

    # Build merged-like frame while preserving unmatched rows (outer behavior).
    for rid_b in sorted(unmatched_b):
        pair_rows.append({"_row_id_b": int(rid_b), "_row_id_a": pd.NA, "_match_stage": "none", "_merge": "left_only"})
    for rid_a in sorted(unmatched_a):
        pair_rows.append({"_row_id_b": pd.NA, "_row_id_a": int(rid_a), "_match_stage": "none", "_merge": "right_only"})

    map_df = pd.DataFrame(pair_rows)
    merged = map_df.merge(b, on="_row_id_b", how="left")
    a_pref = a.add_prefix(f"{label_a}_")
    merged = merged.merge(
        a_pref,
        left_on="_row_id_a",
        right_on=f"{label_a}__row_id_a",
        how="left",
    )
    _debug_unmatched_after_merge(merged, label_a=label_a, n=10)
    _debug_rows(
        "TEMP DEBUG: fallback pairing on Deal ID/Liability Note suffix is DISABLED "
        "(diagnosis mode; primary match_key only)."
    )
    _debug_rows(
        "Reconcile merge rows: "
        f"primary_rows={len(df_b)} m61_rows={len(df_a)} merged_rows={len(merged)} "
        f"merge_indicator={merged['_merge'].value_counts(dropna=False).to_dict()}"
    )
    _debug_rows(f"TEMP DEBUG: row count immediately after pd.merge = {len(merged)}")
    _vc = merged["_merge"].value_counts(dropna=False)
    _debug_rows(
        "TEMP DEBUG: merged['_merge'].value_counts(dropna=False) — "
        "left_only=Fin Inpt only, right_only=M61 only, both=matched key"
    )
    for _k, _v in _vc.items():
        _debug_rows(f"TEMP DEBUG:   {_k} = {_v}")
    _debug_rows(
        "[TEMP VALIDATION] primary read + M61 read + merge: "
        f"primary_rows_after_load={len(df_pri_raw)} "
        f"m61_rows_after_load={m61_row_count_after_load} "
        f"merged_rows={len(merged)} "
        f"merged['_merge'].value_counts(dropna=False)={merged['_merge'].value_counts(dropna=False).to_dict()}"
    )

    liability_extra = [
        "Current Balance",
        "Liability Note",
        "Financial Institution",
        "Maturity Date",
        "Fund Name",
        "Liability Name",
        "target",
        "Status",
    ]

    rows = []

    for _, row in merged.iterrows():
        a_deal = row.get(f"{label_a}_Deal Name")
        b_deal = row.get("Deal Name")
        in_a = not pd.isna(a_deal)
        in_b = not pd.isna(b_deal)

        record = {col: row.get(col) for col in ACP_SHEET_COLS}

        if not in_b and in_a:
            record["Deal Name"] = row.get(f"{label_a}_Deal Name")
            record["Facility"] = row.get(f"{label_a}_Liability Name")
            record["Note Name"] = row.get(f"{label_a}_Liability Note")
            record["Pledge"] = row.get(f"{label_a}_Pledge")
            record["Pledge Date"] = row.get(f"{label_a}_Pledge Date")
            record["Effective Date"] = row.get(f"{label_a}_Effective Date")
            record["Maturity Date"] = row.get(f"{label_a}_Maturity Date")

        facility_raw = record.get("Facility") or row.get(f"{label_a}_Liability Name")
        facility_norm = normalise_facility(facility_raw)
        record["Financial Line"] = f"{record.get('Deal Name', '')} & {str(facility_norm).upper()}"
        record["match_key"] = (
            row.get("match_key")
            if not pd.isna(row.get("match_key"))
            else row.get(f"{label_a}_match_key")
        )

        # Filter-friendly Facility: when the primary model left Facility blank but M61 has a liability name.
        fac_pri = _stripped_nonempty_str(record.get("Facility"))
        liab_name = _stripped_nonempty_str(row.get(f"{label_a}_Liability Name"))
        if not fac_pri and liab_name and in_a:
            record["Facility"] = liab_name

        # Filter-friendly Source: primary "Source" is empty for M61-only rows and sometimes on matches;
        # use M61 Liability Type + Financial Institution when M61 data exists.
        src_pri = _stripped_nonempty_str(record.get("Source"))
        if not src_pri and in_a:
            lt = _stripped_nonempty_str(row.get(f"{label_a}_Liability Type"))
            fi = _stripped_nonempty_str(row.get(f"{label_a}_Financial Institution"))
            parts = [p for p in (lt, fi) if p]
            if parts:
                record["Source"] = " | ".join(parts)

        if in_a and in_b:
            record["Source Indicator"] = "Both"
        elif in_b:
            record["Source Indicator"] = src_ind_primary
        else:
            record["Source Indicator"] = "M61 Only"

        deal_id_acp = row.get("deal_id_value") if in_b else ""
        note_suffix_m61 = row.get(f"{label_a}_liability_note_suffix") if in_a else ""
        deal_id_key = normalise_deal_id_key(deal_id_acp)
        note_suffix_key = normalise_deal_id_key(note_suffix_m61)
        record["Deal ID (ACP)"] = deal_id_acp if deal_id_key else pd.NA
        record["Liability Note Suffix (M61)"] = note_suffix_m61 if note_suffix_key else pd.NA

        record["Target Advance Rate (M61)"] = (
            row.get(f"{label_a}_Target Advance Rate") if in_a else pd.NA
        )

        in_liability_raw = row.get(f"{label_a}_in_liability")
        in_liability_value = "" if pd.isna(in_liability_raw) else str(in_liability_raw).strip().lower()
        only_target_from_invis = in_liability_value == "invis"

        key_field_statuses = []

        for b_field, a_field, ctype in COMPARE_FIELDS:
            val_a = row.get(f"{label_a}_{a_field}")
            val_b = row.get(b_field)
            liability_label = LIABILITY_VALUE_LABELS.get(a_field, f"{a_field} (M61)")

            if only_target_from_invis and a_field != "target":
                record[liability_label] = pd.NA
                record[f"{b_field} Status"] = "MISSING IN M61"
                if b_field in RECON_STATUS_FIELDS:
                    key_field_statuses.append("MISSING")
                continue

            # Advance Rate status comparison basis:
            # - Repo / Non: compare primary Advance Rate vs M61 Target Advance Rate
            # - ACP II (legacy behavior): keep comparing vs Target Advance Rate
            liab_type_lc = (
                str(row.get(f"{label_a}_Liability Type")).strip().lower() if in_a else ""
            )
            use_target_for_adv_compare = liab_type_lc in {"repo", "non"} or primary_file_type == "ACP II"
            if b_field == "Advance Rate" and use_target_for_adv_compare:
                record["Advance Rate (M61)"] = (
                    row.get(f"{label_a}_Current Advance Rate") if in_a else pd.NA
                )
                val_a = row.get(f"{label_a}_Target Advance Rate") if in_a else pd.NA
            else:
                record[liability_label] = val_a

            if b_field == "Effective Date":
                if not in_b:
                    ed_status = f"MISSING IN {miss_pri}"
                elif not in_a:
                    ed_status = "MISSING IN M61"
                else:
                    ed_status = compare_effective_date_status(val_a, val_b, missing_in_primary=miss_pri)
                record[f"{b_field} Status"] = ed_status
                if b_field in RECON_STATUS_FIELDS:
                    key_field_statuses.append(_recon_token_for_effective_date_status(ed_status))
                continue

            if not in_a and not in_b:
                status = "BOTH MISSING"
            elif not in_b:
                status = f"MISSING IN {miss_pri}"
            elif not in_a:
                status = "MISSING IN M61"
            else:
                status = compare_liability_primary_status(
                    val_a, val_b, ctype, missing_in_primary=miss_pri
                )

            record[f"{b_field} Status"] = status
            if b_field in RECON_STATUS_FIELDS:
                key_field_statuses.append(_recon_token_for_compare_status(status))

        for a_col in liability_extra:
            business_col = LIABILITY_VALUE_LABELS.get(a_col, f"{a_col} (M61)")
            if business_col in record:
                continue
            record[business_col] = (
                pd.NA if (only_target_from_invis and a_col != "target") else row.get(f"{label_a}_{a_col}")
            )

        pliab = pd.NA if only_target_from_invis else (row.get(f"{label_a}_Pledge") if in_a else pd.NA)
        pacp = row.get("Pledge") if in_b else pd.NA
        pdt_liab = pd.NA if only_target_from_invis else (row.get(f"{label_a}_Pledge Date") if in_a else pd.NA)
        pdt_acp = row.get("Pledge Date") if in_b else pd.NA

        record["Pledge (ACP)"] = pacp
        record["Pledge (M61)"] = pliab
        record["Pledge Date (ACP)"] = pdt_acp
        record["Pledge Date (M61)"] = pdt_liab
        record["Pledge Date Status"] = compare_pledge_date_status(
            val_liability=pdt_liab,
            val_acp=pdt_acp,
            missing_in_primary=miss_pri,
        )

        m61_fund_s = _stripped_nonempty_str(row.get(f"{label_a}_Fund Name")) if in_a else None
        # Prefer M61 export Fund Name; when missing, use M61-style display name (not export_label).
        if m61_fund_s:
            record["Fund"] = m61_fund_s
        else:
            record["Fund"] = fund_fallback_display
        record["Effective Date (ACP)"] = record.get("Effective Date") if in_b else pd.NA
        record["Advance Rate (ACP)"] = record.get("Advance Rate") if in_b else pd.NA
        record["Spread (ACP)"] = record.get("Spread") if in_b else pd.NA
        record["Undrawn Capacity (ACP)"] = record.get("Current Undrawn Capacity") if in_b else pd.NA
        record["Undrawn Capacity (M61)"] = record.get("Current Undrawn Capacity (M61)")

        # ACP-side values
        record["Index Floor (ACP)"] = record.get("Floor") if in_b else pd.NA
        record["Index Name (ACP)"] = pd.NA
        record["Recourse % (ACP)"] = record.get("Recourse %") if in_b else pd.NA

        # expose existing Undrawn comparison status
        record["Undrawn Capacity Status"] = record.get("Current Undrawn Capacity Status", "N/A")

        # Liability-side values
        if only_target_from_invis or not in_a:
            record["Index Floor (M61)"] = pd.NA
            record["Index Name (M61)"] = pd.NA
            record["Recourse % (M61)"] = pd.NA
        else:
            ix_fl = _merged_liab_col(row, label_a, "IndexFloor")
            ix_nm = _merged_liab_col(row, label_a, "IndexName")
            if pd.notna(ix_nm):
                ix_nm = str(ix_nm).strip() or pd.NA

            record["Index Floor (M61)"] = ix_fl
            record["Index Name (M61)"] = ix_nm
            record["Recourse % (M61)"] = _first_liab_recourse_value(row, label_a)

        record["Index Floor Status"] = compare_optional(
            record.get("Index Floor (M61)"),
            record.get("Index Floor (ACP)"),
            "text",
        )
        record["Index Name Status"] = compare_optional(
            record.get("Index Name (M61)"),
            record.get("Index Name (ACP)"),
            "text",
        )
        record["Recourse % Status"] = compare_optional(
            record.get("Recourse % (M61)"),
            record.get("Recourse % (ACP)"),
            "numeric",
        )

        src_status = "" if pd.isna(row.get(f"{label_a}_Status")) else str(row.get(f"{label_a}_Status")).strip().lower()

        if "red" in src_status:
            recon_status = "MISSING"
        elif not in_a or not in_b:
            recon_status = "MISSING"
        elif "MISSING" in key_field_statuses:
            recon_status = "MISSING"
        elif "MISMATCH" in key_field_statuses:
            recon_status = "MISMATCH"
        else:
            recon_status = "MATCH"

        record["recon_status"] = recon_status
        rows.append(record)

    df_out = pd.DataFrame(rows).reindex(columns=RECON_ORDERED_COLS)
    _debug_rows(f"Reconciliation output rows (df_out)={len(df_out)}")

    adv_rate_col = f"{p_cfg['display_name']} Advance Rate"
    adv_rows = []
    for _, row in merged.iterrows():
        a_deal = row.get(f"{label_a}_Deal Name")
        b_deal = row.get("Deal Name")
        in_a = not pd.isna(a_deal)
        in_b = not pd.isna(b_deal)
        deal = b_deal if in_b else (a_deal if in_a else pd.NA)
        acp_adv = row["Advance Rate"] if in_b and not pd.isna(row.get("Advance Rate")) else pd.NA

        for col in LIABILITY_ADVANCE_RATE_COLUMNS:
            pk = f"{label_a}_{col}"
            liab_val = row.get(pk) if pk in row.index else pd.NA
            if not in_a and not in_b:
                result = "BOTH MISSING"
            elif not in_b:
                result = f"MISSING IN {miss_pri}"
            elif not in_a:
                result = "MISSING IN M61"
            else:
                result = compare_liability_primary_status(
                    liab_val, acp_adv, "numeric", missing_in_primary=miss_pri
                )
            adv_rows.append(
                {
                    "Deal": deal,
                    adv_rate_col: acp_adv,
                    "M61 Column": col,
                    "M61 Value": liab_val,
                    "Result": result,
                }
            )

    df_adv = pd.DataFrame(adv_rows)
    return df_out, df_adv, excluded_by_type_df


# --------------------------------------------------
# 8. EXCEL STYLING HELPERS
# --------------------------------------------------
HEADER_BG = "1F3864"
SUBHDR_BG = "2F5597"
MATCH_BG = "C6EFCE"
MISMATCH_BG = "FFC7CE"
MISSING_BG = "FFEB9C"
WHITE = "FFFFFF"
LIGHT_GRAY = "F2F2F2"

thin = Side(style="thin", color="CCCCCC")
BORDER = Border(left=thin, right=thin, top=thin, bottom=thin)

# Data rows (reconciliation grid from row 5): taller rows + wrap so status text is visible.
EXCEL_RECON_DATA_ROW_HEIGHT = 44


def _fill(hex_color):
    return PatternFill("solid", fgColor=hex_color)


def _hdr_font(size=9, color=WHITE, bold=True):
    return Font(name="Arial", size=size, bold=bold, color=color)


def _body_font(size=9, color="000000", bold=False):
    return Font(name="Arial", size=size, bold=bold, color=color)


def _center():
    return Alignment(horizontal="center", vertical="center", wrap_text=True)


def _left():
    return Alignment(horizontal="left", vertical="center", wrap_text=True)


GROUP_COLORS = {"ID": "2E4057", "ACP": "1A5276", "LIB": "1B4F72", "STATUS": "4A235A"}

COL_DEFS = [
    ("Fund", 18, False, False, "ID"),
    ("Deal Name", 22, False, False, "ID"),
    ("Facility", 14, False, False, "ID"),
    ("Financial Line", 32, False, False, "ID"),
    ("Note Name", 22, False, False, "ID"),
    ("Liability Note (M61)", 26, False, False, "LIB"),
    ("Deal ID (ACP)", 14, False, False, "ACP"),
    ("Liability Note Suffix (M61)", 20, False, False, "LIB"),
    ("Source", 9, False, False, "ID"),
    ("Source Indicator", 15, False, False, "ID"),
    ("Effective Date (ACP)", 14, False, True, "ACP"),
    ("Effective Date (M61)", 16, False, True, "LIB"),
    ("Pledge Date (ACP)", 14, False, True, "ACP"),
    ("Pledge Date (M61)", 16, False, True, "LIB"),
    ("Advance Rate (ACP)", 13, True, False, "ACP"),
    ("Advance Rate (M61)", 16, True, False, "LIB"),
    ("Target Advance Rate (M61)", 19, True, False, "LIB"),
    ("Spread (ACP)", 11, True, False, "ACP"),
    ("Spread (M61)", 14, True, False, "LIB"),
    ("Undrawn Capacity (ACP)", 17, False, False, "ACP"),
    ("Undrawn Capacity (M61)", 20, False, False, "LIB"),
    ("Index Floor (ACP)", 15, False, False, "ACP"),
    ("Index Floor (M61)", 15, False, False, "LIB"),
    ("Index Name (ACP)", 22, False, False, "ACP"),
    ("Index Name (M61)", 22, False, False, "LIB"),
    ("Recourse % (ACP)", 12, True, False, "ACP"),
    ("Recourse % (M61)", 12, True, False, "LIB"),
    ("Advance Rate Status", 15, False, False, "STATUS"),
    ("Spread Status", 13, False, False, "STATUS"),
    ("Effective Date Status", 17, False, False, "STATUS"),
    ("Undrawn Capacity Status", 18, False, False, "STATUS"),
    ("Index Floor Status", 15, False, False, "STATUS"),
    ("Index Name Status", 15, False, False, "STATUS"),
    ("Recourse % Status", 15, False, False, "STATUS"),
    ("Pledge Date Status", 16, False, False, "STATUS"),
    ("Recon Status", 13, False, False, "STATUS"),
]

RECON_COL_MAP = list(RECON_ORDERED_COLS)


def _status_cell(cell, val):
    v = str(val).upper()
    if v in ("N/A", "—", "-"):
        cell.fill = _fill(LIGHT_GRAY)
        cell.font = _body_font(color="888888")
    elif "DIFFERENT" in v or "MISMATCH" in v or "NO MATCH" in v:
        cell.fill = _fill(MISMATCH_BG)
        cell.font = _body_font(bold=True, color="9C0006")
    elif "BOTH MISSING" in v or "MISSING" in v:
        cell.fill = _fill(MISSING_BG)
        cell.font = _body_font(bold=True, color="7D6608")
    elif "MATCH" in v and "MIS" not in v:
        cell.fill = _fill(MATCH_BG)
        cell.font = _body_font(bold=True, color="375623")
    else:
        cell.fill = _fill(LIGHT_GRAY)
        cell.font = _body_font(color="888888")
    cell.alignment = _center()
    cell.border = BORDER


def _fmt_date(v):
    if pd.isna(v) or str(v) in ("NaT", "nan", ""):
        return None
    try:
        return pd.to_datetime(v).date()
    except Exception:
        return None


def _fmt_num(v):
    if pd.isna(v):
        return None
    try:
        return float(v)
    except Exception:
        return None


def _fmt_str_cell(v):
    if pd.isna(v) or str(v).strip().lower() in ("", "nan", "<na>", "nat"):
        return ""
    return str(v).strip()


def _fmt_index_floor_cell(v):
    if pd.isna(v) or str(v).strip().lower() in ("", "nan", "<na>"):
        return None
    try:
        return float(v)
    except (TypeError, ValueError):
        s = str(v).strip()
        return s if s else None


def _fmt_index_name_cell(v):
    if pd.isna(v) or str(v).strip().lower() in ("", "nan", "<na>"):
        return None
    s = str(v).strip()
    return s if s else None


def _fmt_status(v):
    if pd.isna(v) or str(v) in ("nan", "NaN", ""):
        return "N/A"
    return str(v)


def _row_bg(recon_status):
    rs = str(recon_status).upper()
    if "MATCH" in rs and "MIS" not in rs:
        return MATCH_BG
    if "MISMATCH" in rs:
        return MISMATCH_BG
    if "MISSING" in rs:
        return MISSING_BG
    return WHITE


def primary_workbook_context(primary_file_type: str = "ACORE") -> dict:
    cfg = get_primary_config(primary_file_type)
    dn = cfg["display_name"]
    as_of = datetime.now().strftime("%m/%d/%Y")
    return {
        "title": f"Financing Line Reconciliation  —  {dn}  |  As of {as_of}",
        "subtitle": (
            f"Primary Source: {cfg['model_descriptor']}  |  "
            "Comparison Source: M61  |  "
            "Target from M61 file only"
        ),
        "group_primary_header": cfg["primary_group_header"],
        "legend_match_detail": f"All key fields align between {dn} and M61",
        "legend_primary_only_label": cfg["primary_only_legend_label"],
        "legend_primary_only_detail": f"Record found only in {dn} model; not yet in M61",
        "excel_primary_column_suffix": cfg["excel_primary_column_suffix"],
    }


def _excel_header_for_primary(hdr: str, column_suffix: str) -> str:
    if column_suffix == "ACP":
        return hdr
    return hdr.replace("(ACP)", f"({column_suffix})")


def build_workbook(df_recon, primary_file_type: str = "ACORE"):
    wb = Workbook()
    ws = wb.active
    ws.title = "Reconciliation"
    wb_ctx = primary_workbook_context(primary_file_type)
    col_suffix = wb_ctx["excel_primary_column_suffix"]

    total_cols = len(COL_DEFS)
    last_col = get_column_letter(total_cols)

    ws.merge_cells(f"A1:{last_col}1")
    c = ws["A1"]
    c.value = wb_ctx["title"]
    c.font = Font(name="Arial", size=13, bold=True, color=WHITE)
    c.fill = _fill(HEADER_BG)
    c.alignment = _center()
    ws.row_dimensions[1].height = 24

    ws.merge_cells(f"A2:{last_col}2")
    c = ws["A2"]
    c.value = wb_ctx["subtitle"]
    c.font = Font(name="Arial", size=9, italic=True, color=WHITE)
    c.fill = _fill(SUBHDR_BG)
    c.alignment = _center()
    ws.row_dimensions[2].height = 16

    grp_labels = {
        "ID": "IDENTIFICATION",
        "ACP": wb_ctx["group_primary_header"],
        "LIB": "M61 — COMPARISON DATA",
        "STATUS": "RECONCILIATION STATUS",
    }

    runs = []
    run_start = 1
    current_grp = COL_DEFS[0][4]

    for col_idx in range(2, len(COL_DEFS) + 1):
        grp = COL_DEFS[col_idx - 1][4]
        if grp != current_grp:
            runs.append((current_grp, run_start, col_idx - 1))
            run_start = col_idx
            current_grp = grp
    runs.append((current_grp, run_start, len(COL_DEFS)))

    for grp, c_start, c_end in runs:
        if c_start != c_end:
            ws.merge_cells(start_row=3, start_column=c_start, end_row=3, end_column=c_end)

        cell = ws.cell(row=3, column=c_start)
        cell.value = grp_labels[grp]
        cell.font = _hdr_font()
        cell.fill = _fill(GROUP_COLORS[grp])
        cell.alignment = _center()
        cell.border = BORDER

        for cnum in range(c_start, c_end + 1):
            ws.cell(row=3, column=cnum).fill = _fill(GROUP_COLORS[grp])
            ws.cell(row=3, column=cnum).border = BORDER

    ws.row_dimensions[3].height = 35

    for i, (hdr, w, _, _, grp) in enumerate(COL_DEFS, 1):
        cell = ws.cell(row=4, column=i)
        cell.value = _excel_header_for_primary(hdr, col_suffix)
        cell.font = _hdr_font()
        cell.fill = _fill(GROUP_COLORS[grp])
        cell.alignment = _center()
        cell.border = BORDER
        ws.column_dimensions[get_column_letter(i)].width = w

    ws.row_dimensions[4].height = 36

    for data_row_idx, (_, row) in enumerate(df_recon.iterrows(), 5):
        row_bg = _row_bg(row.get("recon_status", ""))

        vals = [
            _fmt_str_cell(row.get("Fund")),
            _fmt_str_cell(row.get("Deal Name", "")),
            _fmt_str_cell(row.get("Facility", "")),
            _fmt_str_cell(row.get("Financial Line", "")),
            _fmt_str_cell(row.get("Note Name", "")),
            _fmt_str_cell(row.get("Liability Note (M61)", "")),
            _fmt_str_cell(row.get("Deal ID (ACP)", "")),
            _fmt_str_cell(row.get("Liability Note Suffix (M61)", "")),
            _fmt_str_cell(row.get("Source", "")),
            _fmt_str_cell(row.get("Source Indicator", "")),
            _fmt_date(row.get("Effective Date (ACP)")),
            _fmt_date(row.get("Effective Date (M61)")),
            _fmt_date(row.get("Pledge Date (ACP)")),
            _fmt_date(row.get("Pledge Date (M61)")),
            _fmt_num(row.get("Advance Rate (ACP)")),
            _fmt_num(row.get("Advance Rate (M61)")),
            _fmt_num(row.get("Target Advance Rate (M61)")),
            _fmt_num(row.get("Spread (ACP)")),
            _fmt_num(row.get("Spread (M61)")),
            _fmt_num(row.get("Undrawn Capacity (ACP)")),
            _fmt_num(row.get("Undrawn Capacity (M61)")),
            _fmt_index_floor_cell(row.get("Index Floor (ACP)")),
            _fmt_index_floor_cell(row.get("Index Floor (M61)")),
            _fmt_index_name_cell(row.get("Index Name (ACP)")),
            _fmt_index_name_cell(row.get("Index Name (M61)")),
            _fmt_num(row.get("Recourse % (ACP)")),
            _fmt_num(row.get("Recourse % (M61)")),
            _fmt_status(row.get("Advance Rate Status")),
            _fmt_status(row.get("Spread Status")),
            _fmt_status(row.get("Effective Date Status")),
            _fmt_status(row.get("Undrawn Capacity Status")),
            _fmt_status(row.get("Index Floor Status")),
            _fmt_status(row.get("Index Name Status")),
            _fmt_status(row.get("Recourse % Status")),
            _fmt_status(row.get("Pledge Date Status")),
            _fmt_status(row.get("recon_status")),
        ]

        for col_idx, (val, (hdr, _w, pct, dt, grp)) in enumerate(zip(vals, COL_DEFS), 1):
            cell = ws.cell(row=data_row_idx, column=col_idx)
            cell.value = val
            cell.border = BORDER

            if grp == "STATUS":
                _status_cell(cell, val)
            else:
                cell.fill = _fill(row_bg)
                cell.font = _body_font()
                if dt and val is not None:
                    cell.number_format = "m/d/yy"
                    cell.alignment = _center()
                elif pct and isinstance(val, float):
                    cell.number_format = "0.00%"
                    cell.alignment = _center()
                elif isinstance(val, str):
                    cell.alignment = _left()
                elif grp in ("ACP", "LIB") and not dt and not pct and isinstance(val, (int, float)):
                    cell.alignment = _center()
                else:
                    cell.alignment = _left()

        ws.row_dimensions[data_row_idx].height = EXCEL_RECON_DATA_ROW_HEIGHT

    ws.freeze_panes = "A5"

    ws_leg = wb.create_sheet("Legend")
    ws_leg.column_dimensions["A"].width = 22
    ws_leg.column_dimensions["B"].width = 50
    ws_leg.cell(1, 1).value = "Legend — Reconciliation Status"
    ws_leg.cell(1, 1).font = Font(name="Arial", size=12, bold=True, color=HEADER_BG)
    ws_leg.cell(1, 1).alignment = _left()

    legends = [
        ("MATCH", MATCH_BG, "375623", wb_ctx["legend_match_detail"]),
        ("MISMATCH", MISMATCH_BG, "9C0006", "Key field values differ between primary model and M61"),
        (
            "MISSING",
            MISSING_BG,
            "7D6608",
            "Record exists in one file only, red-flagged in M61 export, or a key field is "
            "MISSING IN ACORE / MISSING IN AOC II / MISSING IN M61 / BOTH MISSING.",
        ),
        (wb_ctx["legend_primary_only_label"], "D9E1F2", "1F3864", wb_ctx["legend_primary_only_detail"]),
        ("Both", "E2EFDA", "375623", "Record found in primary model and M61 — basis for comparison"),
    ]

    for r, (lbl, bg_hex, fc, desc) in enumerate(legends, 3):
        c1 = ws_leg.cell(r, 1)
        c1.value = lbl
        c1.fill = _fill(bg_hex)
        c1.font = Font(name="Arial", size=9, bold=True, color=fc)
        c1.alignment = _center()
        c1.border = BORDER

        c2 = ws_leg.cell(r, 2)
        c2.value = desc
        c2.font = _body_font()
        c2.alignment = _left()
        ws_leg.row_dimensions[r].height = 18

    # --- Scoped Reconciliation (same layout as main sheet; filtered by Fund for this run) ---
    ws_scoped = wb.create_sheet("Scoped Reconciliation")
    scoped_title = (
        f"{wb_ctx['title']}  —  Scoped: {scope_label_for_primary_type(primary_file_type)}"
    )
    ws_scoped.merge_cells(f"A1:{last_col}1")
    c = ws_scoped["A1"]
    c.value = scoped_title
    c.font = Font(name="Arial", size=13, bold=True, color=WHITE)
    c.fill = _fill(HEADER_BG)
    c.alignment = _center()
    ws_scoped.row_dimensions[1].height = 24

    ws_scoped.merge_cells(f"A2:{last_col}2")
    c = ws_scoped["A2"]
    c.value = wb_ctx["subtitle"]
    c.font = Font(name="Arial", size=9, italic=True, color=WHITE)
    c.fill = _fill(SUBHDR_BG)
    c.alignment = _center()
    ws_scoped.row_dimensions[2].height = 16

    for grp, c_start, c_end in runs:
        if c_start != c_end:
            ws_scoped.merge_cells(start_row=3, start_column=c_start, end_row=3, end_column=c_end)

        cell = ws_scoped.cell(row=3, column=c_start)
        cell.value = grp_labels[grp]
        cell.font = _hdr_font()
        cell.fill = _fill(GROUP_COLORS[grp])
        cell.alignment = _center()
        cell.border = BORDER

        for cnum in range(c_start, c_end + 1):
            ws_scoped.cell(row=3, column=cnum).fill = _fill(GROUP_COLORS[grp])
            ws_scoped.cell(row=3, column=cnum).border = BORDER

    ws_scoped.row_dimensions[3].height = 35

    for i, (hdr, w, _, _, grp) in enumerate(COL_DEFS, 1):
        cell = ws_scoped.cell(row=4, column=i)
        cell.value = _excel_header_for_primary(hdr, col_suffix)
        cell.font = _hdr_font()
        cell.fill = _fill(GROUP_COLORS[grp])
        cell.alignment = _center()
        cell.border = BORDER
        ws_scoped.column_dimensions[get_column_letter(i)].width = w

    ws_scoped.row_dimensions[4].height = 36

    scoped_df = filter_recon_to_selected_fund(df_recon, primary_file_type)
    if scoped_df.empty:
        note = ws_scoped.cell(5, 1)
        note.value = (
            f"No scoped rows found for {scope_label_for_primary_type(primary_file_type)} "
            "(Fund filter did not match any rows)."
        )
        note.font = _body_font()
        note.alignment = _left()
    else:
        for data_row_idx, (_, row) in enumerate(scoped_df.iterrows(), 5):
            row_bg = _row_bg(row.get("recon_status", ""))

            vals = [
                _fmt_str_cell(row.get("Fund")),
                _fmt_str_cell(row.get("Deal Name", "")),
                _fmt_str_cell(row.get("Facility", "")),
                _fmt_str_cell(row.get("Financial Line", "")),
                _fmt_str_cell(row.get("Note Name", "")),
                _fmt_str_cell(row.get("Liability Note (M61)", "")),
                _fmt_str_cell(row.get("Deal ID (ACP)", "")),
                _fmt_str_cell(row.get("Liability Note Suffix (M61)", "")),
                _fmt_str_cell(row.get("Source", "")),
                _fmt_str_cell(row.get("Source Indicator", "")),
                _fmt_date(row.get("Effective Date (ACP)")),
                _fmt_date(row.get("Effective Date (M61)")),
                _fmt_date(row.get("Pledge Date (ACP)")),
                _fmt_date(row.get("Pledge Date (M61)")),
                _fmt_num(row.get("Advance Rate (ACP)")),
                _fmt_num(row.get("Advance Rate (M61)")),
                _fmt_num(row.get("Target Advance Rate (M61)")),
                _fmt_num(row.get("Spread (ACP)")),
                _fmt_num(row.get("Spread (M61)")),
                _fmt_num(row.get("Undrawn Capacity (ACP)")),
                _fmt_num(row.get("Undrawn Capacity (M61)")),
                _fmt_index_floor_cell(row.get("Index Floor (ACP)")),
                _fmt_index_floor_cell(row.get("Index Floor (M61)")),
                _fmt_index_name_cell(row.get("Index Name (ACP)")),
                _fmt_index_name_cell(row.get("Index Name (M61)")),
                _fmt_num(row.get("Recourse % (ACP)")),
                _fmt_num(row.get("Recourse % (M61)")),
                _fmt_status(row.get("Advance Rate Status")),
                _fmt_status(row.get("Spread Status")),
                _fmt_status(row.get("Effective Date Status")),
                _fmt_status(row.get("Undrawn Capacity Status")),
                _fmt_status(row.get("Index Floor Status")),
                _fmt_status(row.get("Index Name Status")),
                _fmt_status(row.get("Recourse % Status")),
                _fmt_status(row.get("Pledge Date Status")),
                _fmt_status(row.get("recon_status")),
            ]

            for col_idx, (val, (hdr, _w, pct, dt, grp)) in enumerate(zip(vals, COL_DEFS), 1):
                cell = ws_scoped.cell(row=data_row_idx, column=col_idx)
                cell.value = val
                cell.border = BORDER

                if grp == "STATUS":
                    _status_cell(cell, val)
                else:
                    cell.fill = _fill(row_bg)
                    cell.font = _body_font()
                    if dt and val is not None:
                        cell.number_format = "m/d/yy"
                        cell.alignment = _center()
                    elif pct and isinstance(val, float):
                        cell.number_format = "0.00%"
                        cell.alignment = _center()
                    elif isinstance(val, str):
                        cell.alignment = _left()
                    elif grp in ("ACP", "LIB") and not dt and not pct and isinstance(val, (int, float)):
                        cell.alignment = _center()
                    else:
                        cell.alignment = _left()

            ws_scoped.row_dimensions[data_row_idx].height = EXCEL_RECON_DATA_ROW_HEIGHT

    ws_scoped.freeze_panes = "A5"
    return wb


def run(
    file_a_path=DEFAULT_FILE_A_PATH,
    file_b_path=DEFAULT_FILE_B_PATH,
    out_path: str | None = None,
    primary_file_type: str = "ACORE",
    mapping_path: str | None = None,
):
    if out_path is None:
        out_path = default_recon_output_path(
            primary_file_type,
            uploaded_filename=os.path.basename(file_b_path),
        )
    print(f"Loading M61 export     : {file_a_path}")
    print(f"Loading primary file   : {file_b_path}  ({primary_file_type})")

    df_recon, df_adv, df_excluded = reconcile(
        file_a_path,
        file_b_path,
        primary_file_type=primary_file_type,
        mapping_path=mapping_path,
        uploaded_primary_filename=os.path.basename(file_b_path),
    )
    df_recon = normalize_recon_fund_for_output(df_recon)

    print("\nRECONCILIATION SUMMARY")
    print("=" * 40)
    for status, count in df_recon["recon_status"].value_counts().items():
        print(f"  {status:<15} {count:>4}")
    print(f"  {'TOTAL':<15} {len(df_recon):>4}")
    print("=" * 40)

    wb = build_workbook(df_recon, primary_file_type=primary_file_type)
    wb.save(out_path)
    print(f"\nSaved → {out_path}")
    print(
        "Excluded M61 rows by Liability Type filter "
        f"(visible in Streamlit diagnostics): {len(df_excluded)}"
    )
    return df_recon, df_adv


def main():
    parser = argparse.ArgumentParser(description="Financing Line Reconciliation — Enhanced Output")
    parser.add_argument("--file-a", dest="file_a", default=DEFAULT_FILE_A_PATH, help="Path to (In) M61 export")
    parser.add_argument("--file-b", dest="file_b", default=DEFAULT_FILE_B_PATH, help="Path to primary business file")
    parser.add_argument(
        "--primary-type",
        default="ACORE",
        choices=tuple(sorted(PRIMARY_FILE_CONFIG.keys())),
        help="Primary workbook template (column mapping)",
    )
    parser.add_argument(
        "--out",
        default=None,
        help="Output Excel path (default: <primary-type> Finance Recon - YYYY-MM-DD.xlsx in script dir)",
    )
    parser.add_argument(
        "--inspect-primary",
        action="store_true",
        help="Print configured primary sheet and column headers for --file-b and --primary-type, then exit",
    )
    parser.add_argument(
        "--mapping",
        default=None,
        help="Path to LiabilityNoteID -> CRENoteID mapping workbook (required for --primary-type AOC II)",
    )
    args = parser.parse_args()
    if args.inspect_primary:
        inspect_primary_workbook(args.file_b, args.primary_type)
        return
    out_path = args.out or default_recon_output_path(
        args.primary_type,
        uploaded_filename=os.path.basename(args.file_b),
    )
    run(
        args.file_a,
        args.file_b,
        out_path,
        primary_file_type=args.primary_type,
        mapping_path=args.mapping,
    )


if __name__ == "__main__":
    main()