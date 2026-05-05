# ACORE vs. M61 Data Reconciliation Tool  
### For Finance & Operations Users

LINK: https://hvrecontool.streamlit.app/ 

---

## Overview

This tool compares two data sources — **ACORE (your primary model file)** and **M61 (the comparison export)** — to identify:

- Where records match  
- Where they differ  
- Where data may be missing  

The goal is to surface discrepancies quickly without manually cross-referencing spreadsheets.

---

## How It Works

### Step 1 — Upload Your Files
Upload your ACORE file and your M61 export.  
The tool processes every row and attempts to match records across both sources.

---

### Step 2 — Select Your Fund Scope

Choose whether you want to view:

- **All Uploaded File Results**  
  → Shows everything across all funds  

- **Selected Primary Fund Only**  
  → Filters to a specific ACORE fund (e.g., ACP II)

Because the M61 file often contains multiple funds, scope selection ensures you are only comparing relevant data.

---

### Step 3 — Review the Results

- Main table → one row per record with status  
- Drilldown → detailed side-by-side comparison  

---

## How Records Are Matched

Records are matched using:
- Deal identifiers  
- Effective date  

---

### Exact Match
All identifiers (including effective date) align → records are paired and compared.

---

### Fallback Pairing (Date Mismatch)
If the same deal exists but dates differ:
- Rows are still paired  
- Shown side by side  
- Marked as **MISMATCH**

---

### No Match Found

- ACORE only → **ACORE Only**  
- M61 only → **M61 Only**

---

⚠️ Even if a deal exists on both sides, it will not match if dates differ and no fallback pairing is possible.

---

## Understanding Results

### Recon Status (Row Level)

| Status | Meaning |
|------|--------|
| MATCH | All fields align |
| MISMATCH | At least one field differs |
| MISSING | Exists on only one side |
| PARTIAL REVIEW | Mixed or unclear results |

---

### Field-Level Status (Drilldown)

| Status | Meaning |
|------|--------|
| MATCH | Values are the same |
| MISMATCH | Values differ |
| MISSING FROM M61 | ACORE has value, M61 does not |
| MISSING FROM ACORE | M61 has value, ACORE does not |

---

### File Source

- **ACORE Only** — found only in ACORE  
- **M61 Only** — found only in M61  
- **Both** — paired and compared  

---

## Using the UI

### Main Results Table
- High-level summary  
- One row per record  
- Shows key statuses  

---

### Drilldown View
Click a row to:
- See full ACORE vs M61 data  
- Understand mismatches  

---

### Filters

You can filter by:
- Match / Mismatch / Missing  
- Fund  
- Deal  

---

### Exporting Results

- Export to CSV or Excel  
- Reflects current filters and scope  

---

## Supported Funds

- ACP I  
- ACP II  
- ACP III  
- AOC I  
- AOC II  

Selecting a fund:
- Filters both ACORE and M61  
- Adjusts matching scope  

---

## Common Scenarios & FAQs

### Why does a deal show as ACORE Only or M61 Only?
Most likely:
- Effective date mismatch  
- Or filtered out by scope  

---

### Why does MISMATCH appear when values look the same?
Check drilldown:
- Formatting differences  
- Null vs 0  
- Hidden spaces  

---

### Why don’t I see certain rows?
They may be:
- Out of scope  
- Not tied to selected fund  

---

### Is it normal for M61 to have more rows than ACORE?
Yes — M61 often contains multiple funds.

---

### What does PARTIAL REVIEW mean?
Mixed results — review manually in drilldown.

---

### Can I view multiple funds?
Yes — use **All Uploaded File Results**


