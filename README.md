ACORE vs. M61 Data Reconciliation Tool
For Finance & Operations Users

LINK: https://hvrecontool.streamlit.app/ 
________________


Overview
This tool compares two data sources — ACORE (your primary model file) and M61 (the comparison export) — to identify where records match, where they differ, and where data may be missing from one side or the other.


The goal is to surface meaningful discrepancies quickly, without requiring you to manually cross-reference spreadsheets. The tool handles the matching logic automatically and presents results in a clear, reviewable format.


________________


How It Works
Step 1 — Upload Your Files
Upload your ACORE file and your M61 export. Once both are loaded, the tool processes every row in each file and attempts to find corresponding records across the two sources.
Step 2 — Select Your Fund Scope
Choose whether you want to view:


* All Uploaded File Results — shows everything from both files, across all funds present in each
* Selected Primary Fund Only — filters to a specific ACORE fund (e.g., ACP II), showing only records relevant to that fund


Because the M61 file often contains data for multiple funds, scope selection ensures you are only comparing what is relevant to your chosen fund.
Step 3 — Review the Results
The main results table shows one row per matched or unmatched record, with a clear status for each. Clicking into any row opens a drilldown view showing the underlying raw data from both files side by side.


________________


How Records Are Matched
Records are paired using a combination of deal identifiers and effective date. Specifically, a row from ACORE is matched to a row in M61 when the deal name, facility or liability name, liability note, and effective date all align.
Exact Match
When all identifiers — including the effective date — match, the records are paired and compared field by field.
Fallback Pairing (Date Mismatch)
If the same deal exists in both files but the effective dates differ, the records are still paired and shown side by side. This allows you to see the discrepancy rather than treating both rows as entirely missing. These will appear as mismatches, with the date difference visible in the drilldown.
No Match Found
If a record exists in ACORE but cannot be paired with anything in M61, it appears as ACORE Only. If a record exists in M61 but has no corresponding ACORE row, it appears as M61 Only.


Important: Even if you recognize a deal on both sides, the tool will not force a match if the effective dates are different and no fallback pair can be determined. In those cases, both records appear as single-sided entries until the underlying data is corrected.


________________


Understanding Results
Reconciliation Status (Row Level)
Each record in the results table carries one of the following overall statuses:


Status
	What It Means
	MATCH
	All compared fields align between ACORE and M61
	MISMATCH
	At least one field differs between the two sources
	MISSING
	The record exists on only one side (ACORE Only or M61 Only)
	PARTIAL REVIEW
	Mixed results where fields partially align but a direct mismatch is not confirmed — warrants manual review
	Field-Level Status (Drilldown)
Within the drilldown for any record, each individual field also carries its own status:


Status
	What It Means
	MATCH
	The field value is the same in both ACORE and M61
	MISMATCH
	The field values differ between the two sources
	MISSING FROM M61
	The field has a value in ACORE but is blank or absent in M61
	MISSING FROM ACORE
	The field has a value in M61 but is blank or absent in ACORE
	File Source
Each row also indicates where its data comes from:


* ACORE Only — record found only in the ACORE file
* M61 Only — record found only in the M61 file
* Both — record was paired across both files and compared


________________


Using the UI
Main Results Table
The main table gives you a high-level summary — one row per record, showing the recon status, deal identifiers, and key field values. Use the filters at the top to narrow results by status, fund, or other criteria.
Drilldown View
Click any row to open the drilldown, which shows the raw underlying data from both ACORE and M61 for that record. The drilldown is the best place to understand why a record is flagged — it surfaces date differences, value discrepancies, and missing fields in detail.
Filters
Additional filters allow you to focus on specific subsets of the data, such as:


* Showing only mismatches
* Showing only records missing from one side
* Filtering to a specific deal or fund


Note: Filters operate on the full dataset for your selected scope — they do not compound in unexpected ways based on what is currently displayed.
Exporting Results
You can export the current view to CSV or Excel. The exported file reflects your active scope and filter selections, so what you see is what you get.


________________


Supported Funds
The tool supports the following ACORE funds:


* ACP I
* ACP II
* ACP III
* AOC I
* AOC II


When you select a fund, the tool filters both ACORE and M61 data to records tied to that fund. Some field definitions or behaviors may vary slightly across funds — if you notice unexpected results, verify that the correct fund is selected before investigating further.


________________


Common Scenarios & FAQs
Q: A deal appears in both files, but the tool shows it as ACORE Only or M61 Only. Why?


The most common reason is an effective date mismatch. The tool requires the effective date to align (or be close enough to trigger fallback pairing) before it will compare the two records. Check the effective dates on both sides in your source files.


Q: I see a MISMATCH status, but the values look the same to me. What's happening?


Open the drilldown and review each field's individual status. Common culprits include trailing spaces, formatting differences in percentages or dates, or one side containing a null where the other has a zero. The field-level status will point you to the exact discrepancy.


Q: Some records I expect to see aren't showing up at all. Why?


The tool filters out records that are out of scope — for example, non-financing rows or rows not tied to the selected fund. If you're in "Selected Primary Fund Only" mode, switching to "All Uploaded File Results" may reveal records that were excluded from the scoped view.


Q: The M61 file has many more rows than ACORE. Is that a problem?


Not necessarily. M61 exports often contain data for multiple funds. Once you select your fund scope, the tool narrows the M61 data to what is relevant for comparison. Any remaining M61-only rows after matching may indicate records in M61 that ACORE has not yet captured.


Q: What does PARTIAL REVIEW mean, and what should I do with it?


PARTIAL REVIEW flags records where the results are mixed — some fields match, others are unclear — but there is no confirmed direct mismatch. Treat these as items worth a manual look. Open the drilldown, review the field-level statuses, and confirm whether the values are acceptable.


Q: Can I use this tool for multiple funds at once?


Yes. Select "All Uploaded File Results" to see data across all funds present in your uploaded files. Keep in mind that the comparison logic will apply to each fund's records based on what is in both files, so results from different funds will all appear together in the main table.


________________




For questions about the tool or to report unexpected behavior, contact your data or operations team.
