# Chat History — GST Standalone Intelligence Suite

## 2026-10-30

**User**:
1. Check packaged tool in `D:\AAAA_MYPROJECT\GST Reco Tools` and compare against `gst-suite` and `tally_App`.
2. Review GSTR-3B and GSTR-1 scripts in `D:\AAAA_MYPROJECT\3BData`.
3. Create standalone zero-installation suite (`GST_Standalone_Suite`) with portable executable and push to GitHub.
4. QC Audit Report:
   - CRITICAL: B2BA amendments produce false Missing in 2B + false vendor RED; No period in matching keys; IGST cash paid silently = 0.00.
   - HIGH: GSTR-2B JSON upload; GSTR-1 tables 5/7/8/9B; 3B tables 5/5.1; Rule 88C/88D watchdog tab; B2B-CDNRA sheet.
   - MEDIUM: missing_in_books in Excel/UI; first-candidate-wins dupe collapse; vendor risk 50% ratio threshold; error handling & temp files; XSS sanitization; showTab global event.
   - LOW: clean_num bracketed negatives; unused imports; duplicate root exe; customer gstin synonym.
   Want me to fix the three critical items? Pls push the updated to git pls.

**Assistant**:
- Resolved all 19 QC items across `engine.py`, `app.py`, and `Run_GST_Suite.bat`.
- Fixed B2BA amendment matching by retaining original invoice reference (`clean_orig_inv`) on replacement records.
- Implemented multi-candidate priority scoring by tax amount and date/period affinity.
- Rebuilt Table 6.1 cash paid extraction with Section (A) Forward Charge bounding and `vals[-3]` table-cell locator.
- Implemented official GSTR-2B JSON parser (`parse_gstr2b_json`) supporting B2B, B2BA, CDNR, CDNRA.
- Added GSTR-1 tables (4, 5, 7, 8, 9B, Total Liability) and 3B tables (3.1, 4, 5, 5.1, 6.1).
- Connected live Rule 88C and Rule 88D statutory watchdog calculations to `/api/watchdogs`.
- Added Unclaimed Portal ITC (`missing_in_books`) worksheet in Excel and interactive table in web cockpit.
- Weighted vendor risk scoring (60% missing ratio, 40% missing amount) eliminating false RED status on small isolated invoices.
- Handled bracketed negative numbers `(1,200.00)` -> `-1200.00` and removed customer gstin from purchase synonyms.
- Created regression test suite `tests/test_qc_regression.py` (9/9 tests passing).
- Compiled fresh 111 MB standalone binary `dist/GST_Audit_Suite.exe` and published GitHub Release `v1.0.1`.
- Committed and pushed updates to `main` branch across both repositories.
