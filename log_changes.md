# Change Log — GST Standalone Intelligence Suite

## [1.0.1] - 2026-10-30
### Fixed & Hardened (Statutory QC Audit Resolution)
- **CRITICAL #1: B2BA Amendment Reference Matching**:
  - Fixed B2BA amendment replacement pass to preserve `clean_orig_inv` reference on amended portal records.
  - Reconciler now matches Books entries carrying original invoice numbers (`AM-OLD`) against amended records (`AM-NEW`), eliminating false "Missing in 2B", false "Unclaimed ITC", and false Vendor RED status.
- **CRITICAL #2: Multi-Month Duplicate Numbering & Candidate Scoring**:
  - Replaced single-key dictionary collapse with multi-candidate priority scoring in Pass 1.
  - Matching now prioritizes tax amount match and date/period match before falling back, preventing multi-month files loaded in reverse order from generating false tax mismatches.
- **CRITICAL #3: Table 6.1 IGST Cash Paid Extraction**:
  - Implemented primary table-cell extraction with Section (A) Forward Charge bounding and 3rd-from-right column position (`vals[-3]`).
  - Added robust regex line-scanner fallback, ensuring IGST cash paid (e.g. ₹ 146,523.00) is extracted accurately without defaulting to 0.00.
- **HIGH #4: Official GSTR-2B JSON Ingestion**:
  - Added `parse_gstr2b_json` in `engine.py` to parse portal JSON exports (`b2b`, `b2ba`, `cdnr`, `cdnra`).
  - Wired `.json` file upload directly into `app.py`.
- **HIGH #5: Complete GSTR-1 Tables Extraction**:
  - Extracted Tables 4 (B2B), 5 (B2CL), 7 (B2CS), 8 (Nil/Exempt), 9B (CDNR / CDNUR), and Total Liability via `extract_section_vals`.
- **HIGH #6 & #7: Statutory Watchdogs & UI Wiring**:
  - Connected Rule 88C (DRC-01B) and Rule 88D (DRC-01C) watchdog computation to live return data with `/api/watchdogs` endpoint and interactive status cards.
- **HIGH #8: B2B-CDNRA Ingestion**:
  - Added support for amended credit/debit notes sheet `B2B-CDNRA`.
- **MEDIUM #9: Unclaimed Portal ITC Tab**:
  - Added dedicated `Unclaimed Portal ITC` (`missing_in_books`) tab in openpyxl Excel report and interactive view in web dashboard.
- **MEDIUM #11: Proportionate Vendor Risk Scoring**:
  - Weighted vendor risk scoring (60% missing ratio, 40% missing ITC amount), ensuring small isolated missing bills do not unfairly trigger RED "HOLD PAYMENT".
- **MEDIUM #12 & #13: Error Handling & Resource Hygiene**:
  - Added structured `try...catch` in API routes and frontend JS; wrapped temporary file parsing in `try...finally` to eliminate file leaks.
- **MEDIUM #14 & #15: XSS Sanitization & Thread Safety**:
  - Sanitized table rendering using HTML entity escaping (`esc()`); added `STATE_LOCK` thread safety; avoided implicit global `event` in `showTab`.
- **LOW #16 & #19: Negatives & Synonyms Safety**:
  - Extended `clean_num` to support bracketed negative values e.g. `(1,200.00)` -> `-1200.00`.
  - Removed `"customer gstin"` from purchase synonyms to prevent sales export pollution.
  - Cleaned up build artifacts, maintaining only `dist/GST_Audit_Suite.exe`.

## [1.0.0] - 2026-10-30
### Added
- **Core Engine (`engine.py`)**:
  - GSTR-2B official multi-sheet Excel ingestion (`B2B`, `B2BA`, `B2B-CDNR`, `B2B-CDNRA`, `IMPG`, `IMPGSEZ`).
  - B2BA amendment replacement pipeline.
  - Credit Note `-1.0` signing convention.
  - Multi-month batch file appending & de-duplication.
  - Dynamic Books Excel column auto-mapping.
  - 4-Pass matching reconciler (Exact, Fuzzy Reference, Gaps).
  - 0–100 Vendor Risk Scoring & RAG classification with ITC-at-Risk ₹ under Sec 16(2)(aa).
  - GSTR-3B PDF Data Extractor (Table 3.1, 4, 5, 5.1, 6.1, ARN).
  - GSTR-1 / 1A PDF Data Extractor (Table 4, 5, 7, 8, 9B, Liability, ARN).
  - Rule 88C DRC-01B and Rule 88D DRC-01C Exposure Watchdogs.
  - 6-Tab CFO/CA-ready formula-first openpyxl Excel report builder.
- **Application Server (`app.py`)**:
  - Single-page responsive dashboard running locally on `http://127.0.0.1:5055`.
  - Auto-launches default browser.
- **Portable Distribution**:
  - `GST_Audit_Suite.exe` (111 MB standalone portable binary).
  - `Run_GST_Suite.bat` (batch runner).
- **Documentation**:
  - Institutional `README.md`, `README_HOW_TO_USE.txt`, `walkthrough.md`, and `learning.md`.
