# Change Log — GST Standalone Intelligence Suite

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
