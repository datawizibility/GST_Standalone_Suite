# Architectural Walkthrough — GST Standalone Intelligence Suite

## 1. System Overview
The GST Standalone Intelligence Suite is an offline, zero-installation reconciliation and compliance platform. It runs as a self-contained local application and launches a browser interface at `http://127.0.0.1:5055`.

## 2. Core Modules & Execution Pipeline
```
[User Ingestion]
   ├── Books Register (Excel: Tally / Busy / SAP / Zoho)
   ├── GSTR-2B Workbooks (Multi-Sheet Excel or JSON)
   ├── GSTR-3B Return Copies (PDFs)
   └── GSTR-1 / 1A Return Copies (PDFs)
          │
          ▼
   [engine.py Engine Pipeline]
   ├── B2BA Amendment Overrides & CDNR Sign Adjustment
   ├── 4-Pass Matching Engine (Exact, Fuzzy Reference, Gaps)
   ├── 0-100 Vendor Risk Evaluator & RAG Matrix
   ├── PDF Regex Extraction (3.1, 4, 5, 5.1, 6.1, ARN)
   └── Statutory Exposure Watchdogs (Rule 88C / 88D)
          │
          ▼
   [Outputs]
   ├── Responsive Interactive Dashboard (:5055)
   ├── 6-Tab CFO/CA-Ready Reconciliation Workbook (.xlsx)
   └── Month-over-Month Consolidated 3B/1 Excel (.xlsx)
```

## 3. Data Safety & Network Isolation
- All file buffers and calculations reside strictly in memory.
- Zero outbound network calls or API dependencies.
- No client financial records are persisted after report generation unless explicitly exported.
