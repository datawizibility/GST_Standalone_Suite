# ⚡ GST Standalone Intelligence Suite

> **Offline, Zero-Installation GST Reconciliation, Vendor Risk Scoring & Return Extraction Platform**  
> Designed for Chartered Accountants, CFOs, Tax Professionals, and Enterprise Accounting Teams.

[![Python](https://img.shields.io/badge/Python-3.10%2B-blue.svg)](https://www.python.org/)
[![License: MIT](https://img.shields.io/badge/License-MIT-green.svg)](LICENSE)
[![Platform](https://img.shields.io/badge/Platform-Windows%20%7C%20Linux%20%7C%20macOS-lightgrey.svg)](#)
[![Compliance](https://img.shields.io/badge/GST-Rule%2088C%20%7C%2088D%20%7C%20Sec%2016(2)(aa)-orange.svg)](#)

---

## 🎯 Executive Overview

GST compliance in India faces three major operational friction points:
1. **Vendor Default Risk (Section 16(2)(aa))**: When suppliers fail to upload invoices in their GSTR-1, the buyer loses input tax credit (ITC), creating unexpected working capital traps.
2. **Amendment Chaos (Table B2BA)**: Suppliers constantly amend typing errors or invoice dates in later returns. Traditional matching scripts treat amended rows as both a "missing" and an "excess" invoice, generating hundreds of false alerts.
3. **DRC-01B & DRC-01C Tax Intimations**: Discrepancies exceeding ₹25 Lakhs or 20% between GSTR-1 liability and GSTR-3B tax paid, or between GSTR-2B ITC and GSTR-3B claim, trigger automated department intimations with immediate recovery threats.

**GST Standalone Suite** solves this in a single, lightweight, 100% offline tool. It requires **no installation, no Python environment setup, and zero cloud logins**.

---

## 🚀 Key Modules

```
┌────────────────────────────────────────────────────────────────────────┐
│                   GST STANDALONE INTELLIGENCE SUITE                    │
└───────────────────────────────────┬────────────────────────────────────┘
                                    │
    ┌───────────────────────────────┼───────────────────────────────┐
    ▼                               ▼                               ▼
┌─────────────────────────┐   ┌───────────────────────────┐   ┌──────────────────────────┐
│ Module 1: Purchase 2B   │   │ Module 2 & 3: Extractor   │   │ Module 4: Watchdogs      │
│ • 4-Pass Invoice Match  │   │ • GSTR-3B PDF to Excel    │   │ • Rule 88C DRC-01B Watch │
│ • B2BA Amendment Match  │   │ • GSTR-1 / 1A PDF to Excel│   │ • Rule 88D DRC-01C Watch │
│ • Multi-Month 2B Append │   │ • Month-over-Month Tables │   │ • Sec 16(2)(aa) Exposure │
│ • 0-100 Vendor Risk RAG │   │ • ARN & Filing Extraction │   │ • ITC-at-Risk Quantifier │
└─────────────────────────┘   └───────────────────────────┘   └──────────────────────────┘
```

### 1. Purchase Reconciliation (GSTR-2B vs ERP Books)
- **Official Multi-Tab Ingestion**: Natively reads GST Portal workbooks across sheets: `B2B`, `B2BA`, `B2B-CDNR`, `B2B-CDNRA`, `IMPG`, and `IMPGSEZ`.
- **B2BA Amendment Replacement Engine**: Automatically resolves amendments filed in Table `B2BA` by replacing the original record and flagging `Amended in B2BA (Orig: X)`. Eliminates duplicate and false discrepancy flags.
- **Credit Note Signing**: Automatically adjusts Credit Notes as negative (`-1.0`) so that net ITC balances match ERP Trial Balances mathematically.
- **Multi-Month Batch Append**: Upload 1 month or 12 months at once — the engine deduplicates and merges records chronologically.
- **Dynamic Books Column Auto-Discovery**: Upload purchase registers from Tally, SAP, Busy, Marg, Zoho Books, or custom Excel. Intelligently maps variations of headers (`Voucher No`, `Bill No`, `Party GSTIN`, `Taxable Value`, `Total Amount`).

### 2. Vendor Credit Risk Engine (0–100 RAG Matrix)
- Analyzes each supplier's historical compliance behavior.
- Computes:
  - **Total Booked ITC vs. Portal ITC**
  - **Unfiled Invoice Ratio & Default Rate**
  - **₹ ITC at Risk** (Cash directly blocked under Section 16(2)(aa))
  - **RAG Status**:
    - 🔴 **RED (Score ≥ 50 or ITC Risk > ₹50,000)**: `HOLD PAYMENT` — Immediate risk of tax reversal + Sec 50 interest.
    - 🟡 **AMBER (Score ≥ 20 or ITC Risk > ₹5,000)**: `SEND REMINDER` — Pending GSTR-1 filing.
    - 🟢 **GREEN**: `COMPLIANT` — Release payment as scheduled.

### 3. GSTR-3B PDF Data Extractor (MOM Consolidated)
- Automatically extracts filed return copies (scans a single PDF or entire folder batch).
- Parses:
  - **Table 3.1**: Outward taxable, zero-rated, nil/exempt, non-GST, and inward RCM supplies.
  - **Table 4**: Eligible ITC (All other ITC, Import of goods/services, ISD, RCM) and Net ITC available.
  - **Table 6.1**: Tax payment breakdown (Cash paid vs ITC utilized).
  - **Filing Details**: ARN, ARN Date, and Filing Status.
- Generates a multi-month consolidated Month-over-Month (MOM) Excel summary.

### 4. GSTR-1 / 1A PDF Data Extractor (MOM Consolidated)
- Parses filed GSTR-1, GSTR-1A, and IFF PDFs.
- Extracts Table 4 (B2B), Table 5 (B2CL), Table 7 (B2CS), Table 8 (Nil/Exempt), Table 9B (CDNR / CDNUR), and Total Outward Liability.

### 5. Statutory Watchdogs (Rule 88C & Rule 88D)
- **Rule 88C Watchdog**: Evaluates Outward Tax Liability reported in GSTR-1 against GSTR-3B. Flags exposure if difference exceeds ₹25 Lakhs or 20% (DRC-01B risk).
- **Rule 88D Watchdog**: Evaluates ITC claimed in GSTR-3B against ITC available in GSTR-2B. Flags exposure if excess credit exceeds ₹25 Lakhs or 20% (DRC-01C risk).

---

## 💻 How to Run

### Method 1: Portable Executable (Zero Installation)
1. Download `GST_Audit_Suite.exe` from the [Latest Release](https://github.com/datawizibility/GST_Standalone_Suite/releases).
2. Double-click **`GST_Audit_Suite.exe`**.
3. A console window will start and your default browser will open automatically at:
   ```
   http://127.0.0.1:5055
   ```
4. To stop the application, simply close the console window.

### Method 2: Python / Developer Setup
```bash
# Clone the repository
git clone https://github.com/datawizibility/GST_Standalone_Suite.git
cd GST_Standalone_Suite

# Install dependencies
pip install -r requirements.txt

# Run via batch script (Windows)
Run_GST_Suite.bat

# Or run directly via Python
python app.py
```

---

## 📊 CA/CFO-Ready Excel Reports

Every generated workbook strictly conforms to professional accounting standards:
- **Font**: Segoe UI
- **Number Format**: `#,##0.00`
- **Formula-First Totals**: `=SUM(...)` formulas for dynamic recalculation in Excel.
- **Multi-Tab Organization**:
  1. `Executive Summary`: High-level metrics & statutory compliance KPI dashboard.
  2. `Matched Invoices`: Exact and tolerance-matched records with audit match type tag.
  3. `Tax Discrepancies`: Invoices with rate, taxable, or tax value variance between Books and Portal.
  4. `Missing in GSTR-2B`: Invoices missing on portal, tagged with Section 16(2)(aa) notice risk.
  5. `Unclaimed Portal ITC`: Missing in Books (immediate unclaimed cash saving opportunity).
  6. `Supplier Risk Matrix`: Vendor-wise default ranking and recommended commercial actions.

---

## 🔒 Security & Data Privacy

- **100% Offline Execution**: All computations, regex parsers, and Excel generators execute locally in memory on your PC.
- **Zero Cloud Transmission**: No client data, GSTINs, invoice numbers, or financial figures are sent over the internet.
- **DPDP Act 2023 Compliant**: Client tax data remains strictly confined to your local workstation.

---

## 📄 License
This project is open-source under the [MIT License](LICENSE).
