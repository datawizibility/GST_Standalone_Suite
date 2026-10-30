# -*- coding: utf-8 -*-
"""
GST Intelligence & Reconciliation Engine (Standalone Suite)
Zero-dependency, production-grade GST statutory engine combining:
1. GSTR-2B vs ERP Purchase Reconciler (Multi-Candidate, B2BA overrides, Multi-Month, Vendor Risk RAG)
2. GSTR-2B Excel & JSON Importer (B2B, B2BA, B2B-CDNR, B2B-CDNRA, IMPG, IMPGSEZ)
3. GSTR-3B PDF Data Extractor (Table 3.1, 3.2, 4, 5, 5.1, 6.1, ARN)
4. GSTR-1 / 1A PDF Data Extractor (Table 4, 5, 7, 8, 9B, Total Liability, ARN)
5. Rule 88C (DRC-01B) and Rule 88D (DRC-01C) Exposure Watchdogs
6. CA/CFO-Ready Excel Workbook Generation (openpyxl Formula-First)
"""

import os
import re
import json
import logging
from datetime import datetime
import pandas as pd
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

logger = logging.getLogger(__name__)

# PDF libraries
try:
    import pdfplumber
    HAS_PDFPLUMBER = True
except ImportError:
    HAS_PDFPLUMBER = False

try:
    import pypdf
    HAS_PYPDF = True
except ImportError:
    HAS_PYPDF = False

FONT_FAMILY = "Segoe UI"
COLOR_HEADER_BG = "1F4E79"
COLOR_HEADER_TEXT = "FFFFFF"
COLOR_ACCENT = "2F5597"
COLOR_MATCHED_BG = "E2EFDA"
COLOR_MATCHED_TEXT = "375623"
COLOR_MISMATCH_BG = "FFF2CC"
COLOR_MISMATCH_TEXT = "7F6000"
COLOR_RISK_BG = "FCE4D6"
COLOR_RISK_TEXT = "C65911"
COLOR_UNCLAIMED_BG = "DDEBF7"
COLOR_UNCLAIMED_TEXT = "1F4E79"
COLOR_BORDER = "D9D9D9"

MONTH_MAP = {
    "jan": 1, "feb": 2, "mar": 3, "apr": 4, "may": 5, "jun": 6,
    "jul": 7, "aug": 8, "sep": 9, "oct": 10, "nov": 11, "dec": 12
}

# Purchase column synonyms - excludes customer gstin to prevent sales export pollution
BOOKS_PURCHASE_COLUMN_SYNONYMS = {
    "gstin": ["supplier gstin", "party gstin", "gstin/uin", "gstin", "gstin no", "vendor gstin"],
    "party_name": ["party name", "supplier name", "vendor name", "ledger name", "particulars", "name"],
    "invoice_number": ["invoice no", "invoice number", "bill no", "voucher no", "inv no", "reference", "ref no"],
    "invoice_date": ["invoice date", "date", "bill date", "voucher date", "inv date"],
    "taxable_value": ["taxable value", "taxable amount", "assessable value", "taxable", "base amount"],
    "igst": ["igst", "igst amount", "integrated tax"],
    "cgst": ["cgst", "cgst amount", "central tax"],
    "sgst": ["sgst", "sgst amount", "state tax", "sgst/utgst"],
    "cess": ["cess", "cess amount"],
    "total_amount": ["total amount", "invoice value", "invoice amount", "total", "gross total"]
}


def clean_num(val):
    """
    Cleans numeric strings into float.
    Handles None, empty, dashes, commas, rupee symbols, and bracketed negative numbers like (1,200.00).
    """
    if val is None or pd.isna(val) or str(val).strip() in ("-", "", ".", "None"):
        return 0.0
    if isinstance(val, (int, float)):
        return round(float(val), 2)
    s = str(val).strip().replace(",", "").replace("₹", "")
    if s.startswith("(") and s.endswith(")"):
        s = "-" + s[1:-1].strip()
    s = re.sub(r"[^\d\.-]", "", s)
    try:
        return round(float(s), 2)
    except Exception:
        return 0.0


def normalize_inv(inv_str):
    """Normalizes invoice numbers by stripping whitespace and non-alphanumeric characters."""
    if not inv_str:
        return ""
    return re.sub(r"[^A-Z0-9]", "", str(inv_str).strip().upper())


def parse_date(date_val):
    """Parses various date representations into YYYY-MM-DD string."""
    if not date_val:
        return ""
    if isinstance(date_val, (datetime, pd.Timestamp)):
        return date_val.strftime("%Y-%m-%d")
    s = str(date_val).strip()
    for fmt in ("%d-%m-%Y", "%Y-%m-%d", "%d/%m/%Y", "%Y%m%d", "%d-%b-%Y", "%d/%m/%y", "%d-%m-%y"):
        try:
            return datetime.strptime(s, fmt).strftime("%Y-%m-%d")
        except ValueError:
            pass
    # If string is already 10 chars YYYY-MM-DD or DD-MM-YYYY
    if re.match(r"^\d{4}-\d{2}-\d{2}$", s):
        return s
    return s[:10]


def extract_period_from_date(dt_str):
    """Extracts YYYY-MM period from YYYY-MM-DD date."""
    if dt_str and len(dt_str) >= 7 and dt_str[4] == "-":
        return dt_str[:7]
    return ""


# ═════════════════════════════════════════════════════════════════════════════
# 1. GSTR-2B INGESTION ENGINE (Excel & JSON)
# ═════════════════════════════════════════════════════════════════════════════

def parse_gstr2b_excel(file_path):
    """
    Parses official GSTR-2B portal Excel.
    Reads sheets: B2B, B2BA, B2B-CDNR, B2B-CDNRA, IMPG, IMPGSEZ.
    Performs B2BA and B2B-CDNRA replacement without dropping the original invoice reference.
    """
    wb = openpyxl.load_workbook(file_path, data_only=True)
    invoices = []
    amendments = []

    # 1. B2B
    if "B2B" in wb.sheetnames:
        ws = wb["B2B"]
        header_row = 5
        for r in range(1, min(10, ws.max_row + 1)):
            vals = [str(ws.cell(r, c).value or "") for c in range(1, 10)]
            if any("GSTIN of supplier" in v for v in vals):
                header_row = r
                break
        for r in range(header_row + 2, ws.max_row + 1):
            gstin = str(ws.cell(r, 1).value or "").strip().upper()
            if not gstin or len(gstin) < 10:
                continue
            name = str(ws.cell(r, 2).value or "").strip()
            inv_no = str(ws.cell(r, 3).value or "").strip()
            inv_dt = parse_date(ws.cell(r, 5).value)
            inv_val = clean_num(ws.cell(r, 6).value)
            taxable = clean_num(ws.cell(r, 10).value)
            igst = clean_num(ws.cell(r, 11).value)
            cgst = clean_num(ws.cell(r, 12).value)
            sgst = clean_num(ws.cell(r, 13).value)
            cess = clean_num(ws.cell(r, 14).value)
            invoices.append({
                "source": "GSTR-2B", "doc_type": "INV", "supplier_gstin": gstin,
                "supplier_name": name, "invoice_number": inv_no, "clean_inv_num": normalize_inv(inv_no),
                "clean_orig_inv": normalize_inv(inv_no), "orig_inv_num": inv_no,
                "invoice_date": inv_dt, "period": extract_period_from_date(inv_dt),
                "invoice_value": inv_val, "taxable_value": taxable,
                "igst": igst, "cgst": cgst, "sgst": sgst, "cess": cess,
                "total_itc": round(igst + cgst + sgst + cess, 2), "amendment_note": ""
            })

    # 2. B2BA (Amendments to B2B Invoices)
    if "B2BA" in wb.sheetnames:
        ws = wb["B2BA"]
        if ws.max_row > 7:
            for r in range(8, ws.max_row + 1):
                orig_inv = str(ws.cell(r, 1).value or "").strip()
                gstin = str(ws.cell(r, 3).value or "").strip().upper()
                name = str(ws.cell(r, 4).value or "").strip()
                rev_inv = str(ws.cell(r, 5).value or "").strip() or orig_inv
                rev_dt = parse_date(ws.cell(r, 7).value)
                rev_val = clean_num(ws.cell(r, 8).value)
                rev_taxable = clean_num(ws.cell(r, 12).value)
                igst = clean_num(ws.cell(r, 13).value)
                cgst = clean_num(ws.cell(r, 14).value)
                sgst = clean_num(ws.cell(r, 15).value)
                cess = clean_num(ws.cell(r, 16).value)
                if gstin and (orig_inv or rev_inv):
                    amendments.append({
                        "source": "GSTR-2B-B2BA", "doc_type": "INV",
                        "orig_inv_num": orig_inv, "clean_orig_inv": normalize_inv(orig_inv),
                        "supplier_gstin": gstin, "supplier_name": name,
                        "invoice_number": rev_inv, "clean_inv_num": normalize_inv(rev_inv),
                        "invoice_date": rev_dt, "period": extract_period_from_date(rev_dt),
                        "invoice_value": rev_val, "taxable_value": rev_taxable,
                        "igst": igst, "cgst": cgst, "sgst": sgst, "cess": cess,
                        "total_itc": round(igst + cgst + sgst + cess, 2),
                        "amendment_note": f"Amended in B2BA (Orig: {orig_inv})" if orig_inv != rev_inv else "Amended in B2BA"
                    })

    # 3. B2B-CDNR (Credit & Debit Notes)
    if "B2B-CDNR" in wb.sheetnames:
        ws = wb["B2B-CDNR"]
        if ws.max_row > 6:
            for r in range(7, ws.max_row + 1):
                gstin = str(ws.cell(r, 1).value or "").strip().upper()
                if not gstin or len(gstin) < 10:
                    continue
                name = str(ws.cell(r, 2).value or "").strip()
                nt_num = str(ws.cell(r, 3).value or "").strip()
                nt_type = str(ws.cell(r, 4).value or "Credit Note").strip().lower()
                sign = -1.0 if "credit" in nt_type else 1.0
                nt_dt = parse_date(ws.cell(r, 6).value)
                val = clean_num(ws.cell(r, 7).value) * sign
                taxable = clean_num(ws.cell(r, 11).value) * sign
                igst = clean_num(ws.cell(r, 12).value) * sign
                cgst = clean_num(ws.cell(r, 13).value) * sign
                sgst = clean_num(ws.cell(r, 14).value) * sign
                cess = clean_num(ws.cell(r, 15).value) * sign
                invoices.append({
                    "source": "GSTR-2B-CDNR", "doc_type": "CR" if sign < 0 else "DR",
                    "supplier_gstin": gstin, "supplier_name": name,
                    "invoice_number": nt_num, "clean_inv_num": normalize_inv(nt_num),
                    "clean_orig_inv": normalize_inv(nt_num), "orig_inv_num": nt_num,
                    "invoice_date": nt_dt, "period": extract_period_from_date(nt_dt),
                    "invoice_value": round(val, 2), "taxable_value": round(taxable, 2),
                    "igst": round(igst, 2), "cgst": round(cgst, 2), "sgst": round(sgst, 2), "cess": round(cess, 2),
                    "total_itc": round(igst + cgst + sgst + cess, 2), "amendment_note": ""
                })

    # 4. B2B-CDNRA (Amended Credit & Debit Notes)
    if "B2B-CDNRA" in wb.sheetnames:
        ws = wb["B2B-CDNRA"]
        if ws.max_row > 7:
            for r in range(8, ws.max_row + 1):
                orig_nt = str(ws.cell(r, 1).value or "").strip()
                gstin = str(ws.cell(r, 3).value or "").strip().upper()
                name = str(ws.cell(r, 4).value or "").strip()
                rev_nt = str(ws.cell(r, 5).value or "").strip() or orig_nt
                nt_type = str(ws.cell(r, 6).value or "Credit Note").strip().lower()
                sign = -1.0 if "credit" in nt_type else 1.0
                rev_dt = parse_date(ws.cell(r, 8).value)
                val = clean_num(ws.cell(r, 9).value) * sign
                taxable = clean_num(ws.cell(r, 13).value) * sign
                igst = clean_num(ws.cell(r, 14).value) * sign
                cgst = clean_num(ws.cell(r, 15).value) * sign
                sgst = clean_num(ws.cell(r, 16).value) * sign
                cess = clean_num(ws.cell(r, 17).value) * sign
                if gstin and (orig_nt or rev_nt):
                    amendments.append({
                        "source": "GSTR-2B-CDNRA", "doc_type": "CR" if sign < 0 else "DR",
                        "orig_inv_num": orig_nt, "clean_orig_inv": normalize_inv(orig_nt),
                        "supplier_gstin": gstin, "supplier_name": name,
                        "invoice_number": rev_nt, "clean_inv_num": normalize_inv(rev_nt),
                        "invoice_date": rev_dt, "period": extract_period_from_date(rev_dt),
                        "invoice_value": round(val, 2), "taxable_value": round(taxable, 2),
                        "igst": round(igst, 2), "cgst": round(cgst, 2), "sgst": round(sgst, 2), "cess": round(cess, 2),
                        "total_itc": round(igst + cgst + sgst + cess, 2),
                        "amendment_note": f"Amended in CDNRA (Orig: {orig_nt})" if orig_nt != rev_nt else "Amended in CDNRA"
                    })

    # Apply statutory amendments replacement
    # We DO NOT delete original numbers from existence. Instead, if the original entry exists in B2B/CDNR,
    # we replace its monetary payload with the amended B2BA record, while preserving clean_orig_inv.
    # If the original entry is from a prior un-ingested period, we add the amendment record as is.
    result_invoices = list(invoices)
    for amend in amendments:
        matched_idx = None
        for idx, inv in enumerate(result_invoices):
            if inv["supplier_gstin"] == amend["supplier_gstin"] and inv["clean_inv_num"] == amend["clean_orig_inv"]:
                matched_idx = idx
                break
        if matched_idx is not None:
            # Replace original invoice with amended, retaining both original and revised identifiers
            result_invoices[matched_idx] = amend
        else:
            result_invoices.append(amend)

    return result_invoices


def parse_gstr2b_json(file_path):
    """
    Parses official GSTR-2B JSON file exported from the GST portal.
    Extracts B2B, B2BA, CDNR, and CDNRA tables.
    """
    with open(file_path, "r", encoding="utf-8") as f:
        data = json.load(f)

    # Portal JSON may have payload wrapped in 'data' or directly at root
    payload = data.get("data", data)
    invoices = []
    amendments = []

    # 1. B2B
    for supplier in payload.get("b2b", []):
        gstin = supplier.get("ctin", "").strip().upper()
        name = supplier.get("trdNm", "").strip() or supplier.get("lglNm", "").strip()
        for inv in supplier.get("inv", []):
            inv_no = inv.get("inum", "").strip()
            inv_dt = parse_date(inv.get("idt"))
            inv_val = clean_num(inv.get("val"))
            taxable = 0.0
            igst = 0.0
            cgst = 0.0
            sgst = 0.0
            cess = 0.0
            for itm in inv.get("items", []):
                taxable += clean_num(itm.get("txval"))
                igst += clean_num(itm.get("igst"))
                cgst += clean_num(itm.get("cgst"))
                sgst += clean_num(itm.get("sgst"))
                cess += clean_num(itm.get("cess"))
            invoices.append({
                "source": "GSTR-2B-JSON", "doc_type": "INV", "supplier_gstin": gstin,
                "supplier_name": name, "invoice_number": inv_no, "clean_inv_num": normalize_inv(inv_no),
                "clean_orig_inv": normalize_inv(inv_no), "orig_inv_num": inv_no,
                "invoice_date": inv_dt, "period": extract_period_from_date(inv_dt),
                "invoice_value": inv_val, "taxable_value": round(taxable, 2),
                "igst": round(igst, 2), "cgst": round(cgst, 2), "sgst": round(sgst, 2), "cess": round(cess, 2),
                "total_itc": round(igst + cgst + sgst + cess, 2), "amendment_note": ""
            })

    # 2. B2BA
    for supplier in payload.get("b2ba", []):
        gstin = supplier.get("ctin", "").strip().upper()
        name = supplier.get("trdNm", "").strip() or supplier.get("lglNm", "").strip()
        for inv in supplier.get("inv", []):
            orig_inv = inv.get("oinum", "").strip()
            rev_inv = inv.get("inum", "").strip() or orig_inv
            rev_dt = parse_date(inv.get("idt"))
            rev_val = clean_num(inv.get("val"))
            taxable = 0.0
            igst = 0.0
            cgst = 0.0
            sgst = 0.0
            cess = 0.0
            for itm in inv.get("items", []):
                taxable += clean_num(itm.get("txval"))
                igst += clean_num(itm.get("igst"))
                cgst += clean_num(itm.get("cgst"))
                sgst += clean_num(itm.get("sgst"))
                cess += clean_num(itm.get("cess"))
            amendments.append({
                "source": "GSTR-2B-B2BA-JSON", "doc_type": "INV",
                "orig_inv_num": orig_inv, "clean_orig_inv": normalize_inv(orig_inv),
                "supplier_gstin": gstin, "supplier_name": name,
                "invoice_number": rev_inv, "clean_inv_num": normalize_inv(rev_inv),
                "invoice_date": rev_dt, "period": extract_period_from_date(rev_dt),
                "invoice_value": rev_val, "taxable_value": round(taxable, 2),
                "igst": round(igst, 2), "cgst": round(cgst, 2), "sgst": round(sgst, 2), "cess": round(cess, 2),
                "total_itc": round(igst + cgst + sgst + cess, 2),
                "amendment_note": f"Amended in B2BA (Orig: {orig_inv})" if orig_inv != rev_inv else "Amended in B2BA"
            })

    # 3. CDNR
    for supplier in payload.get("cdnr", []):
        gstin = supplier.get("ctin", "").strip().upper()
        name = supplier.get("trdNm", "").strip()
        for nt in supplier.get("nt", []):
            nt_num = nt.get("ntnum", "").strip()
            nt_type = str(nt.get("ntty", "C")).strip().upper()
            sign = -1.0 if nt_type == "C" else 1.0
            nt_dt = parse_date(nt.get("nt_dt"))
            val = clean_num(nt.get("val")) * sign
            taxable = 0.0
            igst = 0.0
            cgst = 0.0
            sgst = 0.0
            cess = 0.0
            for itm in nt.get("items", []):
                taxable += clean_num(itm.get("txval")) * sign
                igst += clean_num(itm.get("igst")) * sign
                cgst += clean_num(itm.get("cgst")) * sign
                sgst += clean_num(itm.get("sgst")) * sign
                cess += clean_num(itm.get("cess")) * sign
            invoices.append({
                "source": "GSTR-2B-CDNR-JSON", "doc_type": "CR" if sign < 0 else "DR",
                "supplier_gstin": gstin, "supplier_name": name,
                "invoice_number": nt_num, "clean_inv_num": normalize_inv(nt_num),
                "clean_orig_inv": normalize_inv(nt_num), "orig_inv_num": nt_num,
                "invoice_date": nt_dt, "period": extract_period_from_date(nt_dt),
                "invoice_value": round(val, 2), "taxable_value": round(taxable, 2),
                "igst": round(igst, 2), "cgst": round(cgst, 2), "sgst": round(sgst, 2), "cess": round(cess, 2),
                "total_itc": round(igst + cgst + sgst + cess, 2), "amendment_note": ""
            })

    # Apply amendments replacement
    result_invoices = list(invoices)
    for amend in amendments:
        matched_idx = None
        for idx, inv in enumerate(result_invoices):
            if inv["supplier_gstin"] == amend["supplier_gstin"] and inv["clean_inv_num"] == amend["clean_orig_inv"]:
                matched_idx = idx
                break
        if matched_idx is not None:
            result_invoices[matched_idx] = amend
        else:
            result_invoices.append(amend)

    return result_invoices


def parse_books_excel(file_path):
    """
    Ingests Books / ERP purchase register.
    Uses dynamic column synonym resolution and ignores customer gstin to prevent sales pollution.
    """
    wb = openpyxl.load_workbook(file_path, data_only=True)
    ws = wb.active
    rows = list(ws.iter_rows(values_only=True))
    if not rows:
        return []

    header_idx = 0
    headers = []
    for idx, row in enumerate(rows[:15]):
        row_str = [str(c or "").lower().strip() for c in row]
        if any("gstin" in c or "invoice" in c or "voucher" in c for c in row_str):
            header_idx = idx
            headers = row_str
            break

    col_map = {}
    for standard_col, syns in BOOKS_PURCHASE_COLUMN_SYNONYMS.items():
        for syn in syns:
            for c_idx, h in enumerate(headers):
                norm_h = re.sub(r"[^a-z0-9]", "", h)
                norm_syn = re.sub(r"[^a-z0-9]", "", syn)
                if norm_syn in norm_h or norm_h in norm_syn:
                    col_map[standard_col] = c_idx
                    break
            if standard_col in col_map:
                break

    records = []
    for row in rows[header_idx + 1:]:
        if not any(row):
            continue
        gstin = str(row[col_map["gstin"]] or "").strip().upper() if "gstin" in col_map else ""
        inv_no = str(row[col_map["invoice_number"]] or "").strip() if "invoice_number" in col_map else ""
        if not gstin and not inv_no:
            continue
        name = str(row[col_map["party_name"]] or "").strip() if "party_name" in col_map else ""
        inv_dt = parse_date(row[col_map["invoice_date"]]) if "invoice_date" in col_map else ""
        taxable = clean_num(row[col_map["taxable_value"]]) if "taxable_value" in col_map else 0.0
        igst = clean_num(row[col_map["igst"]]) if "igst" in col_map else 0.0
        cgst = clean_num(row[col_map["cgst"]]) if "cgst" in col_map else 0.0
        sgst = clean_num(row[col_map["sgst"]]) if "sgst" in col_map else 0.0
        cess = clean_num(row[col_map["cess"]]) if "cess" in col_map else 0.0
        tot = clean_num(row[col_map["total_amount"]]) if "total_amount" in col_map else (taxable + igst + cgst + sgst + cess)

        records.append({
            "supplier_gstin": gstin, "supplier_name": name, "party_name": name,
            "invoice_number": inv_no, "clean_inv_num": normalize_inv(inv_no),
            "invoice_date": inv_dt, "period": extract_period_from_date(inv_dt),
            "invoice_value": tot, "taxable_value": taxable,
            "igst": igst, "cgst": cgst, "sgst": sgst, "cess": cess,
            "total_itc": round(igst + cgst + sgst + cess, 2)
        })
    return records


# ═════════════════════════════════════════════════════════════════════════════
# 2. MULTI-CANDIDATE RECONCILER & VENDOR RISK MATRIX
# ═════════════════════════════════════════════════════════════════════════════

def reconcile_purchase_2b(books, portal, tolerance=2.0):
    """
    Statutory 4-Pass Purchase Reconciler.
    - Resolves B2BA original/revised invoice references
    - Uses candidate scoring with period/date matching to prevent multi-month mismatch false alarms
    - Calculates Unclaimed ITC (Missing in Books)
    - Computes proportionate Vendor Risk (0-100 score & RAG)
    """
    matched = []
    mismatch = []
    missing_2b = []
    missing_books = []
    used_portal = set()
    used_books = set()

    # Pass 1: Exact / Candidate-Scored Match
    for b_idx, b in enumerate(books):
        b_gstin = b.get("supplier_gstin")
        b_clean = b.get("clean_inv_num")
        if not b_gstin or not b_clean:
            continue

        # Find all eligible candidates in portal
        candidates = []
        for p_idx, p in enumerate(portal):
            if p_idx in used_portal:
                continue
            if p.get("supplier_gstin") != b_gstin:
                continue

            # Direct invoice match OR B2BA amendment match against original invoice number
            is_inv_match = (b_clean == p.get("clean_inv_num"))
            is_amend_match = (p.get("clean_orig_inv") and b_clean == p.get("clean_orig_inv"))

            if is_inv_match or is_amend_match:
                diff_tx = abs(b["taxable_value"] - p["taxable_value"])
                diff_itc = abs(b["total_itc"] - p["total_itc"])
                is_tax_match = (diff_tx <= tolerance and diff_itc <= tolerance)

                # Date / Period match check
                b_date = b.get("invoice_date", "")
                p_date = p.get("invoice_date", "")
                b_period = b.get("period", "")
                p_period = p.get("period", "")
                is_date_match = (b_date and p_date and b_date == p_date)
                is_period_match = (b_period and p_period and b_period == p_period)

                # Calculate candidate score
                score = 0
                if is_tax_match:
                    score += 100
                if is_date_match:
                    score += 50
                elif is_period_match:
                    score += 30
                if is_inv_match:
                    score += 10

                candidates.append((score, is_tax_match, diff_tx, diff_itc, is_amend_match, p_idx, p))

        if candidates:
            # Sort candidates by:
            # 1. Tax match (True first)
            # 2. Overall score (Highest first)
            # 3. Smallest tax difference
            candidates.sort(key=lambda c: (1 if c[1] else 0, c[0], -c[3]), reverse=True)
            best_cand = candidates[0]
            _, is_tax_match, diff_tx, diff_itc, is_amend_match, p_idx, p = best_cand

            match_label = "Exact Match"
            if is_amend_match:
                match_label = f"Exact (Amended in B2BA, Books Ref: {b['invoice_number']}, Rev: {p['invoice_number']})"
            elif p.get("amendment_note"):
                match_label = f"Exact ({p['amendment_note']})"

            if is_tax_match:
                matched.append({
                    "match_type": match_label,
                    "gstin": b["supplier_gstin"], "supplier_name": b.get("supplier_name") or p.get("supplier_name"),
                    "books_inv": b["invoice_number"], "portal_inv": p["invoice_number"],
                    "books_date": b["invoice_date"], "portal_date": p["invoice_date"],
                    "books_taxable": b["taxable_value"], "portal_taxable": p["taxable_value"],
                    "taxable_diff": round(b["taxable_value"] - p["taxable_value"], 2),
                    "books_itc": b["total_itc"], "portal_itc": p["total_itc"],
                    "itc_diff": round(b["total_itc"] - p["total_itc"], 2), "status": "MATCHED"
                })
                used_portal.add(p_idx)
                used_books.add(b_idx)
            else:
                mismatch.append({
                    "match_type": "Tax / Value Difference",
                    "gstin": b["supplier_gstin"], "supplier_name": b.get("supplier_name") or p.get("supplier_name"),
                    "books_inv": b["invoice_number"], "portal_inv": p["invoice_number"],
                    "books_date": b["invoice_date"], "portal_date": p["invoice_date"],
                    "books_taxable": b["taxable_value"], "portal_taxable": p["taxable_value"],
                    "taxable_diff": round(b["taxable_value"] - p["taxable_value"], 2),
                    "books_itc": b["total_itc"], "portal_itc": p["total_itc"],
                    "itc_diff": round(b["total_itc"] - p["total_itc"], 2), "status": "MISMATCH"
                })
                used_portal.add(p_idx)
                used_books.add(b_idx)

    # Pass 2: Probable / Substring Match (Remaining unmatched books)
    for b_idx, b in enumerate(books):
        if b_idx in used_books:
            continue
        for p_idx, p in enumerate(portal):
            if p_idx in used_portal:
                continue
            if b.get("supplier_gstin") == p.get("supplier_gstin"):
                b_num = b["clean_inv_num"].lstrip("0")
                p_num = p["clean_inv_num"].lstrip("0")
                if b_num and p_num and (b_num == p_num or b_num in p_num or p_num in b_num):
                    diff_tx = abs(b["taxable_value"] - p["taxable_value"])
                    diff_itc = abs(b["total_itc"] - p["total_itc"])
                    if diff_tx <= tolerance and diff_itc <= tolerance:
                        matched.append({
                            "match_type": "Probable Reference Match",
                            "gstin": b["supplier_gstin"], "supplier_name": b.get("supplier_name") or p.get("supplier_name"),
                            "books_inv": b["invoice_number"], "portal_inv": p["invoice_number"],
                            "books_date": b["invoice_date"], "portal_date": p["invoice_date"],
                            "books_taxable": b["taxable_value"], "portal_taxable": p["taxable_value"],
                            "taxable_diff": round(b["taxable_value"] - p["taxable_value"], 2),
                            "books_itc": b["total_itc"], "portal_itc": p["total_itc"],
                            "itc_diff": round(b["total_itc"] - p["total_itc"], 2), "status": "MATCHED"
                        })
                        used_portal.add(p_idx)
                        used_books.add(b_idx)
                        break

    # Pass 3: Missing in 2B (Blocked ITC Sec 16(2)(aa))
    for b_idx, b in enumerate(books):
        if b_idx not in used_books:
            missing_2b.append({
                "gstin": b["supplier_gstin"], "supplier_name": b.get("supplier_name", ""),
                "books_inv": b["invoice_number"], "books_date": b["invoice_date"],
                "books_taxable": b["taxable_value"], "books_itc": b["total_itc"],
                "status": "NOT_IN_2B", "action": "Hold payment; supplier unfiled GSTR-1"
            })

    # Pass 4: Missing in Books (Unclaimed ITC)
    for p_idx, p in enumerate(portal):
        if p_idx not in used_portal:
            missing_books.append({
                "gstin": p["supplier_gstin"], "supplier_name": p.get("supplier_name", ""),
                "portal_inv": p["invoice_number"], "portal_date": p["invoice_date"],
                "portal_taxable": p["taxable_value"], "portal_itc": p["total_itc"],
                "status": "NOT_IN_BOOKS", "action": "Verify with purchase & claim in GSTR-3B"
            })

    # Vendor Risk Scoring Engine (Materiality-weighted 0-100 score & RAG)
    vendors = {}
    for b in books:
        gstin = b.get("supplier_gstin")
        if not gstin:
            continue
        if gstin not in vendors:
            vendors[gstin] = {
                "gstin": gstin, "name": b.get("supplier_name", ""), "invoices": 0,
                "total_itc": 0.0, "missing_itc": 0.0, "missing_count": 0, "score": 0, "rag": "GREEN"
            }
        vendors[gstin]["invoices"] += 1
        vendors[gstin]["total_itc"] += b.get("total_itc", 0.0)

    for m in missing_2b:
        gstin = m["gstin"]
        if gstin in vendors:
            vendors[gstin]["missing_count"] += 1
            vendors[gstin]["missing_itc"] += m["books_itc"]

    for gstin, v in vendors.items():
        v["total_itc"] = round(v["total_itc"], 2)
        v["missing_itc"] = round(v["missing_itc"], 2)
        ratio = v["missing_count"] / max(1, v["invoices"])
        # Score calculation: 60% based on missing ratio, 40% based on missing ITC amount
        amount_factor = min(1.0, v["missing_itc"] / 50000.0)
        score = int((ratio * 60) + (amount_factor * 40))
        v["score"] = score

        # Statutory RAG Classification
        if v["missing_itc"] > 50000 or (score >= 60 and v["missing_itc"] > 10000):
            v["rag"] = "RED"
            v["action"] = "HOLD PAYMENT - High ITC Default Risk (Sec 16(2)(aa))"
        elif v["missing_itc"] > 5000 or (score >= 25 and v["missing_itc"] > 2000):
            v["rag"] = "AMBER"
            v["action"] = "SEND REMINDER - Supplier Return Pending"
        else:
            v["rag"] = "GREEN"
            v["action"] = "COMPLIANT - Regular Payment"

    vendor_list = sorted(vendors.values(), key=lambda x: (x["missing_itc"], x["score"]), reverse=True)

    tot_books_itc = round(sum(b["total_itc"] for b in books), 2)
    tot_portal_itc = round(sum(p["total_itc"] for p in portal), 2)
    tot_matched_itc = round(sum(m["portal_itc"] for m in matched), 2)
    tot_risk_itc = round(sum(m["books_itc"] for m in missing_2b), 2)
    tot_unclaimed_itc = round(sum(p["portal_itc"] for p in missing_books), 2)

    summary = {
        "timestamp": datetime.now().isoformat(),
        "total_books_records": len(books), "total_portal_records": len(portal),
        "matched_count": len(matched), "tax_mismatch_count": len(mismatch),
        "missing_in_2b_count": len(missing_2b), "missing_in_books_count": len(missing_books),
        "total_books_itc": tot_books_itc, "total_portal_itc": tot_portal_itc,
        "matched_itc": tot_matched_itc, "itc_at_risk": tot_risk_itc, "unclaimed_itc": tot_unclaimed_itc,
        "high_risk_vendors": len([v for v in vendor_list if v["rag"] == "RED"])
    }

    return {
        "summary": summary, "matched": matched, "tax_mismatch": mismatch,
        "missing_in_2b": missing_2b, "missing_in_books": missing_books, "vendor_risks": vendor_list
    }


# ═════════════════════════════════════════════════════════════════════════════
# 3. GSTR-3B & GSTR-1 PDF EXTRACTORS
# ═════════════════════════════════════════════════════════════════════════════

def extract_pdf_text_and_tables(pdf_path):
    """Extracts text and structured tables from a PDF."""
    text = ""
    tables = []
    if HAS_PDFPLUMBER:
        try:
            with pdfplumber.open(pdf_path) as pdf:
                for p in pdf.pages:
                    text += (p.extract_text() or "") + "\n"
                    page_tables = p.extract_tables()
                    if page_tables:
                        tables.extend(page_tables)
                if text.strip():
                    return text, tables
        except Exception as e:
            logger.warning(f"pdfplumber extraction failed for {pdf_path}: {e}")

    if HAS_PYPDF:
        try:
            reader = pypdf.PdfReader(pdf_path)
            text = "\n".join((p.extract_text() or "") for p in reader.pages)
        except Exception as e:
            logger.warning(f"pypdf extraction failed for {pdf_path}: {e}")
    return text, tables


def extract_gstr3b_pdf(pdf_path):
    """
    Extracts complete GSTR-3B filed data:
    - Metadata: GSTIN, Legal Name, Financial Year, Period, ARN, ARN Date
    - Table 3.1: (a) Outward Taxable, (b) Zero Rated, (c) Nil/Exempt, (d) RCM, (e) Non-GST
    - Table 3.2: Interstate supplies (Unregistered, Composition, UIN)
    - Table 4: Eligible ITC (Available All, Net Available)
    - Table 5: Inward supplies (Composition/Exempt, Non-GST)
    - Table 5.1: Interest and Late Fee
    - Table 6.1: Payment of Tax (Cash Paid for IGST, CGST, SGST, Cess)
    """
    text, tables = extract_pdf_text_and_tables(pdf_path)
    data = {
        "filename": os.path.basename(pdf_path), "gstin": "N/A", "legal_name": "N/A",
        "year": "N/A", "period": "N/A", "arn": "N/A", "arn_date": "N/A",
        "t31": {
            "outward_taxable": [0.0]*5, "zero_rated": [0.0]*4, "nil_exempt": [0.0],
            "rcm": [0.0]*5, "non_gst": [0.0]
        },
        "t32": {"unreg": [0.0]*2, "comp": [0.0]*2, "uin": [0.0]*2},
        "t4": {"net_itc": [0.0]*4, "avail_all": [0.0]*4},
        "t5": {"comp_exempt": [0.0]*2, "non_gst": [0.0]*2},
        "t51": {"interest": [0.0]*4, "late_fee": [0.0]*4},
        "t61": {"cash_paid": [0.0]*4}
    }
    if not text:
        return data

    m_g = re.search(r"GSTIN.*?([0-9]{2}[A-Z]{5}[0-9]{4}[A-Z]{1}[1-9A-Z]{1}Z[0-9A-Z]{1})", text)
    if m_g:
        data["gstin"] = m_g.group(1).upper()

    m_n = re.search(r"Legal name.*?\s+(.*)", text)
    if m_n:
        data["legal_name"] = m_n.group(1).split("\n")[0].strip()

    m_y = re.search(r"(?:Financial year|Year)\s*[:\-]?\s*(\d{4}\s*-\s*\d{2,4})", text, re.IGNORECASE)
    if m_y:
        data["year"] = re.sub(r"\s+", "", m_y.group(1))

    m_p = re.search(r"(?:Tax period|Period)\s*[:\-]?\s*([A-Za-z-]+)", text, re.IGNORECASE)
    if m_p:
        data["period"] = m_p.group(1).strip()

    m_a = re.search(r"ARN\s*[:\-]?\s*([A-Z0-9]{15,22})", text)
    if m_a:
        data["arn"] = m_a.group(1).strip()

    m_ad = re.search(r"(?:ARN Date|Date)\s*[:\-]?\s*(\d{2}/\d{2}/\d{4})", text, re.IGNORECASE)
    if m_ad:
        data["arn_date"] = m_ad.group(1).strip()

    # Table 3.1
    m_31a = re.search(r"\(a\)[^0-9]+([\d,.-]+)\s+([\d,.-]+)\s+([\d,.-]+)\s+([\d,.-]+)\s+([\d,.-]+)", text)
    if m_31a:
        data["t31"]["outward_taxable"] = [clean_num(v) for v in m_31a.groups()]

    m_31d = re.search(r"\(d\)[^0-9]+([\d,.-]+)\s+([\d,.-]+)\s+([\d,.-]+)\s+([\d,.-]+)\s+([\d,.-]+)", text)
    if m_31d:
        data["t31"]["rcm"] = [clean_num(v) for v in m_31d.groups()]

    # Table 4 Net ITC
    m_4c = re.search(r"\(C\)\s*Net\s+ITC\s+Available[^0-9]+([\d,.-]+)\s+([\d,.-]+)\s+([\d,.-]+)\s+([\d,.-]+)", text, re.IGNORECASE) or \
           re.search(r"C\.\s+Net\s+ITC\s+available[^0-9]+([\d,.-]+)\s+([\d,.-]+)\s+([\d,.-]+)\s+([\d,.-]+)", text, re.IGNORECASE) or \
           re.search(r"\(5\)\s*All\s+other\s+ITC[^0-9]+([\d,.-]+)\s+([\d,.-]+)\s+([\d,.-]+)\s+([\d,.-]+)", text, re.IGNORECASE)
    if m_4c:
        data["t4"]["net_itc"] = [clean_num(v) for v in m_4c.groups()]

    # Table 5
    m_5a = re.search(r"(?:composition scheme|Exempt, Nil)[^\d]+([\d,.-]+)\s+([\d,.-]+)", text, re.IGNORECASE)
    if m_5a:
        data["t5"]["comp_exempt"] = [clean_num(v) for v in m_5a.groups()]
    m_5b = re.search(r"Non\s+GST\s+supply[^\d]+([\d,.-]+)\s+([\d,.-]+)", text, re.IGNORECASE)
    if m_5b:
        data["t5"]["non_gst"] = [clean_num(v) for v in m_5b.groups()]

    # Table 6.1 Payment of Tax (Cash Paid)
    cash_extracted = False

    # 1. Primary: Extract from table objects if available
    for tbl in tables:
        if not tbl or len(tbl) < 4:
            continue
        hdr = " ".join(str(c or "").lower() for c in tbl[0])
        if "tax paid through itc" in hdr or "adjustment" in hdr or "net tax" in hdr:
            in_sec_a = False
            for row in tbl:
                c0 = str(row[0] or "").strip()
                if "(A)" in c0 or "Other than reverse charge" in c0:
                    in_sec_a = True
                    continue
                if "(B)" in c0 or "Reverse charge" in c0:
                    in_sec_a = False
                    break
                if in_sec_a and len(row) >= 6:
                    c0_lower = c0.lower().replace("\n", " ")
                    val = clean_num(row[-3])  # Cash paid is always 3rd column from right
                    if "integrated" in c0_lower:
                        data["t61"]["cash_paid"][0] = val
                    elif "central" in c0_lower:
                        data["t61"]["cash_paid"][1] = val
                    elif "state" in c0_lower or "ut" in c0_lower:
                        data["t61"]["cash_paid"][2] = val
                    elif "cess" in c0_lower:
                        data["t61"]["cash_paid"][3] = val
            cash_extracted = any(v > 0 for v in data["t61"]["cash_paid"])
            if cash_extracted:
                break

    # 2. Robust Text Fallback: Section (A) line scanner
    if not cash_extracted:
        clean_text = re.sub(r"(\d+\.\d+)\s*\n\s*(\d+)", r"\1\2", text)
        a_sec = re.search(r"\(A\)\s*Other than reverse charge(.*?)(?:\(B\)|Breakup|\Z)", clean_text, re.DOTALL | re.IGNORECASE)
        if a_sec:
            sec = a_sec.group(1)
            heads = [(0, r"Integrated"), (1, r"Central"), (2, r"State/UT|State"), (3, r"Cess")]
            for head_idx, head_pat in heads:
                m = re.search(rf"({head_pat})\s*(?:tax)?\s+([^\n\r]+)", sec, re.IGNORECASE)
                if m:
                    row_content = m.group(2).replace("-", " 0.0 ")
                    nums = [clean_num(x) for x in row_content.split() if x.strip()]
                    if len(nums) >= 3:
                        data["t61"]["cash_paid"][head_idx] = nums[-3]

    return data


def extract_section_vals(text, header_regex, num_vals=5):
    """Helper to extract numeric row values under a specific GSTR-1 section header."""
    header_match = re.search(header_regex, text, re.IGNORECASE | re.DOTALL)
    if header_match:
        search_text = text[header_match.end():header_match.end() + 2500]
        lines = search_text.split("\n")
        num_pat = r"[\d\.,-]+"
        for line in lines:
            if "Total" in line:
                found_nums = re.findall(num_pat, line)
                filtered_nums = [n for n in found_nums if "." in n]
                if len(filtered_nums) >= num_vals:
                    return [clean_num(v) for v in filtered_nums[:num_vals]]
    return [0.0] * num_vals


def extract_gstr1_pdf(pdf_path):
    """
    Extracts complete GSTR-1 / 1A filed data:
    - Metadata: GSTIN, Legal Name, Financial Year, Period, Return Type, ARN, ARN Date
    - Table 4: B2B Regular supplies
    - Table 5: B2CL (Large) supplies
    - Table 7: B2CS (Small / Others) supplies
    - Table 8: Nil, Exempted, and Non-GST supplies
    - Table 9B: CDNR (Credit/Debit Notes Registered) and CDNUR (Unregistered)
    - Total Outward Tax Liability
    """
    text, _ = extract_pdf_text_and_tables(pdf_path)
    data = {
        "filename": os.path.basename(pdf_path), "gstin": "N/A", "legal_name": "N/A",
        "year": "N/A", "period": "N/A", "arn": "N/A", "arn_date": "N/A", "type": "GSTR-1",
        "b2b": [0.0]*5, "b2cl": [0.0]*5, "b2cs": [0.0]*5,
        "cdnr": [0.0]*5, "cdnur": [0.0]*5, "nil": [0.0]*3,
        "total_liability": [0.0]*5
    }
    if not text:
        return data

    if "GSTR-1A" in text:
        data["type"] = "GSTR-1A"
    elif "IFF" in text:
        data["type"] = "GSTR-1 IFF"

    m_g = re.search(r"GSTIN.*?([0-9]{2}[A-Z]{5}[0-9]{4}[A-Z]{1}[1-9A-Z]{1}Z[0-9A-Z]{1})", text)
    if m_g:
        data["gstin"] = m_g.group(1).upper()

    m_n = re.search(r"Legal name.*?\s+(.*)", text)
    if m_n:
        data["legal_name"] = m_n.group(1).split("\n")[0].strip()

    m_y = re.search(r"(?:Financial year|Year)\s*[:\-]?\s*(\d{4}\s*-\s*\d{2,4})", text, re.IGNORECASE)
    if m_y:
        data["year"] = re.sub(r"\s+", "", m_y.group(1))

    m_p = re.search(r"(?:Tax period|Period)\s*[:\-]?\s*([A-Za-z-]+)", text, re.IGNORECASE)
    if m_p:
        data["period"] = m_p.group(1).strip()

    m_a = re.search(r"ARN\s*[:\-]?\s*([A-Z0-9]+)", text)
    if m_a:
        data["arn"] = m_a.group(1).strip()

    m_ad = re.search(r"(?:ARN Date|Date)\s*[:\-]?\s*(\d{2}/\d{2}/\d{4})", text, re.IGNORECASE)
    if m_ad:
        data["arn_date"] = m_ad.group(1).strip()

    # Extract standard GSTR-1 tables
    data["b2b"] = extract_section_vals(text, r"4A, 4B.*?B2B|B2B Regular", 5)
    data["b2cl"] = extract_section_vals(text, r"5 -.*?B2CL|B2C \(Large\)", 5)
    data["b2cs"] = extract_section_vals(text, r"7-.*?B2CS|7 -.*?B2CS|B2C \(Others\)", 5)
    data["cdnr"] = extract_section_vals(text, r"9B.*?CDNR", 5)
    data["cdnur"] = extract_section_vals(text, r"9B.*?CDNUR", 5)
    data["nil"] = extract_section_vals(text, r"8 -.*?Nil|8A, 8B.*?Nil", 3)

    # Total Outward Liability
    tl_m = re.search(r"Total Liability.*?outward supplies.*?\s+([\d,.-]+)\s+([\d,.-]+)\s+([\d,.-]+)\s+([\d,.-]+)\s+([\d,.-]+)", text, re.IGNORECASE) or \
           re.search(r"Total Liability.*?\s+([\d,.-]+)\s+([\d,.-]+)\s+([\d,.-]+)\s+([\d,.-]+)\s+([\d,.-]+)", text, re.IGNORECASE)
    if tl_m:
        data["total_liability"] = [clean_num(v) for v in tl_m.groups()]

    return data


# ═════════════════════════════════════════════════════════════════════════════
# 4. STATUTORY WATCHDOG ENGINE (Rule 88C / 88D)
# ═════════════════════════════════════════════════════════════════════════════

def compute_statutory_watchdogs(gstr1_liability, gstr3b_liability, gstr3b_itc, gstr2b_itc):
    """
    Computes DRC-01B (Rule 88C) and DRC-01C (Rule 88D) statutory exposure.
    Threshold: Difference > Rs. 25,00,000 AND > 20% of the baseline.
    """
    # 1. Rule 88C: GSTR-1 vs GSTR-3B Tax Liability
    gap_88c = round(gstr1_liability - gstr3b_liability, 2)
    pct_88c = round((gap_88c / max(1.0, gstr3b_liability)) * 100, 2) if gstr3b_liability > 0 else (100.0 if gap_88c > 0 else 0.0)
    trigger_88c = (gap_88c > 2500000.0) and (pct_88c > 20.0)

    res_88c = {
        "rule": "Rule 88C (Form DRC-01B)",
        "gstr1_liability": round(gstr1_liability, 2),
        "gstr3b_liability": round(gstr3b_liability, 2),
        "gap": gap_88c,
        "gap_percentage": pct_88c,
        "status": "NON_COMPLIANT_TRIGGERED" if trigger_88c else "COMPLIANT",
        "action": "ACTION REQUIRED: File Form DRC-01B reply or deposit differential tax with interest" if trigger_88c else "Within statutory tolerance (No DRC-01B)"
    }

    # 2. Rule 88D: GSTR-3B vs GSTR-2B Input Tax Credit
    gap_88d = round(gstr3b_itc - gstr2b_itc, 2)
    pct_88d = round((gap_88d / max(1.0, gstr2b_itc)) * 100, 2) if gstr2b_itc > 0 else (100.0 if gap_88d > 0 else 0.0)
    trigger_88d = (gap_88d > 2500000.0) and (pct_88d > 20.0)

    res_88d = {
        "rule": "Rule 88D (Form DRC-01C)",
        "gstr3b_itc": round(gstr3b_itc, 2),
        "gstr2b_itc": round(gstr2b_itc, 2),
        "gap": gap_88d,
        "gap_percentage": pct_88d,
        "status": "NON_COMPLIANT_TRIGGERED" if trigger_88d else "COMPLIANT",
        "action": "ACTION REQUIRED: File Form DRC-01C reply or reverse excess ITC under Rule 88D" if trigger_88d else "Within statutory tolerance (No DRC-01C)"
    }

    return {"rule_88c": res_88c, "rule_88d": res_88d}


# ═════════════════════════════════════════════════════════════════════════════
# 5. EXCEL WORKBOOK GENERATOR (openpyxl Formula-First)
# ═════════════════════════════════════════════════════════════════════════════

def generate_recon_workbook(recon_data, output_path):
    """
    Generates CA/CFO-Ready Excel Workbook with Segoe UI, formulas, and autofilters.
    Tabs:
    1. Executive Summary
    2. Matched Invoices
    3. Tax Differences (Mismatch)
    4. Missing in GSTR-2B (Blocked ITC)
    5. Missing in Books (Unclaimed ITC)
    6. Supplier Risk Matrix
    """
    wb = openpyxl.Workbook()
    thin_border = Border(
        left=Side(style='thin', color=COLOR_BORDER), right=Side(style='thin', color=COLOR_BORDER),
        top=Side(style='thin', color=COLOR_BORDER), bottom=Side(style='thin', color=COLOR_BORDER)
    )

    # 1. Executive Summary
    ws = wb.active
    ws.title = "Executive Summary"
    ws.views.sheetView[0].showGridLines = True
    ws.merge_cells("A1:E1")
    ws["A1"] = "GST Reconciliation & Compliance Audit Statement"
    ws["A1"].font = Font(name=FONT_FAMILY, size=15, bold=True, color="FFFFFF")
    ws["A1"].fill = PatternFill(start_color=COLOR_HEADER_BG, end_color=COLOR_HEADER_BG, fill_type="solid")
    ws["A1"].alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[1].height = 36

    s = recon_data["summary"]
    metrics = [
        ("Total Invoices in Books", s["total_books_records"], "Count"),
        ("Total Invoices on Portal", s["total_portal_records"], "Count"),
        ("Fully Matched Invoices (Safe ITC)", s["matched_count"], "Count"),
        ("Tax / Value Mismatches", s["tax_mismatch_count"], "Count"),
        ("Missing in 2B (Blocked ITC - Sec 16(2)(aa))", s["missing_in_2b_count"], "Count"),
        ("Missing in Books (Unclaimed Portal ITC)", s["missing_in_books_count"], "Count"),
        ("High-Risk Suppliers (Action: Hold Payment)", s.get("high_risk_vendors", 0), "Count"),
        ("----------------------------------------", "", ""),
        ("Total ITC Claimed in Books", s["total_books_itc"], "INR"),
        ("Total ITC Reflected on Portal", s["total_portal_itc"], "INR"),
        ("Matched & Eligible ITC", s["matched_itc"], "INR"),
        ("ITC at Risk (Cannot Claim in GSTR-3B)", s["itc_at_risk"], "INR"),
        ("Unclaimed Portal ITC (Cash Advantage)", s["unclaimed_itc"], "INR"),
    ]

    ws["A3"] = "Audit Parameter"
    ws["B3"] = "Metric"
    ws["A3"].font = Font(name=FONT_FAMILY, size=10, bold=True, color=COLOR_HEADER_TEXT)
    ws["A3"].fill = PatternFill(start_color=COLOR_ACCENT, end_color=COLOR_ACCENT, fill_type="solid")
    ws["B3"].font = Font(name=FONT_FAMILY, size=10, bold=True, color=COLOR_HEADER_TEXT)
    ws["B3"].fill = PatternFill(start_color=COLOR_ACCENT, end_color=COLOR_ACCENT, fill_type="solid")

    for r_idx, (label, val, fmt) in enumerate(metrics, start=4):
        ws.cell(row=r_idx, column=1, value=label).font = Font(name=FONT_FAMILY, size=10, bold=(fmt == "INR"))
        c2 = ws.cell(row=r_idx, column=2, value=val)
        c2.font = Font(name=FONT_FAMILY, size=10, bold=(fmt == "INR"))
        c2.border = thin_border
        if fmt == "INR":
            c2.number_format = "#,##0.00"
            if "at Risk" in label:
                c2.fill = PatternFill(start_color=COLOR_RISK_BG, end_color=COLOR_RISK_BG, fill_type="solid")
                c2.font = Font(name=FONT_FAMILY, size=10, bold=True, color=COLOR_RISK_TEXT)
            elif "Eligible" in label:
                c2.fill = PatternFill(start_color=COLOR_MATCHED_BG, end_color=COLOR_MATCHED_BG, fill_type="solid")
                c2.font = Font(name=FONT_FAMILY, size=10, bold=True, color=COLOR_MATCHED_TEXT)
            elif "Unclaimed" in label:
                c2.fill = PatternFill(start_color=COLOR_UNCLAIMED_BG, end_color=COLOR_UNCLAIMED_BG, fill_type="solid")
                c2.font = Font(name=FONT_FAMILY, size=10, bold=True, color=COLOR_UNCLAIMED_TEXT)
        elif fmt == "Count":
            c2.number_format = "#,##0"
        ws.row_dimensions[r_idx].height = 20

    def write_table(ws, title, headers, rows, total_cols=None):
        ws.title = title
        ws.views.sheetView[0].showGridLines = True
        for col_idx, h in enumerate(headers, start=1):
            cell = ws.cell(row=1, column=col_idx, value=h)
            cell.font = Font(name=FONT_FAMILY, size=10, bold=True, color=COLOR_HEADER_TEXT)
            cell.fill = PatternFill(start_color=COLOR_HEADER_BG, end_color=COLOR_HEADER_BG, fill_type="solid")
            cell.alignment = Alignment(horizontal="center", vertical="center")
            cell.border = thin_border
        ws.row_dimensions[1].height = 26

        curr_r = 1
        for r_idx, r in enumerate(rows, start=2):
            curr_r = r_idx
            for c_idx, val in enumerate(r, start=1):
                cell = ws.cell(row=r_idx, column=c_idx, value=val)
                cell.font = Font(name=FONT_FAMILY, size=9)
                cell.border = thin_border
                if isinstance(val, (int, float)):
                    cell.number_format = "#,##0.00"
                    cell.alignment = Alignment(horizontal="right", vertical="center")
                else:
                    cell.alignment = Alignment(horizontal="left", vertical="center")
            ws.row_dimensions[r_idx].height = 20

        if total_cols and rows:
            tot_r = curr_r + 1
            ws.cell(row=tot_r, column=1, value="Total").font = Font(name=FONT_FAMILY, size=10, bold=True)
            for c in total_cols:
                col_letter = get_column_letter(c)
                f_cell = ws.cell(row=tot_r, column=c)
                f_cell.value = f"=SUM({col_letter}2:{col_letter}{curr_r})"
                f_cell.font = Font(name=FONT_FAMILY, size=10, bold=True)
                f_cell.number_format = "#,##0.00"
                f_cell.border = thin_border
            ws.row_dimensions[tot_r].height = 22

        ws.auto_filter.ref = f"A1:{get_column_letter(len(headers))}{max(2, curr_r)}"
        for col in ws.columns:
            max_len = max(len(str(c.value or "")) for c in col)
            ws.column_dimensions[get_column_letter(col[0].column)].width = min(max(max_len + 3, 12), 45)

    # 2. Matched Invoices
    m_hdrs = ["Supplier GSTIN", "Supplier Name", "Books Inv", "Portal Inv", "Date", "Books Taxable (₹)", "Portal Taxable (₹)", "Books ITC (₹)", "Portal ITC (₹)", "Match Type"]
    m_rows = [[m["gstin"], m["supplier_name"], m["books_inv"], m["portal_inv"], m["books_date"], m["books_taxable"], m["portal_taxable"], m["books_itc"], m["portal_itc"], m["match_type"]] for m in recon_data["matched"]]
    write_table(wb.create_sheet(), "Matched Invoices", m_hdrs, m_rows, total_cols=[6, 7, 8, 9])

    # 3. Tax Differences (Mismatch)
    diff_hdrs = ["Supplier GSTIN", "Supplier Name", "Books Inv", "Portal Inv", "Date", "Books Taxable (₹)", "Portal Taxable (₹)", "Taxable Diff (₹)", "Books ITC (₹)", "Portal ITC (₹)", "ITC Diff (₹)"]
    diff_rows = [[m["gstin"], m["supplier_name"], m["books_inv"], m["portal_inv"], m["books_date"], m["books_taxable"], m["portal_taxable"], m["taxable_diff"], m["books_itc"], m["portal_itc"], m["itc_diff"]] for m in recon_data.get("tax_mismatch", [])]
    write_table(wb.create_sheet(), "Tax Discrepancies", diff_hdrs, diff_rows, total_cols=[6, 7, 8, 9, 10, 11])

    # 4. Missing in 2B (Blocked ITC)
    m2b_hdrs = ["Supplier GSTIN", "Supplier Name", "Books Invoice", "Date", "Taxable (₹)", "ITC at Risk (₹)", "Statutory Impact", "Action"]
    m2b_rows = [[m["gstin"], m["supplier_name"], m["books_inv"], m["books_date"], m["books_taxable"], m["books_itc"], "Blocked (Sec 16(2)(aa))", m["action"]] for m in recon_data["missing_in_2b"]]
    write_table(wb.create_sheet(), "Missing in GSTR-2B", m2b_hdrs, m2b_rows, total_cols=[5, 6])

    # 5. Missing in Books (Unclaimed ITC)
    mib_hdrs = ["Supplier GSTIN", "Supplier Name", "Portal Invoice", "Date", "Taxable (₹)", "Portal ITC (₹)", "Statutory Benefit", "Commercial Action"]
    mib_rows = [[m["gstin"], m["supplier_name"], m["portal_inv"], m["portal_date"], m["portal_taxable"], m["portal_itc"], "Unclaimed ITC Available", m["action"]] for m in recon_data.get("missing_in_books", [])]
    write_table(wb.create_sheet(), "Unclaimed Portal ITC", mib_hdrs, mib_rows, total_cols=[5, 6])

    # 6. Supplier Risk Matrix
    vr_hdrs = ["Supplier GSTIN", "Supplier Name", "Total Invoices", "Total ITC (₹)", "Missing Invoices", "ITC at Risk (₹)", "Risk Score (0-100)", "RAG Status", "Statutory Action"]
    vr_rows = [[v["gstin"], v["name"], v["invoices"], v["total_itc"], v["missing_count"], v["missing_itc"], v["score"], v["rag"], v["action"]] for v in recon_data.get("vendor_risks", [])]
    write_table(wb.create_sheet(), "Supplier Risk Matrix", vr_hdrs, vr_rows, total_cols=[4, 6])

    for col in ws.columns:
        max_len = max(len(str(c.value or "")) for c in col)
        ws.column_dimensions[get_column_letter(col[0].column)].width = max(max_len + 4, 25)

    wb.save(output_path)
    return output_path
