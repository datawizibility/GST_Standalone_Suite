# -*- coding: utf-8 -*-
"""
GST Intelligence & Reconciliation Engine (Standalone Suite)
Zero-dependency, production-grade GST statutory engine combining:
1. GSTR-2B vs ERP Purchase Reconciler (4-Pass, B2BA overrides, Multi-Month, Vendor Risk RAG)
2. GSTR-1 vs Books Sales Reconciler
3. GSTR-3B PDF Data Extractor (Table 3.1, 4, 5, 5.1, 6.1, ARN)
4. GSTR-1 / 1A PDF Data Extractor (Table 4, 5, 7, 8, 9B, Total Liability, ARN)
5. Rule 88C (DRC-01B) and Rule 88D (DRC-01C) Exposure Watchdogs
6. CA/CFO-Ready Excel Workbook Generation (openpyxl)
"""

import os
import sys
import re
import json
from datetime import datetime
import pandas as pd
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

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
COLOR_BORDER = "D9D9D9"

MONTH_MAP = {
    "jan": 1, "feb": 2, "mar": 3, "apr": 4, "may": 5, "jun": 6,
    "jul": 7, "aug": 8, "sep": 9, "oct": 10, "nov": 11, "dec": 12
}

BOOKS_COLUMN_SYNONYMS = {
    "gstin": ["gstin", "gstin/uin", "party gstin", "supplier gstin", "customer gstin", "gstin no"],
    "party_name": ["party name", "supplier name", "customer name", "ledger name", "vendor name", "particulars", "name"],
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
    if val is None or pd.isna(val) or str(val).strip() in ("-", "", ".", "None"):
        return 0.0
    if isinstance(val, (int, float)):
        return round(float(val), 2)
    s = re.sub(r"[^\d\.-]", "", str(val).replace(",", "").replace("₹", ""))
    try:
        return round(float(s), 2)
    except Exception:
        return 0.0


def normalize_inv(inv_str):
    if not inv_str:
        return ""
    clean = re.sub(r"[^A-Z0-9]", "", str(inv_str).strip().upper())
    return clean


def parse_date(date_val):
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
    return s[:10]


# ═════════════════════════════════════════════════════════════════════════════
# 1. GSTR-2B INGESTION & B2BA AMENDMENTS ENGINE
# ═════════════════════════════════════════════════════════════════════════════

def parse_gstr2b_excel(file_path):
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
                "invoice_date": inv_dt, "invoice_value": inv_val, "taxable_value": taxable,
                "igst": igst, "cgst": cgst, "sgst": sgst, "cess": cess,
                "total_itc": round(igst + cgst + sgst + cess, 2), "amendment_note": ""
            })

    # 2. B2BA
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
                        "invoice_date": rev_dt, "invoice_value": rev_val, "taxable_value": rev_taxable,
                        "igst": igst, "cgst": cgst, "sgst": sgst, "cess": cess,
                        "total_itc": round(igst + cgst + sgst + cess, 2),
                        "amendment_note": f"Amended in B2BA (Orig: {orig_inv})"
                    })

    # 3. CDNR
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
                    "invoice_date": nt_dt, "invoice_value": round(val, 2), "taxable_value": round(taxable, 2),
                    "igst": round(igst, 2), "cgst": round(cgst, 2), "sgst": round(sgst, 2), "cess": round(cess, 2),
                    "total_itc": round(igst + cgst + sgst + cess, 2), "amendment_note": ""
                })

    # Apply amendments replacement
    inv_map = {(inv["supplier_gstin"], inv["clean_inv_num"]): inv for inv in invoices}
    for amend in amendments:
        orig_key = (amend["supplier_gstin"], amend.get("clean_orig_inv", ""))
        rev_key = (amend["supplier_gstin"], amend["clean_inv_num"])
        if orig_key in inv_map:
            del inv_map[orig_key]
        inv_map[rev_key] = amend

    return list(inv_map.values())


def parse_books_excel(file_path):
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
    for standard_col, syns in BOOKS_COLUMN_SYNONYMS.items():
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
            "supplier_gstin": gstin, "customer_gstin": gstin, "supplier_name": name, "party_name": name,
            "invoice_number": inv_no, "clean_inv_num": normalize_inv(inv_no), "invoice_date": inv_dt,
            "invoice_value": tot, "taxable_value": taxable, "igst": igst, "cgst": cgst, "sgst": sgst, "cess": cess,
            "total_itc": round(igst + cgst + sgst + cess, 2), "total_tax": round(igst + cgst + sgst + cess, 2)
        })
    return records


# ═════════════════════════════════════════════════════════════════════════════
# 2. VENDOR RISK ENGINE & 4-PASS RECONCILER
# ═════════════════════════════════════════════════════════════════════════════

def reconcile_purchase_2b(books, portal, tolerance=2.0):
    matched = []
    mismatch = []
    missing_2b = []
    missing_books = []
    used_portal = set()
    used_books = set()

    # Pass 1: Exact Match
    for b_idx, b in enumerate(books):
        if not b.get("supplier_gstin") or not b.get("clean_inv_num"):
            continue
        for p_idx, p in enumerate(portal):
            if p_idx in used_portal:
                continue
            if b["supplier_gstin"] == p["supplier_gstin"] and b["clean_inv_num"] == p["clean_inv_num"]:
                diff_tx = abs(b["taxable_value"] - p["taxable_value"])
                diff_itc = abs(b["total_itc"] - p["total_itc"])
                if diff_tx <= tolerance and diff_itc <= tolerance:
                    matched.append({
                        "match_type": "Exact Match" if not p.get("amendment_note") else f"Exact ({p['amendment_note']})",
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
                    break

    # Pass 2: Probable / Substring match
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

    # Compute Vendor Risk Matrix (0-100 score & RAG)
    vendors = {}
    for b in books:
        gstin = b.get("supplier_gstin")
        if not gstin: continue
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
        score = int(ratio * 100)
        v["score"] = score
        if score >= 50 or v["missing_itc"] > 50000:
            v["rag"] = "RED"
            v["action"] = "HOLD PAYMENT - High ITC Default Risk (Sec 16(2)(aa))"
        elif score >= 20 or v["missing_itc"] > 5000:
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

def extract_pdf_text(pdf_path):
    text = ""
    if HAS_PDFPLUMBER:
        try:
            with pdfplumber.open(pdf_path) as pdf:
                text = "\n".join((p.extract_text() or "") for p in pdf.pages)
                if text.strip():
                    return text
        except Exception:
            pass
    if HAS_PYPDF:
        try:
            reader = pypdf.PdfReader(pdf_path)
            text = "\n".join((p.extract_text() or "") for p in reader.pages)
        except Exception:
            pass
    return text


def extract_gstr3b_pdf(pdf_path):
    text = extract_pdf_text(pdf_path)
    data = {
        "filename": os.path.basename(pdf_path), "gstin": "N/A", "legal_name": "N/A",
        "year": "N/A", "period": "N/A", "arn": "N/A", "arn_date": "N/A",
        "t31": {"outward_taxable": [0.0]*5, "rcm": [0.0]*5},
        "t4": {"net_itc": [0.0]*4},
        "t61": {"cash_paid": [0.0]*4}
    }
    if not text: return data

    m_g = re.search(r"GSTIN.*?([0-9]{2}[A-Z]{5}[0-9]{4}[A-Z]{1}[1-9A-Z]{1}Z[0-9A-Z]{1})", text)
    if m_g: data["gstin"] = m_g.group(1).upper()

    m_n = re.search(r"Legal name.*?\s+(.*)", text)
    if m_n: data["legal_name"] = m_n.group(1).split("\n")[0].strip()

    m_y = re.search(r"(?:Financial year|Year)\s*[:\-]?\s*(\d{4}\s*-\s*\d{2,4})", text, re.IGNORECASE)
    if m_y: data["year"] = re.sub(r"\s+", "", m_y.group(1))

    m_p = re.search(r"(?:Tax period|Period)\s*[:\-]?\s*([A-Za-z-]+)", text, re.IGNORECASE)
    if m_p: data["period"] = m_p.group(1).strip()

    m_a = re.search(r"ARN\s*[:\-]?\s*([A-Z0-9]{15,22})", text)
    if m_a: data["arn"] = m_a.group(1).strip()

    m_ad = re.search(r"(?:ARN Date|Date)\s*[:\-]?\s*(\d{2}/\d{2}/\d{4})", text)
    if m_ad: data["arn_date"] = m_ad.group(1).strip()

    # 3.1(a) Outward taxable
    m_31a = re.search(r"\(a\)[^0-9]+([\d,.-]+)\s+([\d,.-]+)\s+([\d,.-]+)\s+([\d,.-]+)\s+([\d,.-]+)", text)
    if m_31a:
        data["t31"]["outward_taxable"] = [clean_num(v) for v in m_31a.groups()]

    # 3.1(d) RCM
    m_31d = re.search(r"\(d\)[^0-9]+([\d,.-]+)\s+([\d,.-]+)\s+([\d,.-]+)\s+([\d,.-]+)\s+([\d,.-]+)", text)
    if m_31d:
        data["t31"]["rcm"] = [clean_num(v) for v in m_31d.groups()]

    # Table 4 Net ITC
    m_4c = re.search(r"\(C\)\s*Net\s+ITC\s+Available[^0-9]+([\d,.-]+)\s+([\d,.-]+)\s+([\d,.-]+)\s+([\d,.-]+)", text, re.IGNORECASE) or \
           re.search(r"\(5\)\s*All\s+other\s+ITC[^0-9]+([\d,.-]+)\s+([\d,.-]+)\s+([\d,.-]+)\s+([\d,.-]+)", text, re.IGNORECASE)
    if m_4c:
        data["t4"]["net_itc"] = [clean_num(v) for v in m_4c.groups()]

    # Table 6.1 Cash paid
    m_cash_i = re.search(r"Integrated\s+Tax[^0-9]+[\d,.-]+\s+[\d,.-]+\s+[\d,.-]+\s+[\d,.-]+\s+[\d,.-]+\s+([\d,.-]+)", text)
    m_cash_c = re.search(r"Central\s+Tax[^0-9]+[\d,.-]+\s+[\d,.-]+\s+[\d,.-]+\s+[\d,.-]+\s+([\d,.-]+)", text)
    m_cash_s = re.search(r"(?:State/UT|State)\s+Tax[^0-9]+[\d,.-]+\s+[\d,.-]+\s+[\d,.-]+\s+[\d,.-]+\s+([\d,.-]+)", text)
    data["t61"]["cash_paid"] = [
        clean_num(m_cash_i.group(1)) if m_cash_i else 0.0,
        clean_num(m_cash_c.group(1)) if m_cash_c else 0.0,
        clean_num(m_cash_s.group(1)) if m_cash_s else 0.0,
        0.0
    ]
    return data


def extract_gstr1_pdf(pdf_path):
    text = extract_pdf_text(pdf_path)
    data = {
        "filename": os.path.basename(pdf_path), "gstin": "N/A", "legal_name": "N/A",
        "year": "N/A", "period": "N/A", "arn": "N/A", "arn_date": "N/A", "type": "GSTR-1",
        "b2b": [0.0]*5, "b2cs": [0.0]*5, "total_liability": [0.0]*5
    }
    if not text: return data

    if "GSTR-1A" in text: data["type"] = "GSTR-1A"
    elif "IFF" in text: data["type"] = "GSTR-1 IFF"

    m_g = re.search(r"GSTIN.*?([0-9]{2}[A-Z]{5}[0-9]{4}[A-Z]{1}[1-9A-Z]{1}Z[0-9A-Z]{1})", text)
    if m_g: data["gstin"] = m_g.group(1).upper()

    m_n = re.search(r"Legal name.*?\s+(.*)", text)
    if m_n: data["legal_name"] = m_n.group(1).split("\n")[0].strip()

    m_y = re.search(r"(?:Financial year|Year)\s*[:\-]?\s*(\d{4}\s*-\s*\d{2,4})", text, re.IGNORECASE)
    if m_y: data["year"] = re.sub(r"\s+", "", m_y.group(1))

    m_p = re.search(r"(?:Tax period|Period)\s*[:\-]?\s*([A-Za-z-]+)", text, re.IGNORECASE)
    if m_p: data["period"] = m_p.group(1).strip()

    m_a = re.search(r"ARN\s*[:\-]?\s*([A-Z0-9]+)", text)
    if m_a: data["arn"] = m_a.group(1).strip()

    m_ad = re.search(r"(?:ARN Date|Date)\s*[:\-]?\s*(\d{2}/\d{2}/\d{4})", text)
    if m_ad: data["arn_date"] = m_ad.group(1).strip()

    # B2B values
    m_b2b = re.search(r"(?:4A, 4B.*?B2B|B2B Regular).*?Total\s+([\d,.-]+)\s+([\d,.-]+)\s+([\d,.-]+)\s+([\d,.-]+)", text, re.DOTALL)
    if m_b2b: data["b2b"] = [clean_num(v) for v in m_b2b.groups()] + [0.0]

    # Total Liability
    tl_m = re.search(r"Total Liability.*?\s+([\d,.-]+)\s+([\d,.-]+)\s+([\d,.-]+)\s+([\d,.-]+)\s+([\d,.-]+)", text, re.IGNORECASE)
    if tl_m: data["total_liability"] = [clean_num(v) for v in tl_m.groups()]

    return data


# ═════════════════════════════════════════════════════════════════════════════
# 4. EXCEL WORKBOOK GENERATOR (openpyxl Formula-First)
# ═════════════════════════════════════════════════════════════════════════════

def generate_recon_workbook(recon_data, output_path):
    wb = openpyxl.Workbook()
    thin_border = Border(
        left=Side(style='thin', color=COLOR_BORDER), right=Side(style='thin', color=COLOR_BORDER),
        top=Side(style='thin', color=COLOR_BORDER), bottom=Side(style='thin', color=COLOR_BORDER)
    )

    # 1. Summary Dashboard
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

        for col in ws.columns:
            max_len = max(len(str(c.value or "")) for c in col)
            ws.column_dimensions[get_column_letter(col[0].column)].width = min(max(max_len + 3, 12), 45)

    # 2. Matched
    m_hdrs = ["Supplier GSTIN", "Supplier Name", "Books Inv", "Portal Inv", "Date", "Books Taxable (₹)", "Portal Taxable (₹)", "Books ITC (₹)", "Portal ITC (₹)", "Match Type"]
    m_rows = [[m["gstin"], m["supplier_name"], m["books_inv"], m["portal_inv"], m["books_date"], m["books_taxable"], m["portal_taxable"], m["books_itc"], m["portal_itc"], m["match_type"]] for m in recon_data["matched"]]
    write_table(wb.create_sheet(), "Matched Invoices", m_hdrs, m_rows, total_cols=[6, 7, 8, 9])

    # 3. Missing in 2B (Blocked ITC)
    m2b_hdrs = ["Supplier GSTIN", "Supplier Name", "Books Invoice", "Date", "Taxable (₹)", "ITC at Risk (₹)", "Statutory Impact", "Action"]
    m2b_rows = [[m["gstin"], m["supplier_name"], m["books_inv"], m["books_date"], m["books_taxable"], m["books_itc"], "Blocked (Sec 16(2)(aa))", m["action"]] for m in recon_data["missing_in_2b"]]
    write_table(wb.create_sheet(), "Missing in GSTR-2B", m2b_hdrs, m2b_rows, total_cols=[5, 6])

    # 4. Supplier Risk Matrix
    vr_hdrs = ["Supplier GSTIN", "Supplier Name", "Total Invoices", "Total ITC (₹)", "Missing Invoices", "ITC at Risk (₹)", "Score (0-100)", "RAG", "Commercial Action"]
    vr_rows = [[v["gstin"], v["name"], v["invoices"], v["total_itc"], v["missing_count"], v["missing_itc"], v["score"], v["rag"], v["action"]] for v in recon_data.get("vendor_risks", [])]
    write_table(wb.create_sheet(), "Supplier Risk Matrix", vr_hdrs, vr_rows, total_cols=[4, 6])

    for col in ws.columns:
        max_len = max(len(str(c.value or "")) for c in col)
        ws.column_dimensions[get_column_letter(col[0].column)].width = max(max_len + 4, 25)

    wb.save(output_path)
    return output_path
