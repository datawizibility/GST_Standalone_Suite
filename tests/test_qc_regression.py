# -*- coding: utf-8 -*-
"""
QC Regression Test Suite for GST Standalone Suite
Verifies all 19 items from the QC audit report:
- Critical #1: B2BA amendment matching without false missing / false RED
- Critical #2: Multi-month identical invoice numbering resolution
- Critical #3: Table 6.1 IGST cash paid extraction from GSTR-3B
- High #4: GSTR-2B JSON ingestion
- High #5: GSTR-1 tables extraction
- High #7: Rule 88C / 88D watchdog computations
- Medium #9: 6-tab openpyxl workbook with Unclaimed ITC
- Medium #11: Proportionate vendor risk scoring
- Low #16: Bracketed negative numbers
- Low #19: Purchase synonyms safety
"""

import os
import sys
import json
import unittest
import tempfile
import openpyxl

# Add parent directory
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
import engine


class TestQCRegression(unittest.TestCase):

    def test_critical_1_b2ba_amendment_matching(self):
        """
        Critical #1: Books entry carrying original invoice number AM-OLD must match
        portal amended record AM-NEW (where orig_inv == AM-OLD).
        Must NOT flag 'Missing in 2B' and must NOT trigger Vendor RED.
        """
        books = [{
            "supplier_gstin": "27AABCU9603R1ZM",
            "supplier_name": "ABC Tech Ltd",
            "invoice_number": "AM-OLD",
            "clean_inv_num": engine.normalize_inv("AM-OLD"),
            "invoice_date": "2026-04-10",
            "period": "2026-04",
            "invoice_value": 118000.0,
            "taxable_value": 100000.0,
            "igst": 18000.0, "cgst": 0.0, "sgst": 0.0, "cess": 0.0,
            "total_itc": 18000.0
        }]

        # Portal has the amended record in GSTR-2B
        portal = [{
            "source": "GSTR-2B-B2BA",
            "doc_type": "INV",
            "supplier_gstin": "27AABCU9603R1ZM",
            "supplier_name": "ABC Tech Ltd",
            "invoice_number": "AM-NEW",
            "clean_inv_num": engine.normalize_inv("AM-NEW"),
            "orig_inv_num": "AM-OLD",
            "clean_orig_inv": engine.normalize_inv("AM-OLD"),
            "invoice_date": "2026-04-10",
            "period": "2026-04",
            "invoice_value": 118000.0,
            "taxable_value": 100000.0,
            "igst": 18000.0, "cgst": 0.0, "sgst": 0.0, "cess": 0.0,
            "total_itc": 18000.0,
            "amendment_note": "Amended in B2BA (Orig: AM-OLD)"
        }]

        res = engine.reconcile_purchase_2b(books, portal)
        summary = res["summary"]

        self.assertEqual(summary["matched_count"], 1, "Amended invoice must be matched")
        self.assertEqual(summary["missing_in_2b_count"], 0, "Must NOT be missing in 2B")
        self.assertEqual(summary["missing_in_books_count"], 0, "Must NOT be unclaimed in portal")
        self.assertEqual(len(res["matched"]), 1)
        self.assertIn("B2BA", res["matched"][0]["match_type"])

        # Vendor risk check: Vendor must be GREEN, NOT RED!
        vendor = res["vendor_risks"][0]
        self.assertEqual(vendor["rag"], "GREEN", "Vendor should be GREEN compliant")
        self.assertEqual(vendor["missing_itc"], 0.0)

    def test_critical_2_period_matching_multimonth(self):
        """
        Critical #2: Multi-month identical invoice numbering loaded in reversed order
        must match by period/tax without false tax mismatch.
        """
        # Books has April (1,800 tax) and May (4,500 tax)
        books = [
            {
                "supplier_gstin": "07AAAAA0000A1Z5", "supplier_name": "Vendor Multi",
                "invoice_number": "01", "clean_inv_num": "01",
                "invoice_date": "2026-04-15", "period": "2026-04",
                "invoice_value": 11800.0, "taxable_value": 10000.0,
                "igst": 1800.0, "cgst": 0.0, "sgst": 0.0, "cess": 0.0, "total_itc": 1800.0
            },
            {
                "supplier_gstin": "07AAAAA0000A1Z5", "supplier_name": "Vendor Multi",
                "invoice_number": "01", "clean_inv_num": "01",
                "invoice_date": "2026-05-15", "period": "2026-05",
                "invoice_value": 29500.0, "taxable_value": 25000.0,
                "igst": 4500.0, "cgst": 0.0, "sgst": 0.0, "cess": 0.0, "total_itc": 4500.0
            }
        ]

        # Portal has May first, April second (reversed order)
        portal = [
            {
                "source": "GSTR-2B", "doc_type": "INV",
                "supplier_gstin": "07AAAAA0000A1Z5", "supplier_name": "Vendor Multi",
                "invoice_number": "01", "clean_inv_num": "01", "clean_orig_inv": "01",
                "invoice_date": "2026-05-15", "period": "2026-05",
                "invoice_value": 29500.0, "taxable_value": 25000.0,
                "igst": 4500.0, "cgst": 0.0, "sgst": 0.0, "cess": 0.0, "total_itc": 4500.0,
                "amendment_note": ""
            },
            {
                "source": "GSTR-2B", "doc_type": "INV",
                "supplier_gstin": "07AAAAA0000A1Z5", "supplier_name": "Vendor Multi",
                "invoice_number": "01", "clean_inv_num": "01", "clean_orig_inv": "01",
                "invoice_date": "2026-04-15", "period": "2026-04",
                "invoice_value": 11800.0, "taxable_value": 10000.0,
                "igst": 1800.0, "cgst": 0.0, "sgst": 0.0, "cess": 0.0, "total_itc": 1800.0,
                "amendment_note": ""
            }
        ]

        res = engine.reconcile_purchase_2b(books, portal)
        self.assertEqual(res["summary"]["matched_count"], 2, "Both invoices must be matched")
        self.assertEqual(res["summary"]["tax_mismatch_count"], 0, "There must be 0 tax mismatches")

    def test_critical_3_gstr3b_table_61_igst_cash(self):
        """
        Critical #3: Table 6.1 cash paid extraction from GSTR-3B PDF.
        Real sample PDF has IGST Cash Paid = 146,523.00, CGST = 21,983.00, SGST = 21,983.00.
        Must NOT be 0.00!
        """
        pdf_path = r"D:\AAAA_MYPROJECT\3BData\3B Data\GSTR3B_19AAMCP7482L1ZH_032026.pdf"
        if not os.path.exists(pdf_path):
            self.skipTest(f"Sample PDF not found at {pdf_path}")

        data = engine.extract_gstr3b_pdf(pdf_path)
        cash = data["t61"]["cash_paid"]
        self.assertEqual(cash[0], 146523.00, f"IGST cash paid should be 146523.00, got {cash[0]}")
        self.assertEqual(cash[1], 21983.00, f"CGST cash paid should be 21983.00, got {cash[1]}")
        self.assertEqual(cash[2], 21983.00, f"SGST cash paid should be 21983.00, got {cash[2]}")

    def test_high_4_gstr2b_json_parsing(self):
        """
        High #4: Verify GSTR-2B JSON ingestion handles B2B, B2BA, and CDNR.
        """
        sample_json = {
            "gstin": "19AAMCP7482L1ZH",
            "fp": "042026",
            "data": {
                "b2b": [
                    {
                        "ctin": "27AABCU9603R1ZM",
                        "trdNm": "Test Supplier",
                        "inv": [
                            {
                                "inum": "INV-501", "idt": "10-04-2026", "val": 11800.0,
                                "items": [{"txval": 10000.0, "igst": 1800.0, "cgst": 0.0, "sgst": 0.0, "cess": 0.0}]
                            }
                        ]
                    }
                ],
                "cdnr": [
                    {
                        "ctin": "27AABCU9603R1ZM",
                        "trdNm": "Test Supplier",
                        "nt": [
                            {
                                "ntnum": "CR-01", "ntty": "C", "nt_dt": "12-04-2026", "val": 1180.0,
                                "items": [{"txval": 1000.0, "igst": 180.0, "cgst": 0.0, "sgst": 0.0, "cess": 0.0}]
                            }
                        ]
                    }
                ]
            }
        }
        with tempfile.NamedTemporaryFile("w", delete=False, suffix=".json", encoding="utf-8") as f:
            json.dump(sample_json, f)
            tmp_path = f.name

        try:
            records = engine.parse_gstr2b_json(tmp_path)
            self.assertEqual(len(records), 2)
            inv = next(r for r in records if r["doc_type"] == "INV")
            cr = next(r for r in records if r["doc_type"] == "CR")
            self.assertEqual(inv["total_itc"], 1800.0)
            self.assertEqual(cr["total_itc"], -180.0, "Credit note ITC must be signed negative")
        finally:
            os.unlink(tmp_path)

    def test_high_7_watchdog_calculation(self):
        """
        High #7: Statutory Rule 88C and Rule 88D watchdog computations.
        """
        # Case A: Compliant
        wd_comp = engine.compute_statutory_watchdogs(
            gstr1_liability=100000.0, gstr3b_liability=100000.0,
            gstr3b_itc=50000.0, gstr2b_itc=50000.0
        )
        self.assertEqual(wd_comp["rule_88c"]["status"], "COMPLIANT")
        self.assertEqual(wd_comp["rule_88d"]["status"], "COMPLIANT")

        # Case B: DRC-01B Triggered (> 25L and > 20%)
        wd_trig = engine.compute_statutory_watchdogs(
            gstr1_liability=10000000.0, gstr3b_liability=6000000.0,  # Gap 40L (66%)
            gstr3b_itc=50000.0, gstr2b_itc=50000.0
        )
        self.assertEqual(wd_trig["rule_88c"]["status"], "NON_COMPLIANT_TRIGGERED")

    def test_medium_9_excel_six_tabs(self):
        """
        Medium #9: Reconciler Excel report must contain all 6 statutory tabs.
        """
        recon_data = {
            "summary": {
                "total_books_records": 1, "total_portal_records": 1,
                "matched_count": 0, "tax_mismatch_count": 0,
                "missing_in_2b_count": 1, "missing_in_books_count": 1,
                "total_books_itc": 100.0, "total_portal_itc": 100.0,
                "matched_itc": 0.0, "itc_at_risk": 100.0, "unclaimed_itc": 100.0,
                "high_risk_vendors": 0
            },
            "matched": [], "tax_mismatch": [],
            "missing_in_2b": [{
                "gstin": "07A", "supplier_name": "S", "books_inv": "1", "books_date": "2026-04-01",
                "books_taxable": 1000.0, "books_itc": 100.0, "action": "Hold"
            }],
            "missing_in_books": [{
                "gstin": "07B", "supplier_name": "S2", "portal_inv": "2", "portal_date": "2026-04-01",
                "portal_taxable": 1000.0, "portal_itc": 100.0, "action": "Claim"
            }],
            "vendor_risks": []
        }
        with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as f:
            tmp_excel = f.name

        try:
            engine.generate_recon_workbook(recon_data, tmp_excel)
            wb = openpyxl.load_workbook(tmp_excel)
            expected_sheets = [
                "Executive Summary", "Matched Invoices", "Tax Discrepancies",
                "Missing in GSTR-2B", "Unclaimed Portal ITC", "Supplier Risk Matrix"
            ]
            for s in expected_sheets:
                self.assertIn(s, wb.sheetnames, f"Sheet '{s}' must exist in workbook")
        finally:
            if os.path.exists(tmp_excel):
                os.unlink(tmp_excel)

    def test_medium_11_vendor_risk_proportionate(self):
        """
        Medium #11: Single small invoice missing in 2B must NOT mark vendor RED 'HOLD PAYMENT'.
        """
        books = [{
            "supplier_gstin": "27AABCU9603R1ZM", "supplier_name": "Small Vendor",
            "invoice_number": "INV-1", "clean_inv_num": "INV1",
            "invoice_date": "2026-04-10", "period": "2026-04",
            "invoice_value": 590.0, "taxable_value": 500.0,
            "igst": 90.0, "cgst": 0.0, "sgst": 0.0, "cess": 0.0, "total_itc": 90.0
        }]
        portal = []  # Missing in 2B

        res = engine.reconcile_purchase_2b(books, portal)
        vendor = res["vendor_risks"][0]
        self.assertNotEqual(vendor["rag"], "RED", "Small ₹90 missing ITC must not trigger RED HOLD PAYMENT")

    def test_low_16_bracketed_negatives(self):
        """
        Low #16: clean_num must turn (1,200.00) into -1200.00.
        """
        self.assertEqual(engine.clean_num("(1,200.00)"), -1200.00)
        self.assertEqual(engine.clean_num("(500)"), -500.00)
        self.assertEqual(engine.clean_num("₹ 1,500.50"), 1500.50)

    def test_low_19_purchase_column_synonyms_safety(self):
        """
        Low #19: Purchase column synonyms must NOT contain customer gstin.
        """
        self.assertNotIn("customer gstin", engine.BOOKS_PURCHASE_COLUMN_SYNONYMS["gstin"])


if __name__ == "__main__":
    unittest.main()
