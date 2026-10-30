# Lessons Learned & Technical Discoveries — GST Standalone Intelligence Suite

## 1. GSTR-2B Multi-Tab Excel Ingestion Structure
The official GST portal export uses a multi-tier header: row 5 contains primary headings (`GSTIN of supplier`, `Trade/Legal name`, `Invoice Details`, `Taxable Value (₹)`), while row 6 contains subheadings (`Invoice number`, `Invoice type`, `Invoice Date`, `Invoice Value(₹)`, `Integrated Tax(₹)`, `Central Tax(₹)`, `State/UT Tax(₹)`, `Cess(₹)`).
Data rows begin strictly at `header_row + 2` (row 7).

## 2. Resolving B2BA Invoice Amendments
When suppliers amend previous invoices, the amendment is filed in Table `B2BA`. In `B2BA`, columns 1–2 record `Original Invoice number` and `Original Invoice Date`, while columns 5–16 record `Revised Invoice Details`.
Without an amendment replacement pass, naive reconcilers flag these as both a false "Missing in 2B" (under original invoice) and a false "Missing in Books" (under revised invoice).
**Resolution**: Map `(supplier_gstin, clean_orig_inv)` and replace the record with `(supplier_gstin, clean_rev_inv)` while annotating `Amended in B2BA (Orig: X)` so that the matcher resolves whether the accountant recorded either reference in ERP.

## 3. Credit Note Signing Convention
In `B2B-CDNR`, Credit Notes reduce input tax credit, while Debit Notes increase it. Credit Notes must be multiplied by `-1.0` across taxable base and tax heads (IGST, CGST, SGST, Cess) so that aggregate ITC formulas and mathematical checks balance correctly against the trial balance.

## 4. PyInstaller Exclusion Optimization
Compiling data science tools often pulls in gigantic heavy packages like `torch`, `matplotlib`, and `scipy` if present in the global virtual environment. Adding explicit exclusions `--exclude-module torch --exclude-module scipy --exclude-module matplotlib --exclude-module IPython --exclude-module pytest` reduces build time to under 90 seconds and produces a lean ~111MB portable executable.

## 5. Candidate Scoring vs First-Match Pitfalls in Reconcilers
When suppliers issue monthly recurring invoices (e.g., utility bills, rent, internet) using identical or resetting sequence numbers (e.g. `01`, `OCT-01`), matching solely on `(gstin, inv_no)` without date/period and amount scoring will corrupt the match if multi-month files are loaded in non-chronological order.
A robust reconciler must score candidates by prioritizing:
1. Exact tax match within tolerance + Date match (highest score: 150)
2. Exact tax match within tolerance (score: 100)
3. Date match (score: 50)
This ensures that the correct monthly invoice is paired, preventing spurious tax discrepancies.

## 6. Table 6.1 PDF Cash Paid Extraction Mechanics
In GSTR-3B Table 6.1 ("Payment of tax"), PDF rendering libraries often split wrapped decimal numbers across lines (e.g. `204300.0\n0`). Furthermore, Section (A) represents forward charge tax paid while Section (B) represents reverse charge liability.
Because the number of ITC columns varies by tax head and PDF generator, rigidly hardcoding column offsets fails. However, across both old 8-column and modern 10-column portal layouts, **Tax paid in cash** is universally located at position `-3` (third from the right, immediately preceding Interest and Late Fee). Isolating Section (A) and querying `vals[-3]` guarantees 100% extraction accuracy across all filing formats.

## 7. Proportionate Vendor Risk Weighting
Evaluating vendor compliance solely on the ratio of missing invoices produces severe false positives for one-off vendors (e.g. 1 bill of ₹500 unfiled → 100% ratio → RED "HOLD PAYMENT"). An institutional risk engine must weight both the missing ratio (60%) and the actual ₹ ITC at risk (40%), reserving RED status only for vendors exceeding ₹50,000 at risk or systemic high default rates.
