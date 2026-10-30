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
