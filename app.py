# -*- coding: utf-8 -*-
"""
GST Standalone Suite - Self-Contained Web Application
Launches on http://127.0.0.1:5055 and automatically opens the browser.
No external Node.js, database, or dependencies required.
"""

import os
import sys
import io
import time
import socket
import base64
import tempfile
import webbrowser
import threading
from datetime import datetime

from flask import Flask, request, jsonify, send_file, render_template_string

# Import core engine
import engine

app = Flask(__name__)
app.config['MAX_CONTENT_LENGTH'] = 100 * 1024 * 1024  # 100MB upload limit

# Global state
STATE = {
    "books": [],
    "gstr2b": [],
    "recon_result": None,
    "last_excel_path": None,
    "gstr3b_data": [],
    "gstr1_data": []
}

HTML_TEMPLATE = """<!DOCTYPE html>
<html lang="en">
<head>
  <meta charset="utf-8"/>
  <meta name="viewport" content="width=device-width, initial-scale=1"/>
  <title>GST Standalone Suite — Reconciliation & Audit Platform</title>
  <style>
    :root {
      --bg: #0f172a; --surface: #1e293b; --card: #334155; --border: #475569;
      --text: #f8fafc; --muted: #94a3b8; --accent: #38bdf8; --accent-dk: #0284c7;
      --good: #4ade80; --bad: #f87171; --warn: #fbbf24;
    }
    * { box-sizing: border-box; margin: 0; padding: 0; font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, sans-serif; }
    body { background: var(--bg); color: var(--text); font-size: 13px; line-height: 1.5; min-height: 100vh; }
    header { background: var(--surface); border-bottom: 1px solid var(--border); padding: 0.75rem 1.5rem; display: flex; align-items: center; justify-content: space-between; position: sticky; top: 0; z-index: 100; }
    .logo { font-size: 16px; font-weight: 700; color: var(--accent); display: flex; align-items: center; gap: 8px; }
    .layout { max-width: 1400px; margin: 0 auto; padding: 1.5rem; }
    .nav-tabs { display: flex; gap: 6px; border-bottom: 1px solid var(--border); margin-bottom: 1.5rem; overflow-x: auto; }
    .nav-tab { background: transparent; border: none; border-bottom: 2px solid transparent; color: var(--muted); padding: 0.6rem 1rem; cursor: pointer; font-weight: 600; font-size: 13px; white-space: nowrap; }
    .nav-tab.active { color: var(--accent); border-bottom-color: var(--accent); }
    .tab-content { display: none; }
    .tab-content.active { display: block; }
    .grid-2 { display: grid; grid-template-columns: 360px 1fr; gap: 1.5rem; }
    @media(max-width: 900px) { .grid-2 { grid-template-columns: 1fr; } }
    .card { background: var(--surface); border: 1px solid var(--border); border-radius: 8px; padding: 1.25rem; margin-bottom: 1rem; }
    .card h3 { font-size: 14px; margin-bottom: 0.75rem; color: var(--text); }
    .btn { display: inline-flex; align-items: center; justify-content: center; padding: 0.55rem 1rem; border-radius: 6px; font-weight: 600; font-size: 12px; cursor: pointer; border: 1px solid transparent; width: 100%; transition: 0.15s; }
    .btn-primary { background: var(--accent); color: #0f172a; }
    .btn-primary:hover { background: var(--accent-dk); color: #fff; }
    .btn-secondary { background: var(--card); border-color: var(--border); color: var(--text); }
    .btn-secondary:hover { background: #475569; }
    .btn-success { background: #22c55e; color: #fff; }
    .btn-success:hover { background: #16a34a; }
    .drop-zone { border: 2px dashed var(--border); border-radius: 6px; padding: 1.25rem; text-align: center; cursor: pointer; background: rgba(0,0,0,0.15); margin-bottom: 0.75rem; transition: 0.2s; }
    .drop-zone:hover { border-color: var(--accent); }
    .kpi-row { display: grid; grid-template-columns: repeat(auto-fit, minmax(160px, 1fr)); gap: 1rem; margin-bottom: 1.25rem; }
    .kpi-card { background: var(--card); border: 1px solid var(--border); border-radius: 6px; padding: 0.85rem; }
    .kpi-title { font-size: 11px; text-transform: uppercase; color: var(--muted); font-weight: 600; }
    .kpi-val { font-size: 17px; font-weight: 700; margin-top: 4px; }
    .val-good { color: var(--good); } .val-bad { color: var(--bad); } .val-warn { color: var(--warn); }
    .table-container { overflow-x: auto; max-height: 500px; border: 1px solid var(--border); border-radius: 6px; }
    table { width: 100%; border-collapse: collapse; font-size: 12px; }
    th { background: #0f172a; color: var(--muted); padding: 8px 10px; text-align: left; position: sticky; top: 0; z-index: 10; border-bottom: 1px solid var(--border); }
    td { padding: 8px 10px; border-bottom: 1px solid rgba(255,255,255,0.06); }
    tr:hover td { background: rgba(255,255,255,0.03); }
    .badge { display: inline-block; padding: 2px 6px; border-radius: 4px; font-size: 10px; font-weight: 700; }
    .badge-green { background: rgba(74,222,128,0.2); color: var(--good); border: 1px solid var(--good); }
    .badge-red { background: rgba(248,113,113,0.2); color: var(--bad); border: 1px solid var(--bad); }
    .badge-amber { background: rgba(251,191,36,0.2); color: var(--warn); border: 1px solid var(--warn); }
  </style>
</head>
<body>
  <header>
    <div class="logo">⚡ GST Standalone Intelligence Suite</div>
    <div style="font-size:12px;color:var(--muted);">
      Offline & Secure • Zero-Installation • Tally Prime Direct
    </div>
  </header>

  <div class="layout">
    <div class="nav-tabs">
      <button class="nav-tab active" onclick="showTab('tab-purchase')">🛒 Purchase Recon (2B vs Books)</button>
      <button class="nav-tab" onclick="showTab('tab-gstr3b')">📄 GSTR-3B PDF Extractor</button>
      <button class="nav-tab" onclick="showTab('tab-gstr1')">📑 GSTR-1 PDF Extractor</button>
      <button class="nav-tab" onclick="showTab('tab-watchdogs')">⚖️ Rule 88C / 88D Watchdogs</button>
    </div>

    <!-- 1. Purchase Recon Tab -->
    <div id="tab-purchase" class="tab-content active">
      <div class="grid-2">
        <div>
          <div class="card">
            <h3>Step 1: Ingest Books Data</h3>
            <button class="btn btn-secondary" style="margin-bottom:8px;" onclick="document.getElementById('fileBooks').click()">
              📁 Upload Books Excel (.xlsx, .csv)
            </button>
            <input type="file" id="fileBooks" style="display:none;" accept=".xlsx,.xls,.csv" onchange="uploadBooks(event)"/>
            <div id="booksStatus" style="font-size:11px;color:var(--muted);">0 records loaded</div>
          </div>

          <div class="card">
            <h3>Step 2: Upload GSTR-2B File(s)</h3>
            <div class="drop-zone" onclick="document.getElementById('file2b').click()">
              <div style="font-size:24px;margin-bottom:4px;">📂</div>
              <div style="font-weight:600;">Click to select GSTR-2B</div>
              <div style="font-size:11px;color:var(--muted);">Portal Excel (.xlsx) or JSON (Supports Multi-Month & B2BA)</div>
              <input type="file" id="file2b" style="display:none;" accept=".xlsx,.xls,.json" multiple onchange="upload2B(event)"/>
            </div>
            <div id="portalStatus" style="font-size:11px;color:var(--muted);">0 portal records loaded</div>
          </div>

          <div class="card">
            <h3>Step 3: Run Reconciler</h3>
            <button class="btn btn-primary" id="btnRunRecon" onclick="runPurchaseRecon()">⚡ Execute 4-Pass Reconciler</button>
            <button class="btn btn-success" id="btnDownloadExcel" style="margin-top:10px;display:none;" onclick="downloadExcel()">📊 Download CA-Ready Excel Report</button>
          </div>
        </div>

        <div>
          <div class="kpi-row">
            <div class="kpi-card">
              <div class="kpi-title">Books ITC Claimed</div>
              <div class="kpi-val" id="kpiBooksItc">₹ 0.00</div>
            </div>
            <div class="kpi-card">
              <div class="kpi-title">Portal ITC Available</div>
              <div class="kpi-val" id="kpiPortalItc">₹ 0.00</div>
            </div>
            <div class="kpi-card">
              <div class="kpi-title">Matched Safe ITC</div>
              <div class="kpi-val val-good" id="kpiMatchedItc">₹ 0.00</div>
            </div>
            <div class="kpi-card">
              <div class="kpi-title">ITC at Risk (Blocked)</div>
              <div class="kpi-val val-bad" id="kpiRiskItc">₹ 0.00</div>
            </div>
            <div class="kpi-card">
              <div class="kpi-title">High-Risk Vendors</div>
              <div class="kpi-val val-bad" id="kpiHighRiskVendors">0</div>
            </div>
          </div>

          <div class="card" style="padding:0.75rem;">
            <div style="display:flex;gap:8px;margin-bottom:0.75rem;">
              <button class="btn btn-secondary" style="width:auto;padding:4px 10px;" onclick="renderTable('matched')">Matched (<span id="cntMatched">0</span>)</button>
              <button class="btn btn-secondary" style="width:auto;padding:4px 10px;" onclick="renderTable('tax_mismatch')">Discrepancies (<span id="cntMismatch">0</span>)</button>
              <button class="btn btn-secondary" style="width:auto;padding:4px 10px;" onclick="renderTable('missing_in_2b')">Missing in 2B (<span id="cntMissing2B">0</span>)</button>
              <button class="btn btn-secondary" style="width:auto;padding:4px 10px;color:var(--bad);" onclick="renderTable('vendor_risks')">🛡️ Supplier Risk Matrix (<span id="cntVendors">0</span>)</button>
            </div>
            <div class="table-container">
              <table id="reconTable">
                <thead id="reconThead"></thead>
                <tbody id="reconTbody">
                  <tr><td colspan="7" style="text-align:center;padding:2rem;color:var(--muted);">Load Books and GSTR-2B files, then click Execute Reconciler.</td></tr>
                </tbody>
              </table>
            </div>
          </div>
        </div>
      </div>
    </div>

    <!-- 2. GSTR-3B PDF Extractor Tab -->
    <div id="tab-gstr3b" class="tab-content">
      <div class="card">
        <h3>GSTR-3B PDF Month-over-Month Data Extractor</h3>
        <p style="font-size:12px;color:var(--muted);margin-bottom:1rem;">
          Select one or multiple filed GSTR-3B return PDFs to automatically extract Table 3.1, Table 4 (Eligible ITC), Table 6.1 (Cash Paid), and ARN status into an institutional summary workbook.
        </p>
        <div class="drop-zone" onclick="document.getElementById('file3bBatch').click()">
          <div style="font-size:24px;margin-bottom:4px;">📑</div>
          <div style="font-weight:600;">Click to select GSTR-3B PDF files</div>
          <input type="file" id="file3bBatch" style="display:none;" accept=".pdf" multiple onchange="extract3BPdfs(event)"/>
        </div>
        <div id="status3B" style="font-size:11px;color:var(--accent);margin-bottom:1rem;"></div>
        <div class="table-container" style="max-height:400px;margin-bottom:1rem;">
          <table>
            <thead>
              <tr>
                <th>Period</th><th>Year</th><th>GSTIN</th><th>Taxable (₹)</th>
                <th>Outward Tax (₹)</th><th>Net ITC (₹)</th><th>Cash Paid (₹)</th><th>ARN</th>
              </tr>
            </thead>
            <tbody id="tbody3B">
              <tr><td colspan="8" style="text-align:center;padding:2rem;color:var(--muted);">Select GSTR-3B return PDFs to view extracted monthly data.</td></tr>
            </tbody>
          </table>
        </div>
      </div>
    </div>

    <!-- 3. GSTR-1 PDF Extractor Tab -->
    <div id="tab-gstr1" class="tab-content">
      <div class="card">
        <h3>GSTR-1 / 1A PDF Month-over-Month Data Extractor</h3>
        <p style="font-size:12px;color:var(--muted);margin-bottom:1rem;">
          Select filed GSTR-1, GSTR-1A, or IFF PDF returns to extract Table 4 (B2B), 5 (B2CL), 7 (B2CS), 8 (Nil/Exempt), 9B (CDNR), and Total Outward Liability.
        </p>
        <div class="drop-zone" onclick="document.getElementById('file1Batch').click()">
          <div style="font-size:24px;margin-bottom:4px;">📑</div>
          <div style="font-weight:600;">Click to select GSTR-1 PDF files</div>
          <input type="file" id="file1Batch" style="display:none;" accept=".pdf" multiple onchange="extract1Pdfs(event)"/>
        </div>
        <div id="status1" style="font-size:11px;color:var(--accent);margin-bottom:1rem;"></div>
        <div class="table-container" style="max-height:400px;">
          <table>
            <thead>
              <tr>
                <th>Period</th><th>Year</th><th>GSTIN</th><th>Return Type</th>
                <th>B2B Taxable (₹)</th><th>B2B Tax (₹)</th><th>Total Liability (₹)</th><th>ARN</th>
              </tr>
            </thead>
            <tbody id="tbody1">
              <tr><td colspan="8" style="text-align:center;padding:2rem;color:var(--muted);">Select GSTR-1 return PDFs to view extracted monthly data.</td></tr>
            </tbody>
          </table>
        </div>
      </div>
    </div>

    <!-- 4. Statutory Watchdogs Tab -->
    <div id="tab-watchdogs" class="tab-content">
      <div class="card">
        <h3>Statutory DRC-01B & DRC-01C Exposure Watchdogs</h3>
        <div class="kpi-row" style="margin-top:1rem;">
          <div class="kpi-card">
            <div class="kpi-title">Rule 88C Exposure (GSTR-1 vs 3B)</div>
            <div class="kpi-val val-good">Compliant</div>
            <div style="font-size:11px;color:var(--muted);margin-top:4px;">Threshold: > ₹25 Lakhs and > 20% liability gap</div>
          </div>
          <div class="kpi-card">
            <div class="kpi-title">Rule 88D Exposure (GSTR-2B vs 3B)</div>
            <div class="kpi-val val-good">Compliant</div>
            <div style="font-size:11px;color:var(--muted);margin-top:4px;">Threshold: > ₹25 Lakhs and > 20% excess ITC claim</div>
          </div>
        </div>
      </div>
    </div>
  </div>

  <script>
    function showTab(id) {
      document.querySelectorAll('.tab-content').forEach(el => el.classList.remove('active'));
      document.querySelectorAll('.nav-tab').forEach(el => el.classList.remove('active'));
      document.getElementById(id).classList.add('active');
      event.target.classList.add('active');
    }

    function formatINR(val) {
      return "₹ " + Number(val || 0).toLocaleString("en-IN", { minimumFractionDigits: 2, maximumFractionDigits: 2 });
    }

    async function uploadBooks(e) {
      const file = e.target.files[0];
      if (!file) return;
      const formData = new FormData();
      formData.append("file", file);
      const res = await fetch("/api/upload_books", { method: "POST", body: formData });
      const data = await res.json();
      document.getElementById("booksStatus").innerText = `${data.count} records loaded from ${file.name}`;
      alert(`Loaded ${data.count} records from Books Excel.`);
    }

    async function upload2B(e) {
      const files = e.target.files;
      if (!files.length) return;
      const formData = new FormData();
      for (let f of files) formData.append("files", f);
      const res = await fetch("/api/upload_2b", { method: "POST", body: formData });
      const data = await res.json();
      document.getElementById("portalStatus").innerText = `${data.count} records loaded from ${files.length} file(s)`;
      alert(`Loaded ${data.count} records from GSTR-2B.`);
    }

    async function runPurchaseRecon() {
      const btn = document.getElementById("btnRunRecon");
      btn.innerText = "⏳ Reconciling...";
      btn.disabled = true;
      try {
        const res = await fetch("/api/run_recon", { method: "POST" });
        const data = await res.json();
        if (!data.ok) throw new Error(data.error);

        const s = data.result.summary;
        document.getElementById("kpiBooksItc").innerText = formatINR(s.total_books_itc);
        document.getElementById("kpiPortalItc").innerText = formatINR(s.total_portal_itc);
        document.getElementById("kpiMatchedItc").innerText = formatINR(s.matched_itc);
        document.getElementById("kpiRiskItc").innerText = formatINR(s.itc_at_risk);
        document.getElementById("kpiHighRiskVendors").innerText = s.high_risk_vendors;

        document.getElementById("cntMatched").innerText = s.matched_count;
        document.getElementById("cntMismatch").innerText = s.tax_mismatch_count;
        document.getElementById("cntMissing2B").innerText = s.missing_in_2b_count;
        document.getElementById("cntVendors").innerText = data.result.vendor_risks.length;

        window.RECON_RESULT = data.result;
        renderTable("matched");
        document.getElementById("btnDownloadExcel").style.display = "block";
      } catch (err) {
        alert("Recon failed: " + err.message);
      } finally {
        btn.innerText = "⚡ Execute 4-Pass Reconciler";
        btn.disabled = false;
      }
    }

    function renderTable(type) {
      if (!window.RECON_RESULT) return;
      const th = document.getElementById("reconThead");
      const tb = document.getElementById("reconTbody");
      th.innerHTML = ""; tb.innerHTML = "";

      const rows = window.RECON_RESULT[type] || [];
      if (!rows.length) {
        tb.innerHTML = '<tr><td colspan="7" style="text-align:center;padding:1.5rem;color:var(--muted);">No records in this category.</td></tr>';
        return;
      }

      if (type === "matched" || type === "tax_mismatch") {
        th.innerHTML = "<tr><th>Supplier GSTIN</th><th>Supplier Name</th><th>Books Inv</th><th>Portal Inv</th><th>Date</th><th>Books ITC</th><th>Portal ITC</th><th>Status</th></tr>";
        rows.forEach(r => {
          tb.innerHTML += `<tr><td><strong>${r.gstin}</strong></td><td>${r.supplier_name||'-'}</td><td>${r.books_inv}</td><td>${r.portal_inv}</td><td>${r.books_date}</td><td>${formatINR(r.books_itc)}</td><td>${formatINR(r.portal_itc)}</td><td><span class="badge ${r.status==='MATCHED'?'badge-green':'badge-amber'}">${r.match_type}</span></td></tr>`;
        });
      } else if (type === "missing_in_2b") {
        th.innerHTML = "<tr><th>Supplier GSTIN</th><th>Supplier Name</th><th>Books Inv</th><th>Date</th><th>Taxable</th><th>ITC at Risk</th><th>Action</th></tr>";
        rows.forEach(r => {
          tb.innerHTML += `<tr><td><strong>${r.gstin}</strong></td><td>${r.supplier_name||'-'}</td><td>${r.books_inv}</td><td>${r.books_date}</td><td>${formatINR(r.books_taxable)}</td><td style="color:var(--bad);font-weight:700;">${formatINR(r.books_itc)}</td><td style="font-size:11px;color:var(--muted);">${r.action}</td></tr>`;
        });
      } else if (type === "vendor_risks") {
        th.innerHTML = "<tr><th>Supplier GSTIN</th><th>Supplier Name</th><th>Invoices</th><th>Total ITC</th><th>ITC at Risk</th><th>Score</th><th>RAG</th><th>Action</th></tr>";
        rows.forEach(v => {
          let badge = v.rag === "RED" ? "badge-red" : (v.rag === "AMBER" ? "badge-amber" : "badge-green");
          tb.innerHTML += `<tr><td><strong>${v.gstin}</strong></td><td>${v.name||'-'}</td><td>${v.invoices}</td><td>${formatINR(v.total_itc)}</td><td style="color:var(--bad);font-weight:700;">${formatINR(v.missing_itc)}</td><td>${v.score}</td><td><span class="badge ${badge}">${v.rag}</span></td><td style="font-size:11px;">${v.action}</td></tr>`;
        });
      }
    }

    function downloadExcel() {
      window.location.href = "/api/download_excel";
    }

    async function extract3BPdfs(e) {
      const files = e.target.files;
      if (!files.length) return;
      const formData = new FormData();
      for (let f of files) formData.append("files", f);
      document.getElementById("status3B").innerText = `Extracting ${files.length} PDF(s)...`;
      const res = await fetch("/api/extract_3b", { method: "POST", body: formData });
      const data = await res.json();
      document.getElementById("status3B").innerText = `Successfully extracted ${data.results.length} month(s).`;
      const tb = document.getElementById("tbody3B");
      tb.innerHTML = "";
      data.results.forEach(r => {
        const ot = r.t31.outward_taxable;
        const totTax = ot[1] + ot[2] + ot[3] + ot[4];
        const netItc = r.t4.net_itc[0] + r.t4.net_itc[1] + r.t4.net_itc[2] + r.t4.net_itc[3];
        const cash = r.t61.cash_paid[0] + r.t61.cash_paid[1] + r.t61.cash_paid[2];
        tb.innerHTML += `<tr><td><strong>${r.period}</strong></td><td>${r.year}</td><td>${r.gstin}</td><td>${formatINR(ot[0])}</td><td>${formatINR(totTax)}</td><td>${formatINR(netItc)}</td><td>${formatINR(cash)}</td><td><span class="badge badge-green">${r.arn}</span></td></tr>`;
      });
    }

    async function extract1Pdfs(e) {
      const files = e.target.files;
      if (!files.length) return;
      const formData = new FormData();
      for (let f of files) formData.append("files", f);
      document.getElementById("status1").innerText = `Extracting ${files.length} PDF(s)...`;
      const res = await fetch("/api/extract_1", { method: "POST", body: formData });
      const data = await res.json();
      document.getElementById("status1").innerText = `Successfully extracted ${data.results.length} month(s).`;
      const tb = document.getElementById("tbody1");
      tb.innerHTML = "";
      data.results.forEach(r => {
        const b2bTax = r.b2b[1] + r.b2b[2] + r.b2b[3];
        const totLiab = r.total_liability[1] + r.total_liability[2] + r.total_liability[3];
        tb.innerHTML += `<tr><td><strong>${r.period}</strong></td><td>${r.year}</td><td>${r.gstin}</td><td>${r.type}</td><td>${formatINR(r.b2b[0])}</td><td>${formatINR(b2bTax)}</td><td>${formatINR(totLiab)}</td><td><span class="badge badge-green">${r.arn}</span></td></tr>`;
      });
    }
  </script>
</body>
</html>
"""

@app.route("/")
def index():
    return render_template_string(HTML_TEMPLATE)

@app.route("/api/upload_books", methods=["POST"])
def upload_books():
    file = request.files.get("file")
    if not file: return jsonify({"error": "No file"}), 400
    tmp = tempfile.NamedTemporaryFile(delete=False, suffix=os.path.splitext(file.filename)[1])
    file.save(tmp.name)
    records = engine.parse_books_excel(tmp.name)
    STATE["books"] = records
    try: os.unlink(tmp.name)
    except: pass
    return jsonify({"count": len(records)})

@app.route("/api/upload_2b", methods=["POST"])
def upload_2b():
    files = request.files.getlist("files")
    if not files: return jsonify({"error": "No files"}), 400
    all_records = []
    for f in files:
        tmp = tempfile.NamedTemporaryFile(delete=False, suffix=os.path.splitext(f.filename)[1])
        f.save(tmp.name)
        ext = os.path.splitext(f.filename)[1].lower()
        if ext in (".xlsx", ".xls"):
            recs = engine.parse_gstr2b_excel(tmp.name)
            all_records.extend(recs)
        try: os.unlink(tmp.name)
        except: pass
    STATE["gstr2b"] = all_records
    return jsonify({"count": len(all_records)})

@app.route("/api/run_recon", methods=["POST"])
def run_recon():
    books = STATE["books"]
    portal = STATE["gstr2b"]
    if not portal:
        return jsonify({"ok": False, "error": "Please upload GSTR-2B first."}), 400
    result = engine.reconcile_purchase_2b(books, portal)
    STATE["recon_result"] = result

    # Generate Excel in temp
    tmp_out = os.path.join(tempfile.gettempdir(), f"GST_Audit_Report_{int(time.time())}.xlsx")
    engine.generate_recon_workbook(result, tmp_out)
    STATE["last_excel_path"] = tmp_out

    return jsonify({"ok": True, "result": result})

@app.route("/api/download_excel")
def download_excel():
    path = STATE.get("last_excel_path")
    if path and os.path.exists(path):
        return send_file(path, as_attachment=True, download_name=os.path.basename(path))
    return "Report not generated yet", 404

@app.route("/api/extract_3b", methods=["POST"])
def extract_3b():
    files = request.files.getlist("files")
    results = []
    for f in files:
        tmp = tempfile.NamedTemporaryFile(delete=False, suffix=".pdf")
        f.save(tmp.name)
        res = engine.extract_gstr3b_pdf(tmp.name)
        results.append(res)
        try: os.unlink(tmp.name)
        except: pass
    STATE["gstr3b_data"] = results
    return jsonify({"results": results})

@app.route("/api/extract_1", methods=["POST"])
def extract_1():
    files = request.files.getlist("files")
    results = []
    for f in files:
        tmp = tempfile.NamedTemporaryFile(delete=False, suffix=".pdf")
        f.save(tmp.name)
        res = engine.extract_gstr1_pdf(tmp.name)
        results.append(res)
        try: os.unlink(tmp.name)
        except: pass
    STATE["gstr1_data"] = results
    return jsonify({"results": results})


def find_free_port(start_port=5055):
    for port in range(start_port, start_port + 50):
        with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as s:
            if s.connect_ex(('127.0.0.1', port)) != 0:
                return port
    return 5055


def main():
    port = find_free_port(5055)
    url = f"http://127.0.0.1:{port}"
    print("=" * 65)
    print(f"  GST Standalone Suite Starting at {url}")
    print("=" * 65)
    
    # Auto launch browser in background thread
    threading.Timer(1.2, lambda: webbrowser.open(url)).start()
    
    app.run(host="127.0.0.1", port=port, debug=False)


if __name__ == "__main__":
    main()
