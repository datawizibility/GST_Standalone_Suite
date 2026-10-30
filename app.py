# -*- coding: utf-8 -*-
"""
GST Standalone Suite - Self-Contained Web Application
Launches on http://127.0.0.1:5055 and automatically opens the browser.
No external Node.js, database, or cloud dependencies required.
"""

import os
import time
import socket
import logging
import tempfile
import webbrowser
import threading
from datetime import datetime

from flask import Flask, request, jsonify, send_file, render_template_string

# Import core statutory engine
import engine

logging.basicConfig(level=logging.INFO, format="%(asctime)s [%(levelname)s] %(message)s")
logger = logging.getLogger(__name__)

app = Flask(__name__)
app.config['MAX_CONTENT_LENGTH'] = 100 * 1024 * 1024  # 100MB upload limit

# Thread-safe global state
STATE_LOCK = threading.Lock()
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
  <title>GST Standalone Suite — Statutory Audit & Reconciliation</title>
  <style>
    :root {
      --bg: #0f172a; --surface: #1e293b; --card: #334155; --border: #475569;
      --text: #f8fafc; --muted: #94a3b8; --accent: #38bdf8; --accent-dk: #0284c7;
      --good: #4ade80; --bad: #f87171; --warn: #fbbf24; --info: #60a5fa;
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
    .val-good { color: var(--good); } .val-bad { color: var(--bad); } .val-warn { color: var(--warn); } .val-info { color: var(--info); }
    .table-container { overflow-x: auto; max-height: 520px; border: 1px solid var(--border); border-radius: 6px; }
    table { width: 100%; border-collapse: collapse; font-size: 12px; }
    th { background: #0f172a; color: var(--muted); padding: 8px 10px; text-align: left; position: sticky; top: 0; z-index: 10; border-bottom: 1px solid var(--border); }
    td { padding: 8px 10px; border-bottom: 1px solid rgba(255,255,255,0.06); }
    tr:hover td { background: rgba(255,255,255,0.03); }
    .badge { display: inline-block; padding: 2px 6px; border-radius: 4px; font-size: 10px; font-weight: 700; }
    .badge-green { background: rgba(74,222,128,0.2); color: var(--good); border: 1px solid var(--good); }
    .badge-red { background: rgba(248,113,113,0.2); color: var(--bad); border: 1px solid var(--bad); }
    .badge-amber { background: rgba(251,191,36,0.2); color: var(--warn); border: 1px solid var(--warn); }
    .badge-info { background: rgba(96,165,250,0.2); color: var(--info); border: 1px solid var(--info); }
  </style>
</head>
<body>
  <header>
    <div class="logo">⚡ GST Standalone Intelligence Suite</div>
    <div style="font-size:12px;color:var(--muted);">
      Offline & Secure • Zero-Installation • Statutory Precision Engine
    </div>
  </header>

  <div class="layout">
    <div class="nav-tabs">
      <button class="nav-tab active" onclick="showTab('tab-purchase', this)">🛒 Purchase Recon (2B vs Books)</button>
      <button class="nav-tab" onclick="showTab('tab-gstr3b', this)">📄 GSTR-3B PDF Extractor</button>
      <button class="nav-tab" onclick="showTab('tab-gstr1', this)">📑 GSTR-1 PDF Extractor</button>
      <button class="nav-tab" onclick="showTab('tab-watchdogs', this)">⚖️ Rule 88C / 88D Watchdogs</button>
    </div>

    <!-- 1. Purchase Recon Tab -->
    <div id="tab-purchase" class="tab-content active">
      <div class="grid-2">
        <div>
          <div class="card">
            <h3>Step 1: Ingest Books Purchase Data</h3>
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
              <div style="font-size:11px;color:var(--muted);">Portal Excel (.xlsx) or JSON (Multi-Month & B2BA supported)</div>
              <input type="file" id="file2b" style="display:none;" accept=".xlsx,.xls,.json" multiple onchange="upload2B(event)"/>
            </div>
            <div id="portalStatus" style="font-size:11px;color:var(--muted);">0 portal records loaded</div>
          </div>

          <div class="card">
            <h3>Step 3: Run Statutory Reconciler</h3>
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
              <div class="kpi-title">Unclaimed ITC (Cash Adv)</div>
              <div class="kpi-val val-info" id="kpiUnclaimedItc">₹ 0.00</div>
            </div>
            <div class="kpi-card">
              <div class="kpi-title">High-Risk Vendors</div>
              <div class="kpi-val val-bad" id="kpiHighRiskVendors">0</div>
            </div>
          </div>

          <div class="card" style="padding:0.75rem;">
            <div style="display:flex;gap:8px;margin-bottom:0.75rem;flex-wrap:wrap;">
              <button class="btn btn-secondary" style="width:auto;padding:4px 10px;" onclick="renderTable('matched')">Matched (<span id="cntMatched">0</span>)</button>
              <button class="btn btn-secondary" style="width:auto;padding:4px 10px;" onclick="renderTable('tax_mismatch')">Discrepancies (<span id="cntMismatch">0</span>)</button>
              <button class="btn btn-secondary" style="width:auto;padding:4px 10px;" onclick="renderTable('missing_in_2b')">Missing in 2B (<span id="cntMissing2B">0</span>)</button>
              <button class="btn btn-secondary" style="width:auto;padding:4px 10px;color:var(--info);" onclick="renderTable('missing_in_books')">Unclaimed ITC (<span id="cntMissingBooks">0</span>)</button>
              <button class="btn btn-secondary" style="width:auto;padding:4px 10px;color:var(--bad);" onclick="renderTable('vendor_risks')">🛡️ Supplier Risk Matrix (<span id="cntVendors">0</span>)</button>
            </div>
            <div class="table-container">
              <table id="reconTable">
                <thead id="reconThead"></thead>
                <tbody id="reconTbody">
                  <tr><td colspan="8" style="text-align:center;padding:2rem;color:var(--muted);">Load Books and GSTR-2B files, then click Execute Reconciler.</td></tr>
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
          Select filed GSTR-3B return PDFs to automatically extract Table 3.1, Table 4 (Eligible ITC), Table 5, Table 6.1 (Cash Paid), and ARN status.
        </p>
        <div class="drop-zone" onclick="document.getElementById('file3bBatch').click()">
          <div style="font-size:24px;margin-bottom:4px;">📑</div>
          <div style="font-weight:600;">Click to select GSTR-3B PDF files</div>
          <input type="file" id="file3bBatch" style="display:none;" accept=".pdf" multiple onchange="extract3BPdfs(event)"/>
        </div>
        <div id="status3B" style="font-size:11px;color:var(--accent);margin-bottom:1rem;"></div>
        <div class="table-container" style="max-height:450px;">
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
        <div class="table-container" style="max-height:450px;">
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
        <p style="font-size:12px;color:var(--muted);margin-bottom:1rem;">
          Evaluates statutory risk under Section 75(12) of the CGST Act. Automatically computes tax gaps between filed GSTR-1, GSTR-3B, and GSTR-2B.
        </p>
        <div style="display:flex;gap:10px;margin-bottom:1.5rem;">
          <button class="btn btn-primary" style="width:auto;" onclick="recalculateWatchdogs()">⚡ Recompute from Loaded Returns</button>
        </div>
        <div class="kpi-row">
          <div class="kpi-card" style="border-left:4px solid var(--accent);">
            <div class="kpi-title">Rule 88C Exposure (Form DRC-01B)</div>
            <div class="kpi-val" id="wd88cStatus">Compliant</div>
            <div style="margin-top:8px;font-size:12px;" id="wd88cDetails">
              GSTR-1 Liability: <strong id="wdG1Liab">₹ 0.00</strong> | GSTR-3B Liability: <strong id="wd3bLiab">₹ 0.00</strong><br/>
              Difference Gap: <strong id="wd88cGap">₹ 0.00 (0.00%)</strong>
            </div>
            <div style="font-size:11px;color:var(--muted);margin-top:6px;" id="wd88cAction">Threshold: &gt; ₹25 Lakhs AND &gt; 20% liability gap</div>
          </div>
          <div class="kpi-card" style="border-left:4px solid var(--warn);">
            <div class="kpi-title">Rule 88D Exposure (Form DRC-01C)</div>
            <div class="kpi-val" id="wd88dStatus">Compliant</div>
            <div style="margin-top:8px;font-size:12px;" id="wd88dDetails">
              GSTR-3B Claimed ITC: <strong id="wd3bItc">₹ 0.00</strong> | GSTR-2B Available ITC: <strong id="wd2bItc">₹ 0.00</strong><br/>
              Difference Gap: <strong id="wd88dGap">₹ 0.00 (0.00%)</strong>
            </div>
            <div style="font-size:11px;color:var(--muted);margin-top:6px;" id="wd88dAction">Threshold: &gt; ₹25 Lakhs AND &gt; 20% excess ITC claim</div>
          </div>
        </div>
      </div>
    </div>
  </div>

  <script>
    function esc(str) {
      if (str === null || str === undefined) return "";
      return String(str)
        .replace(/&/g, "&amp;")
        .replace(/</g, "&lt;")
        .replace(/>/g, "&gt;")
        .replace(/"/g, "&quot;")
        .replace(/'/g, "&#39;");
    }

    function showTab(id, btn) {
      document.querySelectorAll('.tab-content').forEach(el => el.classList.remove('active'));
      document.querySelectorAll('.nav-tab').forEach(el => el.classList.remove('active'));
      document.getElementById(id).classList.add('active');
      if (btn) btn.classList.add('active');
    }

    function formatINR(val) {
      return "₹ " + Number(val || 0).toLocaleString("en-IN", { minimumFractionDigits: 2, maximumFractionDigits: 2 });
    }

    async function uploadBooks(e) {
      const file = e.target.files[0];
      if (!file) return;
      const formData = new FormData();
      formData.append("file", file);
      try {
        const res = await fetch("/api/upload_books", { method: "POST", body: formData });
        const data = await res.json();
        if (!data.ok) throw new Error(data.error);
        document.getElementById("booksStatus").innerText = `${data.count} records loaded from ${file.name}`;
        alert(`Loaded ${data.count} purchase records from Books.`);
      } catch (err) {
        alert("Books upload failed: " + err.message);
      }
    }

    async function upload2B(e) {
      const files = e.target.files;
      if (!files.length) return;
      const formData = new FormData();
      for (let f of files) formData.append("files", f);
      try {
        const res = await fetch("/api/upload_2b", { method: "POST", body: formData });
        const data = await res.json();
        if (!data.ok) throw new Error(data.error);
        document.getElementById("portalStatus").innerText = `${data.count} records loaded from ${files.length} file(s)`;
        alert(`Loaded ${data.count} portal records from GSTR-2B.`);
      } catch (err) {
        alert("GSTR-2B upload failed: " + err.message);
      }
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
        document.getElementById("kpiUnclaimedItc").innerText = formatINR(s.unclaimed_itc);
        document.getElementById("kpiHighRiskVendors").innerText = s.high_risk_vendors;

        document.getElementById("cntMatched").innerText = s.matched_count;
        document.getElementById("cntMismatch").innerText = s.tax_mismatch_count;
        document.getElementById("cntMissing2B").innerText = s.missing_in_2b_count;
        document.getElementById("cntMissingBooks").innerText = s.missing_in_books_count;
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
        tb.innerHTML = '<tr><td colspan="8" style="text-align:center;padding:1.5rem;color:var(--muted);">No records in this category.</td></tr>';
        return;
      }

      if (type === "matched" || type === "tax_mismatch") {
        th.innerHTML = "<tr><th>Supplier GSTIN</th><th>Supplier Name</th><th>Books Inv</th><th>Portal Inv</th><th>Date</th><th>Books ITC</th><th>Portal ITC</th><th>Status</th></tr>";
        rows.forEach(r => {
          const badgeClass = r.status === "MATCHED" ? "badge-green" : "badge-amber";
          tb.innerHTML += `<tr><td><strong>${esc(r.gstin)}</strong></td><td>${esc(r.supplier_name)||'-'}</td><td>${esc(r.books_inv)}</td><td>${esc(r.portal_inv)}</td><td>${esc(r.books_date)}</td><td>${formatINR(r.books_itc)}</td><td>${formatINR(r.portal_itc)}</td><td><span class="badge ${badgeClass}">${esc(r.match_type)}</span></td></tr>`;
        });
      } else if (type === "missing_in_2b") {
        th.innerHTML = "<tr><th>Supplier GSTIN</th><th>Supplier Name</th><th>Books Inv</th><th>Date</th><th>Taxable</th><th>ITC at Risk</th><th>Action</th></tr>";
        rows.forEach(r => {
          tb.innerHTML += `<tr><td><strong>${esc(r.gstin)}</strong></td><td>${esc(r.supplier_name)||'-'}</td><td>${esc(r.books_inv)}</td><td>${esc(r.books_date)}</td><td>${formatINR(r.books_taxable)}</td><td style="color:var(--bad);font-weight:700;">${formatINR(r.books_itc)}</td><td style="font-size:11px;color:var(--muted);">${esc(r.action)}</td></tr>`;
        });
      } else if (type === "missing_in_books") {
        th.innerHTML = "<tr><th>Supplier GSTIN</th><th>Supplier Name</th><th>Portal Inv</th><th>Date</th><th>Taxable</th><th>Unclaimed ITC</th><th>Commercial Action</th></tr>";
        rows.forEach(r => {
          tb.innerHTML += `<tr><td><strong>${esc(r.gstin)}</strong></td><td>${esc(r.supplier_name)||'-'}</td><td>${esc(r.portal_inv)}</td><td>${esc(r.portal_date)}</td><td>${formatINR(r.portal_taxable)}</td><td style="color:var(--info);font-weight:700;">${formatINR(r.portal_itc)}</td><td style="font-size:11px;color:var(--muted);">${esc(r.action)}</td></tr>`;
        });
      } else if (type === "vendor_risks") {
        th.innerHTML = "<tr><th>Supplier GSTIN</th><th>Supplier Name</th><th>Invoices</th><th>Total ITC</th><th>ITC at Risk</th><th>Score</th><th>RAG</th><th>Action</th></tr>";
        rows.forEach(v => {
          const badge = v.rag === "RED" ? "badge-red" : (v.rag === "AMBER" ? "badge-amber" : "badge-green");
          tb.innerHTML += `<tr><td><strong>${esc(v.gstin)}</strong></td><td>${esc(v.name)||'-'}</td><td>${v.invoices}</td><td>${formatINR(v.total_itc)}</td><td style="color:var(--bad);font-weight:700;">${formatINR(v.missing_itc)}</td><td>${v.score}</td><td><span class="badge ${badge}">${esc(v.rag)}</span></td><td style="font-size:11px;">${esc(v.action)}</td></tr>`;
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
      try {
        const res = await fetch("/api/extract_3b", { method: "POST", body: formData });
        const data = await res.json();
        if (!data.ok) throw new Error(data.error);
        document.getElementById("status3B").innerText = `Successfully extracted ${data.results.length} month(s).`;
        const tb = document.getElementById("tbody3B");
        tb.innerHTML = "";
        data.results.forEach(r => {
          const ot = r.t31.outward_taxable;
          const totTax = ot[1] + ot[2] + ot[3] + ot[4];
          const netItc = r.t4.net_itc[0] + r.t4.net_itc[1] + r.t4.net_itc[2] + r.t4.net_itc[3];
          const cash = r.t61.cash_paid[0] + r.t61.cash_paid[1] + r.t61.cash_paid[2] + r.t61.cash_paid[3];
          tb.innerHTML += `<tr><td><strong>${esc(r.period)}</strong></td><td>${esc(r.year)}</td><td>${esc(r.gstin)}</td><td>${formatINR(ot[0])}</td><td>${formatINR(totTax)}</td><td>${formatINR(netItc)}</td><td>${formatINR(cash)}</td><td><span class="badge badge-green">${esc(r.arn)}</span></td></tr>`;
        });
        recalculateWatchdogs();
      } catch (err) {
        document.getElementById("status3B").innerText = "Error: " + err.message;
        alert("GSTR-3B extraction failed: " + err.message);
      }
    }

    async function extract1Pdfs(e) {
      const files = e.target.files;
      if (!files.length) return;
      const formData = new FormData();
      for (let f of files) formData.append("files", f);
      document.getElementById("status1").innerText = `Extracting ${files.length} PDF(s)...`;
      try {
        const res = await fetch("/api/extract_1", { method: "POST", body: formData });
        const data = await res.json();
        if (!data.ok) throw new Error(data.error);
        document.getElementById("status1").innerText = `Successfully extracted ${data.results.length} month(s).`;
        const tb = document.getElementById("tbody1");
        tb.innerHTML = "";
        data.results.forEach(r => {
          const b2bTax = r.b2b[1] + r.b2b[2] + r.b2b[3];
          const totLiab = r.total_liability[1] + r.total_liability[2] + r.total_liability[3];
          tb.innerHTML += `<tr><td><strong>${esc(r.period)}</strong></td><td>${esc(r.year)}</td><td>${esc(r.gstin)}</td><td>${esc(r.type)}</td><td>${formatINR(r.b2b[0])}</td><td>${formatINR(b2bTax)}</td><td>${formatINR(totLiab)}</td><td><span class="badge badge-green">${esc(r.arn)}</span></td></tr>`;
        });
        recalculateWatchdogs();
      } catch (err) {
        document.getElementById("status1").innerText = "Error: " + err.message;
        alert("GSTR-1 extraction failed: " + err.message);
      }
    }

    async function recalculateWatchdogs() {
      try {
        const res = await fetch("/api/watchdogs");
        const data = await res.json();
        if (!data.ok) return;

        const w88c = data.rule_88c;
        const w88d = data.rule_88d;

        const el88c = document.getElementById("wd88cStatus");
        el88c.innerText = w88c.status === "COMPLIANT" ? "Compliant" : "DRC-01B Triggered (High Exposure)";
        el88c.className = "kpi-val " + (w88c.status === "COMPLIANT" ? "val-good" : "val-bad");
        document.getElementById("wdG1Liab").innerText = formatINR(w88c.gstr1_liability);
        document.getElementById("wd3bLiab").innerText = formatINR(w88c.gstr3b_liability);
        document.getElementById("wd88cGap").innerText = `${formatINR(w88c.gap)} (${w88c.gap_percentage}%)`;
        document.getElementById("wd88cAction").innerText = w88c.action;

        const el88d = document.getElementById("wd88dStatus");
        el88d.innerText = w88d.status === "COMPLIANT" ? "Compliant" : "DRC-01C Triggered (High Exposure)";
        el88d.className = "kpi-val " + (w88d.status === "COMPLIANT" ? "val-good" : "val-bad");
        document.getElementById("wd3bItc").innerText = formatINR(w88d.gstr3b_itc);
        document.getElementById("wd2bItc").innerText = formatINR(w88d.gstr2b_itc);
        document.getElementById("wd88dGap").innerText = `${formatINR(w88d.gap)} (${w88d.gap_percentage}%)`;
        document.getElementById("wd88dAction").innerText = w88d.action;
      } catch (err) {
        console.error("Watchdog update failed:", err);
      }
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
    if not file:
        return jsonify({"ok": False, "error": "No file uploaded"}), 400
    tmp = tempfile.NamedTemporaryFile(delete=False, suffix=os.path.splitext(file.filename)[1])
    try:
        file.save(tmp.name)
        records = engine.parse_books_excel(tmp.name)
        with STATE_LOCK:
            STATE["books"] = records
        return jsonify({"ok": True, "count": len(records)})
    except Exception as e:
        logger.exception("Failed to parse books")
        return jsonify({"ok": False, "error": str(e)}), 500
    finally:
        try:
            os.unlink(tmp.name)
        except Exception:
            pass

@app.route("/api/upload_2b", methods=["POST"])
def upload_2b():
    files = request.files.getlist("files")
    if not files:
        return jsonify({"ok": False, "error": "No files uploaded"}), 400
    all_records = []
    for f in files:
        ext = os.path.splitext(f.filename)[1].lower()
        tmp = tempfile.NamedTemporaryFile(delete=False, suffix=ext)
        try:
            f.save(tmp.name)
            if ext in (".xlsx", ".xls"):
                recs = engine.parse_gstr2b_excel(tmp.name)
                all_records.extend(recs)
            elif ext == ".json":
                recs = engine.parse_gstr2b_json(tmp.name)
                all_records.extend(recs)
            else:
                logger.warning(f"Unsupported format {ext} for {f.filename}")
        except Exception as e:
            logger.exception(f"Failed to parse 2B file {f.filename}")
            return jsonify({"ok": False, "error": f"Error in {f.filename}: {str(e)}"}), 500
        finally:
            try:
                os.unlink(tmp.name)
            except Exception:
                pass
    with STATE_LOCK:
        STATE["gstr2b"] = all_records
    return jsonify({"ok": True, "count": len(all_records)})

@app.route("/api/run_recon", methods=["POST"])
def run_recon():
    try:
        with STATE_LOCK:
            books = list(STATE["books"])
            portal = list(STATE["gstr2b"])

        if not portal:
            return jsonify({"ok": False, "error": "Please upload at least one GSTR-2B file first."}), 400
        if not books:
            return jsonify({"ok": False, "error": "Please upload Books purchase register first."}), 400

        result = engine.reconcile_purchase_2b(books, portal)

        # Generate openpyxl workbook
        tmp_out = os.path.join(tempfile.gettempdir(), f"GST_Audit_Report_{int(time.time())}.xlsx")
        engine.generate_recon_workbook(result, tmp_out)

        with STATE_LOCK:
            STATE["recon_result"] = result
            STATE["last_excel_path"] = tmp_out

        return jsonify({"ok": True, "result": result})
    except Exception as e:
        logger.exception("Reconciliation failed")
        return jsonify({"ok": False, "error": str(e)}), 500

@app.route("/api/download_excel")
def download_excel():
    with STATE_LOCK:
        path = STATE.get("last_excel_path")
    if not path or not os.path.exists(path):
        return "No report generated yet.", 404
    return send_file(path, as_attachment=True, download_name="GST_Purchase_Recon_Audit_Report.xlsx")

@app.route("/api/extract_3b", methods=["POST"])
def extract_3b():
    files = request.files.getlist("files")
    if not files:
        return jsonify({"ok": False, "error": "No PDF files uploaded"}), 400
    results = []
    for f in files:
        tmp = tempfile.NamedTemporaryFile(delete=False, suffix=".pdf")
        try:
            f.save(tmp.name)
            data = engine.extract_gstr3b_pdf(tmp.name)
            results.append(data)
        except Exception as e:
            logger.exception(f"Failed to extract 3B from {f.filename}")
            return jsonify({"ok": False, "error": f"Error in {f.filename}: {str(e)}"}), 500
        finally:
            try:
                os.unlink(tmp.name)
            except Exception:
                pass
    with STATE_LOCK:
        STATE["gstr3b_data"] = results
    return jsonify({"ok": True, "results": results})

@app.route("/api/extract_1", methods=["POST"])
def extract_1():
    files = request.files.getlist("files")
    if not files:
        return jsonify({"ok": False, "error": "No PDF files uploaded"}), 400
    results = []
    for f in files:
        tmp = tempfile.NamedTemporaryFile(delete=False, suffix=".pdf")
        try:
            f.save(tmp.name)
            data = engine.extract_gstr1_pdf(tmp.name)
            results.append(data)
        except Exception as e:
            logger.exception(f"Failed to extract GSTR-1 from {f.filename}")
            return jsonify({"ok": False, "error": f"Error in {f.filename}: {str(e)}"}), 500
        finally:
            try:
                os.unlink(tmp.name)
            except Exception:
                pass
    with STATE_LOCK:
        STATE["gstr1_data"] = results
    return jsonify({"ok": True, "results": results})

@app.route("/api/watchdogs")
def get_watchdogs():
    with STATE_LOCK:
        g3b = STATE.get("gstr3b_data", [])
        g1 = STATE.get("gstr1_data", [])
        recon = STATE.get("recon_result")
        portal = STATE.get("gstr2b", [])

    tot_g1_liab = sum(r["total_liability"][1] + r["total_liability"][2] + r["total_liability"][3] for r in g1)
    tot_g3b_liab = sum(r["t31"]["outward_taxable"][1] + r["t31"]["outward_taxable"][2] + r["t31"]["outward_taxable"][3] for r in g3b)
    tot_g3b_itc = sum(r["t4"]["net_itc"][0] + r["t4"]["net_itc"][1] + r["t4"]["net_itc"][2] + r["t4"]["net_itc"][3] for r in g3b)
    tot_2b_itc = recon["summary"]["total_portal_itc"] if recon else sum(p.get("total_itc", 0.0) for p in portal)

    watchdog_res = engine.compute_statutory_watchdogs(tot_g1_liab, tot_g3b_liab, tot_g3b_itc, tot_2b_itc)
    return jsonify({"ok": True, **watchdog_res})


def is_port_in_use(port):
    with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as s:
        return s.connect_ex(('127.0.0.1', port)) == 0


def open_browser():
    time.sleep(1.2)
    webbrowser.open("http://127.0.0.1:5055")


if __name__ == "__main__":
    port = 5055
    print("=" * 70)
    print(" ⚡ GST STANDALONE SUITE — STATUTORY AUDIT & RECONCILIATION ENGINE")
    print("=" * 70)
    print(f" Web Cockpit running at: http://127.0.0.1:{port}")
    print(" Press Ctrl+C to terminate.")
    print("=" * 70)

    threading.Thread(target=open_browser, daemon=True).start()
    app.run(host="127.0.0.1", port=port, debug=False, threaded=True)
