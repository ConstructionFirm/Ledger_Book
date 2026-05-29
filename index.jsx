<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8" />
<meta name="viewport" content="width=device-width, initial-scale=1.0"/>
<title>Account Ledger</title>
<script src="https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js"></script>

<!-- ── SUPABASE CONFIG ──────────────────────────────────────────────────────
     Fill these in when you're ready to go cloud-based.
     Then import { createClient } from '@supabase/supabase-js' and connect.
─────────────────────────────────────────────────────────────────────────── -->
<script>
  const SUPABASE_URL      = "https://enfdzpmefhzharrycguz.supabase.co";   // e.g. "https://xyzxyz.supabase.co"
  const SUPABASE_ANON_KEY = "eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6ImVuZmR6cG1lZmh6aGFycnljZ3V6Iiwicm9sZSI6ImFub24iLCJpYXQiOjE3Nzg5MjI2MzIsImV4cCI6MjA5NDQ5ODYzMn0.LzsQtCopd8b_ueHmEu6S05NBHN62Ya2Ek6gcuPAJdPw";   // your project's anon/public key
</script>

<style>
  *, *::before, *::after { box-sizing: border-box; margin: 0; padding: 0; }

  :root {
    --bg:        #070d1a;
    --bg2:       #0b1726;
    --bg3:       #09131f;
    --bg4:       #0d1b2e;
    --border:    #1a3050;
    --border2:   #1e3a5f;
    --text:      #e2e8f0;
    --muted:     #64748b;
    --muted2:    #475569;
    --green:     #4ade80;
    --red:       #f87171;
    --blue:      #93c5fd;
    --blue2:     #3b82f6;
    --blue3:     #60a5fa;
    --blue4:     #7dd3fc;
    --indigo:    #a5b4fc;
  }

  body {
    background: var(--bg);
    color: var(--text);
    font-family: Georgia, 'Times New Roman', serif;
    min-height: 100vh;
  }

  button { font-family: inherit; cursor: pointer; }
  input  { font-family: Georgia, serif; }

  /* ── TOAST ─────────────────────────────────────────── */
  #toast {
    position: fixed; top: 18px; right: 18px; z-index: 9999;
    padding: 11px 18px; border-radius: 8px; color: #fff;
    font-family: 'Courier New', monospace; font-size: 13px; font-weight: 600;
    box-shadow: 0 4px 24px rgba(0,0,0,.5);
    display: none; transition: opacity .3s;
  }

  /* ── MODAL ─────────────────────────────────────────── */
  .overlay {
    position: fixed; inset: 0; background: rgba(0,0,0,.72);
    display: flex; align-items: center; justify-content: center;
    z-index: 1000; backdrop-filter: blur(4px);
  }
  .modal {
    background: var(--bg4); border: 1px solid var(--border2);
    border-radius: 14px; width: 360px; padding: 24px;
    box-shadow: 0 20px 60px rgba(0,0,0,.6);
  }
  .modal-header {
    display: flex; justify-content: space-between; align-items: center;
    margin-bottom: 20px;
  }
  .modal-title { font-size: 16px; font-weight: 700; color: #e0f2fe; }
  .modal-close {
    background: none; border: none; color: var(--muted);
    font-size: 18px; line-height: 1; padding: 2px 6px;
  }
  .modal-close:hover { color: var(--text); }
  .modal-btns { display: flex; gap: 10px; justify-content: flex-end; margin-top: 20px; }
  .hidden { display: none !important; }

  /* ── HEADER ────────────────────────────────────────── */
  header {
    background: linear-gradient(135deg, #07121f 0%, #0c1e3b 100%);
    border-bottom: 1px solid var(--border);
    padding: 18px 28px;
    display: flex; align-items: center; justify-content: space-between;
    flex-wrap: wrap; gap: 12px;
  }
  .header-left { display: flex; align-items: center; gap: 14px; }
  .logo {
    width: 42px; height: 42px; border-radius: 10px;
    background: linear-gradient(135deg, #1d4ed8, #3b82f6);
    display: flex; align-items: center; justify-content: center;
    font-size: 22px; color: #bfdbfe; flex-shrink: 0;
  }
  .app-title { font-size: 21px; font-weight: 700; color: #e0f2fe; letter-spacing: .4px; }
  .app-sub { font-size: 11px; color: var(--muted); letter-spacing: 1.5px; text-transform: uppercase; }
  .balance-chip {
    background: rgba(59,130,246,.1); border: 1px solid rgba(59,130,246,.25);
    border-radius: 12px; padding: 10px 18px; text-align: center;
  }
  .bal-label { display: block; font-size: 10px; color: var(--muted); letter-spacing: 1.2px; text-transform: uppercase; }
  .bal-amt { font-size: 19px; font-weight: 700; }

  /* ── TAB BAR ───────────────────────────────────────── */
  #tab-bar {
    background: var(--bg); border-bottom: 1px solid var(--border);
    display: flex; align-items: center;
    padding: 0 28px; gap: 2px;
    position: sticky; top: 0; z-index: 50;
    overflow-x: auto; white-space: nowrap;
  }
  .tab {
    display: inline-flex; align-items: center; gap: 8px;
    padding: 11px 14px; cursor: pointer;
    font-size: 13px; color: var(--muted);
    border-bottom: 2px solid transparent;
    user-select: none; transition: color .2s; flex-shrink: 0;
  }
  .tab.active {
    color: var(--blue4); border-bottom: 2px solid var(--blue2);
    background: rgba(59,130,246,.06);
  }
  .tab-actions { display: inline-flex; gap: 4px; margin-left: 2px; }
  .tab-btn {
    padding: 1px 5px; border-radius: 4px; font-size: 11px;
    color: #94a3b8; background: rgba(255,255,255,.04);
    border: none;
  }
  .tab-btn:hover { color: var(--text); }
  .tab-btn.del { color: var(--red); }
  #add-tab-btn {
    margin-left: 8px; padding: 7px 14px; flex-shrink: 0;
    background: rgba(59,130,246,.1); border: 1px dashed rgba(59,130,246,.35);
    border-radius: 8px; color: var(--blue3); font-size: 12px;
    letter-spacing: .4px;
  }

  /* ── TOOLBAR ───────────────────────────────────────── */
  #toolbar {
    background: var(--bg3); border-bottom: 1px solid var(--border);
    padding: 10px 28px; display: flex; align-items: center;
    justify-content: space-between; gap: 12px; flex-wrap: wrap;
  }
  .toolbar-left { display: flex; align-items: center; gap: 10px; flex-wrap: wrap; }
  .toolbar-right { display: flex; align-items: center; gap: 10px; }
  .toolbar-label { font-size: 11px; color: var(--muted2); letter-spacing: 1px; text-transform: uppercase; }

  /* ── BUTTONS ───────────────────────────────────────── */
  .btn-export-active {
    background: linear-gradient(135deg, #1d4ed8, #2563eb); border: none;
    border-radius: 8px; color: #fff; padding: 8px 18px;
    font-size: 13px; font-weight: 600; letter-spacing: .3px;
  }
  .btn-export-all {
    background: rgba(37,99,235,.12); border: 1px solid rgba(37,99,235,.35);
    border-radius: 8px; color: var(--blue3); padding: 8px 18px;
    font-size: 13px; letter-spacing: .3px;
  }
  .btn-import {
    background: rgba(99,102,241,.12); border: 1px solid rgba(99,102,241,.35);
    border-radius: 8px; color: var(--indigo); padding: 8px 18px;
    font-size: 13px; letter-spacing: .3px;
  }
  .btn-green {
    background: linear-gradient(135deg, #059669, #10b981); border: none;
    border-radius: 8px; color: #fff; padding: 8px 22px;
    font-size: 14px; font-weight: 700; height: 38px;
  }
  .btn-cancel {
    background: none; border: 1px solid var(--border2);
    border-radius: 8px; color: var(--muted); padding: 8px 14px;
    font-size: 13px; height: 38px;
  }
  .btn-blue {
    background: #1d4ed8; border: none; border-radius: 6px;
    color: #fff; padding: 6px 12px; font-size: 13px;
  }
  .btn-blue-sm {
    background: none; border: 1px solid var(--border2); border-radius: 6px;
    color: var(--muted); padding: 6px 12px; font-size: 13px;
  }
  .btn-red {
    background: #dc2626; border: none; border-radius: 8px;
    color: #fff; padding: 8px 22px; font-size: 14px; font-weight: 700; height: 38px;
  }

  /* ── BODY ──────────────────────────────────────────── */
  #body { padding: 22px 28px; max-width: 1380px; margin: 0 auto; }

  /* ── TOP ROW ───────────────────────────────────────── */
  .top-row { display: flex; gap: 16px; align-items: stretch; margin-bottom: 18px; flex-wrap: wrap; }

  .opening-card {
    background: var(--bg4); border: 1px solid var(--border2);
    border-radius: 12px; padding: 16px 20px; min-width: 230px;
  }
  .opening-lbl {
    font-size: 10px; color: var(--muted); letter-spacing: 1.2px;
    text-transform: uppercase; margin-bottom: 10px; display: block;
  }
  #opening-display {
    background: none; border: 1px dashed var(--border2); border-radius: 8px;
    color: var(--blue); padding: 8px 14px; font-size: 15px;
    width: 100%; text-align: left; font-family: Georgia, serif;
  }
  #opening-edit-row { display: none; gap: 8px; align-items: center; }
  #opening-input {
    background: var(--bg); border: 1px solid var(--border2); border-radius: 8px;
    color: var(--text); padding: 6px 10px; font-size: 14px;
    width: 140px; outline: none;
  }

  /* ── CARDS ─────────────────────────────────────────── */
  #cards {
    display: grid; grid-template-columns: repeat(auto-fit, minmax(150px,1fr));
    gap: 12px; flex: 1;
  }
  .card {
    background: var(--bg4); border: 1px solid var(--border);
    border-radius: 12px; padding: 14px 18px;
    display: flex; flex-direction: column; gap: 4px;
  }
  .card-icon { font-size: 18px; }
  .card-lbl { font-size: 10px; color: var(--muted); letter-spacing: 1px; text-transform: uppercase; }
  .card-val { font-size: 17px; font-weight: 700; font-variant-numeric: tabular-nums; }

  /* ── FORM ──────────────────────────────────────────── */
  #form-card {
    background: var(--bg2); border: 1px solid var(--border);
    border-radius: 12px; padding: 18px 22px; margin-bottom: 16px;
  }
  #form-title { font-size: 11px; color: var(--muted2); letter-spacing: 1px; text-transform: uppercase; margin-bottom: 14px; }
  .form-row { display: flex; gap: 10px; flex-wrap: wrap; align-items: flex-end; }
  .field { display: flex; flex-direction: column; flex: 1; min-width: 120px; }
  .field-wide { flex: 2; }
  .field label { font-size: 10px; color: var(--muted); letter-spacing: 1px; text-transform: uppercase; margin-bottom: 5px; display: block; }
  .field label.green { color: var(--green); }
  .field label.red   { color: var(--red); }
  .field-btns { display: flex; gap: 8px; align-items: flex-end; }

  input[type="text"], input[type="number"], input[type="date"], input[type="search"] {
    background: var(--bg); border: 1px solid var(--border2); border-radius: 8px;
    color: var(--text); padding: 8px 12px; font-size: 14px; outline: none;
    width: 100%;
  }
  input[type="text"]:focus, input[type="number"]:focus, input[type="date"]:focus { border-color: var(--blue2); }

  /* ── SEARCH ────────────────────────────────────────── */
  .search-row { display: flex; align-items: center; gap: 10px; margin-bottom: 12px; }
  #search-input { flex: 1; background: var(--bg2); border: 1px solid var(--border); }
  #clear-search { background: none; border: none; color: var(--muted2); font-size: 14px; }
  #row-count { font-size: 12px; color: #334155; white-space: nowrap; }

  /* ── TABLE ─────────────────────────────────────────── */
  #table-wrap { background: var(--bg3); border: 1px solid var(--border); border-radius: 12px; overflow: hidden; }
  table { width: 100%; border-collapse: collapse; }
  thead th {
    padding: 11px 16px; text-align: left;
    font-size: 10px; letter-spacing: 1.2px; text-transform: uppercase;
    color: var(--muted2); background: var(--bg2); border-bottom: 1px solid var(--border);
    font-weight: 600;
  }
  tbody tr { transition: background .15s; }
  tbody tr:nth-child(even) { background: #0e1a2b; }
  tbody tr:nth-child(odd)  { background: #0b1523; }
  tbody tr.editing-row     { background: #162032 !important; }
  td {
    padding: 10px 16px; font-size: 13px; color: #94a3b8;
    border-bottom: 1px solid var(--bg4);
    font-variant-numeric: tabular-nums;
  }
  td.num  { font-weight: 500; }
  td.green { color: var(--green); }
  td.red   { color: var(--red); }
  td.blue  { color: var(--blue); font-weight: 600; }
  td.dimmed { color: #334155; font-size: 11px; }
  td.bold  { font-weight: 500; }

  tfoot td {
    background: var(--bg4); border-top: 1px solid var(--border);
    font-weight: 700; font-size: 13px;
  }
  tfoot .tfoot-label { color: var(--muted2); text-align: right; letter-spacing: 1px; font-size: 11px; }

  .btn-edit {
    background: rgba(59,130,246,.12); border: 1px solid rgba(59,130,246,.25);
    border-radius: 6px; color: var(--blue); padding: 3px 8px; font-size: 12px;
  }
  .btn-del {
    background: rgba(248,113,113,.08); border: 1px solid rgba(248,113,113,.25);
    border-radius: 6px; color: var(--red); padding: 3px 8px; font-size: 11px;
  }

  /* ── EMPTY ─────────────────────────────────────────── */
  #empty-state { padding: 56px 20px; text-align: center; }
  .empty-icon  { font-size: 44px; margin-bottom: 10px; }
  .empty-text  { font-size: 17px; color: var(--border2); margin-bottom: 6px; }
  .empty-hint  { font-size: 12px; color: var(--border); }
</style>
</head>
<body>

<!-- TOAST -->
<div id="toast"></div>

<!-- MODAL CONTAINER -->
<div id="modal-container"></div>

<!-- HEADER -->
<header>
  <div class="header-left">
    <div class="logo">⬡</div>
    <div>
      <div class="app-title">Account Ledger</div>
      <div class="app-sub">Multi-Sheet Balance Tracker</div>
    </div>
  </div>
  <div class="balance-chip">
    <span class="bal-label" id="bal-sheet-name">Balance · Personal</span>
    <span class="bal-amt" id="bal-amount">₹0.00</span>
  </div>
</header>

<!-- TAB BAR -->
<div id="tab-bar">
  <!-- tabs injected by JS -->
  <button id="add-tab-btn">+ New Sheet</button>
</div>

<!-- TOOLBAR -->
<div id="toolbar">
  <div class="toolbar-left">
    <span class="toolbar-label">Export:</span>
    <button class="btn-export-active" onclick="handleExportActive()">↓ This Sheet</button>
    <button class="btn-export-all"    onclick="handleExportAll()">↓ All Sheets</button>
  </div>
  <div class="toolbar-right">
    <button class="btn-import" onclick="document.getElementById('file-input').click()">↑ Import Excel</button>
    <input id="file-input" type="file" accept=".xlsx,.xls" style="display:none" onchange="handleImport(event)" />
  </div>
</div>

<!-- BODY -->
<div id="body">

  <!-- TOP ROW -->
  <div class="top-row">
    <div class="opening-card">
      <span class="opening-lbl">Opening Balance</span>
      <button id="opening-display" onclick="startEditOpening()">₹0.00 <span style="font-size:11px;opacity:.5">✎ edit</span></button>
      <div id="opening-edit-row">
        <input id="opening-input" type="number" placeholder="0.00" />
        <button class="btn-blue" onclick="confirmOpening()">Set</button>
        <button class="btn-blue-sm" onclick="cancelOpening()">✕</button>
      </div>
    </div>
    <div id="cards">
      <div class="card">
        <span class="card-icon" style="color:var(--green)">↑</span>
        <span class="card-lbl">Total Credits</span>
        <span class="card-val" id="card-credit" style="color:var(--green)">₹0.00</span>
      </div>
      <div class="card">
        <span class="card-icon" style="color:var(--red)">↓</span>
        <span class="card-lbl">Total Debits</span>
        <span class="card-val" id="card-debit" style="color:var(--red)">₹0.00</span>
      </div>
      <div class="card">
        <span class="card-icon" id="card-net-icon" style="color:var(--green)">≈</span>
        <span class="card-lbl">Net Change</span>
        <span class="card-val" id="card-net" style="color:var(--green)">₹0.00</span>
      </div>
      <div class="card">
        <span class="card-icon" style="color:var(--blue)">#</span>
        <span class="card-lbl">Entries</span>
        <span class="card-val" id="card-entries" style="color:var(--blue)">0</span>
      </div>
    </div>
  </div>

  <!-- FORM -->
  <div id="form-card">
    <div id="form-title">+ ADD TRANSACTION · <span id="form-sheet-name" style="color:#334155;font-size:11px"></span></div>
    <div class="form-row">
      <div class="field">
        <label>Date</label>
        <input type="date" id="f-date" />
      </div>
      <div class="field field-wide">
        <label>Name / Description</label>
        <input type="text" id="f-name" placeholder="e.g. Salary, Rent…" />
      </div>
      <div class="field">
        <label class="green">Credit (+)</label>
        <input type="number" id="f-credit" min="0" placeholder="0.00" oninput="clearDebit()" />
      </div>
      <div class="field">
        <label class="red">Debit (−)</label>
        <input type="number" id="f-debit" min="0" placeholder="0.00" oninput="clearCredit()" />
      </div>
      <div class="field-btns">
        <button class="btn-cancel hidden" id="btn-cancel" onclick="cancelEdit()">Cancel</button>
        <button class="btn-green" id="btn-add" onclick="handleAdd()">Add</button>
      </div>
    </div>
  </div>

  <!-- SEARCH -->
  <div class="search-row">
    <input type="search" id="search-input" placeholder="🔍  Search transactions in this sheet…" oninput="renderTable()" />
    <button id="clear-search" onclick="clearSearch()">✕</button>
    <span id="row-count">0 entries</span>
  </div>

  <!-- TABLE -->
  <div id="table-wrap">
    <div id="empty-state">
      <div class="empty-icon">📒</div>
      <div class="empty-text" id="empty-text">No transactions yet</div>
      <div class="empty-hint">Add one above or import an Excel file</div>
    </div>
    <table id="ledger-table" class="hidden">
      <thead>
        <tr>
          <th>#</th><th>Date</th><th>Name / Description</th>
          <th>Credit (+)</th><th>Debit (−)</th><th>Balance</th><th>Actions</th>
        </tr>
      </thead>
      <tbody id="tbody"></tbody>
      <tfoot>
        <tr>
          <td colspan="3" class="tfoot-label">TOTALS</td>
          <td class="green" id="tfoot-credit"></td>
          <td class="red"   id="tfoot-debit"></td>
          <td id="tfoot-bal"></td>
          <td></td>
        </tr>
      </tfoot>
    </table>
  </div>

</div><!-- /body -->

<script>
// ── helpers ───────────────────────────────────────────────────────────────────
const uid = () => Date.now().toString(36) + Math.random().toString(36).slice(2);

function fmt(val) {
  if (val === "" || val === null || val === undefined || isNaN(Number(val))) return "—";
  return new Intl.NumberFormat("en-IN", { style: "currency", currency: "INR", minimumFractionDigits: 2 }).format(val);
}

function fmtDate(d) {
  if (!d) return "";
  const dt = new Date(d);
  if (isNaN(dt)) return d;
  return dt.toLocaleDateString("en-GB", { day: "2-digit", month: "short", year: "numeric" });
}

function today() { return new Date().toISOString().split("T")[0]; }

// ── state ─────────────────────────────────────────────────────────────────────
let sheets   = [{ id: uid(), name: "Personal", openingBalance: 0, transactions: [] }];
let activeId = sheets[0].id;
let editingId = null;
let search = "";

function activeSheet() { return sheets.find(s => s.id === activeId) || sheets[0]; }

function updateSheet(id, fn) {
  sheets = sheets.map(s => s.id === id ? { ...s, ...fn(s) } : s);
}

// ── toast ─────────────────────────────────────────────────────────────────────
let toastTimer;
function showToast(msg, type = "success") {
  const el = document.getElementById("toast");
  el.textContent = (type === "error" ? "✕ " : "✓ ") + msg;
  el.style.background = type === "error" ? "#dc2626" : "#059669";
  el.style.display = "block";
  clearTimeout(toastTimer);
  toastTimer = setTimeout(() => el.style.display = "none", 2800);
}

// ── modal ─────────────────────────────────────────────────────────────────────
function openModal(html) {
  const c = document.getElementById("modal-container");
  c.innerHTML = `<div class="overlay" onclick="closeModal(event)">${html}</div>`;
}
function closeModal(e) {
  if (e && e.target !== e.currentTarget) return;
  document.getElementById("modal-container").innerHTML = "";
}
function closeModalForce() {
  document.getElementById("modal-container").innerHTML = "";
}

// ── SHEET MANAGEMENT ──────────────────────────────────────────────────────────
document.getElementById("add-tab-btn").onclick = () => {
  openModal(`
    <div class="modal" onclick="event.stopPropagation()">
      <div class="modal-header">
        <span class="modal-title">New Sheet</span>
        <button class="modal-close" onclick="closeModalForce()">✕</button>
      </div>
      <div>
        <div class="field"><label>Sheet Name</label>
          <input type="text" id="new-sheet-name" placeholder="e.g. John, Business, Savings…"
            onkeydown="if(event.key==='Enter')doAddSheet()" autofocus />
        </div>
        <div class="modal-btns">
          <button class="btn-cancel" onclick="closeModalForce()">Cancel</button>
          <button class="btn-green" onclick="doAddSheet()">Create Sheet</button>
        </div>
      </div>
    </div>
  `);
  setTimeout(() => document.getElementById("new-sheet-name")?.focus(), 50);
};

function doAddSheet() {
  const name = (document.getElementById("new-sheet-name")?.value || "").trim();
  if (!name) return;
  const s = { id: uid(), name, openingBalance: 0, transactions: [] };
  sheets.push(s);
  activeId = s.id;
  editingId = null;
  closeModalForce();
  resetForm();
  renderAll();
  showToast(`Sheet "${name}" created`);
}

function doRenameSheet(id) {
  const name = (document.getElementById("rename-input")?.value || "").trim();
  if (!name) return;
  updateSheet(id, () => ({ name }));
  closeModalForce();
  renderAll();
  showToast("Sheet renamed");
}

function openRenameModal(id) {
  const s = sheets.find(x => x.id === id);
  openModal(`
    <div class="modal" onclick="event.stopPropagation()">
      <div class="modal-header">
        <span class="modal-title">Rename Sheet</span>
        <button class="modal-close" onclick="closeModalForce()">✕</button>
      </div>
      <div>
        <div class="field"><label>New Name</label>
          <input type="text" id="rename-input" value="${s.name}"
            onkeydown="if(event.key==='Enter')doRenameSheet('${id}')" autofocus />
        </div>
        <div class="modal-btns">
          <button class="btn-cancel" onclick="closeModalForce()">Cancel</button>
          <button class="btn-green" onclick="doRenameSheet('${id}')">Rename</button>
        </div>
      </div>
    </div>
  `);
  setTimeout(() => { const el = document.getElementById("rename-input"); if(el){el.focus();el.select();} }, 50);
}

function openDeleteModal(id) {
  const s = sheets.find(x => x.id === id);
  openModal(`
    <div class="modal" onclick="event.stopPropagation()">
      <div class="modal-header">
        <span class="modal-title">Delete Sheet?</span>
        <button class="modal-close" onclick="closeModalForce()">✕</button>
      </div>
      <div>
        <p style="color:#94a3b8;margin-bottom:20px">
          This will permanently delete <strong style="color:var(--red)">${s.name}</strong> and all its transactions.
        </p>
        <div class="modal-btns">
          <button class="btn-cancel" onclick="closeModalForce()">Cancel</button>
          <button class="btn-red" onclick="doDeleteSheet('${id}')">Delete</button>
        </div>
      </div>
    </div>
  `);
}

function doDeleteSheet(id) {
  if (sheets.length === 1) { showToast("Cannot delete the last sheet", "error"); closeModalForce(); return; }
  sheets = sheets.filter(s => s.id !== id);
  if (activeId === id) activeId = sheets[0].id;
  closeModalForce();
  editingId = null;
  resetForm();
  renderAll();
  showToast("Sheet deleted", "error");
}

function switchSheet(id) {
  activeId = id;
  editingId = null;
  search = "";
  document.getElementById("search-input").value = "";
  resetForm();
  renderAll();
}

// ── TRANSACTIONS ──────────────────────────────────────────────────────────────
function resetForm() {
  document.getElementById("f-date").value   = today();
  document.getElementById("f-name").value   = "";
  document.getElementById("f-credit").value = "";
  document.getElementById("f-debit").value  = "";
  document.getElementById("btn-add").textContent = "Add";
  document.getElementById("btn-cancel").classList.add("hidden");
  document.getElementById("form-title").innerHTML =
    `+ ADD TRANSACTION · <span style="color:#334155;font-size:11px">${activeSheet().name}</span>`;
}

function clearCredit() { document.getElementById("f-credit").value = ""; }
function clearDebit()  { document.getElementById("f-debit").value  = ""; }

function handleAdd() {
  const date   = document.getElementById("f-date").value;
  const name   = document.getElementById("f-name").value.trim();
  const credit = document.getElementById("f-credit").value;
  const debit  = document.getElementById("f-debit").value;

  if (!name)              { showToast("Name is required", "error"); return; }
  if (!credit && !debit)  { showToast("Enter credit or debit amount", "error"); return; }
  if (credit && debit)    { showToast("Enter only credit OR debit", "error"); return; }

  if (editingId) {
    updateSheet(activeId, s => ({
      transactions: s.transactions.map(t => t.id === editingId
        ? { id: editingId, date, name, credit, debit } : t)
    }));
    editingId = null;
    showToast("Transaction updated");
  } else {
    updateSheet(activeId, s => ({
      transactions: [...s.transactions, { id: uid(), date, name, credit, debit }]
    }));
    showToast("Transaction added");
  }
  resetForm();
  renderAll();
}

function startEdit(id) {
  const t = activeSheet().transactions.find(x => x.id === id);
  if (!t) return;
  editingId = id;
  document.getElementById("f-date").value   = t.date;
  document.getElementById("f-name").value   = t.name;
  document.getElementById("f-credit").value = t.credit;
  document.getElementById("f-debit").value  = t.debit;
  document.getElementById("btn-add").textContent = "Update";
  document.getElementById("btn-cancel").classList.remove("hidden");
  document.getElementById("form-title").innerHTML =
    `✎ EDIT TRANSACTION · <span style="color:#334155;font-size:11px">${activeSheet().name}</span>`;
  window.scrollTo({ top: 0, behavior: "smooth" });
}

function cancelEdit() { editingId = null; resetForm(); renderTable(); }

function deleteRow(id) {
  updateSheet(activeId, s => ({ transactions: s.transactions.filter(t => t.id !== id) }));
  if (editingId === id) { editingId = null; resetForm(); }
  showToast("Transaction deleted", "error");
  renderAll();
}

// ── OPENING BALANCE ───────────────────────────────────────────────────────────
function startEditOpening() {
  document.getElementById("opening-display").style.display = "none";
  const row = document.getElementById("opening-edit-row");
  row.style.display = "flex";
  document.getElementById("opening-input").value = activeSheet().openingBalance;
  document.getElementById("opening-input").focus();
}
function confirmOpening() {
  const val = parseFloat(document.getElementById("opening-input").value) || 0;
  updateSheet(activeId, () => ({ openingBalance: val }));
  cancelOpening();
  renderAll();
}
function cancelOpening() {
  document.getElementById("opening-display").style.display = "";
  document.getElementById("opening-edit-row").style.display = "none";
}
document.getElementById("opening-input").addEventListener("keydown", e => {
  if (e.key === "Enter") confirmOpening();
  if (e.key === "Escape") cancelOpening();
});

// ── SEARCH ────────────────────────────────────────────────────────────────────
function clearSearch() {
  document.getElementById("search-input").value = "";
  search = "";
  renderTable();
}

// ── COMPUTE ───────────────────────────────────────────────────────────────────
function computeRows(sheet, filter) {
  const txns = filter
    ? sheet.transactions.filter(t => t.name.toLowerCase().includes(filter.toLowerCase()))
    : sheet.transactions;
  let bal = sheet.openingBalance;
  return txns.map(t => {
    bal += (parseFloat(t.credit) || 0) - (parseFloat(t.debit) || 0);
    return { ...t, balance: bal };
  });
}

function computeTotals(sheet) {
  return sheet.transactions.reduce(
    (a, t) => ({ credit: a.credit + (parseFloat(t.credit)||0), debit: a.debit + (parseFloat(t.debit)||0) }),
    { credit: 0, debit: 0 }
  );
}

// ── RENDER ────────────────────────────────────────────────────────────────────
function renderTabs() {
  const bar = document.getElementById("tab-bar");
  // remove old tabs
  bar.querySelectorAll(".tab").forEach(el => el.remove());
  const addBtn = document.getElementById("add-tab-btn");

  sheets.forEach(s => {
    const div = document.createElement("div");
    div.className = "tab" + (s.id === activeId ? " active" : "");
    div.onclick = () => switchSheet(s.id);

    let actionsHtml = "";
    if (s.id === activeId) {
      actionsHtml = `
        <span class="tab-actions">
          <button class="tab-btn" title="Rename" onclick="event.stopPropagation();openRenameModal('${s.id}')">✎</button>
          ${sheets.length > 1 ? `<button class="tab-btn del" title="Delete" onclick="event.stopPropagation();openDeleteModal('${s.id}')">✕</button>` : ""}
        </span>`;
    }
    div.innerHTML = `<span style="font-weight:500">${s.name}</span>${actionsHtml}`;
    bar.insertBefore(div, addBtn);
  });
}

function renderSummary() {
  const s = activeSheet();
  const totals = computeTotals(s);
  const finalBal = s.openingBalance + totals.credit - totals.debit;
  const net = totals.credit - totals.debit;

  document.getElementById("bal-sheet-name").textContent = "Balance · " + s.name;
  const balAmt = document.getElementById("bal-amount");
  balAmt.textContent = fmt(finalBal);
  balAmt.style.color = finalBal >= 0 ? "#86efac" : "#fca5a5";

  document.getElementById("opening-display").innerHTML =
    `${fmt(s.openingBalance)} <span style="font-size:11px;opacity:.5">✎ edit</span>`;

  document.getElementById("card-credit").textContent = fmt(totals.credit);
  document.getElementById("card-debit").textContent  = fmt(totals.debit);

  const netEl = document.getElementById("card-net");
  netEl.textContent = fmt(net);
  netEl.style.color = net >= 0 ? "var(--green)" : "var(--red)";
  document.getElementById("card-net-icon").style.color = net >= 0 ? "var(--green)" : "var(--red)";

  document.getElementById("card-entries").textContent = s.transactions.length;
}

function renderTable() {
  const s = activeSheet();
  search = document.getElementById("search-input").value;
  const rows = computeRows(s, search);
  const totals = computeTotals(s);
  const finalBal = s.openingBalance + totals.credit - totals.debit;

  const tbody = document.getElementById("tbody");
  const table = document.getElementById("ledger-table");
  const empty = document.getElementById("empty-state");

  document.getElementById("row-count").textContent =
    rows.length + " entr" + (rows.length === 1 ? "y" : "ies");

  if (rows.length === 0) {
    table.classList.add("hidden");
    empty.classList.remove("hidden");
    document.getElementById("empty-text").innerHTML =
      `No transactions in <em>${s.name}</em>`;
    return;
  }
  table.classList.remove("hidden");
  empty.classList.add("hidden");

  tbody.innerHTML = rows.map((t, i) => `
    <tr class="${editingId === t.id ? "editing-row" : ""}">
      <td class="dimmed">${i + 1}</td>
      <td>${fmtDate(t.date)}</td>
      <td class="bold">${t.name}</td>
      <td class="green">${t.credit !== "" ? fmt(t.credit) : ""}</td>
      <td class="red">${t.debit !== "" ? fmt(t.debit) : ""}</td>
      <td class="${t.balance >= 0 ? "blue" : "red"}">${fmt(t.balance)}</td>
      <td>
        <button class="btn-edit" onclick="startEdit('${t.id}')">✎</button>
        <button class="btn-del"  onclick="deleteRow('${t.id}')">✕</button>
      </td>
    </tr>
  `).join("");

  document.getElementById("tfoot-credit").textContent = fmt(totals.credit);
  document.getElementById("tfoot-debit").textContent  = fmt(totals.debit);
  const tfBal = document.getElementById("tfoot-bal");
  tfBal.textContent  = fmt(finalBal);
  tfBal.style.color  = finalBal >= 0 ? "var(--blue)" : "var(--red)";
  tfBal.style.fontWeight = "700";
}

function renderAll() {
  renderTabs();
  renderSummary();
  renderTable();
  document.getElementById("form-sheet-name").textContent = activeSheet().name;
}

// ── EXPORT ────────────────────────────────────────────────────────────────────
function buildWs(sheet) {
  let bal = sheet.openingBalance;
  const dataRows = sheet.transactions.map(t => {
    bal += (parseFloat(t.credit)||0) - (parseFloat(t.debit)||0);
    return [t.date, t.name,
      t.credit !== "" ? parseFloat(t.credit) : "",
      t.debit  !== "" ? parseFloat(t.debit)  : "",
      bal];
  });
  const totC = sheet.transactions.reduce((a,t) => a+(parseFloat(t.credit)||0), 0);
  const totD = sheet.transactions.reduce((a,t) => a+(parseFloat(t.debit)||0),  0);
  const data = [
    [`Ledger: ${sheet.name}`],
    ["Opening Balance", sheet.openingBalance],
    [],
    ["Date","Name / Description","Credit (+)","Debit (−)","Balance"],
    ...dataRows,
    [],
    ["","TOTALS", totC, totD, bal],
  ];
  const ws = XLSX.utils.aoa_to_sheet(data);
  ws["!cols"] = [{wch:14},{wch:28},{wch:14},{wch:14},{wch:16}];
  return ws;
}

function handleExportActive() {
  const s = activeSheet();
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, buildWs(s), s.name.slice(0,31));
  XLSX.writeFile(wb, `Ledger_${s.name}_${today()}.xlsx`);
  showToast("Sheet exported");
}

function handleExportAll() {
  const wb = XLSX.utils.book_new();
  sheets.forEach(s => XLSX.utils.book_append_sheet(wb, buildWs(s), s.name.slice(0,31)));
  XLSX.writeFile(wb, `All_Ledgers_${today()}.xlsx`);
  showToast(`All ${sheets.length} sheet(s) exported`);
}

// ── IMPORT ────────────────────────────────────────────────────────────────────
function handleImport(e) {
  const file = e.target.files[0];
  if (!file) return;
  const reader = new FileReader();
  reader.onload = evt => {
    try {
      const wb = XLSX.read(evt.target.result, { type: "binary", cellDates: true });
      const imported = [];
      wb.SheetNames.forEach(name => {
        const ws  = wb.Sheets[name];
        const raw = XLSX.utils.sheet_to_json(ws, { header: 1, defval: "" });
        let openingBal = 0, dataStart = 0;
        for (let i = 0; i < Math.min(6, raw.length); i++) {
          const r = raw[i];
          if (String(r[0]).toLowerCase().includes("opening")) openingBal = parseFloat(r[1]) || 0;
          if (String(r[0]).toLowerCase().includes("date") && String(r[1]).toLowerCase().includes("name")) {
            dataStart = i + 1; break;
          }
        }
        const txns = [];
        for (let i = dataStart; i < raw.length; i++) {
          const r = raw[i];
          const n = String(r[1]||"").trim();
          if (!n || n.toLowerCase() === "totals" || n.toLowerCase() === "name / description") continue;
          if (!r[0] && !r[1]) continue;
          let dateVal = "";
          if (r[0]) {
            const d = r[0] instanceof Date ? r[0] : new Date(r[0]);
            dateVal = isNaN(d) ? String(r[0]) : d.toISOString().split("T")[0];
          }
          txns.push({
            id: uid(), date: dateVal, name: n,
            credit: r[2] !== "" && !isNaN(parseFloat(r[2])) ? parseFloat(r[2]) : "",
            debit:  r[3] !== "" && !isNaN(parseFloat(r[3])) ? parseFloat(r[3]) : "",
          });
        }
        imported.push({ id: uid(), name, openingBalance: openingBal, transactions: txns });
      });
      if (!imported.length) { showToast("No sheets found in file", "error"); return; }
      sheets   = imported;
      activeId = imported[0].id;
      editingId = null;
      resetForm();
      renderAll();
      showToast(`Imported ${imported.length} sheet(s)`);
    } catch(err) {
      showToast("Import failed — check file format", "error");
    }
  };
  reader.readAsBinaryString(file);
  e.target.value = "";
}

// ── INIT ──────────────────────────────────────────────────────────────────────
document.getElementById("f-date").value = today();
renderAll();
</script>
</body>
</html>
