import { useState, useRef } from "react";
import * as XLSX from "xlsx";

// ─────────────────────────────────────────────────────────────────────────────
// SUPABASE CONFIG — fill these in when you're ready to go cloud-based
// ─────────────────────────────────────────────────────────────────────────────
const SUPABASE_URL = "";        // e.g. "https://xyzxyz.supabase.co"
const SUPABASE_ANON_KEY = "";   // your project's anon/public key
// ─────────────────────────────────────────────────────────────────────────────
// When both values are set, the app will be ready to:
//   - createClient(SUPABASE_URL, SUPABASE_ANON_KEY)
//   - read/write sheets and transactions from Supabase tables
// ─────────────────────────────────────────────────────────────────────────────

// ── helpers ──────────────────────────────────────────────────────────────────
const uid = () => Date.now().toString(36) + Math.random().toString(36).slice(2);

function formatCurrency(val) {
  if (val === "" || val === null || val === undefined || isNaN(Number(val))) return "—";
  return new Intl.NumberFormat("en-IN", {
    style: "currency", currency: "INR", minimumFractionDigits: 2,
  }).format(val);
}

function formatDate(d) {
  if (!d) return "";
  const dt = new Date(d);
  if (isNaN(dt)) return d;
  return dt.toLocaleDateString("en-GB", { day: "2-digit", month: "short", year: "numeric" });
}

const newTxn = () => ({
  id: uid(),
  date: new Date().toISOString().split("T")[0],
  name: "", credit: "", debit: "",
});

const newSheet = (name) => ({
  id: uid(),
  name: name.trim() || "Ledger",
  openingBalance: 0,
  transactions: [],
});

// ── initial state ─────────────────────────────────────────────────────────────
const INITIAL_SHEETS = [newSheet("Personal")];

// ── main component ────────────────────────────────────────────────────────────
export default function App() {
  const [sheets, setSheets] = useState(INITIAL_SHEETS);
  const [activeId, setActiveId] = useState(INITIAL_SHEETS[0].id);

  // sheet-level modal state
  const [addSheetOpen, setAddSheetOpen] = useState(false);
  const [newSheetName, setNewSheetName] = useState("");
  const [renameId, setRenameId] = useState(null);
  const [renameName, setRenameName] = useState("");
  const [deleteConfirm, setDeleteConfirm] = useState(null); // sheet id

  // toast
  const [toast, setToast] = useState(null);

  // file input ref (per sheet, keyed by activeId)
  const fileRef = useRef();

  const showToast = (msg, type = "success") => {
    setToast({ msg, type });
    setTimeout(() => setToast(null), 2800);
  };

  // ── sheet helpers ───────────────────────────────────────────────────────────
  const activeSheet = sheets.find((s) => s.id === activeId) || sheets[0];

  const updateSheet = (id, updater) =>
    setSheets((prev) => prev.map((s) => (s.id === id ? { ...s, ...updater(s) } : s)));

  const addSheet = () => {
    if (!newSheetName.trim()) return;
    const s = newSheet(newSheetName);
    setSheets((prev) => [...prev, s]);
    setActiveId(s.id);
    setNewSheetName("");
    setAddSheetOpen(false);
    showToast(`Sheet "${s.name}" created`);
  };

  const renameSheet = () => {
    if (!renameName.trim()) return;
    updateSheet(renameId, () => ({ name: renameName.trim() }));
    showToast("Sheet renamed");
    setRenameId(null);
    setRenameName("");
  };

  const deleteSheet = (id) => {
    if (sheets.length === 1) { showToast("Cannot delete the last sheet", "error"); return; }
    const remaining = sheets.filter((s) => s.id !== id);
    setSheets(remaining);
    if (activeId === id) setActiveId(remaining[0].id);
    setDeleteConfirm(null);
    showToast("Sheet deleted", "error");
  };

  // ── transaction helpers (operate on activeSheet) ────────────────────────────
  const [form, setForm] = useState(newTxn());
  const [editingId, setEditingId] = useState(null);
  const [search, setSearch] = useState("");

  const handleAdd = () => {
    if (!form.name.trim()) { showToast("Name is required", "error"); return; }
    if (!form.credit && !form.debit) { showToast("Enter credit or debit amount", "error"); return; }
    if (form.credit && form.debit) { showToast("Enter only credit OR debit", "error"); return; }
    if (editingId) {
      updateSheet(activeId, (s) => ({
        transactions: s.transactions.map((t) => t.id === editingId ? { ...form, id: editingId } : t),
      }));
      showToast("Transaction updated");
      setEditingId(null);
    } else {
      updateSheet(activeId, (s) => ({
        transactions: [...s.transactions, { ...form, id: uid() }],
      }));
      showToast("Transaction added");
    }
    setForm(newTxn());
  };

  const handleEdit = (t) => { setForm({ ...t }); setEditingId(t.id); };

  const handleDelete = (id) => {
    updateSheet(activeId, (s) => ({
      transactions: s.transactions.filter((t) => t.id !== id),
    }));
    if (editingId === id) { setEditingId(null); setForm(newTxn()); }
    showToast("Transaction deleted", "error");
  };

  const handleCancel = () => { setEditingId(null); setForm(newTxn()); };

  const setOpeningBalance = (val) =>
    updateSheet(activeId, () => ({ openingBalance: parseFloat(val) || 0 }));

  // ── computed rows with running balance ──────────────────────────────────────
  const computeRows = (sheet) => {
    const txns = sheet.transactions.filter(
      (t) => !search || t.name.toLowerCase().includes(search.toLowerCase())
    );
    let bal = sheet.openingBalance;
    return txns.map((t) => {
      bal += (parseFloat(t.credit) || 0) - (parseFloat(t.debit) || 0);
      return { ...t, balance: bal };
    });
  };

  const rows = computeRows(activeSheet);

  const totals = activeSheet.transactions.reduce(
    (a, t) => ({ credit: a.credit + (parseFloat(t.credit) || 0), debit: a.debit + (parseFloat(t.debit) || 0) }),
    { credit: 0, debit: 0 }
  );
  const finalBalance = activeSheet.openingBalance + totals.credit - totals.debit;

  // ── export (current sheet only, or all sheets) ──────────────────────────────
  const exportSheet = (sheet) => {
    let bal = sheet.openingBalance;
    const rows = sheet.transactions.map((t) => {
      bal += (parseFloat(t.credit) || 0) - (parseFloat(t.debit) || 0);
      return [t.date, t.name, t.credit !== "" ? parseFloat(t.credit) : "", t.debit !== "" ? parseFloat(t.debit) : "", bal];
    });
    const data = [
      [`Ledger: ${sheet.name}`],
      ["Opening Balance", sheet.openingBalance],
      [],
      ["Date", "Name / Description", "Credit (+)", "Debit (−)", "Balance"],
      ...rows,
      [],
      ["", "TOTALS",
        sheet.transactions.reduce((a, t) => a + (parseFloat(t.credit) || 0), 0),
        sheet.transactions.reduce((a, t) => a + (parseFloat(t.debit) || 0), 0),
        bal,
      ],
    ];
    const ws = XLSX.utils.aoa_to_sheet(data);
    ws["!cols"] = [{ wch: 14 }, { wch: 28 }, { wch: 14 }, { wch: 14 }, { wch: 16 }];
    return ws;
  };

  const handleExportActive = () => {
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, exportSheet(activeSheet), activeSheet.name.slice(0, 31));
    XLSX.writeFile(wb, `Ledger_${activeSheet.name}_${new Date().toISOString().split("T")[0]}.xlsx`);
    showToast("Sheet exported");
  };

  const handleExportAll = () => {
    const wb = XLSX.utils.book_new();
    sheets.forEach((s) => {
      XLSX.utils.book_append_sheet(wb, exportSheet(s), s.name.slice(0, 31));
    });
    XLSX.writeFile(wb, `All_Ledgers_${new Date().toISOString().split("T")[0]}.xlsx`);
    showToast(`All ${sheets.length} sheets exported`);
  };

  // ── import ──────────────────────────────────────────────────────────────────
  const handleImport = (e) => {
    const file = e.target.files[0];
    if (!file) return;
    const reader = new FileReader();
    reader.onload = (evt) => {
      try {
        const wb = XLSX.read(evt.target.result, { type: "binary", cellDates: true });
        let importedSheets = [];

        wb.SheetNames.forEach((sheetName) => {
          const ws = wb.Sheets[sheetName];
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
            const nameVal = String(r[1] || "").trim();
            if (!nameVal || nameVal.toLowerCase() === "totals" || nameVal.toLowerCase() === "name / description") continue;
            if (!r[0] && !r[1]) continue;
            let dateVal = "";
            if (r[0]) {
              const d = r[0] instanceof Date ? r[0] : new Date(r[0]);
              dateVal = isNaN(d) ? String(r[0]) : d.toISOString().split("T")[0];
            }
            txns.push({
              id: uid(),
              date: dateVal, name: nameVal,
              credit: r[2] !== "" && !isNaN(parseFloat(r[2])) ? parseFloat(r[2]) : "",
              debit: r[3] !== "" && !isNaN(parseFloat(r[3])) ? parseFloat(r[3]) : "",
            });
          }

          importedSheets.push({ id: uid(), name: sheetName, openingBalance: openingBal, transactions: txns });
        });

        if (importedSheets.length === 0) { showToast("No sheets found in file", "error"); return; }
        setSheets(importedSheets);
        setActiveId(importedSheets[0].id);
        showToast(`Imported ${importedSheets.length} sheet(s)`);
      } catch {
        showToast("Import failed — check file format", "error");
      }
    };
    reader.readAsBinaryString(file);
    e.target.value = "";
  };

  // ── opening balance inline edit ─────────────────────────────────────────────
  const [editingOpening, setEditingOpening] = useState(false);
  const [openingInput, setOpeningInput] = useState("");

  // ── render ──────────────────────────────────────────────────────────────────
  return (
    <div style={S.root}>
      {/* TOAST */}
      {toast && (
        <div style={{ ...S.toast, background: toast.type === "error" ? "#dc2626" : "#059669" }}>
          {toast.type === "error" ? "✕ " : "✓ "}{toast.msg}
        </div>
      )}

      {/* MODALS */}
      {addSheetOpen && (
        <Modal title="New Sheet" onClose={() => setAddSheetOpen(false)}>
          <div style={S.modalBody}>
            <label style={S.label}>Sheet Name</label>
            <input
              style={S.input}
              placeholder="e.g. John, Business, Savings…"
              value={newSheetName}
              onChange={(e) => setNewSheetName(e.target.value)}
              onKeyDown={(e) => e.key === "Enter" && addSheet()}
              autoFocus
            />
            <div style={S.modalBtns}>
              <button style={S.cancelBtn} onClick={() => setAddSheetOpen(false)}>Cancel</button>
              <button style={S.addBtn} onClick={addSheet}>Create Sheet</button>
            </div>
          </div>
        </Modal>
      )}

      {renameId && (
        <Modal title="Rename Sheet" onClose={() => setRenameId(null)}>
          <div style={S.modalBody}>
            <label style={S.label}>New Name</label>
            <input
              style={S.input}
              value={renameName}
              onChange={(e) => setRenameName(e.target.value)}
              onKeyDown={(e) => e.key === "Enter" && renameSheet()}
              autoFocus
            />
            <div style={S.modalBtns}>
              <button style={S.cancelBtn} onClick={() => setRenameId(null)}>Cancel</button>
              <button style={S.addBtn} onClick={renameSheet}>Rename</button>
            </div>
          </div>
        </Modal>
      )}

      {deleteConfirm && (
        <Modal title="Delete Sheet?" onClose={() => setDeleteConfirm(null)}>
          <div style={S.modalBody}>
            <p style={{ color: "#94a3b8", marginBottom: "20px" }}>
              This will permanently delete <strong style={{ color: "#f87171" }}>
              {sheets.find(s => s.id === deleteConfirm)?.name}
              </strong> and all its transactions.
            </p>
            <div style={S.modalBtns}>
              <button style={S.cancelBtn} onClick={() => setDeleteConfirm(null)}>Cancel</button>
              <button style={{ ...S.addBtn, background: "#dc2626" }} onClick={() => deleteSheet(deleteConfirm)}>Delete</button>
            </div>
          </div>
        </Modal>
      )}

      {/* HEADER */}
      <header style={S.header}>
        <div style={S.headerLeft}>
          <div style={S.logo}>⬡</div>
          <div>
            <div style={S.title}>Account Ledger</div>
            <div style={S.subtitle}>Multi-Sheet Balance Tracker</div>
          </div>
        </div>
        <div style={S.balanceChip}>
          <span style={S.balChipLabel}>Balance · {activeSheet.name}</span>
          <span style={{ ...S.balChipAmt, color: finalBalance >= 0 ? "#86efac" : "#fca5a5" }}>
            {formatCurrency(finalBalance)}
          </span>
        </div>
      </header>

      {/* SHEET TABS */}
      <div style={S.tabBar}>
        <div style={S.tabs}>
          {sheets.map((s) => (
            <div
              key={s.id}
              style={{ ...S.tab, ...(s.id === activeId ? S.tabActive : {}) }}
              onClick={() => { setActiveId(s.id); setEditingId(null); setForm(newTxn()); setSearch(""); }}
            >
              <span style={S.tabName}>{s.name}</span>
              {s.id === activeId && (
                <span style={S.tabActions}>
                  <span
                    style={S.tabBtn}
                    title="Rename"
                    onClick={(e) => { e.stopPropagation(); setRenameId(s.id); setRenameName(s.name); }}
                  >✎</span>
                  {sheets.length > 1 && (
                    <span
                      style={{ ...S.tabBtn, color: "#f87171" }}
                      title="Delete"
                      onClick={(e) => { e.stopPropagation(); setDeleteConfirm(s.id); }}
                    >✕</span>
                  )}
                </span>
              )}
            </div>
          ))}
          <button style={S.addTabBtn} onClick={() => setAddSheetOpen(true)} title="Add new sheet">
            + New Sheet
          </button>
        </div>

      </div>

      {/* TOOLBAR — Export / Import */}
      <div style={S.toolbar}>
        <div style={S.toolbarLeft}>
          <span style={S.toolbarLabel}>Export:</span>
          <button style={S.exportActiveBtn} onClick={handleExportActive}>
            ↓ This Sheet
          </button>
          <button style={S.exportAllBtn} onClick={handleExportAll}>
            ↓ All Sheets
          </button>
        </div>
        <div style={S.toolbarRight}>
          <button style={S.importToolBtn} onClick={() => fileRef.current.click()}>
            ↑ Import Excel
          </button>
          <input ref={fileRef} type="file" accept=".xlsx,.xls" style={{ display: "none" }} onChange={handleImport} />
        </div>
      </div>

      {/* BODY */}
      <div style={S.body}>

        {/* Summary + Opening balance row */}
        <div style={S.topRow}>
          {/* opening balance */}
          <div style={S.openingCard}>
            <div style={S.openingLabel}>Opening Balance</div>
            {editingOpening ? (
              <div style={{ display: "flex", gap: "8px", alignItems: "center" }}>
                <input
                  type="number"
                  style={{ ...S.input, width: "140px", padding: "6px 10px" }}
                  value={openingInput}
                  onChange={(e) => setOpeningInput(e.target.value)}
                  onKeyDown={(e) => {
                    if (e.key === "Enter") { setOpeningBalance(openingInput); setEditingOpening(false); }
                    if (e.key === "Escape") setEditingOpening(false);
                  }}
                  autoFocus
                />
                <button style={S.smBtn} onClick={() => { setOpeningBalance(openingInput); setEditingOpening(false); }}>Set</button>
                <button style={{ ...S.smBtn, background: "#1e293b" }} onClick={() => setEditingOpening(false)}>✕</button>
              </div>
            ) : (
              <button
                style={S.openingVal}
                onClick={() => { setEditingOpening(true); setOpeningInput(String(activeSheet.openingBalance)); }}
              >
                {formatCurrency(activeSheet.openingBalance)} <span style={{ fontSize: "11px", opacity: 0.5 }}>✎ edit</span>
              </button>
            )}
          </div>

          {/* summary cards */}
          <div style={S.cards}>
            {[
              { label: "Total Credits", val: totals.credit, color: "#4ade80", icon: "↑" },
              { label: "Total Debits", val: totals.debit, color: "#f87171", icon: "↓" },
              { label: "Net Change", val: totals.credit - totals.debit, color: (totals.credit - totals.debit) >= 0 ? "#4ade80" : "#f87171", icon: "≈" },
              { label: "Entries", val: activeSheet.transactions.length, color: "#93c5fd", icon: "#", raw: true },
            ].map((c) => (
              <div key={c.label} style={S.card}>
                <span style={{ ...S.cardIcon, color: c.color }}>{c.icon}</span>
                <span style={S.cardLabel}>{c.label}</span>
                <span style={{ ...S.cardVal, color: c.color }}>{c.raw ? c.val : formatCurrency(c.val)}</span>
              </div>
            ))}
          </div>
        </div>

        {/* FORM */}
        <div style={S.formCard}>
          <div style={S.formTitle}>{editingId ? "✎  Edit Transaction" : "+  Add Transaction"} <span style={{ color: "#334155", fontSize: "11px" }}>· {activeSheet.name}</span></div>
          <div style={S.formRow}>
            <div style={S.field}>
              <label style={S.label}>Date</label>
              <input style={S.input} type="date" value={form.date} onChange={(e) => setForm({ ...form, date: e.target.value })} />
            </div>
            <div style={{ ...S.field, flex: 2 }}>
              <label style={S.label}>Name / Description</label>
              <input style={S.input} placeholder="e.g. Salary, Rent…" value={form.name} onChange={(e) => setForm({ ...form, name: e.target.value })} />
            </div>
            <div style={S.field}>
              <label style={{ ...S.label, color: "#4ade80" }}>Credit (+)</label>
              <input
                style={{ ...S.input, borderColor: form.credit ? "#4ade8066" : "" }}
                type="number" min="0" placeholder="0.00"
                value={form.credit}
                onChange={(e) => setForm({ ...form, credit: e.target.value, debit: "" })}
              />
            </div>
            <div style={S.field}>
              <label style={{ ...S.label, color: "#f87171" }}>Debit (−)</label>
              <input
                style={{ ...S.input, borderColor: form.debit ? "#f8717166" : "" }}
                type="number" min="0" placeholder="0.00"
                value={form.debit}
                onChange={(e) => setForm({ ...form, debit: e.target.value, credit: "" })}
              />
            </div>
            <div style={S.fieldBtns}>
              {editingId && <button style={S.cancelBtn} onClick={handleCancel}>Cancel</button>}
              <button style={S.addBtn} onClick={handleAdd}>{editingId ? "Update" : "Add"}</button>
            </div>
          </div>
        </div>

        {/* SEARCH */}
        <div style={S.searchRow}>
          <input
            style={S.searchInput}
            placeholder="🔍  Search transactions in this sheet…"
            value={search}
            onChange={(e) => setSearch(e.target.value)}
          />
          {search && <button style={S.clearSearch} onClick={() => setSearch("")}>✕</button>}
          <span style={S.rowCount}>{rows.length} entr{rows.length === 1 ? "y" : "ies"}</span>
        </div>

        {/* TABLE */}
        <div style={S.tableWrap}>
          {rows.length === 0 ? (
            <div style={S.empty}>
              <div style={S.emptyIcon}>📒</div>
              <div style={S.emptyText}>No transactions in <em>{activeSheet.name}</em></div>
              <div style={S.emptyHint}>Add one above or import an Excel file</div>
            </div>
          ) : (
            <table style={S.table}>
              <thead>
                <tr>
                  {["#", "Date", "Name / Description", "Credit (+)", "Debit (−)", "Balance", "Actions"].map((h) => (
                    <th key={h} style={S.th}>{h}</th>
                  ))}
                </tr>
              </thead>
              <tbody>
                {rows.map((t, i) => (
                  <tr key={t.id} style={{ ...S.tr, background: editingId === t.id ? "#162032" : i % 2 === 0 ? "#0b1523" : "#0e1a2b" }}>
                    <td style={{ ...S.td, color: "#334155", fontSize: "11px" }}>{i + 1}</td>
                    <td style={S.td}>{formatDate(t.date)}</td>
                    <td style={{ ...S.td, fontWeight: 500 }}>{t.name}</td>
                    <td style={{ ...S.td, color: "#4ade80" }}>{t.credit !== "" ? formatCurrency(t.credit) : ""}</td>
                    <td style={{ ...S.td, color: "#f87171" }}>{t.debit !== "" ? formatCurrency(t.debit) : ""}</td>
                    <td style={{ ...S.td, color: t.balance >= 0 ? "#93c5fd" : "#f87171", fontWeight: 600 }}>
                      {formatCurrency(t.balance)}
                    </td>
                    <td style={{ ...S.td }}>
                      <div style={{ display: "flex", gap: "6px" }}>
                        <button style={S.editBtn} onClick={() => handleEdit(t)}>✎</button>
                        <button style={S.delBtn} onClick={() => handleDelete(t.id)}>✕</button>
                      </div>
                    </td>
                  </tr>
                ))}
              </tbody>
              <tfoot>
                <tr style={{ background: "#0d1b2e" }}>
                  <td colSpan={3} style={{ ...S.td, color: "#475569", textAlign: "right", fontWeight: 700, fontSize: "11px", letterSpacing: "1px" }}>TOTALS</td>
                  <td style={{ ...S.td, color: "#4ade80", fontWeight: 700 }}>{formatCurrency(totals.credit)}</td>
                  <td style={{ ...S.td, color: "#f87171", fontWeight: 700 }}>{formatCurrency(totals.debit)}</td>
                  <td style={{ ...S.td, color: finalBalance >= 0 ? "#93c5fd" : "#f87171", fontWeight: 700 }}>{formatCurrency(finalBalance)}</td>
                  <td style={S.td} />
                </tr>
              </tfoot>
            </table>
          )}
        </div>

      </div>
    </div>
  );
}

// ── Modal wrapper ─────────────────────────────────────────────────────────────
function Modal({ title, onClose, children }) {
  return (
    <div style={S.overlay} onClick={onClose}>
      <div style={S.modal} onClick={(e) => e.stopPropagation()}>
        <div style={S.modalHeader}>
          <span style={S.modalTitle}>{title}</span>
          <button style={S.modalClose} onClick={onClose}>✕</button>
        </div>
        {children}
      </div>
    </div>
  );
}

// ── styles ────────────────────────────────────────────────────────────────────
const S = {
  root: { minHeight: "100vh", background: "#070d1a", fontFamily: "'Georgia', serif", color: "#e2e8f0" },

  toast: {
    position: "fixed", top: "18px", right: "18px", zIndex: 9999,
    padding: "11px 18px", borderRadius: "8px", color: "#fff",
    fontFamily: "'Courier New', monospace", fontSize: "13px", fontWeight: 600,
    boxShadow: "0 4px 24px rgba(0,0,0,0.5)",
  },

  overlay: {
    position: "fixed", inset: 0, background: "rgba(0,0,0,0.7)",
    display: "flex", alignItems: "center", justifyContent: "center", zIndex: 1000,
    backdropFilter: "blur(4px)",
  },
  modal: {
    background: "#0d1b2e", border: "1px solid #1e3a5f",
    borderRadius: "14px", width: "360px", padding: "24px",
    boxShadow: "0 20px 60px rgba(0,0,0,0.6)",
  },
  modalHeader: { display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: "20px" },
  modalTitle: { fontSize: "16px", fontWeight: 700, color: "#e0f2fe" },
  modalClose: { background: "none", border: "none", color: "#64748b", fontSize: "16px", cursor: "pointer" },
  modalBody: {},
  modalBtns: { display: "flex", gap: "10px", justifyContent: "flex-end", marginTop: "20px" },

  header: {
    background: "linear-gradient(135deg,#07121f 0%,#0c1e3b 100%)",
    borderBottom: "1px solid #1a3050",
    padding: "18px 28px",
    display: "flex", alignItems: "center", justifyContent: "space-between", flexWrap: "wrap", gap: "12px",
  },
  headerLeft: { display: "flex", alignItems: "center", gap: "14px" },
  logo: {
    width: "42px", height: "42px", borderRadius: "10px",
    background: "linear-gradient(135deg,#1d4ed8,#3b82f6)",
    display: "flex", alignItems: "center", justifyContent: "center",
    fontSize: "22px", color: "#bfdbfe",
  },
  title: { fontSize: "21px", fontWeight: 700, color: "#e0f2fe", letterSpacing: "0.4px" },
  subtitle: { fontSize: "11px", color: "#475569", letterSpacing: "1.5px", textTransform: "uppercase" },
  balanceChip: {
    background: "rgba(59,130,246,0.1)", border: "1px solid rgba(59,130,246,0.25)",
    borderRadius: "12px", padding: "10px 18px", textAlign: "center",
  },
  balChipLabel: { display: "block", fontSize: "10px", color: "#64748b", letterSpacing: "1.2px", textTransform: "uppercase" },
  balChipAmt: { fontSize: "19px", fontWeight: 700 },

  // tab bar
  tabBar: {
    background: "#070d1a", borderBottom: "1px solid #1a3050",
    display: "flex", alignItems: "center", justifyContent: "space-between",
    padding: "0 28px", flexWrap: "wrap", gap: "8px",
    position: "sticky", top: 0, zIndex: 50,
  },
  tabs: { display: "flex", alignItems: "center", gap: "2px", overflowX: "auto", paddingBottom: "0", flexWrap: "nowrap" },
  tab: {
    display: "flex", alignItems: "center", gap: "8px",
    padding: "11px 16px", cursor: "pointer",
    fontSize: "13px", color: "#64748b", borderBottom: "2px solid transparent",
    whiteSpace: "nowrap", transition: "color 0.2s",
    userSelect: "none",
  },
  tabActive: {
    color: "#7dd3fc", borderBottom: "2px solid #3b82f6",
    background: "rgba(59,130,246,0.06)",
  },
  tabName: { fontWeight: 500 },
  tabActions: { display: "flex", gap: "4px" },
  tabBtn: {
    padding: "1px 5px", borderRadius: "4px", fontSize: "11px",
    color: "#94a3b8", cursor: "pointer",
    background: "rgba(255,255,255,0.04)",
  },
  addTabBtn: {
    marginLeft: "6px", padding: "8px 14px",
    background: "rgba(59,130,246,0.1)", border: "1px dashed rgba(59,130,246,0.35)",
    borderRadius: "8px", color: "#60a5fa", fontSize: "12px", cursor: "pointer",
    whiteSpace: "nowrap", letterSpacing: "0.4px",
  },
  toolbar: {
    background: "#09131f", borderBottom: "1px solid #1a3050",
    padding: "10px 28px", display: "flex", alignItems: "center",
    justifyContent: "space-between", gap: "12px", flexWrap: "wrap",
  },
  toolbarLeft: { display: "flex", alignItems: "center", gap: "10px" },
  toolbarRight: { display: "flex", alignItems: "center", gap: "10px" },
  toolbarLabel: { fontSize: "11px", color: "#475569", letterSpacing: "1px", textTransform: "uppercase" },
  exportActiveBtn: {
    background: "linear-gradient(135deg,#1d4ed8,#2563eb)", border: "none",
    borderRadius: "8px", color: "#fff", padding: "8px 18px",
    fontSize: "13px", fontWeight: 600, cursor: "pointer", letterSpacing: "0.3px",
  },
  exportAllBtn: {
    background: "rgba(37,99,235,0.12)", border: "1px solid rgba(37,99,235,0.35)",
    borderRadius: "8px", color: "#60a5fa", padding: "8px 18px",
    fontSize: "13px", cursor: "pointer", letterSpacing: "0.3px",
  },
  importToolBtn: {
    background: "rgba(99,102,241,0.12)", border: "1px solid rgba(99,102,241,0.35)",
    borderRadius: "8px", color: "#a5b4fc", padding: "8px 18px",
    fontSize: "13px", cursor: "pointer", letterSpacing: "0.3px",
  },

  body: { padding: "22px 28px", maxWidth: "1380px", margin: "0 auto" },

  topRow: { display: "flex", gap: "16px", alignItems: "stretch", marginBottom: "18px", flexWrap: "wrap" },
  openingCard: {
    background: "#0d1b2e", border: "1px solid #1e3a5f",
    borderRadius: "12px", padding: "16px 20px", minWidth: "220px",
  },
  openingLabel: { fontSize: "10px", color: "#64748b", letterSpacing: "1.2px", textTransform: "uppercase", marginBottom: "10px" },
  openingVal: {
    background: "none", border: "1px dashed #1e3a5f", borderRadius: "8px",
    color: "#93c5fd", padding: "8px 14px", fontSize: "15px",
    cursor: "pointer", fontFamily: "'Georgia', serif", width: "100%", textAlign: "left",
  },

  cards: { display: "grid", gridTemplateColumns: "repeat(auto-fit, minmax(150px,1fr))", gap: "12px", flex: 1 },
  card: {
    background: "#0d1b2e", border: "1px solid #1a3050",
    borderRadius: "12px", padding: "14px 18px",
    display: "flex", flexDirection: "column", gap: "4px",
  },
  cardIcon: { fontSize: "18px" },
  cardLabel: { fontSize: "10px", color: "#64748b", letterSpacing: "1px", textTransform: "uppercase" },
  cardVal: { fontSize: "17px", fontWeight: 700 },

  formCard: {
    background: "#0b1726", border: "1px solid #1a3050",
    borderRadius: "12px", padding: "18px 22px", marginBottom: "16px",
  },
  formTitle: { fontSize: "11px", color: "#475569", letterSpacing: "1px", textTransform: "uppercase", marginBottom: "14px" },
  formRow: { display: "flex", gap: "10px", flexWrap: "wrap", alignItems: "flex-end" },
  field: { display: "flex", flexDirection: "column", flex: 1, minWidth: "120px" },
  label: { fontSize: "10px", color: "#64748b", letterSpacing: "1px", textTransform: "uppercase", marginBottom: "5px" },
  input: {
    background: "#070d1a", border: "1px solid #1e3a5f", borderRadius: "8px",
    color: "#e2e8f0", padding: "8px 12px", fontSize: "14px", outline: "none",
    fontFamily: "'Georgia', serif",
  },
  fieldBtns: { display: "flex", gap: "8px", alignItems: "flex-end" },
  addBtn: {
    background: "linear-gradient(135deg,#059669,#10b981)", border: "none",
    borderRadius: "8px", color: "#fff", padding: "8px 22px",
    fontSize: "14px", fontWeight: 700, cursor: "pointer", height: "38px",
  },
  cancelBtn: {
    background: "none", border: "1px solid #1e3a5f", borderRadius: "8px",
    color: "#64748b", padding: "8px 14px", fontSize: "13px", cursor: "pointer", height: "38px",
  },
  smBtn: {
    background: "#1d4ed8", border: "none", borderRadius: "6px",
    color: "#fff", padding: "6px 12px", fontSize: "13px", cursor: "pointer",
  },

  searchRow: { display: "flex", alignItems: "center", gap: "10px", marginBottom: "12px" },
  searchInput: {
    flex: 1, background: "#0b1726", border: "1px solid #1a3050",
    borderRadius: "8px", color: "#e2e8f0", padding: "8px 14px",
    fontSize: "14px", outline: "none", fontFamily: "'Georgia', serif",
  },
  clearSearch: {
    background: "none", border: "none", color: "#475569", fontSize: "14px", cursor: "pointer",
  },
  rowCount: { fontSize: "12px", color: "#334155", whiteSpace: "nowrap" },

  tableWrap: {
    background: "#09131f", border: "1px solid #1a3050",
    borderRadius: "12px", overflow: "hidden",
  },
  table: { width: "100%", borderCollapse: "collapse" },
  th: {
    padding: "11px 16px", textAlign: "left",
    fontSize: "10px", letterSpacing: "1.2px", textTransform: "uppercase",
    color: "#475569", background: "#0b1726", borderBottom: "1px solid #1a3050",
  },
  tr: {},
  td: { padding: "10px 16px", fontSize: "13px", color: "#94a3b8", borderBottom: "1px solid #0d1b2e" },
  editBtn: {
    background: "rgba(59,130,246,0.12)", border: "1px solid rgba(59,130,246,0.25)",
    borderRadius: "6px", color: "#93c5fd", padding: "3px 8px", cursor: "pointer", fontSize: "12px",
  },
  delBtn: {
    background: "rgba(248,113,113,0.08)", border: "1px solid rgba(248,113,113,0.25)",
    borderRadius: "6px", color: "#f87171", padding: "3px 8px", cursor: "pointer", fontSize: "11px",
  },
  empty: { padding: "56px 20px", textAlign: "center" },
  emptyIcon: { fontSize: "44px", marginBottom: "10px" },
  emptyText: { fontSize: "17px", color: "#1e3a5f", marginBottom: "6px" },
  emptyHint: { fontSize: "12px", color: "#1a3050" },

};
