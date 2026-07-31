import { useState, useEffect, useMemo, useCallback } from "react";
import * as XLSX from "xlsx";
import { Upload, ChevronDown, ChevronRight, Mail, Bell, BellRing, Check, CheckCheck, Users, UserPlus, AlertTriangle, Trash2, Plus, ShieldCheck, FileDown, LogOut } from "lucide-react";
import { supabase } from "./supabaseClient";
import { subscribeToPush } from "./pushSubscribe";

// ---- helpers ----------------------------------------------------------

const todayISO = () => new Date().toISOString().slice(0, 10);
const daysBetween = (fromISO, toISO) => {
  if (!toISO) return null;
  const a = new Date(fromISO + "T00:00:00");
  const b = new Date(toISO + "T00:00:00");
  return Math.round((b - a) / 86400000);
};

function normalizeKey(raw) {
  if (!raw) return "";
  return raw.toString().toUpperCase().replace(/P\.?\s?O\.?\s?BOX.*$/i, "").replace(/[().,]/g, "").replace(/\s+/g, " ").trim();
}

function excelDateToISO(val) {
  if (val === null || val === undefined || val === "") return null;
  if (val instanceof Date) return val.toISOString().slice(0, 10);
  if (typeof val === "number") return new Date(Math.round((val - 25569) * 86400 * 1000)).toISOString().slice(0, 10);
  const d = new Date(val);
  return !isNaN(d) ? d.toISOString().slice(0, 10) : null;
}

function statusFor(daysLeft, delivered, acknowledged) {
  if (delivered) return { key: "delivered", label: "Delivered", tone: "delivered" };
  if (daysLeft === null) return { key: "unknown", label: "No ETA", tone: "unknown" };
  if (daysLeft < 0) return { key: "overdue", label: `${Math.abs(daysLeft)}d overdue`, tone: "overdue" };
  if (daysLeft <= 5) return { key: "critical", label: `${daysLeft}d left`, tone: acknowledged ? "acked" : "critical" };
  if (daysLeft <= 10) return { key: "warning", label: `${daysLeft}d left`, tone: acknowledged ? "acked" : "warning" };
  if (daysLeft <= 15) return { key: "upcoming", label: `${daysLeft}d left`, tone: acknowledged ? "acked" : "upcoming" };
  return { key: "ok", label: `${daysLeft}d left`, tone: "ok" };
}

// Hardcoded customer -> coordinator assignments. Fill these in with real
// pairs — matching checks whether the customer's name CONTAINS your key
// (case/whitespace-insensitive), so one entry like "LG ELECTRONICS" catches
// every legal-entity variant ("LG ELECTRONICS MIDDLE EAST", "LG ELECTRONICS
// GULF FZE", etc.) without listing each one. Every import checks this table
// FIRST, before any Coordinator/PIC column in the sheet. Anyone not listed
// here still always gets assigned — it just falls back to whoever is running
// the import instead of a specific pinned person.
const CUSTOMER_COORDINATOR_MAP = {
  "AL YOUSUF ELECTRONICS": "Rahma",
  "LG ELECTRONICS": "Saron Solomo",
};

function fireBrowserNotification(title, body) {
  try {
    if (typeof Notification === "undefined") return;
    if (Notification.permission === "granted") new Notification(title, { body });
  } catch (e) { /* notifications not available */ }
}

function timeAgo(iso) {
  const mins = Math.round((Date.now() - new Date(iso).getTime()) / 60000);
  if (mins < 1) return "just now";
  if (mins < 60) return `${mins}m ago`;
  const hrs = Math.round(mins / 60);
  if (hrs < 24) return `${hrs}h ago`;
  return `${Math.round(hrs / 24)}d ago`;
}

function reminderText(s, customerName) {
  const lines = [`Shipment reminder: ${s.shipment_no}`, `Customer: ${customerName || "—"}`, `Arrival (ETA): ${s.arrival_date} (${s.status.label})`];
  if (s.tracking_link) lines.push(`Tracking: ${s.tracking_link}`);
  return lines.join("\n");
}

// ---- login screen -------------------------------------------------------

function LoginScreen({ onLoggedIn }) {
  const [mode, setMode] = useState("login"); // "login" | "signup"
  const [email, setEmail] = useState("");
  const [password, setPassword] = useState("");
  const [busy, setBusy] = useState(false);
  const [error, setError] = useState("");
  const [info, setInfo] = useState("");

  async function submit(e) {
    e.preventDefault();
    setBusy(true); setError(""); setInfo("");
    try {
      if (mode === "login") {
        const { error } = await supabase.auth.signInWithPassword({ email, password });
        if (error) throw error;
        onLoggedIn();
      } else {
        const { error } = await supabase.auth.signUp({ email, password });
        if (error) throw error;
        setInfo("Account created. Ask an admin to link your coordinator profile, then sign in.");
        setMode("login");
      }
    } catch (err) {
      setError(err.message || "Something went wrong.");
    } finally {
      setBusy(false);
    }
  }

  return (
    <div style={{ minHeight: "100vh", display: "flex", alignItems: "center", justifyContent: "center", background: "#FAFAFA", fontFamily: "'Inter', system-ui, sans-serif" }}>
      <form onSubmit={submit} style={{ background: "#fff", border: "1px solid #E5E7EB", borderRadius: 14, padding: 28, width: "100%", maxWidth: 340 }}>
        <div style={{ fontWeight: 700, fontSize: 17, color: "#111827", marginBottom: 4 }}>📦 Shipment Manifest</div>
        <div style={{ fontSize: 13, color: "#6B7280", marginBottom: 18 }}>{mode === "login" ? "Sign in to continue" : "Create your account"}</div>
        {error && <div style={{ background: "#FEF2F2", color: "#B91C1C", padding: "8px 10px", borderRadius: 7, fontSize: 12.5, marginBottom: 10 }}>{error}</div>}
        {info && <div style={{ background: "#F0FDF4", color: "#15803D", padding: "8px 10px", borderRadius: 7, fontSize: 12.5, marginBottom: 10 }}>{info}</div>}
        <div style={{ display: "flex", flexDirection: "column", gap: 10 }}>
          <input type="email" required placeholder="you@company.com" value={email} onChange={(e) => setEmail(e.target.value)} style={{ width: "100%" }} />
          <input type="password" required minLength={6} placeholder="Password" value={password} onChange={(e) => setPassword(e.target.value)} style={{ width: "100%" }} />
          <button type="submit" disabled={busy} className="btn" style={{ background: "#DC2626", color: "#fff", padding: "10px", borderRadius: 7, fontWeight: 600, fontSize: 13 }}>
            {busy ? "Please wait…" : mode === "login" ? "Sign in" : "Sign up"}
          </button>
        </div>
        <button type="button" className="btn" onClick={() => { setMode(mode === "login" ? "signup" : "login"); setError(""); setInfo(""); }} style={{ background: "transparent", color: "#6B7280", fontSize: 12.5, marginTop: 14, width: "100%" }}>
          {mode === "login" ? "Need an account? Sign up" : "Already have an account? Sign in"}
        </button>
      </form>
    </div>
  );
}

// ---- main component -------------------------------------------------------

export default function ShipmentApp() {
  const [session, setSession] = useState(undefined); // undefined = loading, null = logged out
  const [me, setMe] = useState(null); // row from `coordinators`
  const [meMissing, setMeMissing] = useState(false);

  const [shipments, setShipments] = useState([]);
  const [customers, setCustomers] = useState({}); // id -> {id, canonicalName, email, coordinators: [ids], templates: []}
  const [coordinators, setCoordinators] = useState([]); // [{id, name, is_admin}]
  const [notifications, setNotifications] = useState([]);

  const [loaded, setLoaded] = useState(false);
  const [viewMode, setViewMode] = useState("mine");
  const [expanded, setExpanded] = useState({});
  const [templatesOpen, setTemplatesOpen] = useState({});
  const [assignOpenKey, setAssignOpenKey] = useState(null);
  const [showAlert, setShowAlert] = useState(false);
  const [showNotifPanel, setShowNotifPanel] = useState(false);
  const [notifPermission, setNotifPermission] = useState(typeof Notification !== "undefined" ? Notification.permission : "unsupported");
  const [dismissedThisSession, setDismissedThisSession] = useState(false);
  const [importSummary, setImportSummary] = useState(null);
  const [error, setError] = useState("");
  const [templateForm, setTemplateForm] = useState(null);
  const [editingShipment, setEditingShipment] = useState(null);
  const today = todayISO();

  const isAdmin = !!(me && me.is_admin);
  const coordinatorNames = coordinators.map((c) => c.name);
  const nameById = useMemo(() => Object.fromEntries(coordinators.map((c) => [c.id, c.name])), [coordinators]);
  const idByName = useMemo(() => Object.fromEntries(coordinators.map((c) => [c.name, c.id])), [coordinators]);

  // ---- auth ----
  useEffect(() => {
    supabase.auth.getSession().then(({ data }) => setSession(data.session));
    const { data: sub } = supabase.auth.onAuthStateChange((_event, sess) => setSession(sess));
    return () => sub.subscription.unsubscribe();
  }, []);

  useEffect(() => {
    if (!session) { setMe(null); return; }
    (async () => {
      const { data, error } = await supabase.from("coordinators").select("*").eq("id", session.user.id).maybeSingle();
      if (error) { setError(error.message); return; }
      if (!data) { setMeMissing(true); return; }
      setMe(data);
    })();
  }, [session]);

  // ---- initial data load ----
  const loadAll = useCallback(async () => {
    const [{ data: coordData }, { data: custData }, { data: shipData }, { data: notifData }] = await Promise.all([
      supabase.from("coordinators").select("*"),
      supabase.from("customers").select("*, customer_coordinators(coordinator_id), customer_email_templates(*)"),
      supabase.from("shipments").select("*"),
      supabase.from("notifications").select("*").order("created_at", { ascending: false }),
    ]);
    setCoordinators(coordData || []);
    const custMap = {};
    (custData || []).forEach((c) => {
      custMap[c.id] = {
        id: c.id,
        canonicalName: c.canonical_name,
        email: c.email || "",
        coordinators: (c.customer_coordinators || []).map((cc) => cc.coordinator_id),
        templates: c.customer_email_templates || [],
      };
    });
    setCustomers(custMap);
    setShipments(shipData || []);
    setNotifications(notifData || []);
    setLoaded(true);
  }, []);

  useEffect(() => { if (me) loadAll(); }, [me, loadAll]);

  // ---- live notifications (works even while this tab sits in the background) ----
  useEffect(() => {
    if (!me) return;
    const channel = supabase
      .channel("notifications-" + me.id)
      .on("postgres_changes", { event: "INSERT", schema: "public", table: "notifications", filter: `recipient_id=eq.${me.id}` }, (payload) => {
        setNotifications((prev) => [payload.new, ...prev]);
        fireBrowserNotification(payload.new.title, payload.new.body);
      })
      .subscribe();
    return () => { supabase.removeChannel(channel); };
  }, [me]);

  function requestNotifPermission() {
    if (typeof Notification === "undefined") return;
    Notification.requestPermission().then((perm) => {
      setNotifPermission(perm);
      if (perm === "granted" && me) subscribeToPush(me.id);
    });
  }

  // ---- import (admin only — enforced both by UI and by RLS) ----
  async function importRows(rows, sheetLabel) {
    setError("");
    let added = 0, skipped = 0, newCustomers = 0;
    const existingShipmentNos = new Set(shipments.map((s) => s.shipment_no));
    const toInsertShipments = [];
    const customerIdCache = {}; // normalizedKey -> customer_id, for this run
    const assignedPairsThisRun = new Set(); // "customerId:coordinatorId" already handled this run

    for (const row of rows) {
      const shipmentNo = row["Shipment No."] || row["Shipment No"];
      if (!shipmentNo) continue;
      if (existingShipmentNos.has(String(shipmentNo))) { skipped++; continue; }

      const cneeRaw = (row["CNEE  Name"] || row["CNEE Name"] || "").toString().trim();
      const key = normalizeKey(cneeRaw);
      let customerId = customerIdCache[key];

      if (!customerId && key) {
        // 1. known alias?
        const { data: alias } = await supabase.from("customer_name_aliases").select("customer_id").eq("raw_name", cneeRaw).maybeSingle();
        if (alias) {
          customerId = alias.customer_id;
        } else {
          // 2. Upsert by canonical_name instead of select-then-insert — avoids a
          // race/duplicate-key error when the same name shows up more than once
          // (either across rows in this file, or against an already-imported name).
          const { data: custRow, error: custErr } = await supabase
            .from("customers")
            .upsert({ canonical_name: cneeRaw }, { onConflict: "canonical_name" })
            .select("id")
            .single();
          if (custErr) { setError(custErr.message); continue; }
          customerId = custRow.id;
          newCustomers++;
          await supabase
            .from("customer_name_aliases")
            .upsert({ raw_name: cneeRaw, customer_id: customerId }, { onConflict: "raw_name" });
        }
        customerIdCache[key] = customerId;
      }

      // 3. Coordinator assignment — ALWAYS resolves to someone, never left "Unassigned".
      // If the sheet names a real coordinator (via Coordinator / PIC / Coordinator Name /
      // Assigned To column), use them. Otherwise fall back to whoever is running this
      // import (the admin), so every customer always has an owner.
      const coordName = (row["Coordinator"] || row["PIC"] || row["Coordinator Name"] || row["Assigned To"] || "").toString().trim();
      const match = coordName ? coordinators.find((c) => c.name.toLowerCase() === coordName.toLowerCase()) : null;
      const hardcodedEntry = Object.entries(CUSTOMER_COORDINATOR_MAP).find(([mapKey]) => key.includes(normalizeKey(mapKey)));
      const hardcodedName = hardcodedEntry ? hardcodedEntry[1] : null;
      const hardcodedMatch = hardcodedName ? coordinators.find((c) => c.name.toLowerCase() === hardcodedName.toLowerCase()) : null;
      const fallback = coordinators.find((c) => c.id === me.id) || null;
      const assignTo = hardcodedMatch || match || fallback;
      if (assignTo && customerId) {
        const pairKey = `${customerId}:${assignTo.id}`;
        if (!assignedPairsThisRun.has(pairKey)) {
          assignedPairsThisRun.add(pairKey);
          await supabase.from("customer_coordinators").upsert(
            { customer_id: customerId, coordinator_id: assignTo.id, assigned_by: me.id },
            { onConflict: "customer_id,coordinator_id", ignoreDuplicates: true }
          );
        }
      }

      toInsertShipments.push({
        hbl_no: row["HBL No."] || "",
        shipment_no: String(shipmentNo),
        mbl_no: row["MBL No."] || "",
        on_board_date: excelDateToISO(row["Onb. Date"]),
        arrival_date: excelDateToISO(row["ARR Date"]),
        shpr_name: (row["SHPR Name"] || "").toString().trim(),
        cnee_name_raw: cneeRaw,
        customer_id: customerId || null,
        carrier_name: row["Carrier Name"] || "",
        pol: row["POL"] || "",
        pod: row["POD"] || "",
        vessel: row["Vessel"] || "",
        do_status_3p: row["3P Letter "] || row["3P Letter"] || "",
        do_status_shared: row["Do Status"] || "",
        do_shared_date: excelDateToISO(row["Do shared date"]),
        invoice_status: row["Invoice status"] || "",
        remark: row["Remark"] || "",
        raw_details: row, // full original row, every column, nothing lost — viewable per shipment
      });
      existingShipmentNos.add(String(shipmentNo));
      added++;
    }

    if (toInsertShipments.length > 0) {
      const { error: insErr } = await supabase.from("shipments").upsert(toInsertShipments, { onConflict: "shipment_no", ignoreDuplicates: true });
      if (insErr) { setError(insErr.message); return; }
    }
    await loadAll();
    setImportSummary({ added, skipped, newCustomers, total: rows.length, sheetUsed: sheetLabel });
  }

  function handleFile(e) {
    const file = e.target.files[0];
    if (!file) return;
    const reader = new FileReader();
    reader.onerror = () => setError("Couldn't read the file from your device — try selecting it again.");
    reader.onload = (evt) => {
      try {
        const wb = XLSX.read(evt.target.result, { type: "array", cellDates: true });
        let sheetName = null, rows = [];
        for (const name of wb.SheetNames) {
          const candidateRows = XLSX.utils.sheet_to_json(wb.Sheets[name], { defval: null });
          if (candidateRows.length === 0) continue;
          const headers = Object.keys(candidateRows[0]).map((h) => h.trim().toLowerCase());
          if (headers.some((h) => h.startsWith("shipment no"))) { sheetName = name; rows = candidateRows; break; }
        }
        if (!sheetName) { setError(`Couldn't find a "Shipment No." column. Sheets in this file: ${wb.SheetNames.join(", ")}.`); return; }
        importRows(rows, sheetName);
      } catch (err) {
        setError(`Import failed: ${err.message || err}`);
      }
    };
    reader.readAsArrayBuffer(file);
    e.target.value = "";
  }

  // ---- grouping / at-risk ----
  const grouped = useMemo(() => {
    const map = {};
    shipments.forEach((s) => {
      const key = s.customer_id || "UNKNOWN";
      (map[key] = map[key] || []).push(s);
    });
    return Object.entries(map)
      .map(([key, ships]) => ({
        key,
        customer: customers[key] || { id: key, canonicalName: ships[0]?.cnee_name_raw || "Unknown", email: "", coordinators: [], templates: [] },
        shipments: ships
          .map((s) => { const daysLeft = daysBetween(today, s.arrival_date); return { ...s, daysLeft, status: statusFor(daysLeft, s.delivered, s.reminder_acknowledged) }; })
          .sort((a, b) => (a.arrival_date || "").localeCompare(b.arrival_date || "")),
      }))
      .filter((g) => (viewMode === "mine" ? (g.customer.coordinators || []).includes(me?.id) : true))
      .sort((a, b) => a.customer.canonicalName.localeCompare(b.customer.canonicalName));
  }, [shipments, customers, viewMode, me, today]);

  const atRisk = useMemo(() => {
    const allGrouped = isAdmin
      ? Object.entries(shipments.reduce((m, s) => { const k = s.customer_id || "UNKNOWN"; (m[k] = m[k] || []).push(s); return m; }, {}))
          .map(([key, ships]) => ({ key, customer: customers[key] || { canonicalName: ships[0]?.cnee_name_raw }, shipments: ships.map((s) => ({ ...s, daysLeft: daysBetween(today, s.arrival_date), status: statusFor(daysBetween(today, s.arrival_date), s.delivered, s.reminder_acknowledged) })) }))
      : grouped;
    const list = [];
    allGrouped.forEach((g) => g.shipments.forEach((s) => {
      if (!s.delivered && !s.reminder_acknowledged && s.daysLeft !== null && s.daysLeft <= 15) list.push({ ...s, customerName: g.customer.canonicalName });
    }));
    return list;
  }, [grouped, isAdmin, shipments, customers, today]);

  useEffect(() => { if (loaded && atRisk.length > 0 && !dismissedThisSession) setShowAlert(true); }, [loaded, atRisk.length, dismissedThisSession]);

  async function updateCustomerEmail(customerId, email) {
    setCustomers((p) => ({ ...p, [customerId]: { ...p[customerId], email } }));
    const { error } = await supabase.from("customers").update({ email }).eq("id", customerId);
    if (error) setError(error.message);
  }

  function toggleExpand(key) { setExpanded((p) => ({ ...p, [key]: !p[key] })); }
  function toggleTemplates(key) { setTemplatesOpen((p) => ({ ...p, [key]: !p[key] })); }

  async function toggleCoordinatorAssignment(customerId, coordinatorId) {
    const c = customers[customerId];
    const isAssigned = (c.coordinators || []).includes(coordinatorId);
    if (isAssigned) {
      const { error } = await supabase.from("customer_coordinators").delete().eq("customer_id", customerId).eq("coordinator_id", coordinatorId);
      if (error) { setError(error.message); return; }
    } else {
      const { error } = await supabase.from("customer_coordinators").insert({ customer_id: customerId, coordinator_id: coordinatorId, assigned_by: me.id });
      if (error) { setError(error.message); return; }
    }
    setCustomers((p) => {
      const list = p[customerId].coordinators || [];
      const next = isAssigned ? list.filter((id) => id !== coordinatorId) : [...list, coordinatorId];
      return { ...p, [customerId]: { ...p[customerId], coordinators: next } };
    });
  }

  async function markNotificationRead(id) {
    setNotifications((p) => p.map((n) => (n.id === id ? { ...n, is_read: true } : n)));
    await supabase.from("notifications").update({ is_read: true }).eq("id", id);
  }
  async function markAllNotificationsRead() {
    setNotifications((p) => p.map((n) => ({ ...n, is_read: true })));
    await supabase.from("notifications").update({ is_read: true }).eq("recipient_id", me.id).eq("is_read", false);
  }

  async function toggleAck(id) {
    const s = shipments.find((x) => x.id === id);
    setShipments((p) => p.map((x) => (x.id === id ? { ...x, reminder_acknowledged: !x.reminder_acknowledged } : x)));
    await supabase.from("shipments").update({ reminder_acknowledged: !s.reminder_acknowledged }).eq("id", id);
  }
  async function toggleDelivered(id) {
    const s = shipments.find((x) => x.id === id);
    setShipments((p) => p.map((x) => (x.id === id ? { ...x, delivered: !x.delivered } : x)));
    await supabase.from("shipments").update({ delivered: !s.delivered }).eq("id", id);
  }
  async function removeShipment(id) {
    setShipments((p) => p.filter((x) => x.id !== id));
    await supabase.from("shipments").delete().eq("id", id);
  }
  async function saveShipmentEdit(updated) {
    setShipments((p) => p.map((s) => (s.id === updated.id ? { ...s, ...updated } : s)));
    setEditingShipment(null);
    const { arrival_date, tracking_link, remark } = updated;
    await supabase.from("shipments").update({ arrival_date, tracking_link, remark }).eq("id", updated.id);
  }

  async function saveTemplate() {
    if (!templateForm || !templateForm.title) return;
    const { customerKey, id, title, subject, body } = templateForm;
    if (id) {
      await supabase.from("customer_email_templates").update({ title, subject, body }).eq("id", id);
    } else {
      await supabase.from("customer_email_templates").insert({ customer_id: customerKey, title, subject, body });
    }
    await loadAll();
    setTemplateForm(null);
  }
  async function deleteTemplate(customerKey, id) {
    await supabase.from("customer_email_templates").delete().eq("id", id);
    setCustomers((p) => ({ ...p, [customerKey]: { ...p[customerKey], templates: (p[customerKey].templates || []).filter((t) => t.id !== id) } }));
  }

  async function copyReminder(s, customerName) {
    try { await navigator.clipboard.writeText(reminderText(s, customerName)); }
    catch (e) { setError("Couldn't copy to clipboard."); }
  }

  function downloadMyShipments() {
    const mine = shipments
      .filter((s) => { const c = customers[s.customer_id]; return c && (c.coordinators || []).includes(me.id); })
      .map((s) => ({
        "Customer": customers[s.customer_id]?.canonicalName || s.cnee_name_raw,
        "Shipment No.": s.shipment_no, "HBL No.": s.hbl_no, "MBL No.": s.mbl_no,
        "Arrival Date": s.arrival_date, "Days Left": daysBetween(today, s.arrival_date),
        "Shipper": s.shpr_name, "Carrier": s.carrier_name,
        "DO Made": s.reminder_acknowledged ? "Yes" : "No", "Delivered": s.delivered ? "Yes" : "No",
        "Tracking Link": s.tracking_link, "Remark": s.remark,
      }));
    if (mine.length === 0) { setError("No shipments assigned to you yet — nothing to download."); return; }
    const ws = XLSX.utils.json_to_sheet(mine);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, me.name);
    XLSX.writeFile(wb, `${me.name}-shipments-${today}.xlsx`);
  }

  const myAssignedCount = me ? Object.values(customers).filter((c) => (c.coordinators || []).includes(me.id)).length : 0;
  const myUnreadCount = notifications.filter((n) => !n.is_read).length;

  if (session === undefined) return <div style={{ padding: 40, textAlign: "center", color: "#9CA3AF" }}>Loading…</div>;
  if (!session) return <LoginScreen onLoggedIn={() => {}} />;
  if (meMissing) {
    return (
      <div style={{ minHeight: "100vh", display: "flex", alignItems: "center", justifyContent: "center", background: "#FAFAFA", fontFamily: "'Inter', system-ui, sans-serif" }}>
        <div style={{ maxWidth: 360, textAlign: "center", padding: 24 }}>
          <div style={{ fontWeight: 700, fontSize: 15, color: "#111827", marginBottom: 8 }}>Account not linked yet</div>
          <div style={{ fontSize: 13, color: "#6B7280", marginBottom: 16 }}>You're signed in, but an admin still needs to link your login to a coordinator profile in Supabase.</div>
          <button className="btn" onClick={() => supabase.auth.signOut()} style={{ background: "#F3F4F6", color: "#374151", padding: "9px 16px", borderRadius: 7, fontSize: 13, fontWeight: 600 }}>Sign out</button>
        </div>
      </div>
    );
  }
  if (!loaded) return <div style={{ padding: 40, textAlign: "center", color: "#9CA3AF" }}>Loading shipments…</div>;

  return (
    <div style={{ minHeight: "100vh", background: "#FAFAFA", color: "#1F2937", fontFamily: "'Inter', system-ui, sans-serif" }}>
      <style>{`
        * { box-sizing: border-box; }
        .btn { cursor: pointer; border: none; transition: filter .15s ease, transform .1s ease; }
        .btn:hover { filter: brightness(0.95); }
        .btn:active { transform: scale(0.98); }
        input, select, textarea { background: #fff; border: 1px solid #E5E7EB; color: #1F2937; border-radius: 6px; padding: 8px 10px; font-size: 13px; font-family: inherit; }
        input:focus, select:focus, textarea:focus { outline: 2px solid #DC2626; outline-offset: 1px; }
      `}</style>

      <header style={{ borderBottom: "1px solid #E5E7EB", background: "#fff", padding: "16px 20px", display: "flex", alignItems: "center", justifyContent: "space-between", flexWrap: "wrap", gap: 12 }}>
        <div style={{ display: "flex", alignItems: "center", gap: 12 }}>
          <div style={{ width: 36, height: 36, borderRadius: 8, background: "#DC2626", display: "flex", alignItems: "center", justifyContent: "center", fontSize: 17 }}>📦</div>
          <div>
            <div style={{ fontWeight: 700, fontSize: 16, color: "#111827" }}>{me.name}{isAdmin && <ShieldCheck size={13} style={{ marginLeft: 6, verticalAlign: "-2px" }} color="#DC2626" />}</div>
            <div style={{ fontSize: 11, color: "#9CA3AF" }}>{shipments.length} shipments · {Object.keys(customers).length} customers</div>
          </div>
        </div>
        <div style={{ display: "flex", gap: 8, alignItems: "center", flexWrap: "wrap" }}>
          {notifPermission === "default" && (
            <button className="btn" onClick={requestNotifPermission} title="Get a notification even when this tab isn't focused" style={{ display: "flex", alignItems: "center", gap: 6, background: "#fff", color: "#6B7280", padding: "8px 10px", borderRadius: 7, fontSize: 11.5, fontWeight: 600, border: "1px solid #E5E7EB" }}>
              <BellRing size={13} /> Enable alerts
            </button>
          )}
          <div style={{ position: "relative" }}>
            <button className="btn" onClick={() => setShowNotifPanel((v) => !v)} style={{ display: "flex", alignItems: "center", gap: 6, background: myUnreadCount > 0 ? "#FEF2F2" : "#fff", color: myUnreadCount > 0 ? "#DC2626" : "#374151", padding: "8px 12px", borderRadius: 7, fontSize: 12, fontWeight: 700, border: `1px solid ${myUnreadCount > 0 ? "#FCA5A5" : "#E5E7EB"}` }}>
              <Bell size={14} /> {myUnreadCount > 0 ? myUnreadCount : ""}
            </button>
            {showNotifPanel && <NotificationPanel notifications={notifications} onClose={() => setShowNotifPanel(false)} onMarkRead={markNotificationRead} onMarkAllRead={markAllNotificationsRead} />}
          </div>
          {atRisk.length > 0 && (
            <button className="btn" onClick={() => setShowAlert(true)} style={{ display: "flex", alignItems: "center", gap: 6, background: "#FEF2F2", color: "#DC2626", padding: "8px 12px", borderRadius: 7, fontSize: 12, fontWeight: 700, border: "1px solid #FCA5A5" }}>
              <AlertTriangle size={14} /> {atRisk.length} due soon
            </button>
          )}
          <button className="btn" onClick={downloadMyShipments} style={{ display: "flex", alignItems: "center", gap: 6, background: "#fff", color: "#374151", padding: "9px 14px", borderRadius: 7, fontSize: 13, fontWeight: 600, border: "1px solid #E5E7EB" }}>
            <FileDown size={14} /> My shipments (.xlsx)
          </button>
          {isAdmin && (
            <>
              <input id="excel-upload-input" type="file" accept=".xlsx,.xls" onChange={handleFile} style={{ position: "absolute", width: 1, height: 1, padding: 0, margin: -1, overflow: "hidden", clip: "rect(0,0,0,0)", border: 0 }} />
              <label htmlFor="excel-upload-input" className="btn" style={{ display: "inline-flex", alignItems: "center", gap: 6, background: "#DC2626", color: "#fff", padding: "9px 14px", borderRadius: 7, fontSize: 13, fontWeight: 600, cursor: "pointer" }}>
                <Upload size={14} /> Import Excel
              </label>
            </>
          )}
          <button className="btn" onClick={() => supabase.auth.signOut()} title="Sign out" style={{ display: "flex", alignItems: "center", gap: 6, background: "transparent", color: "#9CA3AF", padding: "9px 10px", borderRadius: 7, fontSize: 12 }}>
            <LogOut size={14} />
          </button>
        </div>
      </header>

      <main style={{ padding: "20px", maxWidth: 1100, margin: "0 auto" }}>
        {error && <div style={{ background: "#FEF2F2", color: "#B91C1C", padding: "10px 14px", borderRadius: 8, marginBottom: 16, fontSize: 13, border: "1px solid #FCA5A5" }}>{error}</div>}
        {importSummary && (
          <div style={{ background: "#F0FDF4", color: "#15803D", padding: "10px 14px", borderRadius: 8, marginBottom: 16, fontSize: 13, border: "1px solid #86EFAC" }}>
            Imported {importSummary.added} shipment{importSummary.added !== 1 ? "s" : ""} from {importSummary.total} rows in sheet "{importSummary.sheetUsed}"
            {importSummary.skipped > 0 ? ` (${importSummary.skipped} already existed, skipped)` : ""}
            {importSummary.newCustomers > 0 ? ` · ${importSummary.newCustomers} new customer${importSummary.newCustomers !== 1 ? "s" : ""} detected — assign a coordinator below` : ""}.
          </div>
        )}

        <div style={{ display: "flex", gap: 10, marginBottom: 16, flexWrap: "wrap", alignItems: "center" }}>
          <select value={viewMode} onChange={(e) => setViewMode(e.target.value)} style={{ width: 190 }}>
            <option value="mine">My customers ({myAssignedCount})</option>
            <option value="all">All customers</option>
          </select>
        </div>

        {grouped.length === 0 ? (
          <div style={{ border: "1px dashed #E5E7EB", borderRadius: 12, padding: "48px 20px", textAlign: "center", color: "#9CA3AF" }}>
            {viewMode === "mine" ? 'No customers assigned to you yet — switch to "All customers" or ask an admin to assign some.' : isAdmin ? 'No shipments yet. Tap "Import Excel" to get started.' : "Ask your admin to import the master file."}
          </div>
        ) : (
          <div style={{ display: "flex", flexDirection: "column", gap: 10 }}>
            {grouped.map((g) => (
              <CustomerGroup
                key={g.key}
                group={g}
                expanded={!!expanded[g.key]}
                templatesOpenState={!!templatesOpen[g.key]}
                onToggle={() => toggleExpand(g.key)}
                onToggleTemplates={() => toggleTemplates(g.key)}
                onUpdateEmail={(email) => updateCustomerEmail(g.key, email)}
                onToggleAck={toggleAck}
                onToggleDelivered={toggleDelivered}
                onRemove={isAdmin ? removeShipment : null}
                onEdit={setEditingShipment}
                onCopyReminder={copyReminder}
                coordinators={coordinators}
                isAdmin={isAdmin}
                assignOpen={assignOpenKey === g.key}
                onToggleAssignOpen={() => setAssignOpenKey((k) => (k === g.key ? null : g.key))}
                onToggleCoordinator={(id) => toggleCoordinatorAssignment(g.key, id)}
                onAddTemplate={() => setTemplateForm({ customerKey: g.key, id: null, title: "", subject: "", body: "" })}
                onEditTemplate={(t) => setTemplateForm({ customerKey: g.key, ...t })}
                onDeleteTemplate={(id) => deleteTemplate(g.key, id)}
              />
            ))}
          </div>
        )}
      </main>

      {showAlert && atRisk.length > 0 && (
        <Modal onClose={() => { setShowAlert(false); setDismissedThisSession(true); }}>
          <div style={{ display: "flex", alignItems: "center", gap: 10, marginBottom: 14 }}>
            <div style={{ width: 32, height: 32, borderRadius: 8, background: "#FEF2F2", display: "flex", alignItems: "center", justifyContent: "center" }}><AlertTriangle size={16} color="#DC2626" /></div>
            <div style={{ fontWeight: 700, fontSize: 15, color: "#111827" }}>{atRisk.length} shipment{atRisk.length > 1 ? "s" : ""} need attention{isAdmin ? " (company-wide)" : ""}</div>
          </div>
          <div style={{ display: "flex", flexDirection: "column", gap: 8, maxHeight: 320, overflowY: "auto" }}>
            {atRisk.map((s) => (
              <div key={s.id} style={{ display: "flex", justifyContent: "space-between", alignItems: "center", background: "#F9FAFB", border: "1px solid #E5E7EB", borderRadius: 8, padding: "9px 11px", gap: 10 }}>
                <div>
                  <div style={{ fontWeight: 600, fontSize: 12, color: "#111827" }}>{s.shipment_no}</div>
                  <div style={{ fontSize: 11.5, color: "#6B7280" }}>{s.customerName} · arrives {s.arrival_date}</div>
                </div>
                <button className="btn" onClick={() => toggleAck(s.id)} style={{ background: "#DC2626", color: "#fff", padding: "6px 10px", borderRadius: 6, fontSize: 11, fontWeight: 600 }}>DO Made</button>
              </div>
            ))}
          </div>
          <button className="btn" onClick={() => { setShowAlert(false); setDismissedThisSession(true); }} style={{ marginTop: 14, width: "100%", background: "#F3F4F6", color: "#374151", padding: "9px", borderRadius: 7, fontWeight: 600, fontSize: 13 }}>Close for now</button>
        </Modal>
      )}

      {templateForm && (
        <Modal onClose={() => setTemplateForm(null)}>
          <div style={{ fontWeight: 700, fontSize: 15, marginBottom: 14, color: "#111827" }}>{templateForm.id ? "Edit" : "New"} prompt email</div>
          <div style={{ display: "flex", flexDirection: "column", gap: 10 }}>
            <input value={templateForm.title} onChange={(e) => setTemplateForm({ ...templateForm, title: e.target.value })} placeholder="Title, e.g. ETA delay notice" style={{ width: "100%" }} />
            <input value={templateForm.subject} onChange={(e) => setTemplateForm({ ...templateForm, subject: e.target.value })} placeholder="Subject" style={{ width: "100%" }} />
            <textarea value={templateForm.body} onChange={(e) => setTemplateForm({ ...templateForm, body: e.target.value })} rows={5} placeholder="Body" style={{ width: "100%", resize: "vertical" }} />
            <button className="btn" onClick={saveTemplate} style={{ background: "#DC2626", color: "#fff", padding: "10px", borderRadius: 7, fontWeight: 600, fontSize: 13 }}>Save template</button>
          </div>
        </Modal>
      )}

      {editingShipment && <EditShipmentModal shipment={editingShipment} onSave={saveShipmentEdit} onClose={() => setEditingShipment(null)} />}
    </div>
  );
}

// ---- subcomponents -------------------------------------------------------

function CustomerGroup({ group, expanded, templatesOpenState, onToggle, onToggleTemplates, onUpdateEmail, onToggleAck, onToggleDelivered, onRemove, onEdit, coordinators, isAdmin, assignOpen, onToggleAssignOpen, onToggleCoordinator, onAddTemplate, onEditTemplate, onDeleteTemplate }) {
  const { customer, shipments } = group;
  const templates = customer.templates || [];
  const assignedNames = coordinators.filter((c) => (customer.coordinators || []).includes(c.id)).map((c) => c.name);
  const atRiskCount = shipments.filter((s) => !s.delivered && !s.reminder_acknowledged && s.daysLeft !== null && s.daysLeft <= 15).length;

  return (
    <div style={{ border: "1px solid #E5E7EB", borderRadius: 10, background: "#fff", overflow: "hidden" }}>
      <div onClick={onToggle} style={{ display: "flex", alignItems: "center", justifyContent: "space-between", padding: "12px 14px", cursor: "pointer", gap: 10, flexWrap: "wrap" }}>
        <div style={{ display: "flex", alignItems: "center", gap: 10, minWidth: 0 }}>
          {expanded ? <ChevronDown size={16} color="#9CA3AF" /> : <ChevronRight size={16} color="#9CA3AF" />}
          <div style={{ fontWeight: 600, fontSize: 14, color: "#111827" }}>{customer.canonicalName}</div>
          <span style={{ fontSize: 11, color: "#9CA3AF" }}>({shipments.length})</span>
          {atRiskCount > 0 && <span style={{ background: "#FEF2F2", color: "#DC2626", fontSize: 11, padding: "2px 7px", borderRadius: 5, fontWeight: 700, border: "1px solid #FCA5A5" }}>{atRiskCount} due soon</span>}
        </div>
        <div onClick={(e) => e.stopPropagation()} style={{ display: "flex", gap: 8, alignItems: "center", flexWrap: "wrap" }}>
          <input placeholder="customer email" value={customer.email || ""} onChange={(e) => onUpdateEmail(e.target.value)} style={{ width: 170 }} />
          {customer.email && <a href={`mailto:${customer.email}`} style={{ color: "#DC2626", fontSize: 13 }}><Mail size={13} /></a>}
          {isAdmin ? (
            <div style={{ position: "relative" }}>
              <button className="btn" onClick={onToggleAssignOpen} style={{ display: "flex", alignItems: "center", gap: 6, background: "#fff", color: "#374151", padding: "6px 10px", borderRadius: 6, fontSize: 12, border: "1px solid #E5E7EB", maxWidth: 190 }}>
                <UserPlus size={13} />
                <span style={{ overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap" }}>{assignedNames.length > 0 ? assignedNames.join(", ") : "Unassigned"}</span>
              </button>
              {assignOpen && (
                <div style={{ position: "absolute", top: "110%", right: 0, background: "#fff", border: "1px solid #E5E7EB", borderRadius: 8, boxShadow: "0 8px 24px rgba(0,0,0,0.12)", padding: 8, zIndex: 20, minWidth: 170 }}>
                  {coordinators.map((c) => (
                    <label key={c.id} style={{ display: "flex", alignItems: "center", gap: 8, padding: "6px 6px", fontSize: 13, cursor: "pointer" }}>
                      <input type="checkbox" checked={(customer.coordinators || []).includes(c.id)} onChange={() => onToggleCoordinator(c.id)} />
                      {c.name}
                    </label>
                  ))}
                </div>
              )}
            </div>
          ) : (
            <span style={{ fontSize: 12, color: "#6B7280" }}>{assignedNames.length > 0 ? assignedNames.join(", ") : "Unassigned"}</span>
          )}
        </div>
      </div>

      {expanded && (
        <div style={{ borderTop: "1px solid #E5E7EB" }}>
          <div style={{ background: "#FAFAFA", padding: "10px 14px", borderBottom: "1px solid #E5E7EB" }}>
            <div onClick={onToggleTemplates} style={{ display: "flex", alignItems: "center", justifyContent: "space-between", cursor: "pointer" }}>
              <div style={{ display: "flex", alignItems: "center", gap: 6, fontSize: 12, fontWeight: 600, color: "#374151" }}>
                {templatesOpenState ? <ChevronDown size={13} /> : <ChevronRight size={13} />}
                Prompt emails ({templates.length})
              </div>
              {isAdmin && <button className="btn" onClick={(e) => { e.stopPropagation(); onAddTemplate(); }} style={{ display: "flex", alignItems: "center", gap: 4, background: "#fff", color: "#DC2626", border: "1px solid #FCA5A5", padding: "4px 9px", borderRadius: 6, fontSize: 11, fontWeight: 600 }}><Plus size={12} /> New</button>}
            </div>
            {templatesOpenState && (
              <div style={{ marginTop: 10, display: "flex", flexDirection: "column", gap: 6 }}>
                {templates.length === 0 ? <div style={{ fontSize: 12, color: "#9CA3AF" }}>None yet.</div> : templates.map((t) => (
                  <div key={t.id} style={{ display: "flex", justifyContent: "space-between", alignItems: "center", background: "#fff", border: "1px solid #E5E7EB", borderRadius: 7, padding: "8px 10px", gap: 8 }}>
                    <div style={{ minWidth: 0 }}><div style={{ fontSize: 12.5, fontWeight: 600, color: "#111827" }}>{t.title}</div></div>
                    <div style={{ display: "flex", gap: 6, flexShrink: 0 }}>
                      {customer.email && <a href={`mailto:${customer.email}?subject=${encodeURIComponent(t.subject || "")}&body=${encodeURIComponent(t.body || "")}`} style={{ color: "#DC2626", fontSize: 11 }}>Use</a>}
                      {isAdmin && <><button className="btn" onClick={() => onEditTemplate(t)} style={{ background: "transparent", color: "#9CA3AF", fontSize: 11 }}>Edit</button><button className="btn" onClick={() => onDeleteTemplate(t.id)} style={{ background: "transparent", color: "#9CA3AF", fontSize: 11 }}>Delete</button></>}
                    </div>
                  </div>
                ))}
              </div>
            )}
          </div>
          <div style={{ overflowX: "auto" }}>
            <table style={{ width: "100%", borderCollapse: "collapse" }}>
              <thead><tr>{["", "Shipment No.", "Arrival", "Carrier", "Status", ""].map((h) => <th key={h} style={{ textAlign: "left", fontSize: 10.5, color: "#9CA3AF", padding: "8px 12px", borderBottom: "1px solid #E5E7EB", whiteSpace: "nowrap" }}>{h}</th>)}</tr></thead>
              <tbody>
                {shipments.map((s) => (
                  <ShipmentRow key={s.id} s={s} onToggleAck={onToggleAck} onEdit={onEdit} onRemove={onRemove} />
                ))}
              </tbody>
            </table>
          </div>
        </div>
      )}
    </div>
  );
}

function ShipmentRow({ s, onToggleAck, onEdit, onRemove }) {
  const [open, setOpen] = useState(false);
  const alreadyShown = new Set([
    "Shipment No.", "HBL No.", "MBL No.", "MBL No", "Onb. Date", "ARR Date",
    "SHPR Name", "CNEE  Name", "CNEE Name", "Carrier Name", "3P Letter ", "3P Letter",
    "Do Status", "Do shared date", "Invoice status", "Remark",
  ]);
  const extraEntries = s.raw_details
    ? Object.entries(s.raw_details).filter(([k, v]) => !alreadyShown.has(k) && v !== null && v !== undefined && v !== "")
    : [];
  return (
    <>
      <tr onClick={() => setOpen((v) => !v)} style={{ cursor: "pointer" }}>
        <td style={{ padding: "9px 12px", fontSize: 12.5, borderBottom: "1px solid #F3F4F6", width: 24 }}>{open ? <ChevronDown size={13} color="#9CA3AF" /> : <ChevronRight size={13} color="#9CA3AF" />}</td>
        <td style={{ padding: "9px 12px", fontSize: 12.5, borderBottom: "1px solid #F3F4F6" }}>{s.shipment_no}</td>
        <td style={{ padding: "9px 12px", fontSize: 12.5, borderBottom: "1px solid #F3F4F6" }}>{s.arrival_date || "—"}</td>
        <td style={{ padding: "9px 12px", fontSize: 12.5, borderBottom: "1px solid #F3F4F6", maxWidth: 130, overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap" }}>{s.carrier_name}</td>
        <td style={{ padding: "9px 12px", fontSize: 12.5, borderBottom: "1px solid #F3F4F6" }}><StatusBadge status={s.status} /></td>
        <td style={{ padding: "9px 12px", fontSize: 12.5, borderBottom: "1px solid #F3F4F6" }} onClick={(e) => e.stopPropagation()}>
          <div style={{ display: "flex", gap: 5 }}>
            <button className="btn" onClick={() => onToggleAck(s.id)} style={{ background: s.reminder_acknowledged ? "#F0FDF4" : "#F3F4F6", color: s.reminder_acknowledged ? "#15803D" : "#6B7280", padding: "5px 8px", borderRadius: 6, fontSize: 10.5, fontWeight: 600, whiteSpace: "nowrap" }}>{s.reminder_acknowledged ? "✓ DO Made" : "DO Made"}</button>
            <button className="btn" onClick={() => onEdit(s)} style={{ background: "transparent", color: "#9CA3AF", padding: 5, fontSize: 10.5 }}>Edit</button>
            {onRemove && <button className="btn" onClick={() => onRemove(s.id)} style={{ background: "transparent", color: "#9CA3AF", padding: 5 }}><Trash2 size={12} /></button>}
          </div>
        </td>
      </tr>
      {open && (
        <tr>
          <td colSpan={6} style={{ padding: "10px 14px 14px 34px", background: "#FAFAFA", borderBottom: "1px solid #F3F4F6" }}>
            <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fill, minmax(150px, 1fr))", gap: "6px 16px" }}>
              {extraEntries.length === 0 ? (
                <div style={{ fontSize: 11.5, color: "#9CA3AF" }}>No extra details on this shipment.</div>
              ) : extraEntries.map(([k, v]) => (
                <div key={k} style={{ fontSize: 11 }}>
                  <div style={{ color: "#9CA3AF", fontSize: 10 }}>{k.trim()}</div>
                  <div style={{ color: "#374151", fontWeight: 500 }}>{String(v)}</div>
                </div>
              ))}
            </div>
          </td>
        </tr>
      )}
    </>
  );
}

function NotificationPanel({ notifications, onClose, onMarkRead, onMarkAllRead }) {
  const sorted = [...notifications].sort((a, b) => b.created_at.localeCompare(a.created_at));
  const unread = sorted.filter((n) => !n.is_read).length;
  return (
    <>
      <div onClick={onClose} style={{ position: "fixed", inset: 0, zIndex: 19 }} />
      <div style={{ position: "absolute", top: "110%", right: 0, background: "#fff", border: "1px solid #E5E7EB", borderRadius: 10, boxShadow: "0 10px 30px rgba(0,0,0,0.15)", width: 320, maxHeight: 380, display: "flex", flexDirection: "column", zIndex: 21 }}>
        <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between", padding: "10px 12px", borderBottom: "1px solid #F3F4F6" }}>
          <div style={{ fontWeight: 700, fontSize: 13, color: "#111827" }}>Notifications</div>
          {unread > 0 && <button className="btn" onClick={onMarkAllRead} style={{ display: "flex", alignItems: "center", gap: 4, background: "transparent", color: "#DC2626", fontSize: 11.5, fontWeight: 600 }}><CheckCheck size={12} /> Mark all read</button>}
        </div>
        <div style={{ overflowY: "auto" }}>
          {sorted.length === 0 ? <div style={{ padding: "24px 14px", textAlign: "center", color: "#9CA3AF", fontSize: 12.5 }}>No notifications yet.</div> : sorted.map((n) => (
            <div key={n.id} onClick={() => onMarkRead(n.id)} style={{ display: "flex", gap: 9, padding: "10px 12px", borderBottom: "1px solid #F9FAFB", cursor: "pointer", background: n.is_read ? "#fff" : "#FEF9F9" }}>
              <div style={{ width: 7, height: 7, borderRadius: 99, background: n.is_read ? "transparent" : "#DC2626", marginTop: 5, flexShrink: 0 }} />
              <div style={{ minWidth: 0 }}>
                <div style={{ fontSize: 12.5, fontWeight: 600, color: "#111827" }}>{n.title}</div>
                <div style={{ fontSize: 11.5, color: "#6B7280", marginTop: 1 }}>{n.body}</div>
                <div style={{ fontSize: 10.5, color: "#B0B5BE", marginTop: 3 }}>{timeAgo(n.created_at)}</div>
              </div>
            </div>
          ))}
        </div>
      </div>
    </>
  );
}

function EditShipmentModal({ shipment, onSave, onClose }) {
  const [form, setForm] = useState({ ...shipment });
  const [showAllDetails, setShowAllDetails] = useState(false);
  const rawEntries = shipment.raw_details
    ? Object.entries(shipment.raw_details).filter(([, v]) => v !== null && v !== undefined && v !== "")
    : [];
  return (
    <Modal onClose={onClose}>
      <div style={{ fontWeight: 700, fontSize: 15, marginBottom: 14, color: "#111827" }}>Edit shipment {shipment.shipment_no}</div>
      <div style={{ display: "flex", flexDirection: "column", gap: 10 }}>
        <input type="date" value={form.arrival_date || ""} onChange={(e) => setForm({ ...form, arrival_date: e.target.value })} style={{ width: "100%" }} />
        <input value={form.tracking_link || ""} onChange={(e) => setForm({ ...form, tracking_link: e.target.value })} placeholder="Tracking link" style={{ width: "100%" }} />
        <textarea value={form.remark || ""} onChange={(e) => setForm({ ...form, remark: e.target.value })} rows={3} placeholder="Remark" style={{ width: "100%", resize: "vertical" }} />
        <button className="btn" onClick={() => onSave(form)} style={{ background: "#DC2626", color: "#fff", padding: "10px", borderRadius: 7, fontWeight: 600, fontSize: 13 }}>Save changes</button>
      </div>
      {rawEntries.length > 0 && (
        <div style={{ marginTop: 16, borderTop: "1px solid #E5E7EB", paddingTop: 12 }}>
          <button className="btn" onClick={() => setShowAllDetails((v) => !v)} style={{ display: "flex", alignItems: "center", gap: 6, background: "transparent", color: "#6B7280", fontSize: 12, fontWeight: 600 }}>
            {showAllDetails ? <ChevronDown size={13} /> : <ChevronRight size={13} />}
            All details from the master file ({rawEntries.length})
          </button>
          {showAllDetails && (
            <div style={{ marginTop: 8, maxHeight: 260, overflowY: "auto", background: "#FAFAFA", border: "1px solid #E5E7EB", borderRadius: 8, padding: 10 }}>
              {rawEntries.map(([k, v]) => (
                <div key={k} style={{ display: "flex", justifyContent: "space-between", gap: 10, fontSize: 11.5, padding: "3px 0", borderBottom: "1px solid #F3F4F6" }}>
                  <span style={{ color: "#9CA3AF", flexShrink: 0 }}>{k}</span>
                  <span style={{ color: "#374151", textAlign: "right" }}>{String(v)}</span>
                </div>
              ))}
            </div>
          )}
        </div>
      )}
    </Modal>
  );
}

function StatusBadge({ status }) {
  const map = { ok: { bg: "#F0FDF4", fg: "#15803D" }, upcoming: { bg: "#EFF6FF", fg: "#2563EB" }, warning: { bg: "#FFFBEB", fg: "#B45309" }, critical: { bg: "#FEF2F2", fg: "#DC2626" }, overdue: { bg: "#FEE2E2", fg: "#991B1B" }, delivered: { bg: "#F3F4F6", fg: "#6B7280" }, acked: { bg: "#F3F4F6", fg: "#9CA3AF" }, unknown: { bg: "#F3F4F6", fg: "#9CA3AF" } };
  const c = map[status.tone] || map.unknown;
  return <span style={{ background: c.bg, color: c.fg, padding: "3px 8px", borderRadius: 5, fontSize: 11, fontWeight: 700, whiteSpace: "nowrap" }}>{status.label}</span>;
}

function Modal({ children, onClose }) {
  return (
    <div onClick={onClose} style={{ position: "fixed", inset: 0, background: "rgba(17,24,39,0.5)", display: "flex", alignItems: "center", justifyContent: "center", padding: 20, zIndex: 50 }}>
      <div onClick={(e) => e.stopPropagation()} style={{ background: "#fff", border: "1px solid #E5E7EB", borderRadius: 14, padding: 22, width: "100%", maxWidth: 460, maxHeight: "90vh", overflowY: "auto", boxShadow: "0 10px 40px rgba(0,0,0,0.15)" }}>
        {children}
      </div>
    </div>
  );
        }
