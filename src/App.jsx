import React, { useState, useEffect, useMemo, useCallback, useRef } from "react";
import {
  LayoutDashboard, Users, CalendarDays, FileBarChart2, Settings as SettingsIcon,
  Search, Plus, Pencil, Trash2, Download, Printer, X, Clock, ChevronLeft,
  ChevronRight, AlertTriangle, Upload, Check, UserCog, LogIn as LogInIcon
} from "lucide-react";
import * as XLSX from "xlsx";
import Papa from "papaparse";
import {
  BarChart, Bar, XAxis, YAxis, CartesianGrid, Tooltip, ResponsiveContainer, Cell
} from "recharts";

/* =========================================================================
   PUNCHBOOK — multi-employee attendance ledger
   Data model (persistent, shared across everyone who opens this artifact):
     "config"              -> { employees:[], shifts:[], settings:{} }
     "attn:{empId}:{YYYY-MM}" -> { "YYYY-MM-DD": dayRecord, ... }
   All calculations (late-in, early-out, work hours, OT, %) are DERIVED
   live from the raw punches on every render, so editing any past record
   instantly recalculates every summary, dashboard and report that uses it.
   ========================================================================= */

const FONT_IMPORT_ID = "punchbook-fonts";

const DEFAULT_SETTINGS = { graceMinutes: 10, fullDayHours: 8, halfDayHours: 4 };
const DEFAULT_SHIFTS = [
  { code: "GEN", name: "General Shift", start: "09:00", end: "18:00" },
];
const ROLES = ["Admin", "Manager", "Viewer"];

// Shift list transcribed from the "Sangam Electronics co. — Shifts" report.
// Full Day / Half Day are the minimum worked hours (H:MM) converted to decimal hours.
const SANGAM_SHIFTS = [
  { code: "A1", name: "Shift A1", start: "09:00", end: "17:00", breakStart: "", breakEnd: "", fullDayHours: 7, halfDayHours: 3.5 },
  { code: "A2", name: "Shift A2", start: "10:30", end: "18:00", breakStart: "", breakEnd: "", fullDayHours: 7, halfDayHours: 3 },
  { code: "A3", name: "Shift A3", start: "09:15", end: "18:00", breakStart: "", breakEnd: "", fullDayHours: 7.75, halfDayHours: 4 },
  { code: "A4", name: "Shift A4", start: "08:00", end: "18:00", breakStart: "13:00", breakEnd: "14:00", fullDayHours: 8, halfDayHours: 4 },
  { code: "A5", name: "Shift A5", start: "08:30", end: "17:30", breakStart: "", breakEnd: "", fullDayHours: 8, halfDayHours: 4.5 },
  { code: "A6", name: "Shift A6", start: "08:45", end: "17:00", breakStart: "13:00", breakEnd: "13:30", fullDayHours: 7, halfDayHours: 3.5 },
  { code: "A7", name: "Shift A7", start: "08:00", end: "17:00", breakStart: "13:00", breakEnd: "13:30", fullDayHours: 7.5, halfDayHours: 3 },
  { code: "A8", name: "Shift A8", start: "08:00", end: "16:30", breakStart: "13:00", breakEnd: "13:30", fullDayHours: 7.5, halfDayHours: 3.5 },
  { code: "A9", name: "Shift A9", start: "08:45", end: "18:00", breakStart: "13:00", breakEnd: "13:30", fullDayHours: 8, halfDayHours: 4 },
  { code: "B1", name: "Shift B1", start: "10:15", end: "18:00", breakStart: "", breakEnd: "", fullDayHours: 6, halfDayHours: 2.75 },
  { code: "B2", name: "Shift B2", start: "08:30", end: "18:30", breakStart: "", breakEnd: "", fullDayHours: 8, halfDayHours: 4 },
  { code: "G2", name: "Shift G2", start: "08:30", end: "17:00", breakStart: "", breakEnd: "", fullDayHours: 7.75, halfDayHours: 4 },
  { code: "G3", name: "Shift G3", start: "09:00", end: "17:30", breakStart: "", breakEnd: "", fullDayHours: 8, halfDayHours: 4 },
  { code: "G4", name: "Shift G4", start: "09:00", end: "18:00", breakStart: "", breakEnd: "", fullDayHours: 7.75, halfDayHours: 4 },
  { code: "G5", name: "Shift G5", start: "10:00", end: "18:00", breakStart: "", breakEnd: "", fullDayHours: 6.75, halfDayHours: 4 },
  { code: "G6", name: "Shift G6", start: "10:00", end: "18:30", breakStart: "", breakEnd: "", fullDayHours: 7.75, halfDayHours: 4 },
  { code: "G7", name: "Shift G7", start: "09:30", end: "18:00", breakStart: "", breakEnd: "", fullDayHours: 7.25, halfDayHours: 4 },
  { code: "G8", name: "Shift G8", start: "09:00", end: "13:00", breakStart: "", breakEnd: "", fullDayHours: 4, halfDayHours: 2 },
  { code: "G9", name: "Shift G9", start: "09:00", end: "18:30", breakStart: "", breakEnd: "", fullDayHours: 8, halfDayHours: 4 },
  { code: "GS", name: "Shift GS", start: "08:00", end: "18:30", breakStart: "", breakEnd: "", fullDayHours: 9, halfDayHours: 4.5 },
];

/* ---------------------------- utility helpers --------------------------- */

const uid = (p = "id") => `${p}_${Date.now().toString(36)}_${Math.random().toString(36).slice(2, 8)}`;

function pad2(n) { return String(n).padStart(2, "0"); }

function isoDate(d) { return `${d.getFullYear()}-${pad2(d.getMonth() + 1)}-${pad2(d.getDate())}`; }

function yyyymm(iso) { return iso.slice(0, 7); }

function parseTimeToMinutes(str) {
  if (!str) return null;
  const m = /^(\d{1,2}):(\d{2})$/.exec(str.trim());
  if (!m) return null;
  return parseInt(m[1], 10) * 60 + parseInt(m[2], 10);
}

function hoursToHM(hrs) {
  if (hrs === null || hrs === undefined || isNaN(hrs)) return "—";
  const totalMin = Math.round(hrs * 60);
  const sign = totalMin < 0 ? "-" : "";
  const abs = Math.abs(totalMin);
  const h = Math.floor(abs / 60);
  const m = abs % 60;
  return `${sign}${h}:${pad2(m)}`;
}

function round2(n) { return Math.round(n * 100) / 100; }

function monthLabel(ym) {
  const [y, m] = ym.split("-").map(Number);
  return new Date(y, m - 1, 1).toLocaleDateString(undefined, { month: "long", year: "numeric" });
}

function daysInMonth(ym) {
  const [y, m] = ym.split("-").map(Number);
  const count = new Date(y, m, 0).getDate();
  const out = [];
  for (let d = 1; d <= count; d++) out.push(`${ym}-${pad2(d)}`);
  return out;
}

function shiftYm(ym, delta) {
  const [y, m] = ym.split("-").map(Number);
  const d = new Date(y, m - 1 + delta, 1);
  return `${d.getFullYear()}-${pad2(d.getMonth() + 1)}`;
}

function monthsBetween(startIso, endIso) {
  const out = [];
  let cur = yyyymm(startIso);
  const end = yyyymm(endIso);
  let guard = 0;
  while (cur <= end && guard < 600) {
    out.push(cur);
    cur = shiftYm(cur, 1);
    guard++;
  }
  return out;
}

function weekdayShort(iso) {
  const [y, m, d] = iso.split("-").map(Number);
  return new Date(y, m - 1, d).toLocaleDateString(undefined, { weekday: "short" });
}

function todayIso() { return isoDate(new Date()); }

function currentYm() { return todayIso().slice(0, 7); }

/* ---------------------------- calculation engine ------------------------ */

function shiftBreakMinutes(shift) {
  if (!shift || !shift.breakStart || !shift.breakEnd) return 0;
  const a = parseTimeToMinutes(shift.breakStart), b = parseTimeToMinutes(shift.breakEnd);
  if (a === null || b === null) return 0;
  let diff = b - a;
  if (diff < 0) diff += 24 * 60;
  return diff;
}

function calcDay(rec, shift, settings) {
  const fullDayH = shift && shift.fullDayHours != null ? shift.fullDayHours : settings.fullDayHours;
  const halfDayH = shift && shift.halfDayHours != null ? shift.halfDayHours : settings.halfDayHours;
  const breakH = (rec.breakMinutes || 0) / 60;
  let workHours = 0;
  if (rec.firstIn && rec.lastOut) {
    const inM = parseTimeToMinutes(rec.firstIn);
    const outM = parseTimeToMinutes(rec.lastOut);
    if (inM !== null && outM !== null) {
      let diff = outM - inM;
      if (diff < 0) diff += 24 * 60; // overnight shift
      workHours = Math.max(0, diff / 60 - breakH);
    }
  }

  let lateInH = 0, earlyOutH = 0;
  const isWorking = rec.dayType !== "weekoff";
  if (isWorking && shift && rec.firstIn) {
    const inM = parseTimeToMinutes(rec.firstIn);
    const startM = parseTimeToMinutes(shift.start);
    if (inM !== null && startM !== null) {
      const raw = Math.max(0, inM - startM);
      lateInH = Math.max(0, raw - settings.graceMinutes) / 60;
    }
  }
  if (isWorking && shift && rec.lastOut) {
    const outM = parseTimeToMinutes(rec.lastOut);
    const endM = parseTimeToMinutes(shift.end);
    if (outM !== null && endM !== null) {
      const raw = Math.max(0, endM - outM);
      earlyOutH = Math.max(0, raw - settings.graceMinutes) / 60;
    }
  }

  let dayStatus;
  if (rec.dayType === "weekoff") dayStatus = "Week Off";
  else if (rec.dayType === "leave") dayStatus = "On Leave";
  else if (!rec.firstIn || !rec.lastOut) dayStatus = "Absent (No Punches)";
  else if (workHours >= fullDayH) dayStatus = "Present";
  else if (workHours >= halfDayH) dayStatus = "Half Day";
  else dayStatus = "Absent (Short Hours)";

  const presentValue = dayStatus === "Present" ? 1 : dayStatus === "Half Day" ? 0.5 : 0;
  const specialAddition = dayStatus === "Absent (Short Hours)" ? workHours : 0;
  const extraWork = dayStatus === "Present" && workHours > fullDayH
    ? workHours - fullDayH : 0;
  const weekOffWorkedHours = rec.dayType === "weekoff" && workHours > 0 ? workHours : 0;
  const totalAddition = extraWork + specialAddition;
  const totalDeduction = lateInH + earlyOutH;
  const net = totalAddition - totalDeduction;

  return {
    ...rec, workHours: round2(workHours), lateInH: round2(lateInH), earlyOutH: round2(earlyOutH),
    dayStatus, presentValue, specialAddition: round2(specialAddition), extraWork: round2(extraWork),
    weekOffWorkedHours: round2(weekOffWorkedHours), totalAddition: round2(totalAddition),
    totalDeduction: round2(totalDeduction), net: round2(net),
  };
}

function summarize(calcRows) {
  const s = {
    totalDays: calcRows.length, presentDays: 0, halfDays: 0, absentShort: 0, absentNoPunch: 0,
    onLeave: 0, weekOffDays: 0, lateCount: 0, earlyCount: 0, totalLateHours: 0, totalEarlyHours: 0,
    totalWorkHours: 0, totalExtraWork: 0, totalSpecialAddition: 0, totalAddition: 0, totalDeduction: 0,
    weekOffWorkedCount: 0, totalWeekOffOT: 0,
  };
  const additionDates = [];   // days "Absent (Worked <threshold)" whose hours are credited to Addition
  const weekOffOtDates = [];  // days worked on a Week Off, credited as OT (1x)
  for (const r of calcRows) {
    if (r.dayStatus === "Present") s.presentDays++;
    else if (r.dayStatus === "Half Day") s.halfDays++;
    else if (r.dayStatus === "Absent (Short Hours)") s.absentShort++;
    else if (r.dayStatus === "Absent (No Punches)") s.absentNoPunch++;
    else if (r.dayStatus === "On Leave") s.onLeave++;
    else if (r.dayStatus === "Week Off") s.weekOffDays++;
    if (r.lateInH > 0) { s.lateCount++; s.totalLateHours += r.lateInH; }
    if (r.earlyOutH > 0) { s.earlyCount++; s.totalEarlyHours += r.earlyOutH; }
    if (r.weekOffWorkedHours > 0) {
      s.weekOffWorkedCount++; s.totalWeekOffOT += r.weekOffWorkedHours;
      weekOffOtDates.push({ date: r.date, hours: round2(r.weekOffWorkedHours) });
    }
    if (r.dayStatus === "Absent (Short Hours)" && r.specialAddition > 0) {
      additionDates.push({ date: r.date, hours: round2(r.specialAddition) });
    }
    s.totalWorkHours += r.workHours;
    s.totalExtraWork += r.extraWork;
    s.totalSpecialAddition += r.specialAddition;
    s.totalAddition += r.totalAddition;
    s.totalDeduction += r.totalDeduction;
  }
  s.totalAbsent = s.absentShort + s.absentNoPunch;
  s.workingDays = s.totalDays - s.weekOffDays - s.onLeave;
  // Matches the Excel Summary sheet's "Total Working Days (Total Days - Week Off)" exactly
  // (that sheet has no separate Leave concept, so it only ever subtracts Week Off).
  s.totalWorkingDaysStrict = s.totalDays - s.weekOffDays;
  s.finalAttendance = s.presentDays + s.halfDays * 0.5;
  s.attendancePct = s.workingDays > 0 ? round2((s.finalAttendance / s.workingDays) * 100) : 0;
  s.netHours = round2(s.totalAddition - s.totalDeduction);
  s.totalLateHours = round2(s.totalLateHours);
  s.totalEarlyHours = round2(s.totalEarlyHours);
  s.totalWorkHours = round2(s.totalWorkHours);
  s.totalExtraWork = round2(s.totalExtraWork);
  s.totalSpecialAddition = round2(s.totalSpecialAddition);
  s.totalAddition = round2(s.totalAddition);
  s.totalDeduction = round2(s.totalDeduction);
  s.totalWeekOffOT = round2(s.totalWeekOffOT);
  s.additionDates = additionDates;
  s.weekOffOtDates = weekOffOtDates;
  // Excel's "Reconciliation Check": Present + Half + Absent vs Total Working Days.
  // On Leave is folded in on the "accounted for" side since this app (unlike the sheet) has a Leave day type.
  const accountedDays = s.presentDays + s.halfDays + s.totalAbsent + s.onLeave;
  s.reconciliationDiff = accountedDays - s.totalWorkingDaysStrict;
  s.reconciliationOk = s.reconciliationDiff === 0;
  return s;
}

const STATUS_STYLE = {
  "Present": { color: "#2F6F4E", label: "PRESENT" },
  "Half Day": { color: "#B4770F", label: "HALF DAY" },
  "Absent (Short Hours)": { color: "#A63A2E", label: "ABSENT" },
  "Absent (No Punches)": { color: "#A63A2E", label: "ABSENT" },
  "Week Off": { color: "#5B6B78", label: "WEEK OFF" },
  "On Leave": { color: "#6B4FA0", label: "ON LEAVE" },
};

/* ------------------------------ stamp badge ------------------------------ */

function Stamp({ status, small }) {
  const st = STATUS_STYLE[status] || { color: "#666", label: status || "—" };
  return (
    <span
      className="stamp"
      style={{
        color: st.color,
        borderColor: st.color,
        fontSize: small ? "9px" : "10.5px",
        padding: small ? "1px 6px" : "2px 9px",
      }}
    >
      {st.label}
    </span>
  );
}

/* ------------------------------ font loader ------------------------------ */

function useFonts() {
  useEffect(() => {
    if (document.getElementById(FONT_IMPORT_ID)) return;
    const style = document.createElement("style");
    style.id = FONT_IMPORT_ID;
    style.innerHTML = `@import url('https://fonts.googleapis.com/css2?family=Zilla+Slab:wght@400;500;600;700&family=Inter:wght@400;500;600;700&family=IBM+Plex+Mono:wght@400;500;600&display=swap');`;
    document.head.appendChild(style);
  }, []);
}

/* ================================ APP ==================================== */

export default function App() {
  useFonts();

  const [loading, setLoading] = useState(true);
  const [error, setError] = useState("");
  const [config, setConfig] = useState({ employees: [], shifts: DEFAULT_SHIFTS, settings: DEFAULT_SETTINGS });
  const [configVersion, setConfigVersion] = useState("new");
  const [attnCache, setAttnCache] = useState({}); // key -> {version, data:{date:rec}}
  const [view, setView] = useState("dashboard");
  const [session, setSession] = useState(null); // {name, role}
  const [toast, setToast] = useState(null);

  const showToast = useCallback((msg) => {
    setToast(msg);
    window.clearTimeout(showToast._t);
    showToast._t = window.setTimeout(() => setToast(null), 2600);
  }, []);

  /* --------------------------- initial load --------------------------- */
  useEffect(() => {
    (async () => {
      try {
        let cfg = null, cfgVersion = "new";
        try {
          const res = await window.storage.get("config", true);
          if (res) { cfg = JSON.parse(res.value); cfgVersion = res.key ? undefined : undefined; }
        } catch (e) { cfg = null; }
        if (!cfg) {
          cfg = { employees: [], shifts: DEFAULT_SHIFTS, settings: DEFAULT_SETTINGS };
          const saved = await window.storage.set("config", JSON.stringify(cfg), true);
          cfgVersion = saved ? "saved" : "new";
        }
        setConfig(cfg);

        // preload current month's attendance for all employees
        const ym = currentYm();
        const entries = await Promise.allSettled(
          (cfg.employees || []).map(async (e) => {
            const key = `attn:${e.id}:${ym}`;
            try {
              const r = await window.storage.get(key, true);
              return [key, r ? JSON.parse(r.value) : {}];
            } catch (e2) { return [key, {}]; }
          })
        );
        const cache = {};
        entries.forEach((r) => { if (r.status === "fulfilled") cache[r.value[0]] = r.value[1]; });
        setAttnCache(cache);
      } catch (e) {
        setError("Could not load stored data. You can still use the app for this session.");
      } finally {
        setLoading(false);
      }
    })();
  }, []);

  /* --------------------------- config persistence --------------------------- */
  const saveConfig = useCallback(async (nextCfg) => {
    setConfig(nextCfg);
    try {
      await window.storage.set("config", JSON.stringify(nextCfg), true);
    } catch (e) {
      showToast("Could not save — check connection.");
    }
  }, [showToast]);

  /* --------------------------- attendance persistence --------------------------- */
  const getMonthData = useCallback((empId, ym) => {
    const key = `attn:${empId}:${ym}`;
    return attnCache[key] || null; // null = not fetched yet
  }, [attnCache]);

  const ensureMonthLoaded = useCallback(async (empId, ym) => {
    const key = `attn:${empId}:${ym}`;
    if (attnCache[key]) return attnCache[key];
    try {
      const r = await window.storage.get(key, true);
      const data = r ? JSON.parse(r.value) : {};
      setAttnCache((prev) => ({ ...prev, [key]: data }));
      return data;
    } catch (e) {
      setAttnCache((prev) => ({ ...prev, [key]: {} }));
      return {};
    }
  }, [attnCache]);

  const upsertRecord = useCallback(async (empId, date, fields) => {
    const ym = yyyymm(date);
    const key = `attn:${empId}:${ym}`;
    const current = attnCache[key] || (await ensureMonthLoaded(empId, ym));
    const nextMonth = { ...current, [date]: { ...(current[date] || {}), ...fields, date, employeeId: empId } };
    setAttnCache((prev) => ({ ...prev, [key]: nextMonth }));
    try {
      await window.storage.set(key, JSON.stringify(nextMonth), true);
      showToast("Record saved — all summaries recalculated.");
    } catch (e) {
      showToast("Save failed — please retry.");
    }
  }, [attnCache, ensureMonthLoaded, showToast]);

  const deleteRecord = useCallback(async (empId, date) => {
    const ym = yyyymm(date);
    const key = `attn:${empId}:${ym}`;
    const current = attnCache[key] || {};
    const nextMonth = { ...current };
    delete nextMonth[date];
    setAttnCache((prev) => ({ ...prev, [key]: nextMonth }));
    try {
      await window.storage.set(key, JSON.stringify(nextMonth), true);
      showToast("Record removed.");
    } catch (e) {
      showToast("Could not remove — please retry.");
    }
  }, [attnCache, showToast]);

  const bulkUpsert = useCallback(async (empId, recordsByDate) => {
    // group by month, merge, persist each month once
    const byMonth = {};
    Object.entries(recordsByDate).forEach(([date, fields]) => {
      const ym = yyyymm(date);
      byMonth[ym] = byMonth[ym] || {};
      byMonth[ym][date] = { ...fields, date, employeeId: empId };
    });
    for (const ym of Object.keys(byMonth)) {
      const key = `attn:${empId}:${ym}`;
      const current = attnCache[key] || (await ensureMonthLoaded(empId, ym));
      const nextMonth = { ...current, ...byMonth[ym] };
      setAttnCache((prev) => ({ ...prev, [key]: nextMonth }));
      try { await window.storage.set(key, JSON.stringify(nextMonth), true); } catch (e) { /* continue */ }
    }
    showToast("Import complete.");
  }, [attnCache, ensureMonthLoaded, showToast]);

  /* --------------------------- employee CRUD --------------------------- */
  const addEmployee = (emp) => {
    const next = { ...config, employees: [...config.employees, { ...emp, id: uid("emp"), active: true }] };
    saveConfig(next);
  };
  const updateEmployee = (id, fields) => {
    const next = { ...config, employees: config.employees.map((e) => (e.id === id ? { ...e, ...fields } : e)) };
    saveConfig(next);
  };
  const deleteEmployee = (id) => {
    const next = { ...config, employees: config.employees.filter((e) => e.id !== id) };
    saveConfig(next);
  };

  const addShift = (shift) => {
    const next = { ...config, shifts: [...config.shifts, { ...shift, code: shift.code || uid("SH") }] };
    saveConfig(next);
  };
  const bulkAddShifts = (list) => {
    const existingCodes = new Set(config.shifts.map((s) => s.code));
    const additions = list.filter((s) => !existingCodes.has(s.code));
    if (additions.length === 0) { showToast("Those shifts are already on the list."); return; }
    const next = { ...config, shifts: [...config.shifts, ...additions] };
    saveConfig(next);
    showToast(`Added ${additions.length} shift${additions.length !== 1 ? "s" : ""}.`);
  };
  const updateShift = (code, fields) => {
    const next = { ...config, shifts: config.shifts.map((s) => (s.code === code ? { ...s, ...fields } : s)) };
    saveConfig(next);
  };
  const deleteShift = (code) => {
    const next = { ...config, shifts: config.shifts.filter((s) => s.code !== code) };
    saveConfig(next);
  };
  const updateSettings = (fields) => {
    const next = { ...config, settings: { ...config.settings, ...fields } };
    saveConfig(next);
  };

  const shiftByCode = useMemo(() => {
    const m = {};
    (config.shifts || []).forEach((s) => { m[s.code] = s; });
    return m;
  }, [config.shifts]);

  const canEdit = session && (session.role === "Admin" || session.role === "Manager");
  const isAdmin = session && session.role === "Admin";

  if (loading) {
    return (
      <div style={{ ...S.appShell, alignItems: "center", justifyContent: "center", display: "flex" }}>
        <div style={{ fontFamily: "'IBM Plex Mono', monospace", color: "#8A8371", letterSpacing: "0.08em" }}>
          OPENING THE LEDGER…
        </div>
      </div>
    );
  }

  if (!session) {
    return <SignIn onSignIn={setSession} />;
  }

  return (
    <div style={S.appShell}>
      <Sidebar view={view} setView={setView} session={session} onSwitchUser={() => setSession(null)} />
      <div style={S.main}>
        <TopBar session={session} error={error} />
        <div style={S.content}>
          {view === "dashboard" && (
            <Dashboard config={config} attnCache={attnCache} ensureMonthLoaded={ensureMonthLoaded}
              shiftByCode={shiftByCode} setView={setView} />
          )}
          {view === "employees" && (
            <EmployeesView config={config} canEdit={canEdit} isAdmin={isAdmin}
              addEmployee={addEmployee} updateEmployee={updateEmployee} deleteEmployee={deleteEmployee} />
          )}
          {view === "attendance" && (
            <AttendanceView config={config} shiftByCode={shiftByCode} attnCache={attnCache}
              ensureMonthLoaded={ensureMonthLoaded} upsertRecord={upsertRecord} deleteRecord={deleteRecord}
              bulkUpsert={bulkUpsert} canEdit={canEdit} />
          )}
          {view === "reports" && (
            <ReportsView config={config} shiftByCode={shiftByCode} ensureMonthLoaded={ensureMonthLoaded}
              attnCache={attnCache} />
          )}
          {view === "settings" && (
            <SettingsView config={config} isAdmin={isAdmin} addShift={addShift} updateShift={updateShift}
              deleteShift={deleteShift} updateSettings={updateSettings} bulkAddShifts={bulkAddShifts} />
          )}
        </div>
      </div>
      {toast && <div style={S.toast}>{toast}</div>}
    </div>
  );
}

/* ============================== SIGN IN =================================== */

function SignIn({ onSignIn }) {
  useFonts();
  const [name, setName] = useState("");
  const [role, setRole] = useState("Admin");
  return (
    <div style={{ ...S.appShell, display: "flex", alignItems: "center", justifyContent: "center" }}>
      <div style={{ width: 380, background: "#fff", border: "1px solid #DED9CA", borderRadius: 4, padding: "34px 30px", boxShadow: "0 1px 2px rgba(20,20,10,0.04)" }}>
        <div style={{ fontFamily: "'Zilla Slab', serif", fontWeight: 700, fontSize: 26, color: "#20242B", letterSpacing: "-0.01em" }}>
          Punchbook
        </div>
        <div style={{ fontFamily: "'IBM Plex Mono', monospace", fontSize: 11, color: "#8A8371", letterSpacing: "0.08em", marginTop: 2, marginBottom: 26 }}>
          THE COMPANY ATTENDANCE LEDGER
        </div>
        <label style={S.label}>Your name</label>
        <input style={S.input} value={name} onChange={(e) => setName(e.target.value)} placeholder="e.g. Priya Shah" />
        <label style={{ ...S.label, marginTop: 14 }}>Your role</label>
        <div style={{ display: "flex", gap: 8, marginTop: 6 }}>
          {ROLES.map((r) => (
            <button key={r} onClick={() => setRole(r)}
              style={{ ...S.roleBtn, ...(role === r ? S.roleBtnActive : {}) }}>
              {r}
            </button>
          ))}
        </div>
        <div style={{ fontSize: 11.5, color: "#8A8371", marginTop: 10, lineHeight: 1.5 }}>
          Admins and Managers can add or edit records. Viewers see dashboards and reports only.
          This is a role-based view switch shared with everyone using this ledger, not a secured login.
        </div>
        <button
          style={{ ...S.primaryBtn, width: "100%", marginTop: 20, justifyContent: "center" }}
          onClick={() => onSignIn({ name: name.trim() || "Guest", role })}
        >
          <LogInIcon size={15} /> Open the ledger
        </button>
      </div>
    </div>
  );
}

/* ============================== SIDEBAR / TOPBAR =========================== */

function Sidebar({ view, setView, session, onSwitchUser }) {
  const items = [
    { key: "dashboard", label: "Dashboard", icon: LayoutDashboard },
    { key: "employees", label: "Employees", icon: Users },
    { key: "attendance", label: "Attendance", icon: CalendarDays },
    { key: "reports", label: "Reports", icon: FileBarChart2 },
    { key: "settings", label: "Settings", icon: SettingsIcon },
  ];
  return (
    <div style={S.sidebar}>
      <div style={{ padding: "22px 18px 18px" }}>
        <div style={{ fontFamily: "'Zilla Slab', serif", fontWeight: 700, fontSize: 21, color: "#F5F2EA", letterSpacing: "-0.01em" }}>
          Punchbook
        </div>
        <div style={{ fontFamily: "'IBM Plex Mono', monospace", fontSize: 9.5, color: "#8F9BAE", letterSpacing: "0.1em", marginTop: 2 }}>
          ATTENDANCE LEDGER
        </div>
      </div>
      <div style={{ flex: 1, padding: "6px 10px" }}>
        {items.map((it) => {
          const Icon = it.icon;
          const active = view === it.key;
          return (
            <div key={it.key} onClick={() => setView(it.key)}
              style={{ ...S.navItem, ...(active ? S.navItemActive : {}) }}>
              <Icon size={16} strokeWidth={2} />
              <span>{it.label}</span>
            </div>
          );
        })}
      </div>
      <div style={{ padding: "14px 18px 20px", borderTop: "1px solid #34405A" }}>
        <div style={{ fontSize: 11.5, color: "#B7C0CF" }}>Signed in as</div>
        <div style={{ fontSize: 13.5, color: "#F5F2EA", fontWeight: 600, marginTop: 2 }}>{session.name}</div>
        <div style={{ fontFamily: "'IBM Plex Mono', monospace", fontSize: 10, color: "#8F9BAE", marginTop: 1 }}>{session.role.toUpperCase()}</div>
        <button onClick={onSwitchUser} style={{ ...S.ghostBtnSmall, marginTop: 10, width: "100%", justifyContent: "center" }}>
          <UserCog size={13} /> Switch user
        </button>
      </div>
    </div>
  );
}

function TopBar({ session, error }) {
  const dt = new Date().toLocaleDateString(undefined, { weekday: "long", year: "numeric", month: "long", day: "numeric" });
  return (
    <div style={S.topbar}>
      <div style={{ fontFamily: "'IBM Plex Mono', monospace", fontSize: 12, color: "#6B6656" }}>{dt}</div>
      {error && (
        <div style={{ display: "flex", alignItems: "center", gap: 6, color: "#A63A2E", fontSize: 12 }}>
          <AlertTriangle size={14} /> {error}
        </div>
      )}
    </div>
  );
}

/* ============================== DASHBOARD =================================== */

function Dashboard({ config, attnCache, ensureMonthLoaded, shiftByCode, setView }) {
  const [ym, setYm] = useState(currentYm());
  const { employees, settings } = config;

  useEffect(() => {
    employees.forEach((e) => ensureMonthLoaded(e.id, ym));
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [ym, employees.length]);

  const perEmployee = useMemo(() => {
    return employees.map((emp) => {
      const key = `attn:${emp.id}:${ym}`;
      const monthData = attnCache[key] || {};
      const rows = Object.values(monthData).map((rec) =>
        calcDay(rec, shiftByCode[rec.shiftCode] || DEFAULT_SHIFTS[0], settings));
      const sum = summarize(rows);
      return { emp, sum, rows };
    });
  }, [employees, attnCache, ym, shiftByCode, settings]);

  const company = useMemo(() => {
    const allRows = perEmployee.flatMap((p) => p.rows);
    return summarize(allRows);
  }, [perEmployee]);

  const trend = useMemo(() => {
    const days = daysInMonth(ym);
    return days.map((d) => {
      let present = 0, absent = 0, half = 0;
      perEmployee.forEach((p) => {
        const r = p.rows.find((x) => x.date === d);
        if (!r) return;
        if (r.dayStatus === "Present") present++;
        else if (r.dayStatus === "Half Day") half++;
        else if (r.dayStatus.startsWith("Absent")) absent++;
      });
      return { day: d.slice(8), present, half, absent };
    });
  }, [perEmployee, ym]);

  return (
    <div>
      <div style={S.rowBetween}>
        <div>
          <div style={S.h1}>Dashboard</div>
          <div style={S.subtle}>Company-wide attendance for {monthLabel(ym)}</div>
        </div>
        <MonthSwitcher ym={ym} setYm={setYm} />
      </div>

      <div style={S.cardGrid}>
        <StatCard label="Employees" value={employees.length} />
        <StatCard label="Attendance %" value={`${company.attendancePct}%`} accent="#2F6F4E" />
        <StatCard label="Late arrivals" value={company.lateCount} sub={`${hoursToHM(company.totalLateHours)} total`} accent="#B4770F" />
        <StatCard label="Early departures" value={company.earlyCount} sub={`${hoursToHM(company.totalEarlyHours)} total`} accent="#B4770F" />
        <StatCard label="Overtime hours" value={hoursToHM(company.totalExtraWork + company.totalWeekOffOT)} accent="#2B3A55" />
        <StatCard label="Absences" value={company.totalAbsent} accent="#A63A2E" />
      </div>

      <div style={S.card}>
        <div style={S.cardTitle}>Daily headcount — {monthLabel(ym)}</div>
        <div style={{ width: "100%", height: 220 }}>
          <ResponsiveContainer>
            <BarChart data={trend} barCategoryGap={2}>
              <CartesianGrid strokeDasharray="2 4" stroke="#E4DFD0" vertical={false} />
              <XAxis dataKey="day" tick={{ fontSize: 10, fontFamily: "IBM Plex Mono" }} stroke="#B8B199" />
              <YAxis tick={{ fontSize: 10, fontFamily: "IBM Plex Mono" }} stroke="#B8B199" allowDecimals={false} />
              <Tooltip contentStyle={{ fontFamily: "Inter", fontSize: 12, borderRadius: 4 }} />
              <Bar dataKey="present" stackId="a" fill="#2F6F4E" name="Present" />
              <Bar dataKey="half" stackId="a" fill="#C98B12" name="Half day" />
              <Bar dataKey="absent" stackId="a" fill="#A63A2E" name="Absent" />
            </BarChart>
          </ResponsiveContainer>
        </div>
      </div>

      <div style={S.card}>
        <div style={S.cardTitle}>By employee</div>
        {employees.length === 0 ? (
          <EmptyState text="No employees yet." actionLabel="Add employees" onAction={() => setView("employees")} />
        ) : (
          <table style={S.table}>
            <thead>
              <tr>
                <th style={S.th}>Employee</th>
                <th style={S.th}>Present</th>
                <th style={S.th}>Half</th>
                <th style={S.th}>Absent</th>
                <th style={S.th}>Late</th>
                <th style={S.th}>Early-out</th>
                <th style={S.th}>OT hrs</th>
                <th style={S.th}>Attendance %</th>
              </tr>
            </thead>
            <tbody>
              {perEmployee.map(({ emp, sum }) => (
                <tr key={emp.id} style={S.trHover} onClick={() => setView("attendance")}>
                  <td style={S.td}><b>{emp.name}</b><div style={{ fontSize: 10.5, color: "#8A8371" }}>{emp.code}</div></td>
                  <td style={S.tdMono}>{sum.presentDays}</td>
                  <td style={S.tdMono}>{sum.halfDays}</td>
                  <td style={S.tdMono}>{sum.totalAbsent}</td>
                  <td style={S.tdMono}>{sum.lateCount}</td>
                  <td style={S.tdMono}>{sum.earlyCount}</td>
                  <td style={S.tdMono}>{hoursToHM(sum.totalExtraWork + sum.totalWeekOffOT)}</td>
                  <td style={S.tdMono}><b>{sum.attendancePct}%</b></td>
                </tr>
              ))}
            </tbody>
          </table>
        )}
      </div>
    </div>
  );
}

function MonthSwitcher({ ym, setYm }) {
  return (
    <div style={{ display: "flex", alignItems: "center", gap: 6, fontFamily: "'IBM Plex Mono', monospace" }}>
      <button style={S.iconBtn} onClick={() => setYm(shiftYm(ym, -1))}><ChevronLeft size={16} /></button>
      <div style={{ fontSize: 13, minWidth: 130, textAlign: "center", color: "#20242B" }}>{monthLabel(ym)}</div>
      <button style={S.iconBtn} onClick={() => setYm(shiftYm(ym, 1))}><ChevronRight size={16} /></button>
    </div>
  );
}

function StatCard({ label, value, sub, accent }) {
  return (
    <div style={S.statCard}>
      <div style={{ fontSize: 11, color: "#8A8371", letterSpacing: "0.04em" }}>{label.toUpperCase()}</div>
      <div style={{ fontFamily: "'Zilla Slab', serif", fontWeight: 700, fontSize: 26, color: accent || "#20242B", marginTop: 4 }}>{value}</div>
      {sub && <div style={{ fontSize: 11, color: "#8A8371", marginTop: 2 }}>{sub}</div>}
    </div>
  );
}

function EmptyState({ text, actionLabel, onAction }) {
  return (
    <div style={{ textAlign: "center", padding: "34px 10px", color: "#8A8371" }}>
      <div style={{ fontSize: 13.5 }}>{text}</div>
      {actionLabel && (
        <button style={{ ...S.primaryBtn, margin: "14px auto 0" }} onClick={onAction}>
          <Plus size={14} /> {actionLabel}
        </button>
      )}
    </div>
  );
}

/* ============================== EMPLOYEES =================================== */

function EmployeesView({ config, canEdit, isAdmin, addEmployee, updateEmployee, deleteEmployee }) {
  const [q, setQ] = useState("");
  const [modal, setModal] = useState(null); // {mode:'add'|'edit', emp}

  const filtered = config.employees.filter((e) =>
    (e.name + " " + e.code + " " + (e.department || "")).toLowerCase().includes(q.toLowerCase()));

  return (
    <div>
      <div style={S.rowBetween}>
        <div>
          <div style={S.h1}>Employees</div>
          <div style={S.subtle}>{config.employees.length} on record</div>
        </div>
        {canEdit && (
          <button style={S.primaryBtn} onClick={() => setModal({ mode: "add" })}>
            <Plus size={14} /> Add employee
          </button>
        )}
      </div>

      <div style={{ ...S.card, padding: 0 }}>
        <div style={{ padding: 14, borderBottom: "1px solid #EAE5D6" }}>
          <SearchBox value={q} onChange={setQ} placeholder="Search employees…" />
        </div>
        {filtered.length === 0 ? (
          <EmptyState text="No employees match." />
        ) : (
          <table style={S.table}>
            <thead>
              <tr>
                <th style={S.th}>Name</th>
                <th style={S.th}>Code</th>
                <th style={S.th}>Department</th>
                <th style={S.th}>Default shift</th>
                <th style={S.th}>Status</th>
                {canEdit && <th style={S.th}></th>}
              </tr>
            </thead>
            <tbody>
              {filtered.map((e) => (
                <tr key={e.id}>
                  <td style={S.td}><b>{e.name}</b></td>
                  <td style={S.tdMono}>{e.code}</td>
                  <td style={S.td}>{e.department || "—"}</td>
                  <td style={S.tdMono}>{e.defaultShift}</td>
                  <td style={S.td}>{e.active === false
                    ? <span style={{ color: "#A63A2E" }}>Inactive</span>
                    : <span style={{ color: "#2F6F4E" }}>Active</span>}</td>
                  {canEdit && (
                    <td style={{ ...S.td, textAlign: "right", whiteSpace: "nowrap" }}>
                      <button style={S.iconBtn} onClick={() => setModal({ mode: "edit", emp: e })}><Pencil size={14} /></button>
                      {isAdmin && (
                        <button style={S.iconBtn} onClick={() => {
                          if (window.confirm(`Remove ${e.name} from the ledger? Their punch history stays stored.`)) deleteEmployee(e.id);
                        }}><Trash2 size={14} /></button>
                      )}
                    </td>
                  )}
                </tr>
              ))}
            </tbody>
          </table>
        )}
      </div>

      {modal && (
        <EmployeeModal
          shifts={config.shifts}
          initial={modal.emp}
          onClose={() => setModal(null)}
          onSave={(fields) => {
            if (modal.mode === "add") addEmployee(fields);
            else updateEmployee(modal.emp.id, fields);
            setModal(null);
          }}
        />
      )}
    </div>
  );
}

function EmployeeModal({ initial, shifts, onClose, onSave }) {
  const [name, setName] = useState(initial?.name || "");
  const [code, setCode] = useState(initial?.code || "");
  const [department, setDepartment] = useState(initial?.department || "");
  const [defaultShift, setDefaultShift] = useState(initial?.defaultShift || shifts[0]?.code || "");
  const [active, setActive] = useState(initial?.active !== false);

  return (
    <Modal title={initial ? "Edit employee" : "Add employee"} onClose={onClose}>
      <label style={S.label}>Full name</label>
      <input style={S.input} value={name} onChange={(e) => setName(e.target.value)} placeholder="Jordan Lee" />
      <label style={{ ...S.label, marginTop: 12 }}>Employee code</label>
      <input style={S.input} value={code} onChange={(e) => setCode(e.target.value)} placeholder="EMP-014" />
      <label style={{ ...S.label, marginTop: 12 }}>Department</label>
      <input style={S.input} value={department} onChange={(e) => setDepartment(e.target.value)} placeholder="Operations" />
      <label style={{ ...S.label, marginTop: 12 }}>Default shift</label>
      <select style={S.input} value={defaultShift} onChange={(e) => setDefaultShift(e.target.value)}>
        {shifts.map((s) => <option key={s.code} value={s.code}>{s.name} ({s.start}–{s.end})</option>)}
      </select>
      {initial && (
        <label style={{ display: "flex", alignItems: "center", gap: 8, marginTop: 14, fontSize: 13 }}>
          <input type="checkbox" checked={active} onChange={(e) => setActive(e.target.checked)} /> Active employee
        </label>
      )}
      <div style={{ display: "flex", gap: 8, marginTop: 22, justifyContent: "flex-end" }}>
        <button style={S.ghostBtn} onClick={onClose}>Cancel</button>
        <button style={S.primaryBtn} disabled={!name || !code}
          onClick={() => onSave({ name, code, department, defaultShift, active })}>
          <Check size={14} /> Save
        </button>
      </div>
    </Modal>
  );
}

/* ============================== ATTENDANCE =================================== */

function AttendanceView({ config, shiftByCode, attnCache, ensureMonthLoaded, upsertRecord, deleteRecord, bulkUpsert, canEdit }) {
  const { employees, settings } = config;
  const [empId, setEmpId] = useState(employees[0]?.id || "");
  const [ym, setYm] = useState(currentYm());
  const [editDate, setEditDate] = useState(null);
  const [importOpen, setImportOpen] = useState(false);

  useEffect(() => { if (!empId && employees[0]) setEmpId(employees[0].id); }, [employees, empId]);
  useEffect(() => { if (empId) ensureMonthLoaded(empId, ym); }, [empId, ym, ensureMonthLoaded]);

  const emp = employees.find((e) => e.id === empId);
  const monthKey = `attn:${empId}:${ym}`;
  const monthData = attnCache[monthKey] || {};

  const rows = useMemo(() => {
    return daysInMonth(ym).map((d) => {
      const rec = monthData[d] || { date: d, employeeId: empId, dayType: "working", shiftCode: emp?.defaultShift, firstIn: "", lastOut: "", breakMinutes: 0, remark: "" };
      return calcDay(rec, shiftByCode[rec.shiftCode] || DEFAULT_SHIFTS[0], settings);
    });
  }, [monthData, ym, empId, emp, shiftByCode, settings]);

  const sum = useMemo(() => summarize(rows.filter((r) => monthData[r.date])), [rows, monthData]);

  if (employees.length === 0) {
    return <EmptyState text="Add an employee first to start recording attendance." />;
  }

  return (
    <div>
      <div style={S.rowBetween}>
        <div>
          <div style={S.h1}>Attendance</div>
          <div style={S.subtle}>Daily punches, editable at any time — every summary recalculates instantly.</div>
        </div>
        <div style={{ display: "flex", gap: 8 }}>
          {canEdit && (
            <button style={S.ghostBtn} onClick={() => setImportOpen(true)}><Upload size={14} /> Import CSV</button>
          )}
        </div>
      </div>

      <div style={{ display: "flex", gap: 14, marginBottom: 14, flexWrap: "wrap", alignItems: "center" }}>
        <select style={{ ...S.input, width: 240 }} value={empId} onChange={(e) => setEmpId(e.target.value)}>
          {employees.map((e) => <option key={e.id} value={e.id}>{e.name} — {e.code}</option>)}
        </select>
        <MonthSwitcher ym={ym} setYm={setYm} />
      </div>

      <div style={S.cardGrid}>
        <StatCard label="Present" value={sum.presentDays} accent="#2F6F4E" />
        <StatCard label="Half day" value={sum.halfDays} accent="#C98B12" />
        <StatCard label="Absent" value={sum.totalAbsent} accent="#A63A2E" />
        <StatCard label="Attendance %" value={`${sum.attendancePct}%`} />
        <StatCard label="Net hours" value={hoursToHM(sum.netHours)} />
        <StatCard label="Week-off OT" value={hoursToHM(sum.totalWeekOffOT)} />
      </div>

      <div style={{ ...S.card, padding: 0, overflowX: "auto" }}>
        <table style={S.table}>
          <thead>
            <tr>
              <th style={S.th}>Date</th>
              <th style={S.th}>Day</th>
              <th style={S.th}>Type</th>
              <th style={S.th}>First in</th>
              <th style={S.th}>Last out</th>
              <th style={S.th}>Work hrs</th>
              <th style={S.th}>Late-in</th>
              <th style={S.th}>Early-out</th>
              <th style={S.th}>OT</th>
              <th style={S.th}>Status</th>
              {canEdit && <th style={S.th}></th>}
            </tr>
          </thead>
          <tbody>
            {rows.map((r) => {
              const entered = !!monthData[r.date];
              return (
                <tr key={r.date} style={entered ? undefined : { opacity: 0.45 }}>
                  <td style={S.tdMono}>{r.date.slice(8)}</td>
                  <td style={S.td}>{weekdayShort(r.date)}</td>
                  <td style={S.td}>{r.dayType === "weekoff" ? "Week off" : r.dayType === "leave" ? "Leave" : "Working"}</td>
                  <td style={S.tdMono}>{r.firstIn || "—"}</td>
                  <td style={S.tdMono}>{r.lastOut || "—"}</td>
                  <td style={S.tdMono}>{entered ? hoursToHM(r.workHours) : "—"}</td>
                  <td style={S.tdMono}>{entered && r.lateInH > 0 ? hoursToHM(r.lateInH) : "—"}</td>
                  <td style={S.tdMono}>{entered && r.earlyOutH > 0 ? hoursToHM(r.earlyOutH) : "—"}</td>
                  <td style={S.tdMono}>{entered && r.extraWork > 0 ? hoursToHM(r.extraWork) : "—"}</td>
                  <td style={S.td}>{entered ? <Stamp status={r.dayStatus} small /> : <span style={{ fontSize: 11, color: "#B8B199" }}>not entered</span>}</td>
                  {canEdit && (
                    <td style={{ ...S.td, textAlign: "right", whiteSpace: "nowrap" }}>
                      <button style={S.iconBtn} onClick={() => setEditDate(r.date)}><Pencil size={13} /></button>
                    </td>
                  )}
                </tr>
              );
            })}
          </tbody>
        </table>
      </div>

      {editDate && emp && (
        <DayEditModal
          date={editDate}
          record={monthData[editDate]}
          shifts={config.shifts}
          defaultShift={emp.defaultShift}
          onClose={() => setEditDate(null)}
          onSave={(fields) => { upsertRecord(empId, editDate, fields); setEditDate(null); }}
          onDelete={monthData[editDate] ? () => { deleteRecord(empId, editDate); setEditDate(null); } : null}
        />
      )}

      {importOpen && emp && (
        <ImportModal
          employee={emp}
          shifts={config.shifts}
          onClose={() => setImportOpen(false)}
          onImport={(records) => { bulkUpsert(empId, records); setImportOpen(false); }}
        />
      )}
    </div>
  );
}

function DayEditModal({ date, record, shifts, defaultShift, onClose, onSave, onDelete }) {
  const [dayType, setDayType] = useState(record?.dayType || "working");
  const [shiftCode, setShiftCode] = useState(record?.shiftCode || defaultShift);
  const [firstIn, setFirstIn] = useState(record?.firstIn || "");
  const [lastOut, setLastOut] = useState(record?.lastOut || "");
  const [breakMinutes, setBreakMinutes] = useState(record?.breakMinutes || 0);
  const [remark, setRemark] = useState(record?.remark || "");

  return (
    <Modal title={`Edit — ${date}`} onClose={onClose}>
      <label style={S.label}>Day type</label>
      <select style={S.input} value={dayType} onChange={(e) => setDayType(e.target.value)}>
        <option value="working">Working day</option>
        <option value="weekoff">Week off</option>
        <option value="leave">On leave</option>
      </select>

      <label style={{ ...S.label, marginTop: 12 }}>Shift</label>
      <select style={S.input} value={shiftCode} onChange={(e) => setShiftCode(e.target.value)}>
        {shifts.map((s) => <option key={s.code} value={s.code}>{s.name} ({s.start}–{s.end})</option>)}
      </select>

      <div style={{ display: "flex", gap: 12, marginTop: 12 }}>
        <div style={{ flex: 1 }}>
          <label style={S.label}>First in</label>
          <input style={S.input} type="time" value={firstIn} onChange={(e) => setFirstIn(e.target.value)} />
        </div>
        <div style={{ flex: 1 }}>
          <label style={S.label}>Last out</label>
          <input style={S.input} type="time" value={lastOut} onChange={(e) => setLastOut(e.target.value)} />
        </div>
      </div>

      <label style={{ ...S.label, marginTop: 12 }}>Break (minutes)</label>
      <input style={S.input} type="number" min="0" value={breakMinutes} onChange={(e) => setBreakMinutes(Number(e.target.value) || 0)} />

      <label style={{ ...S.label, marginTop: 12 }}>Remark</label>
      <input style={S.input} value={remark} onChange={(e) => setRemark(e.target.value)} placeholder="Optional note" />

      <div style={{ display: "flex", gap: 8, marginTop: 22, justifyContent: "space-between" }}>
        {onDelete ? (
          <button style={{ ...S.ghostBtn, color: "#A63A2E", borderColor: "#E3C5BE" }} onClick={onDelete}>
            <Trash2 size={14} /> Delete record
          </button>
        ) : <span />}
        <div style={{ display: "flex", gap: 8 }}>
          <button style={S.ghostBtn} onClick={onClose}>Cancel</button>
          <button style={S.primaryBtn} onClick={() => onSave({ dayType, shiftCode, firstIn, lastOut, breakMinutes, remark })}>
            <Check size={14} /> Save & recalculate
          </button>
        </div>
      </div>
    </Modal>
  );
}


function normalizeDate(dateStr) {
  if (!dateStr) return "";

  // XLSX date cells can arrive as JS Date objects (when read with cellDates:true).
  // Format those as DD-MM-YYYY text first, then fall through to the string parsing below.
  if (dateStr instanceof Date && !isNaN(dateStr)) {
    dateStr = `${pad2(dateStr.getDate())}-${pad2(dateStr.getMonth() + 1)}-${dateStr.getFullYear()}`;
  }

  dateStr = String(dateStr).trim();

  // Primary expected format: DD-MM-YYYY
  if (/^\d{2}-\d{2}-\d{4}$/.test(dateStr)) {
    const [dd, mm, yyyy] = dateStr.split("-");
    return `${yyyy}-${mm}-${dd}`;
  }
  // Also accepted: DD/MM/YYYY
  if (/^\d{2}\/\d{2}\/\d{4}$/.test(dateStr)) {
    const [dd, mm, yyyy] = dateStr.split("/");
    return `${yyyy}-${mm}-${dd}`;
  }
  // Still accepted for backwards compatibility with older exports: YYYY-MM-DD
  if (/^\d{4}-\d{2}-\d{2}$/.test(dateStr)) return dateStr;

  return "";
}

function ImportModal({ employee, shifts, onClose, onImport }) {
  const [fileName, setFileName] = useState("");
  const [rows, setRows] = useState([]);
  const [err, setErr] = useState("");
  const fileRef = useRef(null);

  // Shared row -> record normalizer, used for both CSV rows (from papaparse)
  // and XLSX/XLS rows (from SheetJS) so downstream logic stays identical.
  const normalizeRows = (rawRows) => {
    try {
      const parsed = rawRows.map((row) => {
        const date = normalizeDate(row.Date ?? row.date ?? "");
        const isWO = /wo|week ?off/i.test(row.Shift || row.dayType || "");
        return {
          date,
          dayType: isWO ? "weekoff" : "working",
          shiftCode: row.Shift && shifts.find((s) => s.code === row.Shift) ? row.Shift : (shifts[0]?.code || ""),
          firstIn: String(row["First IN"] ?? row.firstIn ?? "").trim(),
          lastOut: String(row["Last OUT"] ?? row.lastOut ?? "").trim(),
          breakMinutes: Number(row.Break || row.breakMinutes || 0) || 0,
          remark: row.Remark || row.remark || "",
        };
      }).filter((r) => r.date !== "");
      setRows(parsed);
      if (parsed.length === 0) setErr("No valid rows found. Supported date formats: DD-MM-YYYY, DD/MM/YYYY, YYYY-MM-DD.");
    } catch (e) { setErr("Could not read that file."); }
  };

  const handleFile = (file) => {
    setFileName(file.name);
    setErr("");
    const ext = file.name.split(".").pop().toLowerCase();

    if (ext === "csv") {
      Papa.parse(file, {
        header: true,
        skipEmptyLines: true,
        complete: (res) => normalizeRows(res.data),
        error: () => setErr("Could not read that file."),
      });
    } else if (ext === "xlsx" || ext === "xls") {
      const reader = new FileReader();
      reader.onload = (e) => {
        try {
          const data = new Uint8Array(e.target.result);
          // cellDates:true turns date-formatted cells into JS Date objects instead
          // of Excel's serial-number form, so normalizeDate can format them itself.
          const wb = XLSX.read(data, { type: "array", cellDates: true });
          const sheet = wb.Sheets[wb.SheetNames[0]];
          // dateNF forces any date cell that SheetJS stringifies to DD-MM-YYYY text;
          // raw:false also renders numbers/times as their displayed text, defval fills blanks.
          const rawRows = XLSX.utils.sheet_to_json(sheet, { defval: "", raw: false, dateNF: "dd-mm-yyyy" });
          normalizeRows(rawRows);
        } catch (e) {
          setErr("Could not read that file.");
        }
      };
      reader.onerror = () => setErr("Could not read that file.");
      reader.readAsArrayBuffer(file);
    } else {
      setErr("Unsupported file type. Please upload a .csv, .xlsx or .xls file.");
    }
  };

  return (
    <Modal title={`Import punches — ${employee.name}`} onClose={onClose} wide>
      <div style={{ fontSize: 12.5, color: "#6B6656", marginBottom: 10, lineHeight: 1.6 }}>
        CSV or Excel columns expected: <code>Date</code> (DD-MM-YYYY), <code>Shift</code> (shift code, or contains "WO" for week off),
        <code> First IN</code>, <code>Last OUT</code>, optional <code>Break</code> (minutes) and <code>Remark</code>.
      </div>
      <div style={S.dropZone} onClick={() => fileRef.current?.click()}>
        <Upload size={18} />
        <div style={{ marginTop: 6, fontSize: 13 }}>{fileName || "Click to choose a .csv or .xlsx file"}</div>
      </div>
      <input ref={fileRef} type="file" accept=".csv,.xlsx,.xls" style={{ display: "none" }}
        onChange={(e) => e.target.files[0] && handleFile(e.target.files[0])} />
      {err && <div style={{ color: "#A63A2E", fontSize: 12.5, marginTop: 8 }}>{err}</div>}
      {rows.length > 0 && (
        <div style={{ marginTop: 12, maxHeight: 220, overflow: "auto", border: "1px solid #EAE5D6", borderRadius: 4 }}>
          <table style={S.table}>
            <thead><tr><th style={S.th}>Date</th><th style={S.th}>Type</th><th style={S.th}>In</th><th style={S.th}>Out</th></tr></thead>
            <tbody>
              {rows.slice(0, 50).map((r) => (
                <tr key={r.date}><td style={S.tdMono}>{r.date}</td><td style={S.td}>{r.dayType}</td><td style={S.tdMono}>{r.firstIn}</td><td style={S.tdMono}>{r.lastOut}</td></tr>
              ))}
            </tbody>
          </table>
          <div style={{ fontSize: 11.5, color: "#8A8371", padding: 8 }}>{rows.length} rows ready to import.</div>
        </div>
      )}
      <div style={{ display: "flex", gap: 8, marginTop: 20, justifyContent: "flex-end" }}>
        <button style={S.ghostBtn} onClick={onClose}>Cancel</button>
        <button style={S.primaryBtn} disabled={rows.length === 0}
          onClick={() => { const byDate = {}; rows.forEach((r) => { byDate[r.date] = r; }); onImport(byDate); }}>
          <Check size={14} /> Import {rows.length ? `(${rows.length})` : ""}
        </button>
      </div>
    </Modal>
  );
}

/* ============================== REPORTS =================================== */

function ReportsView({ config, shiftByCode, ensureMonthLoaded, attnCache }) {
  const { employees, settings } = config;
  const [empId, setEmpId] = useState("all");
  const [start, setStart] = useState(`${currentYm()}-01`);
  const [end, setEnd] = useState(todayIso());
  const [busy, setBusy] = useState(false);
  const printRef = useRef(null);

  const targetEmployees = empId === "all" ? employees : employees.filter((e) => e.id === empId);

  const [loadedRows, setLoadedRows] = useState([]);

  const runReport = useCallback(async () => {
    setBusy(true);
    const months = monthsBetween(start, end);
    const allRows = [];
    for (const emp of targetEmployees) {
      for (const ym of months) {
        const data = await ensureMonthLoaded(emp.id, ym);
        Object.values(data).forEach((rec) => {
          if (rec.date >= start && rec.date <= end) {
            const calc = calcDay(rec, shiftByCode[rec.shiftCode] || DEFAULT_SHIFTS[0], settings);
            allRows.push({ ...calc, employeeName: emp.name, employeeCode: emp.code });
          }
        });
      }
    }
    allRows.sort((a, b) => (a.employeeName + a.date).localeCompare(b.employeeName + b.date));
    setLoadedRows(allRows);
    setBusy(false);
  }, [targetEmployees, start, end, ensureMonthLoaded, shiftByCode, settings]);

  useEffect(() => { if (employees.length) runReport(); /* eslint-disable-next-line */ }, [employees.length]);

  const perEmpSummary = useMemo(() => {
    const byEmp = {};
    loadedRows.forEach((r) => {
      byEmp[r.employeeName] = byEmp[r.employeeName] || [];
      byEmp[r.employeeName].push(r);
    });
    return Object.entries(byEmp).map(([name, rows]) => ({ name, sum: summarize(rows), rows }));
  }, [loadedRows]);

  const overall = useMemo(() => summarize(loadedRows), [loadedRows]);

  const exportExcel = () => {
    const wb = XLSX.utils.book_new();
    const sheetRows = loadedRows.map((r) => ({
      Employee: r.employeeName, Code: r.employeeCode, Date: r.date, Day: weekdayShort(r.date),
      Type: r.dayType, "First In": r.firstIn, "Last Out": r.lastOut, "Work Hours": hoursToHM(r.workHours),
      "Late In": hoursToHM(r.lateInH), "Early Out": hoursToHM(r.earlyOutH), "Extra Work": hoursToHM(r.extraWork),
      "Special Addition": hoursToHM(r.specialAddition), "Total Addition": hoursToHM(r.totalAddition),
      "Total Deduction": hoursToHM(r.totalDeduction), "Net": hoursToHM(r.net),
      "Week-Off OT (1x)": r.weekOffWorkedHours > 0 ? hoursToHM(r.weekOffWorkedHours) : "",
      Status: r.dayStatus, Remark: r.remark || "",
    }));
    const ws = XLSX.utils.json_to_sheet(sheetRows);
    XLSX.utils.book_append_sheet(wb, ws, "Daily Detail");

    const summaryRows = perEmpSummary.map(({ name, sum }) => ({
      Employee: name,
      "Total Days": sum.totalDays,
      "Present Days (Full)": sum.presentDays,
      "Half Days": sum.halfDays,
      "Absent - Worked <Threshold": sum.absentShort,
      "Absent - No Punches": sum.absentNoPunch,
      "Total Absent Days": sum.totalAbsent,
      "Week Off Days": sum.weekOffDays,
      "Total Working Days (Total-WeekOff)": sum.totalWorkingDaysStrict,
      "Reconciliation Check": sum.reconciliationOk ? "OK" : `Mismatch (${sum.reconciliationDiff > 0 ? "+" : ""}${sum.reconciliationDiff})`,
      "Final Attendance (Half=0.5)": sum.finalAttendance,
      "Attendance %": sum.attendancePct,
      "Total Late-IN Hours": hoursToHM(sum.totalLateHours),
      "Total Early-OUT Hours": hoursToHM(sum.totalEarlyHours),
      "Total Extra Work Hours": hoursToHM(sum.totalExtraWork),
      "Total Special Addition Hours": hoursToHM(sum.totalSpecialAddition),
      "TOTAL ADDITION HOURS": hoursToHM(sum.totalAddition),
      "TOTAL DEDUCTION HOURS": hoursToHM(sum.totalDeduction),
      "Net (Addition-Deduction)": hoursToHM(sum.netHours),
      "Week Off Days Worked (Count)": sum.weekOffWorkedCount,
      "Total OT (1x) Hours - Week Off": hoursToHM(sum.totalWeekOffOT),
    }));
    const ws2 = XLSX.utils.json_to_sheet(summaryRows);
    XLSX.utils.book_append_sheet(wb, ws2, "Summary");

    const additionRows = perEmpSummary.flatMap(({ name, sum }) =>
      sum.additionDates.map((d) => ({ Employee: name, Date: d.date, "Hours Added": hoursToHM(d.hours) })));
    const ws3 = XLSX.utils.json_to_sheet(additionRows.length ? additionRows : [{ Employee: "", Date: "None this period", "Hours Added": "" }]);
    XLSX.utils.book_append_sheet(wb, ws3, "Addition Dates");

    const weekOffOtRows = perEmpSummary.flatMap(({ name, sum }) =>
      sum.weekOffOtDates.map((d) => ({ Employee: name, Date: d.date, "OT (1x) Hours": hoursToHM(d.hours) })));
    const ws4 = XLSX.utils.json_to_sheet(weekOffOtRows.length ? weekOffOtRows : [{ Employee: "", Date: "None this period", "OT (1x) Hours": "" }]);
    XLSX.utils.book_append_sheet(wb, ws4, "Week-Off OT Dates");

    XLSX.writeFile(wb, `attendance-report_${start}_to_${end}.xlsx`);
  };

  const exportPdf = () => { window.print(); };

  return (
    <div>
      <div style={S.rowBetween}>
        <div>
          <div style={S.h1}>Reports</div>
          <div style={S.subtle}>Search, filter and export any employee or date range — nothing is ever overwritten.</div>
        </div>
        <div style={{ display: "flex", gap: 8 }} className="no-print">
          <button style={S.ghostBtn} onClick={exportExcel} disabled={loadedRows.length === 0}><Download size={14} /> Export Excel</button>
          <button style={S.ghostBtn} onClick={exportPdf} disabled={loadedRows.length === 0}><Printer size={14} /> Export PDF</button>
        </div>
      </div>

      <div style={{ display: "flex", gap: 12, marginBottom: 14, flexWrap: "wrap", alignItems: "flex-end" }} className="no-print">
        <div>
          <label style={S.label}>Employee</label>
          <select style={{ ...S.input, width: 220 }} value={empId} onChange={(e) => setEmpId(e.target.value)}>
            <option value="all">All employees</option>
            {employees.map((e) => <option key={e.id} value={e.id}>{e.name}</option>)}
          </select>
        </div>
        <div>
          <label style={S.label}>From</label>
          <input style={S.input} type="date" value={start} onChange={(e) => setStart(e.target.value)} />
        </div>
        <div>
          <label style={S.label}>To</label>
          <input style={S.input} type="date" value={end} onChange={(e) => setEnd(e.target.value)} />
        </div>
        <button style={S.primaryBtn} onClick={runReport} disabled={busy}>{busy ? "Loading…" : "Run report"}</button>
      </div>

      <div id="print-area" ref={printRef}>
        <div style={{ fontFamily: "'IBM Plex Mono', monospace", fontSize: 11, color: "#8A8371", marginBottom: 4 }}>
          {start} — {end} · {targetEmployees.length} employee{targetEmployees.length !== 1 ? "s" : ""} · {loadedRows.length} recorded days
        </div>
        <div style={{ fontFamily: "'IBM Plex Mono', monospace", fontSize: 11, color: "#8A8371", marginBottom: 10 }}>
          Grace: {settings.graceMinutes}m · Default full day: {hoursToHM(settings.fullDayHours)} · Default half day: {hoursToHM(settings.halfDayHours)} (per-shift overrides apply where set)
        </div>

        <div style={S.cardGrid}>
          <StatCard label="Attendance %" value={`${overall.attendancePct}%`} accent="#2F6F4E" />
          <StatCard label="Present" value={overall.presentDays} />
          <StatCard label="Absent" value={overall.totalAbsent} accent="#A63A2E" />
          <StatCard label="Late arrivals" value={overall.lateCount} sub={hoursToHM(overall.totalLateHours)} />
          <StatCard label="Early departures" value={overall.earlyCount} sub={hoursToHM(overall.totalEarlyHours)} />
          <StatCard label="Overtime" value={hoursToHM(overall.totalExtraWork + overall.totalWeekOffOT)} />
          <StatCard label="Addition hours" value={hoursToHM(overall.totalAddition)} accent="#2F6F4E" sub={`incl. ${hoursToHM(overall.totalSpecialAddition)} special`} />
          <StatCard label="Deduction hours" value={hoursToHM(overall.totalDeduction)} accent="#A63A2E" />
          <StatCard
            label="Reconciliation"
            value={overall.reconciliationOk ? "OK" : "Mismatch"}
            accent={overall.reconciliationOk ? "#2F6F4E" : "#A63A2E"}
            sub={overall.reconciliationOk ? "Present+Half+Absent = Working days" : `Off by ${overall.reconciliationDiff > 0 ? "+" : ""}${overall.reconciliationDiff} day(s)`}
          />
        </div>

        <div style={S.card}>
          <div style={S.cardTitle}>Summary by employee</div>
          <div style={S.tableScroll}>
            <table style={S.table}>
              <thead>
                <tr>
                  <th style={S.th}>Employee</th><th style={S.th}>Total days</th><th style={S.th}>Present</th><th style={S.th}>Half</th>
                  <th style={S.th}>Absent</th><th style={S.th}>Week off</th><th style={S.th}>Working days</th>
                  <th style={S.th}>Final attn.</th><th style={S.th}>Attendance %</th>
                  <th style={S.th}>Late</th><th style={S.th}>Early</th>
                  <th style={S.th}>Special addn.</th><th style={S.th}>Addn. hrs</th><th style={S.th}>Dedn. hrs</th>
                  <th style={S.th}>Net hrs</th><th style={S.th}>WO worked</th><th style={S.th}>WO OT hrs</th>
                  <th style={S.th}>Reconciliation</th>
                </tr>
              </thead>
              <tbody>
                {perEmpSummary.map(({ name, sum }) => (
                  <tr key={name}>
                    <td style={S.td}><b>{name}</b></td>
                    <td style={S.tdMono}>{sum.totalDays}</td>
                    <td style={S.tdMono}>{sum.presentDays}</td>
                    <td style={S.tdMono}>{sum.halfDays}</td>
                    <td style={S.tdMono}>{sum.totalAbsent}</td>
                    <td style={S.tdMono}>{sum.weekOffDays}</td>
                    <td style={S.tdMono}>{sum.totalWorkingDaysStrict}</td>
                    <td style={S.tdMono}>{sum.finalAttendance}</td>
                    <td style={S.tdMono}><b>{sum.attendancePct}%</b></td>
                    <td style={S.tdMono}>{sum.lateCount}</td>
                    <td style={S.tdMono}>{sum.earlyCount}</td>
                    <td style={S.tdMono}>{hoursToHM(sum.totalSpecialAddition)}</td>
                    <td style={S.tdMono}>{hoursToHM(sum.totalAddition)}</td>
                    <td style={S.tdMono}>{hoursToHM(sum.totalDeduction)}</td>
                    <td style={S.tdMono}>{hoursToHM(sum.netHours)}</td>
                    <td style={S.tdMono}>{sum.weekOffWorkedCount}</td>
                    <td style={S.tdMono}>{hoursToHM(sum.totalWeekOffOT)}</td>
                    <td style={S.td}>
                      {sum.reconciliationOk
                        ? <span style={{ color: "#2F6F4E", fontSize: 11, fontWeight: 600 }}>OK</span>
                        : <span style={{ color: "#A63A2E", fontSize: 11, fontWeight: 600 }}>{sum.reconciliationDiff > 0 ? "+" : ""}{sum.reconciliationDiff}</span>}
                    </td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        </div>

        <div style={S.card}>
          <div style={S.cardTitle}>Additions — days worked below the half-day threshold</div>
          <div style={{ fontSize: 11.5, color: "#8A8371", marginBottom: 10 }}>
            These days still count as Absent for attendance purposes, but the hours actually worked on them are credited into Total Addition Hours above.
          </div>
          {perEmpSummary.every(({ sum }) => sum.additionDates.length === 0) ? (
            <EmptyState text="None in this period." />
          ) : (
            perEmpSummary.map(({ name, sum }) => sum.additionDates.length > 0 && (
              <div key={name} style={{ marginBottom: 10 }}>
                <div style={{ fontSize: 12.5, fontWeight: 600, marginBottom: 5 }}>{name}</div>
                <div style={{ display: "flex", flexWrap: "wrap", gap: 6 }}>
                  {sum.additionDates.map((d) => (
                    <span key={d.date} style={S.chip}>{d.date} · {hoursToHM(d.hours)}</span>
                  ))}
                </div>
              </div>
            ))
          )}
        </div>

        <div style={S.card}>
          <div style={S.cardTitle}>Week-Off days worked — OT (1×)</div>
          <div style={{ fontSize: 11.5, color: "#8A8371", marginBottom: 10 }}>
            Days still counted as Week Off for attendance purposes, with the hours actually worked pulled out separately as OT (1×).
          </div>
          {perEmpSummary.every(({ sum }) => sum.weekOffOtDates.length === 0) ? (
            <EmptyState text="None in this period." />
          ) : (
            perEmpSummary.map(({ name, sum }) => sum.weekOffOtDates.length > 0 && (
              <div key={name} style={{ marginBottom: 10 }}>
                <div style={{ fontSize: 12.5, fontWeight: 600, marginBottom: 5 }}>{name}</div>
                <div style={{ display: "flex", flexWrap: "wrap", gap: 6 }}>
                  {sum.weekOffOtDates.map((d) => (
                    <span key={d.date} style={S.chip}>{d.date} · {hoursToHM(d.hours)} OT</span>
                  ))}
                </div>
              </div>
            ))
          )}
        </div>

        <div style={S.card}>
          <div style={S.cardTitle}>Daily detail</div>
          {loadedRows.length === 0 ? <EmptyState text="No records in this range yet." /> : (
            <div style={S.tableScroll}>
            <table style={S.table}>
              <thead>
                <tr>
                  <th style={S.th}>Employee</th><th style={S.th}>Date</th><th style={S.th}>In</th><th style={S.th}>Out</th>
                  <th style={S.th}>Work hrs</th><th style={S.th}>Late</th><th style={S.th}>Early</th>
                  <th style={S.th}>Addn.</th><th style={S.th}>Dedn.</th><th style={S.th}>Net</th><th style={S.th}>Status</th>
                </tr>
              </thead>
              <tbody>
                {loadedRows.map((r) => (
                  <tr key={r.employeeCode + r.date}>
                    <td style={S.td}>{r.employeeName}</td>
                    <td style={S.tdMono}>{r.date}</td>
                    <td style={S.tdMono}>{r.firstIn || "—"}</td>
                    <td style={S.tdMono}>{r.lastOut || "—"}</td>
                    <td style={S.tdMono}>{hoursToHM(r.workHours)}</td>
                    <td style={S.tdMono}>{r.lateInH > 0 ? hoursToHM(r.lateInH) : "—"}</td>
                    <td style={S.tdMono}>{r.earlyOutH > 0 ? hoursToHM(r.earlyOutH) : "—"}</td>
                    <td style={S.tdMono}>{r.totalAddition > 0 ? hoursToHM(r.totalAddition) : "—"}</td>
                    <td style={S.tdMono}>{r.totalDeduction > 0 ? hoursToHM(r.totalDeduction) : "—"}</td>
                    <td style={S.tdMono}>{hoursToHM(r.net)}</td>
                    <td style={S.td}><Stamp status={r.dayStatus} small /></td>
                  </tr>
                ))}
              </tbody>
            </table>
            </div>
          )}
        </div>
      </div>
    </div>
  );
}

/* ============================== SETTINGS =================================== */

function SettingsView({ config, isAdmin, addShift, updateShift, deleteShift, updateSettings, bulkAddShifts }) {
  const { settings, shifts } = config;
  const [grace, setGrace] = useState(settings.graceMinutes);
  const [fullDay, setFullDay] = useState(settings.fullDayHours);
  const [halfDay, setHalfDay] = useState(settings.halfDayHours);
  const [shiftModal, setShiftModal] = useState(null);
  const sangamRemaining = SANGAM_SHIFTS.filter((s) => !shifts.some((x) => x.code === s.code)).length;

  return (
    <div>
      <div style={S.h1}>Settings</div>
      <div style={S.subtle}>Company-wide rules — every dashboard and report recalculates the moment these change.</div>

      <div style={S.card}>
        <div style={S.cardTitle}>Attendance rules</div>
        {!isAdmin && <div style={{ fontSize: 12, color: "#8A8371", marginBottom: 10 }}>Only Admins can change these.</div>}
        <div style={{ display: "flex", gap: 20, flexWrap: "wrap" }}>
          <div>
            <label style={S.label}>Grace time (minutes)</label>
            <input style={S.input} type="number" disabled={!isAdmin} value={grace} onChange={(e) => setGrace(Number(e.target.value))} />
          </div>
          <div>
            <label style={S.label}>Full-day hours</label>
            <input style={S.input} type="number" step="0.5" disabled={!isAdmin} value={fullDay} onChange={(e) => setFullDay(Number(e.target.value))} />
          </div>
          <div>
            <label style={S.label}>Half-day hours</label>
            <input style={S.input} type="number" step="0.5" disabled={!isAdmin} value={halfDay} onChange={(e) => setHalfDay(Number(e.target.value))} />
          </div>
        </div>
        {isAdmin && (
          <button style={{ ...S.primaryBtn, marginTop: 16 }}
            onClick={() => updateSettings({ graceMinutes: grace, fullDayHours: fullDay, halfDayHours: halfDay })}>
            <Check size={14} /> Save rules
          </button>
        )}
      </div>

      <div style={S.card}>
        <div style={S.rowBetween}>
          <div style={S.cardTitle}>Shifts</div>
          {isAdmin && (
            <div style={{ display: "flex", gap: 8 }}>
              {sangamRemaining > 0 && (
                <button style={S.ghostBtn} onClick={() => bulkAddShifts(SANGAM_SHIFTS)}>
                  <Upload size={14} /> Add Sangam Electronics shift list ({sangamRemaining})
                </button>
              )}
              <button style={S.primaryBtn} onClick={() => setShiftModal({ mode: "add" })}><Plus size={14} /> Add shift</button>
            </div>
          )}
        </div>
        <table style={S.table}>
          <thead>
            <tr>
              <th style={S.th}>Code</th><th style={S.th}>Name</th><th style={S.th}>Start</th><th style={S.th}>End</th>
              <th style={S.th}>Break</th><th style={S.th}>Full day</th><th style={S.th}>Half day</th>
              {isAdmin && <th style={S.th}></th>}
            </tr>
          </thead>
          <tbody>
            {shifts.map((s) => (
              <tr key={s.code}>
                <td style={S.tdMono}>{s.code}</td>
                <td style={S.td}>{s.name}</td>
                <td style={S.tdMono}>{s.start}</td>
                <td style={S.tdMono}>{s.end}</td>
                <td style={S.tdMono}>{s.breakStart && s.breakEnd ? `${s.breakStart}–${s.breakEnd}` : "—"}</td>
                <td style={S.tdMono}>{hoursToHM(s.fullDayHours != null ? s.fullDayHours : settings.fullDayHours)}</td>
                <td style={S.tdMono}>{hoursToHM(s.halfDayHours != null ? s.halfDayHours : settings.halfDayHours)}</td>
                {isAdmin && (
                  <td style={{ ...S.td, textAlign: "right" }}>
                    <button style={S.iconBtn} onClick={() => setShiftModal({ mode: "edit", shift: s })}><Pencil size={13} /></button>
                    <button style={S.iconBtn} onClick={() => { if (window.confirm(`Delete shift ${s.code}?`)) deleteShift(s.code); }}><Trash2 size={13} /></button>
                  </td>
                )}
              </tr>
            ))}
          </tbody>
        </table>
      </div>

      <div style={{ ...S.card, background: "#F1EEE3" }}>
        <div style={S.cardTitle}>About access in this ledger</div>
        <div style={{ fontSize: 12.5, color: "#6B6656", lineHeight: 1.7 }}>
          Roles here (Admin / Manager / Viewer) control what the interface lets someone click — they are a
          convenience for teams sharing this ledger, not a secured login system. All data is stored centrally and
          shared with everyone who opens this app. For a version with real password-protected accounts and
          per-user permissions enforced on the server, this is best rebuilt as a hosted application.
        </div>
      </div>

      {shiftModal && (
        <ShiftModal
          initial={shiftModal.shift}
          onClose={() => setShiftModal(null)}
          onSave={(fields) => {
            if (shiftModal.mode === "add") addShift(fields);
            else updateShift(shiftModal.shift.code, fields);
            setShiftModal(null);
          }}
        />
      )}
    </div>
  );
}

function ShiftModal({ initial, onClose, onSave }) {
  const [code, setCode] = useState(initial?.code || "");
  const [name, setName] = useState(initial?.name || "");
  const [start, setStart] = useState(initial?.start || "09:00");
  const [end, setEnd] = useState(initial?.end || "18:00");
  return (
    <Modal title={initial ? "Edit shift" : "Add shift"} onClose={onClose}>
      <label style={S.label}>Shift code</label>
      <input style={S.input} value={code} disabled={!!initial} onChange={(e) => setCode(e.target.value.toUpperCase())} placeholder="GEN" />
      <label style={{ ...S.label, marginTop: 12 }}>Name</label>
      <input style={S.input} value={name} onChange={(e) => setName(e.target.value)} placeholder="General Shift" />
      <div style={{ display: "flex", gap: 12, marginTop: 12 }}>
        <div style={{ flex: 1 }}>
          <label style={S.label}>Start</label>
          <input style={S.input} type="time" value={start} onChange={(e) => setStart(e.target.value)} />
        </div>
        <div style={{ flex: 1 }}>
          <label style={S.label}>End</label>
          <input style={S.input} type="time" value={end} onChange={(e) => setEnd(e.target.value)} />
        </div>
      </div>
      <div style={{ display: "flex", gap: 8, marginTop: 22, justifyContent: "flex-end" }}>
        <button style={S.ghostBtn} onClick={onClose}>Cancel</button>
        <button style={S.primaryBtn} disabled={!code || !name} onClick={() => onSave({ code, name, start, end })}>
          <Check size={14} /> Save
        </button>
      </div>
    </Modal>
  );
}

/* ============================== SHARED WIDGETS =================================== */

function SearchBox({ value, onChange, placeholder }) {
  return (
    <div style={{ display: "flex", alignItems: "center", gap: 8, border: "1px solid #DED9CA", borderRadius: 4, padding: "7px 10px", background: "#FBFAF6" }}>
      <Search size={14} color="#8A8371" />
      <input value={value} onChange={(e) => onChange(e.target.value)} placeholder={placeholder}
        style={{ border: "none", outline: "none", background: "transparent", fontSize: 13, fontFamily: "Inter", width: "100%" }} />
    </div>
  );
}

function Modal({ title, children, onClose, wide }) {
  return (
    <div style={S.modalOverlay} onClick={onClose}>
      <div style={{ ...S.modalBox, width: wide ? 520 : 400 }} onClick={(e) => e.stopPropagation()}>
        <div style={S.rowBetween}>
          <div style={{ fontFamily: "'Zilla Slab', serif", fontWeight: 600, fontSize: 18, color: "#20242B" }}>{title}</div>
          <button style={S.iconBtn} onClick={onClose}><X size={16} /></button>
        </div>
        <div style={{ marginTop: 14 }}>{children}</div>
      </div>
    </div>
  );
}

/* ============================== STYLES =================================== */

const S = {
  appShell: {
    display: "flex", minHeight: "100vh", background: "#F5F2EA",
    fontFamily: "'Inter', sans-serif", color: "#20242B",
  },
  sidebar: {
    width: 210, background: "#20242B", display: "flex", flexDirection: "column",
    flexShrink: 0,
  },
  main: { flex: 1, display: "flex", flexDirection: "column", minWidth: 0 },
  topbar: {
    height: 46, display: "flex", alignItems: "center", justifyContent: "space-between",
    padding: "0 26px", borderBottom: "1px solid #E4DFD0", background: "#F9F7F1",
  },
  content: { padding: "22px 26px 60px", overflowY: "auto" },
  navItem: {
    display: "flex", alignItems: "center", gap: 10, padding: "9px 12px", borderRadius: 4,
    color: "#B7C0CF", fontSize: 13.5, cursor: "pointer", marginBottom: 2,
  },
  navItemActive: { background: "#2B3A55", color: "#F5F2EA" },
  h1: { fontFamily: "'Zilla Slab', serif", fontWeight: 700, fontSize: 24, color: "#20242B" },
  subtle: { fontSize: 12.5, color: "#8A8371", marginTop: 2 },
  rowBetween: { display: "flex", alignItems: "center", justifyContent: "space-between", marginBottom: 18, flexWrap: "wrap", gap: 10 },
  cardGrid: { display: "grid", gridTemplateColumns: "repeat(auto-fit, minmax(150px, 1fr))", gap: 12, marginBottom: 18 },
  statCard: { background: "#fff", border: "1px solid #EAE5D6", borderRadius: 5, padding: "14px 16px" },
  card: { background: "#fff", border: "1px solid #EAE5D6", borderRadius: 5, padding: 18, marginBottom: 18 },
  cardTitle: { fontFamily: "'Zilla Slab', serif", fontWeight: 600, fontSize: 15.5, marginBottom: 12, color: "#20242B" },
  table: { width: "100%", borderCollapse: "collapse", fontSize: 12.8 },
  th: {
    textAlign: "left", padding: "8px 10px", fontSize: 10.5, letterSpacing: "0.05em", color: "#8A8371",
    borderBottom: "1.5px solid #E4DFD0", textTransform: "uppercase", fontWeight: 600,
  },
  td: { padding: "9px 10px", borderBottom: "1px solid #EFEBDF", color: "#3A362C" },
  tdMono: { padding: "9px 10px", borderBottom: "1px solid #EFEBDF", fontFamily: "'IBM Plex Mono', monospace", fontSize: 12, color: "#3A362C" },
  trHover: { cursor: "pointer" },
  label: { display: "block", fontSize: 11.5, color: "#6B6656", marginBottom: 5, letterSpacing: "0.02em" },
  input: {
    width: "100%", border: "1px solid #DED9CA", borderRadius: 4, padding: "8px 10px", fontSize: 13.5,
    fontFamily: "Inter", outline: "none", background: "#FBFAF6", boxSizing: "border-box",
  },
  primaryBtn: {
    display: "inline-flex", alignItems: "center", gap: 6, background: "#2B3A55", color: "#F5F2EA",
    border: "none", borderRadius: 4, padding: "9px 14px", fontSize: 13, cursor: "pointer", fontWeight: 600,
  },
  ghostBtn: {
    display: "inline-flex", alignItems: "center", gap: 6, background: "#fff", color: "#2B3A55",
    border: "1px solid #DED9CA", borderRadius: 4, padding: "8px 13px", fontSize: 13, cursor: "pointer", fontWeight: 500,
  },
  ghostBtnSmall: {
    display: "inline-flex", alignItems: "center", gap: 6, background: "transparent", color: "#B7C0CF",
    border: "1px solid #3E4A63", borderRadius: 4, padding: "6px 10px", fontSize: 11.5, cursor: "pointer",
  },
  iconBtn: {
    display: "inline-flex", alignItems: "center", justifyContent: "center", background: "transparent",
    border: "none", color: "#6B6656", cursor: "pointer", padding: 5, borderRadius: 4,
  },
  roleBtn: {
    flex: 1, padding: "8px 0", fontSize: 12.5, border: "1px solid #DED9CA", borderRadius: 4,
    background: "#FBFAF6", color: "#6B6656", cursor: "pointer",
  },
  roleBtnActive: { background: "#2B3A55", color: "#F5F2EA", borderColor: "#2B3A55" },
  modalOverlay: {
    position: "fixed", inset: 0, background: "rgba(20,20,10,0.35)", display: "flex",
    alignItems: "center", justifyContent: "center", zIndex: 50, padding: 16,
  },
  modalBox: { background: "#fff", borderRadius: 6, padding: 22, maxHeight: "88vh", overflowY: "auto" },
  toast: {
    position: "fixed", bottom: 22, left: "50%", transform: "translateX(-50%)", background: "#20242B",
    color: "#F5F2EA", padding: "10px 18px", borderRadius: 5, fontSize: 13, zIndex: 60,
    boxShadow: "0 4px 14px rgba(0,0,0,0.2)",
  },
  dropZone: {
    border: "1.5px dashed #DED9CA", borderRadius: 5, padding: "26px 10px", textAlign: "center",
    color: "#8A8371", cursor: "pointer", background: "#FBFAF6",
  },
  chip: {
    display: "inline-flex", alignItems: "center", gap: 6, fontFamily: "'IBM Plex Mono', monospace",
    fontSize: 11, color: "#3A362C", background: "#F1EEE3", border: "1px solid #E4DFD0",
    borderRadius: 3, padding: "3px 8px",
  },
  tableScroll: { width: "100%", overflowX: "auto" },
};

/* stamp + print styles injected once */
(function injectGlobalStyle() {
  if (typeof document === "undefined") return;
  if (document.getElementById("punchbook-global")) return;
  const el = document.createElement("style");
  el.id = "punchbook-global";
  el.innerHTML = `
    .stamp {
      display:inline-block; border:1.5px solid; border-radius:3px; font-family:'IBM Plex Mono',monospace;
      font-weight:600; letter-spacing:0.06em; transform: rotate(-1.5deg); background: rgba(0,0,0,0.015);
    }
    tr:hover td { background: #FBFAF6; }
    input:focus, select:focus { border-color:#2B3A55 !important; }
    button:disabled { opacity:0.45; cursor:not-allowed; }
    @media print {
      .no-print, aside { display:none !important; }
      body { background:#fff; }
    }
  `;
  document.head.appendChild(el);
})();
