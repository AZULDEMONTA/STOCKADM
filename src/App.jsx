import { useState, useEffect, useMemo, useCallback, useRef, Component } from "react";
import { supabase } from "./supabaseClient";
import {
  Upload, FileSpreadsheet, ChevronDown, ChevronRight, AlertTriangle,
  TrendingUp, TrendingDown, Search, RotateCcw, Info, Layers,
  LineChart as LineChartIcon, Gauge as GaugeIcon, Zap, Receipt, Boxes,
} from "lucide-react";
import * as XLSX from "xlsx";
import {
  LineChart, Line, XAxis, YAxis, CartesianGrid, Tooltip, Legend, ResponsiveContainer, ReferenceLine,
} from "recharts";

/* ---------------------------------- tokens ---------------------------------- */

const COLORS = {
  bg: "#F3F5F8",
  surface: "#FFFFFF",
  ink: "#16233F",
  inkMuted: "#63718A",
  border: "#E1E5EB",
  primary: "#2C6E8E",
  primaryDark: "#1B4A61",
  primarySoft: "#E7F0F4",
  success: "#2F8F6B",
  successSoft: "#E5F3EE",
  warning: "#C7862A",
  warningSoft: "#FBF1E2",
  critical: "#C1443A",
  criticalSoft: "#FBEAE8",
  violet: "#6E4FA3",
};

const LINE_PALETTE = ["#2C6E8E", "#C1443A", "#2F8F6B", "#D9A441", "#6E4FA3", "#3E6B99", "#8A5A44", "#4C8C8C"];

const FONT_LINK =
  "https://fonts.googleapis.com/css2?family=Space+Grotesk:wght@500;600;700&family=Inter:wght@400;500;600;700&family=IBM+Plex+Mono:wght@400;500;600&display=swap";

const MES_NUM = {
  ENERO: 1, FEBRERO: 2, MARZO: 3, ABRIL: 4, MAYO: 5, JUNIO: 6,
  JULIO: 7, AGOSTO: 8, SEPTIEMBRE: 9, SETIEMBRE: 9, OCTUBRE: 10, NOVIEMBRE: 11, DICIEMBRE: 12,
};
const MES_DISPLAY = { SETIEMBRE: "SEPTIEMBRE" };
const DEFAULT_TARGET = { min: 45, max: 60 };
const STORAGE_KEY = "inventory-data-v2";
const TOTAL_LABEL = "Total Empresa";

/* ---------------------------------- helpers ---------------------------------- */

function normalize(s) {
  return (s ?? "").toString().toUpperCase().normalize("NFD").replace(/[\u0300-\u036f]/g, "").trim();
}
function titleCase(s) {
  return (s ?? "").toString().toLowerCase().replace(/(^|\s)\S/g, (c) => c.toUpperCase());
}
function fmtMoney(n) {
  if (n == null || isNaN(n)) return "—";
  return "$ " + Math.round(n).toLocaleString("es-AR");
}
function fmtMoneyShort(n) {
  if (n == null || isNaN(n)) return "—";
  const abs = Math.abs(n);
  if (abs >= 1e6) return "$ " + (n / 1e6).toLocaleString("es-AR", { maximumFractionDigits: 1 }) + "M";
  if (abs >= 1e3) return "$ " + (n / 1e3).toLocaleString("es-AR", { maximumFractionDigits: 0 }) + "K";
  return fmtMoney(n);
}
function fmtDias(n) {
  if (n == null || isNaN(n)) return "—";
  return Math.round(n).toString();
}
function fmtPct(n) {
  if (n == null || isNaN(n)) return null;
  return `${n.toFixed(1)}%`;
}
function diasStockDe(stockVal, ventaVal) {
  if (stockVal == null || !ventaVal) return null;
  return stockVal / (ventaVal / 30);
}
function normKey(rkey) {
  return normalize(rkey.replace(/\|\|/g, " || "));
}

function parsePeriodo(raw) {
  const n = normalize(raw);
  const mes = Object.keys(MES_NUM).find((m) => n.includes(m));
  if (!mes) return null;
  const yearMatch = n.match(/(\d{2,4})/);
  if (!yearMatch) return null;
  let year = parseInt(yearMatch[1], 10);
  if (year < 100) year += 2000;
  const mesNum = MES_NUM[mes];
  const mesLabel = MES_DISPLAY[mes] || mes;
  return {
    key: `${year}-${String(mesNum).padStart(2, "0")}`,
    label: `${titleCase(mesLabel)} ${year}`,
    sortValue: year * 12 + mesNum,
  };
}

function findHeaderRow(wb, requiredKeywordGroups) {
  // requiredKeywordGroups: array of keyword-arrays; a header row must contain, for each group,
  // at least one cell whose normalized text includes ANY keyword from that group.
  for (const sheetName of wb.SheetNames) {
    const ws = wb.Sheets[sheetName];
    const rows = XLSX.utils.sheet_to_json(ws, { header: 1, defval: null, raw: true });
    for (let i = 0; i < Math.min(rows.length, 25); i++) {
      const row = rows[i] || [];
      const norms = row.map(normalize);
      const ok = requiredKeywordGroups.every((group) =>
        norms.some((n) => n && group.some((kw) => n.includes(kw)))
      );
      if (ok) return { rows, headerRowIdx: i, header: row, sheetName };
    }
  }
  return null;
}

/* ---------------------------------- parsers ---------------------------------- */

function parseStockFile(arrayBuffer, fileName) {
  const wb = XLSX.read(arrayBuffer, { type: "array" });
  const found = findHeaderRow(wb, [["GRUPO"], ["RUB"], ["PERIODO"], ["NETO"]]);
  if (!found) {
    throw new Error("No encontré una hoja con columnas de Grupo, Rubro, Período y Neto. Revisá el archivo de stock.");
  }
  const { rows, headerRowIdx, header } = found;

  let grupoCol = -1, rubroCol = -1, periodoCol = -1, netoCol = -1;
  header.forEach((cell, idx) => {
    const n = normalize(cell);
    if (!n) return;
    if (n.includes("GRUPO")) grupoCol = idx;
    else if (n.includes("RUB")) rubroCol = idx;
    else if (n.includes("PERIODO")) periodoCol = idx;
    else if (n.includes("NETO")) netoCol = idx;
  });

  const agg = {};
  const periodsSeen = new Map();

  for (let i = headerRowIdx + 1; i < rows.length; i++) {
    const row = rows[i];
    if (!row) continue;
    const grupoRaw = row[grupoCol], rubroRaw = row[rubroCol], periodoRaw = row[periodoCol], netoRaw = row[netoCol];
    if (grupoRaw == null || rubroRaw == null || periodoRaw == null) continue;
    const grupo = String(grupoRaw).trim();
    const rubro = String(rubroRaw).trim();
    if (!grupo || !rubro) continue;
    const periodo = parsePeriodo(periodoRaw);
    if (!periodo) continue;
    const neto = typeof netoRaw === "number" ? netoRaw : (parseFloat(netoRaw) || 0);

    const rkey = `${grupo}||${rubro}`;
    const fullKey = `${rkey}##${periodo.key}`;
    agg[fullKey] = (agg[fullKey] || 0) + neto;
    if (!periodsSeen.has(periodo.key)) periodsSeen.set(periodo.key, periodo);
  }

  const rowsOut = {};
  Object.entries(agg).forEach(([fullKey, val]) => {
    const [rkey, periodoKey] = fullKey.split("##");
    const [grupo, rubro] = rkey.split("||");
    if (!rowsOut[rkey]) rowsOut[rkey] = { grupo, rubro, months: {} };
    rowsOut[rkey].months[periodoKey] = val;
  });

  if (Object.keys(rowsOut).length === 0) {
    throw new Error("No se encontraron filas de stock válidas en el archivo.");
  }

  return { rows: rowsOut, periods: [...periodsSeen.values()], fileName };
}

function monthKeyFromDateCell(val) {
  let d;
  if (val instanceof Date) d = val;
  else if (typeof val === "number") d = new Date(Math.round((val - 25569) * 86400 * 1000));
  else if (typeof val === "string" && val.trim()) { d = new Date(val); }
  else return null;
  if (!d || isNaN(d.getTime())) return null;
  return `${d.getUTCFullYear()}-${String(d.getUTCMonth() + 1).padStart(2, "0")}`;
}

function parseVentaFile(arrayBuffer, fileName) {
  const wb = XLSX.read(arrayBuffer, { type: "array", cellDates: true });
  const found = findHeaderRow(wb, [["GRUPO"], ["RUB"], ["VENTA"]]);
  if (!found) {
    throw new Error("No encontré una hoja con columnas de Grupo, Rubro y Venta. Revisá el archivo de venta.");
  }
  const { rows, headerRowIdx, header } = found;

  let grupoCol = -1, rubroCol = -1, ventaCol = -1, fechaCol = -1;
  header.forEach((cell, idx) => {
    const n = normalize(cell);
    if (!n) return;
    if (n.includes("GRUPO")) grupoCol = idx;
    else if (n.includes("RUB")) rubroCol = idx;
    else if (n.includes("VENTA")) ventaCol = idx;
    else if (n.includes("FECHA")) fechaCol = idx;
  });

  const agg = {};
  const mesesVistos = new Set();
  for (let i = headerRowIdx + 1; i < rows.length; i++) {
    const row = rows[i];
    if (!row) continue;
    const grupoRaw = row[grupoCol], rubroRaw = row[rubroCol], ventaRaw = row[ventaCol];
    if (grupoRaw == null || rubroRaw == null) continue;
    const grupo = String(grupoRaw).trim();
    const rubro = String(rubroRaw).trim();
    if (!grupo || !rubro) continue;
    const venta = typeof ventaRaw === "number" ? ventaRaw : (parseFloat(ventaRaw) || 0);
    const rkey = `${grupo}||${rubro}`;
    agg[rkey] = (agg[rkey] || 0) + venta;
    if (fechaCol !== -1) {
      const mk = monthKeyFromDateCell(row[fechaCol]);
      if (mk) mesesVistos.add(mk);
    }
  }

  if (Object.keys(agg).length === 0) {
    throw new Error("No se encontraron filas de venta válidas en el archivo.");
  }

  const numMeses = mesesVistos.size || 1;
  const promedio = {};
  Object.entries(agg).forEach(([rkey, total]) => { promedio[rkey] = total / numMeses; });

  return { venta: promedio, fileName, rowCount: Object.keys(promedio).length, numMeses, hasFecha: fechaCol !== -1 };
}

/* ---------------------------------- data merge ---------------------------------- */

function emptyData() {
  return { periodsMeta: {}, rows: {}, venta: {}, targets: {}, uploadLog: [], ventaMeta: null };
}

function mergeStockData(prevData, parseResult) {
  const data = prevData ? JSON.parse(JSON.stringify(prevData)) : emptyData();
  parseResult.periods.forEach((p) => { data.periodsMeta[p.key] = { label: p.label, sortValue: p.sortValue }; });
  Object.entries(parseResult.rows).forEach(([rkey, r]) => {
    if (!data.rows[rkey]) data.rows[rkey] = { grupo: r.grupo, rubro: r.rubro, months: {} };
    Object.entries(r.months).forEach(([pk, val]) => { data.rows[rkey].months[pk] = val; });
    if (!data.targets[rkey]) data.targets[rkey] = { ...DEFAULT_TARGET };
  });
  data.uploadLog.push({
    id: `s_${Date.now()}`, type: "stock", fileName: parseResult.fileName,
    uploadedAt: new Date().toISOString(),
    periods: parseResult.periods.map((p) => p.label),
    rowCount: Object.keys(parseResult.rows).length,
  });
  return data;
}

function applyVentaData(prevData, parseResult) {
  const data = prevData ? JSON.parse(JSON.stringify(prevData)) : emptyData();
  data.venta = parseResult.venta;
  data.ventaMeta = { fileName: parseResult.fileName, uploadedAt: new Date().toISOString(), rowCount: parseResult.rowCount, numMeses: parseResult.numMeses, hasFecha: parseResult.hasFecha };
  data.uploadLog.push({
    id: `v_${Date.now()}`, type: "venta", fileName: parseResult.fileName,
    uploadedAt: new Date().toISOString(), rowCount: parseResult.rowCount,
  });
  return data;
}

function withTimeout(promise, ms) {
  return Promise.race([
    promise,
    new Promise((_, reject) => setTimeout(() => reject(new Error("timeout")), ms)),
  ]);
}

async function loadData() {
  try {
    const { data: row, error } = await withTimeout(
      supabase.from("app_storage").select("value").eq("key", STORAGE_KEY).maybeSingle(),
      8000
    );
    if (error) throw error;
    return { ok: true, data: row ? row.value : emptyData() };
  } catch (e) {
    console.warn("No había datos guardados previos (o falló la lectura); arrancando vacío.", e);
    return { ok: true, data: emptyData() };
  }
}

async function saveData(data, retries = 2) {
  for (let attempt = 1; attempt <= retries; attempt++) {
    try {
      const { error } = await withTimeout(
        supabase.from("app_storage").upsert({ key: STORAGE_KEY, value: data }, { onConflict: "key" }),
        8000
      );
      if (!error) return true;
      throw error;
    } catch (e) {
      if (attempt === retries) {
        console.error("Error guardando datos", e);
        return false;
      }
      await new Promise((r) => setTimeout(r, 400 * attempt));
    }
  }
  return false;
}

/* ---------------------------------- demo data ---------------------------------- */

const DEMO_GRUPOS = [
  { grupo: "Padoc S.A.", rubros: ["Boos", "Welling Polo C.", "Natural Spirit", "Tascani", "Haramain", "Mua", "Stone", "Feraud", "Portsaid"] },
  { grupo: "Nortex Distribuciones", rubros: ["Denim Co", "Urban Flex", "Cotton Base", "Lino Rey", "Terra Wear"] },
  { grupo: "Sur Textil S.A.", rubros: ["Alpaca Norte", "Merino Sur", "Wool House", "Cashmere Lab"] },
];
const DEMO_PERIODOS = [
  { key: "2025-12", label: "Diciembre 2025", sortValue: 2025 * 12 + 12 },
  { key: "2026-02", label: "Febrero 2026", sortValue: 2026 * 12 + 2 },
  { key: "2026-06", label: "Junio 2026", sortValue: 2026 * 12 + 6 },
];

function generateDemoData() {
  const data = emptyData();
  DEMO_PERIODOS.forEach((p) => { data.periodsMeta[p.key] = { label: p.label, sortValue: p.sortValue }; });

  DEMO_GRUPOS.forEach((g) => {
    g.rubros.forEach((rb) => {
      const rkey = `${g.grupo}||${rb}`;
      let stock = Math.round(300000 + Math.random() * 30000000);
      const months = {};
      DEMO_PERIODOS.forEach((p, i) => {
        if (i > 0) stock = Math.round(stock * (0.75 + Math.random() * 0.7));
        months[p.key] = stock;
      });
      data.rows[rkey] = { grupo: g.grupo, rubro: rb, months };
      data.targets[rkey] = { ...DEFAULT_TARGET };
      const targetDias = Math.round(15 + Math.random() * 130);
      const lastStock = months[DEMO_PERIODOS[DEMO_PERIODOS.length - 1].key];
      data.venta[rkey] = Math.round((lastStock / targetDias) * 30);
    });
  });

  data.ventaMeta = { fileName: "Datos de ejemplo (simulados)", uploadedAt: new Date().toISOString(), rowCount: Object.keys(data.venta).length, numMeses: 1, hasFecha: true };
  data.uploadLog.push({ id: `demo_${Date.now()}`, type: "stock", fileName: "Datos de ejemplo (simulados)", uploadedAt: new Date().toISOString(), periods: DEMO_PERIODOS.map((p) => p.label), rowCount: Object.keys(data.rows).length });
  data.uploadLog.push({ id: `demo2_${Date.now()}`, type: "venta", fileName: "Datos de ejemplo (simulados)", uploadedAt: new Date().toISOString(), rowCount: Object.keys(data.venta).length });
  return data;
}

/* ---------------------------------- small UI pieces ---------------------------------- */

function ConfirmModal({ open, title, message, confirmLabel, onConfirm, onCancel }) {
  if (!open) return null;
  return (
    <div
      onClick={onCancel}
      style={{
        position: "fixed", inset: 0, background: "rgba(22,35,63,0.45)",
        display: "flex", alignItems: "center", justifyContent: "center", zIndex: 50,
      }}
    >
      <div
        onClick={(e) => e.stopPropagation()}
        style={{
          background: COLORS.surface, borderRadius: 14, padding: "22px 24px", maxWidth: 380,
          width: "90%", boxShadow: "0 12px 40px rgba(22,35,63,0.25)",
        }}
      >
        <div style={{ fontWeight: 700, fontSize: 15, color: COLORS.ink, marginBottom: 8 }}>{title}</div>
        <div style={{ fontSize: 13, color: COLORS.inkMuted, lineHeight: 1.5, marginBottom: 20 }}>{message}</div>
        <div style={{ display: "flex", gap: 8, justifyContent: "flex-end" }}>
          <button
            onClick={onCancel}
            style={{ border: `1px solid ${COLORS.border}`, background: "none", color: COLORS.inkMuted, padding: "8px 16px", borderRadius: 8, fontSize: 12.5, fontWeight: 600, cursor: "pointer" }}
          >
            Cancelar
          </button>
          <button
            onClick={onConfirm}
            style={{ border: "none", background: COLORS.critical, color: "#fff", padding: "8px 16px", borderRadius: 8, fontSize: 12.5, fontWeight: 600, cursor: "pointer" }}
          >
            {confirmLabel || "Confirmar"}
          </button>
        </div>
      </div>
    </div>
  );
}

function Gauge({ dias, min, max }) {
  if (dias == null) return <div style={{ fontSize: 11, color: COLORS.inkMuted, width: 118 }}>sin venta cargada</div>;
  min = Math.round(min);
  max = Math.round(max);
  let color = COLORS.success;
  let segment = 1, t = 0.5;
  if (dias < min) {
    color = COLORS.critical;
    segment = 0;
    t = min > 0 ? Math.max(0, Math.min(1, dias / min)) : 0;
  } else if (dias > max) {
    color = COLORS.warning;
    segment = 2;
    const span = Math.max(max * 0.6, 10);
    t = Math.max(0, Math.min(1, (dias - max) / span));
  } else {
    color = COLORS.success;
    segment = 1;
    t = max > min ? (dias - min) / (max - min) : 0.5;
  }
  const markerPct = ((segment + t) / 3) * 100;

  return (
    <div style={{ width: 118 }}>
      <div style={{ display: "flex", alignItems: "baseline", gap: 4, marginBottom: 6 }}>
        <span style={{ fontFamily: "'Space Grotesk', sans-serif", fontSize: 16, fontWeight: 700, color }}>{fmtDias(dias)}</span>
        <span style={{ fontSize: 11, fontWeight: 600, color: COLORS.inkMuted }}>días</span>
      </div>
      <div style={{ position: "relative", height: 14 }}>
        <div style={{
          position: "absolute", left: `calc(${markerPct}% - 4px)`, top: 0,
          width: 0, height: 0,
          borderLeft: "4px solid transparent", borderRight: "4px solid transparent",
          borderTop: `5px solid ${color}`,
        }} />
        <div style={{ position: "absolute", top: 7, left: 0, right: 0, display: "flex", height: 6, borderRadius: 4, overflow: "hidden" }}>
          <div style={{ flex: 1, background: segment === 0 ? COLORS.critical : COLORS.criticalSoft }} />
          <div style={{ flex: 1, background: segment === 1 ? COLORS.success : COLORS.successSoft }} />
          <div style={{ flex: 1, background: segment === 2 ? COLORS.warning : COLORS.warningSoft }} />
        </div>
      </div>
      <div style={{ display: "flex", fontSize: 9.5, color: COLORS.inkMuted, marginTop: 3, fontFamily: "IBM Plex Mono, monospace" }}>
        <span style={{ flex: 1, textAlign: "center" }}>{"<"}{min}</span>
        <span style={{ flex: 1, textAlign: "center" }}>{min}–{max}</span>
        <span style={{ flex: 1, textAlign: "center" }}>{">"}{max}</span>
      </div>
    </div>
  );
}

function KpiCard({ label, value, sub, icon: Icon, accent }) {
  return (
    <div style={{ background: COLORS.surface, border: `1px solid ${COLORS.border}`, borderRadius: 10, padding: "9px 12px", flex: "1 1 150px" }}>
      <div style={{ display: "flex", alignItems: "center", gap: 5, color: COLORS.inkMuted, fontSize: 10, fontWeight: 600, textTransform: "uppercase", letterSpacing: 0.4 }}>
        {Icon && <Icon size={12} strokeWidth={2.2} />} {label}
      </div>
      <div style={{ fontFamily: "'Space Grotesk', sans-serif", fontSize: 18, fontWeight: 700, color: accent || COLORS.ink, marginTop: 3, letterSpacing: -0.3, lineHeight: 1.15 }}>
        {value}
      </div>
      {sub && <div style={{ fontSize: 10.5, color: COLORS.inkMuted, marginTop: 1 }}>{sub}</div>}
    </div>
  );
}

const PERIOD_CELL_STYLE = { borderLeft: `1px solid ${COLORS.border}`, paddingLeft: 12, paddingRight: 4 };

function PeriodCell({ value, prevValue, pct }) {
  if (value == null) return <div style={{ ...PERIOD_CELL_STYLE, textAlign: "right", color: COLORS.inkMuted, fontSize: 12 }}>—</div>;
  let delta = null;
  if (prevValue != null && prevValue !== 0) delta = ((value - prevValue) / prevValue) * 100;
  return (
    <div style={{ ...PERIOD_CELL_STYLE, textAlign: "right" }}>
      <div style={{ fontFamily: "IBM Plex Mono, monospace", fontSize: 12.5 }}>{fmtMoney(value)}</div>
      {delta != null && (
        <div style={{
          fontSize: 10.5, fontFamily: "IBM Plex Mono, monospace", marginTop: 1,
          color: delta > 0 ? COLORS.success : delta < 0 ? COLORS.critical : COLORS.inkMuted,
        }}>
          {delta > 0 ? "▲" : delta < 0 ? "▼" : "•"} {Math.abs(delta).toFixed(1)}%
        </div>
      )}
      {pct != null && (
        <div style={{ fontSize: 10, color: COLORS.primary, marginTop: 1, fontFamily: "IBM Plex Mono, monospace" }}>
          {pct.toFixed(1)}%
        </div>
      )}
    </div>
  );
}

function PctTag({ pct }) {
  if (pct == null) return null;
  return (
    <div style={{ fontSize: 10, color: COLORS.violet, marginTop: 1, fontFamily: "IBM Plex Mono, monospace" }}>
      {pct.toFixed(1)}%
    </div>
  );
}

const TOTAL_PERIOD_CELL_STYLE = { borderLeft: "1px solid rgba(255,255,255,0.2)", paddingLeft: 12, paddingRight: 4 };

function TotalPeriodCell({ value, prevValue }) {
  if (value == null) return <div style={{ ...TOTAL_PERIOD_CELL_STYLE, textAlign: "right", color: "rgba(255,255,255,0.6)", fontSize: 12 }}>—</div>;
  let delta = null;
  if (prevValue != null && prevValue !== 0) delta = ((value - prevValue) / prevValue) * 100;
  return (
    <div style={{ ...TOTAL_PERIOD_CELL_STYLE, textAlign: "right" }}>
      <div style={{ fontFamily: "IBM Plex Mono, monospace", fontSize: 12.5, fontWeight: 700 }}>{fmtMoney(value)}</div>
      {delta != null && (
        <div style={{
          fontSize: 10.5, fontFamily: "IBM Plex Mono, monospace", marginTop: 1,
          color: delta > 0 ? "#9FE3C4" : delta < 0 ? "#F4ABA3" : "rgba(255,255,255,0.7)",
        }}>
          {delta > 0 ? "▲" : delta < 0 ? "▼" : "•"} {Math.abs(delta).toFixed(1)}%
        </div>
      )}
    </div>
  );
}

const STATUS_META = {
  rango: { label: "En rango", shortLabel: "en rango", bg: COLORS.successSoft, fg: COLORS.success },
  riesgo: { label: "Riesgo quiebre", shortLabel: "riesgo", bg: COLORS.criticalSoft, fg: COLORS.critical },
  exceso: { label: "Exceso", shortLabel: "exceso", bg: COLORS.warningSoft, fg: COLORS.warning },
  sinDatos: { label: "Sin datos", shortLabel: "sin datos", bg: COLORS.bg, fg: COLORS.inkMuted },
};
function getStatus(dias, min, max) {
  if (dias == null) return "sinDatos";
  if (dias < min) return "riesgo";
  if (dias > max) return "exceso";
  return "rango";
}

function StatusBadge({ dias, min, max }) {
  const status = getStatus(dias, min, max);
  if (status === "sinDatos") return null;
  const meta = STATUS_META[status];
  return (
    <span style={{ background: meta.bg, color: meta.fg, fontSize: 10.5, fontWeight: 600, padding: "3px 8px", borderRadius: 20, whiteSpace: "nowrap" }}>
      {meta.label}
    </span>
  );
}

function StatusCountBadges({ counts }) {
  const entries = ["rango", "riesgo", "exceso"].filter((k) => counts[k] > 0);
  if (entries.length === 0) return null;
  return (
    <div style={{ display: "flex", gap: 5, marginLeft: 10 }}>
      {entries.map((key) => (
        <span key={key} style={{
          display: "inline-flex", alignItems: "center", gap: 3,
          background: STATUS_META[key].bg, color: STATUS_META[key].fg,
          fontSize: 10.5, fontWeight: 700, padding: "2px 8px", borderRadius: 20, whiteSpace: "nowrap",
        }}>
          {counts[key]} {STATUS_META[key].shortLabel}
        </span>
      ))}
    </div>
  );
}

/* ---------------------------------- views ---------------------------------- */

function useVentaLookup(data) {
  return useMemo(() => {
    const m = {};
    Object.entries(data.venta || {}).forEach(([k, v]) => { m[normKey(k)] = v; });
    return (rkey) => m[normKey(rkey)] ?? null;
  }, [data.venta]);
}

function useMonthOrder(data) {
  return useMemo(() => {
    return Object.entries(data.periodsMeta)
      .sort((a, b) => a[1].sortValue - b[1].sortValue)
      .map(([key, meta]) => ({ key, ...meta }));
  }, [data.periodsMeta]);
}

function DashboardView({ data, targets, setTargets, expanded, toggleExpand, setExpanded }) {
  const [search, setSearch] = useState("");
  const [grupoFilter, setGrupoFilter] = useState("TODOS");
  const [statusFilter, setStatusFilter] = useState("TODOS");
  const [sortBy, setSortBy] = useState("stock");
  const [confirmApplyTarget, setConfirmApplyTarget] = useState(false);
  const [hiddenMonths, setHiddenMonths] = useState(() => new Set());
  const [monthFilterOpen, setMonthFilterOpen] = useState(false);
  const getVenta = useVentaLookup(data);
  const monthOrder = useMonthOrder(data);
  const latest = monthOrder[monthOrder.length - 1];
  const prev = monthOrder[monthOrder.length - 2];
  const visibleMonthOrder = useMemo(
    () => monthOrder.filter((m) => !hiddenMonths.has(m.key)),
    [monthOrder, hiddenMonths]
  );
  const toggleMonthVisible = (key) => {
    setHiddenMonths((prevH) => {
      const next = new Set(prevH);
      if (next.has(key)) {
        next.delete(key);
      } else {
        if (monthOrder.length - (next.size + 1) < 1) return prevH; // no permitir ocultar todos
        next.add(key);
      }
      return next;
    });
  };

  const grupos = useMemo(() => {
    const set = new Set(Object.values(data.rows).map((r) => r.grupo));
    return [...set].sort();
  }, [data.rows]);

  const globalTotals = useMemo(() => {
    const stockByPeriod = {};
    monthOrder.forEach((m) => { stockByPeriod[m.key] = 0; });
    Object.values(data.rows).forEach((r) => {
      monthOrder.forEach((m) => { stockByPeriod[m.key] += r.months[m.key] || 0; });
    });
    const ventaTotal = Object.values(data.venta || {}).reduce((s, v) => s + (v || 0), 0);
    return { stockByPeriod, ventaTotal };
  }, [data.rows, data.venta, monthOrder]);

  const grouped = useMemo(() => {
    const list = grupos
      .filter((g) => grupoFilter === "TODOS" || g === grupoFilter)
      .map((g) => {
        let rubrosAll = Object.entries(data.rows)
          .filter(([, r]) => r.grupo === g)
          .map(([rkey, r]) => ({ rkey, ...r }));
        if (search.trim()) {
          const q = normalize(search);
          rubrosAll = rubrosAll.filter((r) => normalize(r.rubro).includes(q));
        }
        rubrosAll = rubrosAll.map((r) => {
          const t = targets[r.rkey] || DEFAULT_TARGET;
          const venta = getVenta(r.rkey);
          const dias = diasStockDe(r.months[latest?.key], venta);
          return { ...r, venta, dias, target: t, status: getStatus(dias, t.min, t.max) };
        });

        const statusCounts = { rango: 0, riesgo: 0, exceso: 0, sinDatos: 0 };
        rubrosAll.forEach((r) => { statusCounts[r.status]++; });

        let rubros = statusFilter === "TODOS" ? rubrosAll : rubrosAll.filter((r) => r.status === statusFilter);
        if (sortBy === "venta") rubros.sort((a, b) => (b.venta || 0) - (a.venta || 0));
        else if (sortBy === "nombre") rubros.sort((a, b) => a.rubro.localeCompare(b.rubro));
        else rubros.sort((a, b) => (b.months[latest?.key] || 0) - (a.months[latest?.key] || 0));

        const periodTotals = monthOrder.map((m) => rubrosAll.reduce((s, r) => s + (r.months[m.key] || 0), 0));
        const stockLatest = periodTotals[periodTotals.length - 1] || 0;
        const ventaTotal = rubrosAll.reduce((s, r) => s + (getVenta(r.rkey) || 0), 0);
        const diasGrupo = diasStockDe(stockLatest, ventaTotal);
        let wMin = 0, wMax = 0, wDen = 0;
        rubrosAll.forEach((r) => {
          const w = r.months[latest?.key] || 0;
          wMin += r.target.min * w;
          wMax += r.target.max * w;
          wDen += w;
        });
        const targetGrupo = wDen ? { min: Math.round(wMin / wDen), max: Math.round(wMax / wDen) } : DEFAULT_TARGET;
        return { grupo: g, rubros, statusCounts, periodTotals, stockLatest, ventaTotal, diasGrupo, targetGrupo };
      })
      .filter((g) => g.rubros.length > 0);

    if (sortBy === "stock") list.sort((a, b) => b.stockLatest - a.stockLatest);
    else if (sortBy === "venta") list.sort((a, b) => b.ventaTotal - a.ventaTotal);
    else list.sort((a, b) => a.grupo.localeCompare(b.grupo));
    return list;
  }, [grupos, grupoFilter, search, statusFilter, sortBy, data.rows, latest, monthOrder, getVenta, targets]);

  const kpis = useMemo(() => {
    const allRubros = Object.entries(data.rows).map(([rkey, r]) => ({ rkey, ...r }));
    let stockActual = 0, stockPrev = 0, weightedNum = 0, weightedDen = 0;
    allRubros.forEach((r) => {
      const v = r.months[latest?.key] || 0;
      stockActual += v;
      if (prev) stockPrev += r.months[prev.key] || 0;
      const venta = getVenta(r.rkey);
      const dias = diasStockDe(v, venta);
      if (dias != null) {
        weightedNum += dias * v;
        weightedDen += v;
      }
    });
    const ventaTotal = Object.values(data.venta || {}).reduce((s, v) => s + (v || 0), 0);
    const avgDias = weightedDen ? weightedNum / weightedDen : null;
    const variacion = stockPrev ? ((stockActual - stockPrev) / stockPrev) * 100 : null;
    return { stockActual, variacion, avgDias, ventaTotal, totalRubros: allRubros.length };
  }, [data.rows, data.venta, targets, latest, prev, getVenta]);

  const globalPeriodTotals = monthOrder.map((m) => globalTotals.stockByPeriod[m.key] || 0);

  const updateTarget = (key, field, value) => {
    const n = Math.max(0, Number(value) || 0);
    setTargets((prevT) => ({ ...prevT, [key]: { ...(prevT[key] || DEFAULT_TARGET), [field]: n } }));
  };

  const applyTargetToAll = (min, max) => {
    setTargets((prevT) => {
      const next = { ...prevT };
      Object.keys(data.rows).forEach((rkey) => { next[rkey] = { min, max }; });
      return next;
    });
  };

  return (
    <div>
      <ConfirmModal
        open={confirmApplyTarget}
        title="Aplicar objetivo a todos los rubros"
        message="¿Aplicar el objetivo 45–60 días a todos los rubros? Esto sobrescribe cualquier objetivo personalizado que hayas configurado."
        confirmLabel="Aplicar"
        onCancel={() => setConfirmApplyTarget(false)}
        onConfirm={() => { applyTargetToAll(45, 60); setConfirmApplyTarget(false); }}
      />
      <div style={{ display: "flex", gap: 8, flexWrap: "wrap", marginBottom: 14 }}>
        <KpiCard label="Stock actual" value={fmtMoneyShort(kpis.stockActual)} sub={latest ? latest.label : ""} icon={Layers} />
        <KpiCard
          label="Variación vs. período anterior"
          value={kpis.variacion == null ? "—" : `${kpis.variacion >= 0 ? "+" : ""}${kpis.variacion.toFixed(1)}%`}
          sub={prev ? `vs. ${prev.label}` : "sin período previo"}
          icon={kpis.variacion >= 0 ? TrendingUp : TrendingDown}
          accent={kpis.variacion == null ? COLORS.ink : kpis.variacion >= 0 ? COLORS.warning : COLORS.success}
        />
        <KpiCard label="Días de stock (prom. ponderado)" value={kpis.avgDias == null ? "—" : `${Math.round(kpis.avgDias)}d`} sub={`sobre ${kpis.totalRubros} rubros`} icon={GaugeIcon} />
        <KpiCard label="Venta promedio total" value={fmtMoneyShort(kpis.ventaTotal)} sub="suma del archivo de venta" icon={Receipt} />
      </div>

      {!data.ventaMeta && (
        <div style={{ display: "flex", gap: 6, alignItems: "flex-start", background: COLORS.warningSoft, color: "#8A5E1F", fontSize: 12.5, padding: "9px 12px", borderRadius: 8, marginBottom: 14 }}>
          <Info size={14} style={{ marginTop: 1, flexShrink: 0 }} />
          Todavía no cargaste el archivo de venta promedio — los días de stock no se pueden calcular hasta que lo subas.
        </div>
      )}

      <div style={{ display: "flex", gap: 10, marginBottom: 14, flexWrap: "wrap", alignItems: "center" }}>
        <div style={{ position: "relative" }}>
          <Search size={14} style={{ position: "absolute", left: 10, top: 9, color: COLORS.inkMuted }} />
          <input
            value={search}
            onChange={(e) => setSearch(e.target.value)}
            placeholder="Buscar rubro..."
            style={{ padding: "7px 10px 7px 30px", borderRadius: 8, border: `1px solid ${COLORS.border}`, fontSize: 13, width: 200, fontFamily: "Inter, sans-serif" }}
          />
        </div>
        <select
          value={grupoFilter}
          onChange={(e) => setGrupoFilter(e.target.value)}
          style={{ padding: "7px 10px", borderRadius: 8, border: `1px solid ${COLORS.border}`, fontSize: 13, fontFamily: "Inter, sans-serif", color: COLORS.ink, background: COLORS.surface }}
        >
          <option value="TODOS">Todos los proveedores</option>
          {grupos.map((g) => <option key={g} value={g}>{g}</option>)}
        </select>
        <div style={{ display: "flex", alignItems: "center", gap: 8, marginLeft: "auto" }}>
          <span style={{ fontSize: 12, color: COLORS.inkMuted }}>Ordenar proveedores por</span>
          <div style={{ display: "flex", gap: 4, background: COLORS.surface, border: `1px solid ${COLORS.border}`, borderRadius: 8, padding: 3 }}>
            {[{ id: "stock", label: "Stock" }, { id: "venta", label: "Venta" }, { id: "nombre", label: "Nombre" }].map((opt) => (
              <button
                key={opt.id}
                onClick={() => setSortBy(opt.id)}
                style={{
                  border: "none", padding: "5px 11px", borderRadius: 6, fontSize: 12, cursor: "pointer",
                  background: sortBy === opt.id ? COLORS.primary : "transparent",
                  color: sortBy === opt.id ? "#fff" : COLORS.inkMuted,
                  fontWeight: sortBy === opt.id ? 700 : 500,
                }}
              >
                {opt.label}
              </button>
            ))}
          </div>
        </div>
      </div>

      <div style={{ display: "flex", alignItems: "center", gap: 8, marginBottom: 14, flexWrap: "wrap" }}>
        <span style={{ fontSize: 12, color: COLORS.inkMuted }}>Filtrar rubros por etiqueta</span>
        <div style={{ display: "flex", gap: 6, flexWrap: "wrap" }}>
          {["TODOS", "rango", "riesgo", "exceso"].map((key) => {
            const active = statusFilter === key;
            const meta = key === "TODOS" ? { label: "Todos", fg: COLORS.primaryDark, bg: COLORS.primarySoft } : STATUS_META[key];
            return (
              <button
                key={key}
                onClick={() => setStatusFilter(key)}
                style={{
                  border: `1px solid ${active ? meta.fg : COLORS.border}`,
                  background: active ? meta.bg : COLORS.surface,
                  color: meta.fg,
                  padding: "5px 12px", borderRadius: 20, fontSize: 12, fontWeight: 600, cursor: "pointer",
                }}
              >
                {meta.label}
              </button>
            );
          })}
        </div>
        <button
          onClick={() => setExpanded(new Set())}
          style={{ display: "flex", alignItems: "center", gap: 6, background: "none", border: `1px solid ${COLORS.border}`, color: COLORS.primaryDark, padding: "5px 12px", borderRadius: 8, fontSize: 12, fontWeight: 600, cursor: "pointer" }}
        >
          <ChevronRight size={13} /> Colapsar todos
        </button>
        <div style={{ position: "relative", marginLeft: "auto" }}>
          <button
            onClick={() => setMonthFilterOpen((o) => !o)}
            style={{ display: "flex", alignItems: "center", gap: 6, background: hiddenMonths.size > 0 ? COLORS.primarySoft : "none", border: `1px solid ${hiddenMonths.size > 0 ? COLORS.primaryDark : COLORS.border}`, color: COLORS.primaryDark, padding: "5px 12px", borderRadius: 8, fontSize: 12, fontWeight: 600, cursor: "pointer" }}
          >
            <Layers size={13} />
            Meses {hiddenMonths.size > 0 ? `(${monthOrder.length - hiddenMonths.size}/${monthOrder.length})` : ""}
            <ChevronDown size={13} />
          </button>
          {monthFilterOpen && (
            <>
              <div onClick={() => setMonthFilterOpen(false)} style={{ position: "fixed", inset: 0, zIndex: 5 }} />
              <div style={{
                position: "absolute", right: 0, top: "calc(100% + 6px)", zIndex: 6,
                background: COLORS.surface, border: `1px solid ${COLORS.border}`, borderRadius: 10,
                boxShadow: "0 8px 24px rgba(22,35,63,0.14)", padding: 10, minWidth: 190, maxHeight: 280, overflowY: "auto",
              }}>
                <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 6 }}>
                  <span style={{ fontSize: 11, fontWeight: 700, color: COLORS.inkMuted, textTransform: "uppercase", letterSpacing: 0.4 }}>
                    Mostrar períodos
                  </span>
                  <button
                    onClick={() => setHiddenMonths(new Set())}
                    style={{ background: "none", border: "none", color: COLORS.primary, fontSize: 11, fontWeight: 600, cursor: "pointer" }}
                  >
                    Mostrar todos
                  </button>
                </div>
                {monthOrder.map((m) => {
                  const checked = !hiddenMonths.has(m.key);
                  return (
                    <label
                      key={m.key}
                      style={{ display: "flex", alignItems: "center", gap: 8, padding: "5px 4px", fontSize: 13, color: COLORS.ink, cursor: "pointer" }}
                    >
                      <input type="checkbox" checked={checked} onChange={() => toggleMonthVisible(m.key)} style={{ cursor: "pointer" }} />
                      {m.label}
                    </label>
                  );
                })}
              </div>
            </>
          )}
        </div>
        <button
          onClick={() => setConfirmApplyTarget(true)}
          style={{ display: "flex", alignItems: "center", gap: 6, background: "none", border: `1px solid ${COLORS.border}`, color: COLORS.primaryDark, padding: "5px 12px", borderRadius: 8, fontSize: 12, fontWeight: 600, cursor: "pointer" }}
        >
          <GaugeIcon size={13} /> Aplicar objetivo 45–60 a todos
        </button>
      </div>

      <div style={{ display: "flex", gap: 6, alignItems: "flex-start", background: COLORS.primarySoft, color: COLORS.primaryDark, fontSize: 12, padding: "8px 10px", borderRadius: 8, marginBottom: 14 }}>
        <Info size={14} style={{ marginTop: 1, flexShrink: 0 }} />
        Los porcentajes de stock (en celeste) se calculan sobre el stock total de <strong>todos</strong> los proveedores/rubros en cada período. Los porcentajes de venta (en violeta) se calculan sobre la venta promedio total.
      </div>

      <div style={{ background: COLORS.surface, border: `1px solid ${COLORS.border}`, borderRadius: 12, overflow: "auto", maxHeight: "72vh", position: "relative" }}>
        <div style={{ minWidth: 560 + visibleMonthOrder.length * 120 }}>
          <div style={{
            display: "grid", gridTemplateColumns: `1.6fr repeat(${visibleMonthOrder.length}, 120px) 150px 190px 130px`,
            padding: "10px 16px", background: COLORS.bg, borderBottom: `1px solid ${COLORS.border}`,
            fontSize: 11, fontWeight: 700, color: COLORS.inkMuted, textTransform: "uppercase", letterSpacing: 0.4,
            position: "sticky", top: 0, zIndex: 3,
          }}>
            <div>Proveedor / Rubro</div>
            {monthOrder.map((m) => (
              !hiddenMonths.has(m.key) && (
                <div key={m.key} style={{ ...PERIOD_CELL_STYLE, textAlign: "right" }}>{m.label}</div>
              )
            ))}
            <div style={{ textAlign: "right", borderLeft: `1px solid ${COLORS.border}`, paddingLeft: 16, marginLeft: 4 }}>Venta promedio</div>
            <div style={{ paddingLeft: 20 }}>Días de stock</div>
            <div>Objetivo (días)</div>
          </div>

          <div style={{
            display: "grid", gridTemplateColumns: `1.6fr repeat(${visibleMonthOrder.length}, 120px) 150px 190px 130px`,
            padding: "14px 16px", background: COLORS.primaryDark, color: "#fff",
            alignItems: "center", borderBottom: `2px solid ${COLORS.ink}`,
          }}>
            <div style={{ fontWeight: 700, fontSize: 13.5, letterSpacing: 0.3, textTransform: "uppercase" }}>
              {TOTAL_LABEL}
            </div>
            {globalPeriodTotals.map((val, i) => (
              hiddenMonths.has(monthOrder[i].key) ? null : (
                <TotalPeriodCell key={monthOrder[i].key} value={val} prevValue={i > 0 ? globalPeriodTotals[i - 1] : null} />
              )
            ))}
            <div style={{ textAlign: "right", fontFamily: "IBM Plex Mono, monospace", fontSize: 12.5, fontWeight: 700, borderLeft: "1px solid rgba(255,255,255,0.25)", paddingLeft: 16, marginLeft: 4 }}>
              {globalTotals.ventaTotal ? fmtMoney(globalTotals.ventaTotal) : "—"}
            </div>
            <div style={{ paddingLeft: 20, fontFamily: "IBM Plex Mono, monospace", fontSize: 13, fontWeight: 700 }}>
              {kpis.avgDias == null ? "—" : `${Math.round(kpis.avgDias)}d prom.`}
            </div>
            <div />
          </div>

          {grouped.map(({ grupo, rubros, statusCounts, periodTotals, ventaTotal, diasGrupo, targetGrupo }) => {
            const isOpen = statusFilter !== "TODOS" ? true : expanded.has(grupo);
            const ventaPctGrupo = ventaTotal && globalTotals.ventaTotal ? (ventaTotal / globalTotals.ventaTotal) * 100 : null;
            return (
              <div key={grupo}>
                <div
                  onClick={() => toggleExpand(grupo)}
                  style={{
                    display: "grid", gridTemplateColumns: `1.6fr repeat(${visibleMonthOrder.length}, 120px) 150px 190px 130px`,
                    padding: "12px 16px", cursor: "pointer", background: COLORS.primarySoft,
                    borderBottom: `1px solid ${COLORS.border}`, alignItems: "center",
                  }}
                >
                  <div style={{ display: "flex", alignItems: "center", gap: 6, fontWeight: 700, color: COLORS.primaryDark, fontSize: 13.5 }}>
                    {isOpen ? <ChevronDown size={15} /> : <ChevronRight size={15} />}
                    {grupo}
                  </div>
                  {periodTotals.map((val, i) => {
                    const periodKey = monthOrder[i].key;
                    if (hiddenMonths.has(periodKey)) return null;
                    const totalForPeriod = globalTotals.stockByPeriod[periodKey];
                    const pct = val != null && totalForPeriod ? (val / totalForPeriod) * 100 : null;
                    return (
                      <PeriodCell key={periodKey} value={val} prevValue={i > 0 ? periodTotals[i - 1] : null} pct={pct} />
                    );
                  })}
                  <div style={{ textAlign: "right", fontFamily: "IBM Plex Mono, monospace", fontSize: 12.5, color: COLORS.inkMuted, borderLeft: `1px solid ${COLORS.border}`, paddingLeft: 16, marginLeft: 4 }}>
                    {ventaTotal ? fmtMoney(ventaTotal) : "—"}
                    <PctTag pct={ventaPctGrupo} />
                  </div>
                  <div style={{ paddingLeft: 20 }}>
                    <Gauge dias={diasGrupo} min={targetGrupo.min} max={targetGrupo.max} />
                  </div>
                  <div />
                </div>

                {isOpen && rubros.map((r) => {
                  const t = r.target;
                  const venta = r.venta;
                  const dias = r.dias;
                  const ventaPctRubro = venta != null && globalTotals.ventaTotal ? (venta / globalTotals.ventaTotal) * 100 : null;
                  return (
                    <div
                      key={r.rkey}
                      style={{
                        display: "grid", gridTemplateColumns: `1.6fr repeat(${visibleMonthOrder.length}, 120px) 150px 190px 130px`,
                        padding: "10px 16px 10px 34px", borderBottom: `1px solid ${COLORS.border}`, alignItems: "center",
                      }}
                    >
                      <div style={{ display: "flex", alignItems: "center", gap: 8, fontSize: 13, color: COLORS.ink }}>
                        {titleCase(r.rubro)}
                        <StatusBadge dias={dias} min={t.min} max={t.max} />
                      </div>
                      {monthOrder.map((m, i) => {
                        if (hiddenMonths.has(m.key)) return null;
                        const val = r.months[m.key] ?? null;
                        const totalForPeriod = globalTotals.stockByPeriod[m.key];
                        const pct = val != null && totalForPeriod ? (val / totalForPeriod) * 100 : null;
                        return (
                          <PeriodCell
                            key={m.key}
                            value={val}
                            prevValue={i > 0 ? (r.months[monthOrder[i - 1].key] ?? null) : null}
                            pct={pct}
                          />
                        );
                      })}
                      <div style={{ textAlign: "right", fontFamily: "IBM Plex Mono, monospace", fontSize: 12.5, color: COLORS.inkMuted, borderLeft: `1px solid ${COLORS.border}`, paddingLeft: 16, marginLeft: 4 }}>
                        {venta != null ? fmtMoney(venta) : "—"}
                        <PctTag pct={ventaPctRubro} />
                      </div>
                      <div style={{ paddingLeft: 20 }}>
                        <Gauge dias={dias} min={t.min} max={t.max} />
                      </div>
                      <div style={{ display: "flex", gap: 4, alignItems: "center" }}>
                        <input
                          type="number" value={t.min}
                          onChange={(e) => updateTarget(r.rkey, "min", e.target.value)}
                          style={{ width: 42, padding: "3px 4px", fontSize: 11.5, borderRadius: 5, border: `1px solid ${COLORS.border}`, fontFamily: "IBM Plex Mono, monospace" }}
                        />
                        <span style={{ color: COLORS.inkMuted, fontSize: 11 }}>–</span>
                        <input
                          type="number" value={t.max}
                          onChange={(e) => updateTarget(r.rkey, "max", e.target.value)}
                          style={{ width: 42, padding: "3px 4px", fontSize: 11.5, borderRadius: 5, border: `1px solid ${COLORS.border}`, fontFamily: "IBM Plex Mono, monospace" }}
                        />
                      </div>
                    </div>
                  );
                })}
              </div>
            );
          })}
          {grouped.length === 0 && (
            <div style={{ padding: 40, textAlign: "center", color: COLORS.inkMuted, fontSize: 13 }}>
              No hay rubros que coincidan con el filtro.
            </div>
          )}
        </div>
      </div>
    </div>
  );
}

function EvolucionView({ data }) {
  const [compareLevel, setCompareLevel] = useState("rubro");
  const [selectedRubros, setSelectedRubros] = useState(new Set());
  const [selectedGrupos, setSelectedGrupos] = useState(new Set());
  const [showTotalEmpresa, setShowTotalEmpresa] = useState(false);
  const [metric, setMetric] = useState("stock");
  const [listSortBy, setListSortBy] = useState("stock");
  const getVenta = useVentaLookup(data);
  const monthOrder = useMonthOrder(data);
  const latest = monthOrder[monthOrder.length - 1];

  const grupos = useMemo(() => {
    const map = {};
    Object.entries(data.rows).forEach(([rkey, r]) => {
      const stockLatest = r.months[latest?.key] || 0;
      const venta = getVenta(rkey) || 0;
      if (!map[r.grupo]) map[r.grupo] = { rubros: [], stockTotal: 0, ventaTotal: 0 };
      map[r.grupo].rubros.push({ rkey, rubro: r.rubro, stockLatest, venta });
      map[r.grupo].stockTotal += stockLatest;
      map[r.grupo].ventaTotal += venta;
    });
    Object.values(map).forEach((g) => {
      if (listSortBy === "venta") g.rubros.sort((a, b) => b.venta - a.venta);
      else if (listSortBy === "nombre") g.rubros.sort((a, b) => a.rubro.localeCompare(b.rubro));
      else g.rubros.sort((a, b) => b.stockLatest - a.stockLatest);
    });
    let entries = Object.entries(map);
    if (listSortBy === "venta") entries.sort((a, b) => b[1].ventaTotal - a[1].ventaTotal);
    else if (listSortBy === "nombre") entries.sort((a, b) => a[0].localeCompare(b[0]));
    else entries.sort((a, b) => b[1].stockTotal - a[1].stockTotal);
    return entries;
  }, [data.rows, latest, getVenta, listSortBy]);

  const grupoAgg = useMemo(() => {
    const map = {};
    Object.entries(data.rows).forEach(([rkey, r]) => {
      if (!map[r.grupo]) map[r.grupo] = { months: {}, venta: 0 };
      Object.entries(r.months).forEach(([mk, val]) => {
        map[r.grupo].months[mk] = (map[r.grupo].months[mk] || 0) + (val || 0);
      });
      const venta = getVenta(rkey);
      if (venta) map[r.grupo].venta += venta;
    });
    return map;
  }, [data.rows, getVenta]);

  const totalAgg = useMemo(() => {
    const months = {};
    let venta = 0;
    Object.entries(data.rows).forEach(([rkey, r]) => {
      Object.entries(r.months).forEach(([mk, val]) => {
        months[mk] = (months[mk] || 0) + (val || 0);
      });
      const v = getVenta(rkey);
      if (v) venta += v;
    });
    return { months, venta };
  }, [data.rows, getVenta]);

  const grupoList = useMemo(() => {
    let entries = Object.keys(grupoAgg).map((g) => ({
      grupo: g,
      stockLatest: grupoAgg[g].months[latest?.key] || 0,
      ventaTotal: grupoAgg[g].venta,
    }));
    if (listSortBy === "venta") entries.sort((a, b) => b.ventaTotal - a.ventaTotal);
    else if (listSortBy === "nombre") entries.sort((a, b) => a.grupo.localeCompare(b.grupo));
    else entries.sort((a, b) => b.stockLatest - a.stockLatest);
    return entries;
  }, [grupoAgg, latest, listSortBy]);

  const toggleRubro = (key) => {
    setSelectedRubros((prevSel) => {
      const next = new Set(prevSel);
      next.has(key) ? next.delete(key) : next.add(key);
      return next;
    });
  };
  const toggleGrupo = (key) => {
    setSelectedGrupos((prevSel) => {
      const next = new Set(prevSel);
      next.has(key) ? next.delete(key) : next.add(key);
      return next;
    });
  };

  const selectedList = compareLevel === "grupo" ? [...selectedGrupos] : [...selectedRubros];

  const chartData = useMemo(() => {
    return monthOrder.map((m) => {
      const point = { mes: m.label };
      if (compareLevel === "grupo") {
        selectedList.forEach((g) => {
          const agg = grupoAgg[g];
          if (!agg) return;
          if (metric === "stock") {
            point[g] = agg.months[m.key] ?? null;
          } else {
            const stock = agg.months[m.key];
            const venta = agg.venta;
            point[g] = stock != null && venta ? Math.round(stock / (venta / 30)) : null;
          }
        });
      } else {
        selectedList.forEach((rkey) => {
          const row = data.rows[rkey];
          if (!row) return;
          const label = `${row.grupo} · ${titleCase(row.rubro)}`;
          if (metric === "stock") {
            point[label] = row.months[m.key] ?? null;
          } else {
            const stock = row.months[m.key];
            const venta = getVenta(rkey);
            point[label] = stock != null && venta ? Math.round(stock / (venta / 30)) : null;
          }
        });
      }
      if (showTotalEmpresa) {
        if (metric === "stock") {
          point[TOTAL_LABEL] = totalAgg.months[m.key] ?? null;
        } else {
          const stock = totalAgg.months[m.key];
          const venta = totalAgg.venta;
          point[TOTAL_LABEL] = stock != null && venta ? Math.round(stock / (venta / 30)) : null;
        }
      }
      return point;
    });
  }, [selectedList, data, metric, monthOrder, getVenta, compareLevel, grupoAgg, showTotalEmpresa, totalAgg]);

  const seriesLabels = [
    ...(compareLevel === "grupo"
      ? selectedList
      : selectedList.map((rkey) => {
          const row = data.rows[rkey];
          return `${row?.grupo} · ${titleCase(row?.rubro || "")}`;
        })),
    ...(showTotalEmpresa ? [TOTAL_LABEL] : []),
  ];

  const hasSelection = selectedList.length > 0 || showTotalEmpresa;

  return (
    <div style={{ display: "flex", gap: 20, alignItems: "flex-start" }}>
      <div style={{ width: 260, background: COLORS.surface, border: `1px solid ${COLORS.border}`, borderRadius: 12, padding: 14, maxHeight: 620, overflowY: "auto" }}>
        <div style={{ fontSize: 11.5, fontWeight: 700, color: COLORS.inkMuted, textTransform: "uppercase", letterSpacing: 0.4, marginBottom: 10 }}>
          Seleccioná qué comparar
        </div>
        <label style={{
          display: "flex", alignItems: "center", gap: 6, fontSize: 12.5, fontWeight: 700, color: COLORS.ink,
          padding: "7px 8px", marginBottom: 10, cursor: "pointer", background: COLORS.bg, borderRadius: 8,
        }}>
          <input type="checkbox" checked={showTotalEmpresa} onChange={() => setShowTotalEmpresa((v) => !v)} />
          {TOTAL_LABEL}
        </label>
        <div style={{ display: "flex", gap: 4, background: COLORS.bg, borderRadius: 8, padding: 3, marginBottom: 10 }}>
          {[{ id: "rubro", label: "Rubros" }, { id: "grupo", label: "Proveedores" }].map((opt) => (
            <button
              key={opt.id}
              onClick={() => setCompareLevel(opt.id)}
              style={{
                flex: 1, border: "none", padding: "6px 8px", borderRadius: 6, fontSize: 12, cursor: "pointer",
                background: compareLevel === opt.id ? COLORS.primary : "transparent",
                color: compareLevel === opt.id ? "#fff" : COLORS.inkMuted,
                fontWeight: compareLevel === opt.id ? 700 : 500,
              }}
            >
              {opt.label}
            </button>
          ))}
        </div>
        <div style={{ display: "flex", alignItems: "center", gap: 6, marginBottom: 12, flexWrap: "wrap" }}>
          <span style={{ fontSize: 11, color: COLORS.inkMuted }}>Ordenar por</span>
          <div style={{ display: "flex", gap: 3, background: COLORS.bg, borderRadius: 7, padding: 3 }}>
            {[{ id: "stock", label: "Stock" }, { id: "venta", label: "Venta" }, { id: "nombre", label: "Nombre" }].map((opt) => (
              <button
                key={opt.id}
                onClick={() => setListSortBy(opt.id)}
                style={{
                  border: "none", padding: "4px 8px", borderRadius: 5, fontSize: 11, cursor: "pointer",
                  background: listSortBy === opt.id ? COLORS.primary : "transparent",
                  color: listSortBy === opt.id ? "#fff" : COLORS.inkMuted,
                  fontWeight: listSortBy === opt.id ? 700 : 500,
                }}
              >
                {opt.label}
              </button>
            ))}
          </div>
        </div>
        {compareLevel === "grupo" ? (
          grupoList.map(({ grupo }) => (
            <label key={grupo} style={{ display: "flex", alignItems: "center", gap: 6, fontSize: 12.5, padding: "5px 0", cursor: "pointer", color: COLORS.ink, fontWeight: 600 }}>
              <input type="checkbox" checked={selectedGrupos.has(grupo)} onChange={() => toggleGrupo(grupo)} />
              {grupo}
            </label>
          ))
        ) : (
          grupos.map(([g, gData]) => (
            <div key={g} style={{ marginBottom: 12 }}>
              <div style={{ fontWeight: 700, fontSize: 12.5, color: COLORS.primaryDark, marginBottom: 4 }}>{g}</div>
              {gData.rubros.map(({ rkey, rubro }) => (
                <label key={rkey} style={{ display: "flex", alignItems: "center", gap: 6, fontSize: 12, padding: "3px 0", cursor: "pointer", color: COLORS.inkMuted }}>
                  <input type="checkbox" checked={selectedRubros.has(rkey)} onChange={() => toggleRubro(rkey)} />
                  {titleCase(rubro)}
                </label>
              ))}
            </div>
          ))
        )}
      </div>

      <div style={{ flex: 1, background: COLORS.surface, border: `1px solid ${COLORS.border}`, borderRadius: 12, padding: 18 }}>
        <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 14, flexWrap: "wrap", gap: 8 }}>
          <div style={{ fontFamily: "'Space Grotesk', sans-serif", fontWeight: 700, fontSize: 15, color: COLORS.ink }}>Evolución por período</div>
          <div style={{ display: "flex", gap: 6, background: COLORS.bg, borderRadius: 8, padding: 3 }}>
            {[{ id: "stock", label: "Stock $" }, { id: "dias", label: "Días de stock (est.)" }].map((opt) => (
              <button
                key={opt.id}
                onClick={() => setMetric(opt.id)}
                style={{
                  border: "none", padding: "6px 12px", borderRadius: 6, fontSize: 12, cursor: "pointer",
                  background: metric === opt.id ? COLORS.surface : "transparent",
                  color: metric === opt.id ? COLORS.primaryDark : COLORS.inkMuted,
                  fontWeight: metric === opt.id ? 700 : 500,
                  boxShadow: metric === opt.id ? "0 1px 2px rgba(0,0,0,0.08)" : "none",
                }}
              >
                {opt.label}
              </button>
            ))}
          </div>
        </div>

        {metric === "dias" && (
          <div style={{ display: "flex", gap: 6, alignItems: "flex-start", background: COLORS.warningSoft, color: "#8A5E1F", fontSize: 12, padding: "8px 10px", borderRadius: 8, marginBottom: 12 }}>
            <Info size={14} style={{ marginTop: 1, flexShrink: 0 }} />
            Estimado aplicando la venta promedio actual de cada rubro a los períodos históricos de stock.
          </div>
        )}

        {metric === "stock" && hasSelection && (
          <div style={{ display: "flex", gap: 6, alignItems: "flex-start", background: COLORS.primarySoft, color: COLORS.primaryDark, fontSize: 12, padding: "8px 10px", borderRadius: 8, marginBottom: 12 }}>
            <Info size={14} style={{ marginTop: 1, flexShrink: 0 }} />
            Las líneas punteadas horizontales muestran la venta promedio de cada serie, como referencia constante frente a la evolución del stock.
          </div>
        )}

        {!hasSelection ? (
          <div style={{ padding: 60, textAlign: "center", color: COLORS.inkMuted, fontSize: 13 }}>
            Elegí "{TOTAL_LABEL}" y/o uno o más rubros/proveedores en la lista de la izquierda para ver su evolución.
          </div>
        ) : (
          <ResponsiveContainer width="100%" height={420}>
            <LineChart data={chartData} margin={{ top: 4, right: 12, left: 0, bottom: 4 }}>
              <CartesianGrid stroke={COLORS.border} strokeDasharray="3 3" vertical={false} />
              <XAxis dataKey="mes" tick={{ fontSize: 11.5, fill: COLORS.inkMuted, fontFamily: "Inter" }} axisLine={{ stroke: COLORS.border }} tickLine={false} />
              <YAxis
                tick={{ fontSize: 11, fill: COLORS.inkMuted, fontFamily: "IBM Plex Mono" }}
                axisLine={false} tickLine={false}
                tickFormatter={(v) => metric === "stock" ? fmtMoneyShort(v) : `${v}d`}
                width={64}
              />
              <Tooltip
                formatter={(v) => metric === "stock" ? fmtMoney(v) : `${v}d`}
                contentStyle={{ borderRadius: 8, border: `1px solid ${COLORS.border}`, fontSize: 12, fontFamily: "Inter" }}
              />
              <Legend wrapperStyle={{ fontSize: 11.5, fontFamily: "Inter" }} />
              {seriesLabels.map((label, i) => {
                const isTotal = label === TOTAL_LABEL;
                return (
                  <Line
                    key={label}
                    type="monotone"
                    dataKey={label}
                    stroke={isTotal ? COLORS.ink : LINE_PALETTE[i % LINE_PALETTE.length]}
                    strokeWidth={isTotal ? 3.2 : 2.2}
                    dot={{ r: isTotal ? 4 : 3 }}
                    connectNulls
                  />
                );
              })}
              {metric === "stock" && selectedList.map((key, i) => {
                const venta = compareLevel === "grupo" ? grupoAgg[key]?.venta : getVenta(key);
                if (!venta) return null;
                const color = LINE_PALETTE[i % LINE_PALETTE.length];
                return (
                  <ReferenceLine
                    key={`venta-${key}`}
                    y={venta}
                    stroke={color}
                    strokeDasharray="4 4"
                    strokeWidth={1.3}
                    label={{ value: `venta prom. ${fmtMoneyShort(venta)}`, position: "insideTopRight", fontSize: 10, fill: color, fontFamily: "IBM Plex Mono, monospace" }}
                  />
                );
              })}
              {metric === "stock" && showTotalEmpresa && totalAgg.venta > 0 && (
                <ReferenceLine
                  y={totalAgg.venta}
                  stroke={COLORS.ink}
                  strokeDasharray="4 4"
                  strokeWidth={1.6}
                  label={{ value: `venta prom. total ${fmtMoneyShort(totalAgg.venta)}`, position: "insideTopRight", fontSize: 10, fill: COLORS.ink, fontFamily: "IBM Plex Mono, monospace" }}
                />
              )}
            </LineChart>
          </ResponsiveContainer>
        )}
      </div>
    </div>
  );
}

function UploadBox({ title, description, icon: Icon, onFile, busy, status, accentColor }) {
  return (
    <div
      onDragOver={(e) => e.preventDefault()}
      onDrop={(e) => { e.preventDefault(); onFile(e.dataTransfer.files?.[0]); }}
      style={{ border: `2px dashed ${COLORS.border}`, borderRadius: 14, padding: "26px 20px", textAlign: "center", background: COLORS.surface }}
    >
      <Icon size={24} color={accentColor} style={{ marginBottom: 8 }} />
      <div style={{ fontWeight: 700, color: COLORS.ink, marginBottom: 4, fontSize: 14 }}>{title}</div>
      <div style={{ fontSize: 12, color: COLORS.inkMuted, marginBottom: 12, lineHeight: 1.5 }}>{description}</div>
      <label style={{
        display: "inline-flex", alignItems: "center", gap: 8, background: accentColor, color: "#fff",
        padding: "8px 16px", borderRadius: 8, fontSize: 12.5, fontWeight: 600, cursor: "pointer",
      }}>
        <FileSpreadsheet size={14} />
        {busy ? "Procesando..." : "Elegir archivo"}
        <input
          type="file" accept=".xlsx,.xls" style={{ display: "none" }} disabled={busy}
          onChange={(e) => {
            const file = e.target.files?.[0];
            e.target.value = "";
            if (file) onFile(file);
          }}
        />
      </label>
      {status && (
        <div style={{
          marginTop: 12, padding: "8px 12px", borderRadius: 8, fontSize: 12, textAlign: "left",
          background: status.ok ? COLORS.successSoft : COLORS.criticalSoft,
          color: status.ok ? COLORS.success : COLORS.critical,
        }}>
          {status.msg}
        </div>
      )}
    </div>
  );
}

function CargarView({ data, onMergeStock, onApplyVenta, onResetAll, onLoadDemo }) {
  const [stockStatus, setStockStatus] = useState(null);
  const [stockBusy, setStockBusy] = useState(false);
  const [ventaStatus, setVentaStatus] = useState(null);
  const [ventaBusy, setVentaBusy] = useState(false);
  const [confirmReset, setConfirmReset] = useState(false);

  const handleStockFile = useCallback(async (file) => {
    if (!file) return;
    setStockBusy(true); setStockStatus(null);
    try {
      const buf = await file.arrayBuffer();
      const result = parseStockFile(buf, file.name);
      onMergeStock(result);
      setStockStatus({ ok: true, msg: `${Object.keys(result.rows).length} rubros · períodos: ${result.periods.map((p) => p.label).join(", ")}` });
    } catch (e) {
      setStockStatus({ ok: false, msg: e.message || "No se pudo procesar el archivo." });
    } finally { setStockBusy(false); }
  }, [onMergeStock]);

  const handleVentaFile = useCallback(async (file) => {
    if (!file) return;
    setVentaBusy(true); setVentaStatus(null);
    try {
      const buf = await file.arrayBuffer();
      const result = parseVentaFile(buf, file.name);
      onApplyVenta(result);
      if (result.hasFecha) {
        setVentaStatus({ ok: true, msg: `${result.rowCount} combinaciones de grupo/rubro · promediado sobre ${result.numMeses} mes(es) detectado(s) en el archivo. Reemplazó la venta promedio anterior.` });
      } else {
        setVentaStatus({ ok: false, msg: `No encontré una columna de fecha, así que usé el total sin promediar (puede estar sobreestimado si el archivo cubre varios meses). Revisá que el archivo tenga una columna tipo "Fecha_emision".` });
      }
    } catch (e) {
      setVentaStatus({ ok: false, msg: e.message || "No se pudo procesar el archivo." });
    } finally { setVentaBusy(false); }
  }, [onApplyVenta]);

  return (
    <div style={{ maxWidth: 900 }}>
      <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 16, marginBottom: 20 }}>
        <UploadBox
          title="Stock por período"
          description="Detalle de artículos con columnas Grupo, Rubro, Período y Neto ($). Cada período que traiga el archivo se suma al histórico."
          icon={Boxes}
          accentColor={COLORS.primary}
          onFile={handleStockFile}
          busy={stockBusy}
          status={stockStatus}
        />
        <UploadBox
          title="Venta promedio"
          description="Detalle de ventas con columnas Grupo, Rubro y Venta. Se agrupa por grupo + rubro y reemplaza por completo la venta promedio anterior."
          icon={Receipt}
          accentColor={COLORS.violet}
          onFile={handleVentaFile}
          busy={ventaBusy}
          status={ventaStatus}
        />
      </div>

      <div style={{ marginBottom: 20 }}>
        <button
          onClick={onLoadDemo}
          style={{ display: "inline-flex", alignItems: "center", gap: 6, background: "none", border: `1px solid ${COLORS.border}`, color: COLORS.primaryDark, padding: "7px 14px", borderRadius: 8, fontSize: 12.5, fontWeight: 600, cursor: "pointer" }}
        >
          <Zap size={13} /> {data.uploadLog.length > 0 ? "Regenerar datos de ejemplo" : "Probar con datos de ejemplo"}
        </button>
      </div>

      {data.ventaMeta && (
        <div style={{ background: COLORS.surface, border: `1px solid ${COLORS.border}`, borderRadius: 12, padding: 14, marginBottom: 16, display: "flex", justifyContent: "space-between", fontSize: 12.5 }}>
          <div><strong>Venta promedio activa:</strong> {data.ventaMeta.fileName}</div>
          <div style={{ color: COLORS.inkMuted }}>
            {data.ventaMeta.rowCount} filas · {data.ventaMeta.hasFecha ? `promedio sobre ${data.ventaMeta.numMeses} mes(es)` : "sin promediar (no había fecha)"} · cargado el {new Date(data.ventaMeta.uploadedAt).toLocaleDateString("es-AR")}
          </div>
        </div>
      )}

      <div style={{ background: COLORS.surface, border: `1px solid ${COLORS.border}`, borderRadius: 12, padding: 16 }}>
        <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 12 }}>
          <div style={{ fontWeight: 700, fontSize: 13.5, color: COLORS.ink }}>Historial de cargas</div>
          {data.uploadLog.length > 0 && (
            <button
              onClick={() => setConfirmReset(true)}
              style={{ display: "flex", alignItems: "center", gap: 5, background: "none", border: "none", color: COLORS.critical, fontSize: 12, cursor: "pointer" }}
            >
              <RotateCcw size={13} /> Borrar todos los datos
            </button>
          )}
        </div>
        <ConfirmModal
          open={confirmReset}
          title="Borrar todos los datos"
          message="¿Borrar todos los datos cargados (stock y venta)? Esta acción no se puede deshacer."
          confirmLabel="Borrar todo"
          onCancel={() => setConfirmReset(false)}
          onConfirm={() => { onResetAll(); setConfirmReset(false); }}
        />
        {data.uploadLog.length === 0 ? (
          <div style={{ fontSize: 12.5, color: COLORS.inkMuted }}>Todavía no cargaste ningún archivo.</div>
        ) : (
          <div>
            {[...data.uploadLog].reverse().map((p) => (
              <div key={p.id} style={{ display: "flex", justifyContent: "space-between", padding: "8px 0", borderBottom: `1px solid ${COLORS.border}`, fontSize: 12.5, gap: 10 }}>
                <span style={{
                  background: p.type === "stock" ? COLORS.primarySoft : "#EFE9F5",
                  color: p.type === "stock" ? COLORS.primaryDark : COLORS.violet,
                  fontWeight: 700, fontSize: 10.5, padding: "2px 8px", borderRadius: 20, textTransform: "uppercase", flexShrink: 0,
                }}>
                  {p.type === "stock" ? "Stock" : "Venta"}
                </span>
                <div style={{ color: COLORS.ink, fontWeight: 500, flex: 1 }}>{p.fileName}</div>
                <div style={{ color: COLORS.inkMuted }}>{p.rowCount} filas{p.periods ? ` · ${p.periods.join(", ")}` : ""}</div>
                <div style={{ color: COLORS.inkMuted, fontFamily: "IBM Plex Mono, monospace" }}>{new Date(p.uploadedAt).toLocaleDateString("es-AR")}</div>
              </div>
            ))}
          </div>
        )}
      </div>
    </div>
  );
}

/* ---------------------------------- app ---------------------------------- */

class ErrorBoundary extends Component {
  constructor(props) {
    super(props);
    this.state = { error: null };
  }
  static getDerivedStateFromError(error) {
    return { error };
  }
  componentDidCatch(error, info) {
    console.error("Control de Stock crashed:", error, info);
  }
  render() {
    if (this.state.error) {
      return (
        <div style={{ background: COLORS.bg, minHeight: "100%", fontFamily: "Inter, sans-serif", padding: "60px 24px", textAlign: "center" }}>
          <AlertTriangle size={28} color={COLORS.critical} style={{ marginBottom: 10 }} />
          <div style={{ fontWeight: 700, color: COLORS.ink, marginBottom: 6, fontSize: 15 }}>Algo falló al cargar Control de Stock</div>
          <div style={{ fontSize: 12.5, color: COLORS.inkMuted, marginBottom: 16, maxWidth: 480, marginLeft: "auto", marginRight: "auto" }}>
            {this.state.error?.message || "Error desconocido"}
          </div>
          <button
            onClick={() => this.setState({ error: null })}
            style={{ display: "inline-flex", alignItems: "center", gap: 6, background: COLORS.primary, color: "#fff", border: "none", padding: "9px 18px", borderRadius: 8, fontSize: 13, fontWeight: 600, cursor: "pointer" }}
          >
            <RotateCcw size={14} /> Reintentar
          </button>
        </div>
      );
    }
    return this.props.children;
  }
}

function AppInner() {
  const [data, setData] = useState(null);
  const [loading, setLoading] = useState(true);
  const [loadError, setLoadError] = useState(false);
  const [saving, setSaving] = useState(false);
  const [view, setView] = useState("dashboard");
  const [expanded, setExpanded] = useState(new Set());
  const saveTimerRef = useRef(null);

  const bootstrap = useCallback(async () => {
    setLoading(true);
    setLoadError(false);
    const result = await loadData();
    if (!result.ok) {
      setLoadError(true);
      setLoading(false);
      return;
    }
    const d = result.data;
    setData(d);
    const firstGrupo = [...new Set(Object.values(d.rows).map((r) => r.grupo))][0];
    if (firstGrupo) setExpanded(new Set([firstGrupo]));
    setLoading(false);
  }, []);

  useEffect(() => { bootstrap(); }, [bootstrap]);
  useEffect(() => () => { if (saveTimerRef.current) clearTimeout(saveTimerRef.current); }, []);

  const persist = useCallback((next) => {
    setData(next);
    setSaving(true);
    saveData(next).finally(() => setSaving(false));
  }, []);

  const persistDebounced = useCallback((next) => {
    setData(next);
    if (saveTimerRef.current) clearTimeout(saveTimerRef.current);
    setSaving(true);
    saveTimerRef.current = setTimeout(() => {
      saveData(next).finally(() => setSaving(false));
    }, 700);
  }, []);

  const handleMergeStock = useCallback((parseResult) => {
    setData((prevData) => {
      const merged = mergeStockData(prevData, parseResult);
      setSaving(true);
      saveData(merged).finally(() => setSaving(false));
      const newGrupos = [...new Set(Object.values(parseResult.rows).map((r) => r.grupo))];
      setExpanded((exp) => new Set([...exp, ...newGrupos]));
      return merged;
    });
  }, []);

  const handleApplyVenta = useCallback((parseResult) => {
    setData((prevData) => {
      const next = applyVentaData(prevData, parseResult);
      setSaving(true);
      saveData(next).finally(() => setSaving(false));
      return next;
    });
  }, []);

  const setTargets = useCallback((updater) => {
    setData((prevData) => {
      const nextTargets = typeof updater === "function" ? updater(prevData.targets) : updater;
      const next = { ...prevData, targets: nextTargets };
      persistDebounced(next);
      return next;
    });
  }, [persistDebounced]);

  const handleResetAll = useCallback(() => {
    const empty = emptyData();
    persist(empty);
    setExpanded(new Set());
    setView("cargar");
  }, [persist]);

  const handleLoadDemo = useCallback(() => {
    const demo = generateDemoData();
    persist(demo);
    setExpanded(new Set([DEMO_GRUPOS[0].grupo]));
    setView("dashboard");
  }, [persist]);

  const toggleExpand = (grupo) => {
    setExpanded((prevExp) => {
      const next = new Set(prevExp);
      next.has(grupo) ? next.delete(grupo) : next.add(grupo);
      return next;
    });
  };

  const hasData = data && Object.keys(data.rows).length > 0;

  const TABS = [
    { id: "dashboard", label: "Dashboard", icon: Layers },
    { id: "evolucion", label: "Evolución", icon: LineChartIcon },
    { id: "cargar", label: "Cargar datos", icon: Upload },
  ];

  return (
    <div style={{ background: COLORS.bg, minHeight: "100%", fontFamily: "Inter, sans-serif", color: COLORS.ink, padding: "20px 24px 60px" }}>
      <link rel="stylesheet" href={FONT_LINK} />

      <div style={{ display: "flex", justifyContent: "space-between", alignItems: "flex-end", flexWrap: "wrap", gap: 14, marginBottom: 20 }}>
        <div>
          <div style={{ fontFamily: "'Space Grotesk', sans-serif", fontWeight: 700, fontSize: 22, letterSpacing: -0.5 }}>Control de Stock</div>
          <div style={{ fontSize: 12.5, color: COLORS.inkMuted, marginTop: 2, display: "flex", alignItems: "center", gap: 6 }}>
            Evolución de inventario por proveedor y rubro, y días de stock vs. objetivo
            {saving && <span style={{ color: COLORS.primary, fontWeight: 600 }}>· guardando...</span>}
          </div>
        </div>
        <div style={{ display: "flex", gap: 4, background: COLORS.surface, border: `1px solid ${COLORS.border}`, borderRadius: 10, padding: 4 }}>
          {TABS.map((t) => (
            <button
              key={t.id}
              onClick={() => setView(t.id)}
              style={{
                display: "flex", alignItems: "center", gap: 6, border: "none", cursor: "pointer",
                padding: "8px 14px", borderRadius: 7, fontSize: 12.5, fontWeight: 600,
                background: view === t.id ? COLORS.primary : "transparent",
                color: view === t.id ? "#fff" : COLORS.inkMuted,
              }}
            >
              <t.icon size={14} /> {t.label}
            </button>
          ))}
        </div>
      </div>

      {loading ? (
        <div style={{ padding: 60, textAlign: "center", color: COLORS.inkMuted, fontSize: 13 }}>
          <div style={{
            width: 22, height: 22, margin: "0 auto 12px", borderRadius: "50%",
            border: `2.5px solid ${COLORS.border}`, borderTopColor: COLORS.primary,
            animation: "cs-spin 0.8s linear infinite",
          }} />
          <style>{"@keyframes cs-spin { to { transform: rotate(360deg); } }"}</style>
          Cargando datos guardados...
        </div>
      ) : loadError ? (
        <div style={{ background: COLORS.surface, border: `1px dashed ${COLORS.critical}`, borderRadius: 12, padding: 50, textAlign: "center" }}>
          <AlertTriangle size={28} color={COLORS.critical} style={{ marginBottom: 10 }} />
          <div style={{ fontWeight: 700, marginBottom: 4 }}>No se pudo cargar la información guardada</div>
          <div style={{ fontSize: 12.5, color: COLORS.inkMuted, marginBottom: 16 }}>
            Puede ser un problema de conexión momentáneo. Probá de nuevo — tus datos siguen guardados, no se perdieron.
          </div>
          <button onClick={bootstrap} style={{ display: "inline-flex", alignItems: "center", gap: 6, background: COLORS.primary, color: "#fff", border: "none", padding: "9px 18px", borderRadius: 8, fontSize: 13, fontWeight: 600, cursor: "pointer" }}>
            <RotateCcw size={14} /> Reintentar
          </button>
        </div>
      ) : !hasData && view !== "cargar" ? (
        <div style={{ background: COLORS.surface, border: `1px dashed ${COLORS.border}`, borderRadius: 12, padding: 50, textAlign: "center" }}>
          <FileSpreadsheet size={28} color={COLORS.inkMuted} style={{ marginBottom: 10 }} />
          <div style={{ fontWeight: 700, marginBottom: 4 }}>Todavía no hay datos cargados</div>
          <div style={{ fontSize: 12.5, color: COLORS.inkMuted, marginBottom: 16 }}>Subí tus archivos de stock y venta para empezar a ver el dashboard.</div>
          <div style={{ display: "flex", gap: 10, justifyContent: "center" }}>
            <button onClick={() => setView("cargar")} style={{ background: COLORS.primary, color: "#fff", border: "none", padding: "9px 18px", borderRadius: 8, fontSize: 13, fontWeight: 600, cursor: "pointer" }}>
              Ir a cargar datos
            </button>
            <button onClick={handleLoadDemo} style={{ display: "flex", alignItems: "center", gap: 6, background: "none", border: `1px solid ${COLORS.border}`, color: COLORS.primaryDark, padding: "9px 18px", borderRadius: 8, fontSize: 13, fontWeight: 600, cursor: "pointer" }}>
              <Zap size={14} /> Probar con datos de ejemplo
            </button>
          </div>
        </div>
      ) : (
        <>
          {view === "dashboard" && <DashboardView data={data} targets={data.targets} setTargets={setTargets} expanded={expanded} toggleExpand={toggleExpand} setExpanded={setExpanded} />}
          {view === "evolucion" && <EvolucionView data={data} />}
          {view === "cargar" && <CargarView data={data} onMergeStock={handleMergeStock} onApplyVenta={handleApplyVenta} onResetAll={handleResetAll} onLoadDemo={handleLoadDemo} />}
        </>
      )}
    </div>
  );
}

export default function App() {
  return (
    <ErrorBoundary>
      <AppInner />
    </ErrorBoundary>
  );
}
