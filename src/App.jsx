import { useState, useEffect } from "react";
import * as XLSX from "xlsx";
import {
  LineChart, Line, XAxis, YAxis, CartesianGrid, Tooltip,
  Legend, ResponsiveContainer, ReferenceLine,
} from "recharts";

// ── CONFIG ─────────────────────────────────────────────────────────────────────
const FILE_ID    = "1ynIpgfPGAr5F6uQ-t0HkXZNShrzPngzd";
const SHEET_NAME = "Quality Status (HQ)";
// Proxy để bypass CORS khi deploy trên Vercel
const EXPORT_URL = "https://docs.google.com/spreadsheets/d/" + FILE_ID + "/export?format=xlsx";
const GDRIVE_URL = "https://api.allorigins.win/raw?url=" + encodeURIComponent(EXPORT_URL);

// ── THEME ──────────────────────────────────────────────────────────────────────
const C = {
  navy:    "#0D2B55",
  blue:    "#1565C0",
  accent:  "#E53935",
  gold:    "#F9A825",
  ok:      "#2E7D32",
  okBg:    "#E8F5E9",
  warnBg:  "#FFF8E1",
  warn:    "#E65100",
  bg:      "#EEF2F7",
  surface: "#FFFFFF",
  border:  "#D1DBE8",
  text:    "#1A202C",
  muted:   "#64748B",
  rowAlt:  "#F0F5FF",
  lineA:   "#1565C0",
  lineB:   "#E53935",
};

// ── HELPERS ────────────────────────────────────────────────────────────────────
const num = (v) => {
  if (v == null || v === "" || v === "-") return 0;
  const n = parseFloat(String(v).replace(/[%,\s]/g, ""));
  return isNaN(n) ? 0 : n;
};

const MONTHS_SHORT = ["Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec"];

const excelDateToLabel = (serial) => {
  if (!serial || isNaN(Number(serial))) return String(serial ?? "");
  const d = new Date(Math.round((Number(serial) - 25569) * 86400 * 1000));
  return `${MONTHS_SHORT[d.getUTCMonth()]}-${String(d.getUTCFullYear()).slice(2)}`;
};

// ── PARSER ─────────────────────────────────────────────────────────────────────
function parseSheet(wb) {
  const ws = wb.Sheets[SHEET_NAME];
  if (!ws) {
    const names = wb.SheetNames.join(", ");
    throw new Error(`Sheet "${SHEET_NAME}" not found. Available: ${names}`);
  }
  const raw = XLSX.utils.sheet_to_json(ws, { header: 1, defval: null });

  // ── Find "NCR Rate" header row ──
  let ncrRateRow = -1;
  for (let i = 0; i < Math.min(raw.length, 40); i++) {
    if (String(raw[i]?.[0] ?? "").trim() === "NCR Rate") { ncrRateRow = i; break; }
  }
  if (ncrRateRow === -1) throw new Error("Cannot locate 'NCR Rate' row in sheet.");

  const headerRow = raw[ncrRateRow];      // col labels
  const inputRow  = raw[ncrRateRow + 1];  // Input quantity(ea)
  const ncrRow    = raw[ncrRateRow + 2];  // NCR(ea)
  const defRow    = raw[ncrRateRow + 3];  // Defect Rate(%)
  const cumRow    = raw[ncrRateRow + 4];  // Cumulative defect rate(%)

  // ── Build month labels (cols 2–13 → Jan-Dec) ──
  const monthLabels = [];
  for (let c = 2; c <= 13; c++) {
    const v = headerRow[c];
    monthLabels.push(typeof v === "number" && v > 40000 ? excelDateToLabel(v) : String(v ?? ""));
  }
  // col 1 = 2025 baseline, col 14 = YTD Total
  // cols 15-19 = FW weeks, col 20 = 5W Total

  const weekLabels = [];
  for (let c = 15; c <= 19; c++) {
    weekLabels.push(String(headerRow[c] ?? `FW${c - 14 + 15}`));
  }

  // ── KPI values from top of sheet ──
  let target = 0.0015, cumDefect = 0;
  for (let i = 0; i < Math.min(raw.length, 15); i++) {
    const a = String(raw[i]?.[0] ?? "").trim().toLowerCase();
    if (a.includes("y2026 target") && i < 5) {
      const v = raw[i + 1]?.[0] ?? raw[i + 2]?.[0];
      if (v != null) target = num(v);
    }
    if (a.includes("cumulative") && a.includes("defect rate")) {
      const v = raw[i + 1]?.[0] ?? raw[i + 2]?.[0];
      if (v != null) cumDefect = num(v);
    }
  }
  // fallback from summary table
  if (!cumDefect) cumDefect = num(cumRow?.[14]);
  const targetPct   = target * 100;
  const cumDefectPct = cumDefect * 100;

  const totalInput = num(inputRow?.[14]);
  const totalNcr   = num(ncrRow?.[14]);

  // ── Monthly series ──
  // col 1 = 2025, col 2..13 = Jan-Dec
  const allMonthCols = [
    { label: "2025", col: 1 },
    ...monthLabels.map((label, i) => ({ label, col: i + 2 })),
  ];
  const monthly = allMonthCols.map(({ label, col }) => {
    const inp = num(inputRow?.[col]);
    const ncr = num(ncrRow?.[col]);
    const def = num(defRow?.[col]) * 100;
    const cum = num(cumRow?.[col]) * 100;
    return {
      month: label,
      input: inp,
      ncr,
      defect:     inp > 0 ? parseFloat(def.toFixed(4)) : null,
      cumulative: inp > 0 ? parseFloat(cum.toFixed(4)) : null,
    };
  });

  // ── Weekly series ──
  const weekly = weekLabels.map((label, i) => {
    const col = 15 + i;
    return {
      week:   label,
      input:  num(inputRow?.[col]),
      ncr:    num(ncrRow?.[col]),
      defect: parseFloat((num(defRow?.[col]) * 100).toFixed(4)),
    };
  });

  // ── Process rows ──
  const PROCESSES = ["Turning","Boring","Gear Cutting","Induction","Drilling","Hard Turning","Assembly","Coating","Total"];
  const processRows = [];
  for (let i = ncrRateRow + 6; i < Math.min(raw.length, ncrRateRow + 30); i++) {
    const a = String(raw[i]?.[0] ?? "").trim();
    if (!PROCESSES.includes(a)) continue;
    processRows.push({
      process:  a,
      ncr2025:  num(raw[i]?.[1]),
      months:   monthLabels.map((_, mi) => num(raw[i]?.[mi + 2])),
      total:    num(raw[i]?.[14]),
      weeks:    weekLabels.map((_, wi) => num(raw[i]?.[15 + wi])),
      weekTotal: num(raw[i]?.[20]),
      isTotal:  a === "Total",
    });
    if (a === "Total") break;
  }

  return {
    targetPct,
    cumDefectPct,
    totalInput,
    totalNcr,
    monthly,
    weekly,
    monthLabels,
    weekLabels,
    processRows,
  };
}

// ── SHARED STYLES ──────────────────────────────────────────────────────────────
const thStyle = (extra = {}) => ({
  background: C.navy, color: "#fff",
  padding: "6px 10px", fontSize: 11,
  fontWeight: 700, textAlign: "center",
  whiteSpace: "nowrap", ...extra,
});

// ── SUB-COMPONENTS ─────────────────────────────────────────────────────────────
const CustomTooltip = ({ active, payload, label }) => {
  if (!active || !payload?.length) return null;
  return (
    <div style={{ background: C.navy, color: "#fff", padding: "8px 12px", borderRadius: 6, fontSize: 12 }}>
      <div style={{ fontWeight: 700, marginBottom: 4 }}>{label}</div>
      {payload.map((p) => (
        <div key={p.name}>
          <span style={{ color: p.color }}>■ </span>
          {p.name}: {p.value != null ? p.value.toFixed(2) + "%" : "—"}
        </div>
      ))}
    </div>
  );
};

const KpiCard = ({ label, value, sub, highlight, warn }) => (
  <div style={{
    background: highlight ? C.navy : C.surface,
    border: `1.5px solid ${highlight ? C.navy : warn ? C.accent : C.border}`,
    borderRadius: 10, padding: "14px 18px",
    flex: "1 1 130px", minWidth: 130,
  }}>
    <div style={{ fontSize: 10, fontWeight: 700, letterSpacing: "0.08em", textTransform: "uppercase", color: highlight ? "#90CAF9" : C.muted, marginBottom: 6 }}>
      {label}
    </div>
    <div style={{ fontSize: 26, fontWeight: 800, lineHeight: 1, color: highlight ? "#FFF" : warn ? C.accent : C.navy }}>
      {value}
    </div>
    {sub && <div style={{ fontSize: 11, color: highlight ? "#90CAF9" : C.muted, marginTop: 4 }}>{sub}</div>}
  </div>
);

const SectionHeader = ({ title, sub }) => (
  <div style={{ borderLeft: `4px solid ${C.blue}`, paddingLeft: 10, marginBottom: 14 }}>
    <div style={{ fontSize: 12, fontWeight: 700, color: C.navy, textTransform: "uppercase", letterSpacing: "0.07em" }}>{title}</div>
    {sub && <div style={{ fontSize: 11, color: C.muted, marginTop: 1 }}>{sub}</div>}
  </div>
);

const Spinner = () => (
  <>
    <style>{`@keyframes spin{to{transform:rotate(360deg)}}`}</style>
    <div style={{ width: 40, height: 40, border: `4px solid ${C.border}`, borderTopColor: C.blue, borderRadius: "50%", animation: "spin 0.8s linear infinite" }} />
  </>
);

// ── MAIN COMPONENT ─────────────────────────────────────────────────────────────
export default function App() {
  const [data, setData]         = useState(null);
  const [error, setError]       = useState(null);
  const [loading, setLoading]   = useState(true);
  const [tab, setTab]           = useState("monthly");
  const [lastUpdated, setLastUpdated] = useState(null);

  const loadData = async () => {
    setLoading(true);
    setError(null);
    try {
      const res = await fetch(GDRIVE_URL);
      if (!res.ok) throw new Error(`HTTP ${res.status} — kiểm tra file đã được share public chưa.`);
      const buf = await res.arrayBuffer();
      const wb  = XLSX.read(buf, { type: "array" });
      setData(parseSheet(wb));
      setLastUpdated(new Date().toLocaleString("vi-VN"));
    } catch (e) {
      setError(e.message);
    } finally {
      setLoading(false);
    }
  };

  useEffect(() => { loadData(); }, []);

  // ── LOADING ──
  if (loading) return (
    <div style={{ display: "flex", flexDirection: "column", alignItems: "center", justifyContent: "center", height: "100vh", background: C.bg, gap: 16 }}>
      <Spinner />
      <div style={{ color: C.muted, fontSize: 14 }}>Đang tải dữ liệu từ Google Drive…</div>
    </div>
  );

  // ── ERROR ──
  if (error) return (
    <div style={{ display: "flex", flexDirection: "column", alignItems: "center", justifyContent: "center", height: "100vh", background: C.bg, gap: 12, padding: 24, textAlign: "center" }}>
      <div style={{ fontSize: 36 }}>⚠️</div>
      <div style={{ fontSize: 16, fontWeight: 700, color: C.accent }}>Không thể tải file</div>
      <div style={{ fontSize: 13, color: C.muted, maxWidth: 420 }}>{error}</div>
      <div style={{ fontSize: 12, color: C.muted, background: C.surface, padding: "10px 18px", borderRadius: 8, maxWidth: 460, lineHeight: 1.6 }}>
        Đảm bảo file Google Drive đã được set <strong>"Anyone with the link → Viewer"</strong>
      </div>
      <button onClick={loadData} style={{ marginTop: 8, padding: "8px 22px", background: C.blue, color: "#fff", border: "none", borderRadius: 6, cursor: "pointer", fontWeight: 700, fontSize: 13 }}>
        Thử lại
      </button>
    </div>
  );

  const { targetPct, cumDefectPct, totalInput, totalNcr, monthly, weekly, monthLabels, weekLabels, processRows } = data;
  const isOverTarget = cumDefectPct > targetPct;

  const weeklyInputTotal = weekly.reduce((s, w) => s + w.input, 0);
  const weeklyNcrTotal   = weekly.reduce((s, w) => s + w.ncr,   0);

  const chartMonthly = monthly.map(d => ({
    name: d.month,
    "Defect Rate(%)": d.defect,
    "Cumulative(%)":  d.cumulative,
  }));

  const chartWeekly = weekly.map(d => ({
    name: d.week,
    "Defect Rate(%)": d.defect,
  }));

  return (
    <div style={{ fontFamily: "'Inter','Segoe UI',sans-serif", background: C.bg, minHeight: "100vh", padding: "18px 22px", color: C.text }}>

      {/* ── HEADER ── */}
      <div style={{ display: "flex", alignItems: "flex-start", justifyContent: "space-between", marginBottom: 18, flexWrap: "wrap", gap: 10 }}>
        <div>
          <div style={{ fontSize: 10, fontWeight: 700, letterSpacing: "0.1em", textTransform: "uppercase", color: C.blue, marginBottom: 3 }}>
            Internal Process Quality
          </div>
          <h1 style={{ margin: 0, fontSize: 20, fontWeight: 800, color: C.navy, letterSpacing: "-0.02em" }}>
            Y2026 NCR Status Dashboard
          </h1>
          {lastUpdated && (
            <div style={{ fontSize: 11, color: C.muted, marginTop: 4 }}>
              Cập nhật lúc {lastUpdated} ·{" "}
              <span onClick={loadData} style={{ color: C.blue, cursor: "pointer", textDecoration: "underline" }}>
                Refresh
              </span>
            </div>
          )}
        </div>
        <div style={{
          background: isOverTarget ? C.warnBg : C.okBg,
          border: `1.5px solid ${isOverTarget ? C.warn : C.ok}`,
          borderRadius: 8, padding: "8px 18px", textAlign: "center", minWidth: 120,
        }}>
          <div style={{ fontSize: 10, fontWeight: 700, color: C.muted, textTransform: "uppercase" }}>Y2026 Target</div>
          <div style={{ fontSize: 24, fontWeight: 800, color: isOverTarget ? C.warn : C.ok }}>
            {targetPct.toFixed(2)}%
          </div>
          <div style={{ fontSize: 10, color: isOverTarget ? C.warn : C.ok, fontWeight: 600 }}>
            {isOverTarget ? "⚠ Above target" : "✓ On track"}
          </div>
        </div>
      </div>

      {/* ── KPI ROW ── */}
      <div style={{ display: "flex", gap: 10, flexWrap: "wrap", marginBottom: 18 }}>
        <KpiCard label="Cumulative Defect Rate" value={cumDefectPct.toFixed(2) + "%"} highlight
          sub={`Target: ${targetPct.toFixed(2)}%`} />
        <KpiCard label="YTD Input Qty" value={totalInput.toLocaleString()} sub="ea · Jan–May 2026" />
        <KpiCard label="YTD NCR" value={totalNcr} warn={totalNcr > 0} sub="ea · Jan–May 2026" />
        <KpiCard label="5-Week Input" value={weeklyInputTotal.toLocaleString()}
          sub={`ea · ${weekLabels[0]}–${weekLabels[weekLabels.length - 1]}`} />
        <KpiCard label="5-Week NCR" value={weeklyNcrTotal} warn={weeklyNcrTotal > 0}
          sub={`ea · ${weekLabels[0]}–${weekLabels[weekLabels.length - 1]}`} />
      </div>

      {/* ── CHARTS ── */}
      <div style={{ display: "grid", gridTemplateColumns: "3fr 2fr", gap: 14, marginBottom: 18 }}>

        {/* Monthly chart */}
        <div style={{ background: C.surface, borderRadius: 10, padding: 16, border: `1px solid ${C.border}` }}>
          <SectionHeader title="Monthly NCR Status" />
          <ResponsiveContainer width="100%" height={190}>
            <LineChart data={chartMonthly} margin={{ top: 4, right: 12, left: -14, bottom: 0 }}>
              <CartesianGrid strokeDasharray="3 3" stroke="#E2E8F0" />
              <XAxis dataKey="name" tick={{ fontSize: 10, fill: C.muted }} />
              <YAxis tickFormatter={v => v + "%"} tick={{ fontSize: 10, fill: C.muted }} domain={[0, "auto"]} />
              <Tooltip content={<CustomTooltip />} />
              <Legend iconType="circle" wrapperStyle={{ fontSize: 11 }} />
              <ReferenceLine y={targetPct} stroke={C.gold} strokeDasharray="4 4"
                label={{ value: `Target ${targetPct}%`, fontSize: 9, fill: C.gold, position: "insideTopLeft" }} />
              <Line type="monotone" dataKey="Defect Rate(%)" stroke={C.lineA} strokeWidth={2} dot={{ r: 3 }} connectNulls={false} />
              <Line type="monotone" dataKey="Cumulative(%)" stroke={C.lineB} strokeWidth={2} dot={{ r: 3 }} strokeDasharray="5 3" connectNulls={false} />
            </LineChart>
          </ResponsiveContainer>
        </div>

        {/* Weekly chart + mini table */}
        <div style={{ background: C.surface, borderRadius: 10, padding: 16, border: `1px solid ${C.border}` }}>
          <SectionHeader title="Weekly NCR Status" sub="Last 5 weeks" />
          <ResponsiveContainer width="100%" height={130}>
            <LineChart data={chartWeekly} margin={{ top: 4, right: 12, left: -14, bottom: 0 }}>
              <CartesianGrid strokeDasharray="3 3" stroke="#E2E8F0" />
              <XAxis dataKey="name" tick={{ fontSize: 10, fill: C.muted }} />
              <YAxis tickFormatter={v => v + "%"} tick={{ fontSize: 10, fill: C.muted }} domain={[0, "auto"]} />
              <Tooltip content={<CustomTooltip />} />
              <ReferenceLine y={targetPct} stroke={C.gold} strokeDasharray="4 4" />
              <Line type="monotone" dataKey="Defect Rate(%)" stroke={C.lineA} strokeWidth={2} dot={{ r: 4 }} />
            </LineChart>
          </ResponsiveContainer>
          <div style={{ marginTop: 10, overflowX: "auto" }}>
            <table style={{ width: "100%", borderCollapse: "collapse", fontSize: 11 }}>
              <thead>
                <tr style={{ background: C.navy }}>
                  <th style={thStyle({ textAlign: "left" })}></th>
                  {weekLabels.map(w => <th key={w} style={thStyle()}>{w}</th>)}
                  <th style={thStyle()}>Total</th>
                </tr>
              </thead>
              <tbody>
                {[
                  { label: "Input", vals: weekly.map(w => w.input), total: weeklyInputTotal, fmt: v => v.toLocaleString() },
                  { label: "NCR",   vals: weekly.map(w => w.ncr),   total: weeklyNcrTotal,  fmt: v => v },
                  { label: "Rate",  vals: weekly.map(w => w.defect),total: null, fmt: v => v.toFixed(2) + "%" },
                ].map((row, i) => (
                  <tr key={row.label} style={{ background: i % 2 === 0 ? C.surface : C.rowAlt }}>
                    <td style={{ padding: "4px 8px", fontWeight: 700, color: C.navy }}>{row.label}</td>
                    {row.vals.map((v, j) => (
                      <td key={j} style={{
                        textAlign: "center", padding: "4px 6px",
                        color: row.label === "Rate" && v > targetPct ? C.accent : C.text,
                        fontWeight: row.label === "Rate" && v > targetPct ? 700 : 400,
                      }}>
                        {row.fmt(v)}
                      </td>
                    ))}
                    <td style={{ textAlign: "center", padding: "4px 6px", fontWeight: 700, color: C.navy }}>
                      {row.total != null ? row.total.toLocaleString() : "—"}
                    </td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        </div>
      </div>

      {/* ── MONTHLY RATE TABLE ── */}
      <div style={{ background: C.surface, borderRadius: 10, padding: 16, border: `1px solid ${C.border}`, marginBottom: 18 }}>
        <SectionHeader title="NCR Rate Summary — Monthly" />
        <div style={{ overflowX: "auto" }}>
          <table style={{ width: "100%", borderCollapse: "collapse", fontSize: 11 }}>
            <thead>
              <tr style={{ background: C.navy }}>
                <th style={thStyle({ textAlign: "left", minWidth: 130 })}>Metric</th>
                <th style={thStyle()}>2025</th>
                {monthLabels.map(m => <th key={m} style={thStyle()}>{m}</th>)}
                <th style={thStyle()}>Total Y26</th>
              </tr>
            </thead>
            <tbody>
              {[
                { label: "Input (ea)",     vals: monthly.map(m => m.input),      total: totalInput,    fmt: v => v != null ? v.toLocaleString() : "—",        color: false },
                { label: "NCR (ea)",       vals: monthly.map(m => m.ncr),        total: totalNcr,      fmt: v => v != null ? v : "—",                          color: false },
                { label: "Defect Rate(%)", vals: monthly.map(m => m.defect),     total: null,          fmt: v => v != null ? v.toFixed(2) + "%" : "—",         color: true  },
                { label: "Cumulative(%)",  vals: monthly.map(m => m.cumulative), total: cumDefectPct,  fmt: v => v != null ? v.toFixed(2) + "%" : "—",         color: true  },
              ].map((row, i) => (
                <tr key={row.label} style={{ background: i % 2 === 0 ? C.surface : C.rowAlt }}>
                  <td style={{ padding: "5px 10px", fontWeight: 700, color: C.navy }}>{row.label}</td>
                  {row.vals.map((v, j) => (
                    <td key={j} style={{
                      textAlign: "center", padding: "5px 8px",
                      color: row.color && v != null ? (v > targetPct ? C.accent : C.ok) : C.text,
                      fontWeight: row.color && v != null && v > targetPct ? 700 : 400,
                    }}>
                      {row.fmt(v)}
                    </td>
                  ))}
                  <td style={{
                    textAlign: "center", padding: "5px 8px", fontWeight: 700,
                    color: row.color && row.total != null ? (row.total > targetPct ? C.accent : C.ok) : C.navy,
                  }}>
                    {row.total != null
                      ? row.color ? row.total.toFixed(2) + "%" : row.total.toLocaleString()
                      : "—"}
                  </td>
     
