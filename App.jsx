import { useState, useEffect } from "react";
import * as XLSX from "xlsx";
import {
  AreaChart, Area, LineChart, Line, XAxis, YAxis, CartesianGrid, Tooltip,
  Legend, ResponsiveContainer, ReferenceLine,
} from "recharts";

// ── CONFIG ─────────────────────────────────────────────────────────────────────
const FILE_ID    = "1vNrcI1yQE06_1QV5e6KUiILSGapq7Dzw";
const SHEET_NAME = "Quality Status (HQ)";
// Internal Vercel API proxy — no third-party dependency
const GDRIVE_URL = "/api/proxy";

// ── THEME ──────────────────────────────────────────────────────────────────────
const C = {
  // base
  bg0:     "#0B1220",
  bg1:     "#111B30",
  surface: "#FFFFFF",
  text:    "#0F172A",
  muted:   "#64748B",
  border:  "#E2E8F0",
  rowAlt:  "#F5F8FF",
  // brand accents (vivid, multi-color)
  indigo:  "#4F46E5",
  blue:    "#2563EB",
  cyan:    "#06B6D4",
  teal:    "#14B8A6",
  green:   "#22C55E",
  amber:   "#F59E0B",
  orange:  "#FB7185",
  red:     "#EF4444",
  violet:  "#8B5CF6",
  pink:    "#EC4899",
  // semantic
  ok:      "#16A34A",
  okBg:    "#DCFCE7",
  warn:    "#DC2626",
  warnBg:  "#FEE2E2",
  navy:    "#0F2748",
};

// gradient palette for charts / cards
const GRADS = [
  ["#4F46E5", "#8B5CF6"], // indigo→violet
  ["#06B6D4", "#3B82F6"], // cyan→blue
  ["#F59E0B", "#FB7185"], // amber→rose
  ["#22C55E", "#14B8A6"], // green→teal
  ["#EC4899", "#8B5CF6"], // pink→violet
];

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

// ── GLOBAL CSS (animations, responsive, scrollbars) ────────────────────────────
const GlobalStyle = () => (
  <style>{`
    @keyframes spin { to { transform: rotate(360deg); } }
    @keyframes fadeUp { from { opacity: 0; transform: translateY(14px); } to { opacity: 1; transform: translateY(0); } }
    @keyframes shimmer { 0% { background-position: -200% 0; } 100% { background-position: 200% 0; } }
    @keyframes floatBg { 0%,100% { transform: translate(0,0); } 50% { transform: translate(-3%, 4%); } }
    .ncr-card { animation: fadeUp 0.5s cubic-bezier(.2,.7,.3,1) both; }
    .ncr-kpi { transition: transform .25s ease, box-shadow .25s ease; }
    .ncr-kpi:hover { transform: translateY(-4px); box-shadow: 0 18px 40px -16px rgba(15,23,42,.35); }
    .ncr-refresh { transition: transform .4s ease; }
    .ncr-refresh:hover { transform: rotate(180deg); }
    .ncr-tab { transition: all .2s ease; }
    .ncr-tab:active { transform: scale(.96); }
    .ncr-table-wrap::-webkit-scrollbar { height: 8px; }
    .ncr-table-wrap::-webkit-scrollbar-thumb { background: ${C.border}; border-radius: 8px; }
    .ncr-row { transition: background .15s ease; }
    .ncr-row:hover { background: #EEF4FF !important; }
    @media (prefers-reduced-motion: reduce) {
      *, *::before, *::after { animation: none !important; transition: none !important; }
    }
  `}</style>
);

// ── SUB-COMPONENTS ─────────────────────────────────────────────────────────────
const CustomTooltip = ({ active, payload, label }) => {
  if (!active || !payload?.length) return null;
  return (
    <div style={{
      background: "rgba(15,23,42,.92)", color: "#fff", padding: "10px 14px",
      borderRadius: 10, fontSize: 12, backdropFilter: "blur(8px)",
      boxShadow: "0 10px 30px -10px rgba(0,0,0,.5)", border: "1px solid rgba(255,255,255,.08)",
    }}>
      <div style={{ fontWeight: 800, marginBottom: 6, letterSpacing: ".02em" }}>{label}</div>
      {payload.map((p) => (
        <div key={p.name} style={{ display: "flex", alignItems: "center", gap: 7, marginTop: 2 }}>
          <span style={{ width: 9, height: 9, borderRadius: "50%", background: p.color, display: "inline-block" }} />
          <span style={{ color: "#CBD5E1" }}>{p.name}:</span>
          <strong>{p.value != null ? p.value.toFixed(2) + "%" : "—"}</strong>
        </div>
      ))}
    </div>
  );
};

const KpiCard = ({ label, value, sub, grad, icon, delay = 0, danger }) => (
  <div className="ncr-card ncr-kpi" style={{
    position: "relative", overflow: "hidden",
    background: `linear-gradient(135deg, ${grad[0]}, ${grad[1]})`,
    borderRadius: 18, padding: "18px 20px",
    flex: "1 1 160px", minWidth: 150,
    color: "#fff", boxShadow: "0 10px 28px -14px rgba(15,23,42,.45)",
    animationDelay: `${delay}ms`,
  }}>
    {/* glow blob */}
    <div style={{
      position: "absolute", top: -30, right: -30, width: 110, height: 110,
      background: "rgba(255,255,255,.18)", borderRadius: "50%", filter: "blur(4px)",
    }} />
    <div style={{ position: "relative", display: "flex", justifyContent: "space-between", alignItems: "flex-start" }}>
      <div style={{ fontSize: 10.5, fontWeight: 700, letterSpacing: ".09em", textTransform: "uppercase", opacity: .92 }}>
        {label}
      </div>
      <span style={{ fontSize: 18, opacity: .95 }}>{icon}</span>
    </div>
    <div style={{ position: "relative", fontSize: 30, fontWeight: 800, lineHeight: 1.05, marginTop: 10, letterSpacing: "-.02em" }}>
      {value}
      {danger && <span style={{ fontSize: 13, marginLeft: 6 }}>▲</span>}
    </div>
    {sub && <div style={{ position: "relative", fontSize: 11.5, opacity: .9, marginTop: 6 }}>{sub}</div>}
  </div>
);

const Panel = ({ children, delay = 0, style = {} }) => (
  <div className="ncr-card" style={{
    background: C.surface, borderRadius: 18, padding: 18,
    border: `1px solid ${C.border}`,
    boxShadow: "0 1px 3px rgba(15,23,42,.04), 0 12px 30px -22px rgba(15,23,42,.25)",
    animationDelay: `${delay}ms`, ...style,
  }}>
    {children}
  </div>
);

const SectionHeader = ({ title, sub, accent = C.indigo }) => (
  <div style={{ display: "flex", alignItems: "center", gap: 10, marginBottom: 14 }}>
    <span style={{ width: 5, height: 26, borderRadius: 4, background: `linear-gradient(${accent}, ${C.cyan})` }} />
    <div>
      <div style={{ fontSize: 13, fontWeight: 800, color: C.navy, letterSpacing: ".01em" }}>{title}</div>
      {sub && <div style={{ fontSize: 11, color: C.muted, marginTop: 1 }}>{sub}</div>}
    </div>
  </div>
);

const Spinner = () => (
  <div style={{ width: 46, height: 46, borderRadius: "50%",
    background: `conic-gradient(${C.cyan}, ${C.indigo}, ${C.pink}, ${C.cyan})`,
    mask: "radial-gradient(farthest-side, transparent calc(100% - 5px), #000 0)",
    WebkitMask: "radial-gradient(farthest-side, transparent calc(100% - 5px), #000 0)",
    animation: "spin .8s linear infinite" }} />
);

const thStyle = (extra = {}) => ({
  background: "transparent", color: "#fff",
  padding: "9px 10px", fontSize: 11,
  fontWeight: 700, textAlign: "center",
  whiteSpace: "nowrap", letterSpacing: ".02em", ...extra,
});

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
    <div style={{ display: "flex", flexDirection: "column", alignItems: "center", justifyContent: "center", minHeight: "100vh", background: `radial-gradient(1200px 600px at 70% -10%, #1E2A4A, ${C.bg0})`, gap: 18 }}>
      <GlobalStyle />
      <Spinner />
      <div style={{ color: "#94A3B8", fontSize: 14, fontWeight: 500 }}>Loading data from Google Drive…</div>
    </div>
  );

  // ── ERROR ──
  if (error) return (
    <div style={{ display: "flex", flexDirection: "column", alignItems: "center", justifyContent: "center", minHeight: "100vh", background: `radial-gradient(1200px 600px at 70% -10%, #1E2A4A, ${C.bg0})`, gap: 14, padding: 24, textAlign: "center" }}>
      <GlobalStyle />
      <div style={{ fontSize: 40 }}>⚠️</div>
      <div style={{ fontSize: 17, fontWeight: 800, color: "#FCA5A5" }}>Unable to load file</div>
      <div style={{ fontSize: 13, color: "#94A3B8", maxWidth: 420 }}>{error}</div>
      <div style={{ fontSize: 12, color: "#CBD5E1", background: "rgba(255,255,255,.06)", padding: "12px 20px", borderRadius: 12, maxWidth: 460, lineHeight: 1.6, border: "1px solid rgba(255,255,255,.08)" }}>
        Make sure the Google Drive file is set to <strong style={{ color: "#fff" }}>"Anyone with the link → Viewer"</strong>
      </div>
      <button onClick={loadData} style={{ marginTop: 8, padding: "10px 26px", background: `linear-gradient(135deg, ${C.indigo}, ${C.cyan})`, color: "#fff", border: "none", borderRadius: 10, cursor: "pointer", fontWeight: 700, fontSize: 13, boxShadow: "0 10px 24px -10px rgba(79,70,229,.6)" }}>
        Try again
      </button>
    </div>
  );

  const { targetPct, cumDefectPct, totalInput, totalNcr, monthly, weekly, monthLabels, weekLabels, processRows } = data;
  const isOverTarget = cumDefectPct > targetPct;
  const ratio = targetPct > 0 ? Math.min((cumDefectPct / targetPct) * 100, 100) : 0;

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
    <div style={{
      fontFamily: "'Inter','Segoe UI',sans-serif",
      minHeight: "100vh", color: C.text,
      background: `radial-gradient(1100px 520px at 85% -8%, #1B2A4D 0%, transparent 55%), radial-gradient(900px 480px at 0% 0%, #15233F 0%, transparent 50%), linear-gradient(180deg, ${C.bg0}, #0E1730)`,
      padding: "clamp(14px, 3vw, 28px)",
    }}>
      <GlobalStyle />

      {/* ── HEADER ── */}
      <div className="ncr-card" style={{
        display: "flex", alignItems: "center", justifyContent: "space-between",
        flexWrap: "wrap", gap: 14, marginBottom: 18,
        background: "rgba(255,255,255,.04)", border: "1px solid rgba(255,255,255,.08)",
        borderRadius: 18, padding: "16px 20px", backdropFilter: "blur(8px)",
      }}>
        <div style={{ display: "flex", alignItems: "center", gap: 14 }}>
          <div style={{
            width: 46, height: 46, borderRadius: 13, flexShrink: 0,
            background: `linear-gradient(135deg, ${C.indigo}, ${C.cyan})`,
            display: "flex", alignItems: "center", justifyContent: "center",
            fontSize: 22, boxShadow: "0 10px 22px -10px rgba(6,182,212,.7)",
          }}>📊</div>
          <div>
            <div style={{ fontSize: 10, fontWeight: 700, letterSpacing: ".14em", textTransform: "uppercase", color: C.cyan, marginBottom: 3 }}>
              Internal Process Quality
            </div>
            <h1 style={{ margin: 0, fontSize: "clamp(18px, 3.5vw, 24px)", fontWeight: 800, color: "#fff", letterSpacing: "-.02em" }}>
              Y2026 NCR Status Dashboard
            </h1>
            {lastUpdated && (
              <div style={{ fontSize: 11, color: "#94A3B8", marginTop: 4, display: "flex", alignItems: "center", gap: 6 }}>
                <span style={{ width: 7, height: 7, borderRadius: "50%", background: C.green, boxShadow: `0 0 8px ${C.green}` }} />
                Updated {lastUpdated}
                <span onClick={loadData} className="ncr-refresh" style={{ cursor: "pointer", color: C.cyan, fontWeight: 700, marginLeft: 4 }} title="Refresh">↻</span>
              </div>
            )}
          </div>
        </div>

        {/* Target status pill */}
        <div style={{
          background: isOverTarget ? "rgba(220,38,38,.14)" : "rgba(22,163,74,.14)",
          border: `1px solid ${isOverTarget ? "rgba(248,113,113,.4)" : "rgba(74,222,128,.4)"}`,
          borderRadius: 14, padding: "10px 18px", textAlign: "center", minWidth: 140,
        }}>
          <div style={{ fontSize: 9.5, fontWeight: 700, color: "#94A3B8", textTransform: "uppercase", letterSpacing: ".1em" }}>Y2026 Target</div>
          <div style={{ fontSize: 24, fontWeight: 800, color: isOverTarget ? "#FCA5A5" : "#86EFAC" }}>
            {targetPct.toFixed(2)}%
          </div>
          <div style={{ fontSize: 10.5, color: isOverTarget ? "#FCA5A5" : "#86EFAC", fontWeight: 700 }}>
            {isOverTarget ? "⚠ Above target" : "✓ On track"}
          </div>
        </div>
      </div>

      {/* ── KPI ROW ── */}
      <div style={{ display: "flex", gap: 12, flexWrap: "wrap", marginBottom: 18 }}>
        <KpiCard label="Cumulative Defect Rate" value={cumDefectPct.toFixed(2) + "%"} grad={isOverTarget ? [C.orange, C.red] : GRADS[3]} icon="🎯" danger={isOverTarget} sub={`Target ${targetPct.toFixed(2)}%`} delay={0} />
        <KpiCard label="YTD Input Qty"  value={totalInput.toLocaleString()} grad={GRADS[1]} icon="📦" sub="ea · YTD 2026" delay={60} />
        <KpiCard label="YTD NCR"        value={totalNcr} grad={GRADS[0]} icon="🚩" sub="ea · YTD 2026" delay={120} />
        <KpiCard label="Input (5 weeks)" value={weeklyInputTotal.toLocaleString()} grad={GRADS[4]} icon="📈" sub={`${weekLabels[0]}–${weekLabels[weekLabels.length-1]}`} delay={180} />
        <KpiCard label="NCR (5 weeks)"   value={weeklyNcrTotal} grad={GRADS[2]} icon="⚡" sub={`${weekLabels[0]}–${weekLabels[weekLabels.length-1]}`} delay={240} />
      </div>

      {/* ── PROGRESS vs TARGET ── */}
      <Panel delay={120} style={{ marginBottom: 18 }}>
        <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 10, flexWrap: "wrap", gap: 6 }}>
          <div style={{ fontSize: 12.5, fontWeight: 800, color: C.navy }}>Cumulative defect rate vs target</div>
          <div style={{ fontSize: 12, fontWeight: 700, color: isOverTarget ? C.warn : C.ok }}>
            {cumDefectPct.toFixed(3)}% / {targetPct.toFixed(3)}%
          </div>
        </div>
        <div style={{ position: "relative", height: 14, borderRadius: 10, background: C.rowAlt, overflow: "hidden" }}>
          <div style={{
            height: "100%", width: `${ratio}%`, borderRadius: 10,
            background: isOverTarget
              ? `linear-gradient(90deg, ${C.amber}, ${C.red})`
              : `linear-gradient(90deg, ${C.teal}, ${C.green})`,
            transition: "width 1s cubic-bezier(.2,.7,.3,1)",
          }} />
        </div>
        <div style={{ fontSize: 11, color: C.muted, marginTop: 8 }}>
          {isOverTarget
            ? `Above target by ${(cumDefectPct - targetPct).toFixed(3)} pts — corrective action needed.`
            : `${(targetPct - cumDefectPct).toFixed(3)} pts of headroom below the target threshold.`}
        </div>
      </Panel>

      {/* ── CHARTS ── */}
      <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit, minmax(320px, 1fr))", gap: 14, marginBottom: 18 }}>

        {/* Monthly chart */}
        <Panel delay={160}>
          <SectionHeader title="NCR by month" sub="Defect Rate & Cumulative" accent={C.indigo} />
          <ResponsiveContainer width="100%" height={230}>
            <AreaChart data={chartMonthly} margin={{ top: 8, right: 14, left: -12, bottom: 0 }}>
              <defs>
                <linearGradient id="gDef" x1="0" y1="0" x2="0" y2="1">
                  <stop offset="0%" stopColor={C.blue} stopOpacity={0.35} />
                  <stop offset="100%" stopColor={C.blue} stopOpacity={0.02} />
                </linearGradient>
                <linearGradient id="gCum" x1="0" y1="0" x2="0" y2="1">
                  <stop offset="0%" stopColor={C.pink} stopOpacity={0.28} />
                  <stop offset="100%" stopColor={C.pink} stopOpacity={0.02} />
                </linearGradient>
              </defs>
              <CartesianGrid strokeDasharray="3 3" stroke="#EEF1F6" vertical={false} />
              <XAxis dataKey="name" tick={{ fontSize: 10, fill: C.muted }} axisLine={false} tickLine={false} />
              <YAxis tickFormatter={v => v + "%"} tick={{ fontSize: 10, fill: C.muted }} axisLine={false} tickLine={false} domain={[0, "auto"]} />
              <Tooltip content={<CustomTooltip />} />
              <Legend iconType="circle" wrapperStyle={{ fontSize: 11, paddingTop: 6 }} />
              <ReferenceLine y={targetPct} stroke={C.amber} strokeDasharray="5 4"
                label={{ value: `Target ${targetPct}%`, fontSize: 9, fill: C.amber, position: "insideTopLeft" }} />
              <Area type="monotone" dataKey="Defect Rate(%)" stroke={C.blue} strokeWidth={2.5} fill="url(#gDef)" dot={{ r: 3, fill: C.blue }} activeDot={{ r: 5 }} connectNulls={false} />
              <Area type="monotone" dataKey="Cumulative(%)" stroke={C.pink} strokeWidth={2.5} fill="url(#gCum)" strokeDasharray="6 3" dot={{ r: 3, fill: C.pink }} activeDot={{ r: 5 }} connectNulls={false} />
            </AreaChart>
          </ResponsiveContainer>
        </Panel>

        {/* Weekly chart + mini table */}
        <Panel delay={200}>
          <SectionHeader title="NCR by week" sub="Last 5 weeks" accent={C.cyan} />
          <ResponsiveContainer width="100%" height={150}>
            <AreaChart data={chartWeekly} margin={{ top: 8, right: 14, left: -12, bottom: 0 }}>
              <defs>
                <linearGradient id="gWk" x1="0" y1="0" x2="0" y2="1">
                  <stop offset="0%" stopColor={C.cyan} stopOpacity={0.4} />
                  <stop offset="100%" stopColor={C.cyan} stopOpacity={0.02} />
                </linearGradient>
              </defs>
              <CartesianGrid strokeDasharray="3 3" stroke="#EEF1F6" vertical={false} />
              <XAxis dataKey="name" tick={{ fontSize: 10, fill: C.muted }} axisLine={false} tickLine={false} />
              <YAxis tickFormatter={v => v + "%"} tick={{ fontSize: 10, fill: C.muted }} axisLine={false} tickLine={false} domain={[0, "auto"]} />
              <Tooltip content={<CustomTooltip />} />
              <ReferenceLine y={targetPct} stroke={C.amber} strokeDasharray="5 4" />
              <Area type="monotone" dataKey="Defect Rate(%)" stroke={C.cyan} strokeWidth={2.5} fill="url(#gWk)" dot={{ r: 4, fill: C.cyan }} activeDot={{ r: 6 }} />
            </AreaChart>
          </ResponsiveContainer>
          <div className="ncr-table-wrap" style={{ marginTop: 12, overflowX: "auto" }}>
            <table style={{ width: "100%", borderCollapse: "separate", borderSpacing: 0, fontSize: 11 }}>
              <thead>
                <tr style={{ background: `linear-gradient(90deg, ${C.navy}, #16335C)` }}>
                  <th style={thStyle({ textAlign: "left", borderTopLeftRadius: 8, borderBottomLeftRadius: 8 })}></th>
                  {weekLabels.map(w => <th key={w} style={thStyle()}>{w}</th>)}
                  <th style={thStyle({ borderTopRightRadius: 8, borderBottomRightRadius: 8 })}>Total</th>
                </tr>
              </thead>
              <tbody>
                {[
                  { label: "Input", vals: weekly.map(w => w.input), total: weeklyInputTotal, fmt: v => v.toLocaleString() },
                  { label: "NCR",   vals: weekly.map(w => w.ncr),   total: weeklyNcrTotal,  fmt: v => v },
                  { label: "Rate",  vals: weekly.map(w => w.defect),total: null, fmt: v => v.toFixed(2) + "%" },
                ].map((row, i) => (
                  <tr key={row.label} className="ncr-row" style={{ background: i % 2 === 0 ? C.surface : C.rowAlt }}>
                    <td style={{ padding: "6px 8px", fontWeight: 700, color: C.navy }}>{row.label}</td>
                    {row.vals.map((v, j) => (
                      <td key={j} style={{
                        textAlign: "center", padding: "6px 6px",
                        color: row.label === "Rate" && v > targetPct ? C.warn : C.text,
                        fontWeight: row.label === "Rate" && v > targetPct ? 700 : 400,
                      }}>
                        {row.fmt(v)}
                      </td>
                    ))}
                    <td style={{ textAlign: "center", padding: "6px 6px", fontWeight: 700, color: C.indigo }}>
                      {row.total != null ? row.total.toLocaleString() : "—"}
                    </td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        </Panel>
      </div>

      {/* ── MONTHLY RATE TABLE ── */}
      <Panel delay={240} style={{ marginBottom: 18 }}>
        <SectionHeader title="NCR Rate Summary — Monthly" accent={C.violet} />
        <div className="ncr-table-wrap" style={{ overflowX: "auto" }}>
          <table style={{ width: "100%", borderCollapse: "separate", borderSpacing: 0, fontSize: 11 }}>
            <thead>
              <tr style={{ background: `linear-gradient(90deg, ${C.navy}, #16335C)` }}>
                <th style={thStyle({ textAlign: "left", minWidth: 130, borderTopLeftRadius: 8, borderBottomLeftRadius: 8 })}>Metric</th>
                <th style={thStyle()}>2025</th>
                {monthLabels.map(m => <th key={m} style={thStyle()}>{m}</th>)}
                <th style={thStyle({ borderTopRightRadius: 8, borderBottomRightRadius: 8 })}>Total Y26</th>
              </tr>
            </thead>
            <tbody>
              {[
                { label: "Input (ea)",     vals: monthly.map(m => m.input),      total: totalInput,    fmt: v => v != null ? v.toLocaleString() : "—",        color: false },
                { label: "NCR (ea)",       vals: monthly.map(m => m.ncr),        total: totalNcr,      fmt: v => v != null ? v : "—",                          color: false },
                { label: "Defect Rate(%)", vals: monthly.map(m => m.defect),     total: null,          fmt: v => v != null ? v.toFixed(2) + "%" : "—",         color: true  },
                { label: "Cumulative(%)",  vals: monthly.map(m => m.cumulative), total: cumDefectPct,  fmt: v => v != null ? v.toFixed(2) + "%" : "—",         color: true  },
              ].map((row, i) => (
                <tr key={row.label} className="ncr-row" style={{ background: i % 2 === 0 ? C.surface : C.rowAlt }}>
                  <td style={{ padding: "7px 10px", fontWeight: 700, color: C.navy }}>{row.label}</td>
                  {row.vals.map((v, j) => (
                    <td key={j} style={{
                      textAlign: "center", padding: "7px 8px",
                      color: row.color && v != null ? (v > targetPct ? C.warn : C.ok) : C.text,
                      fontWeight: row.color && v != null && v > targetPct ? 700 : 400,
                    }}>
                      {row.fmt(v)}
                    </td>
                  ))}
                  <td style={{
                    textAlign: "center", padding: "7px 8px", fontWeight: 800,
                    color: row.color && row.total != null ? (row.total > targetPct ? C.warn : C.ok) : C.indigo,
                  }}>
                    {row.total != null
                      ? row.color ? row.total.toFixed(2) + "%" : row.total.toLocaleString()
                      : "—"}
                  </td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      </Panel>

      {/* ── PROCESS TABLE ── */}
      <Panel delay={280}>
        <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between", marginBottom: 14, flexWrap: "wrap", gap: 10 }}>
          <SectionHeader title="NCR by process" accent={C.teal} />
          <div style={{ display: "flex", gap: 6, background: C.rowAlt, padding: 4, borderRadius: 10 }}>
            {[["monthly","Monthly"],["weekly","Weekly (5W)"]].map(([t, lbl]) => (
              <button key={t} onClick={() => setTab(t)} className="ncr-tab" style={{
                padding: "6px 16px", fontSize: 11.5, fontWeight: 700,
                borderRadius: 8, border: "none", cursor: "pointer",
                background: tab === t ? `linear-gradient(135deg, ${C.indigo}, ${C.cyan})` : "transparent",
                color: tab === t ? "#fff" : C.muted,
                boxShadow: tab === t ? "0 6px 16px -8px rgba(79,70,229,.7)" : "none",
              }}>
                {lbl}
              </button>
            ))}
          </div>
        </div>
        <div className="ncr-table-wrap" style={{ overflowX: "auto" }}>
          <table style={{ width: "100%", borderCollapse: "separate", borderSpacing: 0, fontSize: 11 }}>
            <thead>
              <tr style={{ background: `linear-gradient(90deg, ${C.navy}, #16335C)` }}>
                <th style={thStyle({ textAlign: "left", minWidth: 110, borderTopLeftRadius: 8, borderBottomLeftRadius: 8 })}>Process</th>
                {tab === "monthly" ? (
                  <>
                    <th style={thStyle()}>2025</th>
                    {monthLabels.map(m => <th key={m} style={thStyle()}>{m}</th>)}
                    <th style={thStyle({ borderTopRightRadius: 8, borderBottomRightRadius: 8 })}>Total Y26</th>
                  </>
                ) : (
                  <>
                    {weekLabels.map(w => <th key={w} style={thStyle()}>{w}</th>)}
                    <th style={thStyle({ borderTopRightRadius: 8, borderBottomRightRadius: 8 })}>Total</th>
                  </>
                )}
              </tr>
            </thead>
            <tbody>
              {processRows.map((row, i) => {
                const vals = tab === "monthly"
                  ? [row.ncr2025, ...row.months, row.total]
                  : [...row.weeks, row.weekTotal];
                return (
                  <tr key={row.process} className="ncr-row" style={{
                    background: row.isTotal ? `linear-gradient(90deg, ${C.navy}, #16335C)` : i % 2 === 0 ? C.surface : C.rowAlt,
                  }}>
                    <td style={{ padding: "7px 10px", fontWeight: row.isTotal ? 800 : 600, color: row.isTotal ? "#fff" : C.navy }}>
                      {row.process}
                    </td>
                    {vals.map((v, j) => (
                      <td key={j} style={{
                        textAlign: "center", padding: "7px 8px",
                        color: row.isTotal ? (v > 0 ? C.amber : "#9DB4D4") : v > 0 ? C.warn : "#CBD5E1",
                        fontWeight: v > 0 ? 700 : 400,
                      }}>
                        {v}
                      </td>
                    ))}
                  </tr>
                );
              })}
            </tbody>
          </table>
        </div>
      </Panel>

      {/* ── FOOTER ── */}
      <div style={{ textAlign: "center", marginTop: 18, fontSize: 11, color: "#64748B" }}>
        Y2026 Internal Process Quality Status (NCR) · Live data from Google Drive
      </div>
    </div>
  );
}
