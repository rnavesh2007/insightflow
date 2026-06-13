/**
 * ╔══════════════════════════════════════════════════════════════════════╗
 * ║              InsightFlow — AI-Powered Data Dashboard                ║
 * ╠══════════════════════════════════════════════════════════════════════╣
 * ║  © 2025 Navesh R. All rights reserved.                             ║
 * ║  Bachelor of Technology in Artificial Intelligence & Data Science   ║
 * ║  Vel Tech High Tech Dr. Rangarajan Dr. Sakunthala Engineering       ║
 * ║  College, Chennai, Tamil Nadu, India                                ║
 * ║                                                                      ║
 * ║  Unauthorized reproduction or distribution of this software,        ║
 * ║  in whole or in part, is strictly prohibited.                       ║
 * ╚══════════════════════════════════════════════════════════════════════╝
 */

import { useState, useEffect, useRef } from "react";
import Papa from "papaparse";
import * as XLSX from "xlsx";
import {
  BarChart, Bar, LineChart, Line, PieChart, Pie, Cell,
  ScatterChart, Scatter, XAxis, YAxis, CartesianGrid, Tooltip,
  Legend, ResponsiveContainer, AreaChart, Area, RadarChart,
  PolarGrid, PolarAngleAxis, PolarRadiusAxis, Radar,
} from "recharts";

const C = {
  bg: "#060D16",
  surface: "#0C1A28",
  card: "#101F30",
  cardHover: "#142436",
  border: "#1B3050",
  borderLight: "#24405F",
  cyan: "#00D4F5",
  cyanDim: "#00D4F520",
  purple: "#A78BFA",
  purpleDim: "#A78BFA20",
  green: "#34D399",
  greenDim: "#34D39920",
  amber: "#FBBF24",
  amberDim: "#FBBF2420",
  rose: "#FB7185",
  roseDim: "#FB718520",
  blue: "#60A5FA",
  text: "#E8F0FE",
  textMid: "#94A3B8",
  muted: "#4A5D74",
};

const PIE_COLORS = [C.cyan, C.purple, C.green, C.amber, C.rose, C.blue, "#F472B6", "#4ADE80", "#FB923C", "#818CF8"];

const TOOLTIP_STYLE = {
  backgroundColor: "#0C1A28",
  border: `1px solid ${C.border}`,
  borderRadius: 10,
  fontSize: 13,
  color: C.text,
  boxShadow: "0 8px 24px #00000060",
};

function detectColumnTypes(rows) {
  if (!rows.length) return {};
  const sample = rows.slice(0, 60);
  const types = {};
  Object.keys(rows[0]).forEach(key => {
    const vals = sample.map(r => r[key]).filter(v => v !== "" && v != null);
    const numericCount = vals.filter(v => !isNaN(parseFloat(v)) && isFinite(v)).length;
    types[key] = numericCount / (vals.length || 1) > 0.7 ? "number" : "string";
  });
  return types;
}

function computeStats(rows, col) {
  const vals = rows.map(r => parseFloat(r[col])).filter(v => !isNaN(v));
  if (!vals.length) return null;
  const sum = vals.reduce((a, b) => a + b, 0);
  const mean = sum / vals.length;
  const sorted = [...vals].sort((a, b) => a - b);
  const median = sorted[Math.floor(sorted.length / 2)];
  const min = sorted[0];
  const max = sorted[sorted.length - 1];
  const variance = vals.reduce((acc, v) => acc + (v - mean) ** 2, 0) / vals.length;
  const q1 = sorted[Math.floor(sorted.length * 0.25)];
  const q3 = sorted[Math.floor(sorted.length * 0.75)];
  return { sum, mean, median, min, max, std: Math.sqrt(variance), count: vals.length, q1, q3 };
}

function fmtNum(n) {
  if (n == null || isNaN(n)) return "—";
  if (Math.abs(n) >= 1e9) return (n / 1e9).toFixed(2) + "B";
  if (Math.abs(n) >= 1e6) return (n / 1e6).toFixed(2) + "M";
  if (Math.abs(n) >= 1e3) return (n / 1e3).toFixed(1) + "K";
  return parseFloat(n.toFixed(2)).toLocaleString();
}

function fmtInt(n) { return Math.round(n).toLocaleString(); }

function AnimatedNumber({ value, duration = 1400, formatter = fmtNum }) {
  const [display, setDisplay] = useState(0);
  const startRef = useRef(null);
  const rafRef = useRef(null);
  useEffect(() => {
    const target = parseFloat(value) || 0;
    startRef.current = null;
    const step = ts => {
      if (!startRef.current) startRef.current = ts;
      const p = Math.min((ts - startRef.current) / duration, 1);
      const eased = 1 - Math.pow(1 - p, 4);
      setDisplay(eased * target);
      if (p < 1) rafRef.current = requestAnimationFrame(step);
    };
    rafRef.current = requestAnimationFrame(step);
    return () => cancelAnimationFrame(rafRef.current);
  }, [value, duration]);
  return <span>{formatter(display)}</span>;
}

async function callClaudeAnalysis(rows, colTypes, fileName) {
  const numCols = Object.keys(colTypes).filter(k => colTypes[k] === "number");
  const strCols = Object.keys(colTypes).filter(k => colTypes[k] === "string");
  const stats = numCols.map(col => {
    const s = computeStats(rows, col);
    return s ? `${col}: mean=${fmtNum(s.mean)}, min=${fmtNum(s.min)}, max=${fmtNum(s.max)}, sum=${fmtNum(s.sum)}, std=${fmtNum(s.std)}` : `${col}: no data`;
  });
  const sample = rows.slice(0, 5);
  const prompt = `You are a senior data analyst. Analyze this dataset and provide actionable insights.

Dataset: "${fileName}"
Total rows: ${rows.length}
Numeric columns: ${numCols.join(", ")}
Categorical columns: ${strCols.join(", ")}

Statistics:
${stats.join("\n")}

Sample rows (first 5):
${JSON.stringify(sample, null, 2)}

Respond ONLY with a valid JSON object (no markdown, no backticks, no extra text):
{
  "summary": "2-3 sentence plain-English overview of what this dataset contains and represents",
  "insights": [
    { "title": "short title", "description": "1-2 sentence insight", "type": "positive|warning|info|critical" }
  ],
  "recommendations": [
    { "title": "short title", "action": "concrete recommendation text" }
  ],
  "anomalies": [
    { "column": "col name", "detail": "what looks unusual" }
  ],
  "topFindings": ["finding 1", "finding 2", "finding 3"]
}

Provide exactly 4-5 insights, 3 recommendations, and 1-3 anomalies (or empty array if none).`;

  const response = await fetch("https://api.anthropic.com/v1/messages", {
    method: "POST",
    headers: {
      "Content-Type": "application/json",
      "x-api-key": "sk-ant-உன்key இங்க",
"anthropic-version": "2023-06-01",
      "anthropic-dangerous-direct-browser-access": "true",
    },
    body: JSON.stringify({
      model: "claude-sonnet-4-6",
      max_tokens: 1000,
      messages: [{ role: "user", content: prompt }],
    }),
  });
  const data = await response.json();
  const raw = data.content?.map(b => b.text || "").join("") || "";
  const clean = raw.replace(/```json|```/g, "").trim();
  return JSON.parse(clean);
}

function KPICard({ label, value, sub, accent, formatter, icon }) {
  const [hovered, setHovered] = useState(false);
  return (
    <div
      onMouseEnter={() => setHovered(true)}
      onMouseLeave={() => setHovered(false)}
      style={{
        background: hovered ? C.cardHover : C.card,
        borderRadius: 16, padding: "20px 22px",
        border: `1px solid ${hovered ? C.borderLight : C.border}`,
        flex: "1 1 160px", minWidth: 155,
        transition: "all 0.2s",
        boxShadow: hovered ? `0 8px 32px ${accent}15` : "none",
        position: "relative", overflow: "hidden",
      }}
    >
      <div style={{ position: "absolute", top: -20, right: -20, width: 80, height: 80, borderRadius: "50%", background: accent, opacity: 0.06 }} />
      <div style={{ display: "flex", alignItems: "flex-start", justifyContent: "space-between", marginBottom: 10 }}>
        <p style={{ color: C.muted, fontSize: 11, fontWeight: 600, letterSpacing: "0.8px", textTransform: "uppercase" }}>{label}</p>
        <span style={{ fontSize: 18 }}>{icon}</span>
      </div>
      <p style={{ color: accent || C.cyan, fontSize: 26, fontWeight: 700, fontFamily: "'Space Grotesk', sans-serif", marginBottom: 4 }}>
        <AnimatedNumber value={value} formatter={formatter || fmtNum} />
      </p>
      <p style={{ color: C.muted, fontSize: 12 }}>{sub}</p>
    </div>
  );
}

function ChartCard({ title, subtitle, children, span = 1, accent }) {
  return (
    <div style={{
      background: C.card, borderRadius: 16, padding: "20px 20px 14px",
      border: `1px solid ${C.border}`,
      gridColumn: span === 2 ? "span 2" : "span 1",
    }}>
      <div style={{ marginBottom: 16, display: "flex", alignItems: "flex-start", justifyContent: "space-between" }}>
        <div>
          <p style={{ color: C.text, fontSize: 14, fontWeight: 600, fontFamily: "'Space Grotesk', sans-serif" }}>{title}</p>
          {subtitle && <p style={{ color: C.muted, fontSize: 12, marginTop: 2 }}>{subtitle}</p>}
        </div>
        {accent && <div style={{ width: 8, height: 8, borderRadius: "50%", background: accent, marginTop: 4, boxShadow: `0 0 8px ${accent}` }} />}
      </div>
      {children}
    </div>
  );
}

function InsightBadge({ type }) {
  const map = { positive: [C.green, "✓ Positive"], warning: [C.amber, "⚠ Warning"], info: [C.cyan, "ℹ Info"], critical: [C.rose, "✕ Critical"] };
  const [color, label] = map[type] || map.info;
  return (
    <span style={{ padding: "2px 10px", borderRadius: 99, background: `${color}18`, color, fontSize: 11, fontWeight: 600, border: `1px solid ${color}30` }}>
      {label}
    </span>
  );
}

function AxisSelect({ label, value, onChange, options }) {
  return (
    <label style={{ display: "flex", alignItems: "center", gap: 8, fontSize: 13 }}>
      <span style={{ color: C.muted, fontWeight: 500 }}>{label}</span>
      <div style={{ position: "relative" }}>
        <select value={value} onChange={e => onChange(e.target.value)} style={{
          background: C.card, border: `1px solid ${C.border}`, color: C.text,
          borderRadius: 8, padding: "5px 28px 5px 10px", fontSize: 13,
          cursor: "pointer", outline: "none", appearance: "none",
        }}>
          {options.map(c => <option key={c} value={c}>{c}</option>)}
        </select>
        <span style={{ position: "absolute", right: 8, top: "50%", transform: "translateY(-50%)", color: C.muted, pointerEvents: "none", fontSize: 10 }}>▾</span>
      </div>
    </label>
  );
}

function UploadScreen({ onData }) {
  const [dragging, setDragging] = useState(false);
  const [loading, setLoading] = useState(false);
  const inputRef = useRef();

  const parse = file => {
    setLoading(true);
    const ext = file.name.split(".").pop().toLowerCase();
    if (ext === "csv") {
      Papa.parse(file, {
        header: true, skipEmptyLines: true,
        complete: r => { setLoading(false); onData(r.data, file.name); },
      });
    } else if (ext === "xlsx" || ext === "xls") {
      const reader = new FileReader();
      reader.onload = e => {
        const wb = XLSX.read(e.target.result, { type: "binary" });
        const ws = wb.Sheets[wb.SheetNames[0]];
        const data = XLSX.utils.sheet_to_json(ws, { defval: "" });
        setLoading(false);
        onData(data, file.name);
      };
      reader.readAsBinaryString(file);
    }
  };

  return (
    <div style={{ minHeight: "100vh", background: C.bg, display: "flex", flexDirection: "column", alignItems: "center", justifyContent: "center", fontFamily: "'Inter', sans-serif", padding: 24, position: "relative", overflow: "hidden" }}>
      <style>{`
        @import url('https://fonts.googleapis.com/css2?family=Space+Grotesk:wght@400;500;600;700&family=Inter:wght@400;500;600&display=swap');
        * { box-sizing: border-box; margin: 0; padding: 0; }
        @keyframes float { 0%,100%{transform:translateY(0px)} 50%{transform:translateY(-10px)} }
        @keyframes spin { from{transform:rotate(0deg)} to{transform:rotate(360deg)} }
        ::-webkit-scrollbar{width:6px;height:6px} ::-webkit-scrollbar-track{background:${C.bg}} ::-webkit-scrollbar-thumb{background:${C.border};border-radius:3px}
      `}</style>
      <div style={{ position: "absolute", inset: 0, backgroundImage: `radial-gradient(${C.border} 1px, transparent 1px)`, backgroundSize: "40px 40px", opacity: 0.4, pointerEvents: "none" }} />
      <div style={{ position: "absolute", inset: 0, background: `radial-gradient(ellipse 80% 60% at 50% 40%, ${C.cyan}08 0%, transparent 70%)`, pointerEvents: "none" }} />
      <div style={{ marginBottom: 52, textAlign: "center", position: "relative", animation: "float 4s ease-in-out infinite" }}>
        <div style={{ display: "inline-flex", alignItems: "center", gap: 14, marginBottom: 14 }}>
          <svg width="48" height="48" viewBox="0 0 48 48" fill="none">
            <rect width="48" height="48" rx="14" fill={C.cyan} fillOpacity="0.15" />
            <rect x="8" y="26" width="7" height="16" rx="2.5" fill={C.cyan} />
            <rect x="20" y="17" width="7" height="25" rx="2.5" fill={C.purple} />
            <rect x="32" y="9" width="7" height="33" rx="2.5" fill={C.green} />
          </svg>
          <div>
            <div style={{ fontFamily: "'Space Grotesk', sans-serif", fontSize: 32, fontWeight: 700, color: C.text, letterSpacing: "-0.8px", lineHeight: 1 }}>
              Insight<span style={{ color: C.cyan }}>Flow</span>
            </div>
            <div style={{ color: C.muted, fontSize: 12, fontWeight: 500, letterSpacing: "1.5px", textTransform: "uppercase", marginTop: 4 }}>
              AI · Data Analytics Platform
            </div>
          </div>
        </div>
        <p style={{ color: C.textMid, fontSize: 15, maxWidth: 460 }}>
          Upload any spreadsheet and instantly get an AI-powered interactive analytics dashboard
        </p>
      </div>
      <div
        onDragOver={e => { e.preventDefault(); setDragging(true); }}
        onDragLeave={() => setDragging(false)}
        onDrop={e => { e.preventDefault(); setDragging(false); const f = e.dataTransfer.files[0]; if (f) parse(f); }}
        onClick={() => !loading && inputRef.current.click()}
        style={{
          width: "100%", maxWidth: 540, border: `2px dashed`,
          borderColor: dragging ? C.cyan : C.border,
          borderRadius: 24, padding: "60px 36px", textAlign: "center",
          background: dragging ? `${C.cyan}06` : C.surface,
          cursor: loading ? "wait" : "pointer",
          transition: "all 0.25s", position: "relative",
          boxShadow: dragging ? `0 0 60px ${C.cyan}18` : "0 20px 60px #00000040",
        }}
      >
        <input ref={inputRef} type="file" accept=".csv,.xlsx,.xls" style={{ display: "none" }}
          onChange={e => e.target.files[0] && parse(e.target.files[0])} />
        {loading ? (
          <div style={{ display: "flex", flexDirection: "column", alignItems: "center", gap: 16 }}>
            <div style={{ width: 48, height: 48, border: `3px solid ${C.border}`, borderTop: `3px solid ${C.cyan}`, borderRadius: "50%", animation: "spin 0.8s linear infinite" }} />
            <p style={{ color: C.textMid, fontSize: 15 }}>Parsing your data…</p>
          </div>
        ) : (
          <>
            <div style={{ width: 72, height: 72, borderRadius: 20, background: `${C.cyan}12`, display: "flex", alignItems: "center", justifyContent: "center", margin: "0 auto 24px", border: `1px solid ${C.cyan}25` }}>
              <svg width="32" height="32" fill="none" viewBox="0 0 24 24" stroke={C.cyan} strokeWidth={1.6}>
                <path strokeLinecap="round" strokeLinejoin="round" d="M3 16.5v2.25A2.25 2.25 0 005.25 21h13.5A2.25 2.25 0 0021 18.75V16.5m-13.5-9L12 3m0 0l4.5 4.5M12 3v13.5" />
              </svg>
            </div>
            <p style={{ color: C.text, fontSize: 18, fontWeight: 600, fontFamily: "'Space Grotesk', sans-serif", marginBottom: 8 }}>Drop your spreadsheet here</p>
            <p style={{ color: C.muted, fontSize: 14, marginBottom: 24 }}>or click to browse your files</p>
            <div style={{ display: "flex", gap: 8, justifyContent: "center" }}>
              {[".CSV", ".XLSX", ".XLS"].map(f => (
                <span key={f} style={{ padding: "4px 14px", borderRadius: 99, border: `1px solid ${C.border}`, color: C.muted, fontSize: 12, fontWeight: 600 }}>{f}</span>
              ))}
            </div>
          </>
        )}
      </div>
      <div style={{ display: "flex", gap: 12, marginTop: 32, flexWrap: "wrap", justifyContent: "center" }}>
        {[["🤖", "AI Insights"], ["📊", "6 Chart Types"], ["📋", "Stats Engine"], ["💾", "CSV Export"]].map(([icon, label]) => (
          <div key={label} style={{ display: "flex", alignItems: "center", gap: 6, padding: "6px 14px", borderRadius: 99, background: C.surface, border: `1px solid ${C.border}`, color: C.textMid, fontSize: 13 }}>
            <span>{icon}</span><span>{label}</span>
          </div>
        ))}
      </div>
      <div style={{ position: "absolute", bottom: 20, left: 0, right: 0, textAlign: "center" }}>
        <p style={{ color: C.muted, fontSize: 12 }}>
          © 2025 <span style={{ color: C.cyan, fontWeight: 600 }}>Navesh R</span> · B.Tech AI & Data Science · Vel Tech High Tech Dr. Rangarajan Dr. Sakunthala Engineering College, Chennai
        </p>
      </div>
    </div>
  );
}

function Dashboard({ rows, fileName, onReset }) {
  const colTypes = detectColumnTypes(rows);
  const numCols = Object.keys(colTypes).filter(k => colTypes[k] === "number");
  const strCols = Object.keys(colTypes).filter(k => colTypes[k] === "string");
  const allCols = Object.keys(rows[0] || {});

  const [xAxis, setXAxis] = useState(strCols[0] || allCols[0] || "");
  const [yAxis, setYAxis] = useState(numCols[0] || "");
  const [yAxis2, setYAxis2] = useState(numCols[1] || "");
  const [activeTab, setActiveTab] = useState("overview");
  const [page, setPage] = useState(0);
  const [searchQuery, setSearchQuery] = useState("");
  const [sortCol, setSortCol] = useState(null);
  const [sortDir, setSortDir] = useState("asc");
  const [aiResult, setAiResult] = useState(null);
  const [aiLoading, setAiLoading] = useState(false);
  const [aiError, setAiError] = useState(null);

  const PAGE_SIZE = 15;
  const primaryStats = yAxis ? computeStats(rows, yAxis) : null;
  const nullCount = yAxis ? rows.filter(r => r[yAxis] === "" || r[yAxis] == null).length : 0;

  const chartData = (() => {
    if (!xAxis || !yAxis) return [];
    const agg = {};
    rows.forEach(r => {
      const k = String(r[xAxis] || "").slice(0, 28);
      if (!agg[k]) agg[k] = { name: k, [yAxis]: 0, [yAxis2]: 0, _count: 0 };
      agg[k][yAxis] += parseFloat(r[yAxis]) || 0;
      if (yAxis2) agg[k][yAxis2] += parseFloat(r[yAxis2]) || 0;
      agg[k]._count++;
    });
    return Object.values(agg).sort((a, b) => b[yAxis] - a[yAxis]).slice(0, 20)
      .map(d => ({ ...d, [yAxis]: parseFloat(d[yAxis].toFixed(2)), [yAxis2]: parseFloat((d[yAxis2] || 0).toFixed(2)) }));
  })();

  const pieData = (() => {
    const col = strCols[0];
    if (!col) return [];
    const counts = {};
    rows.forEach(r => { const v = String(r[col] || "Other").slice(0, 25); counts[v] = (counts[v] || 0) + 1; });
    return Object.entries(counts).sort((a, b) => b[1] - a[1]).slice(0, 9).map(([name, value]) => ({ name, value }));
  })();

  const scatterData = (() => {
    if (numCols.length < 2) return [];
    return rows.slice(0, 300).map(r => ({ x: parseFloat(r[numCols[0]]) || 0, y: parseFloat(r[numCols[1]]) || 0 }));
  })();

  const radarData = (() => {
    if (numCols.length < 3) return [];
    return numCols.slice(0, 6).map(col => {
      const s = computeStats(rows, col);
      return { subject: col.slice(0, 12), value: s ? parseFloat(s.mean.toFixed(2)) : 0 };
    });
  })();

  const filteredRows = rows.filter(r =>
    !searchQuery || allCols.some(c => String(r[c] || "").toLowerCase().includes(searchQuery.toLowerCase()))
  );
  const sortedRows = sortCol
    ? [...filteredRows].sort((a, b) => {
        const av = a[sortCol], bv = b[sortCol];
        const an = parseFloat(av), bn = parseFloat(bv);
        const cmp = (!isNaN(an) && !isNaN(bn)) ? an - bn : String(av).localeCompare(String(bv));
        return sortDir === "asc" ? cmp : -cmp;
      })
    : filteredRows;
  const tableRows = sortedRows.slice(page * PAGE_SIZE, (page + 1) * PAGE_SIZE);
  const totalPages = Math.ceil(sortedRows.length / PAGE_SIZE);

  const exportCSV = () => {
    const header = allCols.join(",");
    const body = rows.map(r => allCols.map(c => `"${String(r[c] || "").replace(/"/g, '""')}"`).join(",")).join("\n");
    const blob = new Blob([header + "\n" + body], { type: "text/csv" });
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a"); a.href = url; a.download = fileName.replace(/\.[^.]+$/, "") + "_export.csv"; a.click();
  };

  const runAI = async () => {
    setAiLoading(true); setAiError(null); setAiResult(null); setActiveTab("ai");
    try {
      const result = await callClaudeAnalysis(rows, colTypes, fileName);
      setAiResult(result);
    } catch (e) {
      setAiError("AI analysis failed. Please try again. (" + e.message + ")");
    }
    setAiLoading(false);
  };

  const tabs = [
    { id: "overview", label: "Overview", icon: "📊" },
    { id: "charts", label: "Charts", icon: "📈" },
    { id: "stats", label: "Statistics", icon: "🔢" },
    { id: "data", label: "Data Table", icon: "🗄️" },
    { id: "ai", label: "AI Insights", icon: "🤖" },
  ];

  return (
    <div style={{ minHeight: "100vh", background: C.bg, fontFamily: "'Inter', sans-serif", color: C.text }}>
      <style>{`
        @import url('https://fonts.googleapis.com/css2?family=Space+Grotesk:wght@400;500;600;700&family=Inter:wght@400;500;600&display=swap');
        *{box-sizing:border-box} ::-webkit-scrollbar{width:6px;height:6px} ::-webkit-scrollbar-track{background:${C.bg}} ::-webkit-scrollbar-thumb{background:${C.border};border-radius:3px}
        select{appearance:none} input::placeholder{color:${C.muted}}
        @keyframes spin{from{transform:rotate(0deg)}to{transform:rotate(360deg)}}
        @keyframes fadeIn{from{opacity:0;transform:translateY(8px)}to{opacity:1;transform:translateY(0)}}
        @keyframes pulse{0%,100%{opacity:1}50%{opacity:0.5}}
      `}</style>

      <header style={{ background: C.surface, borderBottom: `1px solid ${C.border}`, padding: "0 28px", display: "flex", alignItems: "center", justifyContent: "space-between", height: 62, position: "sticky", top: 0, zIndex: 100 }}>
        <div style={{ display: "flex", alignItems: "center", gap: 12 }}>
          <svg width="30" height="30" viewBox="0 0 48 48" fill="none">
            <rect width="48" height="48" rx="12" fill={C.cyan} fillOpacity="0.15" />
            <rect x="8" y="26" width="7" height="16" rx="2.5" fill={C.cyan} />
            <rect x="20" y="17" width="7" height="25" rx="2.5" fill={C.purple} />
            <rect x="32" y="9" width="7" height="33" rx="2.5" fill={C.green} />
          </svg>
          <span style={{ fontFamily: "'Space Grotesk', sans-serif", fontWeight: 700, fontSize: 20 }}>
            Insight<span style={{ color: C.cyan }}>Flow</span>
          </span>
          <div style={{ width: 1, height: 20, background: C.border, margin: "0 4px" }} />
          <span style={{ padding: "3px 12px", borderRadius: 99, background: `${C.cyan}18`, color: C.cyan, fontSize: 12, fontWeight: 600, maxWidth: 220, overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap" }}>
            📄 {fileName}
          </span>
        </div>
        <div style={{ display: "flex", alignItems: "center", gap: 12 }}>
          <span style={{ color: C.muted, fontSize: 13 }}>
            <span style={{ color: C.text, fontWeight: 600 }}>{rows.length.toLocaleString()}</span> rows ·{" "}
            <span style={{ color: C.text, fontWeight: 600 }}>{allCols.length}</span> cols
          </span>
          <button onClick={exportCSV} style={{ padding: "6px 16px", borderRadius: 8, border: `1px solid ${C.border}`, background: `${C.green}18`, color: C.green, fontSize: 13, fontWeight: 600, cursor: "pointer" }}>↓ Export</button>
          <button onClick={runAI} disabled={aiLoading} style={{ padding: "6px 16px", borderRadius: 8, border: "none", background: `linear-gradient(135deg, ${C.cyan}, ${C.purple})`, color: "#fff", fontSize: 13, fontWeight: 600, cursor: "pointer", opacity: aiLoading ? 0.7 : 1 }}>
            {aiLoading ? "⏳ Analyzing…" : "🤖 AI Analysis"}
          </button>
          <button onClick={onReset} style={{ padding: "6px 16px", borderRadius: 8, border: `1px solid ${C.border}`, background: "transparent", color: C.muted, fontSize: 13, cursor: "pointer" }}>↑ New file</button>
        </div>
      </header>

      <div style={{ background: C.surface, borderBottom: `1px solid ${C.border}`, padding: "10px 28px", display: "flex", alignItems: "center", gap: 12, flexWrap: "wrap" }}>
        <div style={{ display: "flex", gap: 2, background: C.bg, borderRadius: 12, padding: 4 }}>
          {tabs.map(t => (
            <button key={t.id} onClick={() => setActiveTab(t.id)} style={{
              padding: "6px 16px", borderRadius: 9, border: "none",
              background: activeTab === t.id ? C.card : "transparent",
              color: activeTab === t.id ? C.text : C.muted,
              fontSize: 13, fontWeight: 600, cursor: "pointer", transition: "all 0.15s",
              display: "flex", alignItems: "center", gap: 6,
            }}>
              <span>{t.icon}</span>{t.label}
              {t.id === "ai" && aiResult && <span style={{ width: 6, height: 6, borderRadius: "50%", background: C.green, display: "inline-block" }} />}
            </button>
          ))}
        </div>
        <div style={{ flex: 1 }} />
        <AxisSelect label="X Axis" value={xAxis} onChange={setXAxis} options={allCols} />
        <AxisSelect label="Y Axis" value={yAxis} onChange={setYAxis} options={numCols} />
        {numCols.length > 1 && <AxisSelect label="Y2" value={yAxis2} onChange={setYAxis2} options={numCols} />}
      </div>

      <main style={{ padding: "24px 28px", maxWidth: 1440, margin: "0 auto" }}>

        {activeTab === "overview" && (
          <div style={{ animation: "fadeIn 0.3s ease" }}>
            <div style={{ display: "flex", gap: 14, flexWrap: "wrap", marginBottom: 24 }}>
              <KPICard label="Total Records" value={rows.length} sub={`${allCols.length} columns`} accent={C.cyan} icon="📋" formatter={fmtInt} />
              <KPICard label="Numeric Cols" value={numCols.length} sub={`${strCols.length} categorical`} accent={C.purple} icon="🔢" formatter={fmtInt} />
              {primaryStats && <>
                <KPICard label={`Sum · ${yAxis}`} value={primaryStats.sum} sub={`${primaryStats.count} records`} accent={C.green} icon="∑" />
                <KPICard label={`Mean · ${yAxis}`} value={primaryStats.mean} sub={`Median: ${fmtNum(primaryStats.median)}`} accent={C.amber} icon="μ" />
                <KPICard label={`Max · ${yAxis}`} value={primaryStats.max} sub={`Min: ${fmtNum(primaryStats.min)}`} accent={C.rose} icon="↑" />
              </>}
            </div>
            <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 20 }}>
              <ChartCard title={`${yAxis} by ${xAxis}`} subtitle="Grouped bar chart" accent={C.cyan}>
                <ResponsiveContainer width="100%" height={270}>
                  <BarChart data={chartData} margin={{ top: 0, right: 10, left: 0, bottom: 50 }}>
                    <CartesianGrid strokeDasharray="3 3" stroke={C.border} />
                    <XAxis dataKey="name" stroke={C.muted} tick={{ fontSize: 10, fill: C.muted }} angle={-38} textAnchor="end" interval={0} />
                    <YAxis stroke={C.muted} tick={{ fontSize: 11, fill: C.muted }} tickFormatter={fmtNum} />
                    <Tooltip contentStyle={TOOLTIP_STYLE} formatter={v => fmtNum(v)} />
                    <Bar dataKey={yAxis} fill={C.cyan} radius={[5, 5, 0, 0]} maxBarSize={38} />
                    {yAxis2 && <Bar dataKey={yAxis2} fill={C.purple} radius={[5, 5, 0, 0]} maxBarSize={38} />}
                  </BarChart>
                </ResponsiveContainer>
              </ChartCard>
              <ChartCard title={`${yAxis} Trend`} subtitle="Area chart" accent={C.purple}>
                <ResponsiveContainer width="100%" height={270}>
                  <AreaChart data={chartData} margin={{ top: 0, right: 10, left: 0, bottom: 50 }}>
                    <defs>
                      <linearGradient id="gCyan" x1="0" y1="0" x2="0" y2="1">
                        <stop offset="5%" stopColor={C.cyan} stopOpacity={0.35} />
                        <stop offset="95%" stopColor={C.cyan} stopOpacity={0} />
                      </linearGradient>
                      <linearGradient id="gPurple" x1="0" y1="0" x2="0" y2="1">
                        <stop offset="5%" stopColor={C.purple} stopOpacity={0.3} />
                        <stop offset="95%" stopColor={C.purple} stopOpacity={0} />
                      </linearGradient>
                    </defs>
                    <CartesianGrid strokeDasharray="3 3" stroke={C.border} />
                    <XAxis dataKey="name" stroke={C.muted} tick={{ fontSize: 10, fill: C.muted }} angle={-38} textAnchor="end" interval={0} />
                    <YAxis stroke={C.muted} tick={{ fontSize: 11, fill: C.muted }} tickFormatter={fmtNum} />
                    <Tooltip contentStyle={TOOLTIP_STYLE} formatter={v => fmtNum(v)} />
                    <Area type="monotone" dataKey={yAxis} stroke={C.cyan} fill="url(#gCyan)" strokeWidth={2.5} dot={false} />
                    {yAxis2 && <Area type="monotone" dataKey={yAxis2} stroke={C.purple} fill="url(#gPurple)" strokeWidth={2} dot={false} />}
                  </AreaChart>
                </ResponsiveContainer>
              </ChartCard>
              {pieData.length > 0 && (
                <ChartCard title={`${strCols[0]} Distribution`} subtitle="Category breakdown" accent={C.green}>
                  <ResponsiveContainer width="100%" height={270}>
                    <PieChart>
                      <Pie data={pieData} dataKey="value" nameKey="name" cx="50%" cy="48%" innerRadius={60} outerRadius={100} paddingAngle={3}>
                        {pieData.map((_, i) => <Cell key={i} fill={PIE_COLORS[i % PIE_COLORS.length]} />)}
                      </Pie>
                      <Tooltip contentStyle={TOOLTIP_STYLE} />
                      <Legend iconType="circle" iconSize={8} formatter={v => <span style={{ color: C.textMid, fontSize: 12 }}>{v}</span>} />
                    </PieChart>
                  </ResponsiveContainer>
                </ChartCard>
              )}
              {scatterData.length > 1 && numCols.length >= 2 && (
                <ChartCard title={`${numCols[0]} vs ${numCols[1]}`} subtitle="Correlation scatter" accent={C.amber}>
                  <ResponsiveContainer width="100%" height={270}>
                    <ScatterChart margin={{ top: 0, right: 10, left: 0, bottom: 0 }}>
                      <CartesianGrid strokeDasharray="3 3" stroke={C.border} />
                      <XAxis dataKey="x" stroke={C.muted} tick={{ fontSize: 11, fill: C.muted }} tickFormatter={fmtNum} />
                      <YAxis dataKey="y" stroke={C.muted} tick={{ fontSize: 11, fill: C.muted }} tickFormatter={fmtNum} />
                      <Tooltip contentStyle={TOOLTIP_STYLE} formatter={v => fmtNum(v)} />
                      <Scatter data={scatterData} fill={C.amber} opacity={0.65} />
                    </ScatterChart>
                  </ResponsiveContainer>
                </ChartCard>
              )}
            </div>
          </div>
        )}

        {activeTab === "charts" && (
          <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 20, animation: "fadeIn 0.3s ease" }}>
            <ChartCard title="Top 10 — Horizontal Bar" subtitle={`Ranked by ${yAxis}`} accent={C.rose}>
              <ResponsiveContainer width="100%" height={340}>
                <BarChart data={[...chartData].slice(0, 10)} layout="vertical" margin={{ top: 0, right: 20, left: 0, bottom: 0 }}>
                  <CartesianGrid strokeDasharray="3 3" stroke={C.border} horizontal={false} />
                  <XAxis type="number" stroke={C.muted} tick={{ fontSize: 11, fill: C.muted }} tickFormatter={fmtNum} />
                  <YAxis type="category" dataKey="name" stroke={C.muted} tick={{ fontSize: 11, fill: C.muted }} width={110} />
                  <Tooltip contentStyle={TOOLTIP_STYLE} formatter={v => fmtNum(v)} />
                  <Bar dataKey={yAxis} fill={C.rose} radius={[0, 5, 5, 0]} maxBarSize={26} />
                </BarChart>
              </ResponsiveContainer>
            </ChartCard>
            <ChartCard title="Multi-metric Line" subtitle="Compare two metrics" accent={C.blue}>
              <ResponsiveContainer width="100%" height={340}>
                <LineChart data={chartData} margin={{ top: 0, right: 10, left: 0, bottom: 50 }}>
                  <CartesianGrid strokeDasharray="3 3" stroke={C.border} />
                  <XAxis dataKey="name" stroke={C.muted} tick={{ fontSize: 10, fill: C.muted }} angle={-38} textAnchor="end" interval={0} />
                  <YAxis stroke={C.muted} tick={{ fontSize: 11, fill: C.muted }} tickFormatter={fmtNum} />
                  <Tooltip contentStyle={TOOLTIP_STYLE} formatter={v => fmtNum(v)} />
                  <Legend formatter={v => <span style={{ color: C.textMid, fontSize: 12 }}>{v}</span>} />
                  <Line type="monotone" dataKey={yAxis} stroke={C.cyan} strokeWidth={2.5} dot={{ fill: C.cyan, r: 3 }} />
                  {yAxis2 && <Line type="monotone" dataKey={yAxis2} stroke={C.rose} strokeWidth={2.5} dot={{ fill: C.rose, r: 3 }} />}
                </LineChart>
              </ResponsiveContainer>
            </ChartCard>
            {radarData.length >= 3 && (
              <ChartCard title="Numeric Means — Radar" subtitle="Relative scale comparison" accent={C.purple}>
                <ResponsiveContainer width="100%" height={320}>
                  <RadarChart cx="50%" cy="50%" outerRadius={110} data={radarData}>
                    <PolarGrid stroke={C.border} />
                    <PolarAngleAxis dataKey="subject" tick={{ fontSize: 12, fill: C.muted }} />
                    <PolarRadiusAxis angle={30} tick={{ fontSize: 10, fill: C.muted }} />
                    <Radar name="Mean" dataKey="value" stroke={C.purple} fill={C.purple} fillOpacity={0.25} strokeWidth={2} />
                    <Tooltip contentStyle={TOOLTIP_STYLE} formatter={v => fmtNum(v)} />
                  </RadarChart>
                </ResponsiveContainer>
              </ChartCard>
            )}
            {pieData.length > 0 && (
              <ChartCard title="Category Share" subtitle="Full pie distribution" accent={C.green}>
                <ResponsiveContainer width="100%" height={320}>
                  <PieChart>
                    <Pie data={pieData} dataKey="value" nameKey="name" cx="50%" cy="48%" outerRadius={110} paddingAngle={2}>
                      {pieData.map((_, i) => <Cell key={i} fill={PIE_COLORS[i % PIE_COLORS.length]} />)}
                    </Pie>
                    <Tooltip contentStyle={TOOLTIP_STYLE} />
                    <Legend iconType="square" iconSize={10} formatter={v => <span style={{ color: C.textMid, fontSize: 12 }}>{v}</span>} />
                  </PieChart>
                </ResponsiveContainer>
              </ChartCard>
            )}
          </div>
        )}

        {activeTab === "stats" && (
          <div style={{ display: "flex", flexDirection: "column", gap: 20, animation: "fadeIn 0.3s ease" }}>
            <ChartCard title="Descriptive Statistics" subtitle="All numeric columns">
              <div style={{ overflowX: "auto" }}>
                <table style={{ width: "100%", borderCollapse: "collapse", fontSize: 13 }}>
                  <thead>
                    <tr>
                      {["Column", "Count", "Sum", "Mean", "Median", "Q1", "Q3", "Min", "Max", "Std Dev"].map(h => (
                        <th key={h} style={{ padding: "10px 14px", textAlign: "left", color: C.muted, fontWeight: 600, fontSize: 11, letterSpacing: "0.6px", textTransform: "uppercase", borderBottom: `1px solid ${C.border}`, whiteSpace: "nowrap" }}>{h}</th>
                      ))}
                    </tr>
                  </thead>
                  <tbody>
                    {numCols.map((col, i) => {
                      const s = computeStats(rows, col);
                      if (!s) return null;
                      return (
                        <tr key={col} style={{ background: i % 2 === 0 ? "transparent" : `${C.surface}60` }}>
                          <td style={{ padding: "10px 14px", color: C.cyan, fontWeight: 600, whiteSpace: "nowrap" }}>{col}</td>
                          {[s.count, s.sum, s.mean, s.median, s.q1, s.q3, s.min, s.max, s.std].map((v, j) => (
                            <td key={j} style={{ padding: "10px 14px", color: C.text }}>{fmtNum(v)}</td>
                          ))}
                        </tr>
                      );
                    })}
                  </tbody>
                </table>
              </div>
            </ChartCard>
            {strCols.length > 0 && (
              <ChartCard title="Categorical Summary" subtitle="Unique values and top frequency">
                <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fill, minmax(220px, 1fr))", gap: 14, marginTop: 4 }}>
                  {strCols.map(col => {
                    const vals = rows.map(r => String(r[col] || ""));
                    const unique = new Set(vals).size;
                    const freq = {};
                    vals.forEach(v => freq[v] = (freq[v] || 0) + 1);
                    const top = Object.entries(freq).sort((a, b) => b[1] - a[1])[0];
                    return (
                      <div key={col} style={{ background: C.surface, borderRadius: 12, padding: "14px 16px", border: `1px solid ${C.border}` }}>
                        <p style={{ color: C.purple, fontSize: 13, fontWeight: 600, marginBottom: 8, overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap" }}>{col}</p>
                        <div style={{ display: "flex", justifyContent: "space-between", marginBottom: 4 }}>
                          <span style={{ color: C.muted, fontSize: 12 }}>Unique values</span>
                          <span style={{ color: C.text, fontSize: 12, fontWeight: 600 }}>{unique.toLocaleString()}</span>
                        </div>
                        <div style={{ display: "flex", justifyContent: "space-between" }}>
                          <span style={{ color: C.muted, fontSize: 12 }}>Most frequent</span>
                          <span style={{ color: C.amber, fontSize: 12, fontWeight: 600, maxWidth: 110, overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap" }}>{top?.[0] || "—"}</span>
                        </div>
                        <div style={{ marginTop: 8, background: C.border, borderRadius: 4, height: 4 }}>
                          <div style={{ height: 4, borderRadius: 4, background: C.purple, width: `${Math.min((top?.[1] || 0) / rows.length * 100, 100)}%` }} />
                        </div>
                        <p style={{ color: C.muted, fontSize: 11, marginTop: 4 }}>{((top?.[1] || 0) / rows.length * 100).toFixed(1)}% of rows</p>
                      </div>
                    );
                  })}
                </div>
              </ChartCard>
            )}
          </div>
        )}

        {activeTab === "data" && (
          <div style={{ animation: "fadeIn 0.3s ease" }}>
            <div style={{ marginBottom: 16, display: "flex", alignItems: "center", gap: 12 }}>
              <div style={{ position: "relative", flex: 1, maxWidth: 380 }}>
                <input value={searchQuery} onChange={e => { setSearchQuery(e.target.value); setPage(0); }}
                  placeholder="Search across all columns…"
                  style={{ width: "100%", background: C.card, border: `1px solid ${C.border}`, color: C.text, borderRadius: 10, padding: "8px 12px 8px 36px", fontSize: 13, outline: "none" }} />
                <span style={{ position: "absolute", left: 11, top: "50%", transform: "translateY(-50%)", color: C.muted, fontSize: 14 }}>🔍</span>
              </div>
              <span style={{ color: C.muted, fontSize: 13 }}>{filteredRows.length.toLocaleString()} rows</span>
            </div>
            <div style={{ background: C.card,borderRadius: 16, border: `1px solid ${C.border}`, overflow: "hidden" }}>
              <div style={{ overflowX: "auto" }}>
                <table style={{ width: "100%", borderCollapse: "collapse", fontSize: 13, minWidth: 600 }}>
                  <thead>
                    <tr style={{ background: C.surface }}>
                      {allCols.map(col => (
                        <th key={col} onClick={() => { setSortCol(col); setSortDir(sortCol === col && sortDir === "asc" ? "desc" : "asc"); setPage(0); }}
                          style={{ padding: "12px 16px", textAlign: "left", color: sortCol === col ? C.cyan : C.muted, fontWeight: 600, fontSize: 11, letterSpacing: "0.6px", textTransform: "uppercase", borderBottom: `1px solid ${C.border}`, whiteSpace: "nowrap", cursor: "pointer", userSelect: "none" }}>
                          {col}
                          <span style={{ marginLeft: 4, color: colTypes[col] === "number" ? C.cyan : C.purple, fontSize: 9 }}>{colTypes[col] === "number" ? "123" : "Abc"}</span>
                          {sortCol === col && <span style={{ marginLeft: 4, fontSize: 10 }}>{sortDir === "asc" ? "▲" : "▼"}</span>}
                        </th>
                      ))}
                    </tr>
                  </thead>
                  <tbody>
                    {tableRows.map((row, i) => (
                      <tr key={i} style={{ background: i % 2 === 0 ? "transparent" : `${C.surface}50`, transition: "background 0.1s" }}
                        onMouseEnter={e => e.currentTarget.style.background = `${C.cyan}06`}
                        onMouseLeave={e => e.currentTarget.style.background = i % 2 === 0 ? "transparent" : `${C.surface}50`}>
                        {allCols.map(col => (
                          <td key={col} style={{ padding: "9px 16px", color: C.text, borderBottom: `1px solid ${C.border}30`, maxWidth: 200, overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap" }}>
                            {String(row[col] ?? "")}
                          </td>
                        ))}
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>
              <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between", padding: "12px 20px", borderTop: `1px solid ${C.border}` }}>
                <span style={{ color: C.muted, fontSize: 13 }}>Showing {page * PAGE_SIZE + 1}–{Math.min((page + 1) * PAGE_SIZE, sortedRows.length)} of {sortedRows.length.toLocaleString()}</span>
                <div style={{ display: "flex", gap: 8, alignItems: "center" }}>
                  <button onClick={() => setPage(0)} disabled={page === 0} style={{ padding: "5px 12px", borderRadius: 8, border: `1px solid ${C.border}`, background: page > 0 ? C.surface : "transparent", color: page > 0 ? C.text : C.muted, fontSize: 12, cursor: page > 0 ? "pointer" : "default" }}>«</button>
                  <button onClick={() => setPage(p => p - 1)} disabled={page === 0} style={{ padding: "5px 14px", borderRadius: 8, border: `1px solid ${C.border}`, background: page > 0 ? C.surface : "transparent", color: page > 0 ? C.text : C.muted, fontSize: 13, cursor: page > 0 ? "pointer" : "default" }}>← Prev</button>
                  <span style={{ color: C.muted, fontSize: 13 }}>{page + 1} / {totalPages}</span>
                  <button onClick={() => setPage(p => p + 1)} disabled={page >= totalPages - 1} style={{ padding: "5px 14px", borderRadius: 8, border: `1px solid ${C.border}`, background: page < totalPages - 1 ? C.surface : "transparent", color: page < totalPages - 1 ? C.text : C.muted, fontSize: 13, cursor: page < totalPages - 1 ? "pointer" : "default" }}>Next →</button>
                  <button onClick={() => setPage(totalPages - 1)} disabled={page >= totalPages - 1} style={{ padding: "5px 12px", borderRadius: 8, border: `1px solid ${C.border}`, background: page < totalPages - 1 ? C.surface : "transparent", color: page < totalPages - 1 ? C.text : C.muted, fontSize: 12, cursor: page < totalPages - 1 ? "pointer" : "default" }}>»</button>
                </div>
              </div>
            </div>
          </div>
        )}

        {activeTab === "ai" && (
          <div style={{ animation: "fadeIn 0.3s ease" }}>
            {!aiResult && !aiLoading && !aiError && (
              <div style={{ textAlign: "center", padding: "80px 24px", background: C.card, borderRadius: 20, border: `1px solid ${C.border}` }}>
                <div style={{ fontSize: 64, marginBottom: 24 }}>🤖</div>
                <h2 style={{ fontFamily: "'Space Grotesk', sans-serif", fontSize: 24, fontWeight: 700, marginBottom: 12 }}>AI-Powered Analysis</h2>
                <p style={{ color: C.textMid, fontSize: 15, marginBottom: 28, maxWidth: 480, margin: "0 auto 28px" }}>
                  Claude AI will analyze your dataset and generate insights, anomaly detection, findings, and recommendations.
                </p>
                <button onClick={runAI} style={{ padding: "12px 32px", borderRadius: 12, border: "none", background: `linear-gradient(135deg, ${C.cyan}, ${C.purple})`, color: "#fff", fontSize: 15, fontWeight: 700, cursor: "pointer" }}>
                  🚀 Run AI Analysis
                </button>
              </div>
            )}
            {aiLoading && (
              <div style={{ textAlign: "center", padding: "80px 24px" }}>
                <div style={{ width: 56, height: 56, border: `3px solid ${C.border}`, borderTop: `3px solid ${C.cyan}`, borderRadius: "50%", animation: "spin 0.8s linear infinite", margin: "0 auto 24px" }} />
                <p style={{ color: C.textMid, fontSize: 16, animation: "pulse 1.5s ease-in-out infinite" }}>Claude AI is analyzing your dataset…</p>
                <p style={{ color: C.muted, fontSize: 13, marginTop: 8 }}>Examining {rows.length.toLocaleString()} rows across {allCols.length} columns</p>
              </div>
            )}
            {aiError && (
              <div style={{ padding: 24, background: `${C.rose}12`, borderRadius: 16, border: `1px solid ${C.rose}30`, color: C.rose, textAlign: "center" }}>
                <p style={{ fontSize: 16, fontWeight: 600, marginBottom: 8 }}>Analysis Failed</p>
                <p style={{ fontSize: 13, color: C.textMid, marginBottom: 16 }}>{aiError}</p>
                <button onClick={runAI} style={{ padding: "8px 24px", borderRadius: 8, border: `1px solid ${C.rose}`, background: "transparent", color: C.rose, fontSize: 13, cursor: "pointer" }}>Try Again</button>
              </div>
              )}
            {aiResult && (
              <div style={{ display: "flex", flexDirection: "column", gap: 20 }}>
                <div style={{ background: `linear-gradient(135deg, ${C.cyan}10, ${C.purple}10)`, borderRadius: 16, padding: 24, border: `1px solid ${C.borderLight}` }}>
                  <div style={{ display: "flex", alignItems: "center", gap: 10, marginBottom: 12 }}>
                    <span style={{ fontSize: 22 }}>📝</span>
                    <h3 style={{ fontFamily: "'Space Grotesk', sans-serif", fontSize: 17, fontWeight: 700, color: C.cyan }}>Dataset Summary</h3>
                  </div>
                  <p style={{ color: C.text, fontSize: 15, lineHeight: 1.7 }}>{aiResult.summary}</p>
                </div>
                {aiResult.topFindings?.length > 0 && (
                  <div style={{ background: C.card, borderRadius: 16, padding: 24, border: `1px solid ${C.border}` }}>
                    <div style={{ display: "flex", alignItems: "center", gap: 10, marginBottom: 16 }}>
                      <span style={{ fontSize: 20 }}>🏆</span>
                      <h3 style={{ fontFamily: "'Space Grotesk', sans-serif", fontSize: 16, fontWeight: 700 }}>Top Findings</h3>
                    </div>
                    <div style={{ display: "flex", flexDirection: "column", gap: 10 }}>
                      {aiResult.topFindings.map((f, i) => (
                        <div key={i} style={{ display: "flex", alignItems: "flex-start", gap: 12 }}>
                          <span style={{ color: C.cyan, fontWeight: 700, fontSize: 14, minWidth: 24, fontFamily: "'Space Grotesk', sans-serif" }}>0{i + 1}</span>
                          <p style={{ color: C.textMid, fontSize: 14, lineHeight: 1.6 }}>{f}</p>
                        </div>
                      ))}
                    </div>
                  </div>
                )}
                {aiResult.insights?.length > 0 && (
                  <div>
                    <h3 style={{ fontFamily: "'Space Grotesk', sans-serif", fontSize: 16, fontWeight: 700, marginBottom: 14, display: "flex", alignItems: "center", gap: 8 }}>
                      <span>💡</span> Key Insights
                    </h3>
                    <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fill, minmax(300px, 1fr))", gap: 14 }}>
                      {aiResult.insights.map((ins, i) => {
                        const typeColor = { positive: C.green, warning: C.amber, info: C.cyan, critical: C.rose }[ins.type] || C.cyan;
                        return (
                          <div key={i} style={{ background: C.card, borderRadius: 14, padding: "18px 20px", border: `1px solid ${C.border}`, borderLeft: `3px solid ${typeColor}` }}>
                            <div style={{ display: "flex", alignItems: "flex-start", justifyContent: "space-between", marginBottom: 10, gap: 8 }}>
                              <p style={{ color: C.text, fontSize: 14, fontWeight: 600 }}>{ins.title}</p>
                              <InsightBadge type={ins.type} />
                            </div>
                            <p style={{ color: C.textMid, fontSize: 13, lineHeight: 1.6 }}>{ins.description}</p>
                          </div>
                        );
                      })}
                    </div>
                  </div>
                )}
                <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 20 }}>
                  {aiResult.recommendations?.length > 0 && (
                    <div style={{ background: C.card, borderRadius: 16, padding: 22, border: `1px solid ${C.border}` }}>
                      <h3 style={{ fontFamily: "'Space Grotesk', sans-serif", fontSize: 15, fontWeight: 700, marginBottom: 16, display: "flex", alignItems: "center", gap: 8 }}>
                        <span>🎯</span> Recommendations
                      </h3>
                      <div style={{ display: "flex", flexDirection: "column", gap: 14 }}>
                        {aiResult.recommendations.map((r, i) => (
                          <div key={i} style={{ padding: "12px 14px", background: `${C.green}08`, borderRadius: 10, border: `1px solid ${C.green}20` }}>
                            <p style={{ color: C.green, fontSize: 13, fontWeight: 600, marginBottom: 4 }}>{r.title}</p>
                            <p style={{ color: C.textMid, fontSize: 13, lineHeight: 1.5 }}>{r.action}</p>
                          </div>
                        ))}
                      </div>
                    </div>
                  )}
                  {aiResult.anomalies?.length > 0 && (
                    <div style={{ background: C.card, borderRadius: 16, padding: 22, border: `1px solid ${C.border}` }}>
                      <h3 style={{ fontFamily: "'Space Grotesk', sans-serif", fontSize: 15, fontWeight: 700, marginBottom: 16, display: "flex", alignItems: "center", gap: 8 }}>
                        <span>⚠️</span> Anomalies Detected
                      </h3>
                      <div style={{ display: "flex", flexDirection: "column", gap: 14 }}>
                        {aiResult.anomalies.map((a, i) => (
                          <div key={i} style={{ padding: "12px 14px", background: `${C.amber}08`, borderRadius: 10, border: `1px solid ${C.amber}25` }}>
                            <p style={{ color: C.amber, fontSize: 13, fontWeight: 600, marginBottom: 4 }}>{a.column}</p>
                            <p style={{ color: C.textMid, fontSize: 13, lineHeight: 1.5 }}>{a.detail}</p>
                          </div>
                        ))}
                      </div>
                    </div>
                  )}
                </div>
                <div style={{ textAlign: "center" }}>
                  <button onClick={runAI} style={{ padding: "9px 28px", borderRadius: 10, border: `1px solid ${C.border}`, background: C.surface, color: C.textMid, fontSize: 13, fontWeight: 600, cursor: "pointer" }}>
                    🔄 Re-run Analysis
                  </button>
                </div>
              </div>
            )}
          </div>
        )}
      </main>

      <footer style={{ background: C.surface, borderTop: `1px solid ${C.border}`, padding: "24px 28px", marginTop: 40 }}>
        <div style={{ maxWidth: 1440, margin: "0 auto", display: "flex", alignItems: "center", justifyContent: "space-between", flexWrap: "wrap", gap: 16 }}>
          <div style={{ display: "flex", alignItems: "center", gap: 12 }}>
            <svg width="24" height="24" viewBox="0 0 48 48" fill="none">
              <rect width="48" height="48" rx="12" fill={C.cyan} fillOpacity="0.15" />
              <rect x="8" y="26" width="7" height="16" rx="2.5" fill={C.cyan} />
              <rect x="20" y="17" width="7" height="25" rx="2.5" fill={C.purple} />
              <rect x="32" y="9" width="7" height="33" rx="2.5" fill={C.green} />
            </svg>
            <span style={{ fontFamily: "'Space Grotesk', sans-serif", fontWeight: 700, fontSize: 16, color: C.text }}>
              Insight<span style={{ color: C.cyan }}>Flow</span>
            </span>
          </div>
          <div style={{ textAlign: "center" }}>
            <p style={{ color: C.text, fontSize: 13, fontWeight: 600, marginBottom: 4 }}>
              © 2025 <span style={{ color: C.cyan }}>Navesh R</span> · All Rights Reserved
            </p>
            <p style={{ color: C.muted, fontSize: 12, lineHeight: 1.6 }}>
              Bachelor of Technology in Artificial Intelligence &amp; Data Science<br />
              Vel Tech High Tech Dr. Rangarajan Dr. Sakunthala Engineering College, Chennai, India
            </p>
          </div>
          <div style={{ display: "flex", gap: 16, fontSize: 12, color: C.muted }}>
            <span>React + Recharts</span>
            <span style={{ color: C.border }}>|</span>
            <span>Powered by Claude AI</span>
            <span style={{ color: C.border }}>|</span>
            <span>v2.0.0</span>
          </div>
        </div>
        <div style={{ maxWidth: 1440, margin: "16px auto 0", paddingTop: 16, borderTop: `1px solid ${C.border}30`, textAlign: "center" }}>
          <p style={{ color: C.muted, fontSize: 11 }}>
            Unauthorized reproduction or distribution of this software, in whole or in part, is strictly prohibited.
            Intellectual property of Navesh R · B.Tech AI &amp; DS · Vel Tech High Tech Dr. Rangarajan Dr. Sakunthala Engineering College.
          </p>
        </div>
      </footer>
    </div>
  );
}

export default function DataDashboard() {
  const [data, setData] = useState(null);
  const [fileName, setFileName] = useState("");
  if (!data) return <UploadScreen onData={(rows, name) => { setData(rows); setFileName(name); }} />;
  return <Dashboard rows={data} fileName={fileName} onReset={() => { setData(null); setFileName(""); }} />;
                }
