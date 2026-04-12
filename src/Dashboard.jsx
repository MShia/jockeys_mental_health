import { useState, useEffect, useMemo, useCallback, useRef } from "react";
import * as Papa from "papaparse";
import { BarChart, Bar, XAxis, YAxis, CartesianGrid, Tooltip, Legend, ResponsiveContainer, PieChart, Pie, Cell, RadarChart, Radar, PolarGrid, PolarAngleAxis, PolarRadiusAxis, ScatterChart, Scatter, LineChart, Line } from "recharts";
import { Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell, HeadingLevel, AlignmentType, BorderStyle, WidthType, ShadingType, PageBreak } from "docx";
import { saveAs } from "file-saver";

// ─── CONSTANTS ───
const GENDER_MAP = { 1: "Male", 2: "Female", 3: "Prefer not to say" };
const EDUCATION_MAP = { 1: "Junior Cert", 2: "Leaving Cert", 3: "University" };
const LOCATION_MAP = { 1: "Ireland", 2: "Great Britain" };
const LICENCE_MAP = { 1: "Amateur", 2: "Apprentice", 3: "Claiming Professional", 4:"Fully Licensed Flat", 5:"Fully Licensed Professional NH" };
const WEIGHT_FREQ_MAP = { 1: "Daily", 2: "4-5/week", 3: "2-3/week", 4: "Once/week", 5: "Once/month", 6: "Never" };
const WEIGHT_TF_MAP = { 1: "<1 day", 2: "1-2 days", 3: "3-4 days", 4: "5-7 days", 5: "N/A" };
const HOLIDAY_MAP = {1: "0-7 days", 2: "7-14 days", 3: "14-21 days", 4: "21 -28 days", 5: ">28 days"};
const WEEKLY_WORK_HRS = {1: "30 -40", 2: "40 - 50", 3: "50 - 60", 4: ">60"};
const RACE_DAYS_WEEK = {1: "0-1", 2: "2-3", 3: "4-5", 4: "6-7" }
const WEEKLY_TRVL_HRS ={1: "0-5", 2:"5-10" , 3:"10-15,", 4:"15-20", 5:">20"}


const WM_METHODS = [
  { key: "wm_food_restriction", label: "Food Restriction" },
  { key: "wm_fluid_restriction", label: "Fluid Restriction" },
  { key: "wm_sweat_suit", label: "Sweat Suit" },
  { key: "wm_sweat_car", label: "Sweating in Car" },
  { key: "wm_sauna", label: "Sauna" },
  { key: "wm_steam_room", label: "Steam Room" },
  { key: "wm_hot_baths", label: "Hot Baths" },
  { key: "wm_excessive_exercise", label: "Excessive Exercise" },
  { key: "wm_smoking", label: "Smoking" },
  { key: "wm_diuretics", label: "Diuretics" },
  { key: "wm_laxatives", label: "Laxatives" },
  { key: "wm_flipping", label: "Flipping" },
  { key: "wm_ice_baths", label: "Ice Baths" },
  { key: "wm_glycogen_burn", label: "Glycogen Burn" },
  { key: "wm_weight_loss_injections", label: "Weight Loss Injections" },
];

const CMD_SCALES = [
  { key: "cesd_total", label: "CES-D", cutoff: 16, binaryKey: "cesd_depressive_risk", range: [0, 60], desc: "Depression (CES-D)" },
  { key: "phq9_total", label: "PHQ-9", cutoff: 10, range: [0, 27], desc: "Depression (PHQ-9)" },
  { key: "gad7_total", label: "GAD-7", cutoff: 10, range: [0, 21], desc: "Anxiety (GAD-7)" },
  { key: "k10_total", label: "K10", cutoff: 20, range: [10, 50], desc: "Psychological Distress (K10)" },
  { key: "wemwbs_total", label: "WEMWBS", cutoff: 40, range: [14, 70], desc: "Mental Wellbeing (WEMWBS)", invert: true },
];

const PHQ9_SEVERITY = [
  { min: 0, max: 4, label: "Minimal", color: "#22c55e" },
  { min: 5, max: 9, label: "Mild", color: "#eab308" },
  { min: 10, max: 14, label: "Moderate", color: "#f97316" },
  { min: 15, max: 19, label: "Mod. Severe", color: "#ef4444" },
  { min: 20, max: 27, label: "Severe", color: "#991b1b" },
];
const GAD7_SEVERITY = [
  { min: 0, max: 4, label: "Minimal", color: "#22c55e" },
  { min: 5, max: 9, label: "Mild", color: "#eab308" },
  { min: 10, max: 14, label: "Moderate", color: "#f97316" },
  { min: 15, max: 21, label: "Severe", color: "#ef4444" },
];
const K10_SEVERITY = [
  { min: 10, max: 15, label: "Low", color: "#22c55e" },
  { min: 16, max: 21, label: "Moderate", color: "#eab308" },
  { min: 22, max: 29, label: "High", color: "#f97316" },
  { min: 30, max: 50, label: "Very High", color: "#ef4444" },
];

const COPE_SUBSCALES = [
  { key: "cope_self_distraction_score", label: "Self-Distraction" },
  { key: "cope_active_coping_score", label: "Active Coping" },
  { key: "cope_denial_score", label: "Denial" },
  { key: "cope_substance_use_score", label: "Substance Use" },
  { key: "cope_emotional_support_score", label: "Emotional Support" },
  { key: "cope_behavioural_disengagement_score", label: "Behavioural Disengagement" },
  { key: "cope_venting_score", label: "Venting" },
  { key: "cope_instrumental_support_score", label: "Instrumental Support" },
  { key: "cope_positive_reframing_score", label: "Positive Reframing" },
  { key: "cope_planning_score", label: "Planning" },
  { key: "cope_humour_score", label: "Humour" },
  { key: "cope_acceptance_score", label: "Acceptance" },
  { key: "cope_religion_score", label: "Religion" },
  { key: "cope_self_blame_score", label: "Self-Blame" },

  // { key: "cope_problem_focused", label: "Problem-Focused Coping" },
  // { key: "cope_emotion_focused", label: "Emotion-Focused Coping" },
  // { key: "cope_avoidant", label: "Avoidant Coping" },
];

const COPE_3CLASS = [
  { key: "cope_problem_focused", label: "Problem-Focused Coping" },
  { key: "cope_emotion_focused", label: "Emotion-Focused Coping" },
  { key: "cope_avoidant", label: "Avoidant Coping" },
];

const PASSQ_SUBS = [
  { key: "passq_emotional_support", label: "Emotional Support" },
  { key: "passq_esteem_support", label: "Esteem Support" },
  { key: "passq_informational_support", label: "Informational Support" },
  { key: "passq_tangible_support", label: "Tangible Support" },
];

const RISK_FACTOR_GROUPS = [
  {
    label: "SARRS Life Events",
    vars: [
      { key: "sarrs_total_stress_score", label: "SARRS Total Stress", type: "continuous" },
      { key: "sarrs_n_events", label: "SARRS N Events", type: "continuous" },
    ],
  },
  {
    label: "Sense of Agency",
    vars: [
      { key: "soa_positive_agency_total", label: "Positive Agency", type: "continuous" },
      { key: "soa_negative_agency_total", label: "Negative Agency", type: "continuous" },
    ],
  },
  {
    label: "Career Satisfaction",
    vars: [{ key: "css_total", label: "CSS Total", type: "continuous" }],
  },
  {
    label: "Social Support (PASS-Q)",
    vars: [
      { key: "passq_total", label: "PASS-Q Total", type: "continuous" },
      ...PASSQ_SUBS.map((s) => ({ ...s, type: "continuous" })),
    ],
  },
  {
    label: "Coping (Brief-COPE)",
    vars: null, // dynamically set via copeMode state
  },
  {
    label: "Sleep & Work",
    vars: [
      { key: "sleep_hours", label: "Sleep Hours", type: "continuous" },
      { key: "weekly_work_hours", label: "Weekly Work Hours", type: "categorical", cats: WEEKLY_WORK_HRS },
      { key: "weekly_travel_hours", label: "Weekly Travel Hours", type: "categorical", cats:WEEKLY_TRVL_HRS },
    ],
  },
  {
    label: "Weight Making",
    vars: [
      { key: "ever_cut_weight", label: "Ever Cut Weight", type: "binary" },
      { key: "max_weight_lost_prep_lbs", label: "Max Weight Lost (lbs)", type: "continuous" },
    ],
  },
  {
    label: "Demographics",
    vars: [
      { key: "age", label: "Age", type: "continuous" },
      { key: "gender", label: "Gender (ref=Male)", type: "categorical", cats: GENDER_MAP },
      { key: "education_level", label: "Education", type: "categorical", cats: EDUCATION_MAP },
      { key: "years_licensed", label: "Years Licensed", type: "continuous" },
    ],
  },
];

const PALETTE = ["#0f766e", "#0e7490", "#7c3aed", "#db2777", "#ea580c", "#ca8a04", "#16a34a", "#6366f1", "#e11d48", "#0d9488", "#8b5cf6", "#f59e0b", "#3b82f6", "#ef4444"];

// ─── STATISTICS HELPERS ───
const num = (v) => { const n = parseFloat(v); return isNaN(n) ? null : n; };
const validNums = (arr) => arr.map(num).filter((v) => v !== null);
const mean = (arr) => { const v = validNums(arr); return v.length ? v.reduce((a, b) => a + b, 0) / v.length : null; };
const sd = (arr) => {
  const v = validNums(arr);
  if (v.length < 2) return null;
  const m = v.reduce((a, b) => a + b, 0) / v.length;
  return Math.sqrt(v.reduce((a, b) => a + (b - m) ** 2, 0) / (v.length - 1));
};
const median = (arr) => {
  const v = validNums(arr).sort((a, b) => a - b);
  if (!v.length) return null;
  const mid = Math.floor(v.length / 2);
  return v.length % 2 ? v[mid] : (v[mid - 1] + v[mid]) / 2;
};
const iqr = (arr) => {
  const v = validNums(arr).sort((a, b) => a - b);
  if (v.length < 4) return [null, null];
  const q1i = Math.floor(v.length * 0.25), q3i = Math.floor(v.length * 0.75);
  return [v[q1i], v[q3i]];
};

// Simple OLS linear regression
function linearRegression(y, X, varNames) {
  // X is array of arrays (each row = observation), y is array
  const n = y.length, p = X[0].length;
  if (n <= p + 1) return null;
  // Add intercept
  const Xa = X.map((r) => [1, ...r]);
  const k = Xa[0].length;
  // X'X
  const XtX = Array.from({ length: k }, (_, i) => Array.from({ length: k }, (_, j) => Xa.reduce((s, r) => s + r[i] * r[j], 0)));
  // Invert XtX (Gauss-Jordan)
  const inv = invertMatrix(XtX);
  if (!inv) return null;
  // X'y
  const Xty = Array.from({ length: k }, (_, i) => Xa.reduce((s, r, idx) => s + r[i] * y[idx], 0));
  // beta = inv(X'X) * X'y
  const beta = inv.map((row) => row.reduce((s, v, j) => s + v * Xty[j], 0));
  // residuals
  const yhat = Xa.map((r) => r.reduce((s, v, j) => s + v * beta[j], 0));
  const resid = y.map((v, i) => v - yhat[i]);
  const sse = resid.reduce((s, r) => s + r * r, 0);
  const mse = sse / (n - k);
  // TSS
  const ym = y.reduce((a, b) => a + b, 0) / n;
  const sst = y.reduce((s, v) => s + (v - ym) ** 2, 0);
  const r2 = 1 - sse / sst;
  const adjR2 = 1 - (1 - r2) * (n - 1) / (n - k);
  // SE and p-values
  const se = inv.map((row, i) => Math.sqrt(Math.max(0, mse * row[i])));
  const tstat = beta.map((b, i) => (se[i] > 0 ? b / se[i] : 0));
  const pvals = tstat.map((t) => 2 * (1 - tCDF(Math.abs(t), n - k)));
  const names = ["Intercept", ...varNames];
  return { coefficients: beta, se, tstat, pvals, r2, adjR2, n, names, mse };
}

// Logistic regression via IRLS
function logisticRegression(y, X, varNames, maxIter = 25) {
  const n = y.length, p = X[0].length;
  if (n <= p + 1) return null;
  const Xa = X.map((r) => [1, ...r]);
  const k = Xa[0].length;
  let beta = new Array(k).fill(0);
  
  for (let iter = 0; iter < maxIter; iter++) {
    const eta = Xa.map((r) => r.reduce((s, v, j) => s + v * beta[j], 0));
    const mu = eta.map((e) => 1 / (1 + Math.exp(-Math.max(-500, Math.min(500, e)))));
    const W = mu.map((m) => Math.max(1e-10, m * (1 - m)));
    // X'WX
    const XtWX = Array.from({ length: k }, (_, i) =>
      Array.from({ length: k }, (_, j) => Xa.reduce((s, r, idx) => s + r[i] * W[idx] * r[j], 0))
    );
    const inv = invertMatrix(XtWX);
    if (!inv) return null;
    // X'(y - mu)
    const Xtr = Array.from({ length: k }, (_, i) => Xa.reduce((s, r, idx) => s + r[i] * (y[idx] - mu[idx]), 0));
    const delta = inv.map((row) => row.reduce((s, v, j) => s + v * Xtr[j], 0));
    beta = beta.map((b, i) => b + delta[i]);
    if (delta.every((d) => Math.abs(d) < 1e-8)) break;
  }
  // Final stats
  const eta = Xa.map((r) => r.reduce((s, v, j) => s + v * beta[j], 0));
  const mu = eta.map((e) => 1 / (1 + Math.exp(-Math.max(-500, Math.min(500, e)))));
  const W = mu.map((m) => Math.max(1e-10, m * (1 - m)));
  const XtWX = Array.from({ length: k }, (_, i) =>
    Array.from({ length: k }, (_, j) => Xa.reduce((s, r, idx) => s + r[i] * W[idx] * r[j], 0))
  );
  const inv = invertMatrix(XtWX);
  if (!inv) return null;
  const se = inv.map((row, i) => Math.sqrt(Math.max(0, row[i])));
  const z = beta.map((b, i) => (se[i] > 0 ? b / se[i] : 0));
  const pvals = z.map((zv) => 2 * (1 - normalCDF(Math.abs(zv))));
  const or = beta.map((b) => Math.exp(b));
  const ciLow = beta.map((b, i) => Math.exp(b - 1.96 * se[i]));
  const ciHigh = beta.map((b, i) => Math.exp(b + 1.96 * se[i]));
  // Log-likelihood
  const ll = y.reduce((s, yi, i) => s + (yi * Math.log(mu[i] + 1e-15) + (1 - yi) * Math.log(1 - mu[i] + 1e-15)), 0);
  const ll0 = (() => { const p0 = y.reduce((a, b) => a + b, 0) / n; return n * (p0 * Math.log(p0 + 1e-15) + (1 - p0) * Math.log(1 - p0 + 1e-15)); })();
  const pseudoR2 = 1 - ll / ll0;
  const names = ["Intercept", ...varNames];
  return { coefficients: beta, se, z, pvals, or, ciLow, ciHigh, n, names, pseudoR2, ll };
}

function invertMatrix(m) {
  const n = m.length;
  const aug = m.map((r, i) => [...r.map(Number), ...Array.from({ length: n }, (_, j) => (i === j ? 1 : 0))]);
  for (let i = 0; i < n; i++) {
    let maxR = i;
    for (let r = i + 1; r < n; r++) if (Math.abs(aug[r][i]) > Math.abs(aug[maxR][i])) maxR = r;
    [aug[i], aug[maxR]] = [aug[maxR], aug[i]];
    if (Math.abs(aug[i][i]) < 1e-12) return null;
    const piv = aug[i][i];
    for (let j = 0; j < 2 * n; j++) aug[i][j] /= piv;
    for (let r = 0; r < n; r++) {
      if (r === i) continue;
      const f = aug[r][i];
      for (let j = 0; j < 2 * n; j++) aug[r][j] -= f * aug[i][j];
    }
  }
  return aug.map((r) => r.slice(n));
}

function normalCDF(x) {
  const a1 = 0.254829592, a2 = -0.284496736, a3 = 1.421413741, a4 = -1.453152027, a5 = 1.061405429, p = 0.3275911;
  const sign = x < 0 ? -1 : 1;
  x = Math.abs(x) / Math.sqrt(2);
  const t = 1 / (1 + p * x);
  const y = 1 - ((((a5 * t + a4) * t + a3) * t + a2) * t + a1) * t * Math.exp(-x * x);
  return 0.5 * (1 + sign * y);
}

function tCDF(t, df) {
  // Approximation via regularized incomplete beta
  const x = df / (df + t * t);
  return 1 - 0.5 * incompleteBeta(df / 2, 0.5, x);
}

function incompleteBeta(a, b, x) {
  // Continued fraction approx
  if (x < 0 || x > 1) return 0;
  if (x === 0 || x === 1) return x;
  const lnBeta = lgamma(a) + lgamma(b) - lgamma(a + b);
  const front = Math.exp(Math.log(x) * a + Math.log(1 - x) * b - lnBeta) / a;
  // Lentz's continued fraction
  let f = 1, c = 1, d = 1 - (a + 1) * x / (a + 1);
  if (Math.abs(d) < 1e-30) d = 1e-30;
  d = 1 / d;
  f = d;
  for (let i = 1; i <= 200; i++) {
    const m = i;
    let an = m * (b - m) * x / ((a + 2 * m - 1) * (a + 2 * m));
    d = 1 + an * d; if (Math.abs(d) < 1e-30) d = 1e-30; d = 1 / d;
    c = 1 + an / c; if (Math.abs(c) < 1e-30) c = 1e-30;
    f *= d * c;
    an = -(a + m) * (a + b + m) * x / ((a + 2 * m) * (a + 2 * m + 1));
    d = 1 + an * d; if (Math.abs(d) < 1e-30) d = 1e-30; d = 1 / d;
    c = 1 + an / c; if (Math.abs(c) < 1e-30) c = 1e-30;
    const delta = d * c;
    f *= delta;
    if (Math.abs(delta - 1) < 1e-10) break;
  }
  return front * f;
}

function lgamma(x) {
  const c = [76.18009172947146, -86.50532032941677, 24.01409824083091, -1.231739572450155, 0.1208650973866179e-2, -0.5395239384953e-5];
  let y = x, tmp = x + 5.5;
  tmp -= (x + 0.5) * Math.log(tmp);
  let ser = 1.000000000190015;
  for (let j = 0; j < 6; j++) ser += c[j] / ++y;
  return -tmp + Math.log(2.5066282746310005 * ser / x);
}

const fmt = (v, d = 2) => (v === null || v === undefined || isNaN(v) ? "—" : Number(v).toFixed(d));
const fmtP = (p) => (p === null || p === undefined || isNaN(p) ? "—" : p < 0.001 ? "<.001" : p.toFixed(3));
const pStar = (p) => (p < 0.001 ? "***" : p < 0.01 ? "**" : p < 0.05 ? "*" : "");

// ─── CSV Export ───
function downloadCSV(rows, filename) {
  const csv = rows.map((r) => r.map((c) => `"${String(c ?? "").replace(/"/g, '""')}"`).join(",")).join("\n");
  const blob = new Blob([csv], { type: "text/csv" });
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a"); a.href = url; a.download = filename; a.click();
  URL.revokeObjectURL(url);
}

// ─── COMPONENTS ───
const Tabs = ({ tabs, active, onChange }) => (
  <div style={{ display: "flex", gap: 2, overflowX: "auto", borderBottom: "2px solid #1e293b", marginBottom: 20, paddingBottom: 0 }}>
    {tabs.map((t) => (
      <button key={t.id} onClick={() => onChange(t.id)}
        style={{
          padding: "10px 18px", fontSize: 13, fontWeight: active === t.id ? 700 : 500,
          background: active === t.id ? "#0f766e" : "transparent",
          color: active === t.id ? "#fff" : "#94a3b8",
          border: "none", borderRadius: "8px 8px 0 0", cursor: "pointer",
          whiteSpace: "nowrap", transition: "all .2s",
          borderBottom: active === t.id ? "2px solid #0f766e" : "2px solid transparent",
        }}
      >{t.label}</button>
    ))}
  </div>
);

const Card = ({ title, children, style }) => (
  <div style={{ background: "#1e293b", borderRadius: 12, padding: 20, marginBottom: 16, border: "1px solid #334155", ...style }}>
    {title && <h3 style={{ margin: "0 0 14px", fontSize: 15, color: "#5eead4", fontWeight: 600 }}>{title}</h3>}
    {children}
  </div>
);

const Stat = ({ label, value, sub }) => (
  <div style={{ textAlign: "center", padding: "12px 8px" }}>
    <div style={{ fontSize: 24, fontWeight: 700, color: "#f0fdfa" }}>{value}</div>
    <div style={{ fontSize: 12, color: "#94a3b8", marginTop: 4 }}>{label}</div>
    {sub && <div style={{ fontSize: 11, color: "#64748b", marginTop: 2 }}>{sub}</div>}
  </div>
);

const ExportBtn = ({ onClick, label = "Export CSV" }) => (
  <button onClick={onClick} style={{ padding: "6px 14px", fontSize: 12, background: "#0f766e", color: "#fff", border: "none", borderRadius: 6, cursor: "pointer", fontWeight: 600 }}>{label}</button>
);

const CheckList = ({ items, selected, onChange, columns = 2 }) => (
  <div style={{ display: "grid", gridTemplateColumns: `repeat(${columns}, 1fr)`, gap: "4px 12px", maxHeight: 300, overflowY: "auto", fontSize: 12 }}>
    {items.map((it) => (
      <label key={it.key} style={{ display: "flex", alignItems: "center", gap: 6, cursor: "pointer", padding: "3px 0", color: "#cbd5e1" }}>
        <input type="checkbox" checked={selected.includes(it.key)} onChange={() => {
          onChange(selected.includes(it.key) ? selected.filter((k) => k !== it.key) : [...selected, it.key]);
        }} style={{ accentColor: "#0f766e" }} />
        <span style={{ lineHeight: 1.3 }}>{it.label}</span>
      </label>
    ))}
  </div>
);

const MultiSelect = ({ label, map, selected, onChange }) => {
  const [open, setOpen] = useState(false);
  const ref = useRef(null);
  useEffect(() => {
    const onClick = (e) => { if (ref.current && !ref.current.contains(e.target)) setOpen(false); };
    document.addEventListener("mousedown", onClick);
    return () => document.removeEventListener("mousedown", onClick);
  }, []);
  const entries = Object.entries(map);
  const summary = selected.length === 0 || selected.length === entries.length ? "All" : selected.length === 1 ? (map[selected[0]] || selected[0]) : `${selected.length} selected`;
  return (
    <div ref={ref} style={{ position: "relative" }}>
      <button type="button" onClick={() => setOpen((o) => !o)}
        style={{ background: "#0f172a", color: "#e2e8f0", border: "1px solid #334155", borderRadius: 6, padding: "6px 10px", fontSize: 12, cursor: "pointer", minWidth: 150, textAlign: "left", display: "flex", alignItems: "center", justifyContent: "space-between", gap: 8 }}>
        <span>{label}: <span style={{ color: selected.length ? "#5eead4" : "#94a3b8", fontWeight: selected.length ? 600 : 400 }}>{summary}</span></span>
        <span style={{ color: "#64748b", fontSize: 10 }}>▼</span>
      </button>
      {open && (
        <div style={{ position: "absolute", top: "calc(100% + 4px)", left: 0, minWidth: 240, background: "#1e293b", border: "1px solid #334155", borderRadius: 8, padding: 8, zIndex: 100, boxShadow: "0 8px 24px rgba(0,0,0,.4)" }}>
          <div style={{ display: "flex", justifyContent: "space-between", marginBottom: 6, paddingBottom: 4, borderBottom: "1px solid #334155" }}>
            <button type="button" onClick={() => onChange([])} style={{ background: "none", border: "none", color: "#94a3b8", fontSize: 11, cursor: "pointer", padding: 0 }}>Clear</button>
            <button type="button" onClick={() => onChange(entries.map(([k]) => k))} style={{ background: "none", border: "none", color: "#5eead4", fontSize: 11, cursor: "pointer", padding: 0 }}>Select All</button>
          </div>
          {entries.map(([k, v]) => (
            <label key={k} style={{ display: "flex", alignItems: "center", gap: 8, cursor: "pointer", padding: "4px 2px", color: "#cbd5e1", fontSize: 12 }}>
              <input type="checkbox" checked={selected.includes(k)} onChange={() => {
                onChange(selected.includes(k) ? selected.filter((x) => x !== k) : [...selected, k]);
              }} style={{ accentColor: "#0f766e" }} />
              <span>{v}</span>
            </label>
          ))}
        </div>
      )}
    </div>
  );
};

const TT = ({ active, payload, label }) => {
  if (!active || !payload?.length) return null;
  return (
    <div style={{ background: "#0f172a", border: "1px solid #334155", borderRadius: 8, padding: "8px 12px", fontSize: 12, color: "#e2e8f0" }}>
      <div style={{ fontWeight: 600, marginBottom: 4 }}>{label}</div>
      {payload.map((p, i) => <div key={i} style={{ color: p.color }}>{p.name}: {typeof p.value === "number" ? p.value.toFixed(2) : p.value}</div>)}
    </div>
  );
};

// ─── MAIN APP ───
export default function Dashboard() {
  const [rawData, setRawData] = useState(null);
  const [columns, setColumns] = useState([]);
  const [tab, setTab] = useState("overview");
  const [filters, setFilters] = useState({ gender: "all", education: "all", location: "all", licence: [], ageMin: "", ageMax: "" });
  // Editable CMD cutoffs (defaults from CMD_SCALES)
  const [cmdCutoffs, setCmdCutoffs] = useState(() => {
    const obj = {};
    CMD_SCALES.forEach((s) => { obj[s.key] = s.cutoff; });
    return obj;
  });
  // Regression state
  const [lrOutcome, setLrOutcome] = useState("cesd_depressive_risk");
  const [lrPredictors, setLrPredictors] = useState(["sarrs_total_stress_score", "soa_negative_agency_total", "css_total", "passq_total", "sleep_hours"]);
  const [lrResult, setLrResult] = useState(null);
  const [lrUniResults, setLrUniResults] = useState(null);
  const [lrMode, setLrMode] = useState("both"); // "univariate" | "multivariable" | "both"
  const [linOutcome, setLinOutcome] = useState("cesd_total");
  const [linPredictors, setLinPredictors] = useState(["sarrs_total_stress_score", "soa_negative_agency_total", "css_total", "passq_total", "sleep_hours"]);
  const [linResult, setLinResult] = useState(null);
  const [linUniResults, setLinUniResults] = useState(null);
  const [linMode, setLinMode] = useState("both"); // "univariate" | "multivariable" | "both"
  const [copeMode, setCopeMode] = useState("14"); // "14" | "3" | "all"

  // Active COPE subscales based on toggle
  const activeCope = copeMode === "3" ? COPE_3CLASS : copeMode === "all" ? [...COPE_SUBSCALES, ...COPE_3CLASS] : COPE_SUBSCALES;

  // Dynamic RISK_FACTOR_GROUPS with active COPE subscales injected
  const ACTIVE_RISK_GROUPS = useMemo(() =>
    RISK_FACTOR_GROUPS.map((g) =>
      g.label === "Coping (Brief-COPE)" ? { ...g, vars: activeCope.map((s) => ({ ...s, type: "continuous" })) } : g
    ), [activeCope]);

  // File upload
  const handleFile = useCallback((e) => {
    const file = e.target.files[0];
    if (!file) return;
    const ext = file.name.split(".").pop().toLowerCase();
    if (ext === "csv" || ext === "tsv") {
      Papa.parse(file, {
        header: true, skipEmptyLines: true, dynamicTyping: true,
        complete: (r) => { setRawData(r.data); setColumns(r.meta.fields || []); },
      });
    } else {
      // Try XLSX via SheetJS
      const reader = new FileReader();
      reader.onload = (ev) => {
        try {
          const XLSX = window.XLSX || require("xlsx");
          // fallback: dynamically parse
        } catch { }
        // Papa fallback
        Papa.parse(file, {
          header: true, skipEmptyLines: true, dynamicTyping: true,
          complete: (r) => { setRawData(r.data); setColumns(r.meta.fields || []); },
        });
      };
      reader.readAsArrayBuffer(file);
    }
  }, []);

  // Set age filter defaults from actual data range
  useEffect(() => {
    if (!rawData || !rawData.length) return;
    const ages = rawData.map((r) => num(r.age)).filter((v) => v !== null);
    if (ages.length) {
      setFilters((f) => ({
        ...f,
        ageMin: f.ageMin || String(Math.floor(Math.min(...ages))),
        ageMax: f.ageMax || String(Math.ceil(Math.max(...ages))),
      }));
    }
  }, [rawData]);

  // Filtered data
  const data = useMemo(() => {
    if (!rawData) return [];
    return rawData.filter((r) => {
      if (filters.gender !== "all" && String(r.gender) !== filters.gender) return false;
      if (filters.education !== "all" && String(r.education_level) !== filters.education) return false;
      if (filters.location !== "all" && String(r.primary_race_location) !== filters.location) return false;
      if (filters.licence.length > 0 && !filters.licence.includes(String(r.licence_type))) return false;
      if (filters.ageMin && num(r.age) !== null && num(r.age) < Number(filters.ageMin)) return false;
      if (filters.ageMax && num(r.age) !== null && num(r.age) > Number(filters.ageMax)) return false;
      return true;
    });
  }, [rawData, filters]);

  // Missing data summary
  const missingInfo = useMemo(() => {
    if (!data.length || !columns.length) return [];
    return columns.map((c) => {
      const miss = data.filter((r) => r[c] === null || r[c] === undefined || r[c] === "" || (typeof r[c] === "number" && isNaN(r[c]))).length;
      return { col: c, missing: miss, pct: ((miss / data.length) * 100).toFixed(1) };
    }).filter((m) => m.missing > 0).sort((a, b) => b.missing - a.missing);
  }, [data, columns]);

  const col = (key) => data.map((r) => r[key]);

  // Severity distribution helper
  const severityDist = (key, bands) => {
    const vals = validNums(col(key));
    return bands.map((b) => ({ ...b, count: vals.filter((v) => v >= b.min && v <= b.max).length, pct: vals.length ? ((vals.filter((v) => v >= b.min && v <= b.max).length / vals.length) * 100).toFixed(1) : 0 }));
  };

  // CMD prevalence
  const cmdPrev = useMemo(() => {
    if (!data.length) return [];
    return CMD_SCALES.map((s) => {
      const co = cmdCutoffs[s.key] ?? s.cutoff;
      const vals = validNums(col(s.key));
      const n = vals.length;
      let nAbove;
      if (s.invert) {
        nAbove = vals.filter((v) => v < co).length;
      } else {
        nAbove = vals.filter((v) => v >= co).length;
      }
      return { ...s, cutoff: co, n, nAbove, prev: n ? ((nAbove / n) * 100).toFixed(1) : "—", mean: mean(col(s.key)), sd: sd(col(s.key)), med: median(col(s.key)) };
    });
  }, [data, cmdCutoffs]);

  // All predictors flat list
  const allPredictors = useMemo(() => ACTIVE_RISK_GROUPS.flatMap((g) => g.vars), [ACTIVE_RISK_GROUPS]);

  // Binary outcomes for logistic (use dynamic cutoffs)
  const binaryOutcomes = useMemo(() => {
    const cesdCo = cmdCutoffs["cesd_total"] ?? 16;
    const outcomes = [{ key: "cesd_depressive_risk", label: `CES-D Risk (≥${cesdCo})`, computeKey: "cesd_total", cutoff: cesdCo, invert: false }];
    CMD_SCALES.forEach((s) => {
      if (s.key !== "cesd_total") {
        const co = cmdCutoffs[s.key] ?? s.cutoff;
        outcomes.push({ key: `${s.key}_binary`, label: `${s.label} ${s.invert ? "<" : "≥"}${co}`, computeKey: s.key, cutoff: co, invert: s.invert });
      }
    });
    return outcomes;
  }, [cmdCutoffs]);

  // Helper: build binary Y array for logistic outcome
  const buildBinaryY = useCallback((outcomeKey) => {
    const outcomeInfo = binaryOutcomes.find((o) => o.key === outcomeKey);
    if (outcomeInfo?.computeKey) {
      return data.map((r) => {
        const v = num(r[outcomeInfo.computeKey]);
        if (v === null) return null;
        return outcomeInfo.invert ? (v < outcomeInfo.cutoff ? 1 : 0) : (v >= outcomeInfo.cutoff ? 1 : 0);
      });
    }
    return data.map((r) => num(r[outcomeKey]));
  }, [data, binaryOutcomes]);

  // Helper: expand a single predictor into X columns + varNames
  const expandPredictor = useCallback((predKey) => {
    const pInfo = allPredictors.find((p) => p.key === predKey);
    if (!pInfo) return null;
    const varNames = [];
    const XRows = data.map((r) => {
      const row = [];
      if (pInfo.type === "categorical" && pInfo.cats) {
        const catKeys = Object.keys(pInfo.cats).slice(1);
        catKeys.forEach((ck) => row.push(String(r[pInfo.key]) === ck ? 1 : 0));
      } else {
        row.push(num(r[pInfo.key]));
      }
      return row;
    });
    if (pInfo.type === "categorical" && pInfo.cats) {
      Object.keys(pInfo.cats).slice(1).forEach((ck) => varNames.push(`${pInfo.label}: ${pInfo.cats[ck]}`));
    } else {
      varNames.push(pInfo.label);
    }
    return { XRows, varNames, pInfo };
  }, [data, allPredictors]);

  // Helper: expand multiple predictors
  const expandPredictors = useCallback((predKeys) => {
    const predInfo = predKeys.map((k) => allPredictors.find((p) => p.key === k)).filter(Boolean);
    let varNames = [];
    const XRows = data.map((r) => {
      const row = [];
      predInfo.forEach((p) => {
        if (p.type === "categorical" && p.cats) {
          Object.keys(p.cats).slice(1).forEach((ck) => row.push(String(r[p.key]) === ck ? 1 : 0));
        } else {
          row.push(num(r[p.key]));
        }
      });
      return row;
    });
    predInfo.forEach((p) => {
      if (p.type === "categorical" && p.cats) {
        Object.keys(p.cats).slice(1).forEach((ck) => varNames.push(`${p.label}: ${p.cats[ck]}`));
      } else {
        varNames.push(p.label);
      }
    });
    return { XRows, varNames };
  }, [data, allPredictors]);

  // Helper: filter complete cases
  const completeCase = (yArr, XRows) => {
    const X = [], y = [];
    for (let i = 0; i < yArr.length; i++) {
      if (yArr[i] === null || yArr[i] === undefined) continue;
      if (XRows[i].some((v) => v === null || v === undefined || isNaN(v))) continue;
      X.push(XRows[i]);
      y.push(yArr[i]);
    }
    return { X, y };
  };

  // Run logistic regression (supports univariate / multivariable / both)
  const runLogistic = useCallback(() => {
    if (!data.length || !lrPredictors.length) return;
    const yArr = buildBinaryY(lrOutcome);

    // ── Multivariable ──
    if (lrMode === "multivariable" || lrMode === "both") {
      const { XRows, varNames } = expandPredictors(lrPredictors);
      const { X, y } = completeCase(yArr, XRows);
      if (y.length < varNames.length + 5) {
        setLrResult({ error: `Insufficient complete cases for multivariable model (n=${y.length})` });
      } else {
        try {
          const res = logisticRegression(y, X, varNames);
          setLrResult(res || { error: "Model did not converge" });
        } catch (e) { setLrResult({ error: e.message }); }
      }
    } else {
      setLrResult(null);
    }

    // ── Univariate (one model per predictor) ──
    if (lrMode === "univariate" || lrMode === "both") {
      const uniResults = [];
      lrPredictors.forEach((predKey) => {
        const exp = expandPredictor(predKey);
        if (!exp) return;
        const { XRows, varNames, pInfo } = exp;
        const { X, y } = completeCase(yArr, XRows);
        if (y.length < varNames.length + 5) {
          uniResults.push({ key: predKey, label: pInfo.label, error: `n=${y.length} insufficient` });
          return;
        }
        try {
          const res = logisticRegression(y, X, varNames);
          if (!res) { uniResults.push({ key: predKey, label: pInfo.label, error: "No convergence" }); return; }
          // Extract predictor rows (skip intercept at index 0)
          res.names.slice(1).forEach((name, i) => {
            uniResults.push({
              key: predKey, label: name, n: res.n,
              or: res.or[i + 1], ciLow: res.ciLow[i + 1], ciHigh: res.ciHigh[i + 1],
              beta: res.coefficients[i + 1], se: res.se[i + 1], z: res.z[i + 1], p: res.pvals[i + 1],
              pseudoR2: res.pseudoR2,
            });
          });
        } catch (e) { uniResults.push({ key: predKey, label: pInfo.label, error: e.message }); }
      });
      setLrUniResults(uniResults);
    } else {
      setLrUniResults(null);
    }
  }, [data, lrOutcome, lrPredictors, lrMode, buildBinaryY, expandPredictor, expandPredictors]);

  // Run linear regression (supports univariate / multivariable / both)
  const runLinear = useCallback(() => {
    if (!data.length || !linPredictors.length) return;
    const yArr = data.map((r) => num(r[linOutcome]));

    // ── Multivariable ──
    if (linMode === "multivariable" || linMode === "both") {
      const { XRows, varNames } = expandPredictors(linPredictors);
      const { X, y } = completeCase(yArr, XRows);
      if (y.length < varNames.length + 5) {
        setLinResult({ error: `Insufficient complete cases for multivariable model (n=${y.length})` });
      } else {
        try {
          const res = linearRegression(y, X, varNames);
          setLinResult(res || { error: "Matrix inversion failed" });
        } catch (e) { setLinResult({ error: e.message }); }
      }
    } else {
      setLinResult(null);
    }

    // ── Univariate (one model per predictor) ──
    if (linMode === "univariate" || linMode === "both") {
      const uniResults = [];
      linPredictors.forEach((predKey) => {
        const exp = expandPredictor(predKey);
        if (!exp) return;
        const { XRows, varNames, pInfo } = exp;
        const { X, y } = completeCase(yArr, XRows);
        if (y.length < varNames.length + 5) {
          uniResults.push({ key: predKey, label: pInfo.label, error: `n=${y.length} insufficient` });
          return;
        }
        try {
          const res = linearRegression(y, X, varNames);
          if (!res) { uniResults.push({ key: predKey, label: pInfo.label, error: "Failed" }); return; }
          res.names.slice(1).forEach((name, i) => {
            uniResults.push({
              key: predKey, label: name, n: res.n,
              beta: res.coefficients[i + 1], se: res.se[i + 1], t: res.tstat[i + 1], p: res.pvals[i + 1],
              r2: res.r2, adjR2: res.adjR2,
            });
          });
        } catch (e) { uniResults.push({ key: predKey, label: pInfo.label, error: e.message }); }
      });
      setLinUniResults(uniResults);
    } else {
      setLinUniResults(null);
    }
  }, [data, linOutcome, linPredictors, linMode, expandPredictor, expandPredictors]);

  // ─── WORD DOCUMENT EXPORT ───
  const exportWord = useCallback(async () => {
    if (!data || !data.length) return;
    const allData = data; // use current filtered dataset
    const irlData = allData.filter((r) => String(r.primary_race_location) === "1");
    const gbData = allData.filter((r) => String(r.primary_race_location) === "2");
    const groups = [
      { label: "Total Sample", data: allData },
      { label: "Ireland", data: irlData },
      { label: "Great Britain", data: gbData },
    ];
    const isFiltered = data.length !== rawData.length;

    const border = { style: BorderStyle.SINGLE, size: 1, color: "999999" };
    const borders = { top: border, bottom: border, left: border, right: border };
    const hdrShade = { fill: "D5E8F0", type: ShadingType.CLEAR };
    const cellMar = { top: 60, bottom: 60, left: 100, right: 100 };
    const tblW = 9360;
    // Helper: make a cell
    const mkCell = (text, opts = {}) => new TableCell({
      borders, margins: cellMar, width: opts.width ? { size: opts.width, type: WidthType.DXA } : undefined,
      shading: opts.header ? hdrShade : undefined, verticalAlign: "center",
      children: [new Paragraph({
        alignment: opts.align || AlignmentType.LEFT,
        children: [new TextRun({ text: String(text ?? "—"), bold: !!opts.bold, size: opts.size || 18, font: "Arial" })],
      })],
    });
    // Helper: stat row for continuous variable
    const contRow = (label, key, colWidths) => {
      const cells = [mkCell(label, { width: colWidths[0], bold: false })];
      groups.forEach((g, gi) => {
        const vals = g.data.map((r) => num(r[key])).filter((v) => v !== null);
        const n = vals.length;
        const m = n ? (vals.reduce((a, b) => a + b, 0) / n).toFixed(1) : "—";
        const s = n > 1 ? Math.sqrt(vals.reduce((a, b) => a + (b - (vals.reduce((x, y) => x + y, 0) / n)) ** 2, 0) / (n - 1)).toFixed(1) : "—";
        const sorted = [...vals].sort((a, b) => a - b);
        const mn = n ? sorted[0].toFixed(0) : "—";
        const mx = n ? sorted[n - 1].toFixed(0) : "—";
        cells.push(mkCell(String(n), { width: colWidths[1 + gi * 3], align: AlignmentType.CENTER }));
        cells.push(mkCell(m === "—" ? "—" : `${m} (${s})`, { width: colWidths[2 + gi * 3], align: AlignmentType.CENTER }));
        cells.push(mkCell(m === "—" ? "—" : `${mn}–${mx}`, { width: colWidths[3 + gi * 3], align: AlignmentType.CENTER }));
      });
      return new TableRow({ children: cells });
    };
    // Helper: categorical row
    const catRow = (label, key, catMap, colWidths) => {
      const rows = [];
      // header row for variable
      rows.push(new TableRow({ children: [
        mkCell(label, { width: colWidths[0], bold: true }),
        ...Array(9).fill(null).map((_, i) => mkCell("", { width: colWidths[i + 1] })),
      ]}));
      Object.entries(catMap).forEach(([k, v]) => {
        const cells = [mkCell(`  ${v}`, { width: colWidths[0] })];
        groups.forEach((g, gi) => {
          const total = g.data.filter((r) => r[key] !== null && r[key] !== undefined && r[key] !== "").length;
          const ct = g.data.filter((r) => String(r[key]) === k).length;
          const pct = total > 0 ? ((ct / total) * 100).toFixed(1) : "—";
          cells.push(mkCell(String(ct), { width: colWidths[1 + gi * 3], align: AlignmentType.CENTER }));
          cells.push(mkCell(`${pct}%`, { width: colWidths[2 + gi * 3], align: AlignmentType.CENTER }));
          cells.push(mkCell("", { width: colWidths[3 + gi * 3], align: AlignmentType.CENTER }));
        });
        rows.push(new TableRow({ children: cells }));
      });
      return rows;
    };

    // Column widths: Variable | (N, Mean(SD), Range) x3 groups
    const c0 = 2100; const cN = 560; const cMS = 1100; const cR = 780;
    const colWidths = [c0, cN, cMS, cR, cN, cMS, cR, cN, cMS, cR];

    // Header rows
    const mkHeaderRow1 = () => new TableRow({ children: [
      mkCell("", { width: c0, header: true }),
      mkCell("Total Sample", { width: cN + cMS + cR, header: true, bold: true, align: AlignmentType.CENTER }),
      mkCell("", { width: cMS, header: true }), mkCell("", { width: cR, header: true }),
      mkCell("Ireland", { width: cN + cMS + cR, header: true, bold: true, align: AlignmentType.CENTER }),
      mkCell("", { width: cMS, header: true }), mkCell("", { width: cR, header: true }),
      mkCell("Great Britain", { width: cN + cMS + cR, header: true, bold: true, align: AlignmentType.CENTER }),
      mkCell("", { width: cMS, header: true }), mkCell("", { width: cR, header: true }),
    ]});
    const mkHeaderRow2 = (c2label, c3label) => new TableRow({ children: [
      mkCell("Variable", { width: c0, header: true, bold: true }),
      mkCell("n", { width: cN, header: true, bold: true, align: AlignmentType.CENTER }),
      mkCell(c2label, { width: cMS, header: true, bold: true, align: AlignmentType.CENTER }),
      mkCell(c3label, { width: cR, header: true, bold: true, align: AlignmentType.CENTER }),
      mkCell("n", { width: cN, header: true, bold: true, align: AlignmentType.CENTER }),
      mkCell(c2label, { width: cMS, header: true, bold: true, align: AlignmentType.CENTER }),
      mkCell(c3label, { width: cR, header: true, bold: true, align: AlignmentType.CENTER }),
      mkCell("n", { width: cN, header: true, bold: true, align: AlignmentType.CENTER }),
      mkCell(c2label, { width: cMS, header: true, bold: true, align: AlignmentType.CENTER }),
      mkCell(c3label, { width: cR, header: true, bold: true, align: AlignmentType.CENTER }),
    ]});
    const sectionTitle = (text) => new Paragraph({
      spacing: { before: 300, after: 120 },
      children: [new TextRun({ text, bold: true, size: 20, font: "Arial", color: "333333" })],
    });
    const sectionNote = (text) => new Paragraph({
      spacing: { before: 40, after: 200 },
      children: [new TextRun({ text, italics: true, size: 16, font: "Arial", color: "666666" })],
    });

    // ── TABLE 1: DEMOGRAPHICS ──
    const demoRows = [
      mkHeaderRow1(), mkHeaderRow2("Mean (SD) / %", "Range"),
      contRow("Age (years)", "age", colWidths),
      ...catRow("Gender", "gender", GENDER_MAP, colWidths),
      ...catRow("Education Level", "education_level", EDUCATION_MAP, colWidths),
      ...catRow("Licence Type", "licence_type", LICENCE_MAP, colWidths),
      contRow("Years Licensed", "years_licensed", colWidths),
      ...catRow("Weekly Work Hours", "weekly_work_hours", WEEKLY_WORK_HRS, colWidths),
      ...catRow("Race Days/Week", "race_days_per_week", RACE_DAYS_WEEK, colWidths),
      ...catRow("Weekly Travel Hours", "weekly_travel_hours", WEEKLY_TRVL_HRS, colWidths),
      ...catRow("Annual Holiday Days", "annual_holiday_days", HOLIDAY_MAP, colWidths),
      contRow("Sleep Hours", "sleep_hours", colWidths),
      contRow("Falls (12m)", "num_falls_12m", colWidths),
      contRow("Injuries (12m)", "num_injuries_12m", colWidths),
      contRow("Days Missed", "days_missed_falls_injuries_bans", colWidths),
    ];
    const table1 = new Table({ width: { size: tblW, type: WidthType.DXA }, columnWidths: colWidths, rows: demoRows });

    // ── TABLE 2: CMD PREVALENCE ──
    const cmdColWidths = [c0, cN, cMS, cR, cN, cMS, cR, cN, cMS, cR];
    const cmdRows = [mkHeaderRow1(), mkHeaderRow2("Mean (SD)", "Prev. %")];
    CMD_SCALES.forEach((s) => {
      const co = cmdCutoffs[s.key] ?? s.cutoff;
      const cells = [mkCell(`${s.label} (${s.desc})`, { width: c0 })];
      groups.forEach((g) => {
        const vals = g.data.map((r) => num(r[s.key])).filter((v) => v !== null);
        const n = vals.length;
        const m = n ? (vals.reduce((a, b) => a + b, 0) / n).toFixed(1) : "—";
        const sd = n > 1 ? Math.sqrt(vals.reduce((a, b) => a + (b - (vals.reduce((x, y) => x + y, 0) / n)) ** 2, 0) / (n - 1)).toFixed(1) : "—";
        const nAbove = s.invert ? vals.filter((v) => v < co).length : vals.filter((v) => v >= co).length;
        const prev = n > 0 ? ((nAbove / n) * 100).toFixed(1) + "%" : "—";
        cells.push(mkCell(String(n), { width: cN, align: AlignmentType.CENTER }));
        cells.push(mkCell(m === "—" ? "—" : `${m} (${sd})`, { width: cMS, align: AlignmentType.CENTER }));
        cells.push(mkCell(prev, { width: cR, align: AlignmentType.CENTER }));
      });
      cmdRows.push(new TableRow({ children: cells }));
    });
    const table2 = new Table({ width: { size: tblW, type: WidthType.DXA }, columnWidths: cmdColWidths, rows: cmdRows });

    // ── TABLE 3: RISK FACTORS ──
    const rfRows = [mkHeaderRow1(), mkHeaderRow2("Mean (SD)", "Range")];
    [
      { label: "SARRS Total Stress", key: "sarrs_total_stress_score" },
      { label: "SARRS N Events", key: "sarrs_n_events" },
      { label: "Positive Agency (SoA)", key: "soa_positive_agency_total" },
      { label: "Negative Agency (SoA)", key: "soa_negative_agency_total" },
      { label: "Career Satisfaction (CSS)", key: "css_total" },
      { label: "PASS-Q Total", key: "passq_total" },
      { label: "Emotional Support", key: "passq_emotional_support" },
      { label: "Esteem Support", key: "passq_esteem_support" },
      { label: "Informational Support", key: "passq_informational_support" },
      { label: "Tangible Support", key: "passq_tangible_support" },
    ].forEach((v) => rfRows.push(contRow(v.label, v.key, colWidths)));
    const table3 = new Table({ width: { size: tblW, type: WidthType.DXA }, columnWidths: colWidths, rows: rfRows });

    // ── TABLE 4: COPING STRATEGIES ──
    const copeRows = [mkHeaderRow1(), mkHeaderRow2("Mean (SD)", "Range")];
    activeCope.forEach((s) => copeRows.push(contRow(s.label, s.key, colWidths)));
    const table4 = new Table({ width: { size: tblW, type: WidthType.DXA }, columnWidths: colWidths, rows: copeRows });

    // ── TABLE 5: WEIGHT MAKING ──
    const wmColWidths = [c0, cN, cMS, cR, cN, cMS, cR, cN, cMS, cR];
    const wmRows = [mkHeaderRow1(), mkHeaderRow2("n / Mean (SD)", "% / Range")];
    // Ever cut weight
    const wmCatCells = [mkCell("Ever Cut Weight", { width: c0, bold: true })];
    groups.forEach((g) => {
      const yes = g.data.filter((r) => num(r.ever_cut_weight) === 2).length;
      const tot = g.data.filter((r) => num(r.ever_cut_weight) === 1 || num(r.ever_cut_weight) === 2).length;
      wmCatCells.push(mkCell(String(yes), { width: cN, align: AlignmentType.CENTER }));
      wmCatCells.push(mkCell(`of ${tot}`, { width: cMS, align: AlignmentType.CENTER }));
      wmCatCells.push(mkCell(tot ? `${((yes / tot) * 100).toFixed(1)}%` : "—", { width: cR, align: AlignmentType.CENTER }));
    });
    wmRows.push(new TableRow({ children: wmCatCells }));
    [
      { label: "Non-Race Weight (lbs)", key: "non_race_weight_lbs" },
      { label: "Race Morning Weight (lbs)", key: "race_morning_weight_lbs" },
      { label: "Max Weight Lost (lbs)", key: "max_weight_lost_prep_lbs" },
    ].forEach((v) => wmRows.push(contRow(v.label, v.key, colWidths)));
    // Method prevalence
    WM_METHODS.forEach((m) => {
      const cells = [mkCell(m.label, { width: c0 })];
      groups.forEach((g) => {
        const n = g.data.filter((r) => num(r[m.key]) === 1 || String(r[m.key]) === "1").length;
        const tot = g.data.length;
        cells.push(mkCell(String(n), { width: cN, align: AlignmentType.CENTER }));
        cells.push(mkCell("", { width: cMS, align: AlignmentType.CENTER }));
        cells.push(mkCell(tot ? `${((n / tot) * 100).toFixed(1)}%` : "—", { width: cR, align: AlignmentType.CENTER }));
      });
      wmRows.push(new TableRow({ children: cells }));
    });
    const table5 = new Table({ width: { size: tblW, type: WidthType.DXA }, columnWidths: wmColWidths, rows: wmRows });

    // ── Assemble Document ──
    const doc = new Document({
      styles: {
        default: { document: { run: { font: "Arial", size: 20 } } },
        paragraphStyles: [
          { id: "Heading1", name: "Heading 1", basedOn: "Normal", next: "Normal", quickFormat: true,
            run: { size: 28, bold: true, font: "Arial" },
            paragraph: { spacing: { before: 240, after: 120 } } },
        ],
      },
      sections: [{
        properties: {
          page: {
            size: { width: 15840, height: 12240, orientation: "landscape" },
            margin: { top: 720, right: 720, bottom: 720, left: 720 },
          },
        },
        children: [
          new Paragraph({ heading: HeadingLevel.HEADING_1, children: [new TextRun({ text: "Jockey Mental Wellbeing Survey — Summary Tables", bold: true })] }),
          new Paragraph({ children: [new TextRun({ text: `Generated: ${new Date().toLocaleDateString()} | N = ${allData.length}${isFiltered ? ` of ${rawData.length} (filtered)` : ""} (Ireland: ${irlData.length}, Great Britain: ${gbData.length})`, size: 18, color: "666666", italics: true })] }),
          new Paragraph({ children: [new TextRun({ text: `Filters applied: ${(() => {
            const parts = [];
            if (filters.gender !== "all") parts.push(`Gender: ${GENDER_MAP[filters.gender] || filters.gender}`);
            if (filters.education !== "all") parts.push(`Education: ${EDUCATION_MAP[filters.education] || filters.education}`);
            if (filters.location !== "all") parts.push(`Location: ${LOCATION_MAP[filters.location] || filters.location}`);
            if (filters.licence.length > 0) parts.push(`Licence: ${filters.licence.map((v) => LICENCE_MAP[v] || v).join(", ")}`);
            if (filters.ageMin) parts.push(`Age ≥ ${filters.ageMin}`);
            if (filters.ageMax) parts.push(`Age ≤ ${filters.ageMax}`);
            return parts.length ? parts.join("; ") : "None (full sample)";
          })()}`, size: 18, color: "666666", italics: true })] }),

          sectionTitle("Table 1. Demographic Characteristics"),
          sectionNote("Values are mean (SD) and range for continuous variables; n (%) for categorical variables."),
          table1,

          new Paragraph({ children: [new PageBreak()] }),
          sectionTitle("Table 2. Common Mental Disorder Prevalence"),
          sectionNote(`Cut-offs: CES-D ≥${cmdCutoffs.cesd_total}, PHQ-9 ≥${cmdCutoffs.phq9_total}, GAD-7 ≥${cmdCutoffs.gad7_total}, K10 ≥${cmdCutoffs.k10_total}, WEMWBS <${cmdCutoffs.wemwbs_total}. Prevalence = % meeting/exceeding clinical threshold.`),
          table2,

          new Paragraph({ children: [new PageBreak()] }),
          sectionTitle("Table 3. Risk Factors — Life Stress, Agency, Career Satisfaction, and Social Support"),
          sectionNote("Values are mean (SD) and range. PASS-Q subscales represent mean of constituent items."),
          table3,

          new Paragraph({ children: [new PageBreak()] }),
          sectionTitle(copeMode === "3" ? "Table 4. Coping Strategies (Brief-COPE 3-Class Composite)" : "Table 4. Coping Strategies (Brief-COPE 14 Subscales)"),
          sectionNote(copeMode === "3"
            ? "Values are mean (SD) and range. 3-class composites: Problem-Focused, Emotion-Focused, and Avoidant coping."
            : "Values are mean (SD) and range. Each subscale ranges from 0 to 6 (sum of two items, each scored 0–3)."),
          table4,

          new Paragraph({ children: [new PageBreak()] }),
          sectionTitle("Table 5. Weight-Making Behaviours"),
          sectionNote("Ever Cut Weight shows n (%) of those who responded. Method prevalence shows n and % of total sample reporting each method."),
          table5,
        ],
      }],
    });
    const blob = await Packer.toBlob(doc);
    saveAs(blob, "Jockey_Wellbeing_Tables.docx");
  }, [data, rawData, cmdCutoffs, activeCope, copeMode, filters]);

  // ─── Shared docx helpers for regression exports ───
  const docxBorder = { style: BorderStyle.SINGLE, size: 1, color: "999999" };
  const docxBorders = { top: docxBorder, bottom: docxBorder, left: docxBorder, right: docxBorder };
  const docxHdr = { fill: "D5E8F0", type: ShadingType.CLEAR };
  const docxHdrDark = { fill: "B8D4E3", type: ShadingType.CLEAR };
  const docxSigBg = { fill: "E8F5E9", type: ShadingType.CLEAR };
  const docxCellMar = { top: 50, bottom: 50, left: 80, right: 80 };
  const dxCell = (text, opts = {}) => new TableCell({
    borders: docxBorders, margins: docxCellMar,
    width: opts.width ? { size: opts.width, type: WidthType.DXA } : undefined,
    shading: opts.shading || undefined,
    children: [new Paragraph({
      alignment: opts.align || AlignmentType.LEFT,
      children: [new TextRun({ text: String(text ?? "—"), bold: !!opts.bold, size: opts.size || 18, font: "Arial", italics: !!opts.italics, color: opts.color })],
    })],
  });
  const dxTitle = (text) => new Paragraph({ spacing: { before: 200, after: 80 }, children: [new TextRun({ text, bold: true, size: 22, font: "Arial" })] });
  const dxNote = (text) => new Paragraph({ spacing: { before: 20, after: 120 }, children: [new TextRun({ text, italics: true, size: 16, font: "Arial", color: "555555" })] });
  const dxSpacer = () => new Paragraph({ spacing: { before: 60, after: 60 }, children: [] });

  // Build filter description string
  const describeFilters = useCallback(() => {
    const parts = [];
    if (filters.gender !== "all") parts.push(`Gender: ${GENDER_MAP[filters.gender] || filters.gender}`);
    if (filters.education !== "all") parts.push(`Education: ${EDUCATION_MAP[filters.education] || filters.education}`);
    if (filters.location !== "all") parts.push(`Location: ${LOCATION_MAP[filters.location] || filters.location}`);
    if (filters.licence.length > 0) parts.push(`Licence: ${filters.licence.map((v) => LICENCE_MAP[v] || v).join(", ")}`);
    if (filters.ageMin) parts.push(`Age ≥ ${filters.ageMin}`);
    if (filters.ageMax) parts.push(`Age ≤ ${filters.ageMax}`);
    return parts.length ? parts.join("; ") : "None (full sample)";
  }, [filters]);

  // ─── LOGISTIC REGRESSION WORD EXPORT ───
  const exportLogisticWord = useCallback(async () => {
    if (!data.length) return;
    const outLabel = binaryOutcomes.find((o) => o.key === lrOutcome)?.label || lrOutcome;
    const tblW = 9360;
    const sections = [];

    // Cover info
    sections.push(dxTitle(`Logistic Regression Results — ${outLabel}`));
    sections.push(dxNote(`Model type: ${lrMode === "both" ? "Univariate + Multivariable" : lrMode === "univariate" ? "Univariate (Unadjusted)" : "Multivariable (Adjusted)"} | Filtered N = ${data.length} | Filters: ${describeFilters()} | Generated: ${new Date().toLocaleDateString()}`));
    sections.push(dxNote(`Clinical cut-offs: CES-D ≥${cmdCutoffs.cesd_total}, PHQ-9 ≥${cmdCutoffs.phq9_total}, GAD-7 ≥${cmdCutoffs.gad7_total}, K10 ≥${cmdCutoffs.k10_total}, WEMWBS <${cmdCutoffs.wemwbs_total}`));

    // ── Univariate Table ──
    if ((lrMode === "univariate" || lrMode === "both") && lrUniResults && lrUniResults.length) {
      sections.push(dxSpacer());
      sections.push(dxTitle("Table A. Univariate (Unadjusted) Logistic Regression"));
      sections.push(dxNote("Each predictor entered individually. OR = Odds Ratio. *p<.05, **p<.01, ***p<.001."));
      const cw = [2200, 600, 900, 900, 1000, 900, 1100, 900, 860];
      const hdrRow = new TableRow({ children: [
        dxCell("Variable", { width: cw[0], bold: true, shading: docxHdr }),
        dxCell("N", { width: cw[1], bold: true, shading: docxHdr, align: AlignmentType.CENTER }),
        dxCell("β", { width: cw[2], bold: true, shading: docxHdr, align: AlignmentType.CENTER }),
        dxCell("SE", { width: cw[3], bold: true, shading: docxHdr, align: AlignmentType.CENTER }),
        dxCell("z", { width: cw[4], bold: true, shading: docxHdr, align: AlignmentType.CENTER }),
        dxCell("OR", { width: cw[5], bold: true, shading: docxHdr, align: AlignmentType.CENTER }),
        dxCell("95% CI", { width: cw[6], bold: true, shading: docxHdr, align: AlignmentType.CENTER }),
        dxCell("p", { width: cw[7], bold: true, shading: docxHdr, align: AlignmentType.CENTER }),
        dxCell("Pseudo R²", { width: cw[8], bold: true, shading: docxHdr, align: AlignmentType.CENTER }),
      ]});
      const dataRows = lrUniResults.map((r) => {
        if (r.error) return new TableRow({ children: [
          dxCell(r.label, { width: cw[0] }),
          ...Array(8).fill(null).map((_, i) => dxCell(i === 0 ? r.error : "", { width: cw[i + 1], align: AlignmentType.CENTER, italics: i === 0, color: "CC0000" })),
        ]});
        const sig = r.p < 0.05;
        const shade = sig ? docxSigBg : undefined;
        return new TableRow({ children: [
          dxCell(r.label, { width: cw[0], shading: shade }),
          dxCell(String(r.n), { width: cw[1], align: AlignmentType.CENTER, shading: shade }),
          dxCell(fmt(r.beta, 3), { width: cw[2], align: AlignmentType.CENTER, shading: shade }),
          dxCell(fmt(r.se, 3), { width: cw[3], align: AlignmentType.CENTER, shading: shade }),
          dxCell(fmt(r.z, 2), { width: cw[4], align: AlignmentType.CENTER, shading: shade }),
          dxCell(fmt(r.or, 3), { width: cw[5], align: AlignmentType.CENTER, bold: sig, shading: shade }),
          dxCell(`[${fmt(r.ciLow, 2)}, ${fmt(r.ciHigh, 2)}]`, { width: cw[6], align: AlignmentType.CENTER, shading: shade }),
          dxCell(`${fmtP(r.p)}${pStar(r.p)}`, { width: cw[7], align: AlignmentType.CENTER, bold: sig, shading: shade }),
          dxCell(fmt(r.pseudoR2, 3), { width: cw[8], align: AlignmentType.CENTER, shading: shade }),
        ]});
      });
      sections.push(new Table({ width: { size: tblW, type: WidthType.DXA }, columnWidths: cw, rows: [hdrRow, ...dataRows] }));
    }

    // ── Multivariable Table ──
    if ((lrMode === "multivariable" || lrMode === "both") && lrResult && !lrResult.error) {
      sections.push(dxSpacer());
      sections.push(dxTitle("Table B. Multivariable (Adjusted) Logistic Regression"));
      sections.push(dxNote(`All predictors entered simultaneously. N = ${lrResult.n}, Pseudo R² = ${fmt(lrResult.pseudoR2, 3)}, Log-Likelihood = ${fmt(lrResult.ll, 1)}. *p<.05, **p<.01, ***p<.001.`));
      const cw = [2400, 900, 900, 900, 1000, 1100, 1160];
      const hdrRow = new TableRow({ children: [
        dxCell("Variable", { width: cw[0], bold: true, shading: docxHdr }),
        dxCell("β", { width: cw[1], bold: true, shading: docxHdr, align: AlignmentType.CENTER }),
        dxCell("SE", { width: cw[2], bold: true, shading: docxHdr, align: AlignmentType.CENTER }),
        dxCell("z", { width: cw[3], bold: true, shading: docxHdr, align: AlignmentType.CENTER }),
        dxCell("OR", { width: cw[4], bold: true, shading: docxHdr, align: AlignmentType.CENTER }),
        dxCell("95% CI", { width: cw[5], bold: true, shading: docxHdr, align: AlignmentType.CENTER }),
        dxCell("p", { width: cw[6], bold: true, shading: docxHdr, align: AlignmentType.CENTER }),
      ]});
      const dataRows = lrResult.names.map((n, i) => {
        const isInt = i === 0;
        const sig = !isInt && lrResult.pvals[i] < 0.05;
        const shade = sig ? docxSigBg : undefined;
        return new TableRow({ children: [
          dxCell(n, { width: cw[0], italics: isInt, shading: shade }),
          dxCell(fmt(lrResult.coefficients[i], 3), { width: cw[1], align: AlignmentType.CENTER, shading: shade }),
          dxCell(fmt(lrResult.se[i], 3), { width: cw[2], align: AlignmentType.CENTER, shading: shade }),
          dxCell(fmt(lrResult.z[i], 2), { width: cw[3], align: AlignmentType.CENTER, shading: shade }),
          dxCell(isInt ? "—" : fmt(lrResult.or[i], 3), { width: cw[4], align: AlignmentType.CENTER, bold: sig, shading: shade }),
          dxCell(isInt ? "—" : `[${fmt(lrResult.ciLow[i], 2)}, ${fmt(lrResult.ciHigh[i], 2)}]`, { width: cw[5], align: AlignmentType.CENTER, shading: shade }),
          dxCell(isInt ? fmtP(lrResult.pvals[i]) : `${fmtP(lrResult.pvals[i])}${pStar(lrResult.pvals[i])}`, { width: cw[6], align: AlignmentType.CENTER, bold: sig, shading: shade }),
        ]});
      });
      sections.push(new Table({ width: { size: tblW, type: WidthType.DXA }, columnWidths: cw, rows: [hdrRow, ...dataRows] }));
    }

    // ── Combined side-by-side table ──
    if (lrMode === "both" && lrUniResults && lrResult && !lrResult.error) {
      sections.push(new Paragraph({ children: [new PageBreak()] }));
      sections.push(dxTitle("Table C. Unadjusted vs Adjusted Odds Ratios — Side by Side"));
      sections.push(dxNote(`Outcome: ${outLabel}. Adjusted model N = ${lrResult.n}, Pseudo R² = ${fmt(lrResult.pseudoR2, 3)}. Significant (p<.05) cells highlighted. *p<.05, **p<.01, ***p<.001.`));
      const cw = [1800, 800, 1000, 800, 800, 1000, 800, 800, 1560];
      const hdr1 = new TableRow({ children: [
        dxCell("", { width: cw[0], shading: docxHdr }),
        dxCell("Unadjusted", { width: cw[1] + cw[2] + cw[3], bold: true, shading: docxHdrDark, align: AlignmentType.CENTER }),
        dxCell("", { width: cw[2], shading: docxHdrDark }), dxCell("", { width: cw[3], shading: docxHdrDark }),
        dxCell("Adjusted", { width: cw[4] + cw[5] + cw[6], bold: true, shading: docxHdrDark, align: AlignmentType.CENTER }),
        dxCell("", { width: cw[5], shading: docxHdrDark }), dxCell("", { width: cw[6], shading: docxHdrDark }),
        dxCell("", { width: cw[7], shading: docxHdr }), dxCell("", { width: cw[8], shading: docxHdr }),
      ]});
      const hdr2 = new TableRow({ children: [
        dxCell("Variable", { width: cw[0], bold: true, shading: docxHdr }),
        dxCell("OR", { width: cw[1], bold: true, shading: docxHdr, align: AlignmentType.CENTER }),
        dxCell("95% CI", { width: cw[2], bold: true, shading: docxHdr, align: AlignmentType.CENTER }),
        dxCell("p", { width: cw[3], bold: true, shading: docxHdr, align: AlignmentType.CENTER }),
        dxCell("OR", { width: cw[4], bold: true, shading: docxHdr, align: AlignmentType.CENTER }),
        dxCell("95% CI", { width: cw[5], bold: true, shading: docxHdr, align: AlignmentType.CENTER }),
        dxCell("p", { width: cw[6], bold: true, shading: docxHdr, align: AlignmentType.CENTER }),
        dxCell("N (unadj)", { width: cw[7], bold: true, shading: docxHdr, align: AlignmentType.CENTER }),
        dxCell("N (adj)", { width: cw[8], bold: true, shading: docxHdr, align: AlignmentType.CENTER }),
      ]});
      const dataRows = lrResult.names.slice(1).map((name, i) => {
        const uni = lrUniResults.find((u) => u.label === name);
        const uSig = uni && !uni.error && uni.p < 0.05;
        const aSig = lrResult.pvals[i + 1] < 0.05;
        return new TableRow({ children: [
          dxCell(name, { width: cw[0] }),
          dxCell(uni && !uni.error ? fmt(uni.or, 3) : "—", { width: cw[1], align: AlignmentType.CENTER, bold: uSig, shading: uSig ? docxSigBg : undefined }),
          dxCell(uni && !uni.error ? `[${fmt(uni.ciLow, 2)}, ${fmt(uni.ciHigh, 2)}]` : "—", { width: cw[2], align: AlignmentType.CENTER, shading: uSig ? docxSigBg : undefined }),
          dxCell(uni && !uni.error ? `${fmtP(uni.p)}${pStar(uni.p)}` : "—", { width: cw[3], align: AlignmentType.CENTER, bold: uSig, shading: uSig ? docxSigBg : undefined }),
          dxCell(fmt(lrResult.or[i + 1], 3), { width: cw[4], align: AlignmentType.CENTER, bold: aSig, shading: aSig ? docxSigBg : undefined }),
          dxCell(`[${fmt(lrResult.ciLow[i + 1], 2)}, ${fmt(lrResult.ciHigh[i + 1], 2)}]`, { width: cw[5], align: AlignmentType.CENTER, shading: aSig ? docxSigBg : undefined }),
          dxCell(`${fmtP(lrResult.pvals[i + 1])}${pStar(lrResult.pvals[i + 1])}`, { width: cw[6], align: AlignmentType.CENTER, bold: aSig, shading: aSig ? docxSigBg : undefined }),
          dxCell(uni && !uni.error ? String(uni.n) : "—", { width: cw[7], align: AlignmentType.CENTER }),
          dxCell(String(lrResult.n), { width: cw[8], align: AlignmentType.CENTER }),
        ]});
      });
      sections.push(new Table({ width: { size: tblW, type: WidthType.DXA }, columnWidths: cw, rows: [hdr1, hdr2, ...dataRows] }));
    }

    const doc = new Document({
      styles: { default: { document: { run: { font: "Arial", size: 20 } } } },
      sections: [{
        properties: { page: { size: { width: 15840, height: 12240, orientation: "landscape" }, margin: { top: 720, right: 720, bottom: 720, left: 720 } } },
        children: sections,
      }],
    });
    const blob = await Packer.toBlob(doc);
    saveAs(blob, `Logistic_Regression_${outLabel.replace(/[^a-zA-Z0-9]/g, "_")}.docx`);
  }, [data, lrOutcome, lrMode, lrResult, lrUniResults, binaryOutcomes, filters, cmdCutoffs, describeFilters]);

  // ─── LINEAR REGRESSION WORD EXPORT ───
  const exportLinearWord = useCallback(async () => {
    if (!data.length) return;
    const outLabel = CMD_SCALES.find((s) => s.key === linOutcome)?.label || linOutcome;
    const tblW = 9360;
    const sections = [];

    sections.push(dxTitle(`Linear Regression Results — ${outLabel} Total`));
    sections.push(dxNote(`Model type: ${linMode === "both" ? "Univariate + Multivariable" : linMode === "univariate" ? "Univariate (Unadjusted)" : "Multivariable (Adjusted)"} | Filtered N = ${data.length} | Filters: ${describeFilters()} | Generated: ${new Date().toLocaleDateString()}`));

    // ── Univariate Table ──
    if ((linMode === "univariate" || linMode === "both") && linUniResults && linUniResults.length) {
      sections.push(dxSpacer());
      sections.push(dxTitle("Table A. Univariate (Unadjusted) Linear Regression"));
      sections.push(dxNote("Each predictor entered individually. *p<.05, **p<.01, ***p<.001."));
      const cw = [2400, 700, 1100, 1100, 1000, 1000, 1000, 1060];
      const hdrRow = new TableRow({ children: [
        dxCell("Variable", { width: cw[0], bold: true, shading: docxHdr }),
        dxCell("N", { width: cw[1], bold: true, shading: docxHdr, align: AlignmentType.CENTER }),
        dxCell("β", { width: cw[2], bold: true, shading: docxHdr, align: AlignmentType.CENTER }),
        dxCell("SE", { width: cw[3], bold: true, shading: docxHdr, align: AlignmentType.CENTER }),
        dxCell("t", { width: cw[4], bold: true, shading: docxHdr, align: AlignmentType.CENTER }),
        dxCell("p", { width: cw[5], bold: true, shading: docxHdr, align: AlignmentType.CENTER }),
        dxCell("R²", { width: cw[6], bold: true, shading: docxHdr, align: AlignmentType.CENTER }),
        dxCell("Adj R²", { width: cw[7], bold: true, shading: docxHdr, align: AlignmentType.CENTER }),
      ]});
      const dataRows = linUniResults.map((r) => {
        if (r.error) return new TableRow({ children: [
          dxCell(r.label, { width: cw[0] }),
          ...Array(7).fill(null).map((_, i) => dxCell(i === 0 ? r.error : "", { width: cw[i + 1], align: AlignmentType.CENTER, italics: i === 0, color: "CC0000" })),
        ]});
        const sig = r.p < 0.05;
        const shade = sig ? docxSigBg : undefined;
        return new TableRow({ children: [
          dxCell(r.label, { width: cw[0], shading: shade }),
          dxCell(String(r.n), { width: cw[1], align: AlignmentType.CENTER, shading: shade }),
          dxCell(fmt(r.beta, 4), { width: cw[2], align: AlignmentType.CENTER, bold: sig, shading: shade }),
          dxCell(fmt(r.se, 4), { width: cw[3], align: AlignmentType.CENTER, shading: shade }),
          dxCell(fmt(r.t, 2), { width: cw[4], align: AlignmentType.CENTER, shading: shade }),
          dxCell(`${fmtP(r.p)}${pStar(r.p)}`, { width: cw[5], align: AlignmentType.CENTER, bold: sig, shading: shade }),
          dxCell(fmt(r.r2, 3), { width: cw[6], align: AlignmentType.CENTER, shading: shade }),
          dxCell(fmt(r.adjR2, 3), { width: cw[7], align: AlignmentType.CENTER, shading: shade }),
        ]});
      });
      sections.push(new Table({ width: { size: tblW, type: WidthType.DXA }, columnWidths: cw, rows: [hdrRow, ...dataRows] }));
    }

    // ── Multivariable Table ──
    if ((linMode === "multivariable" || linMode === "both") && linResult && !linResult.error) {
      sections.push(dxSpacer());
      sections.push(dxTitle("Table B. Multivariable (Adjusted) Linear Regression"));
      sections.push(dxNote(`All predictors entered simultaneously. N = ${linResult.n}, R² = ${fmt(linResult.r2, 3)}, Adj R² = ${fmt(linResult.adjR2, 3)}, RMSE = ${fmt(Math.sqrt(linResult.mse), 2)}. *p<.05, **p<.01, ***p<.001.`));
      const cw = [2600, 1200, 1200, 1200, 1200, 1960];
      const hdrRow = new TableRow({ children: [
        dxCell("Variable", { width: cw[0], bold: true, shading: docxHdr }),
        dxCell("β", { width: cw[1], bold: true, shading: docxHdr, align: AlignmentType.CENTER }),
        dxCell("SE", { width: cw[2], bold: true, shading: docxHdr, align: AlignmentType.CENTER }),
        dxCell("t", { width: cw[3], bold: true, shading: docxHdr, align: AlignmentType.CENTER }),
        dxCell("p", { width: cw[4], bold: true, shading: docxHdr, align: AlignmentType.CENTER }),
        dxCell("Sig", { width: cw[5], bold: true, shading: docxHdr, align: AlignmentType.CENTER }),
      ]});
      const dataRows = linResult.names.map((n, i) => {
        const isInt = i === 0;
        const sig = !isInt && linResult.pvals[i] < 0.05;
        const shade = sig ? docxSigBg : undefined;
        return new TableRow({ children: [
          dxCell(n, { width: cw[0], italics: isInt, shading: shade }),
          dxCell(fmt(linResult.coefficients[i], 4), { width: cw[1], align: AlignmentType.CENTER, bold: sig, shading: shade }),
          dxCell(fmt(linResult.se[i], 4), { width: cw[2], align: AlignmentType.CENTER, shading: shade }),
          dxCell(fmt(linResult.tstat[i], 2), { width: cw[3], align: AlignmentType.CENTER, shading: shade }),
          dxCell(`${fmtP(linResult.pvals[i])}${pStar(linResult.pvals[i])}`, { width: cw[4], align: AlignmentType.CENTER, bold: sig, shading: shade }),
          dxCell(isInt ? "" : pStar(linResult.pvals[i]), { width: cw[5], align: AlignmentType.CENTER, shading: shade }),
        ]});
      });
      sections.push(new Table({ width: { size: tblW, type: WidthType.DXA }, columnWidths: cw, rows: [hdrRow, ...dataRows] }));
    }

    // ── Combined side-by-side ──
    if (linMode === "both" && linUniResults && linResult && !linResult.error) {
      sections.push(new Paragraph({ children: [new PageBreak()] }));
      sections.push(dxTitle("Table C. Unadjusted vs Adjusted Coefficients — Side by Side"));
      sections.push(dxNote(`Outcome: ${outLabel}. Adjusted model N = ${linResult.n}, R² = ${fmt(linResult.r2, 3)}, Adj R² = ${fmt(linResult.adjR2, 3)}. *p<.05, **p<.01, ***p<.001.`));
      const cw = [1900, 900, 900, 800, 900, 900, 800, 800, 1460];
      const hdr1 = new TableRow({ children: [
        dxCell("", { width: cw[0], shading: docxHdr }),
        dxCell("Unadjusted", { width: cw[1] + cw[2] + cw[3], bold: true, shading: docxHdrDark, align: AlignmentType.CENTER }),
        dxCell("", { width: cw[2], shading: docxHdrDark }), dxCell("", { width: cw[3], shading: docxHdrDark }),
        dxCell("Adjusted", { width: cw[4] + cw[5] + cw[6], bold: true, shading: docxHdrDark, align: AlignmentType.CENTER }),
        dxCell("", { width: cw[5], shading: docxHdrDark }), dxCell("", { width: cw[6], shading: docxHdrDark }),
        dxCell("", { width: cw[7], shading: docxHdr }), dxCell("", { width: cw[8], shading: docxHdr }),
      ]});
      const hdr2 = new TableRow({ children: [
        dxCell("Variable", { width: cw[0], bold: true, shading: docxHdr }),
        dxCell("β", { width: cw[1], bold: true, shading: docxHdr, align: AlignmentType.CENTER }),
        dxCell("SE", { width: cw[2], bold: true, shading: docxHdr, align: AlignmentType.CENTER }),
        dxCell("p", { width: cw[3], bold: true, shading: docxHdr, align: AlignmentType.CENTER }),
        dxCell("β", { width: cw[4], bold: true, shading: docxHdr, align: AlignmentType.CENTER }),
        dxCell("SE", { width: cw[5], bold: true, shading: docxHdr, align: AlignmentType.CENTER }),
        dxCell("p", { width: cw[6], bold: true, shading: docxHdr, align: AlignmentType.CENTER }),
        dxCell("R² (unadj)", { width: cw[7], bold: true, shading: docxHdr, align: AlignmentType.CENTER }),
        dxCell("R² (adj model)", { width: cw[8], bold: true, shading: docxHdr, align: AlignmentType.CENTER }),
      ]});
      const dataRows = linResult.names.slice(1).map((name, i) => {
        const uni = linUniResults.find((u) => u.label === name);
        const uSig = uni && !uni.error && uni.p < 0.05;
        const aSig = linResult.pvals[i + 1] < 0.05;
        return new TableRow({ children: [
          dxCell(name, { width: cw[0] }),
          dxCell(uni && !uni.error ? fmt(uni.beta, 4) : "—", { width: cw[1], align: AlignmentType.CENTER, bold: uSig, shading: uSig ? docxSigBg : undefined }),
          dxCell(uni && !uni.error ? fmt(uni.se, 4) : "—", { width: cw[2], align: AlignmentType.CENTER, shading: uSig ? docxSigBg : undefined }),
          dxCell(uni && !uni.error ? `${fmtP(uni.p)}${pStar(uni.p)}` : "—", { width: cw[3], align: AlignmentType.CENTER, bold: uSig, shading: uSig ? docxSigBg : undefined }),
          dxCell(fmt(linResult.coefficients[i + 1], 4), { width: cw[4], align: AlignmentType.CENTER, bold: aSig, shading: aSig ? docxSigBg : undefined }),
          dxCell(fmt(linResult.se[i + 1], 4), { width: cw[5], align: AlignmentType.CENTER, shading: aSig ? docxSigBg : undefined }),
          dxCell(`${fmtP(linResult.pvals[i + 1])}${pStar(linResult.pvals[i + 1])}`, { width: cw[6], align: AlignmentType.CENTER, bold: aSig, shading: aSig ? docxSigBg : undefined }),
          dxCell(uni && !uni.error ? fmt(uni.r2, 3) : "—", { width: cw[7], align: AlignmentType.CENTER }),
          dxCell(fmt(linResult.r2, 3), { width: cw[8], align: AlignmentType.CENTER }),
        ]});
      });
      sections.push(new Table({ width: { size: tblW, type: WidthType.DXA }, columnWidths: cw, rows: [hdr1, hdr2, ...dataRows] }));
    }

    const doc = new Document({
      styles: { default: { document: { run: { font: "Arial", size: 20 } } } },
      sections: [{
        properties: { page: { size: { width: 15840, height: 12240, orientation: "landscape" }, margin: { top: 720, right: 720, bottom: 720, left: 720 } } },
        children: sections,
      }],
    });
    const blob = await Packer.toBlob(doc);
    saveAs(blob, `Linear_Regression_${outLabel.replace(/[^a-zA-Z0-9]/g, "_")}.docx`);
  }, [data, linOutcome, linMode, linResult, linUniResults, filters, describeFilters]);

  const TABS = [
    { id: "overview", label: "Overview" },
    { id: "cmd", label: "CMD Prevalence" },
    { id: "risk", label: "Risk Factors" },
    { id: "coping", label: "Coping & Support" },
    { id: "weight", label: "Weight Making" },
    { id: "logistic", label: "Logistic Regression" },
    { id: "linear", label: "Linear Regression" },
    { id: "missing", label: "Missing Data" },
  ];

  // ─── UPLOAD SCREEN ───
  if (!rawData) {
    return (
      <div style={{ minHeight: "100vh", background: "#0f172a", display: "flex", alignItems: "center", justifyContent: "center", fontFamily: "'IBM Plex Sans', 'Segoe UI', sans-serif" }}>
        <div style={{ textAlign: "center", maxWidth: 520 }}>
          <div style={{ fontSize: 13, letterSpacing: 4, color: "#5eead4", fontWeight: 700, marginBottom: 8, textTransform: "uppercase" }}></div>
          <h1 style={{ fontSize: 32, color: "#f0fdfa", fontWeight: 700, margin: "0 0 8px", lineHeight: 1.2 }}>Jockey Mental Wellbeing</h1>
          <h2 style={{ fontSize: 18, color: "#94a3b8", fontWeight: 400, margin: "0 0 32px" }}>Survey Analytics Dashboard</h2>
          <label style={{
            display: "inline-flex", alignItems: "center", gap: 10, padding: "14px 32px",
            background: "linear-gradient(135deg, #0f766e, #0e7490)", borderRadius: 10,
            color: "#fff", fontSize: 15, fontWeight: 600, cursor: "pointer",
            boxShadow: "0 4px 24px rgba(15,118,110,.4)", transition: "transform .15s",
          }}
            onMouseEnter={(e) => e.currentTarget.style.transform = "translateY(-2px)"}
            onMouseLeave={(e) => e.currentTarget.style.transform = "translateY(0)"}
          >
            <svg width="20" height="20" fill="none" stroke="currentColor" strokeWidth="2" viewBox="0 0 24 24"><path d="M21 15v4a2 2 0 01-2 2H5a2 2 0 01-2-2v-4M17 8l-5-5-5 5M12 3v12" /></svg>
            Upload CSV Data
            <input type="file" accept=".csv,.tsv,.xlsx,.xls" onChange={handleFile} style={{ display: "none" }} />
          </label>
          <div style={{ color: "#64748b", fontSize: 12, marginTop: 16 }}>Supports CSV / TSV files</div>
        </div>
      </div>
    );
  }

  // ─── DASHBOARD ───
  return (
    <div style={{ minHeight: "100vh", background: "#0f172a", fontFamily: "'IBM Plex Sans', 'Segoe UI', sans-serif", color: "#e2e8f0", padding: "16px 20px" }}>
      <div style={{ maxWidth: 1280, margin: "0 auto" }}>
        {/* Header */}
        <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 16, flexWrap: "wrap", gap: 10 }}>
          <div>
            <div style={{ fontSize: 11, letterSpacing: 3, color: "#5eead4", fontWeight: 700, textTransform: "uppercase" }}></div>
            <h1 style={{ fontSize: 22, fontWeight: 700, color: "#f0fdfa", margin: "4px 0 0" }}>Jockey Mental Wellbeing Dashboard</h1>
          </div>
          <div style={{ display: "flex", gap: 8, alignItems: "center", fontSize: 12, color: "#94a3b8" }}>
            <span style={{ background: "#1e293b", padding: "4px 10px", borderRadius: 6, border: "1px solid #334155" }}>
              N = {data.length} / {rawData.length}
            </span>
            <label style={{ padding: "6px 12px", background: "#1e293b", borderRadius: 6, border: "1px solid #334155", cursor: "pointer", color: "#5eead4", fontWeight: 600 }}>
              ↻ New File
              <input type="file" accept=".csv,.tsv" onChange={handleFile} style={{ display: "none" }} />
            </label>
            <button onClick={exportWord}
              style={{ padding: "6px 12px", background: "linear-gradient(135deg, #0f766e, #0e7490)", borderRadius: 6, border: "none", cursor: "pointer", color: "#fff", fontWeight: 600, fontSize: 12 }}>
              ⬇ Export Word Tables
            </button>
          </div>
        </div>

        {/* Filters */}
        <Card style={{ padding: 14, overflow: "visible" }}>
          <div style={{ display: "flex", flexWrap: "wrap", gap: 12, alignItems: "center", fontSize: 12 }}>
            <span style={{ color: "#5eead4", fontWeight: 700, fontSize: 11, textTransform: "uppercase", letterSpacing: 1 }}>Filters</span>
            {[
              { key: "gender", label: "Gender", map: GENDER_MAP },
              { key: "education", label: "Education", map: EDUCATION_MAP },
              { key: "location", label: "Location", map: LOCATION_MAP },
            ].map(({ key, label, map }) => (
              <select key={key} value={filters[key]} onChange={(e) => setFilters((f) => ({ ...f, [key]: e.target.value }))}
                style={{ background: "#0f172a", color: "#e2e8f0", border: "1px solid #334155", borderRadius: 6, padding: "6px 10px", fontSize: 12 }}>
                <option value="all">{label}: All</option>
                {Object.entries(map).map(([k, v]) => <option key={k} value={k}>{v}</option>)}
              </select>
            ))}
            <MultiSelect label="Licence" map={LICENCE_MAP}
              selected={filters.licence}
              onChange={(vals) => setFilters((f) => ({ ...f, licence: vals }))} />
            <div style={{ display: "flex", alignItems: "center", gap: 4 }}>
              <span style={{ color: "#94a3b8" }}>Age:</span>
              <input type="number" placeholder="Min" value={filters.ageMin} onChange={(e) => setFilters((f) => ({ ...f, ageMin: e.target.value }))}
                style={{ width: 52, background: "#0f172a", color: "#e2e8f0", border: "1px solid #334155", borderRadius: 6, padding: "5px 6px", fontSize: 12 }} />
              <span style={{ color: "#475569" }}>–</span>
              <input type="number" placeholder="Max" value={filters.ageMax} onChange={(e) => setFilters((f) => ({ ...f, ageMax: e.target.value }))}
                style={{ width: 52, background: "#0f172a", color: "#e2e8f0", border: "1px solid #334155", borderRadius: 6, padding: "5px 6px", fontSize: 12 }} />
            </div>
            <button onClick={() => {
              const ages = rawData ? rawData.map((r) => num(r.age)).filter((v) => v !== null) : [];
              setFilters({
                gender: "all", education: "all", location: "all", licence: [],
                ageMin: ages.length ? String(Math.floor(Math.min(...ages))) : "",
                ageMax: ages.length ? String(Math.ceil(Math.max(...ages))) : "",
              });
            }}
              style={{ padding: "5px 10px", background: "#334155", color: "#94a3b8", border: "none", borderRadius: 6, cursor: "pointer", fontSize: 11 }}>Reset</button>
          </div>
        </Card>

        <Tabs tabs={TABS} active={tab} onChange={setTab} />

        {/* ─── OVERVIEW ─── */}
        {tab === "overview" && (
          <>
            <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit, minmax(130px, 1fr))", gap: 12, marginBottom: 16 }}>
              {[
                { label: "Participants", value: data.length },
                { label: "Mean Age", value: fmt(mean(col("age")), 1), sub: `SD ${fmt(sd(col("age")), 1)}` },
                { label: "Male %", value: `${data.length ? ((data.filter((r) => String(r.gender) === "1").length / data.length) * 100).toFixed(0) : 0}%` },
                { label: "Mean CES-D", value: fmt(mean(col("cesd_total")), 1) },
                { label: "Mean PHQ-9", value: fmt(mean(col("phq9_total")), 1) },
                { label: "Mean GAD-7", value: fmt(mean(col("gad7_total")), 1) },
                { label: "Mean K10", value: fmt(mean(col("k10_total")), 1) },
                { label: "Mean WEMWBS", value: fmt(mean(col("wemwbs_total")), 1) },
              ].map((s, i) => (
                <Card key={i} style={{ padding: 10 }}><Stat {...s} /></Card>
              ))}
            </div>
            <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 16 }}>
              <Card title="CMD Prevalence Summary">
                <div style={{ display: "flex", justifyContent: "flex-end", marginBottom: 8 }}>
                  <ExportBtn onClick={() => {
                    const rows = [["Scale", "N", "Cases", "Prevalence%", "Mean", "SD", "Median"]];
                    cmdPrev.forEach((c) => rows.push([c.label, c.n, c.nAbove, c.prev, fmt(c.mean), fmt(c.sd), fmt(c.med)]));
                    downloadCSV(rows, "cmd_prevalence.csv");
                  }} />
                </div>
                <ResponsiveContainer width="100%" height={250}>
                  <BarChart data={cmdPrev} margin={{ top: 5, right: 20, bottom: 5, left: 0 }}>
                    <CartesianGrid strokeDasharray="3 3" stroke="#1e293b" />
                    <XAxis dataKey="label" tick={{ fill: "#94a3b8", fontSize: 12 }} />
                    <YAxis tick={{ fill: "#94a3b8", fontSize: 12 }} label={{ value: "Prevalence %", angle: -90, position: "insideLeft", fill: "#64748b", fontSize: 11 }} />
                    <Tooltip content={<TT />} />
                    <Bar dataKey="prev" name="Prevalence %" fill="#0f766e" radius={[6, 6, 0, 0]} />
                  </BarChart>
                </ResponsiveContainer>
              </Card>
              <Card title="Demographics Distribution">
                <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 12 }}>
                  {[
                    { title: "Gender", data: Object.entries(GENDER_MAP).map(([k, v]) => ({ name: v, value: data.filter((r) => String(r.gender) === k).length })).filter((d) => d.value > 0) },
                    { title: "Education", data: Object.entries(EDUCATION_MAP).map(([k, v]) => ({ name: v, value: data.filter((r) => String(r.education_level) === k).length })).filter((d) => d.value > 0) },
                  ].map((chart, ci) => (
                    <div key={ci}>
                      <div style={{ fontSize: 12, color: "#94a3b8", textAlign: "center", marginBottom: 4 }}>{chart.title}</div>
                      <ResponsiveContainer width="100%" height={160}>
                        <PieChart>
                          <Pie data={chart.data} dataKey="value" nameKey="name" cx="50%" cy="50%" outerRadius={55} label={({ name, percent }) => `${name} ${(percent * 100).toFixed(0)}%`} labelLine={{ stroke: "#475569" }} style={{ fontSize: 10 }}>
                            {chart.data.map((_, i) => <Cell key={i} fill={PALETTE[i % PALETTE.length]} />)}
                          </Pie>
                          <Tooltip />
                        </PieChart>
                      </ResponsiveContainer>
                    </div>
                  ))}
                </div>
              </Card>
            </div>
            {/* Summary stats table */}
            <Card title="Summary Statistics — Key Scales">
              <ExportBtn label="Export Summary" onClick={() => {
                const keys = ["cesd_total", "phq9_total", "gad7_total", "k10_total", "wemwbs_total", "sarrs_total_stress_score", "css_total", "passq_total", "soa_positive_agency_total", "soa_negative_agency_total", "sleep_hours"];
                const rows = [["Variable", "N", "Mean", "SD", "Median", "Q1", "Q3", "Min", "Max"]];
                keys.forEach((k) => {
                  const v = validNums(col(k));
                  const [q1, q3] = iqr(col(k));
                  rows.push([k, v.length, fmt(mean(col(k))), fmt(sd(col(k))), fmt(median(col(k))), fmt(q1), fmt(q3), v.length ? fmt(Math.min(...v)) : "—", v.length ? fmt(Math.max(...v)) : "—"]);
                });
                downloadCSV(rows, "summary_statistics.csv");
              }} />
              <div style={{ overflowX: "auto", marginTop: 8 }}>
                <table style={{ width: "100%", borderCollapse: "collapse", fontSize: 12 }}>
                  <thead>
                    <tr style={{ borderBottom: "2px solid #334155" }}>
                      {["Variable", "N", "Mean", "SD", "Median", "Q1", "Q3"].map((h) => (
                        <th key={h} style={{ padding: "6px 8px", textAlign: "left", color: "#94a3b8", fontWeight: 600, fontSize: 11 }}>{h}</th>
                      ))}
                    </tr>
                  </thead>
                  <tbody>
                    {["cesd_total", "phq9_total", "gad7_total", "k10_total", "wemwbs_total", "sarrs_total_stress_score", "sarrs_n_events", "css_total", "passq_total", "soa_positive_agency_total", "soa_negative_agency_total", "sleep_hours"].map((k) => {
                      const v = validNums(col(k));
                      const [q1, q3] = iqr(col(k));
                      return (
                        <tr key={k} style={{ borderBottom: "1px solid #1e293b" }}>
                          <td style={{ padding: "5px 8px", color: "#5eead4", fontWeight: 500 }}>{k}</td>
                          <td style={{ padding: "5px 8px" }}>{v.length}</td>
                          <td style={{ padding: "5px 8px" }}>{fmt(mean(col(k)))}</td>
                          <td style={{ padding: "5px 8px" }}>{fmt(sd(col(k)))}</td>
                          <td style={{ padding: "5px 8px" }}>{fmt(median(col(k)))}</td>
                          <td style={{ padding: "5px 8px" }}>{fmt(q1)}</td>
                          <td style={{ padding: "5px 8px" }}>{fmt(q3)}</td>
                        </tr>
                      );
                    })}
                  </tbody>
                </table>
              </div>
            </Card>
          </>
        )}

        {/* ─── CMD PREVALENCE ─── */}
        {tab === "cmd" && (
          <>
            <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 16 }}>
              {/* PHQ-9 Severity */}
              <Card title="PHQ-9 Severity Distribution">
                <ResponsiveContainer width="100%" height={220}>
                  <BarChart data={severityDist("phq9_total", PHQ9_SEVERITY)}>
                    <CartesianGrid strokeDasharray="3 3" stroke="#1e293b" />
                    <XAxis dataKey="label" tick={{ fill: "#94a3b8", fontSize: 11 }} />
                    <YAxis tick={{ fill: "#94a3b8", fontSize: 11 }} />
                    <Tooltip content={<TT />} />
                    <Bar dataKey="count" name="Count">
                      {PHQ9_SEVERITY.map((s, i) => <Cell key={i} fill={s.color} />)}
                    </Bar>
                  </BarChart>
                </ResponsiveContainer>
              </Card>
              {/* GAD-7 Severity */}
              <Card title="GAD-7 Severity Distribution">
                <ResponsiveContainer width="100%" height={220}>
                  <BarChart data={severityDist("gad7_total", GAD7_SEVERITY)}>
                    <CartesianGrid strokeDasharray="3 3" stroke="#1e293b" />
                    <XAxis dataKey="label" tick={{ fill: "#94a3b8", fontSize: 11 }} />
                    <YAxis tick={{ fill: "#94a3b8", fontSize: 11 }} />
                    <Tooltip content={<TT />} />
                    <Bar dataKey="count" name="Count">
                      {GAD7_SEVERITY.map((s, i) => <Cell key={i} fill={s.color} />)}
                    </Bar>
                  </BarChart>
                </ResponsiveContainer>
              </Card>
              {/* K10 Severity */}
              <Card title="K10 Severity Distribution">
                <ResponsiveContainer width="100%" height={220}>
                  <BarChart data={severityDist("k10_total", K10_SEVERITY)}>
                    <CartesianGrid strokeDasharray="3 3" stroke="#1e293b" />
                    <XAxis dataKey="label" tick={{ fill: "#94a3b8", fontSize: 11 }} />
                    <YAxis tick={{ fill: "#94a3b8", fontSize: 11 }} />
                    <Tooltip content={<TT />} />
                    <Bar dataKey="count" name="Count">
                      {K10_SEVERITY.map((s, i) => <Cell key={i} fill={s.color} />)}
                    </Bar>
                  </BarChart>
                </ResponsiveContainer>
              </Card>
              {/* WEMWBS Distribution */}
              <Card title="WEMWBS Wellbeing — Histogram">
                <ResponsiveContainer width="100%" height={220}>
                  <BarChart data={(() => {
                    const vals = validNums(col("wemwbs_total"));
                    const bins = Array.from({ length: 8 }, (_, i) => ({ bin: `${14 + i * 8}-${21 + i * 8}`, count: 0 }));
                    vals.forEach((v) => { const idx = Math.min(7, Math.floor((v - 14) / 8)); if (bins[idx]) bins[idx].count++; });
                    return bins;
                  })()}>
                    <CartesianGrid strokeDasharray="3 3" stroke="#1e293b" />
                    <XAxis dataKey="bin" tick={{ fill: "#94a3b8", fontSize: 10 }} />
                    <YAxis tick={{ fill: "#94a3b8", fontSize: 11 }} />
                    <Tooltip content={<TT />} />
                    <Bar dataKey="count" name="Count" fill="#7c3aed" radius={[4, 4, 0, 0]} />
                  </BarChart>
                </ResponsiveContainer>
              </Card>
            </div>
            {/* Prevalence Table */}
            <Card title="CMD Clinical Cut-off Prevalence" style={{ marginTop: 16 }}>
              <ExportBtn label="Export Prevalence" onClick={() => {
                const rows = [["Scale", "Cut-off", "N Valid", "N Above", "Prevalence %", "Mean", "SD", "Median"]];
                cmdPrev.forEach((c) => rows.push([c.label, c.cutoff, c.n, c.nAbove, c.prev, fmt(c.mean), fmt(c.sd), fmt(c.med)]));
                downloadCSV(rows, "cmd_prevalence_detail.csv");
              }} />
              <table style={{ width: "100%", borderCollapse: "collapse", fontSize: 12, marginTop: 10 }}>
                <thead>
                  <tr style={{ borderBottom: "2px solid #334155" }}>
                    {["Scale", "Description", "Cut-off", "N", "Cases", "Prevalence %", "Mean (SD)", "Median"].map((h) => (
                      <th key={h} style={{ padding: "6px 8px", textAlign: "left", color: "#94a3b8", fontWeight: 600, fontSize: 11 }}>{h}</th>
                    ))}
                  </tr>
                </thead>
                <tbody>
                  {cmdPrev.map((c) => (
                    <tr key={c.key} style={{ borderBottom: "1px solid #1e293b" }}>
                      <td style={{ padding: "5px 8px", fontWeight: 600, color: "#5eead4" }}>{c.label}</td>
                      <td style={{ padding: "5px 8px" }}>{c.desc}</td>
                      <td style={{ padding: "5px 8px" }}>
                        <div style={{ display: "flex", alignItems: "center", gap: 3 }}>
                          <span style={{ color: "#94a3b8" }}>{c.invert ? "<" : "≥"}</span>
                          <input type="number" value={cmdCutoffs[c.key] ?? c.cutoff}
                            onChange={(e) => setCmdCutoffs((prev) => ({ ...prev, [c.key]: Number(e.target.value) }))}
                            style={{ width: 42, background: "#0f172a", color: "#e2e8f0", border: "1px solid #334155", borderRadius: 4, padding: "2px 4px", fontSize: 12, textAlign: "center" }}
                          />
                        </div>
                      </td>
                      <td style={{ padding: "5px 8px" }}>{c.n}</td>
                      <td style={{ padding: "5px 8px", fontWeight: 600, color: "#f97316" }}>{c.nAbove}</td>
                      <td style={{ padding: "5px 8px", fontWeight: 700, color: "#ef4444" }}>{c.prev}%</td>
                      <td style={{ padding: "5px 8px" }}>{fmt(c.mean)} ({fmt(c.sd)})</td>
                      <td style={{ padding: "5px 8px" }}>{fmt(c.med)}</td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </Card>
            {/* CES-D Item Means */}
            <Card title="CES-D Item Mean Scores" style={{ marginTop: 16 }}>
              <ResponsiveContainer width="100%" height={260}>
                <BarChart data={[
                  "cesd_bothered_by_things", "cesd_poor_appetite", "cesd_could_not_shake_blues", "cesd_as_good_as_others", "cesd_trouble_concentrating",
                  "cesd_felt_depressed", "cesd_everything_an_effort", "cesd_felt_hopeful", "cesd_life_a_failure", "cesd_felt_fearful",
                  "cesd_restless_sleep", "cesd_was_happy", "cesd_talked_less", "cesd_felt_lonely", "cesd_people_unfriendly",
                  "cesd_enjoyed_life", "cesd_crying_spells", "cesd_felt_sad", "cesd_people_dislike_me", "cesd_could_not_get_going"
                ].map((k) => ({ name: k.replace("cesd_", "").replace(/_/g, " "), mean: mean(col(k)) || 0 }))}
                  layout="vertical" margin={{ left: 120, right: 20 }}>
                  <CartesianGrid strokeDasharray="3 3" stroke="#1e293b" />
                  <XAxis type="number" tick={{ fill: "#94a3b8", fontSize: 11 }} domain={[0, 3]} />
                  <YAxis type="category" dataKey="name" tick={{ fill: "#94a3b8", fontSize: 10 }} width={115} />
                  <Tooltip content={<TT />} />
                  <Bar dataKey="mean" name="Mean" fill="#0e7490" radius={[0, 4, 4, 0]} />
                </BarChart>
              </ResponsiveContainer>
            </Card>
          </>
        )}

        {/* ─── RISK FACTORS ─── */}
        {tab === "risk" && (
          <>
            <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 16 }}>
              <Card title="SARRS Life Stress">
                <div style={{ display: "flex", gap: 24, marginBottom: 12 }}>
                  <Stat label="Mean Total Score" value={fmt(mean(col("sarrs_total_stress_score")), 0)} sub={`SD ${fmt(sd(col("sarrs_total_stress_score")), 0)}`} />
                  <Stat label="Mean N Events" value={fmt(mean(col("sarrs_n_events")), 1)} sub={`SD ${fmt(sd(col("sarrs_n_events")), 1)}`} />
                </div>
                <ResponsiveContainer width="100%" height={200}>
                  <BarChart data={(() => {
                    const vals = validNums(col("sarrs_total_stress_score"));
                    const bins = [
                      { bin: "<150 (Low)", count: vals.filter((v) => v < 150).length },
                      { bin: "150-299 (Moderate)", count: vals.filter((v) => v >= 150 && v < 300).length },
                      { bin: "≥300 (High)", count: vals.filter((v) => v >= 300).length },
                    ];
                    return bins;
                  })()}>
                    <CartesianGrid strokeDasharray="3 3" stroke="#1e293b" />
                    <XAxis dataKey="bin" tick={{ fill: "#94a3b8", fontSize: 11 }} />
                    <YAxis tick={{ fill: "#94a3b8", fontSize: 11 }} />
                    <Tooltip content={<TT />} />
                    <Bar dataKey="count" name="Count">
                      <Cell fill="#22c55e" /><Cell fill="#eab308" /><Cell fill="#ef4444" />
                    </Bar>
                  </BarChart>
                </ResponsiveContainer>
              </Card>
              <Card title="Sense of Agency">
                <div style={{ display: "flex", gap: 24, marginBottom: 12 }}>
                  <Stat label="Positive Agency" value={fmt(mean(col("soa_positive_agency_total")), 1)} sub={`SD ${fmt(sd(col("soa_positive_agency_total")), 1)}`} />
                  <Stat label="Negative Agency" value={fmt(mean(col("soa_negative_agency_total")), 1)} sub={`SD ${fmt(sd(col("soa_negative_agency_total")), 1)}`} />
                </div>
                <ResponsiveContainer width="100%" height={200}>
                  <BarChart data={[
                    { name: "Positive", mean: mean(col("soa_positive_agency_total")) || 0 },
                    { name: "Negative", mean: mean(col("soa_negative_agency_total")) || 0 },
                  ]}>
                    <CartesianGrid strokeDasharray="3 3" stroke="#1e293b" />
                    <XAxis dataKey="name" tick={{ fill: "#94a3b8", fontSize: 12 }} />
                    <YAxis tick={{ fill: "#94a3b8", fontSize: 11 }} />
                    <Tooltip content={<TT />} />
                    <Bar dataKey="mean" name="Mean Score" fill="#7c3aed" radius={[6, 6, 0, 0]} />
                  </BarChart>
                </ResponsiveContainer>
              </Card>
              <Card title="Career Satisfaction (CSS)">
                <ResponsiveContainer width="100%" height={220}>
                  <BarChart data={[
                    "css_satisfied_career_success", "css_satisfied_overall_career_goals", "css_satisfied_income_goals",
                    "css_satisfied_career_advancement", "css_satisfied_new_skills_development"
                  ].map((k) => ({ name: k.replace("css_satisfied_", "").replace(/_/g, " "), mean: mean(col(k)) || 0 }))}
                    margin={{ left: 10 }}>
                    <CartesianGrid strokeDasharray="3 3" stroke="#1e293b" />
                    <XAxis dataKey="name" tick={{ fill: "#94a3b8", fontSize: 10 }} />
                    <YAxis tick={{ fill: "#94a3b8", fontSize: 11 }} />
                    <Tooltip content={<TT />} />
                    <Bar dataKey="mean" name="Mean" fill="#0f766e" radius={[6, 6, 0, 0]} />
                  </BarChart>
                </ResponsiveContainer>
              </Card>
              <Card title="Sleep Hours Distribution">
                <Stat label="Mean Sleep" value={`${fmt(mean(col("sleep_hours")), 1)}h`} sub={`SD ${fmt(sd(col("sleep_hours")), 1)}`} />
                <ResponsiveContainer width="100%" height={180}>
                  <BarChart data={(() => {
                    const vals = validNums(col("sleep_hours"));
                    const bins = Array.from({ length: 8 }, (_, i) => ({ bin: `${3 + i}h`, count: vals.filter((v) => Math.round(v) === 3 + i).length }));
                    return bins;
                  })()}>
                    <CartesianGrid strokeDasharray="3 3" stroke="#1e293b" />
                    <XAxis dataKey="bin" tick={{ fill: "#94a3b8", fontSize: 11 }} />
                    <YAxis tick={{ fill: "#94a3b8", fontSize: 11 }} />
                    <Tooltip content={<TT />} />
                    <Bar dataKey="count" name="Count" fill="#0e7490" radius={[4, 4, 0, 0]} />
                  </BarChart>
                </ResponsiveContainer>
              </Card>
            </div>
            <Card title="Risk Factor Export" style={{ marginTop: 16 }}>
              <ExportBtn label="Export Risk Factor Summary" onClick={() => {
                const keys = ["sarrs_total_stress_score", "sarrs_n_events", "soa_positive_agency_total", "soa_negative_agency_total", "css_total", "passq_total", "sleep_hours", "weekly_work_hours"];
                const rows = [["Variable", "N", "Mean", "SD", "Median", "Min", "Max"]];
                keys.forEach((k) => {
                  const v = validNums(col(k));
                  rows.push([k, v.length, fmt(mean(col(k))), fmt(sd(col(k))), fmt(median(col(k))), v.length ? fmt(Math.min(...v)) : "—", v.length ? fmt(Math.max(...v)) : "—"]);
                });
                downloadCSV(rows, "risk_factors_summary.csv");
              }} />
            </Card>
          </>
        )}

        {/* ─── COPING & SUPPORT ─── */}
        {tab === "coping" && (
          <>
            {/* COPE mode toggle */}
            <div style={{ display: "flex", alignItems: "center", gap: 10, marginBottom: 12 }}>
              <span style={{ fontSize: 12, color: "#94a3b8", fontWeight: 600 }}>Brief-COPE grouping:</span>
              <div style={{ display: "flex", gap: 3, background: "#1e293b", borderRadius: 8, padding: 3, border: "1px solid #334155" }}>
                {[{ id: "14", label: "14 Subscales" }, { id: "3", label: "3-Class" }, { id: "all", label: "All (14 + 3)" }].map((m) => (
                  <button key={m.id} onClick={() => setCopeMode(m.id)}
                    style={{
                      padding: "5px 14px", fontSize: 12, fontWeight: copeMode === m.id ? 700 : 500,
                      background: copeMode === m.id ? "#0f766e" : "transparent",
                      color: copeMode === m.id ? "#fff" : "#94a3b8",
                      border: "none", borderRadius: 6, cursor: "pointer",
                    }}>{m.label}</button>
                ))}
              </div>
              <span style={{ fontSize: 11, color: "#475569" }}>
                {copeMode === "3" ? "Problem-Focused, Emotion-Focused, Avoidant" : copeMode === "all" ? "14 subscales + 3 composite dimensions" : "All 14 original Brief-COPE subscales"}
              </span>
            </div>
            <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 16 }}>
              <Card title={copeMode === "3" ? "Brief-COPE 3-Class Means" : copeMode === "all" ? "Brief-COPE — All Subscales + Composites" : "Brief-COPE Subscale Means"}>
                <ExportBtn label="Export COPE" onClick={() => {
                  const rows = [["Subscale", "N", "Mean", "SD"]];
                  activeCope.forEach((s) => rows.push([s.label, validNums(col(s.key)).length, fmt(mean(col(s.key))), fmt(sd(col(s.key)))]));
                  downloadCSV(rows, copeMode === "3" ? "cope_3class.csv" : copeMode === "all" ? "cope_all.csv" : "cope_subscales.csv");
                }} />
                <ResponsiveContainer width="100%" height={copeMode === "3" ? 160 : copeMode === "all" ? 440 : 350}>
                  <BarChart data={activeCope.map((s) => ({ name: s.label, mean: mean(col(s.key)) || 0 }))} layout="vertical" margin={{ left: 160 }}>
                    <CartesianGrid strokeDasharray="3 3" stroke="#1e293b" />
                    <XAxis type="number" tick={{ fill: "#94a3b8", fontSize: 11 }} domain={[0, "auto"]} />
                    <YAxis type="category" dataKey="name" tick={{ fill: "#94a3b8", fontSize: 11 }} width={155} />
                    <Tooltip content={<TT />} />
                    <Bar dataKey="mean" name="Mean" fill="#0f766e" radius={[0, 6, 6, 0]} />
                  </BarChart>
                </ResponsiveContainer>
              </Card>
              <Card title="PASS-Q Social Support">
                <ExportBtn label="Export PASS-Q" onClick={() => {
                  const rows = [["Subscale", "N", "Mean", "SD"]];
                  PASSQ_SUBS.forEach((s) => rows.push([s.label, validNums(col(s.key)).length, fmt(mean(col(s.key))), fmt(sd(col(s.key)))]));
                  rows.push(["Total", validNums(col("passq_total")).length, fmt(mean(col("passq_total"))), fmt(sd(col("passq_total")))]);
                  downloadCSV(rows, "passq_support.csv");
                }} />
                <ResponsiveContainer width="100%" height={280}>
                  <RadarChart cx="50%" cy="50%" outerRadius={90} data={PASSQ_SUBS.map((s) => ({ subject: s.label, value: mean(col(s.key)) || 0 }))}>
                    <PolarGrid stroke="#334155" />
                    <PolarAngleAxis dataKey="subject" tick={{ fill: "#94a3b8", fontSize: 11 }} />
                    <PolarRadiusAxis tick={{ fill: "#64748b", fontSize: 10 }} />
                    <Radar name="Mean" dataKey="value" stroke="#0f766e" fill="#0f766e" fillOpacity={0.3} />
                    <Tooltip />
                  </RadarChart>
                </ResponsiveContainer>
                <div style={{ display: "flex", justifyContent: "center", gap: 20, fontSize: 12, color: "#94a3b8", marginTop: 8 }}>
                  <span>Total: <strong style={{ color: "#5eead4" }}>{fmt(mean(col("passq_total")), 1)}</strong></span>
                </div>
              </Card>
            </div>
          </>
        )}

        {/* ─── WEIGHT MAKING ─── */}
        {tab === "weight" && (
          <>
            <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 16 }}>
              <Card title="Weight Cutting Prevalence">
                {(() => {
                  const yes = data.filter((r) => num(r.ever_cut_weight) === 2).length;
                  const no = data.filter((r) => num(r.ever_cut_weight) === 1).length;
                  const total = yes + no;
                  return (
                    <>
                      <div style={{ display: "flex", gap: 20, marginBottom: 12 }}>
                        <Stat label="Ever Cut Weight" value={total ? `${((yes / total) * 100).toFixed(0)}%` : "—"} sub={`${yes} of ${total}`} />
                      </div>
                      <ResponsiveContainer width="100%" height={180}>
                        <PieChart>
                          <Pie data={[{ name: "Yes", value: yes }, { name: "No", value: no }]} dataKey="value" cx="50%" cy="50%" outerRadius={65} label={({ name, percent }) => `${name} ${(percent * 100).toFixed(0)}%`}>
                            <Cell fill="#ef4444" /><Cell fill="#22c55e" />
                          </Pie>
                          <Tooltip />
                        </PieChart>
                      </ResponsiveContainer>
                    </>
                  );
                })()}
              </Card>
              <Card title="Weight Making Methods (% using)">
                <ExportBtn label="Export Methods" onClick={() => {
                  const rows = [["Method", "N Using", "% Using"]];
                  const total = data.length;
                  WM_METHODS.forEach((m) => {
                    const n = data.filter((r) => num(r[m.key]) === 1 || String(r[m.key]) === "1").length;
                    rows.push([m.label, n, total ? ((n / total) * 100).toFixed(1) : 0]);
                  });
                  downloadCSV(rows, "weight_methods.csv");
                }} />
                <ResponsiveContainer width="100%" height={350}>
                  <BarChart data={WM_METHODS.map((m) => {
                    const n = data.filter((r) => num(r[m.key]) === 1 || String(r[m.key]) === "1").length;
                    return { name: m.label, pct: data.length ? (n / data.length) * 100 : 0 };
                  }).sort((a, b) => b.pct - a.pct)} layout="vertical" margin={{ left: 130 }}>
                    <CartesianGrid strokeDasharray="3 3" stroke="#1e293b" />
                    <XAxis type="number" tick={{ fill: "#94a3b8", fontSize: 11 }} unit="%" />
                    <YAxis type="category" dataKey="name" tick={{ fill: "#94a3b8", fontSize: 11 }} width={125} />
                    <Tooltip content={<TT />} />
                    <Bar dataKey="pct" name="% Using" fill="#db2777" radius={[0, 6, 6, 0]} />
                  </BarChart>
                </ResponsiveContainer>
              </Card>
              <Card title="Weight Stats">
                <table style={{ width: "100%", borderCollapse: "collapse", fontSize: 12 }}>
                  <thead>
                    <tr style={{ borderBottom: "2px solid #334155" }}>
                      {["Metric", "N", "Mean", "SD", "Median"].map((h) => (
                        <th key={h} style={{ padding: "5px 8px", textAlign: "left", color: "#94a3b8", fontWeight: 600, fontSize: 11 }}>{h}</th>
                      ))}
                    </tr>
                  </thead>
                  <tbody>
                    {[
                      { key: "non_race_weight_lbs", label: "Non-Race Weight (lbs)" },
                      { key: "race_morning_weight_lbs", label: "Race Morning Weight (lbs)" },
                      { key: "max_weight_lost_prep_lbs", label: "Max Weight Lost (lbs)" },
                    ].map((r) => (
                      <tr key={r.key} style={{ borderBottom: "1px solid #1e293b" }}>
                        <td style={{ padding: "5px 8px", color: "#5eead4" }}>{r.label}</td>
                        <td style={{ padding: "5px 8px" }}>{validNums(col(r.key)).length}</td>
                        <td style={{ padding: "5px 8px" }}>{fmt(mean(col(r.key)))}</td>
                        <td style={{ padding: "5px 8px" }}>{fmt(sd(col(r.key)))}</td>
                        <td style={{ padding: "5px 8px" }}>{fmt(median(col(r.key)))}</td>
                      </tr>
                    ))}
                  </tbody>
                </table>
              </Card>
              <Card title="Cutting Frequency">
                <ResponsiveContainer width="100%" height={200}>
                  <BarChart data={Object.entries(WEIGHT_FREQ_MAP).map(([k, v]) => ({
                    name: v, count: data.filter((r) => String(r.weight_cutting_frequency) === k).length,
                  }))}>
                    <CartesianGrid strokeDasharray="3 3" stroke="#1e293b" />
                    <XAxis dataKey="name" tick={{ fill: "#94a3b8", fontSize: 10 }} />
                    <YAxis tick={{ fill: "#94a3b8", fontSize: 11 }} />
                    <Tooltip content={<TT />} />
                    <Bar dataKey="count" name="Count" fill="#ea580c" radius={[6, 6, 0, 0]} />
                  </BarChart>
                </ResponsiveContainer>
              </Card>
            </div>
          </>
        )}

        {/* ─── LOGISTIC REGRESSION ─── */}
        {tab === "logistic" && (
          <>
            <Card title="Logistic Regression — Odds Ratios for CMD Risk">
              <div style={{ fontSize: 12, color: "#94a3b8", marginBottom: 12 }}>
                Binary outcome via clinical cut-off. <strong>Univariate</strong> runs each predictor alone (unadjusted OR). <strong>Multivariable</strong> enters all predictors simultaneously (adjusted OR). <strong>Both</strong> shows side-by-side comparison.
              </div>
              <div style={{ display: "grid", gridTemplateColumns: "1fr 2fr", gap: 16 }}>
                <div>
                  <div style={{ marginBottom: 12 }}>
                    <label style={{ fontSize: 11, color: "#5eead4", fontWeight: 600, display: "block", marginBottom: 4 }}>Outcome (Binary)</label>
                    <select value={lrOutcome} onChange={(e) => setLrOutcome(e.target.value)}
                      style={{ width: "100%", background: "#0f172a", color: "#e2e8f0", border: "1px solid #334155", borderRadius: 6, padding: "7px 8px", fontSize: 12 }}>
                      {binaryOutcomes.map((o) => <option key={o.key} value={o.key}>{o.label}</option>)}
                    </select>
                  </div>
                  {/* ── Editable Clinical Cut-offs ── */}
                  <div style={{ marginBottom: 12, background: "#0f172a", borderRadius: 8, padding: 10, border: "1px solid #334155" }}>
                    <label style={{ fontSize: 11, color: "#5eead4", fontWeight: 600, display: "block", marginBottom: 6 }}>Clinical Cut-off Thresholds</label>
                    <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: "6px 10px", fontSize: 11 }}>
                      {CMD_SCALES.map((s) => (
                        <div key={s.key} style={{ display: "flex", alignItems: "center", gap: 4 }}>
                          <span style={{ color: "#94a3b8", minWidth: 56 }}>{s.label} {s.invert ? "<" : "≥"}</span>
                          <input type="number" value={cmdCutoffs[s.key] ?? s.cutoff}
                            onChange={(e) => setCmdCutoffs((prev) => ({ ...prev, [s.key]: Number(e.target.value) }))}
                            style={{ width: 48, background: "#1e293b", color: "#e2e8f0", border: "1px solid #334155", borderRadius: 4, padding: "3px 5px", fontSize: 11, textAlign: "center" }}
                          />
                        </div>
                      ))}
                    </div>
                    <div style={{ fontSize: 10, color: "#475569", marginTop: 4 }}>Defaults: CES-D ≥16, PHQ-9 ≥10, GAD-7 ≥10, K10 ≥20, WEMWBS &lt;40</div>
                  </div>
                  <div style={{ marginBottom: 12 }}>
                    <label style={{ fontSize: 11, color: "#5eead4", fontWeight: 600, display: "block", marginBottom: 6 }}>Model Type</label>
                    <div style={{ display: "flex", gap: 4, background: "#0f172a", borderRadius: 8, padding: 3, border: "1px solid #334155" }}>
                      {[{ id: "univariate", label: "Univariate" }, { id: "multivariable", label: "Multivariable" }, { id: "both", label: "Both" }].map((m) => (
                        <button key={m.id} onClick={() => setLrMode(m.id)}
                          style={{
                            flex: 1, padding: "6px 4px", fontSize: 11, fontWeight: lrMode === m.id ? 700 : 500, border: "none", borderRadius: 6, cursor: "pointer",
                            background: lrMode === m.id ? "#0f766e" : "transparent", color: lrMode === m.id ? "#fff" : "#94a3b8",
                          }}>{m.label}</button>
                      ))}
                    </div>
                  </div>
                  <div style={{ marginBottom: 12 }}>
                    <label style={{ fontSize: 11, color: "#5eead4", fontWeight: 600, display: "block", marginBottom: 4 }}>Predictors</label>
                    {ACTIVE_RISK_GROUPS.map((g) => (
                      <div key={g.label} style={{ marginBottom: 6 }}>
                        <div style={{ fontSize: 10, color: "#64748b", fontWeight: 700, textTransform: "uppercase", marginBottom: 2, letterSpacing: 0.5 }}>{g.label}</div>
                        <CheckList items={g.vars} selected={lrPredictors} onChange={setLrPredictors} columns={1} />
                      </div>
                    ))}
                  </div>
                  <button onClick={runLogistic}
                    style={{ width: "100%", padding: "10px", background: "linear-gradient(135deg, #0f766e, #0e7490)", color: "#fff", border: "none", borderRadius: 8, fontWeight: 700, fontSize: 13, cursor: "pointer" }}>
                    Run Logistic Regression
                  </button>
                  {(lrResult || lrUniResults) && (
                    <button onClick={exportLogisticWord}
                      style={{ width: "100%", padding: "8px", marginTop: 8, background: "#1e293b", color: "#5eead4", border: "1px solid #334155", borderRadius: 8, fontWeight: 600, fontSize: 12, cursor: "pointer" }}>
                      ⬇ Export to Word (.docx)
                    </button>
                  )}
                </div>
                <div>
                  {/* ── Univariate Results ── */}
                  {lrUniResults && (lrMode === "univariate" || lrMode === "both") && (
                    <div style={{ marginBottom: 20 }}>
                      <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 8 }}>
                        <h4 style={{ fontSize: 14, color: "#5eead4", margin: 0, fontWeight: 700 }}>Univariate (Unadjusted) ORs</h4>
                        <ExportBtn label="Export Univariate" onClick={() => {
                          const rows = [["Variable", "N", "β", "SE", "z", "p-value", "OR", "OR 95% CI Lower", "OR 95% CI Upper", "Pseudo R²", "Sig"]];
                          lrUniResults.forEach((r) => {
                            if (r.error) { rows.push([r.label, "", "", "", "", "", "", "", "", "", r.error]); }
                            else { rows.push([r.label, r.n, fmt(r.beta, 4), fmt(r.se, 4), fmt(r.z, 3), fmtP(r.p), fmt(r.or, 3), fmt(r.ciLow, 3), fmt(r.ciHigh, 3), fmt(r.pseudoR2, 4), pStar(r.p)]); }
                          });
                          downloadCSV(rows, "logistic_univariate.csv");
                        }} />
                      </div>
                      <div style={{ overflowX: "auto" }}>
                        <table style={{ width: "100%", borderCollapse: "collapse", fontSize: 12 }}>
                          <thead>
                            <tr style={{ borderBottom: "2px solid #334155" }}>
                              {["Variable", "N", "β", "SE", "z", "p", "OR", "95% CI", "R²", ""].map((h) => (
                                <th key={h} style={{ padding: "6px 6px", textAlign: "left", color: "#94a3b8", fontWeight: 600, fontSize: 11 }}>{h}</th>
                              ))}
                            </tr>
                          </thead>
                          <tbody>
                            {lrUniResults.map((r, i) => (
                              <tr key={i} style={{ borderBottom: "1px solid #1e293b", background: !r.error && r.p < 0.05 ? "rgba(15,118,110,.1)" : "transparent" }}>
                                <td style={{ padding: "5px 6px", fontWeight: 500, color: "#e2e8f0" }}>{r.label}</td>
                                {r.error ? (
                                  <td colSpan={8} style={{ padding: "5px 6px", color: "#ef4444", fontStyle: "italic" }}>{r.error}</td>
                                ) : (<>
                                  <td style={{ padding: "5px 6px" }}>{r.n}</td>
                                  <td style={{ padding: "5px 6px" }}>{fmt(r.beta, 3)}</td>
                                  <td style={{ padding: "5px 6px" }}>{fmt(r.se, 3)}</td>
                                  <td style={{ padding: "5px 6px" }}>{fmt(r.z, 2)}</td>
                                  <td style={{ padding: "5px 6px", fontWeight: r.p < 0.05 ? 700 : 400, color: r.p < 0.05 ? "#5eead4" : "#94a3b8" }}>{fmtP(r.p)}</td>
                                  <td style={{ padding: "5px 6px", fontWeight: 600 }}>{fmt(r.or, 3)}</td>
                                  <td style={{ padding: "5px 6px", fontSize: 11 }}>[{fmt(r.ciLow, 2)}, {fmt(r.ciHigh, 2)}]</td>
                                  <td style={{ padding: "5px 6px", fontSize: 11, color: "#64748b" }}>{fmt(r.pseudoR2, 3)}</td>
                                  <td style={{ padding: "5px 6px", color: "#eab308" }}>{pStar(r.p)}</td>
                                </>)}
                              </tr>
                            ))}
                          </tbody>
                        </table>
                      </div>
                    </div>
                  )}

                  {/* ── Multivariable Results ── */}
                  {lrResult?.error && (lrMode === "multivariable" || lrMode === "both") && (
                    <div style={{ color: "#ef4444", padding: 12, background: "#1c1917", borderRadius: 8, fontSize: 13, marginBottom: 12 }}>⚠ {lrResult.error}</div>
                  )}
                  {lrResult && !lrResult.error && (lrMode === "multivariable" || lrMode === "both") && (
                    <>
                      <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 8 }}>
                        <h4 style={{ fontSize: 14, color: "#5eead4", margin: 0, fontWeight: 700 }}>Multivariable (Adjusted) ORs</h4>
                      </div>
                      <div style={{ display: "flex", gap: 16, marginBottom: 12, fontSize: 12, color: "#94a3b8" }}>
                        <span>N = {lrResult.n}</span>
                        <span>Pseudo R² = {fmt(lrResult.pseudoR2, 3)}</span>
                        <span>Log-Likelihood = {fmt(lrResult.ll, 1)}</span>
                        <ExportBtn label="Export Multivariable" onClick={() => {
                          const rows = [["Variable", "Coefficient", "SE", "z", "p-value", "OR", "OR 95% CI Lower", "OR 95% CI Upper", "Sig"]];
                          lrResult.names.forEach((n, i) => rows.push([n, fmt(lrResult.coefficients[i], 4), fmt(lrResult.se[i], 4), fmt(lrResult.z[i], 3), fmtP(lrResult.pvals[i]), fmt(lrResult.or[i], 3), fmt(lrResult.ciLow[i], 3), fmt(lrResult.ciHigh[i], 3), pStar(lrResult.pvals[i])]));
                          downloadCSV(rows, "logistic_multivariable.csv");
                        }} />
                      </div>
                      <div style={{ overflowX: "auto" }}>
                        <table style={{ width: "100%", borderCollapse: "collapse", fontSize: 12 }}>
                          <thead>
                            <tr style={{ borderBottom: "2px solid #334155" }}>
                              {["Variable", "β", "SE", "z", "p", "OR", "95% CI", ""].map((h) => (
                                <th key={h} style={{ padding: "6px 8px", textAlign: "left", color: "#94a3b8", fontWeight: 600, fontSize: 11 }}>{h}</th>
                              ))}
                            </tr>
                          </thead>
                          <tbody>
                            {lrResult.names.map((n, i) => (
                              <tr key={i} style={{ borderBottom: "1px solid #1e293b", background: lrResult.pvals[i] < 0.05 && i > 0 ? "rgba(15,118,110,.1)" : "transparent" }}>
                                <td style={{ padding: "5px 8px", fontWeight: 500, color: i === 0 ? "#64748b" : "#e2e8f0" }}>{n}</td>
                                <td style={{ padding: "5px 8px" }}>{fmt(lrResult.coefficients[i], 3)}</td>
                                <td style={{ padding: "5px 8px" }}>{fmt(lrResult.se[i], 3)}</td>
                                <td style={{ padding: "5px 8px" }}>{fmt(lrResult.z[i], 2)}</td>
                                <td style={{ padding: "5px 8px", fontWeight: lrResult.pvals[i] < 0.05 ? 700 : 400, color: lrResult.pvals[i] < 0.05 ? "#5eead4" : "#94a3b8" }}>{fmtP(lrResult.pvals[i])}</td>
                                <td style={{ padding: "5px 8px", fontWeight: 600 }}>{i === 0 ? "—" : fmt(lrResult.or[i], 3)}</td>
                                <td style={{ padding: "5px 8px", fontSize: 11 }}>{i === 0 ? "—" : `[${fmt(lrResult.ciLow[i], 2)}, ${fmt(lrResult.ciHigh[i], 2)}]`}</td>
                                <td style={{ padding: "5px 8px", color: "#eab308" }}>{i > 0 ? pStar(lrResult.pvals[i]) : ""}</td>
                              </tr>
                            ))}
                          </tbody>
                        </table>
                      </div>
                      {/* Forest plot of ORs */}
                      <div style={{ marginTop: 16 }}>
                        <div style={{ fontSize: 12, color: "#94a3b8", marginBottom: 8 }}>Forest Plot — Adjusted Odds Ratios (95% CI)</div>
                        <ResponsiveContainer width="100%" height={Math.max(200, (lrResult.names.length - 1) * 28 + 40)}>
                          <ScatterChart margin={{ left: 160, right: 40, top: 10, bottom: 10 }}>
                            <CartesianGrid strokeDasharray="3 3" stroke="#1e293b" />
                            <XAxis type="number" dataKey="or" domain={["auto", "auto"]} tick={{ fill: "#94a3b8", fontSize: 11 }} name="OR" scale="log" />
                            <YAxis type="category" dataKey="name" tick={{ fill: "#94a3b8", fontSize: 11 }} width={155} />
                            <Tooltip content={({ active, payload }) => {
                              if (!active || !payload?.[0]) return null;
                              const d = payload[0].payload;
                              return (
                                <div style={{ background: "#0f172a", border: "1px solid #334155", borderRadius: 8, padding: 8, fontSize: 12, color: "#e2e8f0" }}>
                                  <div style={{ fontWeight: 600 }}>{d.name}</div>
                                  <div>OR: {fmt(d.or, 3)} [{fmt(d.lo, 2)}, {fmt(d.hi, 2)}]</div>
                                  <div>p = {fmtP(d.p)}</div>
                                </div>
                              );
                            }} />
                            <Scatter data={lrResult.names.slice(1).map((n, i) => ({
                              name: n, or: lrResult.or[i + 1], lo: lrResult.ciLow[i + 1], hi: lrResult.ciHigh[i + 1], p: lrResult.pvals[i + 1]
                            }))} fill="#0f766e" shape={<CustomDot />} />
                          </ScatterChart>
                        </ResponsiveContainer>
                      </div>
                    </>
                  )}

                  {/* Combined export when both */}
                  {lrMode === "both" && lrUniResults && lrResult && !lrResult.error && (
                    <div style={{ marginTop: 16 }}>
                      <ExportBtn label="Export Combined (Uni + Multi)" onClick={() => {
                        const rows = [["Variable", "Unadj OR", "Unadj 95% CI", "Unadj p", "Adj OR", "Adj 95% CI", "Adj p"]];
                        // Match by variable name
                        lrResult.names.slice(1).forEach((name, i) => {
                          const uni = lrUniResults.find((u) => u.label === name);
                          rows.push([
                            name,
                            uni && !uni.error ? fmt(uni.or, 3) : "—",
                            uni && !uni.error ? `[${fmt(uni.ciLow, 2)}, ${fmt(uni.ciHigh, 2)}]` : "—",
                            uni && !uni.error ? fmtP(uni.p) : "—",
                            fmt(lrResult.or[i + 1], 3),
                            `[${fmt(lrResult.ciLow[i + 1], 2)}, ${fmt(lrResult.ciHigh[i + 1], 2)}]`,
                            fmtP(lrResult.pvals[i + 1]),
                          ]);
                        });
                        downloadCSV(rows, "logistic_combined_uni_multi.csv");
                      }} />
                    </div>
                  )}
                </div>
              </div>
            </Card>
          </>
        )}

        {/* ─── LINEAR REGRESSION ─── */}
        {tab === "linear" && (
          <>
            <Card title="Linear Regression — Continuous CMD Outcomes">
              <div style={{ fontSize: 12, color: "#94a3b8", marginBottom: 12 }}>
                OLS regression on continuous scale scores. <strong>Univariate</strong> runs each predictor alone (unadjusted β). <strong>Multivariable</strong> enters all simultaneously (adjusted β). <strong>Both</strong> shows side-by-side.
              </div>
              <div style={{ display: "grid", gridTemplateColumns: "1fr 2fr", gap: 16 }}>
                <div>
                  <div style={{ marginBottom: 12 }}>
                    <label style={{ fontSize: 11, color: "#5eead4", fontWeight: 600, display: "block", marginBottom: 4 }}>Outcome (Continuous)</label>
                    <select value={linOutcome} onChange={(e) => setLinOutcome(e.target.value)}
                      style={{ width: "100%", background: "#0f172a", color: "#e2e8f0", border: "1px solid #334155", borderRadius: 6, padding: "7px 8px", fontSize: 12 }}>
                      {CMD_SCALES.map((s) => <option key={s.key} value={s.key}>{s.label} Total</option>)}
                    </select>
                  </div>
                  <div style={{ marginBottom: 12 }}>
                    <label style={{ fontSize: 11, color: "#5eead4", fontWeight: 600, display: "block", marginBottom: 6 }}>Model Type</label>
                    <div style={{ display: "flex", gap: 4, background: "#0f172a", borderRadius: 8, padding: 3, border: "1px solid #334155" }}>
                      {[{ id: "univariate", label: "Univariate" }, { id: "multivariable", label: "Multivariable" }, { id: "both", label: "Both" }].map((m) => (
                        <button key={m.id} onClick={() => setLinMode(m.id)}
                          style={{
                            flex: 1, padding: "6px 4px", fontSize: 11, fontWeight: linMode === m.id ? 700 : 500, border: "none", borderRadius: 6, cursor: "pointer",
                            background: linMode === m.id ? "#7c3aed" : "transparent", color: linMode === m.id ? "#fff" : "#94a3b8",
                          }}>{m.label}</button>
                      ))}
                    </div>
                  </div>
                  <div style={{ marginBottom: 12 }}>
                    <label style={{ fontSize: 11, color: "#5eead4", fontWeight: 600, display: "block", marginBottom: 4 }}>Predictors</label>
                    {ACTIVE_RISK_GROUPS.map((g) => (
                      <div key={g.label} style={{ marginBottom: 6 }}>
                        <div style={{ fontSize: 10, color: "#64748b", fontWeight: 700, textTransform: "uppercase", marginBottom: 2, letterSpacing: 0.5 }}>{g.label}</div>
                        <CheckList items={g.vars} selected={linPredictors} onChange={setLinPredictors} columns={1} />
                      </div>
                    ))}
                  </div>
                  <button onClick={runLinear}
                    style={{ width: "100%", padding: "10px", background: "linear-gradient(135deg, #7c3aed, #6366f1)", color: "#fff", border: "none", borderRadius: 8, fontWeight: 700, fontSize: 13, cursor: "pointer" }}>
                    Run Linear Regression
                  </button>
                  {(linResult || linUniResults) && (
                    <button onClick={exportLinearWord}
                      style={{ width: "100%", padding: "8px", marginTop: 8, background: "#1e293b", color: "#a78bfa", border: "1px solid #334155", borderRadius: 8, fontWeight: 600, fontSize: 12, cursor: "pointer" }}>
                      ⬇ Export to Word (.docx)
                    </button>
                  )}
                </div>
                <div>
                  {/* ── Univariate Results ── */}
                  {linUniResults && (linMode === "univariate" || linMode === "both") && (
                    <div style={{ marginBottom: 20 }}>
                      <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 8 }}>
                        <h4 style={{ fontSize: 14, color: "#a78bfa", margin: 0, fontWeight: 700 }}>Univariate (Unadjusted) Coefficients</h4>
                        <ExportBtn label="Export Univariate" onClick={() => {
                          const rows = [["Variable", "N", "β", "SE", "t", "p-value", "R²", "Adj R²", "Sig"]];
                          linUniResults.forEach((r) => {
                            if (r.error) { rows.push([r.label, "", "", "", "", "", "", "", r.error]); }
                            else { rows.push([r.label, r.n, fmt(r.beta, 4), fmt(r.se, 4), fmt(r.t, 3), fmtP(r.p), fmt(r.r2, 4), fmt(r.adjR2, 4), pStar(r.p)]); }
                          });
                          downloadCSV(rows, "linear_univariate.csv");
                        }} />
                      </div>
                      <div style={{ overflowX: "auto" }}>
                        <table style={{ width: "100%", borderCollapse: "collapse", fontSize: 12 }}>
                          <thead>
                            <tr style={{ borderBottom: "2px solid #334155" }}>
                              {["Variable", "N", "β", "SE", "t", "p", "R²", ""].map((h) => (
                                <th key={h} style={{ padding: "6px 6px", textAlign: "left", color: "#94a3b8", fontWeight: 600, fontSize: 11 }}>{h}</th>
                              ))}
                            </tr>
                          </thead>
                          <tbody>
                            {linUniResults.map((r, i) => (
                              <tr key={i} style={{ borderBottom: "1px solid #1e293b", background: !r.error && r.p < 0.05 ? "rgba(124,58,237,.1)" : "transparent" }}>
                                <td style={{ padding: "5px 6px", fontWeight: 500, color: "#e2e8f0" }}>{r.label}</td>
                                {r.error ? (
                                  <td colSpan={6} style={{ padding: "5px 6px", color: "#ef4444", fontStyle: "italic" }}>{r.error}</td>
                                ) : (<>
                                  <td style={{ padding: "5px 6px" }}>{r.n}</td>
                                  <td style={{ padding: "5px 6px" }}>{fmt(r.beta, 4)}</td>
                                  <td style={{ padding: "5px 6px" }}>{fmt(r.se, 4)}</td>
                                  <td style={{ padding: "5px 6px" }}>{fmt(r.t, 2)}</td>
                                  <td style={{ padding: "5px 6px", fontWeight: r.p < 0.05 ? 700 : 400, color: r.p < 0.05 ? "#a78bfa" : "#94a3b8" }}>{fmtP(r.p)}</td>
                                  <td style={{ padding: "5px 6px", fontSize: 11, color: "#64748b" }}>{fmt(r.r2, 3)}</td>
                                  <td style={{ padding: "5px 6px", color: "#eab308" }}>{pStar(r.p)}</td>
                                </>)}
                              </tr>
                            ))}
                          </tbody>
                        </table>
                      </div>
                    </div>
                  )}

                  {/* ── Multivariable Results ── */}
                  {linResult?.error && (linMode === "multivariable" || linMode === "both") && (
                    <div style={{ color: "#ef4444", padding: 12, background: "#1c1917", borderRadius: 8, fontSize: 13, marginBottom: 12 }}>⚠ {linResult.error}</div>
                  )}
                  {linResult && !linResult.error && (linMode === "multivariable" || linMode === "both") && (
                    <>
                      <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 8 }}>
                        <h4 style={{ fontSize: 14, color: "#a78bfa", margin: 0, fontWeight: 700 }}>Multivariable (Adjusted) Coefficients</h4>
                      </div>
                      <div style={{ display: "flex", gap: 16, marginBottom: 12, fontSize: 12, color: "#94a3b8" }}>
                        <span>N = {linResult.n}</span>
                        <span>R² = {fmt(linResult.r2, 3)}</span>
                        <span>Adj R² = {fmt(linResult.adjR2, 3)}</span>
                        <span>RMSE = {fmt(Math.sqrt(linResult.mse), 2)}</span>
                        <ExportBtn label="Export Multivariable" onClick={() => {
                          const rows = [["Variable", "Coefficient", "SE", "t", "p-value", "Sig"]];
                          linResult.names.forEach((n, i) => rows.push([n, fmt(linResult.coefficients[i], 4), fmt(linResult.se[i], 4), fmt(linResult.tstat[i], 3), fmtP(linResult.pvals[i]), pStar(linResult.pvals[i])]));
                          downloadCSV(rows, "linear_multivariable.csv");
                        }} />
                      </div>
                      <div style={{ overflowX: "auto" }}>
                        <table style={{ width: "100%", borderCollapse: "collapse", fontSize: 12 }}>
                          <thead>
                            <tr style={{ borderBottom: "2px solid #334155" }}>
                              {["Variable", "β", "SE", "t", "p", ""].map((h) => (
                                <th key={h} style={{ padding: "6px 8px", textAlign: "left", color: "#94a3b8", fontWeight: 600, fontSize: 11 }}>{h}</th>
                              ))}
                            </tr>
                          </thead>
                          <tbody>
                            {linResult.names.map((n, i) => (
                              <tr key={i} style={{ borderBottom: "1px solid #1e293b", background: linResult.pvals[i] < 0.05 && i > 0 ? "rgba(124,58,237,.1)" : "transparent" }}>
                                <td style={{ padding: "5px 8px", fontWeight: 500, color: i === 0 ? "#64748b" : "#e2e8f0" }}>{n}</td>
                                <td style={{ padding: "5px 8px" }}>{fmt(linResult.coefficients[i], 4)}</td>
                                <td style={{ padding: "5px 8px" }}>{fmt(linResult.se[i], 4)}</td>
                                <td style={{ padding: "5px 8px" }}>{fmt(linResult.tstat[i], 2)}</td>
                                <td style={{ padding: "5px 8px", fontWeight: linResult.pvals[i] < 0.05 ? 700 : 400, color: linResult.pvals[i] < 0.05 ? "#a78bfa" : "#94a3b8" }}>{fmtP(linResult.pvals[i])}</td>
                                <td style={{ padding: "5px 8px", color: "#eab308" }}>{i > 0 ? pStar(linResult.pvals[i]) : ""}</td>
                              </tr>
                            ))}
                          </tbody>
                        </table>
                      </div>
                      {/* Coefficient bar chart */}
                      <div style={{ marginTop: 16 }}>
                        <div style={{ fontSize: 12, color: "#94a3b8", marginBottom: 8 }}>Adjusted Coefficient Magnitude</div>
                        <ResponsiveContainer width="100%" height={Math.max(180, (linResult.names.length - 1) * 26 + 40)}>
                          <BarChart data={linResult.names.slice(1).map((n, i) => ({ name: n, coef: linResult.coefficients[i + 1], sig: linResult.pvals[i + 1] < 0.05 }))} layout="vertical" margin={{ left: 160 }}>
                            <CartesianGrid strokeDasharray="3 3" stroke="#1e293b" />
                            <XAxis type="number" tick={{ fill: "#94a3b8", fontSize: 11 }} />
                            <YAxis type="category" dataKey="name" tick={{ fill: "#94a3b8", fontSize: 11 }} width={155} />
                            <Tooltip content={<TT />} />
                            <Bar dataKey="coef" name="Coefficient">
                              {linResult.names.slice(1).map((_, i) => (
                                <Cell key={i} fill={linResult.pvals[i + 1] < 0.05 ? "#7c3aed" : "#475569"} />
                              ))}
                            </Bar>
                          </BarChart>
                        </ResponsiveContainer>
                      </div>
                    </>
                  )}

                  {/* Combined export when both */}
                  {linMode === "both" && linUniResults && linResult && !linResult.error && (
                    <div style={{ marginTop: 16 }}>
                      <ExportBtn label="Export Combined (Uni + Multi)" onClick={() => {
                        const rows = [["Variable", "Unadj β", "Unadj SE", "Unadj p", "Unadj R²", "Adj β", "Adj SE", "Adj p"]];
                        linResult.names.slice(1).forEach((name, i) => {
                          const uni = linUniResults.find((u) => u.label === name);
                          rows.push([
                            name,
                            uni && !uni.error ? fmt(uni.beta, 4) : "—",
                            uni && !uni.error ? fmt(uni.se, 4) : "—",
                            uni && !uni.error ? fmtP(uni.p) : "—",
                            uni && !uni.error ? fmt(uni.r2, 4) : "—",
                            fmt(linResult.coefficients[i + 1], 4),
                            fmt(linResult.se[i + 1], 4),
                            fmtP(linResult.pvals[i + 1]),
                          ]);
                        });
                        downloadCSV(rows, "linear_combined_uni_multi.csv");
                      }} />
                    </div>
                  )}
                </div>
              </div>
            </Card>
          </>
        )}

        {/* ─── MISSING DATA ─── */}
        {tab === "missing" && (
          <Card title={`Missing Data Summary (${missingInfo.length} variables with missing values)`}>
            <ExportBtn label="Export Missing Data Report" onClick={() => {
              const rows = [["Variable", "Missing Count", "Missing %", "Total N"]];
              missingInfo.forEach((m) => rows.push([m.col, m.missing, m.pct, data.length]));
              downloadCSV(rows, "missing_data.csv");
            }} />
            {missingInfo.length === 0 ? (
              <div style={{ color: "#22c55e", padding: 20, textAlign: "center" }}>No missing data detected</div>
            ) : (
              <div style={{ maxHeight: 600, overflowY: "auto", marginTop: 10 }}>
                <table style={{ width: "100%", borderCollapse: "collapse", fontSize: 12 }}>
                  <thead style={{ position: "sticky", top: 0, background: "#1e293b" }}>
                    <tr style={{ borderBottom: "2px solid #334155" }}>
                      {["Variable", "Missing", "%", "Present"].map((h) => (
                        <th key={h} style={{ padding: "6px 8px", textAlign: "left", color: "#94a3b8", fontWeight: 600, fontSize: 11 }}>{h}</th>
                      ))}
                    </tr>
                  </thead>
                  <tbody>
                    {missingInfo.map((m) => (
                      <tr key={m.col} style={{ borderBottom: "1px solid #0f172a" }}>
                        <td style={{ padding: "4px 8px", color: "#e2e8f0" }}>{m.col}</td>
                        <td style={{ padding: "4px 8px", color: "#ef4444", fontWeight: 600 }}>{m.missing}</td>
                        <td style={{ padding: "4px 8px" }}>
                          <div style={{ display: "flex", alignItems: "center", gap: 6 }}>
                            <div style={{ width: 60, height: 6, background: "#0f172a", borderRadius: 3, overflow: "hidden" }}>
                              <div style={{ width: `${m.pct}%`, height: "100%", background: parseFloat(m.pct) > 30 ? "#ef4444" : parseFloat(m.pct) > 10 ? "#eab308" : "#22c55e", borderRadius: 3 }} />
                            </div>
                            <span style={{ fontSize: 11 }}>{m.pct}%</span>
                          </div>
                        </td>
                        <td style={{ padding: "4px 8px", color: "#64748b" }}>{data.length - m.missing}</td>
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>
            )}
          </Card>
        )}

        {/* Footer */}
        <div style={{ textAlign: "center", padding: "24px 0 12px", fontSize: 11, color: "#475569" }}>
           Jockey Mental Wellbeing Research Dashboard
        </div>
      </div>
    </div>
  );
}

// Custom dot for forest plot
const CustomDot = (props) => {
  const { cx, cy, payload } = props;
  if (!cx || !cy) return null;
  return (
    <g>
      <circle cx={cx} cy={cy} r={5} fill={payload.p < 0.05 ? "#0f766e" : "#475569"} stroke="#e2e8f0" strokeWidth={1} />
    </g>
  );
};
