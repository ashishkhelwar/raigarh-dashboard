import React, { useState, useMemo } from "react";
import {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
  AlignmentType, WidthType, BorderStyle, ShadingType, HeadingLevel
} from "docx";
import {
  TreePine, MapPin, Users, FileText, Download, ChevronRight,
  BarChart3, Layers, Trees, Shield, AlertTriangle, ArrowUpRight,
  Map as MapIcon, ScrollText, FileBarChart, Building2, Leaf,
  Mountain, Compass, Sparkles, Calendar, Wheat
} from "lucide-react";
import {
  BarChart, Bar, LineChart, Line, PieChart, Pie, Cell,
  XAxis, YAxis, CartesianGrid, Tooltip, ResponsiveContainer, Legend, Area, AreaChart
} from "recharts";

// ============================================================================
// DIVISION DATA — extracted from Working Plan & supporting records
// (Sarangarh & Gomarda excluded — bifurcated to separate division)
// ============================================================================

const DIVISION = {
  name: "Raigarh Forest Division",
  nameHi: "रायगढ़ वनमण्डल",
  circle: "Bilaspur",
  state: "Chhattisgarh",
  hq: "Raigarh",
  planPeriod: "2020-21 → 2029-30",
  currentApoYear: "2026-27",
  planYearIndex: 7, // year 7 of 10
};

const RANGES = ["Raigarh", "Kharsiya", "Gharghoda", "Tamnar"];
const RANGES_HI = { Raigarh:"रायगढ़", Kharsiya:"खरसिया", Gharghoda:"घरघोड़ा", Tamnar:"तमनार" };

const RANGE_DATA = {
  Raigarh:   { compartments: 28, jfmcs: 93, area_ha: 16969, rf_count: 72, rf_area: 22577, pf_count: 48, pf_area: 4338,  oa_count: 47, oa_area: 962 },
  Kharsiya:  { compartments: 33, jfmcs: 47, area_ha:  9397, rf_count: 43, rf_area: 11036, pf_count: 27, pf_area: 1709,  oa_count:  2, oa_area:  54 },
  Gharghoda: { compartments: 32, jfmcs: 59, area_ha: 12249, rf_count: 52, rf_area: 12900, pf_count: 83, pf_area: 5620,  oa_count: 37, oa_area: 893 },
  Tamnar:    { compartments: 57, jfmcs: 66, area_ha: 14757, rf_count: 62, rf_area:  7787, pf_count: 79, pf_area: 4104,  oa_count: 17, oa_area: 310 },
};

const WC_COUNTS = { PWC: 114, SCI: 35, RWC: 37, IWC: 36 };
const WC_NAMES = {
  PWC: "Protection",
  SCI: "Selection-cum-Improvement",
  RWC: "Rehabilitation",
  IWC: "Improvement"
};

// Site quality across the entire division (Ha)
const SITE_QUALITY = [
  { name: "Sal SQ-IVa",     area: 15762.24, group: "Sal" },
  { name: "Mixed SQ-IVa",   area:  8364.67, group: "Mixed" },
  { name: "Mixed SQ-IVb",   area:  7662.24, group: "Mixed" },
  { name: "Sal SQ-IVb",     area:  5443.07, group: "Sal" },
  { name: "Understocked",   area:  4983.40, group: "Other" },
  { name: "Blank",          area:  4056.59, group: "Other" },
  { name: "Sal SQ-III",     area:  3501.17, group: "Sal" },
  { name: "Mixed SQ-Vb",    area:  1765.66, group: "Mixed" },
  { name: "Plantation",     area:   642.70, group: "Plantation" },
  { name: "Mixed SQ-Va",    area:   557.31, group: "Mixed" },
  { name: "Mixed SQ-III",   area:   343.02, group: "Mixed" },
  { name: "Encroached",     area:   248.71, group: "Other" },
  { name: "Sal SQ-II",      area:    27.25, group: "Sal" },
  { name: "FRA Plot",       area:    14.38, group: "Other" },
];

const WATERSHED_SCHEDULE = [
  { year: "2020-21", area: 13891.00, status: "completed" },
  { year: "2021-22", area: 13929.00, status: "completed" },
  { year: "2022-23", area: 14947.87, status: "completed" },
  { year: "2023-24", area: 14217.00, status: "completed" },
  { year: "2024-25", area: 13424.74, status: "completed" },
  { year: "2025-26", area: 11521.56, status: "completed" },
  { year: "2026-27", area: 12312.10, status: "current" },
  { year: "2027-28", area: 13490.64, status: "future" },
  { year: "2028-29", area: 13373.32, status: "future" },
  { year: "2029-30", area: 14660.00, status: "future" },
];

const FRA = { cfr_count: 298, cfr_area_ha: 43102.94, ifr_count: 1174 };

// Aggregate totals
const TOT = {
  compartments: Object.values(RANGE_DATA).reduce((s,r)=>s+r.compartments, 0),
  jfmcs: Object.values(RANGE_DATA).reduce((s,r)=>s+r.jfmcs, 0),
  area: Object.values(RANGE_DATA).reduce((s,r)=>s+r.area_ha, 0),
  rf_count: Object.values(RANGE_DATA).reduce((s,r)=>s+r.rf_count, 0),
  rf_area: Object.values(RANGE_DATA).reduce((s,r)=>s+r.rf_area, 0),
  pf_count: Object.values(RANGE_DATA).reduce((s,r)=>s+r.pf_count, 0),
  pf_area: Object.values(RANGE_DATA).reduce((s,r)=>s+r.pf_area, 0),
  oa_count: Object.values(RANGE_DATA).reduce((s,r)=>s+r.oa_count, 0),
  oa_area: Object.values(RANGE_DATA).reduce((s,r)=>s+r.oa_area, 0),
};

// Available reports
const REPORTS = [
  {
    id: "division_brief",
    title: "Division Profile Brief",
    titleHi: "वनमण्डल परिचय प्रतिवेदन",
    desc: "One-page executive summary with all key division statistics, suitable for VIP visits and inter-departmental briefings.",
    icon: ScrollText,
    pages: "2 pages",
    color: "#0d4e2c",
  },
  {
    id: "range_wise",
    title: "Range-wise Statistical Bulletin",
    titleHi: "परिक्षेत्रवार सांख्यिकीय विवरण",
    desc: "Detailed range-by-range breakdown of compartments, JFMCs, forest area, RF/PF/OA blocks, and management circles.",
    icon: BarChart3,
    pages: "5 pages",
    color: "#14532d",
  },
  {
    id: "site_quality",
    title: "Forest Composition & Site Quality",
    titleHi: "वन संरचना एवं भूमि गुणवत्ता प्रतिवेदन",
    desc: "Sal & Mixed forest distribution by site quality class. Understocked area, blanks, plantations.",
    icon: Trees,
    pages: "4 pages",
    color: "#365314",
  },
  {
    id: "watershed_plan",
    title: "Watershed Treatment Schedule",
    titleHi: "वाटरशेड उपचार अनुसूची",
    desc: "10-year watershed treatment area schedule per Working Plan, with progress against current APO year.",
    icon: Mountain,
    pages: "3 pages",
    color: "#0f766e",
  },
  {
    id: "fra_status",
    title: "Forest Rights Act Status",
    titleHi: "वन अधिकार अधिनियम स्थिति प्रतिवेदन",
    desc: "Community (CFR) and Individual (IFR) forest rights granted within division boundary.",
    icon: Shield,
    pages: "3 pages",
    color: "#7c2d12",
  },
  {
    id: "jfmc_directory",
    title: "JFMC Directory by Range",
    titleHi: "वन प्रबंधन समिति निर्देशिका",
    desc: "265 active Joint Forest Management Committees with range-wise distribution.",
    icon: Users,
    pages: "8 pages",
    color: "#713f12",
  },
];

// ============================================================================
// HELPERS
// ============================================================================
const fmt = (n) => new Intl.NumberFormat("en-IN", { maximumFractionDigits: 0 }).format(n);
const fmt2 = (n) => new Intl.NumberFormat("en-IN", { maximumFractionDigits: 2 }).format(n);
const fmtArea = (n) => fmt(n) + " Ha";

// ============================================================================
// MAIN COMPONENT
// ============================================================================
export default function DivisionDashboard() {
  const [activeReport, setActiveReport] = useState(null);

  return (
    <div style={{
      minHeight: "100vh",
      background: "linear-gradient(180deg, #faf8f3 0%, #f5f1e8 100%)",
      fontFamily: "'Source Sans 3', 'Inter', system-ui, sans-serif",
      color: "#1c1917"
    }}>
      <style>{`
        @import url('https://fonts.googleapis.com/css2?family=Fraunces:opsz,wght@9..144,400;9..144,500;9..144,600;9..144,700&family=Source+Sans+3:wght@400;500;600;700&family=JetBrains+Mono:wght@400;500&display=swap');
        .display-font { font-family: 'Fraunces', Georgia, serif; }
        .mono-font { font-family: 'JetBrains Mono', monospace; }
        .deco-rule {
          background: linear-gradient(90deg, transparent, #0d4e2c 20%, #0d4e2c 80%, transparent);
          height: 1px;
        }
        .ornate-bullet::before {
          content: "❦";
          color: #0d4e2c;
          margin-right: 6px;
          font-size: 0.9em;
        }
        @keyframes fadeUp {
          from { opacity: 0; transform: translateY(8px); }
          to { opacity: 1; transform: translateY(0); }
        }
        .fade-up { animation: fadeUp 0.5s ease-out both; }
      `}</style>

      <Header />

      <main style={{ maxWidth: 1280, margin: "0 auto", padding: "32px 32px 80px" }}>
        <DivisionMasthead />
        <KpiGrid />
        <RangeBreakdown />
        <CompositionCharts />
        <WatershedTimeline />
        <ReportsSection onPick={setActiveReport} />
      </main>

      {activeReport && <ReportModal report={activeReport} onClose={()=>setActiveReport(null)} />}

      <footer style={{
        textAlign: "center",
        padding: "24px",
        fontSize: 11,
        color: "#78716c",
        borderTop: "1px solid rgba(13,78,44,0.12)",
        background: "#0d4e2c08",
        letterSpacing: "0.05em",
      }}>
        Working Plan {DIVISION.planPeriod} · APO Year {DIVISION.currentApoYear} · Year {DIVISION.planYearIndex} of 10
      </footer>
    </div>
  );
}

// ============================================================================
// HEADER
// ============================================================================
function Header() {
  return (
    <header style={{
      background: "linear-gradient(135deg, #0d4e2c 0%, #143d2b 50%, #0a3a25 100%)",
      borderBottom: "3px solid #d97706",
      color: "#fef3c7",
      position: "relative",
      overflow: "hidden",
    }}>
      <div style={{
        position: "absolute",
        inset: 0,
        backgroundImage: `radial-gradient(circle at 1px 1px, rgba(217,119,6,0.08) 1px, transparent 0)`,
        backgroundSize: "20px 20px",
      }}/>
      <div style={{ maxWidth: 1280, margin: "0 auto", padding: "20px 32px", position: "relative", zIndex: 1 }}>
        <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between" }}>
          <div style={{ display: "flex", alignItems: "center", gap: 16 }}>
            <div style={{
              width: 52, height: 52,
              background: "#fef3c7",
              borderRadius: "50%",
              display: "flex", alignItems: "center", justifyContent: "center",
              border: "2px solid #d97706",
              boxShadow: "inset 0 2px 4px rgba(0,0,0,0.1)",
            }}>
              <TreePine size={26} color="#0d4e2c" strokeWidth={1.5}/>
            </div>
            <div>
              <div className="display-font" style={{ fontSize: 26, fontWeight: 600, lineHeight: 1.1, letterSpacing: "-0.02em" }}>
                Raigarh Forest Division
              </div>
              <div style={{ fontSize: 11, letterSpacing: "0.18em", textTransform: "uppercase", color: "#fcd34d", marginTop: 2 }}>
                Government of Chhattisgarh · Bilaspur Circle
              </div>
            </div>
          </div>
          <div style={{ display: "flex", gap: 12, alignItems: "center" }}>
            <div style={{
              padding: "6px 14px",
              background: "#fef3c708",
              border: "1px solid rgba(252,211,77,0.3)",
              borderRadius: 4,
              fontSize: 11,
              letterSpacing: "0.08em",
              color: "#fef3c7",
            }}>
              Division Dashboard · Live Data
            </div>
          </div>
        </div>
      </div>
    </header>
  );
}

// ============================================================================
// MASTHEAD — division at-a-glance
// ============================================================================
function DivisionMasthead() {
  return (
    <section style={{ padding: "40px 0 24px", textAlign: "center" }}>
      <div style={{ fontSize: 11, letterSpacing: "0.32em", textTransform: "uppercase", color: "#78716c", marginBottom: 14 }}>
        Working Plan {DIVISION.planPeriod}
      </div>
      <h1 className="display-font" style={{
        fontSize: 56,
        fontWeight: 600,
        lineHeight: 1,
        margin: "0 0 8px",
        color: "#0d4e2c",
        letterSpacing: "-0.025em",
      }}>
        {DIVISION.nameHi}
      </h1>
      <div style={{ fontSize: 14, color: "#78716c", fontStyle: "italic" }}>
        Reports, statistics, and APO planning data — annual edition {DIVISION.currentApoYear}
      </div>
      <div className="deco-rule" style={{ width: 240, margin: "20px auto 0", opacity: 0.4 }}/>
    </section>
  );
}

// ============================================================================
// KPI GRID — primary statistics
// ============================================================================
function KpiGrid() {
  const kpis = [
    { label: "Compartments", value: TOT.compartments, sub: "in Working Plan", icon: Layers, color: "#0d4e2c" },
    { label: "Forest Area", value: fmt(TOT.area), unit: "Ha", sub: "across 4 ranges", icon: Mountain, color: "#365314" },
    { label: "JFM Committees", value: TOT.jfmcs, sub: "active", icon: Users, color: "#7c2d12" },
    { label: "Plan Year", value: DIVISION.planYearIndex, unit: "/10", sub: "of current cycle", icon: Calendar, color: "#a16207" },
  ];
  return (
    <section style={{
      display: "grid",
      gridTemplateColumns: "repeat(4, 1fr)",
      gap: 16,
      marginBottom: 36,
    }}>
      {kpis.map((k, i) => {
        const Icon = k.icon;
        return (
          <div key={i} className="fade-up" style={{
            background: "white",
            border: "1px solid rgba(13,78,44,0.15)",
            padding: "20px 22px",
            borderRadius: 4,
            position: "relative",
            animationDelay: `${i*60}ms`,
            overflow: "hidden",
          }}>
            <div style={{
              position: "absolute",
              top: 0, left: 0, width: 4, height: "100%",
              background: k.color,
            }}/>
            <div style={{ display: "flex", justifyContent: "space-between", alignItems: "flex-start", marginBottom: 12 }}>
              <div style={{ fontSize: 10, letterSpacing: "0.16em", textTransform: "uppercase", color: "#78716c", fontWeight: 500 }}>
                {k.label}
              </div>
              <Icon size={18} color={k.color} strokeWidth={1.5} style={{ opacity: 0.4 }}/>
            </div>
            <div style={{ display: "flex", alignItems: "baseline", gap: 4 }}>
              <span className="display-font" style={{ fontSize: 38, fontWeight: 600, color: "#1c1917", letterSpacing: "-0.02em", lineHeight: 1 }}>
                {k.value}
              </span>
              {k.unit && <span className="display-font" style={{ fontSize: 18, color: "#78716c", fontWeight: 400 }}>{k.unit}</span>}
            </div>
            <div style={{ fontSize: 11, color: "#78716c", marginTop: 6, fontStyle: "italic" }}>
              {k.sub}
            </div>
          </div>
        );
      })}
    </section>
  );
}

// ============================================================================
// RANGE BREAKDOWN
// ============================================================================
function RangeBreakdown() {
  const rangeChartData = RANGES.map(r => ({
    name: r,
    nameHi: RANGES_HI[r],
    Compartments: RANGE_DATA[r].compartments,
    "JFM Committees": RANGE_DATA[r].jfmcs,
    "Area (k Ha)": Math.round(RANGE_DATA[r].area_ha / 100) / 10,
  }));

  return (
    <section style={{ marginBottom: 36 }}>
      <SectionTitle eyebrow="By Administrative Range" title="Range-wise Distribution" />

      <div style={{ display: "grid", gridTemplateColumns: "1.5fr 1fr", gap: 16 }}>
        <div style={{ background: "white", border: "1px solid rgba(13,78,44,0.15)", borderRadius: 4, padding: "20px 24px" }}>
          <table style={{ width: "100%", fontSize: 13, borderCollapse: "collapse" }}>
            <thead>
              <tr style={{ borderBottom: "2px solid #0d4e2c" }}>
                <th style={th()}>Range</th>
                <th style={th("right")}>Compt.</th>
                <th style={th("right")}>JFMCs</th>
                <th style={th("right")}>Area (Ha)</th>
                <th style={th("right")}>RF / PF / OA</th>
              </tr>
            </thead>
            <tbody>
              {RANGES.map(r => {
                const d = RANGE_DATA[r];
                return (
                  <tr key={r} style={{ borderBottom: "1px solid rgba(13,78,44,0.08)" }}>
                    <td style={td()}>
                      <div style={{ display: "flex", flexDirection: "column" }}>
                        <span className="display-font" style={{ fontSize: 16, fontWeight: 500, color: "#0d4e2c" }}>{r}</span>
                        <span style={{ fontSize: 11, color: "#78716c" }}>{RANGES_HI[r]}</span>
                      </div>
                    </td>
                    <td style={td("right")}>{d.compartments}</td>
                    <td style={td("right")}>{d.jfmcs}</td>
                    <td style={td("right")} className="mono-font">{fmt(d.area_ha)}</td>
                    <td style={td("right")} className="mono-font" >
                      <span style={{ color: "#0d4e2c", fontWeight: 500 }}>{d.rf_count}</span>
                      <span style={{ color: "#78716c" }}> / </span>
                      <span style={{ color: "#a16207" }}>{d.pf_count}</span>
                      <span style={{ color: "#78716c" }}> / </span>
                      <span style={{ color: "#7c2d12" }}>{d.oa_count}</span>
                    </td>
                  </tr>
                );
              })}
              <tr style={{ background: "#fef3c720", borderTop: "2px solid #0d4e2c" }}>
                <td style={{ ...td(), fontWeight: 600, color: "#0d4e2c" }} className="display-font">Total</td>
                <td style={{ ...td("right"), fontWeight: 600 }}>{TOT.compartments}</td>
                <td style={{ ...td("right"), fontWeight: 600 }}>{TOT.jfmcs}</td>
                <td style={{ ...td("right"), fontWeight: 600 }} className="mono-font">{fmt(TOT.area)}</td>
                <td style={{ ...td("right"), fontWeight: 600 }} className="mono-font">
                  {TOT.rf_count} / {TOT.pf_count} / {TOT.oa_count}
                </td>
              </tr>
            </tbody>
          </table>
          <div style={{ fontSize: 11, color: "#78716c", marginTop: 12, fontStyle: "italic" }}>
            RF — Reserved Forest · PF — Protected Forest · OA — Orange Area (excluded land)
          </div>
        </div>

        <div style={{ background: "white", border: "1px solid rgba(13,78,44,0.15)", borderRadius: 4, padding: "20px 20px 8px" }}>
          <div style={{ fontSize: 11, letterSpacing: "0.14em", textTransform: "uppercase", color: "#78716c", marginBottom: 14 }}>
            Working Circles
          </div>
          <div style={{ display: "flex", flexDirection: "column", gap: 12 }}>
            {Object.entries(WC_COUNTS).map(([code, count]) => {
              const total = Object.values(WC_COUNTS).reduce((s,n)=>s+n,0);
              const pct = (count/total*100);
              const colors = { PWC:"#0d4e2c", SCI:"#0f766e", RWC:"#a16207", IWC:"#92400e" };
              return (
                <div key={code}>
                  <div style={{ display: "flex", justifyContent: "space-between", marginBottom: 4 }}>
                    <div>
                      <span className="display-font" style={{ fontWeight: 600, fontSize: 14, color: colors[code] }}>{code}</span>
                      <span style={{ fontSize: 11, color: "#78716c", marginLeft: 8 }}>{WC_NAMES[code]}</span>
                    </div>
                    <span className="mono-font" style={{ fontSize: 13, fontWeight: 500 }}>{count}</span>
                  </div>
                  <div style={{ background: "rgba(13,78,44,0.08)", height: 6, borderRadius: 1 }}>
                    <div style={{ background: colors[code], width: `${pct}%`, height: "100%", borderRadius: 1 }}/>
                  </div>
                </div>
              );
            })}
          </div>
        </div>
      </div>

      <div style={{
        background: "white",
        border: "1px solid rgba(13,78,44,0.15)",
        borderRadius: 4,
        padding: "20px 24px 8px",
        marginTop: 16,
        height: 280,
      }}>
        <div style={{ fontSize: 11, letterSpacing: "0.14em", textTransform: "uppercase", color: "#78716c", marginBottom: 12 }}>
          Range Comparison
        </div>
        <ResponsiveContainer width="100%" height="85%">
          <BarChart data={rangeChartData} margin={{ top: 8, right: 16, bottom: 4, left: 0 }}>
            <CartesianGrid strokeDasharray="2 4" stroke="rgba(13,78,44,0.1)" vertical={false}/>
            <XAxis dataKey="name" tick={{ fontSize: 12, fill: "#1c1917" }} axisLine={{ stroke: "rgba(13,78,44,0.2)" }}/>
            <YAxis tick={{ fontSize: 11, fill: "#78716c" }} axisLine={false} tickLine={false}/>
            <Tooltip contentStyle={{ background: "white", border: "1px solid #0d4e2c", borderRadius: 4, fontSize: 12 }}/>
            <Legend wrapperStyle={{ fontSize: 11, paddingTop: 8 }}/>
            <Bar dataKey="Compartments" fill="#0d4e2c" radius={[2,2,0,0]}/>
            <Bar dataKey="JFM Committees" fill="#a16207" radius={[2,2,0,0]}/>
            <Bar dataKey="Area (k Ha)" fill="#0f766e" radius={[2,2,0,0]}/>
          </BarChart>
        </ResponsiveContainer>
      </div>
    </section>
  );
}

// ============================================================================
// COMPOSITION CHARTS — site quality + RF/PF/OA
// ============================================================================
function CompositionCharts() {
  const grouped = ["Sal","Mixed","Other","Plantation"].map(g => ({
    name: g,
    area: SITE_QUALITY.filter(s=>s.group===g).reduce((sum,s)=>sum+s.area,0)
  }));
  const groupColors = { Sal: "#0d4e2c", Mixed: "#0f766e", Other: "#a16207", Plantation: "#84cc16" };

  const blockData = [
    { name: "Reserved Forest (RF)", count: TOT.rf_count, area: TOT.rf_area, color: "#0d4e2c" },
    { name: "Protected Forest (PF)", count: TOT.pf_count, area: TOT.pf_area, color: "#a16207" },
    { name: "Orange Area (OA)", count: TOT.oa_count, area: TOT.oa_area, color: "#7c2d12" },
  ];
  const blockTotalArea = blockData.reduce((s,b)=>s+b.area,0);

  return (
    <section style={{ marginBottom: 36 }}>
      <SectionTitle eyebrow="Forest Composition" title="Site Quality & Land Classification" />

      <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 16 }}>
        {/* Site quality stacked horizontal */}
        <div style={{ background: "white", border: "1px solid rgba(13,78,44,0.15)", borderRadius: 4, padding: "20px 24px" }}>
          <div style={{ fontSize: 11, letterSpacing: "0.14em", textTransform: "uppercase", color: "#78716c", marginBottom: 14 }}>
            Site Quality Distribution (Ha)
          </div>
          <div style={{ display: "flex", flexDirection: "column", gap: 8 }}>
            {SITE_QUALITY.slice(0,10).map((s, i) => {
              const max = SITE_QUALITY[0].area;
              const w = (s.area/max*100);
              return (
                <div key={i} style={{ display: "grid", gridTemplateColumns: "120px 1fr 70px", alignItems: "center", gap: 8 }}>
                  <div style={{ fontSize: 12, color: "#1c1917" }}>{s.name}</div>
                  <div style={{ background: "rgba(13,78,44,0.06)", height: 18, borderRadius: 1, overflow: "hidden" }}>
                    <div style={{
                      background: groupColors[s.group],
                      width: `${w}%`,
                      height: "100%",
                      transition: "width 0.6s ease-out",
                    }}/>
                  </div>
                  <div className="mono-font" style={{ fontSize: 12, textAlign: "right", color: "#78716c" }}>
                    {fmt(s.area)}
                  </div>
                </div>
              );
            })}
          </div>
          <div style={{ display: "flex", gap: 16, marginTop: 16, paddingTop: 12, borderTop: "1px dashed rgba(13,78,44,0.15)", fontSize: 11 }}>
            {grouped.map(g => (
              <div key={g.name} style={{ display: "flex", alignItems: "center", gap: 6 }}>
                <div style={{ width: 12, height: 12, background: groupColors[g.name] }}/>
                <span style={{ color: "#78716c" }}>{g.name}</span>
                <span className="mono-font" style={{ color: "#1c1917", fontWeight: 500 }}>{fmt(g.area)}</span>
              </div>
            ))}
          </div>
        </div>

        {/* Block classification donut + table */}
        <div style={{ background: "white", border: "1px solid rgba(13,78,44,0.15)", borderRadius: 4, padding: "20px 24px" }}>
          <div style={{ fontSize: 11, letterSpacing: "0.14em", textTransform: "uppercase", color: "#78716c", marginBottom: 14 }}>
            Land Classification (Notified)
          </div>
          <div style={{ display: "flex", gap: 20, alignItems: "center" }}>
            <div style={{ width: 160, height: 160 }}>
              <ResponsiveContainer>
                <PieChart>
                  <Pie data={blockData} dataKey="area" innerRadius={50} outerRadius={75} paddingAngle={2} stroke="none">
                    {blockData.map((b,i) => <Cell key={i} fill={b.color}/>)}
                  </Pie>
                </PieChart>
              </ResponsiveContainer>
            </div>
            <div style={{ flex: 1 }}>
              {blockData.map((b, i) => (
                <div key={i} style={{ paddingBottom: 10, marginBottom: 10, borderBottom: i<2?"1px dashed rgba(13,78,44,0.1)":"none" }}>
                  <div style={{ display: "flex", alignItems: "center", gap: 8, marginBottom: 4 }}>
                    <div style={{ width: 8, height: 8, background: b.color, borderRadius: "50%" }}/>
                    <span style={{ fontSize: 12, fontWeight: 500, color: "#1c1917" }}>{b.name}</span>
                  </div>
                  <div className="mono-font" style={{ fontSize: 13, paddingLeft: 16 }}>
                    <span style={{ color: "#1c1917", fontWeight: 500 }}>{fmt(b.area)} Ha</span>
                    <span style={{ color: "#78716c", marginLeft: 6 }}>· {b.count} blocks</span>
                  </div>
                </div>
              ))}
              <div style={{ paddingLeft: 16, marginTop: 6 }}>
                <div style={{ fontSize: 10, color: "#78716c", letterSpacing: "0.1em" }}>TOTAL NOTIFIED</div>
                <div className="display-font mono-font" style={{ fontSize: 18, fontWeight: 600, color: "#0d4e2c" }}>
                  {fmt(blockTotalArea)} Ha
                </div>
              </div>
            </div>
          </div>
        </div>
      </div>

      {/* FRA strip */}
      <div style={{
        background: "linear-gradient(90deg, #fef3c715 0%, white 50%, #fef3c715 100%)",
        border: "1px solid rgba(124,45,18,0.2)",
        borderRadius: 4,
        padding: "16px 24px",
        marginTop: 16,
        display: "grid",
        gridTemplateColumns: "auto 1fr 1fr 1fr",
        gap: 24,
        alignItems: "center",
      }}>
        <div style={{ display: "flex", alignItems: "center", gap: 12 }}>
          <Shield size={20} color="#7c2d12"/>
          <div>
            <div className="display-font" style={{ fontSize: 16, fontWeight: 600, color: "#7c2d12" }}>Forest Rights Act</div>
            <div style={{ fontSize: 11, color: "#78716c", fontStyle: "italic" }}>Within division boundary</div>
          </div>
        </div>
        <FraStat label="Community FRA (CFR)" value={FRA.cfr_count} sub="patta granted"/>
        <FraStat label="Community Area" value={fmt(FRA.cfr_area_ha) + " Ha"} sub={`avg ${fmt(FRA.cfr_area_ha/FRA.cfr_count)} Ha each`}/>
        <FraStat label="Individual FRA (IFR)" value={fmt(FRA.ifr_count)} sub="patta holders"/>
      </div>
    </section>
  );
}
function FraStat({ label, value, sub }) {
  return (
    <div>
      <div style={{ fontSize: 10, letterSpacing: "0.12em", textTransform: "uppercase", color: "#78716c" }}>{label}</div>
      <div className="display-font mono-font" style={{ fontSize: 22, fontWeight: 600, color: "#7c2d12", letterSpacing: "-0.01em" }}>
        {value}
      </div>
      <div style={{ fontSize: 11, color: "#78716c", fontStyle: "italic" }}>{sub}</div>
    </div>
  );
}

// ============================================================================
// WATERSHED TIMELINE
// ============================================================================
function WatershedTimeline() {
  const totalPlanned = WATERSHED_SCHEDULE.reduce((s,w)=>s+w.area, 0);
  const completed = WATERSHED_SCHEDULE.filter(w=>w.status==="completed").reduce((s,w)=>s+w.area, 0);
  const pct = (completed/totalPlanned*100).toFixed(1);

  return (
    <section style={{ marginBottom: 36 }}>
      <SectionTitle eyebrow="Soil & Water Conservation" title="Watershed Treatment Schedule (10-Year)" />

      <div style={{ background: "white", border: "1px solid rgba(13,78,44,0.15)", borderRadius: 4, padding: "20px 24px" }}>
        <div style={{ display: "grid", gridTemplateColumns: "1fr auto auto auto", gap: 24, marginBottom: 18 }}>
          <div style={{ height: 200 }}>
            <ResponsiveContainer>
              <AreaChart data={WATERSHED_SCHEDULE} margin={{ top: 4, right: 16, bottom: 4, left: 0 }}>
                <defs>
                  <linearGradient id="ws-grad" x1="0" y1="0" x2="0" y2="1">
                    <stop offset="0%" stopColor="#0f766e" stopOpacity={0.7}/>
                    <stop offset="100%" stopColor="#0f766e" stopOpacity={0.05}/>
                  </linearGradient>
                </defs>
                <CartesianGrid strokeDasharray="2 4" stroke="rgba(13,78,44,0.08)" vertical={false}/>
                <XAxis dataKey="year" tick={{ fontSize: 10, fill: "#78716c" }} axisLine={{ stroke: "rgba(13,78,44,0.2)" }}/>
                <YAxis tick={{ fontSize: 10, fill: "#78716c" }} axisLine={false} tickLine={false}
                  tickFormatter={v => Math.round(v/1000) + "k"}/>
                <Tooltip
                  contentStyle={{ background: "white", border: "1px solid #0f766e", borderRadius: 4, fontSize: 12 }}
                  formatter={(v) => [fmt(v) + " Ha", "Treatment Area"]}
                />
                <Area type="monotone" dataKey="area" stroke="#0f766e" strokeWidth={2} fill="url(#ws-grad)"/>
              </AreaChart>
            </ResponsiveContainer>
          </div>
          <Stat label="Planned" value={fmt(totalPlanned) + " Ha"} accent="#0f766e"/>
          <Stat label="Completed" value={fmt(completed) + " Ha"} accent="#0d4e2c"/>
          <Stat label="Progress" value={pct + "%"} accent="#a16207"/>
        </div>

        <div style={{ display: "grid", gridTemplateColumns: "repeat(10, 1fr)", gap: 4 }}>
          {WATERSHED_SCHEDULE.map((w, i) => {
            const colors = {
              completed: { bg: "#0d4e2c", text: "#fef3c7" },
              current: { bg: "#a16207", text: "white" },
              future: { bg: "#fef3c7", text: "#78716c" },
            };
            const c = colors[w.status];
            return (
              <div key={i} style={{
                background: c.bg,
                color: c.text,
                padding: "8px 6px",
                borderRadius: 2,
                textAlign: "center",
                border: w.status === "current" ? "2px solid #d97706" : "none",
                position: "relative",
              }}>
                <div style={{ fontSize: 10, opacity: 0.85, fontWeight: 500 }}>{w.year}</div>
                <div className="mono-font" style={{ fontSize: 12, fontWeight: 600, marginTop: 2 }}>{fmt(w.area)}</div>
                {w.status === "current" && (
                  <div style={{
                    position: "absolute", top: -7, right: -7,
                    background: "#d97706", color: "white",
                    width: 14, height: 14, borderRadius: "50%",
                    display: "flex", alignItems: "center", justifyContent: "center",
                    fontSize: 9, fontWeight: 700,
                  }}>•</div>
                )}
              </div>
            );
          })}
        </div>
      </div>
    </section>
  );
}
function Stat({ label, value, accent }) {
  return (
    <div style={{ borderLeft: `2px solid ${accent}`, paddingLeft: 12 }}>
      <div style={{ fontSize: 10, letterSpacing: "0.12em", textTransform: "uppercase", color: "#78716c" }}>{label}</div>
      <div className="display-font mono-font" style={{ fontSize: 22, fontWeight: 600, color: accent, marginTop: 2 }}>{value}</div>
    </div>
  );
}

// ============================================================================
// REPORTS SECTION — generate downloadable docx reports
// ============================================================================
function ReportsSection({ onPick }) {
  return (
    <section>
      <SectionTitle eyebrow="Document Generator" title="Pre-built Reports & Briefings" />
      <div style={{ display: "grid", gridTemplateColumns: "repeat(3, 1fr)", gap: 16 }}>
        {REPORTS.map((r, i) => {
          const Icon = r.icon;
          return (
            <button key={r.id} onClick={()=>onPick(r)} className="fade-up" style={{
              background: "white",
              border: "1px solid rgba(13,78,44,0.15)",
              borderRadius: 4,
              padding: "20px 22px",
              textAlign: "left",
              cursor: "pointer",
              fontFamily: "inherit",
              transition: "all 0.2s ease",
              animationDelay: `${i*40}ms`,
              position: "relative",
            }}
            onMouseEnter={(e)=>{
              e.currentTarget.style.borderColor = r.color;
              e.currentTarget.style.transform = "translateY(-2px)";
            }}
            onMouseLeave={(e)=>{
              e.currentTarget.style.borderColor = "rgba(13,78,44,0.15)";
              e.currentTarget.style.transform = "translateY(0)";
            }}>
              <div style={{ display: "flex", justifyContent: "space-between", alignItems: "flex-start", marginBottom: 14 }}>
                <div style={{
                  width: 40, height: 40,
                  background: r.color + "12",
                  borderRadius: 4,
                  display: "flex", alignItems: "center", justifyContent: "center",
                }}>
                  <Icon size={20} color={r.color} strokeWidth={1.5}/>
                </div>
                <span style={{
                  fontSize: 10, color: "#78716c",
                  background: "rgba(13,78,44,0.05)",
                  padding: "2px 8px", borderRadius: 2,
                }}>{r.pages}</span>
              </div>
              <div className="display-font" style={{ fontSize: 17, fontWeight: 600, color: "#1c1917", lineHeight: 1.2, marginBottom: 4 }}>
                {r.title}
              </div>
              <div style={{ fontSize: 12, color: r.color, marginBottom: 10, fontStyle: "italic" }}>
                {r.titleHi}
              </div>
              <div style={{ fontSize: 12, color: "#78716c", lineHeight: 1.5 }}>{r.desc}</div>
              <div style={{ display: "flex", alignItems: "center", gap: 4, marginTop: 12, fontSize: 11, color: r.color, fontWeight: 500 }}>
                Generate report <ChevronRight size={12}/>
              </div>
            </button>
          );
        })}
      </div>
    </section>
  );
}

// ============================================================================
// SECTION TITLE
// ============================================================================
function SectionTitle({ eyebrow, title }) {
  return (
    <div style={{ marginBottom: 16 }}>
      <div style={{ fontSize: 10, letterSpacing: "0.22em", textTransform: "uppercase", color: "#a16207", fontWeight: 600, marginBottom: 4 }}>
        {eyebrow}
      </div>
      <h2 className="display-font" style={{ fontSize: 28, fontWeight: 600, margin: 0, color: "#1c1917", letterSpacing: "-0.02em" }}>
        {title}
      </h2>
    </div>
  );
}

const th = (align="left") => ({
  textAlign: align, fontSize: 10, letterSpacing: "0.12em", textTransform: "uppercase",
  color: "#78716c", fontWeight: 500, padding: "10px 12px",
});
const td = (align="left") => ({
  textAlign: align, padding: "12px",
});

// ============================================================================
// REPORT MODAL — preview + download
// ============================================================================
function ReportModal({ report, onClose }) {
  const [generating, setGenerating] = useState(false);

  async function handleDownload() {
    setGenerating(true);
    try {
      const blob = await generateReport(report);
      const url = URL.createObjectURL(blob);
      const a = document.createElement("a");
      a.href = url;
      a.download = `${report.id}_${DIVISION.currentApoYear}.docx`;
      document.body.appendChild(a); a.click(); document.body.removeChild(a);
      URL.revokeObjectURL(url);
    } catch (e) {
      alert("Error: " + e.message);
      console.error(e);
    } finally {
      setGenerating(false);
    }
  }

  const Icon = report.icon;

  return (
    <div onClick={onClose} style={{
      position: "fixed", inset: 0,
      background: "rgba(28,25,23,0.6)",
      backdropFilter: "blur(4px)",
      display: "flex", alignItems: "center", justifyContent: "center",
      padding: 24, zIndex: 50,
    }}>
      <div onClick={(e)=>e.stopPropagation()} style={{
        background: "white", borderRadius: 4,
        maxWidth: 640, width: "100%",
        border: `2px solid ${report.color}`,
        position: "relative",
        animation: "fadeUp 0.3s ease-out",
      }}>
        <div style={{ padding: "28px 32px 24px", borderBottom: "1px solid rgba(13,78,44,0.1)" }}>
          <div style={{ display: "flex", gap: 16, alignItems: "flex-start" }}>
            <div style={{
              width: 48, height: 48,
              background: report.color + "15",
              borderRadius: 4,
              display: "flex", alignItems: "center", justifyContent: "center",
              flexShrink: 0,
            }}>
              <Icon size={24} color={report.color} strokeWidth={1.5}/>
            </div>
            <div style={{ flex: 1 }}>
              <div className="display-font" style={{ fontSize: 22, fontWeight: 600, lineHeight: 1.2 }}>
                {report.title}
              </div>
              <div style={{ fontSize: 14, color: report.color, fontStyle: "italic", marginTop: 2 }}>
                {report.titleHi}
              </div>
            </div>
          </div>
          <div style={{ fontSize: 13, color: "#78716c", marginTop: 16, lineHeight: 1.6 }}>
            {report.desc}
          </div>
        </div>

        <div style={{ padding: "24px 32px", background: "#faf8f3" }}>
          <div style={{ fontSize: 11, letterSpacing: "0.16em", textTransform: "uppercase", color: "#78716c", fontWeight: 500, marginBottom: 10 }}>
            Report Contents
          </div>
          <div style={{ display: "flex", flexDirection: "column", gap: 6, fontSize: 13, color: "#1c1917" }}>
            {getReportContents(report.id).map((line, i) => (
              <div key={i} style={{ display: "flex", gap: 8 }}>
                <span style={{ color: report.color, marginTop: 2 }}>❦</span>
                <span>{line}</span>
              </div>
            ))}
          </div>
        </div>

        <div style={{ padding: "20px 32px", display: "flex", justifyContent: "flex-end", gap: 12, alignItems: "center" }}>
          <span style={{ fontSize: 11, color: "#78716c", marginRight: "auto", fontStyle: "italic" }}>
            Output: Word (.docx) · {report.pages}
          </span>
          <button onClick={onClose} style={{
            border: "1px solid rgba(13,78,44,0.2)", background: "white",
            padding: "10px 20px", borderRadius: 4, fontSize: 13, cursor: "pointer", color: "#78716c",
          }}>Cancel</button>
          <button onClick={handleDownload} disabled={generating} style={{
            background: report.color, color: "white", border: "none",
            padding: "10px 22px", borderRadius: 4, fontSize: 13, fontWeight: 500, cursor: "pointer",
            display: "flex", alignItems: "center", gap: 8,
            opacity: generating ? 0.6 : 1,
          }}>
            <Download size={14}/>
            {generating ? "Generating..." : "Download"}
          </button>
        </div>
      </div>
    </div>
  );
}

function getReportContents(id) {
  const map = {
    division_brief: [
      "Division masthead with circle, headquarters, plan period",
      "Headline statistics (compartments, JFMCs, area, plan year)",
      "4-range summary table with key metrics",
      "Working-circle distribution",
      "Forest Rights Act status snapshot",
    ],
    range_wise: [
      "Detailed range-by-range analytical breakdown",
      "Compartment count, JFMC count, total area per range",
      "RF/PF/OA notified blocks with area",
      "Felling-series distribution per range",
      "Per-range comparison tables",
    ],
    site_quality: [
      "Sal forest distribution by site quality (II, III, IVa, IVb)",
      "Mixed forest distribution by site quality (III–Vb)",
      "Understocked area, blanks, plantations",
      "Per-range site quality breakup",
      "Implications for working circle assignment",
    ],
    watershed_plan: [
      "10-year watershed treatment schedule (2020-21 to 2029-30)",
      "Year-wise area treatment table",
      "Progress against current APO year",
      "Per-range allocation",
      "Linked Working Plan compartment list",
    ],
    fra_status: [
      "Community Forest Rights (CFR) summary — 298 patta",
      "CFR area: 43,103 Ha across division",
      "Individual Forest Rights (IFR) — 1,174 holders",
      "Implications for forest management activities",
    ],
    jfmc_directory: [
      "Range-wise JFMC list (265 active committees)",
      "Villages covered per committee",
      "Bank account & banking branch information",
      "President & Secretary contact details (where available)",
    ],
  };
  return map[id] || [];
}

// ============================================================================
// DOCX REPORT GENERATOR
// ============================================================================
async function generateReport(report) {
  const heading = (text, level=HeadingLevel.HEADING_1) => new Paragraph({
    text, heading: level,
    spacing: { before: 280, after: 140 },
  });
  const para = (text, opts={}) => new Paragraph({
    alignment: opts.align || AlignmentType.LEFT,
    spacing: { after: 120 },
    children: [new TextRun({ text, size: opts.size||22, bold: opts.bold||false, italics: opts.italic||false })],
  });
  const cell = (text, opts={}) => new TableCell({
    shading: opts.bg ? { type: ShadingType.SOLID, color: opts.bg, fill: opts.bg } : undefined,
    children: [new Paragraph({
      alignment: opts.align || AlignmentType.LEFT,
      spacing: { before: 0, after: 0 },
      children: [new TextRun({ text: String(text||""), size: opts.size||20, bold: opts.bold||false })],
    })],
  });
  const tbl = (rows, widthsCm, opts={}) => new Table({
    width: { size: 100, type: WidthType.PERCENTAGE },
    columnWidths: widthsCm.map(c => Math.round(c*567)),
    borders: borderAll(),
    rows: rows.map((row, ri) => new TableRow({
      children: row.map(c => cell(c, {
        bold: opts.headerRow && ri===0,
        bg: opts.headerRow && ri===0 ? "0d4e2c" : undefined,
        size: 20,
      }))
    }))
  });

  // Common header for every report
  const header = [
    para("GOVERNMENT OF CHHATTISGARH", { align: AlignmentType.CENTER, bold: true, size: 24 }),
    para("Forest & Climate Change Department", { align: AlignmentType.CENTER, italic: true, size: 22 }),
    para("Bilaspur Forest Circle · Raigarh Forest Division", { align: AlignmentType.CENTER, bold: true, size: 26 }),
    para(`Working Plan Period: ${DIVISION.planPeriod}`, { align: AlignmentType.CENTER, size: 20 }),
    para(`APO Year: ${DIVISION.currentApoYear} (Year ${DIVISION.planYearIndex} of 10)`, { align: AlignmentType.CENTER, italic: true, size: 20 }),
    para(""),
    para(report.title.toUpperCase(), { align: AlignmentType.CENTER, bold: true, size: 28 }),
    para(report.titleHi, { align: AlignmentType.CENTER, italic: true, size: 22 }),
    para(""),
  ];

  // Body builders per report
  const body = buildReportBody(report.id, { tbl, para, heading, cell });

  const doc = new Document({
    creator: "Raigarh Division Dashboard",
    title: report.title,
    sections: [{
      properties: { page: { margin: { top: 1134, bottom: 1134, left: 1247, right: 1247 } } },
      children: [...header, ...body, para(""),
        para(`Prepared on ${new Date().toLocaleDateString("en-IN", { day:"numeric", month:"long", year:"numeric" })} · Source: Working Plan ${DIVISION.planPeriod}`,
             { italic: true, size: 18, align: AlignmentType.CENTER })],
    }],
  });
  return await Packer.toBlob(doc);
}

function buildReportBody(id, h) {
  const { tbl, para, heading } = h;

  if (id === "division_brief") {
    return [
      heading("1. Division at a Glance"),
      para(`The Raigarh Forest Division covers a managed forest area of approximately ${fmt(TOT.area)} hectares across ${TOT.compartments} compartments, distributed over four administrative ranges. The division is currently in Year ${DIVISION.planYearIndex} of its decadal Working Plan ${DIVISION.planPeriod}.`),
      tbl([
        ["Parameter", "Value"],
        ["Forest Circle", DIVISION.circle],
        ["State", DIVISION.state],
        ["Division Headquarters", DIVISION.hq],
        ["Working Plan Period", DIVISION.planPeriod],
        ["Current APO Year", DIVISION.currentApoYear],
        ["Total Compartments", String(TOT.compartments)],
        ["Total JFM Committees", String(TOT.jfmcs)],
        ["Total Forest Area", `${fmt(TOT.area)} Ha`],
        ["Reserved Forest", `${TOT.rf_count} blocks · ${fmt(TOT.rf_area)} Ha`],
        ["Protected Forest", `${TOT.pf_count} blocks · ${fmt(TOT.pf_area)} Ha`],
        ["Orange Area", `${TOT.oa_count} blocks · ${fmt(TOT.oa_area)} Ha`],
      ], [6, 9], { headerRow: true }),

      heading("2. Range-wise Summary"),
      tbl([
        ["Range","Compt.","JFMCs","Area (Ha)"],
        ...RANGES.map(r => [r, String(RANGE_DATA[r].compartments), String(RANGE_DATA[r].jfmcs), fmt(RANGE_DATA[r].area_ha)]),
        ["Total", String(TOT.compartments), String(TOT.jfmcs), fmt(TOT.area)],
      ], [4, 3, 3, 5], { headerRow: true }),

      heading("3. Working Circles"),
      para(`The division operates four working circles distributed across compartments:`),
      tbl([
        ["Working Circle","Code","Compartments"],
        ["Protection","PWC", String(WC_COUNTS.PWC)],
        ["Selection-cum-Improvement","SCI", String(WC_COUNTS.SCI)],
        ["Rehabilitation","RWC", String(WC_COUNTS.RWC)],
        ["Improvement","IWC", String(WC_COUNTS.IWC)],
      ], [6, 3, 6], { headerRow: true }),

      heading("4. Forest Rights Act"),
      para(`Within the division boundary, ${FRA.cfr_count} community forest rights (CFR) titles have been granted covering ${fmt(FRA.cfr_area_ha)} hectares, alongside ${fmt(FRA.ifr_count)} individual forest rights (IFR) titles.`),
    ];
  }

  if (id === "range_wise") {
    const out = [
      heading("Range-wise Statistical Bulletin"),
      para("This bulletin presents detailed statistics for each of the four administrative ranges of Raigarh Forest Division."),
    ];
    RANGES.forEach((r, i) => {
      const d = RANGE_DATA[r];
      out.push(heading(`${i+1}. ${r} Range (${RANGES_HI[r]})`, HeadingLevel.HEADING_2));
      out.push(para(`The ${r} range contains ${d.compartments} compartments and ${d.jfmcs} JFM committees, managing ${fmt(d.area_ha)} hectares of forest land.`));
      out.push(tbl([
        ["Parameter","Value"],
        ["Compartments", String(d.compartments)],
        ["JFM Committees", String(d.jfmcs)],
        ["Total Forest Area", `${fmt(d.area_ha)} Ha`],
        ["Reserved Forest blocks", `${d.rf_count} blocks`],
        ["RF area", `${fmt(d.rf_area)} Ha`],
        ["Protected Forest blocks", `${d.pf_count} blocks`],
        ["PF area", `${fmt(d.pf_area)} Ha`],
        ["Orange Area blocks", `${d.oa_count} blocks`],
        ["OA area", `${fmt(d.oa_area)} Ha`],
      ], [6, 9], { headerRow: true }));
    });
    out.push(heading("Comparative Summary"));
    out.push(tbl([
      ["Range","Compt.","JFMCs","Area","RF","PF","OA"],
      ...RANGES.map(r => {
        const d = RANGE_DATA[r];
        return [r, String(d.compartments), String(d.jfmcs), fmt(d.area_ha),
                String(d.rf_count), String(d.pf_count), String(d.oa_count)];
      }),
      ["Total", String(TOT.compartments), String(TOT.jfmcs), fmt(TOT.area),
       String(TOT.rf_count), String(TOT.pf_count), String(TOT.oa_count)],
    ], [3, 2, 2, 3, 2, 2, 2], { headerRow: true }));
    return out;
  }

  if (id === "site_quality") {
    const sal = SITE_QUALITY.filter(s => s.group === "Sal");
    const mixed = SITE_QUALITY.filter(s => s.group === "Mixed");
    const other = SITE_QUALITY.filter(s => s.group === "Other" || s.group === "Plantation");
    const totalArea = SITE_QUALITY.reduce((s,x)=>s+x.area,0);

    return [
      heading("Forest Composition & Site Quality"),
      para(`This report presents the site quality distribution across the entire managed area of Raigarh Forest Division. Total enumerated area is ${fmt(totalArea)} hectares.`),

      heading("1. Sal Forest", HeadingLevel.HEADING_2),
      tbl([
        ["Site Quality","Area (Ha)","% of Sal"],
        ...sal.map(s => {
          const totalSal = sal.reduce((a,x)=>a+x.area,0);
          return [s.name, fmt(s.area), ((s.area/totalSal)*100).toFixed(1) + "%"];
        }),
        ["Total Sal", fmt(sal.reduce((a,x)=>a+x.area,0)), "100.0%"],
      ], [6, 5, 4], { headerRow: true }),

      heading("2. Mixed Forest", HeadingLevel.HEADING_2),
      tbl([
        ["Site Quality","Area (Ha)","% of Mixed"],
        ...mixed.map(s => {
          const totalMixed = mixed.reduce((a,x)=>a+x.area,0);
          return [s.name, fmt(s.area), ((s.area/totalMixed)*100).toFixed(1) + "%"];
        }),
        ["Total Mixed", fmt(mixed.reduce((a,x)=>a+x.area,0)), "100.0%"],
      ], [6, 5, 4], { headerRow: true }),

      heading("3. Non-forest & Special Categories", HeadingLevel.HEADING_2),
      tbl([
        ["Category","Area (Ha)"],
        ...other.map(s => [s.name, fmt(s.area)]),
        ["Total", fmt(other.reduce((a,x)=>a+x.area,0))],
      ], [9, 6], { headerRow: true }),

      heading("4. Implications"),
      para(`A significant proportion of the division (${fmt(SITE_QUALITY.find(s=>s.name==="Sal SQ-IVa").area)} Ha or ${((SITE_QUALITY.find(s=>s.name==="Sal SQ-IVa").area/totalArea)*100).toFixed(1)}%) falls in Sal SQ-IVa. This is consistent with degraded-but-recoverable forest, suitable for Selection-cum-Improvement and Rehabilitation operations.`),
      para(`Understocked area of ${fmt(SITE_QUALITY.find(s=>s.name==="Understocked").area)} Ha and blank area of ${fmt(SITE_QUALITY.find(s=>s.name==="Blank").area)} Ha together comprise ${fmt(SITE_QUALITY.find(s=>s.name==="Understocked").area + SITE_QUALITY.find(s=>s.name==="Blank").area)} Ha that are candidates for active rehabilitation through CPT works, invasive species removal, and silvicultural intervention.`),
    ];
  }

  if (id === "watershed_plan") {
    const total = WATERSHED_SCHEDULE.reduce((s,w)=>s+w.area, 0);
    const completed = WATERSHED_SCHEDULE.filter(w=>w.status==="completed").reduce((s,w)=>s+w.area, 0);
    const pct = (completed/total*100).toFixed(1);
    return [
      heading("10-Year Watershed Treatment Schedule"),
      para(`As per the Working Plan ${DIVISION.planPeriod}, watershed treatment is scheduled across ${fmt(total)} hectares over ten years. As of the start of APO ${DIVISION.currentApoYear}, ${fmt(completed)} hectares (${pct}%) of the planned area has been treated.`),
      tbl([
        ["Year","Status","Area (Ha)"],
        ...WATERSHED_SCHEDULE.map(w => [w.year, w.status === "completed" ? "Completed" : w.status === "current" ? "Current" : "Future", fmt(w.area)]),
        ["Total","",fmt(total)],
      ], [4, 5, 5], { headerRow: true }),

      heading("Status Summary"),
      tbl([
        ["Phase","Years","Area (Ha)","%"],
        ["Completed", "2020-21 to 2025-26", fmt(completed), pct + "%"],
        ["Current", "2026-27", fmt(WATERSHED_SCHEDULE.find(w=>w.status==="current").area), ((WATERSHED_SCHEDULE.find(w=>w.status==="current").area/total)*100).toFixed(1)+"%"],
        ["Future", "2027-28 to 2029-30", fmt(WATERSHED_SCHEDULE.filter(w=>w.status==="future").reduce((s,w)=>s+w.area,0)), ((WATERSHED_SCHEDULE.filter(w=>w.status==="future").reduce((s,w)=>s+w.area,0)/total)*100).toFixed(1)+"%"],
      ], [4, 5, 4, 3], { headerRow: true }),

      heading("Next Steps"),
      para(`For APO ${DIVISION.currentApoYear}, ${fmt(WATERSHED_SCHEDULE.find(w=>w.status==="current").area)} hectares are scheduled for treatment. This is to be distributed across ranges through individual Detail Project Reports prepared at compartment level by the concerned Range Officer, vetted by the Sub-Divisional Officer, and sanctioned by the Divisional Forest Officer.`),
    ];
  }

  if (id === "fra_status") {
    return [
      heading("Forest Rights Act — Status Within Division"),
      para(`The Forest Rights Act 2006 (FRA) provides recognition and vesting of forest rights to forest-dwelling Scheduled Tribes and other traditional forest dwellers. This report summarises CFR (Community Forest Rights) and IFR (Individual Forest Rights) granted within the boundary of Raigarh Forest Division.`),

      heading("1. Community Forest Rights (CFR)", HeadingLevel.HEADING_2),
      tbl([
        ["Parameter","Value"],
        ["Total CFR titles granted", fmt(FRA.cfr_count)],
        ["Total CFR area", `${fmt(FRA.cfr_area_ha)} Ha`],
        ["Average area per CFR", `${fmt(FRA.cfr_area_ha/FRA.cfr_count)} Ha`],
      ], [6, 9], { headerRow: true }),

      heading("2. Individual Forest Rights (IFR)", HeadingLevel.HEADING_2),
      tbl([
        ["Parameter","Value"],
        ["Total IFR titles granted", fmt(FRA.ifr_count)],
      ], [6, 9], { headerRow: true }),

      heading("3. Implications for Forest Management"),
      para(`CFR areas overlap with managed forest compartments and require coordination between forest management activities and the gram sabha exercising customary rights. Working Plan prescriptions in CFR compartments are subject to consultation under the Act.`),
      para(`The 298 CFR titles cover ${fmt(FRA.cfr_area_ha)} hectares — a significant portion of the division — and shape the operational space for activities such as regulated extraction, plantation, and protection works.`),
    ];
  }

  if (id === "jfmc_directory") {
    return [
      heading("JFM Committee Directory"),
      para(`Raigarh Forest Division operates through ${TOT.jfmcs} active Joint Forest Management Committees (JFMCs / वन प्रबंधन समितियाँ). These committees, formed under the JFM resolution, partner with the Forest Department for protection, regeneration, and benefit-sharing.`),

      heading("Range-wise Distribution"),
      tbl([
        ["Range","JFMCs","Compartments","Avg. Compt. per JFMC"],
        ...RANGES.map(r => {
          const d = RANGE_DATA[r];
          return [r, String(d.jfmcs), String(d.compartments), (d.compartments/d.jfmcs).toFixed(2)];
        }),
        ["Total", String(TOT.jfmcs), String(TOT.compartments), (TOT.compartments/TOT.jfmcs).toFixed(2)],
      ], [4, 3, 4, 4], { headerRow: true }),

      heading("Note"),
      para(`This summary report shows aggregate JFMC counts. The full directory with committee names, included villages, bank account details, and office bearer information is maintained in the Working Plan Annexure XVI and is available in the division office.`),
      para(`For DPR preparation, the relevant JFMC details (name, bank account, villages, president, secretary) are required and form part of the 27-point project report at compartment level.`),
    ];
  }

  return [para("Report content not yet defined for: " + id)];
}

function borderAll() {
  const b = { style: BorderStyle.SINGLE, size: 4, color: "0d4e2c" };
  return { top: b, bottom: b, left: b, right: b, insideHorizontal: b, insideVertical: b };
}
