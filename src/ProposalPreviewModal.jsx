import { useState, useCallback, useEffect } from "react";
import { fmtM, fmtK, fmtDollar } from "./formatters";
import { recomputeReviewedData } from "./proposalUpgrades";
import { getAllFunds, getSubGroupTotals } from "./portfolioData";
import { getFundAlternatives } from "./fundAlternatives";

// ─── Inline-editable row ─────────────────────────────────────────────────

function Row({ label, value, sub, highlight, field, localData, onChange, placeholder }) {
  const rawVal    = field ? (localData[field] ?? "") : "";
  const isEmpty   = field && !rawVal; // treat 0 as empty (no valid field here uses 0 as a real value)
  const [editing, setEditing] = useState(false);
  const [draft, setDraft]     = useState("");

  const isEditable = !!field;

  function startEdit() {
    if (!isEditable) return;
    setDraft(isEmpty ? "" : String(rawVal));
    setEditing(true);
  }

  function commit() {
    const n = parseFloat(draft);
    if (!isNaN(n)) onChange(field, n);
    setEditing(false);
  }

  // Auto-open input for empty editable fields
  const showInput = editing || (isEditable && isEmpty);

  return (
    <div style={{
      display: "flex", justifyContent: "space-between", alignItems: "center",
      padding: "7px 0", borderBottom: "1px solid #EEF0F5",
      cursor: isEditable && !showInput ? "pointer" : "default",
    }}>
      <span style={{ fontSize: 12, color: "#6b7a99", flexShrink: 0, marginRight: 8 }}>{label}</span>
      <div style={{ display: "flex", alignItems: "center", gap: 5 }}>
        {showInput ? (
          <input
            autoFocus={editing}
            value={draft}
            placeholder={placeholder || "enter value"}
            onChange={e => setDraft(e.target.value)}
            onFocus={() => { if (!editing) { setDraft(""); setEditing(true); } }}
            onBlur={commit}
            onKeyDown={e => { if (e.key === "Enter") commit(); if (e.key === "Escape") setEditing(false); }}
            style={{
              width: 110, textAlign: "right", fontSize: 13, fontWeight: 700,
              color: highlight || "var(--navy)",
              border: "1.5px solid " + (highlight || "var(--blue)"),
              borderRadius: 5, padding: "3px 8px",
              background: "#F4F8FF", outline: "none",
            }}
          />
        ) : (
          <>
            <span
              onClick={startEdit}
              title="Click to edit"
              style={{
                fontSize: 14, fontWeight: 700, color: highlight || "var(--navy)",
                borderBottom: `1px dashed ${highlight || "#9bacc8"}`,
              }}
            >{value}</span>
            <span onClick={startEdit} style={{ fontSize: 10, color: "#9bacc8", userSelect: "none" }}>✏</span>
          </>
        )}
      </div>
    </div>
  );
}

// Sub-label shown on its own line when needed
function RowSub({ sub }) {
  return sub
    ? <div style={{ fontSize: 10, color: "#9bacc8", textAlign: "right", marginTop: -4, marginBottom: 2 }}>{sub}</div>
    : null;
}

// ─── Allocation row (shows both $ and % editable) ────────────────────────

function AllocationRow({ label, pctField, baseAmount, minDollarM, pctMin = 0, pctMax = 100, localData, onChange, accent, note }) {
  const pctVal   = localData[pctField] ?? 30;
  const rawDollar = (baseAmount || 0) * (pctVal / 100);
  const dollar    = minDollarM ? Math.max(rawDollar, minDollarM) : rawDollar;
  // Sub-$1M amounts (e.g. CRT income) are edited/displayed in thousands for readability.
  const isK = dollar < 1;

  const [editingDollar, setEditingDollar] = useState(false);
  const [draftDollar,   setDraftDollar]   = useState("");
  const [editingPct,    setEditingPct]    = useState(false);
  const [draftPct,      setDraftPct]      = useState("");

  function commitDollar() {
    const raw = parseFloat(draftDollar);
    if (!isNaN(raw) && baseAmount > 0) {
      const n = isK ? raw / 1000 : raw; // K → M if editing in thousands
      // Convert M → pct, clamp to [pctMin, pctMax]
      const newPct = Math.min(pctMax, Math.max(pctMin, (n / baseAmount) * 100));
      onChange(pctField, newPct);
    }
    setEditingDollar(false);
  }

  function commitPct() {
    const n = parseFloat(draftPct);
    if (!isNaN(n)) onChange(pctField, Math.min(pctMax, Math.max(pctMin, n)));
    setEditingPct(false);
  }

  const na = !baseAmount;
  const ac = accent || "var(--navy)";

  return (
    <div>
      <div style={{
        display: "flex", justifyContent: "space-between", alignItems: "center",
        padding: "8px 0", borderBottom: "1px solid #EEF0F5",
      }}>
        <span style={{ fontSize: 12, color: "#6b7a99", flexShrink: 0, marginRight: 8 }}>{label}</span>
        {na ? (
          <span style={{ fontSize: 13, fontWeight: 700, color: "#9bacc8" }}>—</span>
        ) : (
          <div style={{ display: "flex", alignItems: "center", gap: 10 }}>
            {/* Dollar amount */}
            {editingDollar ? (
              <input
                autoFocus
                value={draftDollar}
                placeholder={isK ? "e.g. 75 (K)" : "e.g. 3.0 (M)"}
                onChange={e => setDraftDollar(e.target.value)}
                onBlur={commitDollar}
                onKeyDown={e => { if (e.key === "Enter") commitDollar(); if (e.key === "Escape") setEditingDollar(false); }}
                style={{
                  width: 80, textAlign: "right", fontSize: 13, fontWeight: 700,
                  color: ac, border: `1.5px solid ${ac}`, borderRadius: 5,
                  padding: "3px 7px", background: "#F4F8FF", outline: "none",
                }}
              />
            ) : (
              <span
                onClick={() => { setDraftDollar(isK ? Math.round(dollar * 1000).toString() : dollar.toFixed(2)); setEditingDollar(true); }}
                style={{ fontSize: 14, fontWeight: 700, color: ac, borderBottom: `1px dashed ${ac}`, cursor: "pointer" }}
                title="Click to edit dollar amount"
              >
                {fmtK(dollar)}
              </span>
            )}
            <span style={{ fontSize: 11, color: "#9bacc8" }}>·</span>
            {/* Percentage */}
            {editingPct ? (
              <input
                autoFocus
                value={draftPct}
                placeholder="%"
                onChange={e => setDraftPct(e.target.value)}
                onBlur={commitPct}
                onKeyDown={e => { if (e.key === "Enter") commitPct(); if (e.key === "Escape") setEditingPct(false); }}
                style={{
                  width: 60, textAlign: "right", fontSize: 13, fontWeight: 700,
                  color: "var(--gold)", border: "1.5px solid var(--gold)", borderRadius: 5,
                  padding: "3px 7px", background: "#FEF9EE", outline: "none",
                }}
              />
            ) : (
              <span
                onClick={() => { setDraftPct(pctVal.toFixed(1)); setEditingPct(true); }}
                style={{ fontSize: 12, fontWeight: 600, color: "var(--gold)", borderBottom: "1px dashed var(--gold-light)", cursor: "pointer" }}
                title="Click to edit percentage"
              >
                {pctVal.toFixed(1)}%
              </span>
            )}
            <span style={{ fontSize: 10, color: "#9bacc8" }}>✏</span>
          </div>
        )}
      </div>
      {note && !na && (
        <div style={{ fontSize: 10, color: "#9bacc8", textAlign: "right", marginTop: -2, marginBottom: 4 }}>{note}</div>
      )}
    </div>
  );
}

// ─── Card ─────────────────────────────────────────────────────────────────

function Card({ title, accent, children }) {
  return (
    <div style={{ background: "#fff", borderRadius: 14, border: `1.5px solid ${accent || "#E8ECF0"}`, padding: "16px 18px", boxShadow: "0 1px 3px rgba(15,23,42,0.04)" }}>
      {title && (
        <div style={{ fontSize: 10, fontWeight: 700, color: accent || "var(--gold)", textTransform: "uppercase", letterSpacing: "0.07em", marginBottom: 8 }}>
          {title}
        </div>
      )}
      {children}
    </div>
  );
}

// ─── Section ──────────────────────────────────────────────────────────────

function Section({ title, children }) {
  return (
    <div style={{ marginBottom: 28 }}>
      <div style={{ fontSize: 11, fontWeight: 700, color: "var(--gold)", textTransform: "uppercase", letterSpacing: "0.08em", borderBottom: "2px solid var(--gold-light)", paddingBottom: 6, marginBottom: 14 }}>
        {title}
      </div>
      {children}
    </div>
  );
}

// ─── Donut chart ──────────────────────────────────────────────────────────

function DonutChart({ value, total, color, label, sub }) {
  const r = 44, circ = 2 * Math.PI * r;
  const dash = total > 0 ? Math.min(value / total, 1) * circ : 0;
  return (
    <div style={{ textAlign: "center" }}>
      <svg width="100" height="100" viewBox="0 0 100 100">
        <circle cx="50" cy="50" r={r} fill="none" stroke="#E8ECF0" strokeWidth="11" />
        <circle cx="50" cy="50" r={r} fill="none" stroke={color} strokeWidth="11"
          strokeDasharray={`${dash} ${circ - dash}`} strokeLinecap="round" transform="rotate(-90 50 50)" />
        <text x="50" y="46" textAnchor="middle" fontSize="14" fontWeight="700" fill="var(--navy)">
          {total > 0 ? `${(Math.min(value/total,1)*100).toFixed(0)}%` : "—"}
        </text>
        <text x="50" y="60" textAnchor="middle" fontSize="8" fill="#6b7a99">{sub}</text>
      </svg>
      <div style={{ fontSize: 10, color: "#6b7a99", marginTop: 2 }}>{label}</div>
    </div>
  );
}

// ─── Risk / Return Scatter Chart ──────────────────────────────────────────
// Standard efficient frontier: X = Risk (volatility), Y = Return.
// Includes gap annotations showing exactly how much excess risk the current
// portfolio carries, and a callout summary below.

function RiskReturnChart({ points }) {
  const valid = points.filter(p => p.x != null && p.y != null);
  if (valid.length === 0) return null;

  const target  = valid[0]; // assumed on frontier
  const current = valid[1] || null;

  // ── Efficient frontier ───────────────────────────────────────────────────
  // Use a square-root curve: μ(σ) = mv_ret + scale·√(σ − mv_sig)
  // This produces the classic Investopedia shape — steeply rising on the left,
  // visibly flattening to the right — unlike a hyperbola which becomes linear.
  // Anchored so the curve passes exactly through the Target point.
  const mv_sig = target.x * 0.28;          // min-variance vol (~28% of target's)
  const mv_ret = target.y * 0.30;          // min-variance return (~30% of target's)
  // scale from target: target.y = mv_ret + scale·√(target.x − mv_sig)
  const sqrtScale = (target.y - mv_ret) / Math.sqrt(target.x - mv_sig);

  // μ at a given σ (returns null left of the minimum-variance point)
  const frontierRet = sig => {
    const d = sig - mv_sig;
    return d >= 0 ? mv_ret + sqrtScale * Math.sqrt(d) : null;
  };
  // Inverse: σ on frontier to achieve a given return
  const frontierSig = ret => {
    if (ret <= mv_ret) return mv_sig;
    return mv_sig + ((ret - mv_ret) / sqrtScale) ** 2;
  };

  // Extend curve well past data so flattening is clearly visible
  const sigMax  = Math.max(...valid.map(p => p.x)) * 2.6;
  const nSeg    = 120;
  const frontierPts = Array.from({ length: nSeg + 1 }, (_, i) => {
    const sig = mv_sig + (sigMax - mv_sig) * i / nSeg;
    const ret = frontierRet(sig);
    return ret != null ? { x: sig, y: ret } : null;
  }).filter(Boolean);

  // ── Gap annotations (when current exists) ───────────────────────────────
  // 1. Horizontal: vol needed on frontier to match current's return
  const effSigAtCurrentRet = current ? frontierSig(current.y) : null;
  // 2. Vertical: return achievable on frontier at current's vol
  const effRetAtCurrentSig = current ? frontierRet(current.x) : null;

  const deltaVolPct = current ? (current.x - target.x) * 100 : null;
  const deltaRetPct = current ? (target.y - current.y) * 100 : null;
  const excessRiskPct = current && effSigAtCurrentRet != null
    ? (current.x - effSigAtCurrentRet) * 100 : null;
  const foregoneRetPct = current && effRetAtCurrentSig != null
    ? (effRetAtCurrentSig - current.y) * 100 : null;

  // ── Axis bounds ──────────────────────────────────────────────────────────
  // X: start just left of the minimum-variance point so the steep rise is visible.
  // Y: span from mv_ret (bottom of curve) up through all data + full frontier.
  const xa = mv_sig * 0.55;
  const xb = sigMax * 1.03;

  const allY = [...valid.map(p => p.y), mv_ret, ...frontierPts.map(p => p.y)];
  const ySpan = Math.max(...allY) - Math.min(...allY);
  const ya = Math.min(...allY) - ySpan * 0.10;
  const yb = Math.max(...allY) + ySpan * 0.22;

  const W = 500, H = 380;
  const ml = 58, mr = 22, mt = 22, mb = 46;
  const pw = W - ml - mr, ph = H - mt - mb;

  const toX = v => ml + ((v - xa) / (xb - xa)) * pw;
  const toY = v => mt + ph - ((v - ya) / (yb - ya)) * ph;

  const pathD = frontierPts
    .map(({ x, y }, i) => `${i === 0 ? "M" : "L"}${toX(x).toFixed(1)},${toY(y).toFixed(1)}`)
    .join(" ");

  // Frontier label: ~65% along curve
  const lIdx  = Math.floor(frontierPts.length * 0.65);
  const lPt   = frontierPts[lIdx];
  const lPtN  = frontierPts[Math.min(lIdx + 5, frontierPts.length - 1)];
  const lAng  = Math.atan2(toY(lPtN.y) - toY(lPt.y), toX(lPtN.x) - toX(lPt.x)) * 180 / Math.PI;

  const N = 4;
  const gxs = Array.from({ length: N + 1 }, (_, i) => xa + (xb - xa) * i / N);
  const gys = Array.from({ length: N + 1 }, (_, i) => ya + (yb - ya) * i / N);
  const clipId = "efClip2";

  return (
    <div style={{ marginTop: 18, marginBottom: 4 }}>
      <div style={{ fontSize: 11, fontWeight: 700, color: "#9bacc8", textTransform: "uppercase", letterSpacing: "0.06em", marginBottom: 6 }}>
        Risk / Return — Efficient Frontier
      </div>

      <svg viewBox={`0 0 ${W} ${H}`} style={{ width: "100%", height: "auto", display: "block" }}>
        <defs>
          <clipPath id={clipId}>
            <rect x={ml} y={mt} width={pw} height={ph} />
          </clipPath>
          {/* Arrowhead marker */}
          <marker id="arrowCoral" markerWidth="7" markerHeight="7" refX="6" refY="3.5" orient="auto">
            <polygon points="0 0, 7 3.5, 0 7" fill="var(--coral)" opacity="0.8" />
          </marker>
          <marker id="arrowTeal" markerWidth="7" markerHeight="7" refX="6" refY="3.5" orient="auto">
            <polygon points="0 0, 7 3.5, 0 7" fill="var(--teal)" opacity="0.8" />
          </marker>
        </defs>

        {/* Grid */}
        {gxs.map((v, i) => (
          <g key={`gx${i}`}>
            <line x1={toX(v)} y1={mt} x2={toX(v)} y2={mt + ph} stroke="#EEF0F5" strokeWidth="1" />
            <text x={toX(v)} y={mt + ph + 14} textAnchor="middle" fontSize="9" fill="#9bacc8">
              {(v * 100).toFixed(1)}%
            </text>
          </g>
        ))}
        {gys.map((v, i) => (
          <g key={`gy${i}`}>
            <line x1={ml} y1={toY(v)} x2={ml + pw} y2={toY(v)} stroke="#EEF0F5" strokeWidth="1" />
            <text x={ml - 6} y={toY(v) + 3.5} textAnchor="end" fontSize="9" fill="#9bacc8">
              {(v * 100).toFixed(1)}%
            </text>
          </g>
        ))}

        {/* Axes */}
        <line x1={ml} y1={mt} x2={ml} y2={mt + ph} stroke="#C8D0DE" strokeWidth="1.5" />
        <line x1={ml} y1={mt + ph} x2={ml + pw} y2={mt + ph} stroke="#C8D0DE" strokeWidth="1.5" />

        <text x={ml + pw / 2} y={H - 4} textAnchor="middle" fontSize="10" fontWeight="600" fill="#6b7a99">
          Risk — Annualized Volatility
        </text>
        <text transform={`rotate(-90 14 ${mt + ph / 2})`} x={14} y={mt + ph / 2}
          textAnchor="middle" fontSize="10" fontWeight="600" fill="#6b7a99">
          Return — 10-Yr Avg
        </text>

        {/* Sub-optimal shading */}
        <g clipPath={`url(#${clipId})`}>
          <path
            d={`${pathD} L${toX(sigMax).toFixed(1)},${toY(ya).toFixed(1)} L${toX(mv_sig).toFixed(1)},${toY(ya).toFixed(1)} Z`}
            fill="#FDF0ED" opacity="0.6"
          />
          <path d={pathD} fill="none" stroke="var(--navy)" strokeWidth="2.4" opacity="0.8" />
          <circle cx={toX(mv_sig)} cy={toY(mv_ret)} r="4" fill="var(--navy)" opacity="0.55" />
        </g>

        {/* Frontier label */}
        <text transform={`translate(${toX(lPt.x)},${toY(lPt.y)}) rotate(${lAng})`}
          dy="-9" textAnchor="middle" fontSize="8.5" fontWeight="700" fontStyle="italic"
          fill="var(--navy)" opacity="0.7">
          Efficient Frontier (Illustrative)
        </text>

        {/* ── Gap annotations (only when current exists) ── */}
        {current && effSigAtCurrentRet != null && (() => {
          const fx = toX(effSigAtCurrentRet); // frontier x at current's return level
          const cy_ = toY(current.y);          // current's y position (its return level)
          const cx_ = toX(current.x);          // current's x position

          // Horizontal arrow: frontier point → current (shows excess risk)
          const showHoriz = cx_ - fx > 8;

          // Vertical arrow: current → frontier point at current's vol (shows foregone return)
          const fRetY = effRetAtCurrentSig != null ? toY(effRetAtCurrentSig) : null;
          const showVert = fRetY != null && cy_ - fRetY > 8;

          return (
            <g clipPath={`url(#${clipId})`}>
              {/* Ghost dot on frontier at current's return level */}
              {showHoriz && (
                <circle cx={fx} cy={cy_} r="4" fill="var(--coral)" opacity="0.35" />
              )}

              {/* Horizontal dashed line: frontier → current */}
              {showHoriz && (
                <>
                  <line x1={fx} y1={cy_} x2={cx_ - 10} y2={cy_}
                    stroke="var(--coral)" strokeWidth="1.5" strokeDasharray="4 2.5"
                    opacity="0.7" markerEnd="url(#arrowCoral)" />
                  <rect x={(fx + cx_) / 2 - 34} y={cy_ - 18} width="68" height="14"
                    rx="3" fill="white" opacity="0.85" />
                  <text x={(fx + cx_) / 2} y={cy_ - 8}
                    textAnchor="middle" fontSize="9" fontWeight="700" fill="var(--coral)">
                    +{excessRiskPct != null ? excessRiskPct.toFixed(1) : "?"}% excess risk
                  </text>
                </>
              )}

              {/* Vertical dashed line: current → frontier (foregone return) */}
              {showVert && (
                <>
                  <line x1={cx_} y1={cy_ - 10} x2={cx_} y2={fRetY + 10}
                    stroke="var(--teal)" strokeWidth="1.5" strokeDasharray="4 2.5"
                    opacity="0.7" markerEnd="url(#arrowTeal)" />
                  <rect x={cx_ + 7} y={(cy_ + fRetY) / 2 - 7} width="72" height="14"
                    rx="3" fill="white" opacity="0.85" />
                  <text x={cx_ + 43} y={(cy_ + fRetY) / 2 + 3}
                    textAnchor="middle" fontSize="9" fontWeight="700" fill="var(--teal)">
                    +{foregoneRetPct != null ? foregoneRetPct.toFixed(1) : "?"}% more return
                  </text>
                </>
              )}
            </g>
          );
        })()}

        {/* Data points */}
        {valid.map((p, i) => {
          const cx = toX(p.x), cy = toY(p.y);
          const isTarget = i === 0;
          return (
            <g key={i}>
              {isTarget && <circle cx={cx} cy={cy} r="15" fill={p.color} opacity="0.12" />}
              <circle cx={cx} cy={cy} r="9" fill={p.color} opacity="0.92" />
              <text x={cx} y={isTarget ? cy - 14 : cy + 22}
                textAnchor="middle" fontSize="10.5" fontWeight="800" fill={p.color}>
                {p.label}
              </text>
              <text x={cx} y={isTarget ? cy - 3 : cy + 33}
                textAnchor="middle" fontSize="8.5" fill="#9bacc8">
                {(p.x * 100).toFixed(1)}% vol · {(p.y * 100).toFixed(1)}% ret
              </text>
            </g>
          );
        })}
      </svg>

      {/* ── Below-chart summary callout ── */}
      {current && (
        <div style={{
          display: "grid",
          gridTemplateColumns: deltaRetPct != null ? "1fr 1fr 1fr" : "1fr 1fr",
          gap: 10, marginTop: 12,
        }}>
          {/* Volatility reduction */}
          <div style={{
            background: "linear-gradient(135deg,#E8F6F5,#F2FBFA)",
            border: "1.5px solid var(--teal)", borderRadius: 10, padding: "10px 12px",
          }}>
            <div style={{ fontSize: 9, textTransform: "uppercase", color: "#6b7a99", letterSpacing: "0.05em", marginBottom: 3 }}>
              Volatility Reduction
            </div>
            <div style={{ fontSize: 22, fontWeight: 800, color: "var(--teal)", lineHeight: 1 }}>
              {deltaVolPct != null ? `${deltaVolPct > 0 ? "−" : "+"}${Math.abs(deltaVolPct).toFixed(1)}%` : "—"}
            </div>
            <div style={{ fontSize: 10, color: "#6b7a99", marginTop: 3 }}>
              {deltaVolPct > 0 ? "less risk vs. current" : "more risk vs. current"}
            </div>
          </div>

          {/* Excess risk current carries */}
          {excessRiskPct != null && excessRiskPct > 0.05 && (
            <div style={{
              background: "linear-gradient(135deg,#FEF0ED,#FFF7F5)",
              border: "1.5px solid var(--coral)", borderRadius: 10, padding: "10px 12px",
            }}>
              <div style={{ fontSize: 9, textTransform: "uppercase", color: "#6b7a99", letterSpacing: "0.05em", marginBottom: 3 }}>
                Current Excess Risk
              </div>
              <div style={{ fontSize: 22, fontWeight: 800, color: "var(--coral)", lineHeight: 1 }}>
                +{excessRiskPct.toFixed(1)}%
              </div>
              <div style={{ fontSize: 10, color: "#6b7a99", marginTop: 3 }}>
                beyond what's needed for same return
              </div>
            </div>
          )}

          {/* Return delta */}
          {deltaRetPct != null && (
            <div style={{
              background: "linear-gradient(135deg,#EEF2FA,#F5F8FF)",
              border: "1.5px solid #9BACC8", borderRadius: 10, padding: "10px 12px",
            }}>
              <div style={{ fontSize: 9, textTransform: "uppercase", color: "#6b7a99", letterSpacing: "0.05em", marginBottom: 3 }}>
                Return Differential
              </div>
              <div style={{ fontSize: 22, fontWeight: 800, color: "var(--navy)", lineHeight: 1 }}>
                {deltaRetPct >= 0 ? "+" : ""}{deltaRetPct.toFixed(1)}%
              </div>
              <div style={{ fontSize: 10, color: "#6b7a99", marginTop: 3 }}>
                {deltaRetPct >= 0 ? "target outperformed" : "target underperformed"} (10-yr avg)
              </div>
            </div>
          )}
        </div>
      )}
    </div>
  );
}

// ─── Main ─────────────────────────────────────────────────────────────────

// Sub-group color palette — must match pptGenerator.js SEGMENTS
const SUBGROUP_COLORS = {
  "Domestic Equity":  "var(--navy)",
  "Intl Developed":   "#1480C8",
  "Emerging Markets": "#27A066",
  "Fixed Income":     "#7056A0",
  "Alternatives":     "#C4872A",
  "Cash":             "#9BAAB8",
};
const SUBGROUP_ORDER = ["Domestic Equity","Intl Developed","Emerging Markets","Fixed Income","Alternatives","Cash"];

// Sub-group tints for the fund table — splits Equity into Domestic / Intl Developed /
// Emerging Markets (matching the donut/stacked-bar breakdown above) instead of lumping
// all equity funds into one bucket. Fixed Income / Alternatives / Cash are unchanged.
const SUBGROUP_TINTS = {
  "Domestic Equity":  { bg: "#EEF2FA", text: "var(--navy)" },
  "Intl Developed":   { bg: "#E7F1FA", text: "#0F6097" },
  "Emerging Markets": { bg: "#E8F6EE", text: "#1C7A4D" },
  "Fixed Income":     { bg: "#EDE8F5", text: "#5040A0" },
  "Alternatives":     { bg: "#FBF3E4", text: "#8A5A10" },
  "Cash":             { bg: "#F2F4F6", text: "#6b7a99" },
};

// Shared grid template for the fund table — 5 columns: name, asset class, alloc,
// fee, and the swap-alternative control.
const FUND_ROW_COLUMNS = "1fr 124px 50px 46px 230px";

// Alternatives (Core Private / Select Liquidity) require $10M+ net worth
const ALT_NET_WORTH_THRESHOLD_M = 10;
const ALT_MODEL_KEYS = new Set(["corePrivate", "selectLiquidity", "selectLiquidityUsBias"]);

export default function ProposalPreviewModal({ data, name, selectedStrategies,
  portfolioModel, riskProfile, portfolioLabel, riskLabel, altFallback,
  backtestResult, backtestLoading, backtestError,
  onConfirm, onBack }) {
  const [local, setLocal]           = useState(() => ({ ...data }));
  const [page, setPage]             = useState(1);
  const hasStrategies = selectedStrategies?.crt || selectedStrategies?.harvesting || selectedStrategies?.collar;
  const TOTAL_PAGES = 3;
  const [priceStatus, setPriceStatus] = useState(
    data.stockPrice > 0 ? null : (data.ticker ? "loading" : null)
  );

  // Per-line fund swaps chosen in the Portfolio Allocation table below — keyed by
  // the fund's stable `name` field. Value is the chosen alternative's manager-list
  // entry ({ name, ticker, fee, passive, ... }), or absent/undefined to keep the
  // original recommended fund.
  // Advisor-selected providers for each strategy
  const HARVESTING_PROVIDERS = [
    { key: "aperio",   label: "Aperio 130:30",    sub: "BlackRock · 0.35% · $1M min",  detail: "Strategy target: ~20–30 long + ~30 short positions. Schwab custodied. Contact: Maureen Sullivan — Maureen.Sullivan@blackrock.com" },
    { key: "invesco",  label: "Invesco 130:30",    sub: "Invesco · 0.30% · $1M min",    detail: "Step 1: reach out to Josh Rogers / Mark Linnecke. Step 2: transition analysis for tax budgeting. Contact: Josh.Rogers@invesco.com" },
    { key: "gateway",  label: "Gateway 130:30",    sub: "Natixis · 0.23% · $1M min",   detail: "~50 yrs quantitative equity/options experience. 0.23% on gross exposure. Contact: Dylan Barlow — DBarlow@gia.com" },
  ];
  const COLLAR_PROVIDERS = [
    { key: "spiderrock", label: "SpiderRock",  sub: "BlackRock · 0.50%–0.85% · $500K min", detail: "Concentrated positions, hedging, put-selling for income, structured note replication. Contact: Adam Butterfield — Adam.Butterfield@blackrock.com" },
    { key: "gateway",    label: "Gateway",     sub: "Natixis · 0.50% · $500K min",          detail: "Collars, covered calls, yield/protection strategies. Contact: Stephen Solaka — SSolaka@gia.com" },
  ];
  const [harvestingProvider, setHarvestingProvider] = useState("aperio");
  const [collarProvider,     setCollarProvider]     = useState("spiderrock");

  const [fundSwaps, setFundSwaps] = useState({});
  function swapFund(fundKey, candidate) {
    setFundSwaps(prev => {
      if (!candidate) {
        if (!(fundKey in prev)) return prev;
        const next = { ...prev };
        delete next[fundKey];
        return next;
      }
      return { ...prev, [fundKey]: candidate };
    });
  }

  // Auto-fetch live stock price on mount if missing
  useEffect(() => {
    const ticker = data.ticker;
    if (!ticker || data.stockPrice > 0) return;

    setPriceStatus("loading");
    fetch(`/api/quote/${encodeURIComponent(ticker)}`)
      .then(async r => {
        if (!r.ok) {
          const body = await r.text().catch(() => "");
          throw new Error(`Quote API ${r.status}: ${body || r.statusText}`);
        }
        return r.json();
      })
      .then(q => {
        const price = Number(q.previousClose || q.currentPrice || 0);
        if (price > 0) {
          setLocal(prev => recomputeReviewedData({ ...prev, stockPrice: price }));
          setPriceStatus("live");
        } else {
          console.warn(`[stock price] No usable price in response for ${ticker}:`, q);
          setPriceStatus("failed");
        }
      })
      .catch(err => {
        // Most common cause: the quote server (server.js on :5174) isn't running.
        // `npm run dev` now starts both it and Vite together — if you started
        // Vite alone, the /api proxy has nothing to forward to.
        console.warn(`[stock price] Fetch failed for ${ticker}:`, err.message || err);
        setPriceStatus("failed");
      });
  }, []); // eslint-disable-line react-hooks/exhaustive-deps

  // Recompute derived values whenever a base field changes
  const handleChange = useCallback((field, value) => {
    setLocal(prev => {
      const updated = { ...prev, [field]: value };
      return recomputeReviewedData(updated);
    });
  }, []);

  const d = local;
  if (!d) return null;

  // Live qualification check — fires if the advisor edits net worth down
  // below $10M inside this modal while an alternatives model is selected.
  const livePctWarning = ALT_MODEL_KEYS.has(portfolioModel) && (d.netWorth || 0) < ALT_NET_WORTH_THRESHOLD_M;

  const taxRateFrac         = (d.totalTaxRate || d.taxRate || 37.1) / 100;
  const crtTaxAvoided       = (d.crtAllocation || 0) * (1 - (d.costBasisPct || 0) / 100) * taxRateFrac;
  const crtDeductionBenefit = (d.charitableDeductionHigh || 0) * taxRateFrac;
  const harvestingBenefit   = d.taxSavings || 0;
  const combinedTaxSavings  = crtTaxAvoided + crtDeductionBenefit + harvestingBenefit;

  const crtAmt    = selectedStrategies?.crt       ? (d.crtAllocation    || 0) : 0;
  const sleeveAmt = selectedStrategies?.harvesting ? (d.harvestingSleeve || 0) : 0;
  const collarAmt = selectedStrategies?.collar     ? (d.collarAllocation || 0) : 0;
  const remStock  = Math.max((d.stockPosition || 0) - crtAmt - sleeveAmt - collarAmt, 0);
  const otherAmt  = Math.max((d.investableAssets || 0) - (d.stockPosition || 0), 0);
  const newInv    = (d.investableAssets || 0) - crtAmt;
  const newConc   = newInv > 0 ? ((remStock + collarAmt) / newInv * 100) : 0;
  const total     = d.investableAssets || 1;

  const afterSegs = [
    ...(crtAmt    ? [{ label: "CRT → Charitable Trust",       val: crtAmt,    color: "var(--gold-light)", text: "#7a5f10" }] : []),
    ...(sleeveAmt ? [{ label: "Harvesting Sleeve (130/30)",    val: sleeveAmt, color: "var(--teal)", text: "#fff"    }] : []),
    ...(collarAmt ? [{ label: "Collared Position (Protected)", val: collarAmt, color: "var(--blue)", text: "#fff"    }] : []),
    ...(remStock  ? [{ label: "Remaining Stock",               val: remStock,  color: "var(--coral)", text: "#fff"    }] : []),
                    { label: "Other Portfolio Assets",          val: otherAmt,  color: "#9BACC8", text: "#fff"    },
  ];

  const ep = { localData: d, onChange: handleChange }; // shorthand for editable row props

  return (
    <div style={{
      position: "fixed", inset: 0, zIndex: 1000,
      background: "rgba(15,23,42,0.55)", backdropFilter: "blur(6px)",
      display: "flex", alignItems: "flex-start", justifyContent: "center",
      overflowY: "auto", padding: "32px 16px",
      animation: "ppmFadeIn 0.18s ease",
    }}>
      <style>{`
        @keyframes ppmFadeIn { from { opacity: 0; } to { opacity: 1; } }
        @keyframes ppmSlideUp { from { opacity: 0; transform: translateY(14px) scale(0.99); } to { opacity: 1; transform: translateY(0) scale(1); } }
        .ppm-btn { transition: transform 0.16s cubic-bezier(0.16,1,0.3,1), box-shadow 0.16s ease, background 0.16s ease, border-color 0.16s ease; }
        .ppm-btn:hover { transform: translateY(-1px); }
        .ppm-btn-primary:hover { background: var(--gold-light) !important; box-shadow: 0 10px 24px rgba(184,137,42,0.32); }
        .ppm-btn-ghost:hover { background: rgba(255,255,255,0.12) !important; border-color: rgba(255,255,255,0.55) !important; }
        .ppm-btn-secondary:hover { background: #fafbfd !important; border-color: #b9c2d6 !important; }
        .ppm-btn-dark:hover { background: var(--navy-2) !important; box-shadow: 0 10px 24px rgba(26,39,68,0.3); }
      `}</style>
      <div style={{
        background: "#F6F7FA", borderRadius: 20, width: "100%", maxWidth: 960,
        boxShadow: "0 28px 80px rgba(15,23,42,0.32)",
        fontFamily: "-apple-system, BlinkMacSystemFont, 'Inter', 'Segoe UI', Roboto, sans-serif",
        animation: "ppmSlideUp 0.32s cubic-bezier(0.16,1,0.3,1)",
        overflow: "hidden",
      }}>

        {/* Header */}
        <div style={{
          background: "var(--navy)", color: "#fff", borderRadius: "20px 20px 0 0",
          padding: "20px 32px",
        }}>
          {/* Top row: name + actions */}
          <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between", marginBottom: 16 }}>
            <div>
              <h2 style={{ margin: 0, fontSize: 21, fontWeight: 600, letterSpacing: "-0.01em" }}>{name}</h2>
              <div style={{ fontSize: 12, color: "#9BACC8", marginTop: 3 }}>
                {[d.ticker, d.clientAge ? `Age ${d.clientAge}` : null].filter(Boolean).join(" · ")}
                <span style={{ marginLeft: 10, color: "rgba(155,172,200,0.6)" }}>· Click any ✏ value to edit</span>
              </div>
            </div>
            <div style={{ display: "flex", gap: 10 }}>
              <button className="ppm-btn ppm-btn-ghost" onClick={onBack} style={{
                background: "transparent", border: "1.5px solid rgba(255,255,255,0.35)",
                color: "#fff", borderRadius: 10, padding: "9px 18px", cursor: "pointer", fontSize: 13,
              }}>← Edit Panel</button>
              <button className="ppm-btn ppm-btn-primary" onClick={() => onConfirm(Object.keys(fundSwaps).length ? { ...local, fundSwaps } : local)} style={{
                background: "var(--gold-light)", border: "none", color: "var(--navy)",
                borderRadius: 10, padding: "9px 22px", cursor: "pointer", fontSize: 13, fontWeight: 700,
              }}>Generate Proposal →</button>
            </div>
          </div>

          {/* Page tabs */}
          <div style={{ display: "flex", gap: 4 }}>
            {[
              { n: 1, label: "Client Profile" },
              { n: 2, label: "Strategies" },
              { n: 3, label: "Portfolio" },
            ].map(({ n, label }) => (
              <button
                key={n}
                onClick={() => setPage(n)}
                style={{
                  background: page === n ? "rgba(255,255,255,0.15)" : "transparent",
                  border: page === n ? "1.5px solid rgba(255,255,255,0.4)" : "1.5px solid rgba(255,255,255,0.12)",
                  color: page === n ? "#fff" : "rgba(255,255,255,0.5)",
                  borderRadius: 8, padding: "6px 18px", cursor: "pointer",
                  fontSize: 12, fontWeight: page === n ? 700 : 400,
                  transition: "all 0.15s",
                }}
              >
                <span style={{
                  display: "inline-flex", alignItems: "center", justifyContent: "center",
                  width: 16, height: 16, borderRadius: "50%",
                  background: page === n ? "var(--gold-light)" : "rgba(255,255,255,0.15)",
                  color: page === n ? "var(--navy)" : "rgba(255,255,255,0.5)",
                  fontSize: 9, fontWeight: 800, marginRight: 7,
                }}>{n}</span>
                {label}
              </button>
            ))}
          </div>
        </div>

        {/* Body */}
        <div style={{ padding: "28px 32px 36px" }}>

          {/* Qualification gate banners */}
          {altFallback && (
            <div style={{
              background: "#FEF3DA", border: "1.5px solid var(--gold-light)", borderRadius: 8,
              padding: "10px 16px", marginBottom: 18, fontSize: 12.5, color: "#7a5f10",
            }}>
              <strong>Portfolio model switched: {altFallback.from} → {altFallback.to}.</strong>{" "}
              Alternatives require $10M+ net worth; this client's net worth was below that threshold.
            </div>
          )}
          {!altFallback && livePctWarning && (
            <div style={{
              background: "#FBEAEA", border: "1.5px solid #E0A0A0", borderRadius: 8,
              padding: "10px 16px", marginBottom: 18, fontSize: 12.5, color: "#8a2a2a",
            }}>
              <strong>Net worth is below $10M.</strong>{" "}
              {portfolioLabel || "This model"} requires alternatives access — generating now will switch the model to Traditional.
            </div>
          )}

          {/* ═══ PAGE 1: Client Profile ═══ */}
          {page === 1 && <>

          {/* 1. Portfolio & Tax Profile */}
          <Section title="Portfolio & Tax Profile">
            <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 14 }}>
              <Card title="Wealth Summary" accent="var(--navy)">
                <Row label="Net Worth ($M)"            value={fmtM(d.netWorth)}          field="netWorth"        highlight="var(--navy)" {...ep} />
                <Row label="Investable Assets ($M)"    value={fmtM(d.investableAssets)}   field="investableAssets" highlight="var(--navy)" {...ep} />
                <Row label="Annual Income ($M)"        value={d.income ? fmtM(d.income) : "—"} field="income"   highlight="var(--navy)" {...ep} />
                <Row label="Concentrated Position ($M)" value={fmtM(d.stockPosition)}    field="stockPosition"   highlight="var(--coral)" {...ep} />
                <Row label="Concentration %"           value={`${(d.concentration||0).toFixed(1)}%`} sub={d.ticker} />
              </Card>
              <Card title="Tax Rates & Basis" accent="var(--gold)">
                <Row label="Federal LTCG Rate (%)"  value={`${d.federalTaxRate || 23.8}%`}         field="federalTaxRate" highlight="var(--gold)" {...ep} />
                <Row label="State Tax Rate (%)"     value={`${(d.stateTaxRate||0).toFixed(1)}%`}    field="stateTaxRate"   highlight="var(--gold)" {...ep} />
                <Row label="Combined Rate (%)"      value={`${d.totalTaxRate || d.taxRate || 0}%`}  highlight="var(--coral)" />
                <Row label="Cost Basis (%)"         value={`${d.costBasisPct || 0}%`}              field="costBasisPct"   highlight="var(--gold)" {...ep} />
                <Row label="Cost Basis ($M)"        value={fmtM(d.costBasis)} />
                <Row label="Embedded Gain ($M)"     value={fmtM(d.embeddedGain)}                   highlight="#E67E22" />
                <Row label="Immediate Tax Liability" value={fmtM(d.immediateTax)}                  highlight="var(--coral)" />
                <Row label="40% Drawdown Exposure"  value={fmtM(d.drawdown40Impact)} />
              </Card>
            </div>
          </Section>

          {/* 2. Concentration visuals */}
          <Section title="Concentration — Before vs. After">
            <div style={{ display: "flex", gap: 24, alignItems: "center", justifyContent: "center", flexWrap: "wrap" }}>
              <DonutChart value={d.stockPosition||0} total={d.investableAssets||1}
                color="var(--coral)" label="Before" sub="Current" />
              <div style={{ fontSize: 20, color: "var(--gold)", fontWeight: 700 }}>→</div>
              {selectedStrategies?.crt && (
                <DonutChart value={(d.stockPosition||0)-crtAmt} total={(d.investableAssets||1)-crtAmt}
                  color="var(--gold)" label="After CRT" sub={`${(d.afterCrtConcentration||0).toFixed(1)}%`} />
              )}
              {(selectedStrategies?.harvesting || selectedStrategies?.collar) && (
                <DonutChart value={remStock+collarAmt} total={newInv||1}
                  color="var(--blue)" label="Final (all strategies)" sub={`${newConc.toFixed(1)}%`} />
              )}
            </div>
          </Section>

          </>}

          {/* ═══ PAGE 2: Strategies ═══ */}
          {page === 2 && <>

          {/* 3. CRT */}
          {selectedStrategies?.crt && (
            <Section title="Charitable Remainder Trust">
              <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 14 }}>
                <Card title="Inputs" accent="var(--navy)">
                  <Row label="Client Age"             value={d.clientAge ? `${d.clientAge} yrs` : "—"} field="clientAge"   highlight="var(--navy)" {...ep} />
                  <AllocationRow label="CRT Contribution" pctField="crtPct" baseAmount={d.stockPosition} minDollarM={1} localData={d} onChange={handleChange} accent="var(--navy)" note="min $1M · click $ or % to edit" />
                  <AllocationRow label="Payout Rate" pctField="crtPayoutRate" baseAmount={d.crtAllocation} pctMin={5} pctMax={50} localData={d} onChange={handleChange} accent="var(--teal)" note="5%–50% per IRC §664 · click $ or % to edit" />
                  <Row label="IRS Deduction Factor"   value={`${d.crtAllocation ? ((d.charitableDeductionHigh/d.crtAllocation)*100).toFixed(1) : "—"}%`} sub="from IRS Table S by age" />
                </Card>
                <Card title="Outputs" accent="var(--teal)">
                  <Row label="Annual CRT Income"         value={fmtK(d.crtIncome)}             highlight="var(--teal)" sub="/year" />
                  <Row label="Charitable Tax Deduction"  value={fmtM(d.charitableDeductionHigh)} highlight="var(--teal)" />
                  <Row label="Deduction Tax Benefit"     value={fmtM(crtDeductionBenefit)}      sub={`deduction × ${(d.totalTaxRate||37.1).toFixed(1)}%`} />
                  <Row label="Capital Gains Tax Avoided" value={fmtM(crtTaxAvoided)}            highlight="var(--teal)" />
                  <Row label="Concentration After CRT"   value={`${(d.afterCrtConcentration||0).toFixed(1)}%`} sub="of reduced portfolio" />
                </Card>
              </div>

              {/* CRT Investment Portfolio — what the trust holds after contribution */}
              {portfolioModel && riskProfile && (() => {
                const crtFunds = getAllFunds(portfolioModel, riskProfile).filter(f => f.alloc > 0);
                if (crtFunds.length === 0) return null;

                // Group by broad category for the summary bars
                const groups = {};
                for (const f of crtFunds) {
                  groups[f.group] = (groups[f.group] || 0) + f.alloc;
                }
                const groupColors = {
                  Equity: "var(--navy)",
                  "Fixed Income": "var(--teal)",
                  Alternatives: "var(--gold)",
                  Cash: "#9BACC8",
                };

                return (
                  <Card title="CRT Investment Portfolio" accent="var(--gold)" style={{ marginTop: 14 }}>
                    {/* Explanation */}
                    <div style={{ fontSize: 11, color: "#6b7a99", marginBottom: 14, lineHeight: 1.6, padding: "10px 12px", background: "#FFFBF0", border: "1px solid #F5DFA0", borderRadius: 8 }}>
                      <strong style={{ color: "var(--navy)" }}>How it works:</strong> Upon contribution, the {d.ticker || "concentrated stock"} shares are transferred into the CRT. The trustee then sells them <em>tax-free inside the trust</em> — no capital gains recognized at sale. The full proceeds (~{fmtM(d.crtAllocation)}) are immediately reinvested in the diversified portfolio below, generating income paid to the client for the trust term. At termination, the remainder passes to the designated charity.
                    </div>

                    {/* Asset-class summary bar */}
                    <div style={{ marginBottom: 14 }}>
                      <div style={{ fontSize: 10, color: "#9bacc8", fontWeight: 700, textTransform: "uppercase", letterSpacing: "0.05em", marginBottom: 6 }}>
                        Asset Allocation
                      </div>
                      <div style={{ display: "flex", borderRadius: 6, overflow: "hidden", height: 22 }}>
                        {Object.entries(groups).filter(([, v]) => v > 0).map(([g, v]) => (
                          <div key={g} style={{
                            width: `${v}%`, background: groupColors[g] || "#9BACC8",
                            display: "flex", alignItems: "center", justifyContent: "center",
                            fontSize: 9, fontWeight: 700, color: "#fff", whiteSpace: "nowrap", overflow: "hidden",
                          }}>
                            {v >= 8 ? `${g} ${v.toFixed(0)}%` : v >= 4 ? `${v.toFixed(0)}%` : ""}
                          </div>
                        ))}
                      </div>
                      <div style={{ display: "flex", flexWrap: "wrap", gap: "6px 14px", marginTop: 6 }}>
                        {Object.entries(groups).filter(([, v]) => v > 0).map(([g, v]) => (
                          <div key={g} style={{ display: "flex", alignItems: "center", gap: 5, fontSize: 10, color: "#6b7a99" }}>
                            <div style={{ width: 8, height: 8, borderRadius: 2, background: groupColors[g] || "#9BACC8", flexShrink: 0 }} />
                            {g} — {v.toFixed(1)}%
                          </div>
                        ))}
                      </div>
                    </div>

                    {/* Fund-level table */}
                    <div style={{ fontSize: 10, color: "#9bacc8", fontWeight: 700, textTransform: "uppercase", letterSpacing: "0.05em", marginBottom: 6 }}>
                      Holdings
                    </div>
                    <div style={{ border: "1px solid #EEF0F5", borderRadius: 8, overflow: "hidden" }}>
                      {/* Header */}
                      <div style={{ display: "grid", gridTemplateColumns: "1fr auto auto", gap: 8, padding: "7px 12px", background: "#F5F7FB", borderBottom: "1px solid #EEF0F5" }}>
                        <span style={{ fontSize: 10, fontWeight: 700, color: "#9bacc8", textTransform: "uppercase" }}>Fund / Manager</span>
                        <span style={{ fontSize: 10, fontWeight: 700, color: "#9bacc8", textTransform: "uppercase", textAlign: "right" }}>Alloc</span>
                        <span style={{ fontSize: 10, fontWeight: 700, color: "#9bacc8", textTransform: "uppercase", textAlign: "right", minWidth: 70 }}>Est. Value</span>
                      </div>
                      {/* Rows grouped by category */}
                      {["Equity", "Fixed Income", "Alternatives", "Cash"].map(grp => {
                        const rows = crtFunds.filter(f => f.group === grp);
                        if (rows.length === 0) return null;
                        return (
                          <div key={grp}>
                            <div style={{ padding: "5px 12px", background: "#FAFBFD", borderBottom: "1px solid #EEF0F5" }}>
                              <span style={{ fontSize: 10, fontWeight: 700, color: groupColors[grp] || "#9bacc8", textTransform: "uppercase", letterSpacing: "0.04em" }}>
                                {grp}
                              </span>
                            </div>
                            {rows.map((f, i) => {
                              const estVal = (d.crtAllocation || 0) * (f.alloc / 100);
                              return (
                                <div key={i} style={{
                                  display: "grid", gridTemplateColumns: "1fr auto auto", gap: 8,
                                  padding: "7px 12px", borderBottom: "1px solid #EEF0F5",
                                  background: i % 2 === 0 ? "#fff" : "#FAFBFD",
                                }}>
                                  <div>
                                    <div style={{ fontSize: 12, fontWeight: 600, color: "var(--navy)" }}>{f.name}</div>
                                    <div style={{ fontSize: 10, color: "#9bacc8" }}>{f.assetClass}</div>
                                  </div>
                                  <div style={{ textAlign: "right", fontSize: 13, fontWeight: 700, color: "var(--navy)" }}>
                                    {f.alloc.toFixed(1)}%
                                  </div>
                                  <div style={{ textAlign: "right", fontSize: 12, fontWeight: 600, color: "var(--teal)", minWidth: 70 }}>
                                    {fmtM(estVal)}
                                  </div>
                                </div>
                              );
                            })}
                          </div>
                        );
                      })}
                      {/* Total */}
                      <div style={{ display: "grid", gridTemplateColumns: "1fr auto auto", gap: 8, padding: "8px 12px", background: "#F0F4FF", borderTop: "1px solid #DDE3F5" }}>
                        <span style={{ fontSize: 12, fontWeight: 700, color: "var(--navy)" }}>Total CRT Assets</span>
                        <span style={{ fontSize: 13, fontWeight: 700, color: "var(--navy)", textAlign: "right" }}>100%</span>
                        <span style={{ fontSize: 12, fontWeight: 700, color: "var(--teal)", textAlign: "right", minWidth: 70 }}>{fmtM(d.crtAllocation || 0)}</span>
                      </div>
                    </div>
                    <div style={{ fontSize: 10, color: "#9bacc8", marginTop: 8 }}>
                      Allocation mirrors the client's recommended {portfolioLabel || "target"} portfolio at {riskLabel || "the selected"} risk profile. Exact holdings subject to trustee discretion and available investment minimums.
                    </div>
                  </Card>
                );
              })()}
            </Section>
          )}

          {/* 4. Harvesting */}
          {selectedStrategies?.harvesting && (
            <Section title="Leveraged Tax-Loss Harvesting (130/30)">
              <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 14, marginBottom: 14 }}>
                <Card title="Structure" accent="var(--gold)">
                  <AllocationRow label="Sleeve Size" pctField="harvestingPct" baseAmount={d.stockPosition} localData={d} onChange={handleChange} accent="var(--gold)" note="click $ or % to edit" />
                  <Row label="Annual Harvest Losses ($M)" value={fmtM(d.annualHarvestLosses)}   highlight="var(--gold)" sub="25% of sleeve" />
                </Card>
                <Card title="Annual Tax Savings" accent="var(--teal)">
                  <Row label="Federal Rate"         value={`${d.federalTaxRate || 23.8}%`} />
                  <Row label="Federal Tax Savings"  value={fmtK(d.federalTaxSavings)}   highlight="var(--teal)" />
                  <Row label="State Rate"           value={`${(d.stateTaxRate||0).toFixed(1)}%`} />
                  <Row label="State Tax Savings"    value={fmtK(d.stateTaxSavings)} />
                  <Row label="Total Annual Savings" value={fmtK(d.taxSavings)}          highlight="var(--teal)" />
                </Card>
              </div>
              {/* Provider selector */}
              <div>
                <div style={{ fontSize: 10, fontWeight: 700, color: "#9bacc8", textTransform: "uppercase", letterSpacing: "0.06em", marginBottom: 8 }}>
                  Select Provider
                </div>
                <div style={{ display: "flex", gap: 8, flexWrap: "wrap" }}>
                  {HARVESTING_PROVIDERS.map(p => (
                    <button key={p.key} onClick={() => setHarvestingProvider(p.key)} style={{
                      padding: "7px 14px", borderRadius: 8, cursor: "pointer", fontSize: 12, fontWeight: 600,
                      border: harvestingProvider === p.key ? "2px solid var(--gold)" : "1.5px solid #E0E5EF",
                      background: harvestingProvider === p.key ? "#FEF9EE" : "#fff",
                      color: harvestingProvider === p.key ? "var(--gold)" : "#6b7a99",
                      transition: "all 0.15s",
                    }}>
                      {p.label}
                    </button>
                  ))}
                </div>
                {(() => {
                  const p = HARVESTING_PROVIDERS.find(p => p.key === harvestingProvider);
                  return p ? (
                    <div style={{ marginTop: 8, padding: "10px 14px", background: "#FFFBF0", border: "1px solid #F5DFA0", borderRadius: 8 }}>
                      <div style={{ fontSize: 12, fontWeight: 700, color: "var(--navy)", marginBottom: 3 }}>{p.label} <span style={{ fontWeight: 400, color: "#9bacc8" }}>— {p.sub}</span></div>
                      <div style={{ fontSize: 11, color: "#6b7a99" }}>{p.detail}</div>
                    </div>
                  ) : null;
                })()}
              </div>
            </Section>
          )}

          {/* 5. Collar */}
          {selectedStrategies?.collar && (
            <Section title="Option Collar">
              <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 14, marginBottom: 14 }}>
                <Card title="Structure" accent="var(--blue)">
                  <AllocationRow label="Collar Size" pctField="collarPct" baseAmount={d.stockPosition} localData={d} onChange={handleChange} accent="var(--blue)" note="click $ or % to edit" />
                  <Row
                    label={
                      priceStatus === "loading" ? "Stock Price ($)  ⏳ fetching…" :
                      priceStatus === "live"    ? `Stock Price ($)  ✓ live` :
                      priceStatus === "failed"  ? "Stock Price ($)  ⚠ enter manually" :
                      "Stock Price ($)"
                    }
                    value={d.stockPrice ? `$${Number(d.stockPrice).toFixed(2)}` : "—"}
                    field="stockPrice" highlight="var(--blue)" placeholder="e.g. 185" {...ep}
                  />
                  <Row label="Put Strike (−15%)"     value={d.putStrike  ? `$${d.putStrike.toFixed(2)}`  : "—"} />
                  <Row label="Call Strike (+19%)"    value={d.callStrike ? `$${d.callStrike.toFixed(2)}` : "—"} />
                </Card>
                <Card title="Dollar Range" accent="var(--blue)">
                  <Row label="Downside Floor (put)" value={fmtM(d.putFloorValue)} highlight="var(--coral)" sub="max loss on this sleeve" />
                  <Row label="Upside Cap (call)"    value={fmtM(d.callCapValue)}  highlight="var(--teal)" sub="max gain on this sleeve" />
                  <Row label="Protected Range"      value={`${fmtM(d.putFloorValue)} – ${fmtM(d.callCapValue)}`} />
                </Card>
              </div>
              {/* Provider selector */}
              <div>
                <div style={{ fontSize: 10, fontWeight: 700, color: "#9bacc8", textTransform: "uppercase", letterSpacing: "0.06em", marginBottom: 8 }}>
                  Select Provider
                </div>
                <div style={{ display: "flex", gap: 8, flexWrap: "wrap" }}>
                  {COLLAR_PROVIDERS.map(p => (
                    <button key={p.key} onClick={() => setCollarProvider(p.key)} style={{
                      padding: "7px 14px", borderRadius: 8, cursor: "pointer", fontSize: 12, fontWeight: 600,
                      border: collarProvider === p.key ? "2px solid var(--blue)" : "1.5px solid #E0E5EF",
                      background: collarProvider === p.key ? "#EEF4FF" : "#fff",
                      color: collarProvider === p.key ? "var(--blue)" : "#6b7a99",
                      transition: "all 0.15s",
                    }}>
                      {p.label}
                    </button>
                  ))}
                </div>
                {(() => {
                  const p = COLLAR_PROVIDERS.find(p => p.key === collarProvider);
                  return p ? (
                    <div style={{ marginTop: 8, padding: "10px 14px", background: "#EEF4FF", border: "1px solid #C0D4F5", borderRadius: 8 }}>
                      <div style={{ fontSize: 12, fontWeight: 700, color: "var(--navy)", marginBottom: 3 }}>{p.label} <span style={{ fontWeight: 400, color: "#9bacc8" }}>— {p.sub}</span></div>
                      <div style={{ fontSize: 11, color: "#6b7a99" }}>{p.detail}</div>
                    </div>
                  ) : null;
                })()}
              </div>
            </Section>
          )}

          {/* 6. Combined Tax Impact */}
          <Section title="Combined Tax Impact">
            <Card accent="var(--teal)">
              {selectedStrategies?.crt && (
                <Row label="CRT — Capital Gains Tax Avoided" value={fmtM(crtTaxAvoided)}       highlight="var(--teal)" />
              )}
              {selectedStrategies?.crt && (
                <Row label="CRT — Deduction Tax Benefit"     value={fmtM(crtDeductionBenefit)} highlight="var(--teal)" />
              )}
              {selectedStrategies?.harvesting && (
                <Row label="Harvesting — Annual Tax Savings" value={fmtK(harvestingBenefit)}   highlight="var(--gold)" sub="recurring each year" />
              )}
            </Card>
            <div style={{
              marginTop: 10, padding: "16px 20px", background: "linear-gradient(135deg, var(--navy), #243459)", borderRadius: 14,
              display: "flex", justifyContent: "space-between", alignItems: "center",
              boxShadow: "0 10px 24px rgba(26,39,68,0.25)",
            }}>
              <span style={{ fontSize: 13, color: "#9BACC8" }}>Combined Tax Benefit (All Strategies)</span>
              <span style={{ fontSize: 23, fontWeight: 700, color: "var(--gold-light)" }}>{fmtM(combinedTaxSavings)}</span>
            </div>
          </Section>

          {/* 7. Portfolio After */}
          {d.stockPosition > 0 && (selectedStrategies?.crt || selectedStrategies?.harvesting || selectedStrategies?.collar) && (
            <Section title="Portfolio After Strategy Implementation">
              <div style={{ marginBottom: 6 }}>
                <div style={{ fontSize: 10, color: "#6b7a99", marginBottom: 3 }}>BEFORE — {fmtM(d.investableAssets)}</div>
                <div style={{ display: "flex", borderRadius: 5, overflow: "hidden", height: 28 }}>
                  <div style={{ width:`${(d.stockPosition||0)/total*100}%`, background:"var(--coral)", display:"flex", alignItems:"center", justifyContent:"center", fontSize:10, color:"#fff", fontWeight:700, padding:"0 4px", whiteSpace:"nowrap", overflow:"hidden" }}>
                    {fmtM(d.stockPosition)} stock
                  </div>
                  <div style={{ flex:1, background:"#9BACC8", display:"flex", alignItems:"center", justifyContent:"center", fontSize:10, color:"#fff", fontWeight:700 }}>
                    {fmtM(otherAmt)} other
                  </div>
                </div>
              </div>
              <div style={{ marginBottom: 14 }}>
                <div style={{ fontSize: 10, color: "var(--teal)", marginBottom: 3 }}>
                  AFTER — {fmtM(newInv)} invested{crtAmt ? ` + ${fmtM(crtAmt)} to CRT` : ""}
                </div>
                <div style={{ display: "flex", borderRadius: 5, overflow: "hidden", height: 28 }}>
                  {afterSegs.map((s, i) => (
                    <div key={i} style={{
                      width:`${s.val/total*100}%`, background:s.color,
                      display:"flex", alignItems:"center", justifyContent:"center",
                      fontSize:10, color:s.text, fontWeight:700, padding:"0 4px",
                      whiteSpace:"nowrap", overflow:"hidden",
                    }}>
                      {s.val/total > 0.06 ? fmtM(s.val) : ""}
                    </div>
                  ))}
                </div>
              </div>
              <Card accent="#E8ECF0">
                {afterSegs.map((s, i) => (
                  <Row key={i} label={s.label} value={fmtM(s.val)}
                    sub={`${(s.val/total*100).toFixed(1)}% of original portfolio`}
                    highlight={s.color} />
                ))}
                <div style={{ borderTop: "1.5px solid #E8ECF0", marginTop: 6, paddingTop: 8 }}>
                  <Row label="New Concentration (remaining + collared)"
                    value={`${newConc.toFixed(1)}%`}
                    sub={`was ${(d.concentration||0).toFixed(1)}%`}
                    highlight="var(--teal)" />
                </div>
              </Card>
            </Section>
          )}

          </>}

          {/* ═══ PAGE 3: Portfolio ═══ */}
          {page === 3 && <>

          {/* 8. Portfolio Allocation */}
          {portfolioModel && riskProfile && (() => {
            const funds   = getAllFunds(portfolioModel, riskProfile);
            const sgt     = getSubGroupTotals(portfolioModel, riskProfile);
            const sgTotal = SUBGROUP_ORDER.reduce((s, k) => s + (sgt[k] || 0), 0) || 100;

            // Stacked bar segments
            const barSegs = SUBGROUP_ORDER
              .map(k => ({ label: k, pct: sgt[k] || 0, color: SUBGROUP_COLORS[k] }))
              .filter(s => s.pct > 0.1);

            // Current-portfolio comparison bar — reuses the exact same
            // concentrated-position + benchmark approximation already disclosed
            // in the Portfolio Transition Analysis backtest below (no separate
            // or fabricated breakdown, just visualized alongside the target).
            const benchmarkTicker = backtestResult?.benchmarkTicker;
            const currentConc = Number(d.concentration) || 0;
            const hasCurrentAlloc = !!d.ticker && currentConc > 0;
            const currentSegs = hasCurrentAlloc
              ? [
                  { label: d.ticker, pct: currentConc, color: "var(--coral)" },
                  ...(benchmarkTicker && currentConc < 100
                    ? [{ label: benchmarkTicker, pct: 100 - currentConc, color: "#9BACC8" }]
                    : []),
                ]
              : [];

            // Fund table grouped by sub-group (Domestic Equity / Intl Developed /
            // Emerging Markets / Fixed Income / Alternatives / Cash) so the equity
            // sleeve is broken out the same way as the donut + stacked bar above.
            const groups = SUBGROUP_ORDER;
            const grouped = {};
            for (const g of groups) grouped[g] = funds.filter(f => f.subGroup === g);

            return (
              <Section title="Portfolio Allocation">
                {/* Model + risk badges */}
                <div style={{ display: "flex", gap: 10, marginBottom: 14, flexWrap: "wrap" }}>
                  {portfolioLabel && (
                    <div style={{
                      background: "var(--navy)", color: "#fff", borderRadius: 6,
                      padding: "5px 14px", fontSize: 12, fontWeight: 700,
                    }}>{portfolioLabel}</div>
                  )}
                  {riskLabel && (
                    <div style={{
                      background: "#F5EDD6", color: "var(--gold)", borderRadius: 6,
                      padding: "5px 14px", fontSize: 12, fontWeight: 700,
                      border: "1.5px solid var(--gold-light)",
                    }}>{riskLabel}</div>
                  )}
                  <div style={{ fontSize: 12, color: "#6b7a99", alignSelf: "center" }}>
                    {funds.filter(f => f.alloc > 0).length} active · {funds.length} total
                  </div>
                </div>

                {/* Horizontal stacked bar — Target */}
                <div style={{ marginBottom: 6 }}>
                  <div style={{ fontSize: 10, fontWeight: 700, color: "#9bacc8", textTransform: "uppercase", letterSpacing: "0.05em", marginBottom: 4 }}>
                    Target Allocation
                  </div>
                  <div style={{ display: "flex", borderRadius: 6, overflow: "hidden", height: 22 }}>
                    {barSegs.map((s, i) => (
                      <div key={i} style={{
                        width: `${(s.pct / sgTotal) * 100}%`,
                        background: s.color,
                        transition: "width 0.3s",
                      }} title={`${s.label}: ${s.pct.toFixed(1)}%`} />
                    ))}
                  </div>
                  {/* Legend */}
                  <div style={{ display: "flex", flexWrap: "wrap", gap: "4px 14px", marginTop: 8 }}>
                    {barSegs.map((s, i) => (
                      <div key={i} style={{ display: "flex", alignItems: "center", gap: 5 }}>
                        <div style={{ width: 10, height: 10, borderRadius: 2, background: s.color, flexShrink: 0 }} />
                        <span style={{ fontSize: 11, color: "#6b7a99" }}>
                          {s.label} <strong style={{ color: "var(--navy)" }}>{s.pct.toFixed(1)}%</strong>
                        </span>
                      </div>
                    ))}
                  </div>
                </div>

                {/* Horizontal stacked bar — Current (approximate) */}
                {hasCurrentAlloc && (
                  <div style={{ marginBottom: 6, marginTop: 16 }}>
                    <div style={{ fontSize: 10, fontWeight: 700, color: "#9bacc8", textTransform: "uppercase", letterSpacing: "0.05em", marginBottom: 4 }}>
                      Current Allocation (approximate)
                    </div>
                    <div style={{ display: "flex", borderRadius: 6, overflow: "hidden", height: 22 }}>
                      {currentSegs.map((s, i) => (
                        <div key={i} style={{
                          width: `${s.pct}%`,
                          background: s.color,
                          transition: "width 0.3s",
                        }} title={`${s.label}: ${s.pct.toFixed(1)}%`} />
                      ))}
                    </div>
                    {/* Legend */}
                    <div style={{ display: "flex", flexWrap: "wrap", gap: "4px 14px", marginTop: 8 }}>
                      {currentSegs.map((s, i) => (
                        <div key={i} style={{ display: "flex", alignItems: "center", gap: 5 }}>
                          <div style={{ width: 10, height: 10, borderRadius: 2, background: s.color, flexShrink: 0 }} />
                          <span style={{ fontSize: 11, color: "#6b7a99" }}>
                            {s.label} <strong style={{ color: "var(--navy)" }}>{s.pct.toFixed(1)}%</strong>
                          </span>
                        </div>
                      ))}
                    </div>
                    <div style={{ fontSize: 9, color: "#9bacc8", marginTop: 6, fontStyle: "italic" }}>
                      Based on the concentrated position noted for this client
                      {benchmarkTicker ? ` — remainder approximated as ${benchmarkTicker}, a diversified benchmark` : ""}.
                      Not a full reconstruction of actual current holdings.
                    </div>
                  </div>
                )}

                {/* Fund table */}
                <div style={{ marginTop: 14 }}>
                  {/* Header */}
                  <div style={{
                    display: "grid", gridTemplateColumns: FUND_ROW_COLUMNS,
                    padding: "5px 8px", borderBottom: "2px solid var(--navy)",
                    fontSize: 10, fontWeight: 700, color: "var(--navy)", textTransform: "uppercase", letterSpacing: "0.06em",
                  }}>
                    <span>Fund / Manager</span>
                    <span>Asset Class</span>
                    <span style={{ textAlign: "right" }}>Alloc</span>
                    <span style={{ textAlign: "right" }}>Fee</span>
                    <span style={{ paddingLeft: 8 }}>Swap Alternative</span>
                  </div>

                  {groups.map(group => {
                    const gFunds = grouped[group];
                    if (!gFunds || gFunds.length === 0) return null;
                    const gc = SUBGROUP_TINTS[group] || { bg: "#F8F9FC", text: "var(--navy)" };
                    const gTotal = gFunds.reduce((s, f) => s + f.alloc, 0);
                    return (
                      <div key={group}>
                        {/* Section header */}
                        <div style={{
                          display: "grid", gridTemplateColumns: FUND_ROW_COLUMNS,
                          padding: "5px 8px", background: gc.bg,
                          fontSize: 10, fontWeight: 700, color: gc.text, textTransform: "uppercase",
                        }}>
                          <span>{group}</span>
                          <span />
                          <span style={{ textAlign: "right" }}>{gTotal.toFixed(1)}%</span>
                          <span />
                          <span />
                        </div>
                        {/* Fund rows */}
                        {gFunds.map((f, i) => {
                          const inactive = f.alloc === 0;
                          const dollarAllocated = (d.investableAssets || 0) * 1000000 * (f.alloc / 100);
                          const swap = fundSwaps[f.name] || null;
                          const displayName = swap ? swap.name : f.fullName;
                          const displayFee  = swap ? swap.fee  : f.fee;

                          const alternatives = !inactive ? getFundAlternatives(f, dollarAllocated || Infinity) : [];
                          const options = [
                            { optKey: "__original__", label: f.fullName, fee: f.fee, candidate: null },
                            ...alternatives.map(a => ({ optKey: (a.ticker || a.name), label: a.name, fee: a.fee, candidate: a })),
                          ];
                          const selectedKey = swap ? (swap.ticker || swap.name) : "__original__";

                          const feeDeltaBps    = swap && Number.isFinite(f.fee) && Number.isFinite(swap.fee) ? (f.fee - swap.fee) * 100 : 0;
                          const feeDeltaDollar = swap && dollarAllocated ? dollarAllocated * (f.fee - swap.fee) / 100 : 0;
                          const hasSwapChoices = !inactive && options.length > 1;

                          return (
                            <div key={i} style={{
                              display: "grid", gridTemplateColumns: FUND_ROW_COLUMNS,
                              padding: "5px 8px",
                              background: inactive ? "#F8F8FA" : (i % 2 === 0 ? "#FAFBFD" : "#ffffff"),
                              borderBottom: "1px solid #F0F2F6",
                              fontSize: 11,
                              opacity: inactive ? 0.5 : 1,
                            }}>
                              <div style={{ overflow: "hidden" }}>
                                <div style={{
                                  color: inactive ? "#9bacc8" : "var(--navy)",
                                  fontWeight: inactive ? 400 : 700,
                                  overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap",
                                  fontSize: 11,
                                }}>
                                  {displayName}
                                  {swap && (
                                    <span style={{
                                      marginLeft: 6, fontSize: 8, fontWeight: 700, color: "#1C7A4D",
                                      background: "#E8F6EE", borderRadius: 4, padding: "1px 5px",
                                      textTransform: "uppercase", letterSpacing: "0.03em",
                                    }}>swapped</span>
                                  )}
                                </div>
                                {swap && (
                                  <div style={{ fontSize: 9, color: "#9bacc8", overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap" }}>
                                    was {f.fullName}
                                  </div>
                                )}
                              </div>
                              <span style={{ color: "#6b7a99", fontSize: 10, alignSelf: "center" }}>{f.assetClass}</span>
                              <span style={{
                                textAlign: "right", alignSelf: "center",
                                fontWeight: inactive ? 400 : 700,
                                color: inactive ? "#9bacc8" : "var(--navy)",
                              }}>
                                {inactive ? "—" : `${f.alloc.toFixed(1)}%`}
                              </span>
                              <span style={{ textAlign: "right", color: "#9bacc8", fontSize: 10, alignSelf: "center" }}>
                                {displayFee > 0 ? `${displayFee.toFixed(2)}%` : "—"}
                              </span>
                              <div style={{ alignSelf: "center", paddingLeft: 8 }}>
                                {hasSwapChoices ? (
                                  <>
                                    <select
                                      value={selectedKey}
                                      onChange={e => {
                                        const chosen = options.find(o => o.optKey === e.target.value);
                                        swapFund(f.name, chosen ? chosen.candidate : null);
                                      }}
                                      style={{
                                        width: "100%", fontSize: 10, color: "var(--navy)",
                                        border: "1px solid #D8DEE8", borderRadius: 5,
                                        padding: "3px 4px", background: "#fff",
                                      }}
                                    >
                                      {options.map(o => (
                                        <option key={o.optKey} value={o.optKey}>
                                          {o.optKey === "__original__"
                                            ? `Keep current (${o.fee > 0 ? o.fee.toFixed(2) + "%" : "—"})`
                                            : `${o.candidate.passive ? "Passive" : "Active"}: ${o.label} (${o.fee.toFixed(2)}%)`}
                                        </option>
                                      ))}
                                    </select>
                                    {swap && (
                                      <div style={{
                                        fontSize: 9, marginTop: 2,
                                        color: feeDeltaBps >= 0 ? "#1C7A4D" : "#B23A3A",
                                      }}>
                                        {feeDeltaBps >= 0 ? "▼" : "▲"} {Math.abs(feeDeltaBps).toFixed(0)}bps
                                        {dollarAllocated > 0 ? ` · ${fmtDollar(Math.abs(feeDeltaDollar))}/yr ${feeDeltaBps >= 0 ? "saved" : "more"}` : ""}
                                      </div>
                                    )}
                                  </>
                                ) : (
                                  <span style={{ fontSize: 10, color: "#c3cbdb" }}>—</span>
                                )}
                              </div>
                            </div>
                          );
                        })}
                      </div>
                    );
                  })}
                </div>
              </Section>
            );
          })()}

          {/* 9. Portfolio Transition Analysis */}
          {portfolioModel && riskProfile && (
            <Section title="Portfolio Transition Analysis">
              {backtestLoading && (
                <div style={{ fontSize: 12, color: "#6b7a99", padding: "14px 0" }}>
                  Running historical backtest on real market data…
                </div>
              )}

              {!backtestLoading && backtestError && (
                <div style={{ fontSize: 12, color: "#B23A3A", padding: "14px 0" }}>
                  Backtest unavailable: {backtestError}
                </div>
              )}

              {!backtestLoading && !backtestError && backtestResult?.note && (
                <div style={{ fontSize: 12, color: "#6b7a99", padding: "14px 0" }}>
                  {backtestResult.note}
                </div>
              )}

              {!backtestLoading && !backtestError && backtestResult && !backtestResult.note && backtestResult.target?.summary && (() => {
                const t = backtestResult.target;
                const c = backtestResult.current;
                const ts = t.summary;
                const cs = c?.summary;
                const hasOverlap = !!(ts?.months && cs?.months);

                const pct = (v, dp = 1) => (v == null ? "—" : `${(v * 100).toFixed(dp)}%`);
                const fee = (v) => (v == null ? "—" : `${v.toFixed(2)}%`);
                const portVal = (d.investableAssets || 0) * 1e6;
                const drawdownDollar = (md) =>
                  md != null && portVal > 0
                    ? ` (−${fmtDollar(portVal * Math.abs(md))})`
                    : "";

                return (
                  <>
                    {ts.startDate && ts.endDate && (
                      <div style={{ fontSize: 11, color: "#6b7a99", marginBottom: 14 }}>
                        Backtested {ts.startDate} – {ts.endDate} ({ts.months} months of overlapping history)
                      </div>
                    )}

                    <div style={{ display: "grid", gridTemplateColumns: c ? "1fr 1fr" : "1fr", gap: 14 }}>
                      <Card title="Recommended (Target) Portfolio" accent="var(--navy)">
                        <Row label="Annualized Return (10-Yr Avg)" value={pct(t.avgAnnualReturn?.value)} highlight="var(--teal)" />
                        <Row label="Annualized Volatility"    value={pct(ts.annualizedVolatility)} />
                        <Row label="Max Drawdown"              value={ts.maxDrawdown != null ? `${pct(ts.maxDrawdown)}${drawdownDollar(ts.maxDrawdown)}` : "—"} />
                        <Row label="Weighted Avg. Fee"          value={fee(t.weightedFeePct)} />
                        {t.avgAnnualReturn && (t.avgAnnualReturn.minYearsUsed < 10 || t.avgAnnualReturn.coveragePct < 100) && (
                          <div style={{ fontSize: 10, color: "#9bacc8", marginTop: 8 }}>
                            Average of each holding's trailing annual returns
                            {t.avgAnnualReturn.minYearsUsed < 10 ? ` (as little as ${t.avgAnnualReturn.minYearsUsed} yr${t.avgAnnualReturn.minYearsUsed === 1 ? "" : "s"} of history for some holdings)` : ""}
                            {t.avgAnnualReturn.coveragePct < 100 ? `, covering ${t.avgAnnualReturn.coveragePct.toFixed(0)}% of the model by allocation` : ""}.
                          </div>
                        )}
                        {t.coveragePct < 100 && (
                          <div style={{ fontSize: 10, color: "#9bacc8", marginTop: 8 }}>
                            Reflects {t.coveragePct.toFixed(0)}% of the model by allocation
                            {t.excluded?.length ? ` — excludes ${t.excluded.map(e => e.name).join(", ")} (no public price history available)` : ""}.
                          </div>
                        )}
                      </Card>

                      {c && (
                        <Card title="Current Portfolio (approximate)" accent="#6b7a99">
                          <Row label="Annualized Return (10-Yr Avg)" value={pct(c.avgAnnualReturn?.value)} />
                          <Row label="Annualized Volatility"    value={pct(cs?.annualizedVolatility)} />
                          <Row label="Max Drawdown"              value={cs?.maxDrawdown != null ? `${pct(cs.maxDrawdown)}${drawdownDollar(cs.maxDrawdown)}` : "—"} />
                          <Row label="Weighted Avg. Fee"          value={fee(c.weightedFeePct)} />
                          {c.avgAnnualReturn && c.avgAnnualReturn.minYearsUsed < 10 && (
                            <div style={{ fontSize: 10, color: "#9bacc8", marginTop: 8 }}>
                              Average of trailing annual returns (as little as {c.avgAnnualReturn.minYearsUsed} yr{c.avgAnnualReturn.minYearsUsed === 1 ? "" : "s"} of history available).
                            </div>
                          )}
                          <div style={{ fontSize: 10, color: "#9bacc8", marginTop: 8 }}>
                            Approximated as {(c.concentration||0).toFixed(1)}% {c.ticker} + {(100 - (c.concentration||0)).toFixed(1)}% {backtestResult.benchmarkTicker} (a diversified benchmark standing in for the untracked remainder), based on the concentrated position noted for this client. Not a full reconstruction of the actual current portfolio.
                          </div>
                        </Card>
                      )}
                    </div>

                    {!c && (
                      <div style={{ fontSize: 11, color: "#9bacc8", marginTop: 14 }}>
                        No current-vs-target comparison shown: no concentrated stock position/concentration % was found in this client's notes. Add one to enable a side-by-side transition comparison.
                      </div>
                    )}

                    {c && hasOverlap && (() => {
                      const retDelta = (t.avgAnnualReturn?.value ?? 0) - (c.avgAnnualReturn?.value ?? 0);
                      const volDelta = (ts.annualizedVolatility ?? 0) - (cs.annualizedVolatility ?? 0);
                      const feeDelta = (t.weightedFeePct ?? 0) - (c.weightedFeePct ?? 0);
                      return (
                        <Card title="Transition Summary" accent="var(--gold)">
                          <div style={{ display: "flex", flexDirection: "column", gap: 6, fontSize: 12, color: "var(--navy)" }}>
                            <div>
                              {volDelta <= 0 ? "▼" : "▲"} {volDelta <= 0 ? "Lower" : "Higher"} historical volatility: <strong>{pct(Math.abs(volDelta))}</strong> {volDelta <= 0 ? "reduction" : "increase"} vs. current portfolio.
                            </div>
                            <div>
                              {retDelta >= 0 ? "▲" : "▼"} {retDelta >= 0 ? "Higher" : "Lower"} historical annualized return: <strong>{pct(Math.abs(retDelta))}</strong> {retDelta >= 0 ? "above" : "below"} current portfolio.
                            </div>
                            <div>
                              {feeDelta <= 0 ? "▼" : "▲"} {feeDelta <= 0 ? "Lower" : "Higher"} weighted fee: <strong>{Math.abs(feeDelta).toFixed(2)}%</strong> {feeDelta <= 0 ? "savings" : "increase"} vs. current portfolio.
                            </div>
                          </div>
                        </Card>
                      );
                    })()}

                    {/* Risk/Return chart — X=volatility, Y=return (standard orientation) */}
                    {t.avgAnnualReturn?.value != null && ts.annualizedVolatility != null && (
                      <RiskReturnChart points={[
                        {
                          x: ts.annualizedVolatility,   // risk on X axis
                          y: t.avgAnnualReturn.value,   // return on Y axis
                          color: "var(--teal)",
                          label: "Target",
                        },
                        ...(c && c.avgAnnualReturn?.value != null && cs?.annualizedVolatility != null
                          ? [{
                              x: cs.annualizedVolatility,
                              y: c.avgAnnualReturn.value,
                              color: "var(--coral)",
                              label: "Current",
                            }]
                          : []),
                      ]} />
                    )}

                    {backtestResult.missingTickers?.length > 0 && (
                      <div style={{ fontSize: 10, color: "#9bacc8", marginTop: 14 }}>
                        No price history available for: {backtestResult.missingTickers.join(", ")}.
                      </div>
                    )}

                    <div style={{
                      fontSize: 10, color: "#9bacc8", marginTop: 16, paddingTop: 12,
                      borderTop: "1px solid #EEF0F5", fontStyle: "italic",
                    }}>
                      Based on real historical monthly prices. For illustrative purposes only — past performance does not guarantee future results.
                      Efficient frontier curve is a theoretical illustration anchored to the target portfolio; it is not derived from a full mean-variance optimization.
                    </div>
                  </>
                );
              })()}
            </Section>
          )}

          </>}

          {/* Footer — page navigation + generate */}
          <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginTop: 16, paddingTop: 20, borderTop: "1px solid #E3E7EF" }}>
            <div style={{ display: "flex", alignItems: "center", gap: 10 }}>
              <button className="ppm-btn ppm-btn-secondary" onClick={onBack} style={{
                background: "#fff", border: "1.5px solid #D0D6E0",
                color: "var(--navy)", borderRadius: 10, padding: "9px 18px",
                cursor: "pointer", fontSize: 13, fontWeight: 600,
              }}>← Edit Panel</button>
              <span style={{ fontSize: 11, color: "#9bacc8" }}>Page {page} of {TOTAL_PAGES}</span>
            </div>
            <div style={{ display: "flex", gap: 10 }}>
              {page > 1 && (
                <button className="ppm-btn ppm-btn-secondary" onClick={() => setPage(p => p - 1)} style={{
                  background: "#fff", border: "1.5px solid #D0D6E0",
                  color: "var(--navy)", borderRadius: 10, padding: "11px 22px",
                  cursor: "pointer", fontSize: 13, fontWeight: 600,
                }}>← Previous</button>
              )}
              {page < TOTAL_PAGES ? (
                <button className="ppm-btn ppm-btn-dark" onClick={() => setPage(p => p + 1)} style={{
                  background: "var(--navy)", border: "none",
                  color: "#fff", borderRadius: 10, padding: "11px 24px",
                  cursor: "pointer", fontSize: 13, fontWeight: 700,
                }}>Next →</button>
              ) : (
                <button className="ppm-btn ppm-btn-dark" onClick={() => onConfirm(Object.keys(fundSwaps).length ? { ...local, fundSwaps } : local)} style={{
                  background: "var(--navy)", border: "none",
                  color: "#fff", borderRadius: 10, padding: "11px 24px",
                  cursor: "pointer", fontSize: 13, fontWeight: 700,
                }}>Generate Proposal →</button>
              )}
            </div>
          </div>

        </div>
      </div>
    </div>
  );
}
