// src/pptVisuals.js
//
// Reusable pptxgenjs visual functions — VISUAL DESIGN ONLY.
//
// These functions render slides using pptxgenjs shapes and text. They do NOT
// extract client data and they do NOT make recommendations. They only draw what
// they are handed in `data`, formatting it for a 13.333 x 7.5 (LAYOUT_WIDE) slide.
//
// Every exported function shares one signature:
//
//     fn(slide, pptx, data, C, helpers)
//
//   slide   - a pptxgenjs slide created by `pptx.addSlide()`
//   pptx    - the pptxgenjs instance (used for `pptx.ShapeType.*`)
//   data    - the proposal/client data object (read-only here)
//   C       - the color palette object from App.jsx (optional; falls back below)
//   helpers - optional { fmtM, fmtK, fmtDollar, pct, ... } formatters from App.jsx
//
// Design rules followed here:
//   - pptxgenjs shapes + text only. No HTML, no CSS.
//   - No fabricated/static numbers. Values come from `data`.
//   - When a value is missing or unparseable, the slide shows "Review".
//   - Labels and numbers are large and readable.
//   - Layout fits 13.333 x 7.5 and stays clear of the footer band (no footer drawn).

const MISSING = "Review";

// Slide geometry (LAYOUT_WIDE, inches).
const SLIDE_W = 13.333;
const MARGIN = 0.55;
const CONTENT_W = SLIDE_W - MARGIN * 2; // 12.233

// Fallback palette mirrors the `C` object in App.jsx so this module renders
// correctly even if `C` is not passed. A provided `C` always wins.
const FALLBACK_C = {
  navy: "1A2744",
  navy2: "243459",
  gold: "B8892A",
  goldLight: "D4A845",
  goldPale: "F5EDDA",
  teal: "1E7A6E",
  tealPale: "DFF1EE",
  coral: "C94F3A",
  coralPale: "FBEAE7",
  blue: "3A6BBF",
  bluePale: "E8EEF8",
  text: "1A2030",
  muted: "6E7E9A",
  border: "D5DAE5",
  bg: "F7F5F0",
  white: "FFFFFF",
  lightBar: "E8EDF5",
};

// ---------------------------------------------------------------------------
// Small internal utilities (formatting + drawing primitives)
// ---------------------------------------------------------------------------

function palette(C) {
  return { ...FALLBACK_C, ...(C || {}) };
}

// Parse a loose numeric value ("$1,200", "60%", 1200) into a finite Number or NaN.
function num(value) {
  if (value === null || value === undefined || value === "") return NaN;
  const n = Number(String(value).replace(/[$,%\s]/g, ""));
  return Number.isFinite(n) ? n : NaN;
}

// Normalize a money value to MILLIONS, matching App.jsx's estateToMillions:
// values over 1,000 are treated as raw dollars, otherwise already millions.
function toMillions(value) {
  const n = num(value);
  if (!Number.isFinite(n)) return NaN;
  return n > 1000 ? n / 1000000 : n;
}

// Normalize a rate that may be given as 0.40 or 40 into a fraction (0.40).
function asRate(value) {
  const n = num(value);
  if (!Number.isFinite(n)) return NaN;
  return n > 1 ? n / 100 : n;
}

// Normalize a percentage that may be 0.60 or 60 into a percent number (60).
function asPercent(value) {
  const n = num(value);
  if (!Number.isFinite(n)) return NaN;
  return n > 1 ? n : n * 100;
}

// First finite candidate, else NaN.
function firstFinite(...candidates) {
  for (const c of candidates) {
    const n = num(c);
    if (Number.isFinite(n)) return n;
  }
  return NaN;
}

// Compact money formatter for MILLIONS input -> "$30.0M" / "$1.2B".
function shortM(millions) {
  if (!Number.isFinite(millions)) return MISSING;
  if (Math.abs(millions) >= 1000) return `$${(millions / 1000).toFixed(1)}B`;
  return `$${millions.toFixed(Math.abs(millions) >= 100 ? 0 : 1)}M`;
}

// Percent formatter: 60 -> "60%".
function pctStr(percent) {
  if (!Number.isFinite(percent)) return MISSING;
  return `${percent.toFixed(percent >= 10 ? 0 : 1)}%`;
}

// Resolve formatters: prefer helpers from App.jsx, fall back to local versions.
// App.jsx fmtM expects a value in millions (same as shortM), pct expects a
// percent number (same as pctStr) — so they are interchangeable here.
function resolveFmt(helpers) {
  const H = helpers || {};
  return {
    money: typeof H.fmtM === "function" ? H.fmtM : shortM,
    pct: typeof H.pct === "function" ? H.pct : pctStr,
  };
}

// Standard background for every visual.
function applyBackground(slide, P) {
  slide.background = { color: P.bg };
}

// Self-contained slide header (kicker + heading + accent rule). Kept local so
// each visual controls its own layout and stays clear of the footer band.
function drawHeader(slide, pptx, P, kicker, heading, subtitle = "") {
  // top accent strip
  slide.addShape(pptx.ShapeType.rect, {
    x: 0, y: 0, w: SLIDE_W, h: 0.08,
    fill: { color: P.navy }, line: { color: P.navy },
  });
  slide.addShape(pptx.ShapeType.rect, {
    x: 0, y: 0, w: 2.2, h: 0.08,
    fill: { color: P.goldLight }, line: { color: P.goldLight },
  });

  slide.addText(String(kicker || "").toUpperCase(), {
    x: MARGIN, y: 0.34, w: CONTENT_W, h: 0.22,
    fontSize: 9, bold: true, color: P.blue, charSpace: 2, margin: 0,
  });

  slide.addText(heading || "", {
    x: MARGIN, y: 0.62, w: CONTENT_W, h: 0.6,
    fontFace: "Georgia", fontSize: 26, bold: true, color: P.navy,
    margin: 0, fit: "shrink",
  });

  slide.addShape(pptx.ShapeType.line, {
    x: MARGIN, y: 1.32, w: CONTENT_W, h: 0,
    line: { color: P.border, width: 0.9 },
  });
  slide.addShape(pptx.ShapeType.line, {
    x: MARGIN, y: 1.32, w: 0.95, h: 0,
    line: { color: P.goldLight, width: 2.6 },
  });

  if (subtitle) {
    slide.addText(subtitle, {
      x: MARGIN, y: 1.46, w: CONTENT_W, h: 0.3,
      fontSize: 11, color: P.text, margin: 0, fit: "shrink",
    });
  }
}

// A labeled value "card": large value, small label, optional left accent bar.
function drawMetricCard(slide, pptx, P, opts) {
  const {
    x, y, w, h,
    label, value,
    fill = P.white,
    valueColor = P.navy,
    labelColor = P.muted,
    accent = P.goldLight,
    valueSize = 22,
  } = opts;

  slide.addShape(pptx.ShapeType.roundRect, {
    x, y, w, h, rectRadius: 0.06,
    fill: { color: fill },
    line: { color: P.border, width: 0.75 },
    shadow: { type: "outer", color: "D9DEE8", opacity: 0.14, blur: 1, angle: 45, distance: 1 },
  });
  slide.addShape(pptx.ShapeType.rect, {
    x, y, w: 0.06, h,
    fill: { color: accent }, line: { color: accent },
  });

  slide.addText(String(label || "").toUpperCase(), {
    x: x + 0.2, y: y + 0.14, w: w - 0.32, h: 0.22,
    fontSize: 9, bold: true, color: labelColor, charSpace: 1, margin: 0, fit: "shrink",
  });
  slide.addText(value === undefined || value === null || value === "" ? MISSING : String(value), {
    x: x + 0.2, y: y + 0.4, w: w - 0.32, h: h - 0.52,
    fontSize: valueSize, bold: true, color: valueColor, margin: 0, fit: "shrink",
    valign: "middle",
  });
}

// ---------------------------------------------------------------------------
// 1. Estate Flow — asset composition -> total estate -> tax exposure to heirs
// ---------------------------------------------------------------------------

export function addEstateFlowVisual(slide, pptx, data = {}, C, helpers = {}) {
  const P = palette(C);
  const money = helpers.fmtM || shortM;

  const toM = (value) => {
    const n = Number(String(value ?? "").replace(/[$,%\s,]/g, ""));
    if (!Number.isFinite(n)) return NaN;
    return n > 1000 ? n / 1000000 : n;
  };

  const fmtMoney = (value) => {
    const n = toM(value);
    return Number.isFinite(n) ? money(n) : "Review";
  };

  const rawTotalEstate =
    toM(data.netWorth || data.totalEstate) ||
    [
      toM(data.managedAssets || data.investableAssets),
      toM(data.realEstateValue || data.realEstateHoldings),
      toM(data.otherAssets || data.otherPrivateAssets),
    ]
      .filter(Number.isFinite)
      .reduce((a, b) => a + b, 0);

  const totalEstate = Number.isFinite(rawTotalEstate) ? rawTotalEstate : NaN;
  const halfEstate = Number.isFinite(totalEstate) ? totalEstate / 2 : NaN;

  const managedAssets = toM(data.managedAssets || data.investableAssets);
  const realEstate = toM(data.realEstateValue || data.realEstateHoldings);
  const otherAssets = toM(data.otherAssets || data.otherPrivateAssets);

  const halfManaged = Number.isFinite(managedAssets) ? managedAssets / 2 : NaN;
  const halfRealEstate = Number.isFinite(realEstate) ? realEstate / 2 : NaN;
  const halfOther = Number.isFinite(otherAssets) ? otherAssets / 2 : NaN;

  const exemption = toM(data.estateTaxExemption || data.estateExemption || 30);
  const taxRateRaw = Number(data.estateTaxRate ?? 0.4);
  const taxRate = taxRateRaw > 1 ? taxRateRaw / 100 : taxRateRaw;

  const taxableEstate =
    Number.isFinite(totalEstate) && Number.isFinite(exemption)
      ? Math.max(0, totalEstate - exemption)
      : NaN;

  const estateTax =
    Number.isFinite(taxableEstate) && Number.isFinite(taxRate)
      ? taxableEstate * taxRate
      : NaN;

  const toHeirs =
    Number.isFinite(totalEstate) && Number.isFinite(estateTax)
      ? Math.max(0, totalEstate - estateTax)
      : NaN;

  const effectiveRate =
    Number.isFinite(totalEstate) && totalEstate > 0 && Number.isFinite(estateTax)
      ? estateTax / totalEstate
      : NaN;

  const clientName = data.clientName || "Client Household";
  const isFuture = String(data.estateTitle || "").toLowerCase().includes("life") ||
    String(data.estateTitle || "").toLowerCase().includes("death");

  const personA =
    data.clientFirstName ||
    data.primaryClientName ||
    (String(clientName).toLowerCase().includes("mercer") ? "Daniel" : "Client");

  const personB =
    data.spouseFirstName ||
    data.spouseName ||
    (String(clientName).toLowerCase().includes("mercer") ? "Claire" : "Spouse");

  const heirs =
    data.heirs ||
    (String(clientName).toLowerCase().includes("mercer") ? "Olivia & Ethan" : "Heirs");

  const titleText = data.estateTitle || "Estate Value Today";
  const subtitleText =
    data.estateSubtitle ||
    "Estate flow, tax exposure, and estimated wealth transfer outcome.";

  const growthYears = data.estateProjectionYears || data.projectionYears || "";
  const growthRate = data.estateGrowthRate || data.growthRate || "";

  slide.background = { color: P.bg || "F7F5F0" };

  // Header band
  slide.addShape(pptx.ShapeType.rect, {
    x: 0,
    y: 0,
    w: 13.333,
    h: 0.92,
    fill: { color: P.navy || "1A2744" },
    line: { color: P.navy || "1A2744" },
  });

  slide.addText(`${clientName} · Estate Planning`, {
    x: 0.65,
    y: 0.18,
    w: 5.8,
    h: 0.12,
    fontSize: 7.4,
    bold: true,
    color: P.goldLight || "D4A845",
    charSpace: 1.3,
    margin: 0,
    fit: "shrink",
  });

  slide.addText(titleText, {
    x: 0.65,
    y: 0.40,
    w: 7.4,
    h: 0.26,
    fontSize: 19.5,
    bold: true,
    color: P.white || "FFFFFF",
    margin: 0,
    fit: "shrink",
  });

  slide.addText(subtitleText, {
    x: 7.25,
    y: 0.39,
    w: 5.35,
    h: 0.18,
    fontSize: 7.8,
    color: "D8E4F5",
    align: "right",
    margin: 0,
    fit: "shrink",
  });

  // Alert stripe
  slide.addShape(pptx.ShapeType.roundRect, {
    x: 0.65,
    y: 1.08,
    w: 12.05,
    h: 0.42,
    rectRadius: 0.035,
    fill: { color: P.goldPale || "F5EDDA" },
    line: { color: P.goldLight || "D4A845", width: 0.65 },
  });

  slide.addText("⚠", {
    x: 0.85,
    y: 1.20,
    w: 0.25,
    h: 0.1,
    fontSize: 8.2,
    color: P.gold || "B8892A",
    margin: 0,
  });

  slide.addText(
    `Assets above ${fmtMoney(exemption)} are illustrated at ${(taxRate * 100).toFixed(0)}% estate tax. Marital deduction defers tax on first death; combined estate is exposed on second death.`,
    {
      x: 1.18,
      y: 1.20,
      w: 11.1,
      h: 0.1,
      fontSize: 7.4,
      color: P.text || "1A2030",
      margin: 0,
      fit: "shrink",
    }
  );

  // Helper styles
  const drawEstateBar = (x, y, w, label, value) => {
    slide.addShape(pptx.ShapeType.roundRect, {
      x,
      y,
      w,
      h: 0.72,
      rectRadius: 0.04,
      fill: { color: P.navy || "1A2744" },
      line: { color: P.navy || "1A2744" },
    });

    slide.addText(label, {
      x: x + 0.18,
      y: y + 0.15,
      w: w - 0.36,
      h: 0.1,
      fontSize: 7.1,
      color: "C8D2E3",
      margin: 0,
      fit: "shrink",
    });

    slide.addText(value, {
      x: x + 0.18,
      y: y + 0.35,
      w: w - 0.36,
      h: 0.16,
      fontSize: 14.5,
      bold: true,
      color: P.white || "FFFFFF",
      margin: 0,
      fit: "shrink",
    });
  };

  const drawMiniRow = (x, y, w, label, value, fill, color) => {
    slide.addShape(pptx.ShapeType.roundRect, {
      x,
      y,
      w,
      h: 0.28,
      rectRadius: 0.025,
      fill: { color: fill },
      line: { color: fill },
    });

    slide.addText(label, {
      x: x + 0.12,
      y: y + 0.085,
      w: w * 0.52,
      h: 0.08,
      fontSize: 6.4,
      color,
      bold: true,
      margin: 0,
      fit: "shrink",
    });

    slide.addText(value, {
      x: x + w * 0.56,
      y: y + 0.085,
      w: w * 0.36,
      h: 0.08,
      fontSize: 6.4,
      color,
      bold: true,
      align: "right",
      margin: 0,
      fit: "shrink",
    });
  };

  const drawSummaryCard = (x, y, w, title, value, sub, fill, color) => {
    slide.addShape(pptx.ShapeType.roundRect, {
      x,
      y,
      w,
      h: 0.92,
      rectRadius: 0.04,
      fill: { color: fill },
      line: { color: fill },
    });

    slide.addText(title, {
      x: x + 0.18,
      y: y + 0.14,
      w: w - 0.36,
      h: 0.1,
      fontSize: 7.3,
      bold: true,
      color,
      charSpace: 0.8,
      margin: 0,
      fit: "shrink",
    });

    slide.addText(value, {
      x: x + 0.18,
      y: y + 0.35,
      w: w - 0.36,
      h: 0.16,
      fontSize: 14.2,
      bold: true,
      color,
      margin: 0,
      fit: "shrink",
    });

    slide.addText(sub, {
      x: x + 0.18,
      y: y + 0.63,
      w: w - 0.36,
      h: 0.1,
      fontSize: 6.5,
      color,
      margin: 0,
      fit: "shrink",
    });
  };

  const arrow = (x, y, text = "⇢") => {
    slide.addText(text, {
      x,
      y,
      w: 0.45,
      h: 0.25,
      fontSize: 17,
      bold: true,
      color: P.muted || "6E7E9A",
      align: "center",
      margin: 0,
    });
  };

  // Section title
  slide.addText(isFuture ? "Estate Flow — Projected" : "Estate Flow — Today", {
    x: 0.8,
    y: 1.78,
    w: 4.6,
    h: 0.15,
    fontSize: 10,
    bold: true,
    color: P.navy || "1A2744",
    margin: 0,
  });

  // Person split
  const leftX = 0.8;
  const rightX = 7.0;
  const topY = 2.12;
  const boxW = 5.2;

  slide.addText(`${personA}'s ${isFuture ? "projected estate" : "estate"}`, {
    x: leftX,
    y: topY,
    w: boxW,
    h: 0.1,
    fontSize: 7.2,
    bold: true,
    color: P.muted || "6E7E9A",
    charSpace: 0.7,
    margin: 0,
    fit: "shrink",
  });

  slide.addText(`${personB}'s ${isFuture ? "projected estate" : "estate"}`, {
    x: rightX,
    y: topY,
    w: boxW,
    h: 0.1,
    fontSize: 7.2,
    bold: true,
    color: P.muted || "6E7E9A",
    charSpace: 0.7,
    margin: 0,
    fit: "shrink",
  });

  drawEstateBar(
    leftX,
    2.35,
    boxW,
    isFuture ? "Projected value at life expectancy" : "Total assets (50% split)",
    fmtMoney(halfEstate)
  );

  drawEstateBar(
    rightX,
    2.35,
    boxW,
    isFuture ? "Projected value at life expectancy" : "Total assets (50% split)",
    fmtMoney(halfEstate)
  );

  arrow(6.22, 2.62);

  if (!isFuture) {
    slide.addText(
      `Managed: ${fmtMoney(halfManaged)} · Real Estate: ${fmtMoney(halfRealEstate)} · Other: ${fmtMoney(halfOther)}`,
      {
        x: leftX + 0.1,
        y: 3.16,
        w: boxW - 0.2,
        h: 0.1,
        fontSize: 6.8,
        color: P.muted || "6E7E9A",
        margin: 0,
        fit: "shrink",
      }
    );
  } else {
    const projectionText =
      growthYears || growthRate
        ? `${fmtMoney(halfEstate)} projected based on growth assumptions`
        : "Projected value based on planning assumptions";

    slide.addText(projectionText, {
      x: leftX + 0.1,
      y: 3.16,
      w: boxW - 0.2,
      h: 0.1,
      fontSize: 6.8,
      color: P.muted || "6E7E9A",
      margin: 0,
      fit: "shrink",
    });
  }

  drawMiniRow(
    rightX,
    3.08,
    boxW,
    "Taxes on first death",
    "$0 marital deduction",
    P.coralPale || "FBEAE7",
    P.coral || "C94F3A"
  );

  drawMiniRow(
    rightX,
    3.42,
    boxW,
    "Inheritance passed to survivor",
    fmtMoney(halfEstate),
    P.bluePale || "E8EEF8",
    P.blue || "3A6BBF"
  );

  // Combined estate
  slide.addText(isFuture ? "↓ Combined estate at last death" : "↓ Combined estate on second death", {
    x: 0.8,
    y: 3.95,
    w: 11.8,
    h: 0.1,
    fontSize: 6.8,
    color: P.muted || "6E7E9A",
    align: "center",
    margin: 0,
  });

  slide.addShape(pptx.ShapeType.roundRect, {
    x: 0.8,
    y: 4.18,
    w: 11.8,
    h: 0.58,
    rectRadius: 0.04,
    fill: { color: P.navy || "1A2744" },
    line: { color: P.navy || "1A2744" },
  });

  slide.addText(isFuture ? "Combined projected estate" : "Combined estate — total net worth", {
    x: 1.05,
    y: 4.38,
    w: 4.2,
    h: 0.1,
    fontSize: 7.5,
    color: "C8D2E3",
    margin: 0,
    fit: "shrink",
  });

  slide.addText(fmtMoney(totalEstate), {
    x: 7.1,
    y: 4.32,
    w: 4.9,
    h: 0.16,
    fontSize: 15.8,
    bold: true,
    color: P.white || "FFFFFF",
    align: "right",
    margin: 0,
    fit: "shrink",
  });

  slide.addText(`↓ After ${(taxRate * 100).toFixed(0)}% estate tax on amounts above ${fmtMoney(exemption)}`, {
    x: 0.8,
    y: 4.94,
    w: 11.8,
    h: 0.1,
    fontSize: 6.8,
    color: P.muted || "6E7E9A",
    align: "center",
    margin: 0,
  });

  const effectiveText = Number.isFinite(effectiveRate)
    ? `${(effectiveRate * 100).toFixed(1)}% effective rate`
    : "Effective rate requires review";

  drawSummaryCard(
    0.8,
    5.22,
    5.65,
    isFuture ? "Estate Tax at Life Expectancy" : "Estate Tax Today",
    `(${fmtMoney(estateTax)})`,
    `${(taxRate * 100).toFixed(0)}% on ${fmtMoney(taxableEstate)} taxable · ${effectiveText}`,
    P.coralPale || "FBEAE7",
    P.coral || "C94F3A"
  );

  drawSummaryCard(
    6.95,
    5.22,
    5.65,
    isFuture ? "Net to Heirs at Life Expectancy" : "Net to Heirs Today",
    fmtMoney(toHeirs),
    `${heirs} — without further planning`,
    P.tealPale || "DFF1EE",
    P.teal || "1E7A6E"
  );

  // Assumption footer strip
  slide.addShape(pptx.ShapeType.roundRect, {
    x: 0.8,
    y: 6.36,
    w: 11.8,
    h: 0.36,
    rectRadius: 0.035,
    fill: { color: P.white || "FFFFFF" },
    line: { color: P.border || "D5DAE5", width: 0.6 },
  });

  const assumptionText = isFuture
    ? `Projection assumptions: ${growthRate ? `${growthRate} growth` : "growth assumptions"} · ${growthYears ? `${growthYears} years` : "projection horizon"} · ${fmtMoney(exemption)} exemption`
    : `Net worth split: ${fmtMoney(halfEstate)} each · Managed assets: ${fmtMoney(managedAssets)} · Real estate: ${fmtMoney(realEstate)} · Other assets: ${fmtMoney(otherAssets)}`;

  slide.addText(assumptionText, {
    x: 1.0,
    y: 6.48,
    w: 11.4,
    h: 0.08,
    fontSize: 6.5,
    color: P.text || "1A2030",
    align: "center",
    margin: 0,
    fit: "shrink",
  });
}


export function addRiskOverviewVisual(slide, pptx, data, C, helpers) {
  const P = palette(C);
  const { pct } = resolveFmt(helpers);
  const d = data || {};

  applyBackground(slide, P);
  drawHeader(
    slide, pptx, P,
    "Portfolio Module",
    "Risk Profile & Allocation",
    "Recommended risk posture and the resulting growth / stability split."
  );

  // --- Risk profile name (dynamic) ---
  const rawLabel = d.riskLabel || d.riskProfileLabel || d.riskProfile || "";
  // App.jsx labels look like "Balanced — 50/50"; show the name, parse the mix.
  const profileName = rawLabel ? String(rawLabel).split("—")[0].trim() : MISSING;

  // --- Equity / fixed split (dynamic) ---
  // Accept explicit percentages, or parse a "60/40" mix from label/field.
  let equity = firstFinite(d.equityPct, d.equityAllocation);
  let fixed = firstFinite(d.fixedIncomePct, d.fixedAllocation, d.bondPct);
  const mixSource = String(d.riskMix || rawLabel || "");
  const mixMatch = mixSource.match(/(\d{1,3})\s*\/\s*(\d{1,3})/);
  if ((!Number.isFinite(equity) || !Number.isFinite(fixed)) && mixMatch) {
    equity = Number(mixMatch[1]);
    fixed = Number(mixMatch[2]);
  }
  const haveSplit = Number.isFinite(equity) && Number.isFinite(fixed) && equity + fixed > 0;

  // Left: large profile card + risk number.
  const leftX = MARGIN;
  const leftW = 4.1;
  slide.addShape(pptx.ShapeType.roundRect, {
    x: leftX, y: 1.95, w: leftW, h: 2.0, rectRadius: 0.06,
    fill: { color: P.navy }, line: { color: P.navy },
  });
  slide.addText("RECOMMENDED RISK PROFILE", {
    x: leftX + 0.28, y: 2.18, w: leftW - 0.5, h: 0.25,
    fontSize: 9.5, bold: true, color: P.goldLight, charSpace: 1, margin: 0,
  });
  slide.addText(profileName, {
    x: leftX + 0.28, y: 2.5, w: leftW - 0.5, h: 0.9,
    fontFace: "Georgia", fontSize: 26, bold: true, color: P.white, margin: 0, fit: "shrink", valign: "middle",
  });
  const riskNumber = firstFinite(d.riskNumber, d.riskScore);
  slide.addText(
    Number.isFinite(riskNumber) ? `Risk Number: ${Math.round(riskNumber)}` : "Risk Number: Review",
    {
      x: leftX + 0.28, y: 3.45, w: leftW - 0.5, h: 0.3,
      fontSize: 12, color: P.lightBar, margin: 0,
    }
  );

  // Right: allocation split bar with large percentage labels.
  const barX = leftX + leftW + 0.4;
  const barW = SLIDE_W - MARGIN - barX;
  slide.addText("GROWTH / STABILITY SPLIT", {
    x: barX, y: 1.95, w: barW, h: 0.25,
    fontSize: 9.5, bold: true, color: P.muted, charSpace: 1, margin: 0,
  });

  const splitY = 2.4;
  const splitH = 0.95;
  if (haveSplit) {
    const total = equity + fixed;
    const eqW = Math.max(0.4, barW * (equity / total));
    const fiW = Math.max(0.4, barW - eqW);
    // Equity (growth) segment
    slide.addShape(pptx.ShapeType.roundRect, {
      x: barX, y: splitY, w: eqW, h: splitH, rectRadius: 0.04,
      fill: { color: P.teal }, line: { color: P.teal },
    });
    slide.addText(`${Math.round(equity)}%`, {
      x: barX, y: splitY, w: eqW, h: splitH,
      fontSize: 28, bold: true, color: P.white, align: "center", valign: "middle", margin: 0,
    });
    // Fixed income (stability) segment
    slide.addShape(pptx.ShapeType.roundRect, {
      x: barX + eqW, y: splitY, w: fiW, h: splitH, rectRadius: 0.04,
      fill: { color: P.blue }, line: { color: P.blue },
    });
    slide.addText(`${Math.round(fixed)}%`, {
      x: barX + eqW, y: splitY, w: fiW, h: splitH,
      fontSize: 28, bold: true, color: P.white, align: "center", valign: "middle", margin: 0,
    });
    slide.addText("Equity / Growth", {
      x: barX, y: splitY + splitH + 0.08, w: barW / 2, h: 0.25,
      fontSize: 10, bold: true, color: P.teal, margin: 0,
    });
    slide.addText("Fixed Income / Stability", {
      x: barX + barW / 2, y: splitY + splitH + 0.08, w: barW / 2, h: 0.25,
      fontSize: 10, bold: true, color: P.blue, align: "right", margin: 0,
    });
  } else {
    slide.addShape(pptx.ShapeType.roundRect, {
      x: barX, y: splitY, w: barW, h: splitH, rectRadius: 0.04,
      fill: { color: P.lightBar }, line: { color: P.border, width: 0.75 },
    });
    slide.addText(MISSING, {
      x: barX, y: splitY, w: barW, h: splitH,
      fontSize: 24, bold: true, color: P.muted, align: "center", valign: "middle", margin: 0,
    });
  }

  // --- Supporting facts row (large stat cards) ---
  const factsY = 4.55;
  const fGap = 0.3;
  const fW = (CONTENT_W - fGap * 3) / 4;
  const horizon = firstFinite(d.timeHorizon, d.timeHorizonYears, d.investmentHorizon);
  const liquidity = toMillions(d.liquidityNeed ?? d.withdrawalNeed ?? d.retirementIncomeNeed);
  const concentration = asPercent(d.concentration);
  const age = firstFinite(d.age, d.clientAge);

  drawMetricCard(slide, pptx, P, {
    x: MARGIN, y: factsY, w: fW, h: 1.15,
    label: "Time Horizon",
    value: Number.isFinite(horizon) ? `${Math.round(horizon)} yrs` : MISSING,
    valueColor: P.navy, accent: P.gold, valueSize: 24,
  });
  drawMetricCard(slide, pptx, P, {
    x: MARGIN + (fW + fGap), y: factsY, w: fW, h: 1.15,
    label: "Liquidity Need",
    value: Number.isFinite(liquidity) ? (helpers && helpers.fmtM ? helpers.fmtM(liquidity) : shortM(liquidity)) : MISSING,
    valueColor: P.navy, accent: P.gold, valueSize: 24,
  });
  drawMetricCard(slide, pptx, P, {
    x: MARGIN + (fW + fGap) * 2, y: factsY, w: fW, h: 1.15,
    label: "Concentration",
    value: Number.isFinite(concentration) ? pct(concentration) : MISSING,
    valueColor: P.coral, accent: P.coral, valueSize: 24,
  });
  drawMetricCard(slide, pptx, P, {
    x: MARGIN + (fW + fGap) * 3, y: factsY, w: fW, h: 1.15,
    label: "Client Age",
    value: Number.isFinite(age) ? String(Math.round(age)) : MISSING,
    valueColor: P.navy, accent: P.gold, valueSize: 24,
  });
}

// ---------------------------------------------------------------------------
// 3. Next Steps — numbered action items, each with a dynamic data chip
// ---------------------------------------------------------------------------

export function addNextStepsVisual(slide, pptx, data, C, helpers) {
  const P = palette(C);
  const { money, pct } = resolveFmt(helpers);
  const d = data || {};

  applyBackground(slide, P);
  drawHeader(
    slide, pptx, P,
    "Next Steps",
    "Recommended Action Plan",
    "Move from proposal review to implementation. Items marked “Review” need confirmation."
  );

  // Dynamic values surfaced alongside each step (NaN/empty => "Review").
  const ticker = d.ticker ? String(d.ticker).toUpperCase() : "";
  const netWorth = (() => {
    const nw = toMillions(d.netWorth);
    return Number.isFinite(nw) ? nw : NaN;
  })();
  const riskLabel = d.riskLabel || d.riskProfile || "";
  const riskName = riskLabel ? String(riskLabel).split("—")[0].trim() : "";
  const concentration = asPercent(d.concentration);
  const position = toMillions(d.stockPosition);
  const rate = asRate(d.totalTaxRate ?? d.taxRate);

  // [title, description, chipLabel, chipValue]
  const steps = [
    [
      "Confirm client facts",
      "Validate goals, time horizon, liquidity needs, and the extracted financial data.",
      "Net Worth",
      money(netWorth),
    ],
    [
      "Validate the risk profile",
      "Confirm the recommended risk posture and target allocation are suitable.",
      "Risk Profile",
      riskName || MISSING,
    ],
    [
      "Address concentration",
      `Review strategies to reduce single-stock exposure${ticker ? ` in ${ticker}` : ""}.`,
      "Concentration",
      Number.isFinite(concentration) ? pct(concentration) : MISSING,
    ],
    [
      "Review tax & estate exposure",
      "Coordinate tax, trust, and estate items with the CPA and attorney.",
      ticker ? `${ticker} Position` : "Position",
      money(position),
    ],
    [
      "Approve & implement",
      "Approve the final action plan and prepare client-ready documents.",
      "Tax Rate",
      Number.isFinite(rate) ? pct(rate * 100) : MISSING,
    ],
  ];

  const startY = 1.95;
  const rowH = 0.9;
  const chipW = 2.6;
  const chipX = SLIDE_W - MARGIN - chipW;

  steps.forEach((step, i) => {
    const [titleText, desc, chipLabel, chipValue] = step;
    const y = startY + i * rowH;

    // Numbered circle
    slide.addShape(pptx.ShapeType.ellipse, {
      x: MARGIN, y: y + 0.05, w: 0.5, h: 0.5,
      fill: { color: P.navy }, line: { color: P.navy },
    });
    slide.addText(String(i + 1), {
      x: MARGIN, y: y + 0.05, w: 0.5, h: 0.5,
      fontSize: 16, bold: true, color: P.white, align: "center", valign: "middle", margin: 0,
    });

    // Step title + description
    const textX = MARGIN + 0.75;
    slide.addText(titleText, {
      x: textX, y: y, w: chipX - textX - 0.3, h: 0.3,
      fontSize: 15, bold: true, color: P.navy, margin: 0, fit: "shrink",
    });
    slide.addText(desc, {
      x: textX, y: y + 0.33, w: chipX - textX - 0.3, h: 0.3,
      fontSize: 10.5, color: P.text, margin: 0, fit: "shrink",
    });

    // Dynamic data chip (right aligned)
    const isReview = chipValue === undefined || chipValue === null || chipValue === "" || chipValue === MISSING;
    slide.addShape(pptx.ShapeType.roundRect, {
      x: chipX, y: y + 0.02, w: chipW, h: rowH - 0.22, rectRadius: 0.06,
      fill: { color: isReview ? P.bg : P.goldPale },
      line: { color: P.border, width: 0.75 },
    });
    slide.addText(String(chipLabel).toUpperCase(), {
      x: chipX + 0.18, y: y + 0.1, w: chipW - 0.36, h: 0.18,
      fontSize: 8, bold: true, color: P.muted, charSpace: 1, margin: 0,
    });
    slide.addText(isReview ? MISSING : String(chipValue), {
      x: chipX + 0.18, y: y + 0.3, w: chipW - 0.36, h: 0.32,
      fontSize: 18, bold: true, color: isReview ? P.muted : P.gold, margin: 0, fit: "shrink",
    });

    // Divider line under the row (skip after the last row)
    if (i < steps.length - 1) {
      slide.addShape(pptx.ShapeType.line, {
        x: textX, y: y + rowH - 0.12, w: chipX - textX - 0.3, h: 0,
        line: { color: P.border, width: 0.6 },
      });
    }
  });
}

export default {
  addEstateFlowVisual,
  addRiskOverviewVisual,
  addNextStepsVisual,
};
