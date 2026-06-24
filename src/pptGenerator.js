import pptxgen from "pptxgenjs";
import { toPng } from "html-to-image";
import { addEstateFlowVisual, addNextStepsVisual } from "./pptVisuals";
import { buildEstateSlideHtml } from "./htmlVisualTemplates";
import { cleanNum, fmtM, fmtK, fmtDollar, pct } from "./formatters";
import { getFunds, getGroupTotals, getSubGroupTotals } from "./portfolioData";

async function loadSvgDataUri(publicPath) {
  const response = await fetch(publicPath);
  if (!response.ok) throw new Error(`Could not load SVG visual: ${publicPath}`);
  const svgText = await response.text();
  return "data:image/svg+xml;base64," + btoa(unescape(encodeURIComponent(svgText)));
}

async function captureHtmlSlideAsPng(html) {
  const wrapper = document.createElement("div");
  wrapper.style.position = "fixed";
  wrapper.style.left = "-9999px";
  wrapper.style.top = "0";
  wrapper.style.width = "1280px";
  wrapper.style.height = "720px";
  wrapper.style.background = "#FFFFFF";
  wrapper.innerHTML = html;
  document.body.appendChild(wrapper);
  try {
    await new Promise((resolve) => requestAnimationFrame(resolve));
    return await toPng(wrapper.firstElementChild, {
      cacheBust: true,
      pixelRatio: 2,
      backgroundColor: "#FFFFFF",
    });
  } finally {
    document.body.removeChild(wrapper);
  }
}

/**
 * Generate a PowerPoint blob.
 * @param {object} params
 * @returns {Promise<Blob>}
 */
export async function generatePowerPoint({
  name,
  notes,
  data,
  selectedStrategies = {},
  selectedProposalModules = {},
  selectedPortfolioStrategies = {},
  selectedRiskProfile = "",
  selectedServices = {},
  includePortfolioStrategySlides = true,
  includeConcentratedStockSlides = false,
  clientType = "",
  strategyLabels = [],
  assumptions = [],
  riskNumber = null,
  fundSwaps = {},
  backtest = null,
}) {

  // Convenience aliases that match the in-component variable names
  const modules = selectedProposalModules;

  const RISK_LABEL_MAP = {
    conservative: "Conservative — 20/80",
    moderatelyConservative: "Moderately Conservative — 30/70",
    conservativePlus: "Conservative Plus — 40/60",
    balanced: "Balanced — 50/50",
    balancedPlus: "Balanced Plus — 60/40",
    growth: "Growth — 70/30",
    growthPlus: "Growth Plus — 80/20",
    aggressive: "Aggressive — 100/0",
  };
  const riskProfileLabel = RISK_LABEL_MAP[selectedRiskProfile] || "Review";

  const selectedPortfolioStrategyLabel = (() => {
    const PORTFOLIO_LABEL_MAP = {
      corePrivate: "Core Private",
      selectLiquidity: "Select Liquidity",
      traditional: "Traditional",
      focusedB: "Focused B",
      selectLiquidityUsBias: "Select Liquidity — U.S. Bias",
      traditionalUsBias: "Traditional — U.S. Bias",
    };
    const key = Object.keys(selectedPortfolioStrategies || {}).find(k => selectedPortfolioStrategies[k]);
    return PORTFOLIO_LABEL_MAP[key] || "Review";
  })();

      const pptx = new pptxgen();
      pptx.layout = "LAYOUT_WIDE";
      pptx.author = "Investment Proposal Agent";
      pptx.subject = "Concentrated Stock Investment Proposal";
      pptx.title = `${name} Styled Investment Proposal`;

      const C = {
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
        bg: "FFFFFF",
        white: "FFFFFF",
        lightBar: "E8EDF5",
      };

      function title(slide, num, heading, subtitle = "") {
        slide.background = { color: C.bg };

        // subtle top accent
        slide.addShape(pptx.ShapeType.rect, {
          x: 0,
          y: 0,
          w: 13.33,
          h: 0.08,
          fill: { color: C.navy },
          line: { color: C.navy },
        });

        slide.addShape(pptx.ShapeType.rect, {
          x: 0,
          y: 0,
          w: 2.2,
          h: 0.08,
          fill: { color: C.goldLight },
          line: { color: C.goldLight },
        });

        slide.addText(num, {
          x: 0.55,
          y: 0.34,
          w: 2.7,
          h: 0.2,
          fontSize: 7.5,
          bold: true,
          color: C.blue,
          charSpace: 2,
          margin: 0,
        });

        slide.addText(heading, {
          x: 0.55,
          y: 0.62,
          w: 11.7,
          h: 0.55,
          fontFace: "Georgia",
          fontSize: 26,
          bold: true,
          color: C.navy,
          margin: 0,
          fit: "shrink",
        });

        slide.addShape(pptx.ShapeType.line, {
          x: 0.55,
          y: 1.24,
          w: 11.9,
          h: 0,
          line: { color: C.border, width: 0.9 },
        });

        slide.addShape(pptx.ShapeType.line, {
          x: 0.55,
          y: 1.24,
          w: 0.95,
          h: 0,
          line: { color: C.goldLight, width: 2.6 },
        });

        if (subtitle) {
          slide.addText(subtitle, {
            x: 0.55,
            y: 1.42,
            w: 11.7,
            h: 0.36,
            fontSize: 10,
            color: C.text,
            fit: "shrink",
            margin: 0,
            breakLine: false,
          });
        }
      }

      function statBox(slide, x, y, w, label, value, fill = C.white, valueColor = C.navy) {
        slide.addShape(pptx.ShapeType.roundRect, {
          x,
          y,
          w,
          h: 0.75,
          rectRadius: 0.07,
          fill: { color: fill },
          line: { color: C.border, width: 0.5 },
          shadow: { type: "outer", color: "D9DEE8", opacity: 0.10, blur: 1, angle: 45, distance: 1 },
        });

        slide.addShape(pptx.ShapeType.rect, {
          x,
          y,
          w: 0.05,
          h: 0.75,
          fill: { color: valueColor },
          line: { color: valueColor },
        });

        slide.addText(value, {
          x: x + 0.08,
          y: y + 0.10,
          w: w - 0.16,
          h: 0.26,
          fontFace: "Georgia",
          fontSize: 16,
          bold: true,
          align: "center",
          color: valueColor,
          fit: "shrink",
          margin: 0,
        });

        slide.addText(label.toUpperCase(), {
          x: x + 0.1,
          y: y + 0.46,
          w: w - 0.2,
          h: 0.14,
          fontSize: 6.5,
          bold: true,
          color: C.muted,
          align: "center",
          charSpace: 0.8,
          margin: 0,
        });
      }

      function card(slide, x, y, w, h, heading, body, fill = C.white) {
        slide.addShape(pptx.ShapeType.roundRect, {
          x,
          y,
          w,
          h,
          rectRadius: 0.08,
          fill: { color: fill },
          line: { color: C.border, width: 0.7 },
          shadow: { type: "outer", color: "D9DEE8", opacity: 0.12, blur: 1, angle: 45, distance: 1 },
        });

        slide.addShape(pptx.ShapeType.rect, {
          x,
          y,
          w: 0.07,
          h,
          fill: { color: C.goldLight },
          line: { color: C.goldLight },
        });

        slide.addText(heading.toUpperCase(), {
          x: x + 0.22,
          y: y + 0.17,
          w: w - 0.44,
          h: 0.18,
          fontSize: 7.2,
          bold: true,
          color: C.blue,
          charSpace: 1.1,
          margin: 0,
          fit: "shrink",
        });

        slide.addShape(pptx.ShapeType.line, {
          x: x + 0.22,
          y: y + 0.47,
          w: w - 0.44,
          h: 0,
          line: { color: C.border, width: 0.6 },
        });

        slide.addText(body, {
          x: x + 0.22,
          y: y + 0.61,
          w: w - 0.44,
          h: h - 0.76,
          fontSize: 11,
          color: C.text,
          fit: "shrink",
          valign: "top",
          margin: 0,
          breakLine: false,
        });
      }

      function footer(slide) {
        slide.addShape(pptx.ShapeType.line, {
          x: 0.55,
          y: 6.82,
          w: 11.9,
          h: 0,
          line: { color: C.border, width: 0.55 },
        });

        slide.addText(
          "For discussion purposes only. Not legal, tax, or investment advice. Figures are illustrative estimates.",
          {
            x: 0.55,
            y: 6.92,
            w: 11.9,
            h: 0.16,
            fontSize: 6.4,
            color: C.muted,
            align: "center",
            margin: 0,
          }
        );
      }

      function hCompare(slide, x, y, label, withoutVal, withVal, formatter, note = "") {
        const maxVal = Math.max(withoutVal, withVal, 1);
        const trackW = 5.0;

        slide.addText(label, {
          x,
          y,
          w: 3.7,
          h: 0.18,
          fontSize: 9.2,
          bold: true,
          color: C.text,
          margin: 0,
        });

        slide.addText("Without strategy", {
          x,
          y: y + 0.28,
          w: 1.35,
          h: 0.16,
          fontSize: 7.2,
          color: C.muted,
          margin: 0,
        });

        slide.addShape(pptx.ShapeType.roundRect, {
          x: x + 1.6,
          y: y + 0.30,
          w: trackW,
          h: 0.16,
          rectRadius: 0.08,
          fill: { color: C.lightBar },
          line: { color: C.lightBar },
        });

        slide.addShape(pptx.ShapeType.roundRect, {
          x: x + 1.6,
          y: y + 0.30,
          w: Math.max(0.12, trackW * (withoutVal / maxVal)),
          h: 0.16,
          rectRadius: 0.08,
          fill: { color: C.coral },
          line: { color: C.coral },
        });

        slide.addText(formatter(withoutVal), {
          x: x + 6.8,
          y: y + 0.25,
          w: 1.0,
          h: 0.18,
          fontSize: 8,
          bold: true,
          color: C.coral,
          align: "right",
          margin: 0,
        });

        slide.addText("With strategy", {
          x,
          y: y + 0.60,
          w: 1.35,
          h: 0.16,
          fontSize: 7.2,
          color: C.muted,
          margin: 0,
        });

        slide.addShape(pptx.ShapeType.roundRect, {
          x: x + 1.6,
          y: y + 0.62,
          w: trackW,
          h: 0.16,
          rectRadius: 0.08,
          fill: { color: C.lightBar },
          line: { color: C.lightBar },
        });

        slide.addShape(pptx.ShapeType.roundRect, {
          x: x + 1.6,
          y: y + 0.62,
          w: Math.max(0.12, trackW * (withVal / maxVal)),
          h: 0.16,
          rectRadius: 0.08,
          fill: { color: C.teal },
          line: { color: C.teal },
        });

        slide.addText(formatter(withVal), {
          x: x + 6.8,
          y: y + 0.57,
          w: 1.0,
          h: 0.18,
          fontSize: 8,
          bold: true,
          color: C.teal,
          align: "right",
          margin: 0,
        });

        if (note) {
          slide.addText(note, {
            x,
            y: y + 0.93,
            w: 7.8,
            h: 0.15,
            fontSize: 7,
            color: C.muted,
            margin: 0,
            fit: "shrink",
          });
        }
      }


      function addSvg(slide, svg, x, y, w, h) {
        const data = "data:image/svg+xml;base64," + btoa(unescape(encodeURIComponent(svg)));
        slide.addImage({ data, x, y, w, h });
      }

      function makeHarvestSvg(data) {
        const longOnly = [11, 18, 26, 34, 41, 47, 53, 58, 63, 67];
        const enhanced = [22, 42, 61, 79, 97, 114, 131, 148, 164, 180];
        const labels = ["Y1","Y2","Y3","Y4","Y5","Y6","Y7","Y8","Y9","Y10"];
        const maxVal = 190;

        function pt(i, val) {
          const x = 70 + (i / (labels.length - 1)) * 820;
          const y = 285 - (val / maxVal) * 235;
          return `${x},${y}`;
        }

        const longPts = longOnly.map((v, i) => pt(i, v)).join(" ");
        const enhPts = enhanced.map((v, i) => pt(i, v)).join(" ");

        const grid = [0,50,100,150].map(v => {
          const y = 285 - (v / maxVal) * 235;
          return `<line x1="70" y1="${y}" x2="890" y2="${y}" stroke="#E5E8EF" stroke-width="1"/>
                  <text x="55" y="${y+4}" text-anchor="end" font-size="12" fill="#6E7E9A">${v}%</text>`;
        }).join("");

        const xLabels = labels.map((l, i) => {
          const x = 70 + (i / (labels.length - 1)) * 820;
          return `<text x="${x}" y="315" text-anchor="middle" font-size="11" fill="#6E7E9A">${l}</text>`;
        }).join("");

        return `
          <svg xmlns="http://www.w3.org/2000/svg" width="960" height="350" viewBox="0 0 960 350">
            <rect width="960" height="350" fill="#FFFFFF"/>
            ${grid}
            <line x1="70" y1="285" x2="890" y2="285" stroke="#D5DAE5"/>
            <line x1="70" y1="50" x2="70" y2="285" stroke="#D5DAE5"/>

            <polyline points="${longPts}" fill="none" stroke="#3A6BBF" stroke-width="4" stroke-linecap="round" stroke-linejoin="round"/>
            <polyline points="${enhPts}" fill="none" stroke="#D4A845" stroke-width="4" stroke-linecap="round" stroke-linejoin="round"/>

            <line x1="630" y1="25" x2="670" y2="25" stroke="#D4A845" stroke-width="4"/>
            <text x="680" y="30" font-size="13" fill="#1A2030">130/30 Enhanced</text>
            <line x1="630" y1="48" x2="670" y2="48" stroke="#3A6BBF" stroke-width="4"/>
            <text x="680" y="53" font-size="13" fill="#1A2030">Long-Only</text>

            ${xLabels}
          </svg>
        `;
      }

      function makeCollarSvg(data) {
        const shares = Math.round((data.collarAllocation * 1000000) / Math.max(data.stockPrice, 1));
        const minPrice = Math.max(5, Math.round(data.stockPrice * 0.45));
        const maxPrice = Math.round(data.stockPrice * 1.65);

        const points = [];
        for (let i = 0; i < 25; i++) {
          points.push(minPrice + ((maxPrice - minPrice) * i) / 24);
        }

        function underlyingValue(price) {
          return (price * shares) / 1000000;
        }

        function collarValue(price) {
          const effectivePrice = Math.max(data.putStrike, Math.min(price, data.callStrike));
          return (effectivePrice * shares) / 1000000;
        }

        const underlying = points.map(underlyingValue);
        const collar = points.map(collarValue);
        const maxVal = Math.max(...underlying, ...collar) * 1.12;

        function pt(price, value) {
          const x = 80 + ((price - minPrice) / (maxPrice - minPrice)) * 800;
          const y = 285 - (value / maxVal) * 235;
          return `${x},${y}`;
        }

        const underlyingPts = points.map((p, i) => pt(p, underlying[i])).join(" ");
        const collarPts = points.map((p, i) => pt(p, collar[i])).join(" ");

        const grid = [0, .25, .5, .75, 1].map(r => {
          const val = maxVal * r;
          const y = 285 - r * 235;
          return `<line x1="80" y1="${y}" x2="880" y2="${y}" stroke="#E5E8EF" stroke-width="1"/>
                  <text x="65" y="${y+4}" text-anchor="end" font-size="11" fill="#6E7E9A">${fmtM(val)}</text>`;
        }).join("");

        const priceTicks = [minPrice, data.putStrike, data.stockPrice, data.callStrike, maxPrice];
        const xLabels = priceTicks.map(v => {
          const x = 80 + ((v - minPrice) / (maxPrice - minPrice)) * 800;
          return `<text x="${x}" y="315" text-anchor="middle" font-size="11" fill="#6E7E9A">$${Math.round(v)}</text>`;
        }).join("");

        const putX = 80 + ((data.putStrike - minPrice) / (maxPrice - minPrice)) * 800;
        const callX = 80 + ((data.callStrike - minPrice) / (maxPrice - minPrice)) * 800;

        return `
          <svg xmlns="http://www.w3.org/2000/svg" width="960" height="350" viewBox="0 0 960 350">
            <rect width="960" height="350" fill="#FFFFFF"/>
            ${grid}
            <line x1="80" y1="285" x2="880" y2="285" stroke="#D5DAE5"/>
            <line x1="80" y1="50" x2="80" y2="285" stroke="#D5DAE5"/>

            <rect x="80" y="50" width="${Math.max(0, putX-80)}" height="235" fill="#FBEAE7" opacity="0.65"/>
            <rect x="${putX}" y="50" width="${Math.max(0, callX-putX)}" height="235" fill="#F5EDDA" opacity="0.65"/>
            <rect x="${callX}" y="50" width="${Math.max(0, 880-callX)}" height="235" fill="#E8EEF8" opacity="0.65"/>

            <polyline points="${underlyingPts}" fill="none" stroke="#5D87B9" stroke-width="4" stroke-linecap="round" stroke-linejoin="round"/>
            <polyline points="${collarPts}" fill="none" stroke="#1A2744" stroke-width="4" stroke-linecap="round" stroke-linejoin="round"/>

            <line x1="${putX}" y1="50" x2="${putX}" y2="285" stroke="#C94F3A" stroke-width="2" stroke-dasharray="5,4"/>
            <line x1="${callX}" y1="50" x2="${callX}" y2="285" stroke="#3A6BBF" stroke-width="2" stroke-dasharray="5,4"/>

            <text x="${putX}" y="42" text-anchor="middle" font-size="12" font-weight="700" fill="#C94F3A">Put $${Math.round(data.putStrike)}</text>
            <text x="${callX}" y="42" text-anchor="middle" font-size="12" font-weight="700" fill="#3A6BBF">Call $${Math.round(data.callStrike)}</text>

            <line x1="610" y1="25" x2="650" y2="25" stroke="#5D87B9" stroke-width="4"/>
            <text x="660" y="30" font-size="13" fill="#1A2030">Underlying</text>
            <line x1="760" y1="25" x2="800" y2="25" stroke="#1A2744" stroke-width="4"/>
            <text x="810" y="30" font-size="13" fill="#1A2030">Protective Collar</text>

            ${xLabels}
          </svg>
        `;
      }

      // ── Portfolio Transformation stacked-bar SVG ──────────────────────────
      function makeAfterPortfolioSvg(data, strats) {
        const inv   = data.investableAssets || 1;
        const stock = data.stockPosition   || 0;
        const other = Math.max(inv - stock, 0);

        const crt     = strats?.crt       ? (data.crtAllocation     || 0) : 0;
        const sleeve  = strats?.harvesting ? (data.harvestingSleeve  || 0) : 0;
        const collar  = strats?.collar     ? (data.collarAllocation  || 0) : 0;
        const remStock = Math.max(stock - crt - sleeve - collar, 0);
        const newInv  = inv - crt; // CRT leaves portfolio

        // colours
        const COL = { crt:"E8C86B", sleeve:"1A7A4A", collar:"2471A3", rem:"C0392B", other:"9BACC8" };

        const W = 620, H = 300, barH = 54, bx = 120, bw = 460;

        function bar(y, total, segs) {
          let out = "", cx = bx;
          for (const { val, col, lbl } of segs) {
            if (!val || !total) continue;
            const sw = (val / total) * bw;
            out += `<rect x="${cx}" y="${y}" width="${sw}" height="${barH}" fill="#${col}"/>`;
            if (sw > 32) out += `<text x="${cx + sw/2}" y="${y + barH/2 + 4}" text-anchor="middle" font-size="10" font-weight="700" fill="#fff">${fmtM(val)}</text>`;
            cx += sw;
          }
          return out;
        }

        const beforeSegs = [
          { val: stock, col: COL.rem,   lbl: "Concentrated Stock" },
          { val: other, col: COL.other, lbl: "Other Assets" },
        ];
        const afterSegs = [
          { val: crt,      col: COL.crt,    lbl: "CRT (Charitable)" },
          { val: sleeve,   col: COL.sleeve, lbl: "Harvesting Sleeve" },
          { val: collar,   col: COL.collar, lbl: "Collared Position" },
          { val: remStock, col: COL.rem,    lbl: "Remaining Stock" },
          { val: other,    col: COL.other,  lbl: "Other Assets" },
        ];

        const legend = [
          ...(crt     ? [{ col: COL.crt,    lbl: "CRT → Charitable Trust" }]         : []),
          ...(sleeve  ? [{ col: COL.sleeve,  lbl: "Harvesting Sleeve (130/30)" }]     : []),
          ...(collar  ? [{ col: COL.collar,  lbl: "Collared Position (Protected)" }]  : []),
          { col: COL.rem,   lbl: "Remaining Concentrated Stock" },
          { col: COL.other, lbl: "Other Portfolio Assets" },
        ];

        let legendSvg = "";
        legend.forEach((l, i) => {
          const lx = 10 + (i % 3) * 200, ly = 248 + Math.floor(i / 3) * 18;
          legendSvg += `<rect x="${lx}" y="${ly}" width="12" height="12" rx="2" fill="#${l.col}"/>`;
          legendSvg += `<text x="${lx+16}" y="${ly+10}" font-size="9" fill="#4A5568">${l.lbl}</text>`;
        });

        return `<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 ${W} ${H}">
          <text x="${bx}" y="26" font-size="11" font-weight="700" fill="#1a2744" letter-spacing="0.06em">BEFORE</text>
          <text x="${bx + bw + 8}" y="26" font-size="9" fill="#9BACC8">${fmtM(inv)}</text>
          ${bar(34, inv, beforeSegs)}
          <text x="${bx}" y="126" font-size="11" font-weight="700" fill="#1a2744" letter-spacing="0.06em">AFTER</text>
          <text x="${bx + bw + 8}" y="126" font-size="9" fill="#1A7A4A">${fmtM(newInv)} invested</text>
          ${bar(134, inv, afterSegs)}
          ${crt ? `<text x="${bx + bw + 8}" y="168" font-size="9" fill="#E8C86B">+${fmtM(crt)} to CRT</text>` : ""}
          ${legendSvg}
        </svg>`;
      }

      function makeConcentrationSvg(data) {
        const labels = ["Today", "Y1", "Y3", "Y5", "Y10"];
        const without = [
          data.concentration,
          Math.min(data.concentration + 2, 95),
          Math.min(data.concentration + 5, 95),
          Math.min(data.concentration + 8, 95),
          Math.min(data.concentration + 12, 95),
        ];
        const withPlan = [
          data.concentration,
          data.afterCrtConcentration,
          Math.max(data.afterCrtConcentration - 8, 20),
          Math.max(data.afterCrtConcentration - 15, 15),
          Math.max(data.afterCrtConcentration - 25, 10),
        ];

        function pt(i, val) {
          const x = 80 + (i / (labels.length - 1)) * 800;
          const y = 285 - (val / 100) * 235;
          return `${x},${y}`;
        }

        const withoutPts = without.map((v, i) => pt(i, v)).join(" ");
        const withPts = withPlan.map((v, i) => pt(i, v)).join(" ");

        const grid = [0,25,50,75,100].map(v => {
          const y = 285 - (v / 100) * 235;
          return `<line x1="80" y1="${y}" x2="880" y2="${y}" stroke="#E5E8EF"/>
                  <text x="65" y="${y+4}" text-anchor="end" font-size="11" fill="#6E7E9A">${v}%</text>`;
        }).join("");

        const xLabels = labels.map((l, i) => {
          const x = 80 + (i / (labels.length - 1)) * 800;
          return `<text x="${x}" y="315" text-anchor="middle" font-size="11" fill="#6E7E9A">${l}</text>`;
        }).join("");

        return `
          <svg xmlns="http://www.w3.org/2000/svg" width="960" height="350" viewBox="0 0 960 350">
            <rect width="960" height="350" fill="#FFFFFF"/>
            ${grid}
            <line x1="80" y1="285" x2="880" y2="285" stroke="#D5DAE5"/>
            <line x1="80" y1="50" x2="80" y2="285" stroke="#D5DAE5"/>

            <polyline points="${withoutPts}" fill="none" stroke="#C94F3A" stroke-width="4" stroke-dasharray="8,6" stroke-linecap="round" stroke-linejoin="round"/>
            <polyline points="${withPts}" fill="none" stroke="#1E7A6E" stroke-width="4" stroke-linecap="round" stroke-linejoin="round"/>

            <line x1="600" y1="25" x2="640" y2="25" stroke="#C94F3A" stroke-width="4" stroke-dasharray="8,6"/>
            <text x="650" y="30" font-size="13" fill="#1A2030">Without Strategy</text>
            <line x1="760" y1="25" x2="800" y2="25" stroke="#1E7A6E" stroke-width="4"/>
            <text x="810" y="30" font-size="13" fill="#1A2030">With Strategy</text>

            ${xLabels}
          </svg>
        `;
      }


      const afterCrtStock = Math.max(data.stockPosition - data.crtAllocation, 0);
      const uncollaredRemaining = Math.max(afterCrtStock - data.collarAllocation, 0);
      const collaredLoss = data.collarAllocation * 0.15;
      const uncollaredLoss = uncollaredRemaining * 0.4;
      const withDownside = collaredLoss + uncollaredLoss;





      // Slide 1 Cover
      let slide = pptx.addSlide();
      slide.background = { color: C.navy };
      slide.addShape(pptx.ShapeType.rect, {
        x: 7.4,
        y: 0,
        w: 5.93,
        h: 7.5,
        fill: { color: C.navy2 },
        line: { color: C.navy2 },
      });

      slide.addText("CONFIDENTIAL · PREPARED EXCLUSIVELY FOR", {
        x: 0.75,
        y: 1.0,
        w: 6.5,
        h: 0.24,
        fontSize: 8.5,
        color: C.goldLight,
        charSpace: 2.8,
        margin: 0,
      });

      slide.addText(name, {
        x: 0.75,
        y: 1.5,
        w: 6.7,
        h: 0.9,
        fontFace: "Georgia",
        fontSize: 34,
        bold: true,
        color: C.white,
        margin: 0,
      });

      slide.addText("Wealth Strategy & Investment Proposal", {
        x: 0.75,
        y: 2.45,
        w: 6.0,
        h: 0.3,
        fontSize: 14,
        color: "D1D7E2",
        margin: 0,
      });

      slide.addShape(pptx.ShapeType.line, {
        x: 0.75,
        y: 3.0,
        w: 0.9,
        h: 0,
        line: { color: C.goldLight, width: 2.2 },
      });

      // Firm name and tagline (no client numbers on cover)
      slide.addText("MERIDIAN WEALTH PARTNERS", {
        x: 0.75, y: 3.95, w: 6.4, h: 0.24,
        fontSize: 11.5, color: C.goldLight, charSpace: 2.8, bold: true, margin: 0,
      });
      slide.addShape(pptx.ShapeType.line, {
        x: 0.75, y: 4.27, w: 1.5, h: 0,
        line: { color: C.goldLight, width: 1.5 },
      });
      const prepDate = new Date().toLocaleDateString("en-US", { month: "long", year: "numeric" });
      slide.addText(`Prepared ${prepDate}  ·  Private & Confidential`, {
        x: 0.75, y: 4.45, w: 6.4, h: 0.18,
        fontSize: 8.5, color: "8A9AB5", margin: 0,
      });

      // Right panel tagline
      slide.addText("WEALTH STRATEGY", {
        x: 8.1, y: 3.7, w: 4.5, h: 0.22,
        fontSize: 10, color: C.goldLight, charSpace: 3.5, bold: true, align: "center", margin: 0,
      });
      slide.addShape(pptx.ShapeType.line, {
        x: 9.4, y: 4.0, w: 1.9, h: 0,
        line: { color: C.goldLight, width: 0.9 },
      });
      slide.addText("Tailored Planning · Disciplined Execution · Lasting Impact", {
        x: 8.1, y: 4.18, w: 4.5, h: 0.18,
        fontSize: 7.8, color: "8A9AB5", align: "center", margin: 0,
      });


      // Slide 2 Overview
      slide = pptx.addSlide();
      title(
        slide,
        "01 · OVERVIEW",
        "Overview",
        ""
      );

      statBox(slide, 0.8, 1.38, 2.55, "Concentrated Position", fmtM(data.stockPosition), C.coralPale, C.coral);
      statBox(slide, 3.65, 1.38, 2.55, "Current Concentration", pct(data.concentration), C.goldPale, C.gold);
      statBox(slide, 6.5, 1.38, 2.55, "Tax If Fully Sold", fmtM(data.immediateTax), C.bluePale, C.blue);
      statBox(slide, 9.35, 1.38, 2.55, "After CRT Concentration", pct(data.afterCrtConcentration), C.tealPale, C.teal);

      // Dynamic goal bullets based on selected strategies
      const bOpts = { bullet: { type: "bullet" }, paraSpaceAfter: 6 };
      const overviewBullets = [];
      overviewBullets.push({ text: `Align portfolio with risk profile, income needs, and long-term objectives`, options: bOpts });
      if (data.stockPosition > 0 && data.ticker && data.ticker !== "Concentrated Stock") {
        overviewBullets.push({ text: `Reduce ${data.ticker} concentration (${pct(data.concentration)}) in a staged, tax-aware manner`, options: bOpts });
      }
      if (selectedStrategies?.crt) {
        overviewBullets.push({ text: `CRT: ~${fmtK(data.crtIncome)}/yr income + estimated charitable deduction`, options: bOpts });
      }
      if (selectedStrategies?.harvesting) {
        overviewBullets.push({ text: `Tax-loss harvesting: ~${fmtM(data.annualHarvestLosses)} first-year losses to offset gains`, options: bOpts });
      }
      if (selectedStrategies?.collar) {
        overviewBullets.push({ text: `Option collar: protect ${fmtM(data.collarAllocation)} in ${data.ticker} with bounded downside`, options: bOpts });
      }
      if (selectedStrategies?.estatePlanning || modules?.estatePlanningReview) {
        overviewBullets.push({ text: `Coordinate estate planning, trust structures, and legacy goals`, options: bOpts });
      }
      if (selectedStrategies?.muniBonds) {
        overviewBullets.push({ text: `Municipal bonds: tax-exempt income from diversification proceeds`, options: bOpts });
      }
      if (selectedStrategies?.donorAdvisedFund) {
        overviewBullets.push({ text: `Donor Advised Fund: immediate deduction + flexible grant-making`, options: bOpts });
      }
      if (overviewBullets.length < 3) {
        overviewBullets.push({ text: `Implement all selected strategies in a sequenced, tax-sensitive manner`, options: bOpts });
      }

      // Summary of Goals card
      slide.addShape(pptx.ShapeType.roundRect, { x: 0.8, y: 2.62, w: 6.5, h: 3.6, rectRadius: 0.08, fill: { color: C.white }, line: { color: C.border, width: 0.7 } });
      slide.addShape(pptx.ShapeType.rect, { x: 0.8, y: 2.62, w: 0.07, h: 3.6, fill: { color: C.goldLight }, line: { color: C.goldLight } });
      slide.addText("SUMMARY OF GOALS & AGENDA", { x: 1.02, y: 2.79, w: 6.0, h: 0.18, fontSize: 7.2, bold: true, color: C.blue, charSpace: 1.1, margin: 0 });
      slide.addShape(pptx.ShapeType.line, { x: 1.02, y: 3.09, w: 6.06, h: 0, line: { color: C.border, width: 0.6 } });
      slide.addText(overviewBullets, { x: 1.02, y: 3.2, w: 6.06, h: 2.8, fontSize: 11, color: C.text, margin: 0, valign: "top" });

      // Initial Observations — clean bullet list
      const obsLines = [];
      if (data.ticker && data.stockPosition > 0) {
        obsLines.push({ text: `${pct(data.concentration)} in ${data.ticker}  (concentration)`, options: { bullet: { type: "bullet" }, bold: false, paraSpaceAfter: 6, color: C.coral, fontSize: 11 } });
        obsLines.push({ text: `${fmtM(data.immediateTax)} tax if sold outright today`, options: { bullet: { type: "bullet" }, paraSpaceAfter: 6, color: C.coral, fontSize: 11 } });
        obsLines.push({ text: `${fmtM(data.drawdown40Impact)} at risk in a 40% drawdown`, options: { bullet: { type: "bullet" }, paraSpaceAfter: 6, color: C.coral, fontSize: 11 } });
      } else {
        obsLines.push({ text: `No concentrated single-stock position`, options: { bullet: { type: "bullet" }, bold: false, paraSpaceAfter: 6, color: C.teal, fontSize: 11 } });
        obsLines.push({ text: `${fmtM(data.investableAssets)} investable assets under management`, options: { bullet: { type: "bullet" }, paraSpaceAfter: 6, color: C.blue, fontSize: 11 } });
      }
      if (selectedStrategies?.crt) {
        obsLines.push({ text: `${pct(data.afterCrtConcentration)} concentration after CRT`, options: { bullet: { type: "bullet" }, paraSpaceAfter: 6, color: C.teal, fontSize: 11 } });
        obsLines.push({ text: `${fmtK(data.crtIncome)}/yr estimated CRT income`, options: { bullet: { type: "bullet" }, paraSpaceAfter: 0, color: C.teal, fontSize: 11 } });
      } else {
        obsLines.push({ text: `${fmtM(data.netWorth)} total net worth`, options: { bullet: { type: "bullet" }, paraSpaceAfter: 6, color: C.blue, fontSize: 11 } });
        obsLines.push({ text: `Risk profile: ${riskProfileLabel || "Review"}`, options: { bullet: { type: "bullet" }, paraSpaceAfter: 0, color: C.teal, fontSize: 11 } });
      }

      slide.addShape(pptx.ShapeType.roundRect, { x: 7.65, y: 2.62, w: 4.55, h: 3.6, rectRadius: 0.08, fill: { color: C.white }, line: { color: C.border, width: 0.7 } });
      slide.addShape(pptx.ShapeType.rect, { x: 7.65, y: 2.62, w: 0.07, h: 3.6, fill: { color: C.goldLight }, line: { color: C.goldLight } });
      slide.addText("INITIAL OBSERVATIONS", { x: 7.87, y: 2.79, w: 4.1, h: 0.18, fontSize: 7.2, bold: true, color: C.blue, charSpace: 1.1, margin: 0 });
      slide.addShape(pptx.ShapeType.line, { x: 7.87, y: 3.09, w: 4.1, h: 0, line: { color: C.border, width: 0.6 } });
      slide.addText(obsLines, { x: 7.87, y: 3.2, w: 4.1, h: 2.8, fontSize: 11, color: C.text, margin: 0, valign: "top" });

      footer(slide);

      // Removed assumptions slide.

      // Slide 3 Where Are We Today
      slide = pptx.addSlide();
      title(
        slide,
        "02 · WHERE ARE WE TODAY?",
        "Current Position & Planning Snapshot",
        `This section summarizes the client's current portfolio position, key risks, and planning context.`
      );

      statBox(slide, 0.65, 1.95, 2.55, "Total Net Worth", fmtM(data.netWorth), C.goldPale, C.gold);
      statBox(slide, 3.45, 1.95, 2.55, "Investable Assets", fmtM(data.investableAssets), C.white, C.navy);
      statBox(slide, 6.25, 1.95, 2.55, data.ticker ? `${data.ticker} Position` : "Concentrated Position", data.stockPosition > 0 ? fmtM(data.stockPosition) : "None", data.stockPosition > 0 ? C.coralPale : C.white, data.stockPosition > 0 ? C.coral : C.muted);
      statBox(slide, 9.05, 1.95, 2.55, "Annual Income", data.income ? fmtM(data.income) : "N/A", C.tealPale, C.teal);

      slide.addText(`Current Portfolio Concentration`, {
        x: 0.75,
        y: 3.15,
        w: 3.2,
        h: 0.18,
        fontSize: 10,
        bold: true,
        color: C.text,
        margin: 0,
      });

      slide.addShape(pptx.ShapeType.roundRect, {
        x: 0.75,
        y: 3.5,
        w: 11.7,
        h: 0.32,
        rectRadius: 0.1,
        fill: { color: C.lightBar },
        line: { color: C.lightBar },
      });

      const currentBarW = Math.max(0.3, 11.7 * Math.min(data.concentration, 100) / 100);
      slide.addShape(pptx.ShapeType.roundRect, {
        x: 0.75,
        y: 3.5,
        w: currentBarW,
        h: 0.32,
        rectRadius: 0.1,
        fill: { color: C.navy },
        line: { color: C.navy },
      });

      slide.addText(data.ticker && data.stockPosition > 0 ? `${data.ticker} · ${pct(data.concentration)}` : "Diversified", {
        x: 0.95,
        y: 3.58,
        w: 2.2,
        h: 0.12,
        fontSize: 8,
        bold: true,
        color: C.white,
        margin: 0,
      });

      card(
        slide,
        0.75,
        4.25,
        5.7,
        1.6,
        "Primary Risk",
        `• 40% drawdown in ${data.ticker} = ~${fmtM(data.drawdown40Impact)} loss\n• Single-stock risk until diversification executes\n• Staged plan required to manage timing and taxes`
      );

      card(
        slide,
        6.8,
        4.25,
        5.7,
        1.6,
        "Tax Challenge",
        `• Selling outright triggers ~${fmtM(data.immediateTax)} in taxes\n• Combined rate: ${pct(data.taxRate)} (federal + state)\n• Strategy must reduce concentration without a forced tax event`
      );

      footer(slide);





      // Removed old 3rd slide: Portfolio Mix / Donut Chart.

      // Slide 4 Where Are We Today — Portfolio Mix
      slide = pptx.addSlide();
      title(
        slide,
        "02B · WHERE ARE WE TODAY?",
        "Current Portfolio Mix",
        `This view shows the current portfolio composition and investable asset breakdown.`
      );

      const chartInvestable = Math.max(data.investableAssets, data.stockPosition, 0.0001);
      const chartStock = Math.min(data.stockPosition, chartInvestable);
      const chartOther = Math.max(chartInvestable - chartStock, 0);

      const concentrationPct = chartInvestable > 0 ? (chartStock / chartInvestable) * 100 : 0;

      // Specific illustrative breakdown of the remaining assets
      const diversifiedEquities = chartOther * 0.40;
      const fixedIncome = chartOther * 0.30;
      const cashLiquidity = chartOther * 0.20;
      const alternatives = chartOther * 0.10;

      const breakdown = [
        { label: `${data.ticker} Position`, value: chartStock, color: "#1A2744" },
        { label: "Diversified Equities", value: diversifiedEquities, color: "#D4A845" },
        { label: "Fixed Income", value: fixedIncome, color: "#6F8FB8" },
        { label: "Cash / Liquidity", value: cashLiquidity, color: "#7D9D9C" },
        { label: "Alternatives", value: alternatives, color: "#C9CED8" },
      ].map((item) => ({
        ...item,
        pct: chartInvestable > 0 ? (item.value / chartInvestable) * 100 : 0,
      }));

      function polarToCartesian(cx, cy, r, angleDeg) {
        const angleRad = (angleDeg - 90) * Math.PI / 180.0;
        return {
          x: cx + (r * Math.cos(angleRad)),
          y: cy + (r * Math.sin(angleRad)),
        };
      }

      function donutSegment(cx, cy, outerR, innerR, startAngle, endAngle, color) {
        const largeArc = endAngle - startAngle > 180 ? 1 : 0;

        const p1 = polarToCartesian(cx, cy, outerR, startAngle);
        const p2 = polarToCartesian(cx, cy, outerR, endAngle);
        const p3 = polarToCartesian(cx, cy, innerR, endAngle);
        const p4 = polarToCartesian(cx, cy, innerR, startAngle);

        return `
          <path
            d="
              M ${p1.x} ${p1.y}
              A ${outerR} ${outerR} 0 ${largeArc} 1 ${p2.x} ${p2.y}
              L ${p3.x} ${p3.y}
              A ${innerR} ${innerR} 0 ${largeArc} 0 ${p4.x} ${p4.y}
              Z
            "
            fill="${color}"
          />
        `;
      }

      function makeTrueDonutSvg(parts, centerLabel, subLabel) {
        const cx = 230;
        const cy = 230;
        const outerR = 155;
        const innerR = 78;

        let angle = 0;
        const segs = parts.map((part) => {
          const sweep = (part.pct / 100) * 360;
          const seg = donutSegment(cx, cy, outerR, innerR, angle, angle + sweep, part.color);
          angle += sweep;
          return seg;
        }).join("");

        return `
          <svg xmlns="http://www.w3.org/2000/svg" width="500" height="500" viewBox="0 0 500 500">
            <rect width="500" height="500" fill="#FFFFFF"/>
            ${segs}
            <circle cx="${cx}" cy="${cy}" r="${innerR - 3}" fill="#FFFFFF"/>
            <text x="${cx}" y="${cy - 6}" text-anchor="middle" font-family="Georgia" font-size="38" font-weight="700" fill="#1A2744">${centerLabel}</text>
            <text x="${cx}" y="${cy + 27}" text-anchor="middle" font-family="Inter, Arial" font-size="14" fill="#6E7E9A">${subLabel}</text>
          </svg>
        `;
      }

      const donutSvg = makeTrueDonutSvg(
        breakdown,
        `${concentrationPct.toFixed(1)}%`,
        `${data.ticker} Concentration`
      );

      addSvg(slide, donutSvg, 0.75, 1.78, 4.55, 4.55);

      slide.addShape(pptx.ShapeType.roundRect, {
        x: 6.05,
        y: 1.95,
        w: 5.75,
        h: 3.85,
        rectRadius: 0.08,
        fill: { color: C.white },
        line: { color: C.border, width: 0.8 },
      });

      slide.addText("Portfolio Breakdown", {
        x: 6.4,
        y: 2.22,
        w: 5.0,
        h: 0.22,
        fontSize: 13,
        bold: true,
        color: C.navy,
        margin: 0,
      });

      slide.addShape(pptx.ShapeType.line, {
        x: 6.4,
        y: 2.62,
        w: 5.0,
        h: 0,
        line: { color: C.border, width: 0.8 },
      });

      breakdown.forEach((item, i) => {
        const y = 3.0 + i * 0.40;

        slide.addShape(pptx.ShapeType.rect, {
          x: 6.45,
          y,
          w: 0.18,
          h: 0.18,
          fill: { color: item.color },
          line: { color: item.color },
        });

        slide.addText(item.label, {
          x: 6.78,
          y: y - 0.02,
          w: 2.8,
          h: 0.18,
          fontSize: 11,
          color: C.text,
          margin: 0,
        });

        slide.addText(`${fmtM(item.value)} | ${pct(item.pct)}`, {
          x: 9.25,
          y: y - 0.02,
          w: 1.95,
          h: 0.18,
          fontSize: 11,
          bold: true,
          color: i === 0 ? C.navy : C.gold,
          align: "right",
          margin: 0,
        });
      });

      slide.addShape(pptx.ShapeType.line, {
        x: 6.4,
        y: 5.05,
        w: 5.0,
        h: 0,
        line: { color: C.border, width: 0.8 },
      });

      slide.addText("Investable Assets", {
        x: 6.45,
        y: 5.28,
        w: 2.6,
        h: 0.18,
        fontSize: 9.5,
        color: C.muted,
        margin: 0,
      });

      slide.addText(fmtM(chartInvestable), {
        x: 9.2,
        y: 5.28,
        w: 1.95,
        h: 0.18,
        fontSize: 9.5,
        bold: true,
        color: C.text,
        align: "right",
        margin: 0,
      });

      slide.addText("Note", {
        x: 0.95,
        y: 6.18,
        w: 0.55,
        h: 0.16,
        fontSize: 8,
        bold: true,
        color: C.blue,
        charSpace: 1.0,
        margin: 0,
      });

      slide.addText(
        `Other asset categories are illustrative unless the client notes specify cash, bonds, equities, or alternatives separately.`,
        {
          x: 1.45,
          y: 6.16,
          w: 10.2,
          h: 0.18,
          fontSize: 9.4,
          color: C.text,
          fit: "shrink",
          margin: 0,
        }
      );

      footer(slide);

      // Slide 5 — Dynamic Action Plan + Implementation Pathway
      slide = pptx.addSlide();
      title(
        slide,
        "03 · ACTION PLAN",
        "Recommended Action Plan",
        `Advisor-selected strategies and implementation pathway for ${name}.`
      );

      // Build dynamic strategy card list
      const actionStrategies = [];
      if (selectedStrategies?.crt) {
        actionStrategies.push({
          title: "Charitable Remainder Trust",
          stat: fmtM(data.crtAllocation),
          statLabel: "Contribution",
          body: `• Contribute ${fmtM(data.crtAllocation)} of ${data.ticker} into CRT\n• ${pct(data.crtPayoutRate)} payout = ~${fmtK(data.crtIncome)}/yr income\n• Est. deduction: ${fmtM(data.charitableDeductionHigh)}`,
          fill: C.goldPale, color: C.gold,
        });
      }
      if (selectedStrategies?.harvesting) {
        actionStrategies.push({
          title: "Tax-Loss Harvesting",
          stat: fmtM(data.harvestingSleeve),
          statLabel: "Sleeve Size",
          body: `• ${fmtM(data.harvestingSleeve)} harvesting sleeve\n• ~${fmtM(data.annualHarvestLosses)} first-year losses\n• Est. tax benefit: ~${fmtK(data.taxSavings)} at ${pct(data.taxRate)}`,
          fill: C.tealPale, color: C.teal,
        });
      }
      if (selectedStrategies?.collar) {
        actionStrategies.push({
          title: "Option Collar",
          stat: fmtM(data.collarAllocation),
          statLabel: "Protected Value",
          body: `• Protect ${fmtM(data.collarAllocation)} in ${data.ticker}\n• Put floor ~$${(data.putStrike||0).toFixed(0)} / call cap ~$${(data.callStrike||0).toFixed(0)}\n• Bounded risk without forced sale`,
          fill: C.bluePale, color: C.blue,
        });
      }
      if (selectedStrategies?.estatePlanning) {
        actionStrategies.push({
          title: "Estate Planning",
          stat: fmtM(data.netWorth),
          statLabel: "Estate Value",
          body: `• Trust structures + beneficiary review\n• GRATs, ILITs, charitable vehicles\n• Coordinate with estate attorney + CPA`,
          fill: C.coralPale, color: C.coral,
        });
      }
      if (selectedStrategies?.muniBonds) {
        actionStrategies.push({
          title: "Municipal Bonds",
          stat: "Tax-Exempt",
          statLabel: "Income Type",
          body: `• Redirect proceeds into muni bond portfolio\n• Federally tax-free income\n• Reduces concentration + improves diversification`,
          fill: C.goldPale, color: C.gold,
        });
      }
      if (selectedStrategies?.donorAdvisedFund) {
        actionStrategies.push({
          title: "Donor Advised Fund",
          stat: "Immediate",
          statLabel: "Deduction Timing",
          body: `• Contribute appreciated shares to DAF\n• Immediate tax deduction\n• Flexible grant-making over time`,
          fill: C.tealPale, color: C.teal,
        });
      }
      if (selectedStrategies?.exchangeFund) {
        actionStrategies.push({
          title: "Exchange Fund",
          stat: "Deferred",
          statLabel: "Tax Treatment",
          body: `• Contribute concentrated shares to fund\n• Diversify without triggering capital gains\n• Deferred tax treatment`,
          fill: C.bluePale, color: C.blue,
        });
      }
      if (modules?.estatePlanningReview && !selectedStrategies?.estatePlanning) {
        actionStrategies.push({
          title: "Estate Review Module",
          stat: fmtM(data.netWorth),
          statLabel: "Estate Value",
          body: `• Review trust structures + gifting strategy\n• See Estate Planning slides for detail`,
          fill: C.coralPale, color: C.coral,
        });
      }
      if (actionStrategies.length === 0) {
        actionStrategies.push({
          title: "Portfolio Strategy",
          stat: fmtM(data.investableAssets),
          statLabel: "Investable Assets",
          body: `• Risk-aligned, diversified allocation\n• Based on selected profile + strategy\n• See strategy slides for detail`,
          fill: C.bluePale, color: C.blue,
        });
      }

      // Layout: 1–3 cards per row, up to 2 rows
      // Content starts at y:1.88 — safe below subtitle (title() subtitle ends ~y:1.78)
      const apCardCount = Math.min(actionStrategies.length, 6);
      const apPerRow = Math.min(apCardCount, 3);
      const apRows = Math.ceil(apCardCount / apPerRow);
      const apCardGap = 0.22;
      const apAreaW = 11.55;
      const apCardW = (apAreaW - (apPerRow - 1) * apCardGap) / apPerRow;
      const apCardH = apRows === 1 ? 2.2 : 1.45;
      const apCardStartY = 1.88;

      actionStrategies.slice(0, 6).forEach((strat, i) => {
        const row = Math.floor(i / apPerRow);
        const col = i % apPerRow;
        const cx = 0.85 + col * (apCardW + apCardGap);
        const cy = apCardStartY + row * (apCardH + 0.18);

        slide.addShape(pptx.ShapeType.roundRect, { x: cx, y: cy, w: apCardW, h: apCardH, rectRadius: 0.08, fill: { color: strat.fill }, line: { color: C.border, width: 0.7 } });
        slide.addShape(pptx.ShapeType.rect, { x: cx, y: cy, w: 0.06, h: apCardH, fill: { color: strat.color }, line: { color: strat.color } });

        slide.addText(strat.stat, { x: cx + 0.18, y: cy + 0.13, w: apCardW - 0.26, h: 0.27, fontFace: "Georgia", fontSize: 15, bold: true, color: strat.color, margin: 0, fit: "shrink" });
        slide.addText(strat.statLabel.toUpperCase(), { x: cx + 0.18, y: cy + 0.43, w: apCardW - 0.26, h: 0.13, fontSize: 6, bold: true, color: C.muted, charSpace: 0.8, margin: 0 });
        slide.addText(strat.title, { x: cx + 0.18, y: cy + 0.60, w: apCardW - 0.26, h: 0.18, fontSize: 9.5, bold: true, color: C.navy, margin: 0, fit: "shrink" });
        slide.addShape(pptx.ShapeType.line, { x: cx + 0.18, y: cy + 0.82, w: apCardW - 0.36, h: 0, line: { color: C.border, width: 0.5 } });
        slide.addText(strat.body, { x: cx + 0.18, y: cy + 0.92, w: apCardW - 0.26, h: apCardH - 1.0, fontSize: 10, color: C.text, margin: 0, valign: "top", fit: "shrink" });
      });

      // Implementation Pathway — separator line first, then label, then steps
      const roadmapSepY = apCardStartY + apRows * (apCardH + 0.18) + 0.2;
      const roadmapY    = roadmapSepY + 0.32;
      const rmAvailH    = 6.62 - roadmapY;

      const roadmapSteps = [{ label: "Validate Facts", body: "Confirm cost basis, tax lots, goals, and restrictions." }];
      if (selectedStrategies?.crt)      roadmapSteps.push({ label: "Establish CRT",    body: "Work with estate attorney + CPA. Transfer shares." });
      if (selectedStrategies?.harvesting) roadmapSteps.push({ label: "Build Sleeve",  body: `Set up ~${fmtM(data.harvestingSleeve)} harvesting allocation.` });
      if (selectedStrategies?.collar)   roadmapSteps.push({ label: "Price Collar",     body: `Model put ~$${(data.putStrike||0).toFixed(0)} / call ~$${(data.callStrike||0).toFixed(0)} with options desk.` });
      if (selectedStrategies?.estatePlanning || modules?.estatePlanningReview) roadmapSteps.push({ label: "Estate Review", body: "Review trusts, gifting, and beneficiary docs." });
      roadmapSteps.push({ label: "Quarterly Review", body: "Monitor concentration, taxes, income, and performance." });

      const rmSteps = roadmapSteps.slice(0, 5);
      const rmW = (apAreaW - (rmSteps.length - 1) * 0.15) / rmSteps.length;
      const rmH = Math.max(0.7, Math.min(1.15, rmAvailH - 0.02));

      if (rmAvailH > 0.75) {
        // Separator line above the section label
        slide.addShape(pptx.ShapeType.line, { x: 0.85, y: roadmapSepY, w: 11.55, h: 0, line: { color: C.border, width: 0.55 } });
        // Section label sits between separator and steps
        slide.addText("IMPLEMENTATION PATHWAY", { x: 0.85, y: roadmapSepY + 0.06, w: 4, h: 0.16, fontSize: 6.2, bold: true, color: C.blue, charSpace: 1.2, margin: 0 });

        rmSteps.forEach((step, i) => {
          const rx = 0.85 + i * (rmW + 0.15);
          const isFirst = i === 0;
          slide.addShape(pptx.ShapeType.roundRect, { x: rx, y: roadmapY, w: rmW, h: rmH, rectRadius: 0.07, fill: { color: isFirst ? C.navy : C.white }, line: { color: isFirst ? C.navy : C.border, width: 0.65 } });
          slide.addShape(pptx.ShapeType.ellipse,   { x: rx + 0.14, y: roadmapY + 0.13, w: 0.3, h: 0.3, fill: { color: isFirst ? C.goldLight : C.navy }, line: { color: isFirst ? C.goldLight : C.navy } });
          slide.addText(String(i + 1), { x: rx + 0.195, y: roadmapY + 0.185, w: 0.18, h: 0.14, fontSize: 7, bold: true, color: C.white, margin: 0, align: "center" });
          slide.addText(step.label, { x: rx + 0.52, y: roadmapY + 0.13, w: rmW - 0.62, h: 0.2,      fontSize: 9,   bold: true, color: isFirst ? C.white : C.navy, margin: 0, fit: "shrink" });
          slide.addText(step.body,  { x: rx + 0.14, y: roadmapY + 0.48, w: rmW - 0.22, h: rmH - 0.55, fontSize: 9, color: isFirst ? "C8D5E8" : C.text,           margin: 0, valign: "top", fit: "shrink" });
          if (i < rmSteps.length - 1) {
            slide.addText("›", { x: rx + rmW - 0.01, y: roadmapY + rmH / 2 - 0.14, w: 0.16, h: 0.28, fontSize: 16, bold: true, color: C.muted, align: "center", margin: 0 });
          }
        });
      }

      footer(slide);

      // Slide 4 Strategy Details
      // Removed redundant slide: 03 · KEY OUTPUTS

      if (includeConcentratedStockSlides && selectedStrategies.crt) {
      // Combined CRT slide: stat boxes + 4-step flow + concentration bars
      slide = pptx.addSlide();
      title(
        slide,
        "03 · STRATEGY DETAIL",
        "Charitable Remainder Trust",
        `Contribute appreciated ${data.ticker} shares into a trust to reduce concentration, generate income, and support charitable goals.`
      );

      // Stat boxes
      statBox(slide, 0.85, 1.78, 2.65, "Contribution", fmtM(data.crtAllocation), C.bluePale, C.blue);
      statBox(slide, 3.75, 1.78, 2.65, "Payout Rate", pct(data.crtPayoutRate), C.goldPale, C.gold);
      statBox(slide, 6.65, 1.78, 2.65, "Annual Income", fmtK(data.crtIncome), C.tealPale, C.teal);
      statBox(slide, 9.55, 1.78, 2.65, "Tax Deduction Est.", fmtM(data.charitableDeductionHigh), C.coralPale, C.coral);

      // 4-step flow
      const crtSteps = [
        { title: "1. Transfer Shares", body: `${fmtM(data.crtAllocation)} of ${data.ticker} → CRT`, fill: C.bluePale, color: C.blue },
        { title: "2. Trust Sells", body: "Diversifies tax-free inside the trust", fill: C.goldPale, color: C.gold },
        { title: "3. Income to Client", body: `${pct(data.crtPayoutRate)} payout = ~${fmtK(data.crtIncome)}/yr`, fill: C.tealPale, color: C.teal },
        { title: "4. Remainder to Charity", body: "Remaining assets to charity at term end", fill: C.coralPale, color: C.coral },
      ];
      const flowY = 3.25, boxW = 2.35, boxH = 1.3, gap = 0.55, startX = 0.95;
      crtSteps.forEach((step, i) => {
        const x = startX + i * (boxW + gap);
        slide.addShape(pptx.ShapeType.roundRect, { x, y: flowY, w: boxW, h: boxH, rectRadius: 0.08, fill: { color: step.fill }, line: { color: C.border, width: 0.8 } });
        slide.addText(step.title, { x: x + 0.12, y: flowY + 0.18, w: boxW - 0.24, h: 0.18, fontSize: 9, bold: true, color: C.navy, margin: 0, fit: "shrink" });
        slide.addText(step.body, { x: x + 0.12, y: flowY + 0.52, w: boxW - 0.24, h: 0.52, fontSize: 10, color: C.text, margin: 0, fit: "shrink" });
        if (i < crtSteps.length - 1) {
          slide.addShape(pptx.ShapeType.chevron, { x: x + boxW + 0.14, y: flowY + 0.46, w: 0.28, h: 0.38, fill: { color: C.gold }, line: { color: C.gold } });
        }
      });

      // Concentration reduction bars
      slide.addText("Concentration reduction", { x: 0.95, y: 4.82, w: 2.5, h: 0.18, fontSize: 9.5, bold: true, color: C.text, margin: 0 });

      slide.addText("Before", { x: 0.98, y: 5.12, w: 0.65, h: 0.14, fontSize: 7.8, color: C.muted, margin: 0 });
      slide.addShape(pptx.ShapeType.roundRect, { x: 1.7, y: 5.12, w: 5.2, h: 0.22, rectRadius: 0.08, fill: { color: C.lightBar }, line: { color: C.lightBar } });
      slide.addShape(pptx.ShapeType.roundRect, { x: 1.7, y: 5.12, w: 5.2 * Math.min(data.concentration / 100, 1), h: 0.22, rectRadius: 0.08, fill: { color: C.coral }, line: { color: C.coral } });
      slide.addText(pct(data.concentration), { x: 7.05, y: 5.12, w: 0.8, h: 0.14, fontSize: 8, bold: true, color: C.coral, margin: 0 });

      slide.addText("After", { x: 0.98, y: 5.52, w: 0.65, h: 0.14, fontSize: 7.8, color: C.muted, margin: 0 });
      slide.addShape(pptx.ShapeType.roundRect, { x: 1.7, y: 5.52, w: 5.2, h: 0.22, rectRadius: 0.08, fill: { color: C.lightBar }, line: { color: C.lightBar } });
      slide.addShape(pptx.ShapeType.roundRect, { x: 1.7, y: 5.52, w: 5.2 * Math.min(data.afterCrtConcentration / 100, 1), h: 0.22, rectRadius: 0.08, fill: { color: C.teal }, line: { color: C.teal } });
      slide.addText(pct(data.afterCrtConcentration), { x: 7.05, y: 5.52, w: 0.8, h: 0.14, fontSize: 8, bold: true, color: C.teal, margin: 0 });

      footer(slide);
      }


      // =========================
      if (includeConcentratedStockSlides && selectedStrategies.crt) {
      // CRT_LIFETIME_AND_COMPARISON_SLIDES
      // =========================
      {
        const moneyM = (v) => {
          const n = Number(v || 0);
          if (n >= 1000) return `$${(n / 1000).toFixed(1)}B`;
          if (n >= 1) return `$${n.toFixed(1)}M`;
          return `$${Math.round(n * 1000)}K`;
        };

        const cleanData = data;
        const stockValue = Number(cleanData.stockPosition || 0);
        const crtAmount = Number(cleanData.crtAllocation || stockValue * 0.32 || 0);
        const payoutRate = Number(cleanData.crtPayoutRate || 5);
        const growthRate = 7.5;
        const taxRate = Number(cleanData.totalTaxRate || cleanData.taxRate || 37.1);
        const deduction = Number(cleanData.charitableDeductionHigh || crtAmount * 0.37 || 0);
        const yr1Income = crtAmount * (payoutRate / 100);

        let balance = crtAmount;
        let cumulativeIncome = 0;
        const points = [];

        for (let yr = 1; yr <= 30; yr++) {
          const income = balance * (payoutRate / 100);
          cumulativeIncome += income;
          balance = Math.max(0, (balance - income) * (1 + growthRate / 100));

          if ([1, 5, 10, 15, 20, 25, 30].includes(yr)) {
            points.push({ yr, balance, cumulativeIncome });
          }
        }

        const endingBalance = points[points.length - 1]?.balance || balance;
        const totalIncome = points[points.length - 1]?.cumulativeIncome || cumulativeIncome;

        // Slide 2: CRT vs Outright Sale
        // Rebuilt as Beacon-style spreadsheet table from sample CRT proposal
        let compareSlide = pptx.addSlide();
        compareSlide.background = { color: "F7F8FB" };

        title(
          compareSlide,
          "04C · CRT COMPARISON",
          "Tax & Earnings Comparison of CRT vs. Outright Stock Sale",
          "A side-by-side view of taxes, capital available, and income potential."
        );

        const fmtWholeDollar = (v) => {
          const n = Math.round(Number(v || 0) * 1000000);
          return n.toLocaleString();
        };

        const fmtWholeDollarRaw = (v) => {
          const n = Math.round(Number(v || 0));
          return n.toLocaleString();
        };

        const federalRate = Number(cleanData.federalTaxRate || 37.0);
        const stateRate = Number(cleanData.stateTaxRate || 12.0);

        // Keep this similar to the sample table:
        // $5.0M CRT -> ~$1.104M federal deduction and ~$883K CA deduction.
        const federalDeduction = crtAmount * 0.22084;
        const stateDeduction = crtAmount * 0.176672;

        const federalRefund = federalDeduction * (federalRate / 100);
        const stateRefund = stateDeduction * (stateRate / 100);
        const refundTotal = federalRefund + stateRefund;

        const saleTaxPct = taxRate;
        const saleAfterTax = crtAmount * (1 - saleTaxPct / 100);
        const crtAvailableDay1 = crtAmount + refundTotal;

        const saleIncome = saleAfterTax * (payoutRate / 100);
        const crtIncome = crtAmount * (payoutRate / 100);
        const addedIncomeFromTaxSavings = refundTotal * (payoutRate / 100);
        const totalCrtIncome = crtIncome + addedIncomeFromTaxSavings;

        const tx = 0.8;
        const ty = 1.75;
        const tableW = 11.75;
        const headerH = 0.55;
        const rowH = 0.58;

        const labelW = 4.7;
        const sellDollarW = 0.35;
        const sellValueW = 2.9;
        const crtDollarW = 0.35;
        const crtValueW = 3.45;

        const xLabel = tx;
        const xSellDollar = xLabel + labelW;
        const xSellValue = xSellDollar + sellDollarW;
        const xCrtDollar = xSellValue + sellValueW;
        const xCrtValue = xCrtDollar + crtDollarW;

        const tableRows = [
          {
            label: "Amount contributed / sold",
            sellDollar: "$",
            sell: fmtWholeDollar(crtAmount),
            crtDollar: "$",
            crt: fmtWholeDollar(crtAmount),
            bold: true,
          },
          {
            label: "Upfront capital-gains tax",
            sellDollar: "",
            sell: `${saleTaxPct.toFixed(1)}% tax drag`,
            crtDollar: "",
            crt: "0.0% upfront tax",
          },
          {
            label: "Available to invest on Day 1",
            sellDollar: "$",
            sell: fmtWholeDollar(saleAfterTax),
            crtDollar: "$",
            crt: fmtWholeDollar(crtAvailableDay1),
            bold: true,
            green: true,
          },
          {
            label: "Estimated tax deduction",
            sellDollar: "$",
            sell: "-",
            crtDollar: "$",
            crt: fmtWholeDollar(federalDeduction + stateDeduction),
          },
          {
            label: "Estimated tax refund / savings",
            sellDollar: "$",
            sell: "-",
            crtDollar: "$",
            crt: fmtWholeDollar(refundTotal),
            bold: true,
          },
          {
            label: `Estimated Year 1 CRT payout`,
            sellDollar: "$",
            sell: fmtWholeDollar(saleIncome),
            crtDollar: "$",
            crt: fmtWholeDollar(totalCrtIncome),
            bold: true,
            green: true,
          },
        ];

        const tableH = headerH + tableRows.length * rowH;

        // Outer table border
        compareSlide.addShape(pptx.ShapeType.rect, {
          x: tx,
          y: ty,
          w: tableW,
          h: tableH,
          fill: { color: "FFFFFF" },
          line: { color: "000000", width: 1.05 },
        });

        // Header row
        compareSlide.addShape(pptx.ShapeType.rect, {
          x: tx,
          y: ty,
          w: tableW,
          h: headerH,
          fill: { color: "D9D9D9" },
          line: { color: "000000", width: 1 },
        });

        compareSlide.addText("Comparison of Stock Sale vs. CRT", {
          x: xLabel + 0.12,
          y: ty + 0.18,
          w: labelW - 0.2,
          h: 0.17,
          fontSize: 9.5,
          bold: true,
          color: "000000",
          margin: 0,
        });

        compareSlide.addText(`Sell ${cleanData.ticker || "Stock"}`, {
          x: xSellDollar,
          y: ty + 0.18,
          w: sellDollarW + sellValueW,
          h: 0.17,
          fontSize: 9.5,
          bold: true,
          color: "000000",
          align: "center",
          margin: 0,
        });

        compareSlide.addText("Fund CRT", {
          x: xCrtDollar,
          y: ty + 0.18,
          w: crtDollarW + crtValueW,
          h: 0.17,
          fontSize: 9.5,
          bold: true,
          color: "000000",
          align: "center",
          margin: 0,
        });

        // Vertical borders
        [xSellDollar, xCrtDollar].forEach((x) => {
          compareSlide.addShape(pptx.ShapeType.line, {
            x,
            y: ty,
            w: 0,
            h: tableH,
            line: { color: "000000", width: 1 },
          });
        });

        [xSellValue, xCrtValue].forEach((x) => {
          compareSlide.addShape(pptx.ShapeType.line, {
            x,
            y: ty + headerH,
            w: 0,
            h: tableH - headerH,
            line: { color: "000000", width: 0.75 },
          });
        });

        let y = ty + headerH;

        tableRows.forEach((r, i) => {
          const fill = r.green ? "DFF2D8" : i % 2 === 0 ? "FFFFFF" : "F7F7F7";

          compareSlide.addShape(pptx.ShapeType.rect, {
            x: tx,
            y,
            w: tableW,
            h: rowH,
            fill: { color: fill },
            line: { color: "CFCFCF", width: 0.35 },
          });

          compareSlide.addText(r.label, {
            x: xLabel + 0.12,
            y: y + 0.2,
            w: labelW - 0.25,
            h: 0.16,
            fontSize: 8.8,
            bold: !!r.bold,
            color: "000000",
            margin: 0,
            fit: "shrink",
          });

          compareSlide.addText(r.sellDollar, {
            x: xSellDollar + 0.08,
            y: y + 0.2,
            w: sellDollarW - 0.1,
            h: 0.16,
            fontSize: 8.8,
            color: "000000",
            margin: 0,
          });

          compareSlide.addText(r.sell, {
            x: xSellValue + 0.1,
            y: y + 0.2,
            w: sellValueW - 0.2,
            h: 0.16,
            fontSize: 8.8,
            bold: !!r.bold,
            color: "000000",
            align: "right",
            margin: 0,
            fit: "shrink",
          });

          compareSlide.addText(r.crtDollar, {
            x: xCrtDollar + 0.08,
            y: y + 0.2,
            w: crtDollarW - 0.1,
            h: 0.16,
            fontSize: 8.8,
            color: "000000",
            margin: 0,
          });

          compareSlide.addText(r.crt, {
            x: xCrtValue + 0.1,
            y: y + 0.2,
            w: crtValueW - 0.2,
            h: 0.16,
            fontSize: 8.8,
            bold: !!r.bold,
            color: "000000",
            align: "right",
            margin: 0,
            fit: "shrink",
          });

          y += rowH;
        });

        compareSlide.addText(
          "Illustrative only. Actual deduction, payout, taxation, and trust economics depend on trust design, AFR/7520 rate, client age, state tax rules, and CPA/legal review.",
          {
            x: 0.82,
            y: 6.68,
            w: 11.7,
            h: 0.18,
            fontSize: 5.8,
            color: "555555",
            margin: 0,
            fit: "shrink",
          }
        );

        footer(compareSlide);
      }
      }




      if (includeConcentratedStockSlides && selectedStrategies.harvesting) {
      // Slide 6 · Leveraged Tax-Loss Harvesting Detail
      // Modeled after enhanced long/short tax-loss harvesting strategy visual
      slide = pptx.addSlide();
      title(
        slide,
        "06 · TAX-LOSS HARVESTING",
        "Enhanced Long/Short Leveraged Tax-Loss Harvesting Strategy",
        "Modest leverage and shorting can help create a more durable stream of harvested tax losses."
      );

      // Dynamic 130/30 dollar amounts based on the client-specific harvesting sleeve
      const tlhLongCore = Number(data.harvestingSleeve || 0);
      const tlhLongExtension = tlhLongCore * 0.30;
      const tlhShortExtension = tlhLongCore * 0.30;


      statBox(slide, 0.85, 1.82, 2.25, "Harvesting Sleeve", fmtM(data.harvestingSleeve), C.white, C.navy);
      statBox(slide, 3.35, 1.82, 2.25, "Year 1 Losses", fmtM(data.annualHarvestLosses), C.tealPale, C.teal);
      statBox(slide, 5.85, 1.82, 2.25, "Tax Savings", fmtK(data.taxSavings), C.goldPale, C.gold);
      statBox(slide, 8.35, 1.82, 2.25, "Offset Goal", `${data.ticker} Gains`, C.bluePale, C.blue);

      slide.addText("Enhanced Tax-Optimized SMA", {
        x: 0.95,
        y: 3.0,
        w: 4.4,
        h: 0.2,
        fontSize: 12,
        bold: true,
        color: C.navy,
        margin: 0,
      });

      slide.addShape(pptx.ShapeType.line, {
        x: 0.95,
        y: 3.28,
        w: 4.5,
        h: 0,
        line: { color: C.navy, width: 1.2 },
      });

      const barX = 1.95;
      const barY = 3.72;
      const barW = 1.55;
      const longExtH = 0.52;
      const coreH = 1.25;
      const shortH = 0.48;

      slide.addShape(pptx.ShapeType.line, {
        x: 1.45,
        y: 3.55,
        w: 0,
        h: 2.35,
        line: { color: "8A8A8A", width: 0.8 },
      });

      slide.addShape(pptx.ShapeType.line, {
        x: 1.45,
        y: 5.28,
        w: 2.55,
        h: 0,
        line: { color: "8A8A8A", width: 0.8 },
      });

      slide.addShape(pptx.ShapeType.rect, {
        x: barX,
        y: barY,
        w: barW,
        h: longExtH,
        fill: { color: "159BE8" },
        line: { color: "159BE8" },
      });

      slide.addText(fmtM(tlhLongExtension), {
        x: barX,
        y: barY + 0.18,
        w: barW,
        h: 0.12,
        fontSize: 7.4,
        bold: true,
        color: C.white,
        align: "center",
        margin: 0,
      });

      slide.addShape(pptx.ShapeType.rect, {
        x: barX,
        y: barY + longExtH,
        w: barW,
        h: coreH,
        fill: { color: "101078" },
        line: { color: "101078" },
      });

      slide.addText(fmtM(tlhLongCore), {
        x: barX,
        y: barY + longExtH + 0.52,
        w: barW,
        h: 0.12,
        fontSize: 7.4,
        bold: true,
        color: C.white,
        align: "center",
        margin: 0,
      });

      slide.addShape(pptx.ShapeType.rect, {
        x: barX,
        y: barY + longExtH + coreH,
        w: barW,
        h: shortH,
        fill: { color: "4D44C6" },
        line: { color: "4D44C6" },
      });

      slide.addText(`-${fmtM(tlhShortExtension)}`, {
        x: barX,
        y: barY + longExtH + coreH + 0.16,
        w: barW,
        h: 0.12,
        fontSize: 7.4,
        bold: true,
        color: C.white,
        align: "center",
        margin: 0,
      });

      [
        { y: barY + 0.08, text: `Long\nExtension\n${fmtM(tlhLongExtension)}`, color: "159BE8" },
        { y: barY + longExtH + 0.45, text: `Long Equity\nPortfolio\n${fmtM(tlhLongCore)}`, color: "101078" },
        { y: barY + longExtH + coreH + 0.06, text: `Short\nExtension\n-${fmtM(tlhShortExtension)}`, color: "4D44C6" },
      ].forEach((a) => {
        slide.addShape(pptx.ShapeType.chevron, {
          x: barX + barW - 0.02,
          y: a.y,
          w: 0.35,
          h: 0.32,
          fill: { color: a.color },
          line: { color: a.color },
        });
        slide.addText(a.text, {
          x: barX + barW + 0.55,
          y: a.y - 0.02,
          w: 1.35,
          h: 0.32,
          fontSize: 7.6,
          bold: true,
          color: C.text,
          align: "center",
          margin: 0,
          fit: "shrink",
        });
      });

      slide.addShape(pptx.ShapeType.rect, {
        x: 6.45,
        y: 3.15,
        w: 5.55,
        h: 2.65,
        fill: { color: "F1F1F1" },
        line: { color: "F1F1F1" },
      });

      slide.addText("Why Add Short Positions?", {
        x: 6.75,
        y: 3.45,
        w: 4.7,
        h: 0.2,
        fontSize: 13,
        bold: true,
        color: C.navy,
        margin: 0,
      });

      slide.addText(
        "Short positions create additional harvesting opportunities — especially in rising markets where long-only portfolios produce fewer losses.",
        {
          x: 6.75,
          y: 3.88,
          w: 4.85,
          h: 0.48,
          fontSize: 10,
          color: C.text,
          margin: 0,
          fit: "shrink",
        }
      );

      slide.addText(
        "• More harvesting opportunities than long-only\n• Short-term losses offset taxable gains\n• Maintains market exposure + tax flexibility for future sales",
        {
          x: 6.85,
          y: 4.45,
          w: 4.75,
          h: 0.95,
          fontSize: 10,
          color: C.text,
          margin: 0,
          fit: "shrink",
        }
      );

      footer(slide);

      // Slide 7 · Direct Indexing & Leveraged Tax-Loss Harvesting
      // Modeled after the attached Direct Indexing & Leveraged Tax-Loss Harvesting visual
      slide = pptx.addSlide();
      title(
        slide,
        "06B · TAX-LOSS HARVESTING",
        "Direct Indexing & Leveraged Tax-Loss Harvesting",
        "A realized capital loss can become a planning asset because it may offset gains or carry forward into future years."
      );

      statBox(slide, 0.85, 1.82, 2.45, "Illustrative Losses", fmtM(data.annualHarvestLosses), C.white, C.navy);
      statBox(slide, 3.6, 1.82, 2.45, "Federal Savings", fmtK(data.federalTaxSavings), C.goldPale, C.gold);
      statBox(slide, 6.35, 1.82, 2.45, "State Savings", fmtK(data.stateTaxSavings), C.tealPale, C.teal);
      statBox(slide, 9.1, 1.82, 2.45, "Total Savings", fmtK(data.taxSavings), C.bluePale, C.blue);

      const tlhStart = Number(data.harvestingSleeve || 0);
      const tlhLoss = Number(data.annualHarvestLosses || 0);
      const tlhSaleValue = Math.max(tlhStart - tlhLoss, 0);
      const tlhRecoveryValue = Math.max(tlhStart * 1.03, tlhSaleValue * 1.24);

      const tlhMax = Math.max(tlhStart, tlhRecoveryValue) * 1.08;
      const tlhMin = Math.max(0, tlhSaleValue * 0.9);

      const yFor = (v) => {
        const top = 75;
        const bottom = 325;
        const span = tlhMax - tlhMin || 1;
        return bottom - ((v - tlhMin) / span) * (bottom - top);
      };

      const fmtChart = (v) => `$${Number(v || 0).toFixed(1)}M`;

      const yTicks = [
        tlhMax,
        tlhMin + (tlhMax - tlhMin) * 0.75,
        tlhMin + (tlhMax - tlhMin) * 0.5,
        tlhMin + (tlhMax - tlhMin) * 0.25,
        tlhMin,
      ];

      const declinePts = [
        [70, yFor(tlhStart * 0.995)],
        [115, yFor(tlhStart * 0.99)],
        [155, yFor(tlhStart * 0.97)],
        [190, yFor(tlhStart * 0.88)],
        [225, yFor(tlhStart * 0.80)],
        [260, yFor(tlhStart * 0.68)],
        [300, yFor(tlhSaleValue * 1.04)],
        [335, yFor(tlhSaleValue)],
      ];

      const recoverPts = [
        [335, yFor(tlhSaleValue)],
        [360, yFor(tlhSaleValue * 0.93)],
        [390, yFor(tlhSaleValue * 1.00)],
        [425, yFor(tlhSaleValue * 0.97)],
        [465, yFor(tlhSaleValue * 1.08)],
        [505, yFor(tlhSaleValue * 1.02)],
        [545, yFor(tlhSaleValue * 1.12)],
        [585, yFor(tlhSaleValue * 1.08)],
        [630, yFor(tlhSaleValue * 1.18)],
        [675, yFor(tlhSaleValue * 1.28)],
        [720, yFor(tlhSaleValue * 1.20)],
        [765, yFor(tlhRecoveryValue)],
        [820, yFor(tlhRecoveryValue * 0.98)],
        [860, yFor(tlhRecoveryValue * 1.01)],
      ];

      const declinePoints = declinePts.map(([x, y]) => `${x},${y}`).join(" ");
      const recoverPoints = recoverPts.map(([x, y]) => `${x},${y}`).join(" ");

      const tlhSvg = `
<svg xmlns="http://www.w3.org/2000/svg" width="980" height="420" viewBox="0 0 980 420">
  <rect x="0" y="0" width="980" height="420" fill="#FFFFFF"/>

  <text x="22" y="28" font-size="19" font-weight="700" fill="#1A2744">Illustrative Example of Leveraged Tax-Loss Harvesting</text>
  <text x="22" y="48" font-size="12" fill="#6E7E9A">Modeled using proposal assumptions</text>

  <rect x="70" y="75" width="790" height="250" fill="#FFFFFF" stroke="#D5DAE5" stroke-width="1.2"/>

  ${yTicks.map(v => `
    <line x1="70" y1="${yFor(v)}" x2="860" y2="${yFor(v)}" stroke="#E6EBF3" stroke-width="1"/>
    <text x="58" y="${yFor(v) + 4}" text-anchor="end" font-size="11" fill="#70809C">${fmtChart(v)}</text>
  `).join("")}

  <text x="70" y="348" font-size="11" fill="#70809C">Jan</text>
  <text x="180" y="348" font-size="11" fill="#70809C">Mar</text>
  <text x="320" y="348" font-size="11" fill="#70809C">Harvest</text>
  <text x="515" y="348" font-size="11" fill="#70809C">Reinvest</text>
  <text x="810" y="348" font-size="11" fill="#70809C">Recover</text>

  <polyline points="${declinePoints}" fill="none" stroke="#159BE8" stroke-width="4" stroke-linecap="round" stroke-linejoin="round"/>
  <polyline points="${recoverPoints}" fill="none" stroke="#101078" stroke-width="4" stroke-linecap="round" stroke-linejoin="round"/>

  <defs>
    <linearGradient id="arrowFade" x1="0" x2="0" y1="0" y2="1">
      <stop offset="0%" stop-color="#BFC7D8" stop-opacity="0.1"/>
      <stop offset="100%" stop-color="#3346C7" stop-opacity="0.95"/>
    </linearGradient>
  </defs>

  <polygon points="252,172 284,172 284,218 305,218 268,252 231,218 252,218" fill="url(#arrowFade)"/>

  <rect x="95" y="42" width="220" height="62" rx="2" fill="#FFFFFF" stroke="#DADFE8"/>
  <text x="205" y="67" text-anchor="middle" font-size="14" font-weight="600" fill="#222">Begin with ${fmtChart(tlhStart)}</text>
  <text x="205" y="88" text-anchor="middle" font-size="13" fill="#222">invested in Stock X</text>

  <rect x="300" y="155" width="230" height="90" rx="2" fill="#FFFFFF" stroke="#DADFE8"/>
  <text x="415" y="183" text-anchor="middle" font-size="13" font-weight="600" fill="#222">Sell Stock X at ${fmtChart(tlhSaleValue)}</text>
  <text x="415" y="207" text-anchor="middle" font-size="13" fill="#222">Realize ${fmtChart(tlhLoss)} capital loss</text>
  <text x="415" y="229" text-anchor="middle" font-size="13" fill="#222">Buy Stock Y with proceeds</text>

  <rect x="560" y="42" width="240" height="72" rx="2" fill="#FFFFFF" stroke="#DADFE8"/>
  <text x="680" y="70" text-anchor="middle" font-size="14" font-weight="600" fill="#222">Investor now has ${fmtChart(tlhLoss)}</text>
  <text x="680" y="92" text-anchor="middle" font-size="13" fill="#222">capital loss to offset gains</text>

  <text x="22" y="392" font-size="11" fill="#70809C">
    Illustrative path uses actual proposal assumptions: ${fmtChart(tlhStart)} harvesting sleeve and ${fmtChart(tlhLoss)} estimated harvested losses.
  </text>
</svg>
`;

      const tlhSvgData = "data:image/svg+xml;base64," + btoa(unescape(encodeURIComponent(tlhSvg)));

      slide.addImage({
        data: tlhSvgData,
        x: 0.75,
        y: 3.02,
        w: 7.65,
        h: 3.35,
      });

      card(
        slide,
        8.75,
        3.15,
        3.4,
        2.95,
        "A LOSS CAN BE AN ASSET",
        "• Realized losses may offset capital gains + limited ordinary income\n• Unused losses carry forward to future tax years\n• Client preserves market exposure after reinvestment",
        C.white
      );

      footer(slide);


      }

      // Slide 7 · Option Collar Detail
      if (includeConcentratedStockSlides && selectedStrategies.collar) {
      slide = pptx.addSlide();
      title(
        slide,
        "05 · STRATEGY DETAIL",
        `Option Collar on ${data.ticker}`,
        ""
      );

      slide.addText(
        "The payoff chart shows the collar structure: downside floor below the put, participation between the put and call, and capped upside above the call.",
        {
          x: 0.85,
          y: 1.38,
          w: 11.2,
          h: 0.28,
          fontSize: 10.2,
          color: C.text,
          margin: 0,
        }
      );

      statBox(slide, 0.85, 1.82, 2.05, "Protected Position", fmtM(data.collarAllocation), C.goldPale, C.gold);
      statBox(slide, 3.1,  1.82, 2.05, data.collarLive ? "Current Price" : "Prior Close Ref.", `$${data.stockPrice.toFixed(2)}`, C.white, C.navy);
      statBox(slide, 5.35, 1.82, 2.05, "Put Floor", `$${data.putStrike.toFixed(2)}`, C.coralPale, C.coral);
      statBox(slide, 7.6,  1.82, 2.05, "Call Cap", `$${data.callStrike.toFixed(2)}`, C.bluePale, C.blue);
      statBox(slide, 9.85, 1.82, 2.35, data.collarLive ? "Net Cost/Credit" : "Est. Net Cost", data.collarLive ? data.collarNetCostLabel?.replace(/Net cost:|Net credit:/, "").trim() || "—" : "See advisor", C.tealPale, C.teal);

      const collarShares = Math.round((data.collarAllocation * 1000000) / Math.max(data.stockPrice, 1));
      const minPrice = Math.max(1, data.putStrike * 0.6);
      const maxPrice = data.callStrike * 1.4;

      const xVals = [];
      for (let i = 0; i < 21; i++) {
        xVals.push(minPrice + ((maxPrice - minPrice) * i) / 20);
      }

      function underlyingVal(price) {
        return (price * collarShares) / 1000000;
      }

      function collarVal(price) {
        if (price <= data.putStrike) return (data.putStrike * collarShares) / 1000000;
        if (price >= data.callStrike) return (data.callStrike * collarShares) / 1000000;
        return (price * collarShares) / 1000000;
      }

      const yUnderlying = xVals.map(underlyingVal);
      const yCollar = xVals.map(collarVal);
      const yMax = Math.max(...yUnderlying, ...yCollar) * 1.12;

      function point(xv, yv) {
        const px = 80 + ((xv - minPrice) / (maxPrice - minPrice)) * 790;
        const py = 285 - (yv / yMax) * 220;
        return `${px},${py}`;
      }

      const underlyingPts = xVals.map((xv, i) => point(xv, yUnderlying[i])).join(" ");
      const collarPts = xVals.map((xv, i) => point(xv, yCollar[i])).join(" ");

      const putX = 80 + ((data.putStrike - minPrice) / (maxPrice - minPrice)) * 790;
      const spotX = 80 + ((data.stockPrice - minPrice) / (maxPrice - minPrice)) * 790;
      const callX = 80 + ((data.callStrike - minPrice) / (maxPrice - minPrice)) * 790;

      const yGrid = [0, 0.25, 0.5, 0.75, 1].map(r => {
        const val = yMax * r;
        const y = 285 - r * 220;
        return `
          <line x1="80" y1="${y}" x2="870" y2="${y}" stroke="#E5E8EF" stroke-width="1"/>
          <text x="68" y="${y + 4}" text-anchor="end" font-size="11" fill="#6E7E9A">${fmtM(val)}</text>
        `;
      }).join("");

      const xLabels = [minPrice, data.putStrike, data.stockPrice, data.callStrike, maxPrice].map(v => {
        const x = 80 + ((v - minPrice) / (maxPrice - minPrice)) * 790;
        return `<text x="${x}" y="318" text-anchor="middle" font-size="11" fill="#6E7E9A">$${Math.round(v)}</text>`;
      }).join("");

      const svg = `
        <svg xmlns="http://www.w3.org/2000/svg" width="960" height="350" viewBox="0 0 960 350">
          <rect width="960" height="350" fill="#FFFFFF"/>
          <rect x="80" y="65" width="${putX - 80}" height="220" fill="#FBEAE7" opacity="0.55"/>
          <rect x="${putX}" y="65" width="${callX - putX}" height="220" fill="#F5EDDA" opacity="0.55"/>
          <rect x="${callX}" y="65" width="${870 - callX}" height="220" fill="#E8EEF8" opacity="0.55"/>

          ${yGrid}

          <line x1="80" y1="285" x2="870" y2="285" stroke="#D5DAE5"/>
          <line x1="80" y1="65" x2="80" y2="285" stroke="#D5DAE5"/>

          <line x1="${putX}" y1="65" x2="${putX}" y2="285" stroke="#C94F3A" stroke-width="2" stroke-dasharray="6,4"/>
          <line x1="${spotX}" y1="65" x2="${spotX}" y2="285" stroke="#B8892A" stroke-width="2" stroke-dasharray="6,4"/>
          <line x1="${callX}" y1="65" x2="${callX}" y2="285" stroke="#3A6BBF" stroke-width="2" stroke-dasharray="6,4"/>

          <polyline points="${underlyingPts}" fill="none" stroke="#5D87B9" stroke-width="4" stroke-linecap="round" stroke-linejoin="round"/>
          <polyline points="${collarPts}" fill="none" stroke="#1A2744" stroke-width="4" stroke-linecap="round" stroke-linejoin="round"/>

          <text x="${putX}" y="52" text-anchor="middle" font-size="11" font-weight="700" fill="#C94F3A">Put $${data.putStrike.toFixed(2)}</text>
          <text x="${spotX}" y="40" text-anchor="middle" font-size="11" font-weight="700" fill="#B8892A">Spot $${data.stockPrice.toFixed(2)}</text>
          <text x="${callX}" y="52" text-anchor="middle" font-size="11" font-weight="700" fill="#3A6BBF">Call $${data.callStrike.toFixed(2)}</text>

          <line x1="625" y1="22" x2="665" y2="22" stroke="#5D87B9" stroke-width="4"/>
          <text x="675" y="27" font-size="12" fill="#1A2030">Underlying</text>
          <line x1="775" y1="22" x2="815" y2="22" stroke="#1A2744" stroke-width="4"/>
          <text x="825" y="27" font-size="12" fill="#1A2030">Protective Collar</text>

          ${xLabels}
        </svg>
      `;

      const svgData = "data:image/svg+xml;base64," + btoa(unescape(encodeURIComponent(svg)));
      slide.addImage({
        data: svgData,
        x: 0.85,
        y: 3.0,
        w: 7.35,
        h: 2.55,
      });

      card(
        slide,
        8.55,
        3.0,
        3.45,
        2.35,
        "Outcome Zones",
        `• Below $${data.putStrike.toFixed(2)}: downside protected by put\n• $${data.putStrike.toFixed(2)}–$${data.callStrike.toFixed(2)}: normal participation\n• Above $${data.callStrike.toFixed(2)}: upside capped by call\n${data.collarLive ? `• Live: ${data.collarExpiration || "current market"}` : `• Ref. price: prior close${data.priorCloseDate ? ` (${data.priorCloseDate})` : ""}`}`
      );

      footer(slide);
      }




      // Slide 8 — CRT Before / After, Cohesive Version
      // Removed redundant slide: 04B · CRT IMPACT

      // Removed redundant slide: 05B · 130/30 HARVESTING

      // Removed redundant slide: 06B · COLLAR OUTCOMES

      slide = pptx.addSlide();
      title(
        slide,
        "06 · INTEGRATED IMPACT",
        "Integrated Impact Dashboard",
        `A coordinated plan reduces concentration, creates tax capacity, generates income, and moderates downside risk.`
      );

      // Top row: 5 stat boxes — start at y:1.88 (clear of subtitle at y:1.78)
      const planConc = data.afterCrtConcentration || Math.max(data.concentration - 20, 10);

      // Combined tax savings across all strategies
      const taxRateFrac = (data.totalTaxRate || data.taxRate || 37.1) / 100;
      const crtEmbeddedGain = (data.crtAllocation || 0) * (1 - (data.costBasisPct || 15) / 100);
      const crtTaxAvoided   = crtEmbeddedGain * taxRateFrac;
      const crtDeductionBenefit = ((data.charitableDeductionLow || 0) + (data.charitableDeductionHigh || 0)) / 2 * taxRateFrac;
      const harvestingBenefit   = data.taxSavings || 0;
      const totalTaxBenefit     = crtTaxAvoided + crtDeductionBenefit + harvestingBenefit;

      const annualIncome = (data.crtIncome || 0);

      // Collar floor and cap in dollars
      const collarFloor   = (data.collarAllocation || 0) * 0.85;  // minimum value (put protects here)
      const collarCap     = (data.collarAllocation || 0) * 1.19;  // maximum value (call caps here)
      const downsideProtected = (data.collarAllocation || 0) - collarFloor; // max dollar loss avoided

      statBox(slide, 0.55, 1.88, 2.2, "Starting Concentration", pct(data.concentration), C.coralPale, C.coral);
      statBox(slide, 2.97, 1.88, 2.2, "Target Concentration", pct(planConc), C.tealPale, C.teal);
      statBox(slide, 5.39, 1.88, 2.2, "Combined Tax Savings", fmtM(totalTaxBenefit), C.goldPale, C.gold);
      statBox(slide, 7.81, 1.88, 2.2, "Annual Income (CRT)", annualIncome ? fmtK(annualIncome) : "N/A", C.bluePale, C.blue);
      statBox(slide, 10.23, 1.88, 2.2, "Collar Downside Floor", collarFloor > 0 ? fmtM(collarFloor) : "N/A", C.white, C.navy);

      // Concentration chart (left) — 0.1" gap below stat boxes (end at 2.78)
      addSvg(slide, makeConcentrationSvg(data), 0.55, 2.88, 6.6, 2.5);

      // Right side: Before/After comparison table
      const compY = 2.88;
      const compX = 7.5;
      const compW = 5.05;

      slide.addShape(pptx.ShapeType.roundRect, { x: compX, y: compY, w: compW, h: 2.5, rectRadius: 0.08, fill: { color: C.white }, line: { color: C.border, width: 0.7 } });
      slide.addShape(pptx.ShapeType.rect, { x: compX, y: compY, w: compW, h: 0.36, rectRadius: 0, fill: { color: C.navy }, line: { color: C.navy } });
      slide.addText("BEFORE VS. AFTER PLAN", { x: compX + 0.18, y: compY + 0.1, w: 2.2, h: 0.18, fontSize: 6.5, bold: true, color: C.goldLight, charSpace: 0.8, margin: 0 });
      slide.addText("WITHOUT", { x: compX + 2.5, y: compY + 0.1, w: 1.1, h: 0.18, fontSize: 6.5, bold: true, color: "8A9AB5", align: "center", margin: 0 });
      slide.addText("WITH PLAN", { x: compX + 3.7, y: compY + 0.1, w: 1.1, h: 0.18, fontSize: 6.5, bold: true, color: C.goldLight, align: "center", margin: 0 });

      const compRows = [
        ["Concentration", pct(data.concentration), pct(planConc)],
        ["Combined Tax Savings", "—", fmtM(totalTaxBenefit)],
        ["Annual CRT Income", "—", annualIncome ? fmtK(annualIncome) : "—"],
        ["Collar Floor / Cap", "—", collarFloor > 0 ? `${fmtM(collarFloor)} / ${fmtM(collarCap)}` : "—"],
      ];

      compRows.forEach((row, i) => {
        const ry = compY + 0.44 + i * 0.52;
        const bg = i % 2 === 0 ? "F7F8FB" : C.white;
        slide.addShape(pptx.ShapeType.rect, { x: compX + 0.01, y: ry, w: compW - 0.02, h: 0.5, fill: { color: bg }, line: { color: bg } });
        slide.addText(row[0], { x: compX + 0.18, y: ry + 0.13, w: 2.1, h: 0.22, fontSize: 10, color: C.text, margin: 0 });
        slide.addText(row[1], { x: compX + 2.5, y: ry + 0.13, w: 1.1, h: 0.22, fontSize: 10, bold: true, color: C.coral, align: "center", margin: 0 });
        slide.addText(row[2], { x: compX + 3.7, y: ry + 0.13, w: 1.1, h: 0.22, fontSize: 10, bold: true, color: C.teal, align: "center", margin: 0 });
      });

      // Bottom: three impact cards
      const impactCards = [
        {
          title: "Concentration Reduction",
          body: `${pct(data.concentration)} → ~${pct(planConc)} via CRT + staged diversification`,
          fill: C.coralPale, color: C.coral,
        },
        {
          title: "Combined Tax Savings",
          body: `CRT avoided: ${fmtM(crtTaxAvoided)}  ·  Deduction benefit: ${fmtM(crtDeductionBenefit)}  ·  Harvesting: ${fmtK(harvestingBenefit)}/yr`,
          fill: C.goldPale, color: C.gold,
        },
        {
          title: "Income & Collar Protection",
          body: collarFloor > 0
            ? `CRT income: ~${fmtK(annualIncome || 0)}/yr  ·  Floor: ${fmtM(collarFloor)}  ·  Cap: ${fmtM(collarCap)}`
            : `CRT income: ~${fmtK(annualIncome || 0)}/yr`,
          fill: C.tealPale, color: C.teal,
        },
      ];

      impactCards.forEach((ic, i) => {
        const icx = 0.55 + i * 3.93;
        slide.addShape(pptx.ShapeType.roundRect, { x: icx, y: 5.55, w: 3.73, h: 0.9, rectRadius: 0.07, fill: { color: ic.fill }, line: { color: C.border, width: 0.6 } });
        slide.addShape(pptx.ShapeType.rect, { x: icx, y: 5.55, w: 0.06, h: 0.9, fill: { color: ic.color }, line: { color: ic.color } });
        slide.addText(ic.title, { x: icx + 0.18, y: 5.60, w: 3.35, h: 0.18, fontSize: 8.5, bold: true, color: ic.color, margin: 0 });
        slide.addText(ic.body,  { x: icx + 0.18, y: 5.82, w: 3.35, h: 0.52, fontSize: 10, color: C.text, margin: 0, valign: "top", fit: "shrink" });
      });

      footer(slide);

      // ── Portfolio Transformation slide ──────────────────────────────────
      if (data.stockPosition > 0 && (selectedStrategies?.crt || selectedStrategies?.harvesting || selectedStrategies?.collar)) {
        slide = pptx.addSlide();
        title(
          slide,
          "07 · PORTFOLIO TRANSFORMATION",
          "Portfolio After Strategy Implementation",
          `How the portfolio is repositioned across CRT, harvesting sleeve, collared position, and remaining holdings.`
        );

        const crtAmt    = selectedStrategies?.crt       ? (data.crtAllocation    || 0) : 0;
        const sleeveAmt = selectedStrategies?.harvesting ? (data.harvestingSleeve || 0) : 0;
        const collarAmt = selectedStrategies?.collar     ? (data.collarAllocation || 0) : 0;
        const remStock  = Math.max((data.stockPosition || 0) - crtAmt - sleeveAmt - collarAmt, 0);
        const otherAmt  = Math.max((data.investableAssets || 0) - (data.stockPosition || 0), 0);
        const newInv    = (data.investableAssets || 0) - crtAmt;
        const newConc   = newInv > 0 ? ((remStock + collarAmt) / newInv * 100) : 0;

        // Stacked bar chart (left 7")
        addSvg(slide, makeAfterPortfolioSvg(data, selectedStrategies), 0.55, 1.88, 7.0, 3.5);

        // Summary table (right)
        const tX = 7.8, tY = 1.88, tW = 4.7;
        slide.addShape(pptx.ShapeType.roundRect, { x: tX, y: tY, w: tW, h: 4.5, rectRadius: 0.08, fill: { color: C.white }, line: { color: C.border, width: 0.7 } });
        slide.addShape(pptx.ShapeType.rect, { x: tX, y: tY, w: tW, h: 0.36, fill: { color: C.navy }, line: { color: C.navy } });
        slide.addText("PORTFOLIO BREAKDOWN — AFTER", { x: tX + 0.14, y: tY + 0.1, w: tW - 0.28, h: 0.18, fontSize: 6.5, bold: true, color: C.goldLight, charSpace: 0.8, margin: 0 });

        const rows = [
          ...(crtAmt    ? [["CRT → Charitable Trust",    fmtM(crtAmt),    `${pct(crtAmt    / (data.investableAssets||1) * 100)} of portfolio`]] : []),
          ...(sleeveAmt ? [["Harvesting Sleeve (130/30)", fmtM(sleeveAmt), `${pct(sleeveAmt / (data.investableAssets||1) * 100)} of portfolio`]] : []),
          ...(collarAmt ? [["Collared Position",          fmtM(collarAmt), `${pct(collarAmt / (data.investableAssets||1) * 100)} of portfolio`]] : []),
          ...(remStock  ? [["Remaining Stock",            fmtM(remStock),  `${pct(remStock  / (data.investableAssets||1) * 100)} of portfolio`]] : []),
                          ["Other Portfolio Assets",      fmtM(otherAmt),  `${pct(otherAmt  / (data.investableAssets||1) * 100)} of portfolio`],
        ];

        rows.forEach((row, i) => {
          const ry = tY + 0.44 + i * 0.68;
          const bg = i % 2 === 0 ? "F7F8FB" : C.white;
          slide.addShape(pptx.ShapeType.rect, { x: tX + 0.01, y: ry, w: tW - 0.02, h: 0.66, fill: { color: bg }, line: { color: bg } });
          slide.addText(row[0], { x: tX + 0.14, y: ry + 0.04, w: tW - 0.28, h: 0.2,  fontSize: 9,    color: C.text,  margin: 0, bold: false });
          slide.addText(row[1], { x: tX + 0.14, y: ry + 0.26, w: 1.6,       h: 0.22, fontSize: 12.5, color: C.navy,  margin: 0, bold: true });
          slide.addText(row[2], { x: tX + 1.9,  y: ry + 0.30, w: tW - 2.1,  h: 0.18, fontSize: 9,    color: "6B7A99",margin: 0 });
        });

        // Key metric below table
        const metY = tY + 0.44 + rows.length * 0.68 + 0.08;
        slide.addShape(pptx.ShapeType.roundRect, { x: tX + 0.14, y: metY, w: tW - 0.28, h: 0.72, rectRadius: 0.06, fill: { color: C.tealPale }, line: { color: C.teal, width: 0.7 } });
        slide.addText("NEW CONCENTRATION", { x: tX + 0.28, y: metY + 0.06, w: 2.2, h: 0.18, fontSize: 7.5, color: C.teal, bold: true, charSpace: 0.5, margin: 0 });
        slide.addText(pct(newConc), { x: tX + 0.28, y: metY + 0.28, w: 2.2, h: 0.32, fontSize: 18, bold: true, color: C.teal, margin: 0 });
        slide.addText(`was ${pct(data.concentration || 0)}`, { x: tX + 2.0, y: metY + 0.34, w: 1.8, h: 0.26, fontSize: 9, color: C.teal, margin: 0 });

        footer(slide);
      }

      // Slide 8 Disclosures
      
      // Removed Important Assumptions slide from client-facing deck.


      // CORE_PROPOSAL_MODULE_SLIDES removed to prevent duplicate early-deck slides.

      // DYNAMIC_PORTFOLIO_STRATEGY_SLIDES
      // Adds portfolio strategy slides only when advisor includes them.
      if (includePortfolioStrategySlides !== false) {
        // DYNAMIC_PORTFOLIO_STRATEGY_SLIDES
      // Adds two portfolio strategy slides based on selected model + risk profile.
      {
        const selectedPortfolioMap = selectedPortfolioStrategies || {};
        const selectedPortfolioKey =
          Object.keys(selectedPortfolioMap).find((key) => selectedPortfolioMap[key]) ||
          "corePrivate";

        const selectedRiskKey = selectedRiskProfile || "";

        const portfolioStrategyConfig = {
          corePrivate: {
            label: "Core Private",
            subtitle: "Traditional assets plus illiquid alternatives for larger qualified clients.",
            entails:
              "• Public equity + fixed income\n• Illiquid alternatives: private equity, real estate, hedge funds, private credit",
            why:
              "• Qualified, high-net-worth clients\n• Can tolerate reduced liquidity\n• Want diversification beyond stocks + bonds",
            allocations: {
              conservative: { equity: 14.25, fixedIncome: 68.75, alternatives: 17.00, cash: 0.00, growth: 21.00, stability: 79.00 },
              moderatelyConservative: { equity: 24.75, fixedIncome: 58.25, alternatives: 17.00, cash: 0.00, growth: 31.50, stability: 68.50 },
              conservativePlus: { equity: 35.25, fixedIncome: 47.75, alternatives: 17.00, cash: 0.00, growth: 42.00, stability: 58.00 },
              balanced: { equity: 44.75, fixedIncome: 37.25, alternatives: 18.00, cash: 0.00, growth: 52.50, stability: 47.50 },
              balancedPlus: { equity: 53.50, fixedIncome: 24.50, alternatives: 22.00, cash: 0.00, growth: 63.00, stability: 37.00 },
              growth: { equity: 64.00, fixedIncome: 14.00, alternatives: 22.00, cash: 0.00, growth: 73.50, stability: 26.50 },
              growthPlus: { equity: 74.50, fixedIncome: 3.00, alternatives: 22.50, cash: 0.00, growth: 84.00, stability: 16.00 },
              aggressive: { equity: 85.00, fixedIncome: 0.00, alternatives: 15.00, cash: 0.00, growth: 95.00, stability: 5.00 },
            },
          },

          selectLiquidity: {
            label: "Select Liquidity",
            subtitle: "Broad multi-asset allocation with interval-fund alternatives — no illiquid private equity.",
            entails:
              "• Equity + fixed income + semi-liquid alternatives\n• Interval funds: GRIFX, CCLFX, JHQDX, RAPIX\n• Real estate, private credit, hedged equity, real assets",
            why:
              "• Lower minimums, no capital-call structure\n• Quarterly redemption windows\n• Ideal for clients transitioning from concentrated positions",
            allocations: {
              conservative: { equity: 14.25, fixedIncome: 68.75, alternatives: 17.00, cash: 0.00, growth: 21.00, stability: 79.00 },
              moderatelyConservative: { equity: 24.75, fixedIncome: 58.25, alternatives: 17.00, cash: 0.00, growth: 31.50, stability: 68.50 },
              conservativePlus: { equity: 35.25, fixedIncome: 47.75, alternatives: 17.00, cash: 0.00, growth: 42.00, stability: 58.00 },
              balanced: { equity: 44.75, fixedIncome: 37.25, alternatives: 18.00, cash: 0.00, growth: 52.50, stability: 47.50 },
              balancedPlus: { equity: 53.50, fixedIncome: 24.50, alternatives: 22.00, cash: 0.00, growth: 63.00, stability: 37.00 },
              growth: { equity: 64.00, fixedIncome: 14.00, alternatives: 22.00, cash: 0.00, growth: 73.50, stability: 26.50 },
              growthPlus: { equity: 74.50, fixedIncome: 3.00, alternatives: 22.50, cash: 0.00, growth: 84.00, stability: 16.00 },
              aggressive: { equity: 85.00, fixedIncome: 0.00, alternatives: 15.00, cash: 0.00, growth: 95.00, stability: 5.00 },
            },
          },

          traditional: {
            label: "Traditional",
            subtitle: "Classic stock-and-bond allocation with no alternatives.",
            entails:
              "• Public equity + fixed income only\n• No alternatives allocation\n• Daily liquidity, transparent implementation",
            why:
              "• Clients prioritizing simplicity + liquidity\n• No alternatives eligibility required\n• Straightforward implementation",
            allocations: {
              conservative: { equity: 21.00, fixedIncome: 79.00, alternatives: 0.00, cash: 0.00, growth: 21.00, stability: 79.00 },
              moderatelyConservative: { equity: 31.50, fixedIncome: 68.50, alternatives: 0.00, cash: 0.00, growth: 31.50, stability: 68.50 },
              conservativePlus: { equity: 42.00, fixedIncome: 58.00, alternatives: 0.00, cash: 0.00, growth: 42.00, stability: 58.00 },
              balanced: { equity: 52.50, fixedIncome: 47.50, alternatives: 0.00, cash: 0.00, growth: 52.50, stability: 47.50 },
              balancedPlus: { equity: 63.00, fixedIncome: 37.00, alternatives: 0.00, cash: 0.00, growth: 63.00, stability: 37.00 },
              growth: { equity: 73.50, fixedIncome: 26.50, alternatives: 0.00, cash: 0.00, growth: 73.50, stability: 26.50 },
              growthPlus: { equity: 84.00, fixedIncome: 16.00, alternatives: 0.00, cash: 0.00, growth: 84.00, stability: 16.00 },
              aggressive: { equity: 100.00, fixedIncome: 0.00, alternatives: 0.00, cash: 0.00, growth: 100.00, stability: 0.00 },
            },
          },

          focusedB: {
            label: "Focused B",
            subtitle: "Traditional allocation with no U.S. Core equity and greater Growth/Value expression.",
            entails:
              "• Traditional stock + bond framework\n• No U.S. Core equity sleeve\n• Greater U.S. Value + Growth tilt",
            why:
              "• Advisor wants intentional equity factor tilt\n• Avoids broad U.S. Core blended exposure",
            allocations: {
              conservative: { equity: 21.00, fixedIncome: 79.00, alternatives: 0.00, cash: 0.00, growth: 21.00, stability: 79.00 },
              moderatelyConservative: { equity: 31.50, fixedIncome: 68.50, alternatives: 0.00, cash: 0.00, growth: 31.50, stability: 68.50 },
              conservativePlus: { equity: 42.00, fixedIncome: 58.00, alternatives: 0.00, cash: 0.00, growth: 42.00, stability: 58.00 },
              balanced: { equity: 52.50, fixedIncome: 47.50, alternatives: 0.00, cash: 0.00, growth: 52.50, stability: 47.50 },
              balancedPlus: { equity: 63.00, fixedIncome: 37.00, alternatives: 0.00, cash: 0.00, growth: 63.00, stability: 37.00 },
              growth: { equity: 73.50, fixedIncome: 26.50, alternatives: 0.00, cash: 0.00, growth: 73.50, stability: 26.50 },
              growthPlus: { equity: 84.00, fixedIncome: 16.00, alternatives: 0.00, cash: 0.00, growth: 84.00, stability: 16.00 },
              aggressive: { equity: 100.00, fixedIncome: 0.00, alternatives: 0.00, cash: 0.00, growth: 100.00, stability: 0.00 },
            },
          },

          selectLiquidityUsBias: {
            label: "Select Liquidity — U.S. Bias",
            subtitle: "Select Liquidity framework with a tilted domestic equity overweight.",
            entails:
              "• Select Liquidity structure + U.S. equity overweight\n• Domestic large/mid/small-cap overweighted\n• Same interval fund alternatives (GRIFX, CCLFX, JHQDX, RAPIX)",
            why:
              "• Clients preferring domestic equity exposure\n• Want Select Liquidity benefits with U.S. market orientation",
            allocations: null,
            fallbackKey: "selectLiquidity",
          },

          traditionalUsBias: {
            label: "Traditional — U.S. Bias",
            subtitle: "Traditional stock-and-bond framework with a greater domestic equity orientation.",
            entails:
              "• Traditional equity + fixed income framework\n• Greater U.S. equity orientation\n• No alternatives allocation",
            why:
              "• Simple, liquid implementation\n• Clients preferring U.S. market exposure",
            allocations: null,
            fallbackKey: "traditional",
          },
        };

        const selectedConfigRaw = portfolioStrategyConfig[selectedPortfolioKey] || portfolioStrategyConfig.corePrivate;
        const selectedConfig = selectedConfigRaw.allocations
          ? selectedConfigRaw
          : {
              ...selectedConfigRaw,
              allocations: portfolioStrategyConfig[selectedConfigRaw.fallbackKey || "traditional"].allocations,
              usingFallback: true,
            };

        const allocation =
          selectedRiskKey && selectedConfig.allocations[selectedRiskKey]
            ? selectedConfig.allocations[selectedRiskKey]
            : null;

        if (!allocation) {
          // No risk profile/allocation selected, so skip portfolio allocation slides.
        } else {

        const profileLabel =
          true
            ? riskProfileLabel
            : "Balanced — 50/50";

        function pctLabel(v) {
          return `${Number(v || 0).toFixed(1)}%`;
        }

        function piePath(cx, cy, r, startAngle, endAngle) {
          const start = polarToCartesian(cx, cy, r, endAngle);
          const end = polarToCartesian(cx, cy, r, startAngle);
          const largeArcFlag = endAngle - startAngle <= 180 ? "0" : "1";
          return [
            "M", cx, cy,
            "L", start.x, start.y,
            "A", r, r, 0, largeArcFlag, 0, end.x, end.y,
            "Z",
          ].join(" ");
        }

        function polarToCartesian(cx, cy, r, angleInDegrees) {
          const angleInRadians = (angleInDegrees - 90) * Math.PI / 180.0;
          return {
            x: cx + r * Math.cos(angleInRadians),
            y: cy + r * Math.sin(angleInRadians),
          };
        }

        function makeAllocationPieSvg(items) {
          const colors = ["#1A2744", "#3A6BBF", "#B8892A", "#8A98AD"];
          const total = items.reduce((sum, item) => sum + Math.max(0, item.value), 0) || 1;
          let currentAngle = 0;

          const slices = items
            .filter((item) => item.value > 0)
            .map((item, i) => {
              const angle = (item.value / total) * 360;
              const path = piePath(150, 150, 115, currentAngle, currentAngle + angle);
              currentAngle += angle;
              return `<path d="${path}" fill="${colors[i % colors.length]}" stroke="#FFFFFF" stroke-width="2"/>`;
            })
            .join("");

          const legend = items
            .map((item, i) => {
              const y = 42 + i * 34;
              return `
                <rect x="320" y="${y}" width="16" height="16" rx="3" fill="${colors[i % colors.length]}"/>
                <text x="346" y="${y + 13}" font-size="16" font-family="Arial" fill="#1A2030">${item.label}: ${pctLabel(item.value)}</text>
              `;
            })
            .join("");

          return `
            <svg xmlns="http://www.w3.org/2000/svg" width="620" height="320" viewBox="0 0 620 320">
              <rect width="620" height="320" fill="#FFFFFF"/>
              ${slices}
              <circle cx="150" cy="150" r="58" fill="#FFFFFF"/>
              <text x="150" y="142" text-anchor="middle" font-size="15" font-family="Arial" font-weight="700" fill="#1A2744">Model</text>
              <text x="150" y="165" text-anchor="middle" font-size="15" font-family="Arial" font-weight="700" fill="#1A2744">Allocation</text>
              ${legend}
            </svg>
          `;
        }

        function addAllocationBar(slide, x, y, w, label, value, fillColor) {
          // Label — narrowed to 1.5" so bars + pct labels stay in the left column
          slide.addText(label, {
            x,
            y,
            w: 1.5,
            h: 0.22,
            fontSize: 11,
            bold: true,
            color: C.text,
            margin: 0,
          });

          // Track background — starts at x+1.65
          slide.addShape(pptx.ShapeType.rect, {
            x: x + 1.65,
            y: y + 0.04,
            w,
            h: 0.14,
            fill: { color: "E8EDF5" },
            line: { color: "E8EDF5" },
          });

          // Filled bar
          slide.addShape(pptx.ShapeType.rect, {
            x: x + 1.65,
            y: y + 0.04,
            w: w * Math.max(0, Math.min(100, value)) / 100,
            h: 0.14,
            fill: { color: fillColor },
            line: { color: fillColor },
          });

          // Pct label — sits just after the bar, well left of x:7.1
          slide.addText(pctLabel(value), {
            x: x + 1.65 + w + 0.1,
            y,
            w: 0.6,
            h: 0.22,
            fontSize: 11,
            bold: true,
            color: C.navy,
            margin: 0,
          });
        }

        // ── REAL FUND DATA ──────────────────────────────────────────────
        // Apply any advisor fund substitutions chosen in the preview modal
        // (keyed by the original recommended fund's `name`) — swaps change
        // the displayed name/ticker/fee but never the allocation %, so
        // groupTotals/subGroupTotals (used for the donut + bars) don't need
        // to change.
        const realFunds = getFunds(selectedPortfolioKey, selectedRiskKey).map(f => {
          const swap = fundSwaps[f.name];
          if (!swap) return f;
          return {
            ...f,
            name: swap.name,
            fullName: swap.name,
            ticker: swap.ticker || f.ticker,
            fee: Number.isFinite(swap.fee) ? swap.fee : f.fee,
            swappedFrom: f.name,
          };
        });
        const groupTotals    = getGroupTotals(selectedPortfolioKey, selectedRiskKey);
        const subGroupTotals = getSubGroupTotals(selectedPortfolioKey, selectedRiskKey);
        const hasAnySwap     = realFunds.some(f => f.swappedFrom);
        const equityPct      = groupTotals["Equity"]       || 0;
        const fixedIncPct    = groupTotals["Fixed Income"] || 0;
        const altsPct        = groupTotals["Alternatives"] || 0;

        // 6-segment donut SVG — viewBox 530×444 matches slide ratio (4.6 × 3.85)
        // so pptxgenjs won't distort the circle when stretching to fill w/h.
        // Segments: Domestic Equity · Intl Developed · Emerging Markets ·
        //           Fixed Income · Alternatives · Cash
        function makePortfolioDonutSvg(sgt) {
          const SEGMENTS = [
            { key: "Domestic Equity",  color: "#1A2744" },  // deep navy
            { key: "Intl Developed",   color: "#1480C8" },  // vivid sky blue
            { key: "Emerging Markets", color: "#27A066" },  // medium green
            { key: "Fixed Income",     color: "#7056A0" },  // muted purple
            { key: "Alternatives",     color: "#C4872A" },  // amber gold
            { key: "Cash",             color: "#9BAAB8" },  // cool gray
          ];

          // Only include segments with allocation > 0.1%; normalize to 100% so no gap
          const rawParts = SEGMENTS
            .map(s => ({ label: s.key, pct: sgt[s.key] || 0, color: s.color }))
            .filter(p => p.pct > 0.1);
          const rawTotal  = rawParts.reduce((s, p) => s + p.pct, 0) || 100;
          const parts     = rawParts.map(p => ({ ...p, pct: (p.pct / rawTotal) * 100 }));

          // Center label: total equity % (re-derived from normalized parts)
          const totalEquityNorm = parts
            .filter(p => ["Domestic Equity","Intl Developed","Emerging Markets"].includes(p.label))
            .reduce((s, p) => s + p.pct, 0);
          // Original (unnormalized) equity total for the label number shown to the advisor
          const totalEquity = (sgt["Domestic Equity"] || 0) + (sgt["Intl Developed"] || 0) + (sgt["Emerging Markets"] || 0);

          const cx = 155, cy = 222, outerR = 155, innerR = 74;
          let angle = 0;
          const segs = parts.map(p => {
            const sweep = (p.pct / 100) * 360;
            const s = donutSegment(cx, cy, outerR, innerR, angle, angle + sweep, p.color);
            angle += sweep;
            return s;
          }).join("");

          // Legend: right side, 2 items per row to fit 6 entries cleanly
          const legendItems = parts.map((p, i) => {
            const col  = i % 2;           // 0 = left col, 1 = right col
            const row  = Math.floor(i / 2);
            const lx   = 330 + col * 96;
            const ly   = 42 + row * 72;
            return `
              <rect x="${lx}" y="${ly}" width="14" height="14" rx="2" fill="${p.color}"/>
              <text x="${lx + 18}" y="${ly + 11}" font-family="Inter, Arial, sans-serif" font-size="13" font-weight="700" fill="#1A2744">${p.pct.toFixed(1)}%</text>
              <text x="${lx + 18}" y="${ly + 24}" font-family="Inter, Arial, sans-serif" font-size="10" fill="#6E7E9A">${p.label}</text>`;
          }).join("");

          return `<svg xmlns="http://www.w3.org/2000/svg" width="530" height="444" viewBox="0 0 530 444">
            <rect width="530" height="444" fill="#FFFFFF"/>
            ${segs}
            <circle cx="${cx}" cy="${cy}" r="${innerR - 4}" fill="#FFFFFF"/>
            <text x="${cx}" y="${cy - 10}" text-anchor="middle" font-family="Georgia, serif" font-size="32" font-weight="700" fill="#1A2744">${totalEquity.toFixed(1)}%</text>
            <text x="${cx}" y="${cy + 20}" text-anchor="middle" font-family="Inter, Arial, sans-serif" font-size="13" fill="#6E7E9A">Total Equity</text>
            ${legendItems}
          </svg>`;
        }

        // ── SLIDE 1: PORTFOLIO STRATEGY OVERVIEW ────────────────────────
        slide = pptx.addSlide();
        title(
          slide,
          "PORTFOLIO STRATEGY",
          `${selectedConfig.label} — Allocation Overview`,
          `${profileLabel} risk profile · ${realFunds.length} active fund positions`
        );

        statBox(slide, 0.85, 1.78, 2.85, "Strategy Model",   selectedConfig.label,                       C.white,    C.navy);
        statBox(slide, 3.90, 1.78, 2.85, "Risk Profile",     String(profileLabel).split("—")[0].trim(),   C.goldPale, C.gold);
        statBox(slide, 6.95, 1.78, 2.85, "Equity",           pctLabel(equityPct),                         C.tealPale, C.teal);
        statBox(slide, 9.95, 1.78, 3.00, "Fixed Income",     pctLabel(fixedIncPct),                       C.bluePale, C.blue);

        // Donut
        const overviewDonutSvg = makePortfolioDonutSvg(subGroupTotals);
        addSvg(slide, overviewDonutSvg, 0.85, 2.95, 4.6, 3.85);

        // Fund table — right side
        const tblX      = 5.75;
        const tblW      = 7.05;
        const colFund   = tblX;
        const colAC     = tblX + 4.15;
        const colAlloc  = tblX + 6.15;
        const hdrY      = 3.0;
        const hdrOpts   = { fontSize: 9.5, bold: true, color: C.navy, margin: 0 };
        const sepColor  = "D0D8E8";

        // Header
        slide.addText("Fund / Manager",  { x: colFund,  y: hdrY, w: 4.0,  h: 0.24, ...hdrOpts });
        slide.addText("Asset Class",     { x: colAC,    y: hdrY, w: 1.85, h: 0.24, ...hdrOpts });
        slide.addText("Alloc",           { x: colAlloc, y: hdrY, w: 0.75, h: 0.24, ...hdrOpts, align: "right" });
        slide.addShape(pptx.ShapeType.rect, {
          x: colFund, y: hdrY + 0.26, w: tblW, h: 0.015, fill: { color: C.navy }, line: { color: C.navy }
        });

        // Sub-group colors for row tints — splits Equity into Domestic / Intl
        // Developed / Emerging Markets to match the 6-segment donut above
        // (and the preview modal's fund table) instead of lumping all equity
        // funds under one "Equity" tint.
        const subGroupFill = {
          "Domestic Equity": "EEF2FA", "Intl Developed": "E7F1FA", "Emerging Markets": "E8F6EE",
          "Fixed Income": "EDE8F5", Alternatives: "FBF3E4", Cash: "F5F5F5",
        };
        const subGroupColor = {
          "Domestic Equity": C.navy, "Intl Developed": "0F6097", "Emerging Markets": "1C7A4D",
          "Fixed Income": "5040A0", Alternatives: C.gold, Cash: C.muted,
        };

        const rowH   = Math.min(0.235, 3.35 / Math.max(realFunds.length, 1));
        let rowY     = hdrY + 0.31;
        let rowCount = 0;

        for (const f of realFunds) {
          if (rowY + rowH > 6.75) break;
          const fill = subGroupFill[f.subGroup] || "FFFFFF";
          const gc   = subGroupColor[f.subGroup] || C.text;

          slide.addShape(pptx.ShapeType.rect, {
            x: colFund, y: rowY, w: tblW, h: rowH,
            fill: { color: rowCount % 2 === 0 ? fill : "FFFFFF" },
            line: { color: "FFFFFF", width: 0 },
          });
          slide.addText(f.swappedFrom ? `${f.name} *` : f.name, {
            x: colFund + 0.06, y: rowY + 0.01, w: 3.95, h: rowH,
            fontSize: 9, color: C.text, bold: false, margin: 0, valign: "middle", fit: "shrink",
          });
          slide.addText(f.assetClass, {
            x: colAC + 0.04, y: rowY + 0.01, w: 1.8, h: rowH,
            fontSize: 8.5, color: gc, bold: false, margin: 0, valign: "middle", fit: "shrink",
          });
          slide.addText(pctLabel(f.alloc), {
            x: colAlloc, y: rowY + 0.01, w: 0.72, h: rowH,
            fontSize: 9, color: C.navy, bold: true, margin: 0, valign: "middle", align: "right",
          });

          rowY += rowH;
          rowCount++;
        }

        // Bottom separator
        slide.addShape(pptx.ShapeType.rect, {
          x: colFund, y: rowY, w: tblW, h: 0.015, fill: { color: sepColor }, line: { color: sepColor }
        });

        if (hasAnySwap) {
          slide.addText("* Advisor-selected fund substitution", {
            x: colFund, y: rowY + 0.04, w: tblW, h: 0.16,
            fontSize: 7, italic: true, color: C.muted, margin: 0,
          });
        }

        footer(slide);

        // ── SLIDE 2: PORTFOLIO STRATEGY DETAIL ──────────────────────────
        slide = pptx.addSlide();
        title(
          slide,
          "PORTFOLIO STRATEGY DETAIL",
          `${selectedConfig.label}: Fund-Level Allocation`,
          `${profileLabel} · Individual fund positions, tickers, and fees`
        );

        // Left: fund detail grouped by sub-category — splits Equity into
        // Domestic / Intl Developed / Emerging Markets to match Slide 1 and
        // the preview modal's fund table.
        const GROUPS = ["Domestic Equity", "Intl Developed", "Emerging Markets", "Fixed Income", "Alternatives", "Cash"];
        const groupedFunds = {};
        for (const g of GROUPS) {
          groupedFunds[g] = realFunds.filter(f => f.subGroup === g);
        }

        const dtlX     = 0.85;
        const dtlW     = 8.0;
        const dtlColF  = dtlX;
        const dtlColT  = dtlX + 4.65;
        const dtlColA  = dtlX + 5.6;
        const dtlColFe = dtlX + 6.5;
        const hdrOpts2 = { fontSize: 9, bold: true, color: C.navy, margin: 0 };

        // Column headers
        const dtlHdrY = 2.02;
        slide.addText("Fund / Manager",   { x: dtlColF  + 0.05, y: dtlHdrY, w: 4.55, h: 0.24, ...hdrOpts2 });
        slide.addText("Ticker",           { x: dtlColT  + 0.03, y: dtlHdrY, w: 0.85, h: 0.24, ...hdrOpts2 });
        slide.addText("Alloc",            { x: dtlColA  + 0.03, y: dtlHdrY, w: 0.8,  h: 0.24, ...hdrOpts2, align: "right" });
        slide.addText("Fee",              { x: dtlColFe + 0.03, y: dtlHdrY, w: 0.7,  h: 0.24, ...hdrOpts2, align: "right" });
        slide.addShape(pptx.ShapeType.rect, {
          x: dtlX, y: dtlHdrY + 0.26, w: dtlW, h: 0.015, fill: { color: C.navy }, line: { color: C.navy }
        });

        const groupTitleColor = {
          "Domestic Equity": C.navy, "Intl Developed": "0F6097", "Emerging Markets": "1C7A4D",
          "Fixed Income": C.blue, Alternatives: C.gold, Cash: C.muted,
        };
        const groupTitleFill  = {
          "Domestic Equity": "EEF2FA", "Intl Developed": "E0EFFA", "Emerging Markets": "E8F6EE",
          "Fixed Income": "E8F0FF", Alternatives: "FAF3E0", Cash: "F5F5F5",
        };

        let dtlY = dtlHdrY + 0.30;
        let altCount = 0;

        for (const group of GROUPS) {
          const gFunds = groupedFunds[group];
          if (!gFunds || gFunds.length === 0) continue;

          const gc2   = groupTitleColor[group] || C.navy;
          const gfill = groupTitleFill[group]  || "F5F5F5";
          const gTot  = gFunds.reduce((s, f) => s + f.alloc, 0);

          // Section header
          slide.addShape(pptx.ShapeType.rect, {
            x: dtlX, y: dtlY, w: dtlW, h: 0.28,
            fill: { color: gfill }, line: { color: gfill }
          });
          slide.addText(`${group.toUpperCase()}`, {
            x: dtlX + 0.08, y: dtlY + 0.01, w: 3.5, h: 0.26,
            fontSize: 9, bold: true, color: gc2, margin: 0, valign: "middle",
          });
          slide.addText(pctLabel(gTot), {
            x: dtlColA, y: dtlY + 0.01, w: 0.85, h: 0.26,
            fontSize: 9, bold: true, color: gc2, margin: 0, valign: "middle", align: "right",
          });
          dtlY += 0.28;

          // Fund rows
          const fRowH = Math.min(0.24, 1.4 / Math.max(gFunds.length, 1));
          for (const f of gFunds) {
            if (dtlY + fRowH > 6.7) break;
            const even = altCount % 2 === 0;
            slide.addShape(pptx.ShapeType.rect, {
              x: dtlX, y: dtlY, w: dtlW, h: fRowH,
              fill: { color: even ? "F8F9FC" : "FFFFFF" }, line: { color: "FFFFFF", width: 0 }
            });
            slide.addText(f.swappedFrom ? `${f.name} *` : f.name, {
              x: dtlColF + 0.1, y: dtlY, w: 4.45, h: fRowH,
              fontSize: 8.5, color: C.text, margin: 0, valign: "middle", fit: "shrink",
            });
            slide.addText(f.ticker !== "N/A" ? f.ticker : "—", {
              x: dtlColT + 0.04, y: dtlY, w: 0.82, h: fRowH,
              fontSize: 8, color: C.muted, margin: 0, valign: "middle",
            });
            slide.addText(pctLabel(f.alloc), {
              x: dtlColA, y: dtlY, w: 0.82, h: fRowH,
              fontSize: 8.5, bold: true, color: C.navy, margin: 0, valign: "middle", align: "right",
            });
            slide.addText(f.fee > 0 ? `${f.fee.toFixed(2)}%` : "—", {
              x: dtlColFe, y: dtlY, w: 0.72, h: fRowH,
              fontSize: 8, color: C.muted, margin: 0, valign: "middle", align: "right",
            });
            dtlY += fRowH;
            altCount++;
          }

          // Thin gap between groups
          dtlY += 0.05;
        }

        // Right side: info cards
        const cardX  = 9.1;
        const cardW  = 4.0;
        card(slide, cardX, 2.02, cardW, 1.45,
          "Client Fit",
          `• ${selectedConfig.label} paired with ${profileLabel}\n• Post-diversification allocation framework\n• Addresses tax, liquidity, and risk objectives`,
          C.white);

        card(slide, cardX, 3.65, cardW, 1.35,
          "Advisor Review Points",
          `• Confirm liquidity needs + restrictions\n• Verify alternatives eligibility\n• Validate risk tolerance + tax constraints`,
          C.bluePale);

        const isSelectLiq2 = selectedConfig.label.includes("Select Liquidity");
        if (isSelectLiq2 || selectedPortfolioKey === "corePrivate") {
          const altFunds = realFunds.filter(f => f.group === "Alternatives");
          const altBullets = altFunds.length > 0
            ? altFunds.map(f => `• ${f.ticker !== "N/A" ? f.ticker : f.name.split(" ")[0]} — ${f.name}`).join("\n")
            : "• See alternatives sleeve for detail";
          card(slide, cardX, 5.18, cardW, 1.55,
            "Alternative Holdings",
            altBullets,
            C.goldPale);
        } else {
          card(slide, cardX, 5.18, cardW, 0.85,
            "Implementation Note",
            "• Validate against current firm model guidance\n• Confirm client-specific suitability before finalizing",
            C.goldPale);
        }

        if (hasAnySwap) {
          slide.addText("* Advisor-selected fund substitution", {
            x: dtlX, y: 6.55, w: dtlW, h: 0.16,
            fontSize: 7, italic: true, color: C.muted, margin: 0,
          });
        }

        footer(slide);

        // ── SLIDE 3: PORTFOLIO TRANSITION ANALYSIS ──────────────────────
        // Backtested current-vs-target comparison computed in App.jsx
        // (see runPortfolioBacktest / backtest.js) using real historical
        // monthly prices — never fabricated numbers. Mirrors the "Portfolio
        // Transition Analysis" section shown in the preview modal so the
        // advisor and client see the same figures here.
        {
          const bt = backtest;
          const hasBacktest = bt && !bt.note && bt.target?.summary;

          slide = pptx.addSlide();
          title(
            slide,
            "PORTFOLIO STRATEGY",
            "Portfolio Transition Analysis",
            hasBacktest && bt.target.summary.startDate
              ? `Backtested ${bt.target.summary.startDate} – ${bt.target.summary.endDate} · real historical monthly prices`
              : "Current vs. recommended portfolio — historical comparison"
          );

          if (!hasBacktest) {
            card(slide, 0.85, 2.0, 11.6, 1.6,
              "Backtest Not Available",
              bt?.note || "A historical backtest could not be computed for this proposal (e.g. insufficient publicly-tradable tickers in the model). Run the proposal through the preview screen to populate this analysis.",
              C.white);
            footer(slide);
          } else {
            const t  = bt.target;
            const c  = bt.current;
            const ts = t.summary;
            const cs = c?.summary;
            const hasOverlap = !!(ts?.months && cs?.months);

            const pctF = (v, d = 1) => (v == null ? "—" : `${(v * 100).toFixed(d)}%`);
            const feeF = (v) => (v == null ? "—" : `${v.toFixed(2)}%`);
            const growth10k = (cumRet) => (cumRet == null ? "—" : `$${Math.round(10000 * (1 + cumRet)).toLocaleString("en-US")}`);

            const panelW = c ? 5.6 : 11.6;
            const panelX1 = 0.85;
            const panelX2 = 6.85;
            const panelY = 2.0;
            const panelH = 2.95;

            // Target (recommended) panel
            card(slide, panelX1, panelY, panelW, panelH, "Recommended (Target) Portfolio", "", C.white);
            statBox(slide, panelX1 + 0.22, panelY + 0.62, (panelW - 0.66) / 2, "Annualized Return (CAGR)", pctF(ts.annualizedReturn), C.tealPale, C.teal);
            statBox(slide, panelX1 + 0.22 + (panelW - 0.66) / 2 + 0.22, panelY + 0.62, (panelW - 0.66) / 2, "Annualized Volatility", pctF(ts.annualizedVolatility), C.white, C.navy);
            statBox(slide, panelX1 + 0.22, panelY + 1.46, (panelW - 0.66) / 2, "Max Drawdown", pctF(ts.maxDrawdown), C.white, C.coral);
            statBox(slide, panelX1 + 0.22 + (panelW - 0.66) / 2 + 0.22, panelY + 1.46, (panelW - 0.66) / 2, "Growth of $10,000", growth10k(ts.cumulativeReturn), C.tealPale, C.teal);
            slide.addText(
              `Weighted avg. fee: ${feeF(t.weightedFeePct)}` +
                (t.coveragePct < 100 ? `  ·  reflects ${t.coveragePct.toFixed(0)}% of model by allocation` : ""),
              { x: panelX1 + 0.22, y: panelY + 2.30, w: panelW - 0.44, h: 0.5, fontSize: 7.5, color: C.muted, margin: 0, fit: "shrink" }
            );

            // Current portfolio panel — from uploaded holdings or concentrated-position approximation
            if (c) {
              const currentTitle = c.fromUploadedHoldings ? "Current Portfolio" : "Current Portfolio (Approximate)";
              const currentNote = c.fromUploadedHoldings
                ? `Weighted avg. fee: ${feeF(c.weightedFeePct)}  ·  from ${c.holdingCount || "uploaded"} holdings`
                : `Weighted avg. fee: ${feeF(c.weightedFeePct)}  ·  approx. ${(c.concentration || 0).toFixed(1)}% ${c.ticker} + ${(100 - (c.concentration || 0)).toFixed(1)}% ${bt.benchmarkTicker}`;
              card(slide, panelX2, panelY, panelW, panelH, currentTitle, "", C.bluePale);
              statBox(slide, panelX2 + 0.22, panelY + 0.62, (panelW - 0.66) / 2, "Annualized Return (CAGR)", pctF(cs?.annualizedReturn), C.white, C.navy);
              statBox(slide, panelX2 + 0.22 + (panelW - 0.66) / 2 + 0.22, panelY + 0.62, (panelW - 0.66) / 2, "Annualized Volatility", pctF(cs?.annualizedVolatility), C.white, C.navy);
              statBox(slide, panelX2 + 0.22, panelY + 1.46, (panelW - 0.66) / 2, "Max Drawdown", pctF(cs?.maxDrawdown), C.white, C.coral);
              statBox(slide, panelX2 + 0.22 + (panelW - 0.66) / 2 + 0.22, panelY + 1.46, (panelW - 0.66) / 2, "Growth of $10,000", growth10k(cs?.cumulativeReturn), C.white, C.navy);
              slide.addText(
                currentNote,
                { x: panelX2 + 0.22, y: panelY + 2.30, w: panelW - 0.44, h: 0.5, fontSize: 7.5, color: C.muted, margin: 0, fit: "shrink" }
              );
            } else {
              slide.addText(
                "No current-vs-target comparison shown: no concentrated stock position / concentration % was found in this client's notes. Add one to enable a side-by-side transition comparison.",
                { x: panelX1, y: panelY + panelH + 0.18, w: 11.6, h: 0.4, fontSize: 9, italic: true, color: C.muted, margin: 0, fit: "shrink" }
              );
            }

            // Transition summary deltas — use CAGR (consistent with volatility window)
            if (c && hasOverlap) {
              const retDelta = (ts.annualizedReturn ?? 0) - (cs?.annualizedReturn ?? 0);
              const volDelta = (ts.annualizedVolatility ?? 0) - (cs?.annualizedVolatility ?? 0);
              const feeDelta = (t.weightedFeePct ?? 0) - (c.weightedFeePct ?? 0);
              const summaryBody =
                `• ${volDelta <= 0 ? "Lower" : "Higher"} historical volatility: ${pctF(Math.abs(volDelta))} ${volDelta <= 0 ? "reduction" : "increase"} vs. current portfolio\n` +
                `• ${retDelta >= 0 ? "Higher" : "Lower"} historical annualized return: ${pctF(Math.abs(retDelta))} ${retDelta >= 0 ? "above" : "below"} current portfolio\n` +
                `• ${feeDelta <= 0 ? "Lower" : "Higher"} weighted fee: ${Math.abs(feeDelta).toFixed(2)}% ${feeDelta <= 0 ? "savings" : "increase"} vs. current portfolio`;
              card(slide, panelX1, panelY + panelH + 0.18, 11.6, 1.1, "Transition Summary", summaryBody, C.goldPale);
            }

            slide.addText(
              "Based on real historical monthly prices. For illustrative purposes only — past performance does not guarantee future results.",
              { x: 0.85, y: 6.62, w: 11.6, h: 0.18, fontSize: 7, italic: true, color: C.muted, margin: 0, align: "center" }
            );

            footer(slide);

            // ── SLIDE 4: PORTFOLIO ANALYTICS ─────────────────────────────────
            // Growth of $10k line chart (left) + Risk/Return efficient frontier
            // (right) — mirrors every visual shown in the modal's backtest section.
            {
              slide = pptx.addSlide();
              title(
                slide,
                "PORTFOLIO STRATEGY",
                "Portfolio Analytics",
                hasBacktest && ts.startDate
                  ? `Backtested ${ts.startDate} – ${ts.endDate} · past performance does not guarantee future results`
                  : "Portfolio risk / return analysis"
              );

              // ── Growth of $10k SVG ───────────────────────────────────────
              function makeGrowth10kSvg(tSeries, cSeries) {
                if (!tSeries || tSeries.length < 2) return null;
                // Build unified date axis
                const dateSet = new Set(tSeries.map(p => p.date));
                if (cSeries) cSeries.forEach(p => dateSet.add(p.date));
                const dates = [...dateSet].sort();
                const n = dates.length;
                const tMap = Object.fromEntries(tSeries.map(p => [p.date, p.value]));
                const cMap = cSeries ? Object.fromEntries(cSeries.map(p => [p.date, p.value])) : {};
                const allVals = [...Object.values(tMap), ...Object.values(cMap)];
                const minV = Math.min(...allVals) * 0.96;
                const maxV = Math.max(...allVals) * 1.04;

                const W = 620, H = 340;
                const ml = 72, mr = 18, mt = 22, mb = 46;
                const pw = W - ml - mr, ph = H - mt - mb;
                const toX = i => ml + (i / Math.max(n - 1, 1)) * pw;
                const toY = v => mt + ph - ((v - minV) / Math.max(maxV - minV, 1)) * ph;

                // Paths
                const tPath = dates.reduce((acc, d, i) => {
                  if (tMap[d] == null) return acc;
                  return acc + (acc === "" ? "M" : " L") + `${toX(i).toFixed(1)},${toY(tMap[d]).toFixed(1)}`;
                }, "");
                const cPath = cSeries ? dates.reduce((acc, d, i) => {
                  if (cMap[d] == null) return acc;
                  return acc + (acc === "" ? "M" : " L") + `${toX(i).toFixed(1)},${toY(cMap[d]).toFixed(1)}`;
                }, "") : "";

                // Gridlines — 5 nice Y values
                const step = (maxV - minV) / 4;
                const yTicks = Array.from({ length: 5 }, (_, i) => Math.round((minV + step * i) / 1000) * 1000);
                const gridLines = yTicks.map(v => {
                  const y = toY(v).toFixed(1);
                  return `<line x1="${ml}" y1="${y}" x2="${W - mr}" y2="${y}" stroke="#E8ECF0" stroke-width="1"/>
                    <text x="${ml - 5}" y="${(parseFloat(y) + 4).toFixed(0)}" text-anchor="end" font-size="9" fill="#6E7E9A">$${(v / 1000).toFixed(0)}k</text>`;
                }).join("");

                // X labels: first, middle, last
                const xIdx = [0, Math.floor(n / 2), n - 1];
                const xLabels = xIdx.map(i => `<text x="${toX(i).toFixed(1)}" y="${mt + ph + 16}" text-anchor="middle" font-size="9" fill="#6E7E9A">${dates[i]?.substring(0, 7) || ""}</text>`).join("");

                // Endpoint labels
                const tLast = tSeries[tSeries.length - 1];
                const cLast = cSeries ? cSeries[cSeries.length - 1] : null;
                const tEndX = toX(dates.indexOf(tLast.date));
                const tEndY = toY(tLast.value);
                const cEndX = cLast ? toX(dates.indexOf(cLast.date)) : 0;
                const cEndY = cLast ? toY(cLast.value) : 0;

                return `<svg xmlns="http://www.w3.org/2000/svg" width="${W}" height="${H}" viewBox="0 0 ${W} ${H}">
                  <rect width="${W}" height="${H}" fill="#FFFFFF"/>
                  ${gridLines}
                  <line x1="${ml}" y1="${mt}" x2="${ml}" y2="${mt + ph}" stroke="#D5DAE5" stroke-width="1"/>
                  <line x1="${ml}" y1="${mt + ph}" x2="${W - mr}" y2="${mt + ph}" stroke="#D5DAE5" stroke-width="1"/>
                  ${tPath ? `<path d="${tPath}" fill="none" stroke="#1E7A6E" stroke-width="2.8" stroke-linejoin="round" stroke-linecap="round"/>` : ""}
                  ${cPath ? `<path d="${cPath}" fill="none" stroke="#C94F3A" stroke-width="2.8" stroke-linejoin="round" stroke-linecap="round" stroke-dasharray="8,4"/>` : ""}
                  ${tLast ? `<circle cx="${tEndX.toFixed(1)}" cy="${tEndY.toFixed(1)}" r="5" fill="#1E7A6E"/>
                    <text x="${(tEndX - 6).toFixed(1)}" y="${(tEndY - 9).toFixed(1)}" text-anchor="middle" font-size="10" font-weight="700" fill="#1E7A6E">$${Math.round(tLast.value).toLocaleString("en-US")}</text>` : ""}
                  ${cLast ? `<circle cx="${cEndX.toFixed(1)}" cy="${cEndY.toFixed(1)}" r="5" fill="#C94F3A"/>
                    <text x="${(cEndX - 6).toFixed(1)}" y="${(cEndY + 18).toFixed(1)}" text-anchor="middle" font-size="10" font-weight="700" fill="#C94F3A">$${Math.round(cLast.value).toLocaleString("en-US")}</text>` : ""}
                  <line x1="${ml + pw - 210}" y1="14" x2="${ml + pw - 185}" y2="14" stroke="#1E7A6E" stroke-width="3"/>
                  <text x="${ml + pw - 178}" y="18" font-size="10" fill="#1A2030">Target Portfolio</text>
                  ${cSeries ? `<line x1="${ml + pw - 100}" y1="14" x2="${ml + pw - 75}" y2="14" stroke="#C94F3A" stroke-width="3" stroke-dasharray="8,4"/>
                    <text x="${ml + pw - 68}" y="18" font-size="10" fill="#1A2030">Current</text>` : ""}
                  ${xLabels}
                  <text x="${(ml + pw / 2).toFixed(0)}" y="${H - 3}" text-anchor="middle" font-size="9" fill="#6E7E9A">Growth of $10,000 Invested</text>
                </svg>`;
              }

              // ── Risk / Return SVG (matches modal RiskReturnChart math) ────
              function makeRiskReturnSvg(tVol, tRet, cVol, cRet) {
                if (tVol == null || tRet == null) return null;
                const mv_sig = 0.01;
                let mv_ret = 0.015;
                let sqrtScale = (tRet - mv_ret) / Math.sqrt(tVol - mv_sig);
                const hasCurrent = cVol != null && cRet != null;

                if (hasCurrent) {
                  const fAtCX = mv_ret + sqrtScale * Math.sqrt(Math.max(cVol - mv_sig, 0));
                  if (cRet > fAtCX && cVol > tVol) {
                    const r = Math.sqrt((cVol - mv_sig) / (tVol - mv_sig));
                    if (Math.abs(1 - r) > 0.001) mv_ret = (cRet - r * tRet) / (1 - r);
                    sqrtScale = (tRet - mv_ret) / Math.sqrt(tVol - mv_sig);
                  }
                }

                const frontierRet = sig => { const d = sig - mv_sig; return d >= 0 ? mv_ret + sqrtScale * Math.sqrt(d) : null; };
                const frontierSig = ret => ret <= mv_ret ? mv_sig : mv_sig + ((ret - mv_ret) / sqrtScale) ** 2;

                const sigMax = Math.max(tVol, hasCurrent ? cVol : 0) * 2.6;
                const frontierPts = Array.from({ length: 81 }, (_, i) => {
                  const sig = mv_sig + (sigMax - mv_sig) * i / 80;
                  const ret = frontierRet(sig);
                  return ret != null ? { x: sig, y: ret } : null;
                }).filter(Boolean);

                const allPtsY = [tRet, hasCurrent ? cRet : tRet, mv_ret, ...frontierPts.map(p => p.y)];
                const xa = mv_sig * 0.55, xb = sigMax * 1.03;
                const ySpan = Math.max(...allPtsY) - Math.min(...allPtsY);
                const ya = Math.min(...allPtsY) - ySpan * 0.12;
                const yb = Math.max(...allPtsY) + ySpan * 0.25;

                const W = 480, H = 360;
                const ml = 52, mr = 18, mt = 20, mb = 44;
                const pw = W - ml - mr, ph = H - mt - mb;
                const toX = v => ml + ((v - xa) / (xb - xa)) * pw;
                const toY = v => mt + ph - ((v - ya) / (yb - ya)) * ph;

                const pathD = frontierPts.map(({ x, y }, i) => `${i === 0 ? "M" : "L"}${toX(x).toFixed(1)},${toY(y).toFixed(1)}`).join(" ");

                const N = 4;
                const gxs = Array.from({ length: N + 1 }, (_, i) => xa + (xb - xa) * i / N);
                const gys = Array.from({ length: N + 1 }, (_, i) => ya + (yb - ya) * i / N);
                const gridX = gxs.map(v => `<line x1="${toX(v).toFixed(1)}" y1="${mt}" x2="${toX(v).toFixed(1)}" y2="${mt + ph}" stroke="#E8ECF0" stroke-width="1"/>
                  <text x="${toX(v).toFixed(1)}" y="${mt + ph + 14}" text-anchor="middle" font-size="9" fill="#6E7E9A">${(v * 100).toFixed(1)}%</text>`).join("");
                const gridY = gys.map(v => `<line x1="${ml}" y1="${toY(v).toFixed(1)}" x2="${ml + pw}" y2="${toY(v).toFixed(1)}" stroke="#E8ECF0" stroke-width="1"/>
                  <text x="${ml - 5}" y="${(parseFloat(toY(v).toFixed(1)) + 3.5).toFixed(0)}" text-anchor="end" font-size="9" fill="#6E7E9A">${(v * 100).toFixed(1)}%</text>`).join("");

                // Gap annotations
                let gaps = "";
                if (hasCurrent) {
                  const effSig = frontierSig(cRet);
                  const effRet = frontierRet(cVol);
                  const fx = toX(effSig), cy_ = toY(cRet), cx_ = toX(cVol);
                  if (cx_ - fx > 8) {
                    const exRisk = (cVol - effSig) * 100;
                    gaps += `<line x1="${fx.toFixed(1)}" y1="${cy_.toFixed(1)}" x2="${(cx_ - 10).toFixed(1)}" y2="${cy_.toFixed(1)}" stroke="#C94F3A" stroke-width="1.5" stroke-dasharray="4 2.5" opacity="0.75"/>
                      <rect x="${((fx + cx_) / 2 - 40).toFixed(1)}" y="${(cy_ - 19).toFixed(1)}" width="80" height="14" rx="3" fill="white" opacity="0.85"/>
                      <text x="${((fx + cx_) / 2).toFixed(1)}" y="${(cy_ - 8).toFixed(1)}" text-anchor="middle" font-size="9" font-weight="700" fill="#C94F3A">+${exRisk.toFixed(1)}% excess risk</text>`;
                  }
                  if (effRet != null && cy_ - toY(effRet) > 8) {
                    const fRetY = toY(effRet);
                    const foregone = (effRet - cRet) * 100;
                    gaps += `<line x1="${cx_.toFixed(1)}" y1="${(cy_ - 10).toFixed(1)}" x2="${cx_.toFixed(1)}" y2="${(fRetY + 8).toFixed(1)}" stroke="#1E7A6E" stroke-width="1.5" stroke-dasharray="4 2.5" opacity="0.75"/>
                      <rect x="${(cx_ + 4).toFixed(1)}" y="${((cy_ + fRetY) / 2 - 7).toFixed(1)}" width="76" height="14" rx="3" fill="white" opacity="0.85"/>
                      <text x="${(cx_ + 42).toFixed(1)}" y="${((cy_ + fRetY) / 2 + 4).toFixed(1)}" text-anchor="middle" font-size="9" font-weight="700" fill="#1E7A6E">+${foregone.toFixed(1)}% foregone</text>`;
                  }
                }

                const tDotX = toX(tVol), tDotY = toY(tRet);
                const cDotX = hasCurrent ? toX(cVol) : 0, cDotY = hasCurrent ? toY(cRet) : 0;
                const tLabelY = tDotY - 12 < mt ? tDotY + 16 : tDotY - 12;
                const cLabelY = hasCurrent ? (cDotY + 16 > mt + ph ? cDotY - 12 : cDotY + 16) : 0;

                return `<svg xmlns="http://www.w3.org/2000/svg" width="${W}" height="${H}" viewBox="0 0 ${W} ${H}">
                  <rect width="${W}" height="${H}" fill="#FFFFFF"/>
                  <defs><clipPath id="rrPpt"><rect x="${ml}" y="${mt}" width="${pw}" height="${ph}"/></clipPath></defs>
                  ${gridX}${gridY}
                  <line x1="${ml}" y1="${mt}" x2="${ml}" y2="${mt + ph}" stroke="#C8D0DE" stroke-width="1.5"/>
                  <line x1="${ml}" y1="${mt + ph}" x2="${ml + pw}" y2="${mt + ph}" stroke="#C8D0DE" stroke-width="1.5"/>
                  <path d="${pathD} L${toX(sigMax).toFixed(1)},${(mt + ph).toFixed(1)} L${toX(mv_sig).toFixed(1)},${(mt + ph).toFixed(1)} Z" fill="#FDF0ED" opacity="0.5" clip-path="url(#rrPpt)"/>
                  <path d="${pathD}" fill="none" stroke="#1A2744" stroke-width="2.2" opacity="0.75" clip-path="url(#rrPpt)"/>
                  <circle cx="${toX(mv_sig).toFixed(1)}" cy="${toY(mv_ret).toFixed(1)}" r="3.5" fill="#1A2744" opacity="0.5"/>
                  <text x="${(toX(mv_sig) + 6).toFixed(1)}" y="${(toY(mv_ret) + 4).toFixed(1)}" font-size="8" fill="#6b7a99" font-style="italic">Efficient Frontier</text>
                  ${gaps}
                  <circle cx="${tDotX.toFixed(1)}" cy="${tDotY.toFixed(1)}" r="7" fill="#1E7A6E"/>
                  <text x="${(tDotX + 11).toFixed(1)}" y="${tLabelY.toFixed(1)}" font-size="9.5" font-weight="700" fill="#1E7A6E">Target  ${(tRet * 100).toFixed(1)}% / ${(tVol * 100).toFixed(1)}% vol</text>
                  ${hasCurrent ? `<circle cx="${cDotX.toFixed(1)}" cy="${cDotY.toFixed(1)}" r="7" fill="#C94F3A"/>
                    <text x="${(cDotX + 11).toFixed(1)}" y="${cLabelY.toFixed(1)}" font-size="9.5" font-weight="700" fill="#C94F3A">Current  ${(cRet * 100).toFixed(1)}% / ${(cVol * 100).toFixed(1)}% vol</text>` : ""}
                  <text x="${(ml + pw / 2).toFixed(0)}" y="${H - 4}" text-anchor="middle" font-size="9.5" font-weight="600" fill="#6b7a99">Risk — Annualized Volatility</text>
                  <text transform="rotate(-90 11 ${(mt + ph / 2).toFixed(0)})" x="11" y="${(mt + ph / 2).toFixed(0)}" text-anchor="middle" font-size="9.5" font-weight="600" fill="#6b7a99">Return — CAGR</text>
                </svg>`;
              }

              const growthSvg = makeGrowth10kSvg(ts.growthSeries, cs?.growthSeries);
              const rrSvg = makeRiskReturnSvg(ts.annualizedVolatility, ts.annualizedReturn, cs?.annualizedVolatility, cs?.annualizedReturn);

              if (growthSvg) addSvg(slide, growthSvg, 0.65, 1.88, 7.2, 4.3);
              if (rrSvg)    addSvg(slide, rrSvg,     8.1, 1.88, 5.0, 3.8);

              // Section labels
              slide.addText("GROWTH OF $10,000", { x: 0.65, y: 1.70, w: 4.0, h: 0.16, fontSize: 7, bold: true, color: C.blue, charSpace: 1.2, margin: 0 });
              slide.addText("RISK / RETURN  ·  EFFICIENT FRONTIER", { x: 8.1, y: 1.70, w: 5.0, h: 0.16, fontSize: 7, bold: true, color: C.blue, charSpace: 1.2, margin: 0 });

              // Key-stat strip above charts (pulled from same window as charts)
              const statY = 6.24;
              const nStats = c ? 4 : 2;
              const statW = 11.6 / nStats;
              statBox(slide, 0.65,                  statY, statW - 0.15, "Target CAGR",         pctF(ts.annualizedReturn),       C.tealPale,  C.teal);
              statBox(slide, 0.65 + statW,           statY, statW - 0.15, "Target Max Drawdown", pctF(ts.maxDrawdown),            C.white,     C.coral);
              if (c) {
                statBox(slide, 0.65 + statW * 2,     statY, statW - 0.15, "Current CAGR",        pctF(cs?.annualizedReturn),      C.white,     C.navy);
                statBox(slide, 0.65 + statW * 3,     statY, statW - 0.15, "Current Max Drawdown",pctF(cs?.maxDrawdown),           C.coralPale, C.coral);
              }

              slide.addText(
                "Based on real historical monthly prices. Efficient frontier is a theoretical illustration only — not derived from mean-variance optimization. Past performance does not guarantee future results.",
                { x: 0.65, y: 6.88, w: 11.6, h: 0.18, fontSize: 6.5, italic: true, color: C.muted, margin: 0, align: "center" }
              );

              footer(slide);
            }
          }
        }
      }

      // PLANNING_MODULE_SLIDES
      // Adds planning module slides based on selected Proposal Modules.
      {
        // modules already defined above
        const notesLower = String(notes || "").toLowerCase();

        const concentrationPctModule =
          Number(data.concentration) > 1
            ? Number(data.concentration)
            : Number(data.concentration || 0) * 100;

        const hasHighConcentrationModule =
          concentrationPctModule >= 25 &&
          Number(data.stockPosition || 0) > 0 &&
          !!data.ticker;

        function cleanNumber(value) {
          const n = Number(value || 0);
          return Number.isFinite(n) ? n : 0;
        }

        function moduleBulletList(slide, bullets, x, y, w, fontSize = 11) {
          slide.addText(
            bullets.map((b) => `• ${b}`).join("\n"),
            {
              x,
              y,
              w,
              h: Math.max(0.45, bullets.length * 0.28),
              fontSize,
              color: C.text,
              margin: 0,
              breakLine: true,
              fit: "shrink",
              valign: "top",
            }
          );
        }

        function moduleTwoColumnSlide(slide, leftTitle, leftBullets, rightTitle, rightBullets) {
          slide.addShape(pptx.ShapeType.roundRect, {
            x: 0.85,
            y: 2.0,
            w: 5.45,
            h: 3.9,
            rectRadius: 0.06,
            fill: { color: C.white },
            line: { color: C.border, width: 1 },
          });

          slide.addShape(pptx.ShapeType.rect, {
            x: 0.85,
            y: 2.0,
            w: 0.09,
            h: 3.9,
            fill: { color: C.gold },
            line: { color: C.gold },
          });

          slide.addText(leftTitle, {
            x: 1.12,
            y: 2.15,
            w: 4.8,
            h: 0.25,
            fontSize: 15,
            bold: true,
            color: C.navy,
            margin: 0,
          });

          moduleBulletList(slide, leftBullets, 1.15, 2.62, 4.75, 11);

          slide.addShape(pptx.ShapeType.roundRect, {
            x: 6.65,
            y: 2.0,
            w: 5.45,
            h: 3.9,
            rectRadius: 0.06,
            fill: { color: C.goldPale },
            line: { color: C.border, width: 1 },
          });

          slide.addShape(pptx.ShapeType.rect, {
            x: 6.65,
            y: 2.0,
            w: 0.09,
            h: 3.9,
            fill: { color: C.navy },
            line: { color: C.navy },
          });

          slide.addText(rightTitle, {
            x: 6.92,
            y: 2.15,
            w: 4.8,
            h: 0.25,
            fontSize: 15,
            bold: true,
            color: C.navy,
            margin: 0,
          });

          moduleBulletList(slide, rightBullets, 6.95, 2.62, 4.75, 11);
        }

        function timelineNode(slide, label, titleText, bodyText, x, y, accent) {
          slide.addShape(pptx.ShapeType.ellipse, {
            x,
            y,
            w: 0.65,
            h: 0.65,
            fill: { color: accent },
            line: { color: accent },
          });

          slide.addText(label, {
            x,
            y: y + 0.22,
            w: 0.65,
            h: 0.12,
            fontSize: 8,
            bold: true,
            color: C.white,
            align: "center",
            margin: 0,
          });

          slide.addText(titleText, {
            x: x - 0.65,
            y: y + 0.88,
            w: 1.95,
            h: 0.24,
            fontSize: 11,
            bold: true,
            color: C.navy,
            align: "center",
            margin: 0,
            fit: "shrink",
          });

          slide.addText(bodyText, {
            x: x - 0.75,
            y: y + 1.22,
            w: 2.15,
            h: 0.55,
            fontSize: 8,
            color: C.text,
            align: "center",
            margin: 0,
            fit: "shrink",
          });
        }

        const highConcentrationNeed = hasHighConcentrationModule
          ? `Transition the concentrated ${data.ticker} position through a staged, tax-aware process.`
          : "Align the portfolio to client goals, liquidity needs, and risk tolerance.";

        // Risk Assessment Overview
        if (modules.riskManagementOverview) {
          slide = pptx.addSlide();
          title(
            slide,
            "PLANNING MODULE",
            "Risk Assessment Overview",
            `Risk profile, allocation targets, and scenario analysis for ${name}.`
          );

          const riskProfiles = [
            { key: "conservative", label: "Conservative", equity: 20, fi: 80 },
            { key: "moderatelyConservative", label: "Mod. Conservative", equity: 30, fi: 70 },
            { key: "conservativePlus", label: "Conservative Plus", equity: 40, fi: 60 },
            { key: "balanced", label: "Balanced", equity: 50, fi: 50 },
            { key: "balancedPlus", label: "Balanced Plus", equity: 60, fi: 40 },
            { key: "growth", label: "Growth", equity: 70, fi: 30 },
            { key: "growthPlus", label: "Growth Plus", equity: 80, fi: 20 },
            { key: "aggressive", label: "Aggressive", equity: 100, fi: 0 },
          ];
          const selectedRP = riskProfiles.find(p => p.key === selectedRiskProfile) || riskProfiles[3];
          const currentConc = cleanNumber(data.concentration);

          // Stat boxes — y:1.88 (clears subtitle)
          // If a risk number was sourced from client notes, show it; otherwise "Not on File"
          const riskNumDisplay = riskNumber != null ? String(riskNumber) : "Not on File";
          const riskNumColor   = riskNumber != null ? C.navy : C.muted;
          const riskNumFill    = riskNumber != null ? C.white : "F5F6F8";

          statBox(slide, 0.55, 1.88, 2.6,  "Client Risk Number",  riskNumDisplay,      riskNumFill, riskNumColor);
          statBox(slide, 3.35, 1.88, 2.6,  "Recommended Profile", selectedRP.label,    C.goldPale,  C.gold);
          statBox(slide, 6.15, 1.88, 2.6,  "Equity / Fixed",      `${selectedRP.equity} / ${selectedRP.fi}`, C.bluePale, C.blue);
          statBox(slide, 8.95, 1.88, 2.6,  "Concentration",       pct(currentConc),    currentConc > 50 ? C.coralPale : C.tealPale, currentConc > 50 ? C.coral : C.teal);

          // Risk-number footnote (only when not on file)
          if (riskNumber == null) {
            slide.addText("* Request Nitrogen / Riskalyze score from advisor if available.", {
              x: 0.55, y: 2.84, w: 7.0, h: 0.16,
              fontSize: 7.5, color: C.muted, italic: true, margin: 0,
            });
          }

          // Risk spectrum — starts at y:3.08
          slide.addText("RISK SPECTRUM", { x: 0.55, y: 3.08, w: 3, h: 0.16, fontSize: 7, bold: true, color: C.blue, charSpace: 1.2, margin: 0 });
          riskProfiles.forEach((rp, i) => {
            const sx = 0.55 + i * 1.53;
            const isSelected = rp.key === selectedRiskProfile;
            slide.addShape(pptx.ShapeType.roundRect, {
              x: sx, y: 3.28, w: 1.43, h: isSelected ? 0.54 : 0.44,
              rectRadius: 0.06,
              fill: { color: isSelected ? C.navy : "F0F2F6" },
              line: { color: isSelected ? C.gold : C.border, width: isSelected ? 1.5 : 0.5 },
            });
            slide.addText(rp.label, {
              x: sx + 0.04, y: 3.32 + (isSelected ? 0.02 : 0.08), w: 1.35, h: 0.28,
              fontSize: isSelected ? 8 : 7, bold: isSelected,
              color: isSelected ? C.white : C.muted,
              align: "center", margin: 0, fit: "shrink",
            });
          });

          // Scenario cards — spectrum ends ~3.82, cards at 4.08
          const downside30 = cleanNumber(data.investableAssets) * 0.3 * (selectedRP.equity / 100);
          const upside20   = cleanNumber(data.investableAssets) * 0.2 * (selectedRP.equity / 100);

          // Downside
          slide.addShape(pptx.ShapeType.roundRect, { x: 0.55, y: 4.08, w: 5.8, h: 1.72, rectRadius: 0.08, fill: { color: C.coralPale }, line: { color: C.coral, width: 0.7 } });
          slide.addShape(pptx.ShapeType.rect,      { x: 0.55, y: 4.08, w: 0.06, h: 1.72, fill: { color: C.coral }, line: { color: C.coral } });
          slide.addText("DOWNSIDE  —  30% EQUITY DECLINE", { x: 0.75, y: 4.18, w: 5.4, h: 0.18, fontSize: 7.5, bold: true, color: C.coral, charSpace: 0.8, margin: 0 });
          slide.addText(`~(${fmtM(downside30)}) estimated impact`, { x: 0.75, y: 4.42, w: 5.4, h: 0.26, fontFace: "Georgia", fontSize: 16, bold: true, color: C.coral, margin: 0 });
          slide.addText([
            { text: `• ${selectedRP.equity}% equity × 30% decline on ${fmtM(cleanNumber(data.investableAssets))} investable assets`, options: { bullet: false } },
            { text: `\n• Concentrated position amplifies drawdown until diversification executes`, options: { bullet: false } },
          ], { x: 0.75, y: 4.74, w: 5.4, h: 0.85, fontSize: 10, color: C.text, margin: 0, valign: "top" });

          // Upside
          slide.addShape(pptx.ShapeType.roundRect, { x: 6.7,  y: 4.08, w: 5.8, h: 1.72, rectRadius: 0.08, fill: { color: C.tealPale }, line: { color: C.teal, width: 0.7 } });
          slide.addShape(pptx.ShapeType.rect,      { x: 6.7,  y: 4.08, w: 0.06, h: 1.72, fill: { color: C.teal }, line: { color: C.teal } });
          slide.addText("UPSIDE  —  20% EQUITY GAIN",         { x: 6.9,  y: 4.18, w: 5.4, h: 0.18, fontSize: 7.5, bold: true, color: C.teal, charSpace: 0.8, margin: 0 });
          slide.addText(`+${fmtM(upside20)} estimated benefit`, { x: 6.9,  y: 4.42, w: 5.4, h: 0.26, fontFace: "Georgia", fontSize: 16, bold: true, color: C.teal, margin: 0 });
          slide.addText([
            { text: `• ${selectedRP.equity}% equity participates in market upside`, options: { bullet: false } },
            { text: `\n• Fixed income (${selectedRP.fi}%) buffers volatility across cycles`, options: { bullet: false } },
          ], { x: 6.9, y: 4.74, w: 5.4, h: 0.85, fontSize: 10, color: C.text, margin: 0, valign: "top" });

          // Rationale bullets — cards end ~5.8, note at 5.96
          slide.addText(
            riskNumber != null
              ? `Risk Number ${riskNumber} → ${selectedRP.label} (${selectedRP.equity}/${selectedRP.fi}).  Concentration at ${pct(currentConc)} adds above-model volatility until diversification plan executes.`
              : `${selectedRP.label} profile (${selectedRP.equity}/${selectedRP.fi}) selected based on risk language, time horizon, and planning objectives.  Concentration at ${pct(currentConc)} adds above-model volatility.`,
            { x: 0.55, y: 5.88, w: 11.95, h: 0.22, fontSize: 8.5, color: C.muted, margin: 0, fit: "shrink" }
          );

          footer(slide);
        }

        // Goals Timeline
        if (modules.goalsTimeline) {
          slide = pptx.addSlide();
          title(
            slide,
            "PLANNING MODULE",
            "Goals Timeline",
            "Organizes the client’s planning priorities by time horizon."
          );

          slide.addShape(pptx.ShapeType.line, {
            x: 1.35,
            y: 3.25,
            w: 10.3,
            h: 0,
            line: { color: C.border, width: 2.2 },
          });

          timelineNode(
            slide,
            "1",
            "Near Term",
            notesLower.includes("liquidity") || notesLower.includes("cash")
              ? "Address cash needs, liquidity reserve, and near-term implementation constraints."
              : "Confirm immediate planning items and implementation readiness.",
            1.2,
            2.9,
            C.navy
          );

          timelineNode(
            slide,
            "2",
            "1–3 Years",
            hasHighConcentrationModule
              ? `Stage ${data.ticker} diversification and coordinate taxes.`
              : "Rebalance portfolio and align allocation with risk profile.",
            4.45,
            2.9,
            C.gold
          );

          timelineNode(
            slide,
            "3",
            "3–7 Years",
            notesLower.includes("retirement")
              ? "Support retirement income planning and portfolio durability."
              : "Monitor goals, risk tolerance, and portfolio strategy fit.",
            7.7,
            2.9,
            C.teal
          );

          timelineNode(
            slide,
            "4",
            "Long Term",
            notesLower.includes("legacy") || notesLower.includes("estate") || notesLower.includes("charitable")
              ? "Coordinate legacy, estate, and charitable planning objectives."
              : "Maintain diversified capital for future goals and family needs.",
            10.95,
            2.9,
            C.blue
          );

          card(
            slide,
            0.9,
            5.25,
            11.35,
            0.85,
            "Advisor Use",
            "Use this timeline to confirm target dollar amounts, priorities, timing, and owners for each goal before implementation.",
            C.white
          );

          footer(slide);
        }

        // Liquidity Needs Review
        if (modules.liquidityNeedsReview) {
          slide = pptx.addSlide();
          title(
            slide,
            "PLANNING MODULE",
            "Liquidity Needs Review",
            "Identifies cash needs, withdrawal timing, and implementation flexibility."
          );

          statBox(slide, 0.85, 1.8, 2.5, "Investable Assets", fmtM(data.investableAssets), C.white, C.navy);
          statBox(slide, 3.65, 1.8, 2.5, "Liquid Planning Need", notesLower.includes("liquidity") || notesLower.includes("cash") ? "Stated" : "Review", C.goldPale, C.gold);
          statBox(slide, 6.45, 1.8, 2.5, "Income / Withdrawal", notesLower.includes("income") || notesLower.includes("withdraw") ? "Relevant" : "Review", C.tealPale, C.teal);
          statBox(slide, 9.25, 1.8, 2.5, "Implementation", "Staged", C.bluePale, C.blue);

          moduleTwoColumnSlide(
            slide,
            "Liquidity Questions",
            [
              "What cash is needed in the next 12–24 months?",
              "Are there expected withdrawals, home purchases, tuition, gifts, or tax payments?",
              "How much portfolio value should remain liquid during implementation?",
              "Does the selected strategy include illiquidity or lock-up constraints?",
            ],
            "Planning Implications",
            [
              highConcentrationNeed,
              "Maintain a liquidity reserve before allocating to longer-term or less-liquid strategies.",
              "Coordinate portfolio transition with tax calendar and spending needs.",
              "Revisit liquidity assumptions before final allocation implementation.",
            ]
          );

          footer(slide);
        }

        // Tax Planning Overview
        if (modules.taxPlanningOverview) {
          slide = pptx.addSlide();
          title(
            slide,
            "PLANNING MODULE",
            "Tax Planning Overview",
            "Frames tax-aware implementation considerations before final recommendations."
          );

          const taxRate = cleanNumber(data.taxRate || data.totalTaxRate || 0);
          const costBasisPct = cleanNumber(data.costBasisPct || 0);
          const embeddedGain = Math.max(0, cleanNumber(data.stockPosition) * (1 - costBasisPct / 100));
          const estimatedTax = data.immediateTax || (embeddedGain * taxRate / 100);

          statBox(slide, 0.85, 1.8, 2.35, "Tax Rate", taxRate ? `${taxRate.toFixed(1)}%` : "Review", C.white, C.navy);
          statBox(slide, 3.45, 1.8, 2.35, "Cost Basis", costBasisPct ? `${costBasisPct.toFixed(1)}%` : "Review", C.goldPale, C.gold);
          statBox(slide, 6.05, 1.8, 2.35, "Est. Embedded Gain", fmtM(embeddedGain), C.tealPale, C.teal);
          statBox(slide, 8.65, 1.8, 2.35, "Immediate Sale Tax", fmtM(estimatedTax), C.bluePale, C.blue);

          moduleTwoColumnSlide(
            slide,
            "Tax Issues to Review",
            [
              "Embedded gains and cost basis assumptions.",
              "Federal, state, and local tax rates.",
              "Timing of sales, charitable transfers, and harvesting.",
              "Coordination with CPA before implementation.",
            ],
            "Potential Planning Tools",
            [
              "Charitable Remainder Trust for charitable and tax planning.",
              "Leveraged tax-loss harvesting to create loss capacity.",
              "Option collar to manage risk without immediate full sale.",
              "Staged sales to reduce timing and market risk.",
            ]
          );

          footer(slide);
        }

        // Income & Expense Snapshot — removed per advisor preference.

        // Retirement Planning
        if (modules.retirementPlanning) {
          slide = pptx.addSlide();
          title(
            slide,
            "PLANNING MODULE",
            "Retirement Planning",
            "Connects portfolio strategy to retirement income, timing, and sustainability."
          );

          const retirementRelevant = notesLower.includes("retirement") || notesLower.includes("retire") || notesLower.includes("income") || notesLower.includes("withdraw");

          statBox(slide, 0.85, 1.8, 2.6, "Retirement Goal", retirementRelevant ? "Mentioned" : "Review", C.white, C.navy);
          statBox(slide, 3.75, 1.8, 2.6, "Risk Profile", riskProfileLabel, C.goldPale, C.gold);
          statBox(slide, 6.65, 1.8, 2.6, "Portfolio Strategy", selectedPortfolioStrategyLabel, C.tealPale, C.teal);
          statBox(slide, 9.55, 1.8, 2.1, "Income Need", notesLower.includes("income") ? "Relevant" : "Review", C.bluePale, C.blue);

          moduleTwoColumnSlide(
            slide,
            "Retirement Questions",
            [
              "When does the client expect to retire or reduce work?",
              "What annual income is needed from the portfolio?",
              "How long should the assets last?",
              "What legacy or permanent capital goals remain after retirement needs?",
            ],
            "Planning Direction",
            [
              "Match risk profile to retirement timing and behavioral tolerance.",
              "Preserve enough liquidity for near-term spending and taxes.",
              "Use diversified portfolio strategy to support long-term withdrawals.",
              "Revisit allocation if retirement timing, spending, or market conditions change.",
            ]
          );

          footer(slide);
        }

        // Legacy & Wealth Transfer
        if (modules.legacyWealthTransfer) {
          slide = pptx.addSlide();
          title(
            slide,
            "PLANNING MODULE",
            "Legacy & Wealth Transfer",
            "Frames family, charitable, and multigenerational planning objectives."
          );

          moduleTwoColumnSlide(
            slide,
            "Legacy Objectives",
            [
              "Identify heirs, family priorities, and charitable intentions.",
              "Clarify whether assets are for lifetime spending, heirs, or permanent capital.",
              "Review gifting, trust, and beneficiary planning opportunities.",
              "Coordinate investment strategy with wealth-transfer goals.",
            ],
            "Potential Planning Actions",
            [
              "Coordinate with estate attorney and CPA.",
              "Evaluate charitable structures if philanthropy is important.",
              "Align portfolio liquidity with estate-tax and transfer needs.",
              "Document goals so investment strategy reflects family priorities.",
            ]
          );

          statBox(slide, 0.85, 6.0, 2.8, "Legacy Mentioned", notesLower.includes("legacy") || notesLower.includes("heirs") || notesLower.includes("estate") ? "Yes" : "Review", C.white, C.navy);
          statBox(slide, 3.95, 6.0, 2.8, "Charitable Intent", notesLower.includes("charitable") || notesLower.includes("donor") ? "Yes" : "Review", C.goldPale, C.gold);
          statBox(slide, 7.05, 6.0, 2.8, "Estate Coordination", "Recommended", C.tealPale, C.teal);

          footer(slide);
        }

        // Estate Planning Review
        if (modules.estatePlanningReview) {
          const extractEstateMoneyFromNotes = (patterns) => {
            const raw = String(notes || "").replace(/\s+/g, " ");
            for (const pattern of patterns) {
              const match = raw.match(pattern);
              if (match?.[1]) {
                const cleaned = String(match[1]).replace(/[$,]/g, "").trim();
                const value = Number(cleaned);
                if (Number.isFinite(value)) {
                  return value > 1000 ? value / 1000000 : value;
                }
              }
            }
            return 0;
          };

          const estateManagedAssets =
            Number(data.managedAssets || data.investableAssets || 0) ||
            extractEstateMoneyFromNotes([
              /Managed Assets\s*[:\-]?\s*\$?([\d,]+(?:\.\d+)?)/i,
              /Managed Investment Assets\s*\$?([\d,]+(?:\.\d+)?)/i,
            ]);

          const parsedRealEstateValue =
            extractEstateMoneyFromNotes([
              /Total Real Estate Value\s*\$?([\d,]+(?:\.\d+)?)/i,
              /Real Estate Holdings\s*\$?([\d,]+(?:\.\d+)?)/i,
              /approximately\s*\$?([\d,]+(?:\.\d+)?)\s*million in real estate/i,
            ]);

          const rawRealEstateValue = Number(data.realEstateValue || data.realEstateHoldings || 0);

          // Prefer explicit document text for real estate because some extractors confuse
          // 50/50 estate split values, e.g. $59M each, with real estate.
          const estateRealEstateValue =
            parsedRealEstateValue ||
            (
              rawRealEstateValue &&
              rawRealEstateValue !== estateNetWorth / 2
                ? rawRealEstateValue
                : 0
            );

          const estateNetWorth =
            Number(data.netWorth || 0) ||
            extractEstateMoneyFromNotes([
              /Estimated Net Worth\s*[:\-]?\s*\$?([\d,]+(?:\.\d+)?)/i,
              /Total Net Worth\s*\$?([\d,]+(?:\.\d+)?)/i,
            ]);

          const estateOtherAssets =
            Number(data.otherPrivateAssets || data.otherAssets || 0) ||
            extractEstateMoneyFromNotes([
              /Other Private\s*\/\s*Personal Assets\s*\$?([\d,]+(?:\.\d+)?)/i,
              /Other Private.*?Assets\s*\$?([\d,]+(?:\.\d+)?)/i,
            ]) ||
            Math.max(0, estateNetWorth - estateManagedAssets - estateRealEstateValue);

          const estateTaxExemption = data.estateTaxExemption || data.estateExemption || 30;
          const estateTaxRate = data.estateTaxRate || 0.40;

          const projectedYears = Number(data.estateProjectionYears || data.projectionYears || 25) || 25;
          const projectedGrowthRate =
            Number(data.estateGrowthRate || data.growthRate || 0.07) > 1
              ? Number(data.estateGrowthRate || data.growthRate || 7) / 100
              : Number(data.estateGrowthRate || data.growthRate || 0.07);

          const baseEstateData = {
            ...data,
            clientName: name || data.clientName || "Client Household",
            netWorth: estateNetWorth,
            managedAssets: estateManagedAssets,
            investableAssets: estateManagedAssets,
            realEstateValue: estateRealEstateValue,
            realEstateHoldings: estateRealEstateValue,
            otherAssets: estateOtherAssets,
            otherPrivateAssets: estateOtherAssets,
            estateTaxExemption,
            estateTaxRate,
            estateProjectionYears: projectedYears,
            estateGrowthRate: projectedGrowthRate,
          };

          const todayPng = await captureHtmlSlideAsPng(
            buildEstateSlideHtml({
              data: baseEstateData,
              isFuture: false,
            })
          );

          slide = pptx.addSlide();
          slide.addImage({
            data: todayPng,
            x: 0,
            y: 0,
            w: 13.333,
            h: 7.5,
          });

          const futureManagedAssets = estateManagedAssets * Math.pow(1 + projectedGrowthRate, projectedYears);
          const futureRealEstateValue = estateRealEstateValue * Math.pow(1 + projectedGrowthRate, projectedYears);
          const futureOtherAssets = estateOtherAssets * Math.pow(1 + projectedGrowthRate, projectedYears);
          const futureNetWorth = futureManagedAssets + futureRealEstateValue + futureOtherAssets;

          const futurePng = await captureHtmlSlideAsPng(
            buildEstateSlideHtml({
              data: {
                ...baseEstateData,
                netWorth: futureNetWorth,
                managedAssets: futureManagedAssets,
                investableAssets: futureManagedAssets,
                realEstateValue: futureRealEstateValue,
                realEstateHoldings: futureRealEstateValue,
                otherAssets: futureOtherAssets,
                otherPrivateAssets: futureOtherAssets,
              },
              isFuture: true,
            })
          );

          slide = pptx.addSlide();
          slide.addImage({
            data: futurePng,
            x: 0,
            y: 0,
            w: 13.333,
            h: 7.5,
          });
        }

        // Restrictions & Implementation Notes
        if (modules.restrictionsImplementationNotes) {
          slide = pptx.addSlide();
          title(
            slide,
            "PLANNING MODULE",
            "Restrictions & Implementation Notes",
            "Highlights implementation constraints that should be reviewed before execution."
          );

          const restrictionsDetected =
            notesLower.includes("restricted") ||
            notesLower.includes("rsu") ||
            notesLower.includes("insider") ||
            notesLower.includes("10b5") ||
            notesLower.includes("trading window") ||
            notesLower.includes("blackout");

          statBox(slide, 0.85, 1.8, 2.65, "Restrictions", restrictionsDetected ? "Detected" : "Review", C.white, C.navy);
          statBox(slide, 3.75, 1.8, 2.65, "Options Review", selectedStrategies?.collar ? "Needed" : "As Applicable", C.goldPale, C.gold);
          statBox(slide, 6.65, 1.8, 2.65, "Tax Review", "Recommended", C.tealPale, C.teal);
          statBox(slide, 9.55, 1.8, 2.1, "Legal Review", "As Needed", C.bluePale, C.blue);

          moduleTwoColumnSlide(
            slide,
            "Possible Restrictions",
            [
              "Insider status, blackout windows, or 10b5-1 requirements.",
              "Restricted stock, RSUs, options, or employer trading limitations.",
              "Tax-lot, cost basis, or holding-period constraints.",
              "Client preferences, ESG restrictions, or manager limitations.",
            ],
            "Implementation Notes",
            [
              "Confirm all trading restrictions before liquidation, collar, or transfer.",
              "Coordinate with compliance, CPA, attorney, and options desk as applicable.",
              "Sequence actions around tax year, liquidity needs, and market conditions.",
              "Document assumptions before final recommendation.",
            ]
          );

          footer(slide);
        }

        // Implementation Timeline module removed from generated deck.


        // Next Steps
        
        // CONCLUSION_ADVISOR_SUMMARY_SLIDE
        {
          const selectedFocusAreas = [];

          if (modules.riskOverview) selectedFocusAreas.push("Risk alignment");
          if (modules.estatePlanningReview) selectedFocusAreas.push("Estate planning");
          if (modules.taxPlanning) selectedFocusAreas.push("Tax coordination");
          if (modules.liquidityNeeds) selectedFocusAreas.push("Liquidity planning");
          if (modules.retirementPlanning) selectedFocusAreas.push("Retirement sustainability");
          if (modules.legacyWealthTransfer) selectedFocusAreas.push("Legacy transfer");
          if (modules.goalsTimeline) selectedFocusAreas.push("Goal sequencing");

          const concentratedStrategySelected =
            selectedStrategies?.crt ||
            selectedStrategies?.harvesting ||
            selectedStrategies?.collar;

          if (concentratedStrategySelected) {
            selectedFocusAreas.push("Concentration reduction");
          }

          const portfolioModelSelected =
            Object.values(selectedPortfolioStrategies || {}).some(Boolean);

          if (portfolioModelSelected) {
            selectedFocusAreas.push("Portfolio model selection");
          }

          const focusText =
            selectedFocusAreas.length > 0
              ? selectedFocusAreas.slice(0, 7).join(" • ")
              : "Advisor-led planning discussion";

          const conclusionText =
            concentratedStrategySelected
              ? `The selected proposal sections support a staged, tax-aware plan to reduce concentrated ${data.ticker || "single-stock"} risk while preserving flexibility for estate, liquidity, tax, and family planning objectives.`
              : "The selected proposal sections support a planning-focused discussion that connects the client’s goals, risk profile, tax picture, liquidity needs, and long-term family priorities.";

          slide = pptx.addSlide();

          title(
            slide,
            "CONCLUSION",
            "Proposal Conclusion",
            ""
          );

          // Large dark summary panel — starts below title heading (~y:1.22)
          slide.addShape(pptx.ShapeType.roundRect, {
            x: 0.85,
            y: 1.38,
            w: 11.6,
            h: 1.45,
            rectRadius: 0.06,
            fill: { color: C.navy },
            line: { color: C.navy },
          });

          slide.addText("Recommended Direction", {
            x: 1.15,
            y: 1.54,
            w: 3.2,
            h: 0.18,
            fontSize: 10,
            bold: true,
            color: C.gold,
            margin: 0,
            charSpace: 0.8,
          });

          slide.addText(conclusionText, {
            x: 1.15,
            y: 1.80,
            w: 10.9,
            h: 0.72,
            fontSize: 12.2,
            bold: true,
            color: C.white,
            margin: 0,
            fit: "shrink",
          });

          // Focus areas strip
          slide.addShape(pptx.ShapeType.roundRect, {
            x: 0.85,
            y: 3.08,
            w: 11.6,
            h: 0.62,
            rectRadius: 0.05,
            fill: { color: C.goldPale },
            line: { color: C.gold, width: 0.7 },
          });

          slide.addText("Selected Focus Areas", {
            x: 1.15,
            y: 3.35,
            w: 2.4,
            h: 0.12,
            fontSize: 8.6,
            bold: true,
            color: C.gold,
            margin: 0,
          });

          slide.addText(focusText, {
            x: 3.35,
            y: 3.33,
            w: 8.75,
            h: 0.13,
            fontSize: 8.6,
            bold: true,
            color: C.text,
            margin: 0,
            fit: "shrink",
          });

          const cards = [
            [
              "1",
              "Validate",
              "Confirm facts, goals, tax assumptions, and client data",
              C.bluePale,
              C.blue,
            ],
            [
              "2",
              "Coordinate",
              "Review with CPA, estate attorney, and portfolio team",
              C.tealPale,
              C.teal,
            ],
            [
              "3",
              "Finalize",
              "Approve modules, sequencing, and final proposal materials",
              C.goldPale,
              C.gold,
            ],
          ];

          cards.forEach((cardItem, index) => {
            const x = 0.85 + index * 3.95;

            slide.addShape(pptx.ShapeType.roundRect, {
              x,
              y: 4.25,
              w: 3.55,
              h: 1.28,
              rectRadius: 0.06,
              fill: { color: cardItem[3] },
              line: { color: cardItem[4], width: 0.75 },
            });

            slide.addShape(pptx.ShapeType.ellipse, {
              x: x + 0.22,
              y: 4.55,
              w: 0.42,
              h: 0.42,
              fill: { color: cardItem[4] },
              line: { color: cardItem[4] },
            });

            slide.addText(cardItem[0], {
              x: x + 0.37,
              y: 4.68,
              w: 0.1,
              h: 0.1,
              fontSize: 8,
              bold: true,
              color: C.white,
              margin: 0,
            });

            slide.addText(cardItem[1], {
              x: x + 0.8,
              y: 4.52,
              w: 2.4,
              h: 0.14,
              fontSize: 11,
              bold: true,
              color: cardItem[4],
              margin: 0,
            });

            slide.addText(cardItem[2], {
              x: x + 0.8,
              y: 4.86,
              w: 2.45,
              h: 0.28,
              fontSize: 10,
              color: C.text,
              margin: 0,
              fit: "shrink",
            });
          });

          slide.addText(
            "Conclusion: move forward only after advisor, client, tax, legal, and implementation review.",
            {
              x: 0.85,
              y: 6.25,
              w: 11.6,
              h: 0.16,
              fontSize: 9.4,
              bold: true,
              color: C.navy,
              align: "center",
              margin: 0,
            }
          );

          footer(slide);
        }

        if (modules.nextSteps) {
          slide = pptx.addSlide();

          addNextStepsVisual(
            slide,
            pptx,
            {
              ...data,
              clientName: name || data.clientName || "Client Household",
              ticker: data.ticker,
              riskProfile: data.riskProfile || riskProfileLabel || selectedRiskProfile || "",
              selectedPortfolioStrategy:
                data.selectedPortfolioStrategy ||
                selectedPortfolioStrategyLabel ||
                "",
              nextSteps: [
                "Confirm client facts, goals, constraints, and extracted financial data.",
                "Finalize selected planning modules and proposal sections.",
                "Review tax, estate, legal, and implementation assumptions with specialists.",
                "Confirm recommended action plan and sequencing.",
                "Generate final Word document and PowerPoint for client presentation.",
              ],
            },
            C,
            {
              fmtM,
              fmtK,
              pct,
              title,
              footer,
            }
          );

          footer(slide);
        }
      }


      }


        } // END_SKIP_IF_NO_RISK_PROFILE
      // FINAL_BOTTOM_SUGGESTED_NEXT_STEPS
      // Final client-facing next steps intentionally placed at the bottom.
      slide = pptx.addSlide();
      title(
        slide,
        "NEXT STEPS",
        "Suggested Next Steps",
        "Recommended follow-up items before final implementation."
      );

      const finalSteps = [
        ["1", "Confirm client facts", "Validate goals, risk profile, time horizon, liquidity needs, and extracted financial data."],
        ["2", "Finalize module selections", "Confirm which planning modules and strategy sections should remain in the final proposal."],
        ["3", "Review with specialists", "Coordinate tax, legal, estate, options, and portfolio implementation review as needed."],
        ["4", "Approve action plan", "Confirm recommended portfolio strategy, allocation profile, and implementation sequencing."],
        ["5", "Prepare final proposal", "Generate the final Word document and PowerPoint for client presentation."],
      ];

      finalSteps.forEach((item, i) => {
        const y = 1.8 + i * 0.82;

        slide.addShape(pptx.ShapeType.ellipse, {
          x: 0.85,
          y,
          w: 0.42,
          h: 0.42,
          fill: { color: C.navy },
          line: { color: C.navy },
        });

        slide.addText(item[0], {
          x: 0.99,
          y: y + 0.12,
          w: 0.12,
          h: 0.1,
          fontSize: 8,
          bold: true,
          color: C.white,
          margin: 0,
        });

        slide.addText(item[1], {
          x: 1.45,
          y: y,
          w: 3.2,
          h: 0.18,
          fontSize: 11,
          bold: true,
          color: C.navy,
          margin: 0,
        });

        slide.addText(item[2], {
          x: 1.45,
          y: y + 0.27,
          w: 9.6,
          h: 0.18,
          fontSize: 8.8,
          color: C.text,
          margin: 0,
        });

        slide.addShape(pptx.ShapeType.line, {
          x: 1.45,
          y: y + 0.6,
          w: 10.25,
          h: 0,
          line: { color: C.border, width: 0.6 },
        });
      });

      footer(slide);


  const result = await pptx.write({ outputType: "blob" });
  const blob =
    result instanceof Blob
      ? result
      : new Blob([result], {
          type: "application/vnd.openxmlformats-officedocument.presentationml.presentation",
        });
  return blob;
}
