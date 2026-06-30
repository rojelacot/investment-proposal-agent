function toMillions(value) {
  const n = Number(String(value ?? "").replace(/[$,%\s,]/g, ""));
  if (!Number.isFinite(n)) return 0;
  return n > 1000 ? n / 1000000 : n;
}

function moneyM(millions) {
  const n = Number(millions || 0);
  if (!Number.isFinite(n)) return "Review";
  if (Math.abs(n) >= 1000) return `$${(n / 1000).toFixed(2)}B`;
  return `$${n.toFixed(n >= 100 ? 0 : 1)}M`;
}

function fullMoney(millions) {
  const n = Number(millions || 0);
  if (!Number.isFinite(n)) return "Review";
  return `$${Math.round(n * 1000000).toLocaleString()}`;
}

function pct(rate) {
  const n = Number(rate || 0);
  return `${(n > 1 ? n : n * 100).toFixed(0)}%`;
}

// Shared mini helpers for slide HTML
function _statCard(accentColor, bgColor, label, value, sub) {
  return `<div style="position:relative;background:${bgColor};border:0.5px solid #D5DAE5;border-radius:7px;padding:10px 14px 10px 20px;box-shadow:0 1px 3px rgba(0,0,0,0.05);overflow:hidden;">
    <div style="position:absolute;top:0;left:0;width:5px;height:100%;background:${accentColor};border-radius:7px 0 0 7px;"></div>
    <div style="font-size:19px;font-weight:700;color:${accentColor};font-family:Georgia,serif;line-height:1.1;">${value}</div>
    <div style="font-size:7.5px;font-weight:700;color:#6E7E9A;letter-spacing:0.8px;text-transform:uppercase;margin-top:5px;">${label}</div>
    ${sub ? `<div style="font-size:10px;color:#6E7E9A;margin-top:3px;">${sub}</div>` : ''}
  </div>`;
}

function _slideHeader(sectionLabel, heading) {
  return `
    <div style="height:8px;background:#1A2744;position:relative;">
      <div style="position:absolute;top:0;left:0;width:210px;height:8px;background:#D4A845;"></div>
    </div>
    <div style="padding:20px 54px 0;">
      <div style="font-size:8.5px;font-weight:700;color:#3A6BBF;letter-spacing:2px;text-transform:uppercase;margin-bottom:5px;">${sectionLabel}</div>
      <div style="font-size:26px;font-weight:700;color:#1A2744;font-family:Georgia,serif;line-height:1.15;">${heading}</div>
      <div style="display:flex;align-items:center;margin-top:9px;margin-bottom:16px;">
        <div style="width:90px;height:2.5px;background:#D4A845;"></div>
        <div style="flex:1;height:1px;background:#D5DAE5;"></div>
      </div>
    </div>`;
}

export function buildEstateSlideHtml(input = {}) {
  const data = input.data || {};
  const isFuture = !!input.isFuture;

  const clientName = data.clientName || "Client Household";
  const personA =
    data.clientFirstName ||
    data.primaryClientName ||
    clientName.split(" ")[0] ||
    "Client";

  const heirs = data.heirs || "Heirs";

  const managedAssets = toMillions(data.managedAssets || data.investableAssets);
  const realEstate = toMillions(data.realEstateValue || data.realEstateHoldings);
  const otherAssets = toMillions(data.otherAssets || data.otherPrivateAssets);
  const totalEstate =
    toMillions(data.netWorth || data.totalEstate) ||
    managedAssets + realEstate + otherAssets;

  const exemption = toMillions(data.estateTaxExemption || data.estateExemption || 30);
  const rawTaxRate = Number(data.estateTaxRate ?? 0.4);
  const taxRate = rawTaxRate > 1 ? rawTaxRate / 100 : rawTaxRate;

  const projectionYears = Number(data.estateProjectionYears || data.projectionYears || 37);
  const rawGrowth = Number(data.estateGrowthRate || data.growthRate || 0.07);
  const growthRate = rawGrowth > 1 ? rawGrowth / 100 : rawGrowth;

  const taxableEstate = Math.max(0, totalEstate - exemption);
  const estateTax = taxableEstate * taxRate;
  const toHeirs = Math.max(0, totalEstate - estateTax);
  const effectiveRate = totalEstate > 0 ? estateTax / totalEstate : 0;

  // ── TODAY slide ───────────────────────────────────────────────────────────
  if (!isFuture) {
    const hasBreakdown = managedAssets > 0 || realEstate > 0 || otherAssets > 0;
    return `<div style="width:1280px;height:720px;background:#FFFFFF;font-family:-apple-system,BlinkMacSystemFont,'Segoe UI',Arial,sans-serif;color:#1A2030;box-sizing:border-box;overflow:hidden;position:relative;">

      ${_slideHeader("ESTATE PLANNING", "Estate Value Today")}

      <div style="padding:0 54px;">

        <!-- 4 stat boxes -->
        <div style="display:grid;grid-template-columns:repeat(4,1fr);gap:14px;margin-bottom:18px;">
          ${_statCard("#1A2744","#FFFFFF","Total Estate",fullMoney(totalEstate),"")}
          ${_statCard("#C94F3A","#FAECE7","Taxable Amount",fullMoney(taxableEstate),`Over ${moneyM(exemption)} exemption`)}
          ${_statCard("#C94F3A","#FAECE7","Est. Estate Tax","("+fullMoney(estateTax)+")",`${(effectiveRate*100).toFixed(1)}% effective rate`)}
          ${_statCard("#1E7A6E","#DFF1EE","Net to Heirs",fullMoney(toHeirs),"Without further planning")}
        </div>

        <!-- Two-column content -->
        <div style="display:grid;grid-template-columns:1.1fr 0.9fr;gap:16px;">

          <!-- Left: estate composition -->
          <div style="background:#F7F9FC;border:1px solid #D5DAE5;border-radius:8px;padding:16px 20px;">
            <div style="font-size:8.5px;font-weight:700;color:#3A6BBF;letter-spacing:1px;text-transform:uppercase;margin-bottom:8px;">ESTATE COMPOSITION</div>
            <div style="height:1px;background:#D5DAE5;margin-bottom:12px;"></div>
            ${managedAssets > 0 ? `<div style="display:flex;justify-content:space-between;align-items:center;padding:6px 0;border-bottom:1px solid #EEF1F6;">
              <span style="font-size:13px;color:#1A2030;">Managed / Investable Assets</span>
              <span style="font-size:14px;font-weight:700;color:#1A2744;font-family:Georgia,serif;">${fullMoney(managedAssets)}</span>
            </div>` : ''}
            ${realEstate > 0 ? `<div style="display:flex;justify-content:space-between;align-items:center;padding:6px 0;border-bottom:1px solid #EEF1F6;">
              <span style="font-size:13px;color:#1A2030;">Real Estate Holdings</span>
              <span style="font-size:14px;font-weight:700;color:#1A2744;font-family:Georgia,serif;">${fullMoney(realEstate)}</span>
            </div>` : ''}
            ${otherAssets > 0 ? `<div style="display:flex;justify-content:space-between;align-items:center;padding:6px 0;border-bottom:1px solid #EEF1F6;">
              <span style="font-size:13px;color:#1A2030;">Other Private Assets</span>
              <span style="font-size:14px;font-weight:700;color:#1A2744;font-family:Georgia,serif;">${fullMoney(otherAssets)}</span>
            </div>` : ''}
            ${!hasBreakdown ? `<div style="font-size:13px;color:#6E7E9A;font-style:italic;">Asset breakdown not available — review with advisor.</div>` : ''}
            <div style="display:flex;justify-content:space-between;align-items:center;padding:8px 0 0;">
              <span style="font-size:13px;font-weight:700;color:#1A2744;">Total Estate</span>
              <span style="font-size:16px;font-weight:700;color:#1A2744;font-family:Georgia,serif;">${fullMoney(totalEstate)}</span>
            </div>
          </div>

          <!-- Right: tax flow + planning note -->
          <div style="display:flex;flex-direction:column;gap:14px;">

            <!-- Tax flow -->
            <div style="background:#FFFFFF;border:1px solid #D5DAE5;border-radius:8px;padding:14px 18px;">
              <div style="font-size:8.5px;font-weight:700;color:#3A6BBF;letter-spacing:1px;text-transform:uppercase;margin-bottom:10px;">ESTATE TAX OVERVIEW</div>
              <div style="display:flex;align-items:center;justify-content:space-between;gap:6px;">
                <div style="text-align:center;flex:1;">
                  <div style="font-size:11px;color:#6E7E9A;margin-bottom:4px;">${personA}'s Estate</div>
                  <div style="font-size:16px;font-weight:700;color:#1A2744;font-family:Georgia,serif;">${fullMoney(totalEstate)}</div>
                </div>
                <div style="font-size:20px;color:#D5DAE5;padding-bottom:4px;">→</div>
                <div style="text-align:center;flex:1;">
                  <div style="font-size:11px;color:#6E7E9A;margin-bottom:4px;">Est. Tax</div>
                  <div style="font-size:16px;font-weight:700;color:#C94F3A;font-family:Georgia,serif;">(${fullMoney(estateTax)})</div>
                </div>
                <div style="font-size:20px;color:#D5DAE5;padding-bottom:4px;">→</div>
                <div style="text-align:center;flex:1;">
                  <div style="font-size:11px;color:#6E7E9A;margin-bottom:4px;">Net to ${heirs}</div>
                  <div style="font-size:16px;font-weight:700;color:#1E7A6E;font-family:Georgia,serif;">${fullMoney(toHeirs)}</div>
                </div>
              </div>
              <div style="margin-top:8px;font-size:10px;color:#6E7E9A;text-align:center;">
                ${pct(taxRate)} on ${moneyM(taxableEstate)} taxable · ${(effectiveRate*100).toFixed(1)}% effective rate · exemption: ${moneyM(exemption)}
              </div>
            </div>

            <!-- Planning opportunity -->
            <div style="background:#F5EDDA;border:1px solid #D4A845;border-radius:8px;padding:14px 18px;flex:1;">
              <div style="font-size:8.5px;font-weight:700;color:#B8892A;letter-spacing:1px;text-transform:uppercase;margin-bottom:8px;">PLANNING OPPORTUNITY</div>
              <div style="font-size:12px;color:#1A2030;line-height:1.55;">
                Coordinated estate planning — trusts, GRATs, charitable vehicles, and beneficiary alignment — can significantly reduce the estimated estate tax and increase the net amount transferred to heirs.
              </div>
            </div>

          </div>
        </div>
      </div>

      <!-- Footer line -->
      <div style="position:absolute;bottom:14px;left:54px;right:54px;height:1px;background:#D5DAE5;"></div>
    </div>`;
  }

  // ── FUTURE / PROJECTED slide ──────────────────────────────────────────────
  const futureTotal = totalEstate;
  const futureTaxable = Math.max(0, futureTotal - exemption);
  const futureTax = futureTaxable * taxRate;
  const futureHeirs = Math.max(0, futureTotal - futureTax);
  const futureEffective = futureTotal > 0 ? futureTax / futureTotal : 0;


  return `<div style="width:1280px;height:720px;background:#FFFFFF;font-family:-apple-system,BlinkMacSystemFont,'Segoe UI',Arial,sans-serif;color:#1A2030;box-sizing:border-box;overflow:hidden;position:relative;">

    ${_slideHeader("ESTATE PLANNING", `Estate Value — Projected in ${projectionYears} Years`)}

    <div style="padding:0 54px;">

      <!-- 4 stat boxes -->
      <div style="display:grid;grid-template-columns:repeat(4,1fr);gap:14px;margin-bottom:18px;">
        ${_statCard("#1A2744","#FFFFFF","Projected Estate",fullMoney(futureTotal),`${projectionYears} yrs · ${(growthRate*100).toFixed(1)}%/yr`)}
        ${_statCard("#C94F3A","#FAECE7","Taxable Amount",fullMoney(futureTaxable),`Over ${moneyM(exemption)} exemption`)}
        ${_statCard("#C94F3A","#FAECE7","Est. Estate Tax","("+fullMoney(futureTax)+")",`${(futureEffective*100).toFixed(1)}% effective rate`)}
        ${_statCard("#1E7A6E","#DFF1EE","Net to Heirs",fullMoney(futureHeirs),"Without further planning")}
      </div>

      <!-- Two-column content -->
      <div style="display:grid;grid-template-columns:1.1fr 0.9fr;gap:16px;">

        <!-- Left: tax breakdown + assumptions -->
        <div style="background:#F7F9FC;border:1px solid #D5DAE5;border-radius:8px;padding:16px 20px;">
          <div style="font-size:8.5px;font-weight:700;color:#3A6BBF;letter-spacing:1px;text-transform:uppercase;margin-bottom:8px;">PROJECTED ESTATE BREAKDOWN</div>
          <div style="height:1px;background:#D5DAE5;margin-bottom:12px;"></div>

          <div style="display:flex;justify-content:space-between;align-items:center;padding:6px 0;border-bottom:1px solid #EEF1F6;">
            <span style="font-size:13px;color:#1A2030;">Projected Total Estate</span>
            <span style="font-size:14px;font-weight:700;color:#1A2744;font-family:Georgia,serif;">${fullMoney(futureTotal)}</span>
          </div>
          <div style="display:flex;justify-content:space-between;align-items:center;padding:6px 0;border-bottom:1px solid #EEF1F6;">
            <span style="font-size:13px;color:#1A2030;">Federal Estate Tax Exemption</span>
            <span style="font-size:14px;font-weight:700;color:#6E7E9A;font-family:Georgia,serif;">(${moneyM(exemption)})</span>
          </div>
          <div style="display:flex;justify-content:space-between;align-items:center;padding:6px 0;border-bottom:1px solid #EEF1F6;">
            <span style="font-size:13px;color:#1A2030;">Taxable Amount</span>
            <span style="font-size:14px;font-weight:700;color:#C94F3A;font-family:Georgia,serif;">${fullMoney(futureTaxable)}</span>
          </div>
          <div style="display:flex;justify-content:space-between;align-items:center;padding:6px 0;border-bottom:1px solid #EEF1F6;">
            <span style="font-size:13px;color:#1A2030;">Est. Estate Tax (${pct(taxRate)})</span>
            <span style="font-size:14px;font-weight:700;color:#C94F3A;font-family:Georgia,serif;">(${fullMoney(futureTax)})</span>
          </div>
          <div style="display:flex;justify-content:space-between;align-items:center;padding:8px 0 0;">
            <span style="font-size:13px;font-weight:700;color:#1A2744;">Net to Heirs</span>
            <span style="font-size:16px;font-weight:700;color:#1E7A6E;font-family:Georgia,serif;">${fullMoney(futureHeirs)}</span>
          </div>
        </div>

        <!-- Right: planning opportunity + assumptions -->
        <div style="display:flex;flex-direction:column;gap:14px;">

          <div style="background:#F5EDDA;border:1px solid #D4A845;border-radius:8px;padding:14px 18px;flex:1;">
            <div style="font-size:8.5px;font-weight:700;color:#B8892A;letter-spacing:1px;text-transform:uppercase;margin-bottom:8px;">PLANNING OPPORTUNITY</div>
            <div style="font-size:12px;color:#1A2030;line-height:1.55;">
              Coordinated implementation of GRATs, irrevocable trust structures, and charitable vehicles may substantially reduce the projected estate tax. Compounding materially increases the cost of delayed planning.
            </div>
          </div>

          <div style="background:#1A2744;border-radius:8px;padding:13px 18px;">
            <div style="font-size:8.5px;font-weight:700;color:#D4A845;letter-spacing:1px;text-transform:uppercase;margin-bottom:6px;">ASSUMPTIONS</div>
            <div style="font-size:11px;color:#B0BBC8;line-height:1.6;">
              ${moneyM(futureTotal / Math.pow(1 + growthRate, projectionYears) || futureTotal)} starting value · ${(growthRate*100).toFixed(1)}%/yr growth · ${projectionYears} years · ${moneyM(exemption)} exemption · ${pct(taxRate)} tax on taxable amount.
            </div>
          </div>

        </div>
      </div>
    </div>

    <!-- Footer line -->
    <div style="position:absolute;bottom:14px;left:54px;right:54px;height:1px;background:#D5DAE5;"></div>
  </div>`;
}
