// Shared number/currency formatters used across PPT and Word generation.

export function cleanNum(n, decimals = 2) {
  if (!Number.isFinite(n)) return "N/A";
  const rounded = Number(n.toFixed(decimals));
  return Number.isInteger(rounded) ? String(rounded) : String(rounded);
}

export function fmtM(num) {
  if (!Number.isFinite(num)) return "N/A";
  if (num >= 1000) return `$${cleanNum(num / 1000, 2)}B`;
  return `$${cleanNum(num, 2)}M`;
}

export function fmtK(num) {
  if (!Number.isFinite(num)) return "N/A";
  if (num >= 1) return fmtM(num);
  return `$${Math.round(num * 1000).toLocaleString()}K`;
}

export function fmtDollar(num) {
  if (!Number.isFinite(num)) return "N/A";
  return `$${Math.round(num).toLocaleString()}`;
}

export function pct(num) {
  return `${num.toFixed(1)}%`;
}
