/* ============================================================
   api.js  —  load via GitHub raw URL on every Blogger page
   GitHub raw URL format:
   https://raw.githubusercontent.com/YOUR_USERNAME/YOUR_REPO/main/api.js
   ============================================================ */

const BASE_API = "https://all-in-one.stockmarketsinindia.workers.dev/api/marketMovers";

/* ---------- route map: Blogger page path -> endpoint config ----------
   Slugs taken directly from your existing PAGE_MAP — no changes needed. */
const ROUTE_MAP = {
  "/p/nse-gainers.html"    : { path: "/nse/gainers",    type: "equity_nse", label: "NSE Gainers",        exchange: "NSE" },
  "/p/nse-losers.html"     : { path: "/nse/losers",     type: "equity_nse", label: "NSE Losers",         exchange: "NSE" },
  "/p/nse-volumes.html"    : { path: "/nse/volume",     type: "equity_nse", label: "NSE Volume Toppers", exchange: "NSE" },
  "/p/bse-gainers.html"    : { path: "/bse/gainers",    type: "equity_bse", label: "BSE Gainers",        exchange: "BSE" },
  "/p/bse-losers.html"     : { path: "/bse/losers",     type: "equity_bse", label: "BSE Losers",         exchange: "BSE" },
  "/p/bse-volumes.html"    : { path: "/bse/volume",     type: "equity_bse", label: "BSE Volume Toppers", exchange: "BSE" },
  "/p/futstk-gainers.html" : { path: "/futstk/gainers", type: "future",     label: "Future Gainers"      },
  "/p/futstk-losers.html"  : { path: "/futstk/losers",  type: "future",     label: "Future Losers"       },
  "/p/optstk-gainers.html" : { path: "/optstk/gainers", type: "option",     label: "Option Gainers"      },
  "/p/optstk-losers.html"  : { path: "/optstk/losers",  type: "option",     label: "Option Losers"       },
};

/* ============================================================
   Column definitions per data type
   key = exact JSON field name from your API
   fmt = "num" | "pct" | "chg" | "vol" | "val" | "str"
   ============================================================ */
const COLS = {
  /* NSE equity — closing price key is "close_price" (lowercase c) */
  equity_nse: [
    { key: "symbol",      label: "Symbol",      fmt: "str" },
    { key: "co_name",     label: "Company",     fmt: "str" },
    { key: "close_price", label: "LTP",         fmt: "num" },
    { key: "perchg",      label: "Chg %",       fmt: "chg" },
    { key: "netchg",      label: "Net Chg",     fmt: "chg" },
    { key: "Open_Price",  label: "Open",        fmt: "num" },
    { key: "high_price",  label: "High",        fmt: "num" },
    { key: "low_price",   label: "Low",         fmt: "num" },
    { key: "PrevClose",   label: "Prev Close",  fmt: "num" },
    { key: "vol_traded",  label: "Volume",      fmt: "vol" },
    { key: "val_traded",  label: "Value (Cr)",  fmt: "val" },
    { key: "52WeekHigh",  label: "52W High",    fmt: "num" },
    { key: "52WeekLow",   label: "52W Low",     fmt: "num" },
  ],
  /* BSE equity — closing price key is "Closing_price" (capital C) */
  equity_bse: [
    { key: "symbol",        label: "Symbol",      fmt: "str" },
    { key: "co_name",       label: "Company",     fmt: "str" },
    { key: "Close_price", label: "LTP",         fmt: "num" },
    { key: "perchg",        label: "Chg %",       fmt: "chg" },
    { key: "netchg",        label: "Net Chg",     fmt: "chg" },
    { key: "Open_Price",    label: "Open",        fmt: "num" },
    { key: "high_price",    label: "High",        fmt: "num" },
    { key: "low_price",     label: "Low",         fmt: "num" },
    { key: "PrevClose",     label: "Prev Close",  fmt: "num" },
    { key: "vol_traded",    label: "Volume",      fmt: "vol" },
    { key: "val_traded",    label: "Value (Cr)",  fmt: "val" },
    { key: "52WeekHigh",    label: "52W High",    fmt: "num" },
    { key: "52WeekLow",     label: "52W Low",     fmt: "num" },
  ],
  future: [
    { key: "Symbol",      label: "Symbol",      fmt: "str" },
    { key: "ExpDate",     label: "Expiry",      fmt: "str" },
    { key: "LTP",         label: "LTP",         fmt: "num" },
    { key: "FaOchg",      label: "Chg %",       fmt: "chg" },
    { key: "Volume",      label: "Volume",      fmt: "vol" },
    { key: "Value",       label: "Value (Cr)",  fmt: "val" },
    { key: "FaOQtyDiff",  label: "OI Qty Diff", fmt: "vol" },
    { key: "FaOQtyChg",   label: "OI Qty Chg%", fmt: "chg" },
    { key: "OIdiff",      label: "OI Diff",     fmt: "vol" },
    { key: "OIchg",       label: "OI Chg%",     fmt: "chg" },
  ],
  option: [
    { key: "Symbol",      label: "Symbol",      fmt: "str" },
    { key: "OPTTYPE",     label: "Type",        fmt: "str" },
    { key: "StrikePrice", label: "Strike",      fmt: "num" },
    { key: "ExpDate",     label: "Expiry",      fmt: "str" },
    { key: "LTP",         label: "LTP",         fmt: "num" },
    { key: "FaOchg",      label: "Chg %",       fmt: "chg" },
    { key: "Volume",      label: "Volume",      fmt: "vol" },
    { key: "Value",       label: "Value (Cr)",  fmt: "val" },
    { key: "FaOQtyDiff",  label: "OI Qty Diff", fmt: "vol" },
    { key: "OIchg",       label: "OI Chg%",     fmt: "chg" },
  ],
};

/* ============================================================
   Formatters
   ============================================================ */
function fmtVal(value, type) {
  if (value === null || value === undefined || value === "") return "—";
  const n = parseFloat(value);
  switch (type) {
    case "num": return isNaN(n) ? String(value).trim()
                                : n.toLocaleString("en-IN", { minimumFractionDigits: 2, maximumFractionDigits: 2 });
    case "vol": return isNaN(n) ? String(value).trim()
                                : Math.round(n).toLocaleString("en-IN");
    case "val": return isNaN(n) ? String(value).trim() : n.toFixed(2);
    case "str": return String(value).trim();
    default:    return String(value).trim();
  }
}

function chgCell(value) {
  const n = parseFloat(value);
  if (isNaN(n)) return "<td>—</td>";
  const cls  = n > 0 ? "sw-pos" : n < 0 ? "sw-neg" : "sw-neu";
  const sign = n > 0 ? "▲ "    : n < 0 ? "▼ "     : "";
  return `<td class="${cls}">${sign}${Math.abs(n).toFixed(2)}</td>`;
}

/* ============================================================
   Build a full <table> string from rows + col config
   ============================================================ */
function buildTable(rows, cols) {
  const thead = cols.map(c => `<th>${c.label}</th>`).join("");
  const tbody = rows.map(row => {
    const cells = cols.map(col => {
      if (col.fmt === "chg") return chgCell(row[col.key]);
      const v   = fmtVal(row[col.key], col.fmt);
      const cls = (col.key === "symbol" || col.key === "Symbol") ? ' class="sw-sym"' : "";
      return `<td${cls}>${v}</td>`;
    }).join("");
    return `<tr>${cells}</tr>`;
  }).join("");
  return `<table class="sw-table"><thead><tr>${thead}</tr></thead><tbody>${tbody}</tbody></table>`;
}

/* ============================================================
   Single-page widget  →  auto-detects route from URL
   Call: loadStockTable("stock-widget")
   ============================================================ */
async function loadStockTable(containerId) {
  const el = document.getElementById(containerId);
  if (!el) return;

  const route = ROUTE_MAP[window.location.pathname];
  if (!route) {
    el.innerHTML = `<p class="sw-err">No route configured for: ${window.location.pathname}</p>`;
    return;
  }

  el.innerHTML = `<div class="sw-loading"><span class="sw-spinner"></span> Loading ${route.label}…</div>`;

  try {
    const res  = await fetch(BASE_API + route.path);
    if (!res.ok) throw new Error("HTTP " + res.status);
    const json = await res.json();
    if (!json.success || !Array.isArray(json.data) || !json.data.length) throw new Error("Empty data");

    const cols    = COLS[route.type];
    const exch    = route.exchange ? `<span class="sw-exch">${route.exchange}</span>` : "";
    const updTime = json.data[0].Upd_Time || json.data[0].UpdTime || "";

    el.innerHTML = `
      <div class="sw-header">
        <span class="sw-title">${route.label} ${exch}</span>
        ${updTime ? `<span class="sw-time">Updated: ${updTime}</span>` : ""}
        <span class="sw-count">${json.data.length} stocks</span>
      </div>
      <div class="sw-scroll">${buildTable(json.data, cols)}</div>`;
  } catch (err) {
    el.innerHTML = `<p class="sw-err">Could not load data — ${err.message}</p>`;
  }
}

/* ============================================================
   Homepage multi-section widget
   Call: loadHomepageWidget("stock-widget-home", [...sections])
   Each section: { path, type, label, exchange?, limit? }
   Use the same path values as ROUTE_MAP (right-hand side)
   ============================================================ */
async function loadHomepageWidget(containerId, sections) {
  const el = document.getElementById(containerId);
  if (!el) return;
  el.innerHTML = `<div class="sw-loading"><span class="sw-spinner"></span> Loading market overview…</div>`;

  const fetches = sections.map(s =>
    fetch(BASE_API + s.path)
      .then(r => r.json())
      .then(json => ({ ...s, rows: (json.data || []).slice(0, s.limit || 5), ok: json.success }))
      .catch(() => ({ ...s, rows: [], ok: false }))
  );
  const results = await Promise.all(fetches);

  const html = results.map(r => {
    const exch = r.exchange ? `<span class="sw-exch">${r.exchange}</span>` : "";
    const inner = r.ok && r.rows.length
      ? `<div class="sw-scroll">${buildTable(r.rows, COLS[r.type])}</div>`
      : `<p class="sw-err" style="padding:12px 14px">No data</p>`;
    return `
      <div class="sw-home-section">
        <div class="sw-header">
          <span class="sw-title">${r.label} ${exch}</span>
          <a class="sw-more" href="#">View all →</a>
        </div>
        ${inner}
      </div>`;
  }).join("");

  el.innerHTML = `<div class="sw-home-grid">${html}</div>`;
}

/* ============================================================
   Auto-init: if #stock-widget exists, load it automatically
   ============================================================ */
document.addEventListener("DOMContentLoaded", () => {
  if (document.getElementById("stock-widget"))
    loadStockTable("stock-widget");
});
