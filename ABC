<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>Maruti Suzuki — Executive Analytics Dashboard</title>
<link href="https://fonts.googleapis.com/css2?family=Cormorant+Garamond:wght@300;400;600;700&family=DM+Sans:wght@300;400;500;600&display=swap" rel="stylesheet">
<script src="https://cdnjs.cloudflare.com/ajax/libs/Chart.js/4.4.1/chart.umd.min.js">
</script>
<script src="https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js">
</script>
<style>
  :root {
    --bg: #080d1a;
    --bg2: #0e1525;
    --card: #111929;
    --card2: #162035;
    --border: #1e2d4a;
    --gold: #c8a84b;
    --gold2: #e8c96a;
    --gold-dim: rgba(200,168,75,0.15);
    --teal: #2dd4bf;
    --coral: #f87171;
    --lavender: #a78bfa;
    --text: #e8edf5;
    --text2: #8fa3c8;
    --text3: #4d6080;
    --green: #4ade80;
    --red: #f87171;
  }
  * { margin: 0; padding: 0; box-sizing: border-box; }
  body { background: var(--bg); color: var(--text); font-family: 'DM Sans', sans-serif; font-size: 14px; min-height: 100vh; }

  /* TOP HEADER */
  .top-header {
    background: linear-gradient(135deg, #0a1428 0%, #0e1e38 50%, #080f20 100%);
    border-bottom: 1px solid var(--border);
    padding: 0 32px;
    display: flex; align-items: center; justify-content: space-between;
    height: 68px; position: sticky; top: 0; z-index: 100;
    box-shadow: 0 4px 24px rgba(0,0,0,0.4);
  }
  .brand { display: flex; align-items: center; gap: 16px; }
  .brand-icon { width: 38px; height: 38px; border-radius: 8px; background: var(--gold-dim); border: 1px solid var(--gold); display: grid; place-items: center; font-size: 18px; }
  .brand-title { font-family: 'Cormorant Garamond', serif; font-size: 22px; font-weight: 600; letter-spacing: 0.5px; color: var(--text); }
  .brand-sub { font-size: 11px; color: var(--text3); letter-spacing: 1.5px; text-transform: uppercase; margin-top: 1px; }
  .header-meta { display: flex; align-items: center; gap: 24px; }
  .meta-pill { display: flex; flex-direction: column; align-items: flex-end; }
  .meta-val { font-size: 17px; font-weight: 600; color: var(--gold2); }
  .meta-lbl { font-size: 10px; color: var(--text3); text-transform: uppercase; letter-spacing: 1px; }
  .upload-btn {
    background: var(--gold-dim); border: 1px solid var(--gold); color: var(--gold2);
    padding: 8px 18px; border-radius: 6px; font-size: 12px; font-weight: 500;
    cursor: pointer; letter-spacing: 0.5px; transition: all 0.2s;
    display: flex; align-items: center; gap: 8px;
  }
  .upload-btn:hover { background: rgba(200,168,75,0.25); }
  #fileInput { display: none; }

  /* KPI BAR */
  .kpi-bar { display: grid; grid-template-columns: repeat(6, 1fr); gap: 12px; padding: 20px 32px; background: var(--bg2); border-bottom: 1px solid var(--border); }
  .kpi-card {
    background: var(--card); border: 1px solid var(--border); border-radius: 10px;
    padding: 14px 16px; position: relative; overflow: hidden; transition: transform 0.2s;
  }
  .kpi-card:hover { transform: translateY(-2px); border-color: var(--gold); }
  .kpi-card::before { content: ''; position: absolute; top: 0; left: 0; right: 0; height: 2px; }
  .kpi-card.gold::before { background: var(--gold); }
  .kpi-card.teal::before { background: var(--teal); }
  .kpi-card.coral::before { background: var(--coral); }
  .kpi-card.lavender::before { background: var(--lavender); }
  .kpi-card.green::before { background: var(--green); }
  .kpi-card.orange::before { background: #fb923c; }
  .kpi-lbl { font-size: 10px; color: var(--text3); text-transform: uppercase; letter-spacing: 1.2px; margin-bottom: 6px; }
  .kpi-val { font-size: 22px; font-weight: 600; color: var(--text); line-height: 1; }
  .kpi-val span { font-size: 12px; color: var(--text2); font-weight: 400; }
  .kpi-chg { font-size: 11px; margin-top: 5px; }
  .kpi-chg.up { color: var(--green); }
  .kpi-chg.dn { color: var(--red); }
  .kpi-sub { font-size: 10px; color: var(--text3); margin-top: 3px; }

  /* TABS */
  .tabs-bar { display: flex; align-items: center; gap: 2px; padding: 12px 32px 0; background: var(--bg2); border-bottom: 1px solid var(--border); overflow-x: auto; }
  .tab-btn {
    padding: 10px 20px; font-size: 12px; font-weight: 500; letter-spacing: 0.5px;
    color: var(--text2); border: none; background: transparent; cursor: pointer;
    border-bottom: 2px solid transparent; transition: all 0.2s; white-space: nowrap; border-radius: 4px 4px 0 0;
    text-transform: uppercase;
  }
  .tab-btn:hover { color: var(--text); background: rgba(255,255,255,0.03); }
  .tab-btn.active { color: var(--gold2); border-bottom-color: var(--gold); background: rgba(200,168,75,0.05); }

  /* CONTENT */
  .main { padding: 24px 32px; }
  .tab-panel { display: none; }
  .tab-panel.active { display: block; }

  /* SECTION HEADERS */
  .sec-hdr { margin-bottom: 16px; display: flex; align-items: baseline; gap: 12px; }
  .sec-title { font-family: 'Cormorant Garamond', serif; font-size: 20px; font-weight: 600; color: var(--gold2); }
  .sec-note { font-size: 11px; color: var(--text3); }

  /* GRID LAYOUTS */
  .grid-2 { display: grid; grid-template-columns: 1fr 1fr; gap: 16px; }
  .grid-3 { display: grid; grid-template-columns: 1fr 1fr 1fr; gap: 16px; }
  .grid-21 { display: grid; grid-template-columns: 2fr 1fr; gap: 16px; }
  .grid-12 { display: grid; grid-template-columns: 1fr 2fr; gap: 16px; }
  .mb16 { margin-bottom: 16px; }
  .mb24 { margin-bottom: 24px; }

  /* CHART CARDS */
  .chart-card {
    background: var(--card); border: 1px solid var(--border); border-radius: 12px;
    padding: 20px; position: relative;
  }
  .chart-card.full { grid-column: 1 / -1; }
  .chart-title { font-size: 13px; font-weight: 600; color: var(--text2); text-transform: uppercase; letter-spacing: 0.8px; margin-bottom: 4px; }
  .chart-subtitle { font-size: 11px; color: var(--text3); margin-bottom: 16px; }
  .chart-wrap { position: relative; }
  .chart-wrap canvas { max-height: 280px; }
  .chart-wrap.tall canvas { max-height: 340px; }

  /* COST TABLE */
  .cost-table { width: 100%; border-collapse: collapse; font-size: 12px; }
  .cost-table th { color: var(--text3); font-weight: 500; text-transform: uppercase; font-size: 10px; letter-spacing: 0.8px; padding: 8px 10px; border-bottom: 1px solid var(--border); text-align: right; }
  .cost-table th:first-child { text-align: left; }
  .cost-table td { padding: 7px 10px; border-bottom: 1px solid rgba(30,45,74,0.5); text-align: right; color: var(--text); }
  .cost-table td:first-child { text-align: left; color: var(--text2); }
  .cost-table tr:hover td { background: rgba(255,255,255,0.02); }
  .cost-table tr:last-child td { border-bottom: none; }
  .pct-bar { display: inline-flex; align-items: center; gap: 6px; }
  .bar-bg { width: 60px; height: 4px; background: var(--border); border-radius: 2px; display: inline-block; }
  .bar-fill { height: 4px; border-radius: 2px; display: inline-block; }

  /* METRIC TILES */
  .metric-grid { display: grid; grid-template-columns: repeat(3, 1fr); gap: 10px; }
  .metric-tile { background: var(--card2); border: 1px solid var(--border); border-radius: 8px; padding: 14px; }
  .metric-lbl { font-size: 10px; color: var(--text3); text-transform: uppercase; letter-spacing: 1px; margin-bottom: 5px; }
  .metric-val { font-size: 18px; font-weight: 600; color: var(--text); }
  .metric-chg { font-size: 11px; margin-top: 3px; }

  /* UPLOAD OVERLAY */
  .upload-overlay {
    position: fixed; inset: 0; background: rgba(8,13,26,0.9); z-index: 200;
    display: none; place-items: center;
  }
  .upload-overlay.show { display: grid; }
  .upload-box {
    background: var(--card); border: 2px dashed var(--border); border-radius: 16px;
    padding: 60px; text-align: center; max-width: 460px; width: 90%;
    transition: border-color 0.2s;
  }
  .upload-box.drag { border-color: var(--gold); }
  .upload-icon { font-size: 48px; margin-bottom: 16px; }
  .upload-box h3 { font-family: 'Cormorant Garamond', serif; font-size: 22px; color: var(--gold2); margin-bottom: 8px; }
  .upload-box p { font-size: 13px; color: var(--text2); margin-bottom: 24px; }
  .upload-close { position: absolute; top: 16px; right: 16px; cursor: pointer; color: var(--text2); font-size: 20px; }
  .upload-box label { cursor: pointer; }
  .parse-btn {
    background: var(--gold); color: #080d1a; border: none; padding: 12px 32px;
    border-radius: 8px; font-size: 14px; font-weight: 600; cursor: pointer; width: 100%;
    margin-top: 16px; transition: background 0.2s;
  }
  .parse-btn:hover { background: var(--gold2); }

  /* TOAST */
  .toast { position: fixed; bottom: 24px; right: 24px; background: var(--card2); border: 1px solid var(--gold); border-radius: 8px; padding: 12px 20px; font-size: 13px; color: var(--gold2); z-index: 300; transform: translateY(80px); transition: transform 0.3s; }
  .toast.show { transform: translateY(0); }

  /* DATA NOTE */
  .data-note { font-size: 10px; color: var(--text3); margin-top: 8px; font-style: italic; }

  /* WATERFALL */
  .wf-row { display: flex; align-items: center; gap: 10px; padding: 4px 0; border-bottom: 1px solid rgba(30,45,74,0.4); }
  .wf-label { width: 180px; font-size: 11px; color: var(--text2); flex-shrink: 0; }
  .wf-bar-area { flex: 1; display: flex; align-items: center; gap: 6px; }
  .wf-bar { height: 14px; border-radius: 3px; min-width: 2px; }
  .wf-amount { font-size: 11px; color: var(--text); width: 80px; text-align: right; flex-shrink: 0; }
  .wf-pct { font-size: 10px; color: var(--text3); width: 45px; text-align: right; flex-shrink: 0; }
  .wf-header { display: flex; gap: 10px; padding: 6px 0; border-bottom: 1px solid var(--border); }
  .wf-header span { font-size: 10px; color: var(--text3); text-transform: uppercase; letter-spacing: 1px; }
  .wf-header .l { width: 180px; flex-shrink: 0; }
  .wf-header .r { flex: 1; }
  .wf-header .v { width: 80px; text-align: right; flex-shrink: 0; }
  .wf-header .p { width: 45px; text-align: right; flex-shrink: 0; }

  /* SCROLLBAR */
  ::-webkit-scrollbar { width: 4px; height: 4px; }
  ::-webkit-scrollbar-track { background: var(--bg); }
  ::-webkit-scrollbar-thumb { background: var(--border); border-radius: 4px; }

  /* TREND BADGE */
  .trend { display: inline-flex; align-items: center; gap: 3px; font-size: 11px; padding: 2px 7px; border-radius: 12px; }
  .trend.up { background: rgba(74,222,128,0.12); color: var(--green); }
  .trend.dn { background: rgba(248,113,113,0.12); color: var(--red); }
  .trend.nt { background: rgba(100,120,160,0.12); color: var(--text2); }

  /* RESPONSIVE */
  @media (max-width: 900px) {
    .kpi-bar { grid-template-columns: repeat(3, 1fr); }
    .grid-2, .grid-21, .grid-12, .grid-3 { grid-template-columns: 1fr; }
    .main { padding: 16px; }
    .kpi-bar { padding: 12px 16px; }
    .tabs-bar { padding: 8px 16px 0; }
    .top-header { padding: 0 16px; }
    .header-meta { gap: 12px; }
  }
</style>
</head>
<body>

<!-- TOP HEADER -->
<div class="top-header">
  <div class="brand">
    <div class="brand-icon">🚗</div>
    <div>
      <div class="brand-title">Maruti Suzuki India Ltd</div>
      <div class="brand-sub">Executive Cost Analytics Dashboard &nbsp;·&nbsp; NSE: MARUTI</div>
    </div>
  </div>
  <div class="header-meta">
    <div class="meta-pill">
      <div class="meta-val" id="hdr-price">₹13,314</div>
      <div class="meta-lbl">Current Price</div>
    </div>
    <div class="meta-pill">
      <div class="meta-val" id="hdr-mcap">₹4,18,596 Cr</div>
      <div class="meta-lbl">Market Cap</div>
    </div>
    <button class="upload-btn" onclick="showUpload()">
      <span>⬆</span> Update Data
    </button>
  </div>
</div>

<!-- KPI BAR -->
<div class="kpi-bar">
  <div class="kpi-card gold">
    <div class="kpi-lbl">Revenue FY26</div>
    <div class="kpi-val" id="kpi-rev">₹1,83,316 <span>Cr</span></div>
    <div class="kpi-chg up" id="kpi-rev-chg">▲ 19.9% YoY</div>
    <div class="kpi-sub">vs ₹1,52,913 Cr in FY25</div>
  </div>
  <div class="kpi-card teal">
    <div class="kpi-lbl">Net Profit FY26</div>
    <div class="kpi-val" id="kpi-np">₹14,680 <span>Cr</span></div>
    <div class="kpi-chg up" id="kpi-np-chg">▲ 1.2% YoY</div>
    <div class="kpi-sub">vs ₹14,500 Cr in FY25</div>
  </div>
  <div class="kpi-card coral">
    <div class="kpi-lbl">Raw Mat % of Rev</div>
    <div class="kpi-val" id="kpi-rm">73.6 <span>%</span></div>
    <div class="kpi-chg dn" id="kpi-rm-chg">▲ 250 bps vs FY25</div>
    <div class="kpi-sub">Key cost driver — under pressure</div>
  </div>
  <div class="kpi-card lavender">
    <div class="kpi-lbl">EBITDA Margin</div>
    <div class="kpi-val" id="kpi-opm">—</div>
    <div class="kpi-chg" id="kpi-opm-chg">Q4 FY26: 11.7%</div>
    <div class="kpi-sub">FY26 quarterly avg</div>
  </div>
  <div class="kpi-card green">
    <div class="kpi-lbl">Debt-to-Equity</div>
    <div class="kpi-val" id="kpi-de">0.001 <span>x</span></div>
    <div class="kpi-chg nt" id="kpi-de-chg">Near debt-free</div>
    <div class="kpi-sub">Borrowings: ₹102.5 Cr</div>
  </div>
  <div class="kpi-card orange">
    <div class="kpi-lbl">OCF FY26</div>
    <div class="kpi-val" id="kpi-ocf">₹19,100 <span>Cr</span></div>
    <div class="kpi-chg up" id="kpi-ocf-chg">▲ 18.4% YoY</div>
    <div class="kpi-sub">Operating Cash Flow</div>
  </div>
</div>

<!-- TABS -->
<div class="tabs-bar">
  <button class="tab-btn active" onclick="switchTab('overview')">Overview</button>
  <button class="tab-btn" onclick="switchTab('cost')">Cost Analytics</button>
  <button class="tab-btn" onclick="switchTab('pl')">P&amp;L Trends</button>
  <button class="tab-btn" onclick="switchTab('quarterly')">Quarterly</button>
  <button class="tab-btn" onclick="switchTab('balance')">Balance Sheet</button>
  <button class="tab-btn" onclick="switchTab('cashflow')">Cash Flow</button>
  <button class="tab-btn" onclick="switchTab('costcvp')">Cost &amp; CVP</button>
</div>

<!-- MAIN CONTENT -->
<div class="main">

  <!-- OVERVIEW TAB -->
  <div class="tab-panel active" id="tab-overview">
    <div class="sec-hdr"><div class="sec-title">Financial Overview</div><div class="sec-note">FY2017 – FY2026 · ₹ Crores</div></div>
    <div class="grid-21 mb16">
      <div class="chart-card">
        <div class="chart-title">Revenue vs Net Profit</div>
        <div class="chart-subtitle">Annual trend — bars: Revenue, line: Net Profit</div>
        <div class="chart-wrap"><canvas id="ch-rev-np"></canvas></div>
      </div>
      <div class="chart-card">
        <div class="chart-title">Cost Structure FY26</div>
        <div class="chart-subtitle">Breakdown of total costs (% of revenue)</div>
        <div class="chart-wrap"><canvas id="ch-cost-donut"></canvas></div>
      </div>
    </div>
    <div class="grid-3">
      <div class="chart-card">
        <div class="chart-title">Profitability Margins</div>
        <div class="chart-subtitle">Gross / PBT / PAT margins %</div>
        <div class="chart-wrap"><canvas id="ch-margins"></canvas></div>
      </div>
      <div class="chart-card">
        <div class="chart-title">Cash Flow Health</div>
        <div class="chart-subtitle">Operating vs Investing vs Financing</div>
        <div class="chart-wrap"><canvas id="ch-ocf-overview"></canvas></div>
      </div>
      <div class="chart-card">
        <div class="chart-title">Balance Sheet Strength</div>
        <div class="chart-subtitle">Total Assets vs Borrowings (₹ Cr)</div>
        <div class="chart-wrap"><canvas id="ch-bs-overview"></canvas></div>
      </div>
    </div>
  </div>

  <!-- COST ANALYTICS TAB -->
  <div class="tab-panel" id="tab-cost">
    <div class="sec-hdr"><div class="sec-title">Cost Analytics Deep-Dive</div><div class="sec-note">Primary focus · FY2017–FY2026</div></div>

    <!-- Cost waterfall latest year -->
    <div class="grid-21 mb16">
      <div class="chart-card">
        <div class="chart-title">Cost Breakdown — FY2025 (Latest Complete)</div>
        <div class="chart-subtitle">Absolute cost by category (₹ Crores)</div>
        <div class="chart-wrap tall"><canvas id="ch-cost-bar"></canvas></div>
      </div>
      <div class="chart-card" id="wf-card">
        <div class="chart-title">Revenue Bridge FY2025</div>
        <div class="chart-subtitle">From Revenue → Net Profit</div>
        <div id="wf-container" style="margin-top: 8px;"></div>
      </div>
    </div>

    <div class="grid-2 mb16">
      <div class="chart-card">
        <div class="chart-title">Cost as % of Revenue — Historical Trend</div>
        <div class="chart-subtitle">Key cost lines as percentage of sales</div>
        <div class="chart-wrap tall"><canvas id="ch-cost-pct"></canvas></div>
      </div>
      <div class="chart-card">
        <div class="chart-title">Raw Material Cost Analysis</div>
        <div class="chart-subtitle">Absolute (₹ Cr, bars) vs % of Revenue (line)</div>
        <div class="chart-wrap tall"><canvas id="ch-rm-deep"></canvas></div>
      </div>
    </div>

    <div class="chart-card mb16">
      <div class="chart-title">Cost Composition — Stacked Area (₹ Crores)</div>
      <div class="chart-subtitle">How total cost is composed year-over-year</div>
      <div class="chart-wrap" style="height:300px"><canvas id="ch-cost-stack"></canvas></div>
    </div>

    <div class="chart-card">
      <div class="chart-title">Cost Efficiency Metrics</div>
      <div class="chart-subtitle">Key cost ratios across 10 years</div>
      <div style="overflow-x:auto; margin-top:4px;">
        <table class="cost-table" id="cost-eff-table">
          <thead><tr>
            <th>Metric</th>
            <th>FY17</th><th>FY18</th><th>FY19</th><th>FY20</th><th>FY21</th>
            <th>FY22</th><th>FY23</th><th>FY24</th><th>FY25</th><th>FY26</th>
          </tr></thead>
          <tbody id="cost-eff-body"></tbody>
        </table>
      </div>
    </div>
  </div>

  <!-- P&L TRENDS TAB -->
  <div class="tab-panel" id="tab-pl">
    <div class="sec-hdr"><div class="sec-title">Profit & Loss Trends</div><div class="sec-note">FY2017 – FY2026</div></div>
    <div class="chart-card mb16">
      <div class="chart-title">All P&L Line Items — Annual</div>
      <div class="chart-subtitle">Revenue, all cost components, and profit (₹ Crores)</div>
      <div class="chart-wrap" style="height:320px"><canvas id="ch-pl-all"></canvas></div>
    </div>
    <div class="grid-2">
      <div class="chart-card">
        <div class="chart-title">Employee Cost & Depreciation</div>
        <div class="chart-subtitle">Rising fixed cost base over time</div>
        <div class="chart-wrap"><canvas id="ch-emp-dep"></canvas></div>
      </div>
      <div class="chart-card">
        <div class="chart-title">Other Income vs Interest</div>
        <div class="chart-subtitle">Non-operating earnings power</div>
        <div class="chart-wrap"><canvas id="ch-oi-int"></canvas></div>
      </div>
    </div>
  </div>

  <!-- QUARTERLY TAB -->
  <div class="tab-panel" id="tab-quarterly">
    <div class="sec-hdr"><div class="sec-title">Quarterly Performance</div><div class="sec-note">Last 10 Quarters</div></div>
    <div class="chart-card mb16">
      <div class="chart-title">Quarterly Revenue & Operating Profit</div>
      <div class="chart-subtitle">Bars: Revenue · Line: Operating Profit (₹ Crores)</div>
      <div class="chart-wrap"><canvas id="ch-q-rev-op"></canvas></div>
    </div>
    <div class="grid-2">
      <div class="chart-card">
        <div class="chart-title">Quarterly Net Profit</div>
        <div class="chart-subtitle">PAT trend (₹ Crores)</div>
        <div class="chart-wrap"><canvas id="ch-q-np"></canvas></div>
      </div>
      <div class="chart-card">
        <div class="chart-title">Quarterly OPM & Cost Ratio</div>
        <div class="chart-subtitle">Operating margin % vs cost/revenue %</div>
        <div class="chart-wrap"><canvas id="ch-q-margins"></canvas></div>
      </div>
    </div>
  </div>

  <!-- BALANCE SHEET TAB -->
  <div class="tab-panel" id="tab-balance">
    <div class="sec-hdr"><div class="sec-title">Balance Sheet Analysis</div><div class="sec-note">FY2017 – FY2026</div></div>
    <div class="grid-2 mb16">
      <div class="chart-card">
        <div class="chart-title">Asset Composition</div>
        <div class="chart-subtitle">Net Block, Investments, Other Assets (₹ Cr)</div>
        <div class="chart-wrap tall"><canvas id="ch-assets"></canvas></div>
      </div>
      <div class="chart-card">
        <div class="chart-title">Reserves vs Borrowings</div>
        <div class="chart-subtitle">Equity base growth vs minimal debt</div>
        <div class="chart-wrap tall"><canvas id="ch-res-borrow"></canvas></div>
      </div>
    </div>
    <div class="grid-3">
      <div class="chart-card">
        <div class="chart-title">Working Capital</div>
        <div class="chart-subtitle">Receivables & Inventory (₹ Cr)</div>
        <div class="chart-wrap"><canvas id="ch-wc"></canvas></div>
      </div>
      <div class="chart-card">
        <div class="chart-title">Capital Allocation</div>
        <div class="chart-subtitle">CWIP & Net Block — Capex visibility</div>
        <div class="chart-wrap"><canvas id="ch-capex"></canvas></div>
      </div>
      <div class="chart-card">
        <div class="chart-title">Investments Corpus</div>
        <div class="chart-subtitle">Treasury strength (₹ Cr)</div>
        <div class="chart-wrap"><canvas id="ch-invest"></canvas></div>
      </div>
    </div>
  </div>

  <!-- CASH FLOW TAB -->
  <div class="tab-panel" id="tab-cashflow">
    <div class="sec-hdr"><div class="sec-title">Cash Flow Analysis</div><div class="sec-note">FY2017 – FY2026</div></div>
    <div class="chart-card mb16">
      <div class="chart-title">Three-Statement Cash Flow</div>
      <div class="chart-subtitle">Operating (green) · Investing (coral) · Financing (lavender) — ₹ Crores</div>
      <div class="chart-wrap" style="height: 300px"><canvas id="ch-cf-3"></canvas></div>
    </div>
    <div class="grid-2">
      <div class="chart-card">
        <div class="chart-title">OCF vs Net Profit</div>
        <div class="chart-subtitle">Cash conversion quality</div>
        <div class="chart-wrap"><canvas id="ch-cf-np"></canvas></div>
      </div>
      <div class="chart-card">
        <div class="chart-title">Free Cash Flow Proxy</div>
        <div class="chart-subtitle">OCF + Investing CF (indicative)</div>
        <div class="chart-wrap"><canvas id="ch-fcf"></canvas></div>
      </div>
    </div>
  </div>

  <!-- COST SHEET & CVP TAB -->
  <div class="tab-panel" id="tab-costcvp">
    <div class="sec-hdr"><div class="sec-title">Cost Sheet & CVP Analysis</div><div class="sec-note">FY2017 – FY2026 · ₹ Crores</div></div>
    <div class="grid-2 mb16">
      <div class="chart-card">
        <div class="chart-title">Cost Sheet — FY2025 (Latest Complete)</div>
        <div class="chart-subtitle">Derived from P&L line items</div>
        <div style="overflow-x:auto; margin-top:12px;">
          <table class="cost-table" id="cost-sheet-table">
            <thead><tr><th>Particulars</th><th>₹ Cr</th><th>% of Sales</th></tr></thead>
            <tbody id="cost-sheet-body"></tbody>
          </table>
        </div>
      </div>
      <div class="chart-card">
        <div class="chart-title">Cost-Volume-Profit (CVP) Trend</div>
        <div class="chart-subtitle">Break‑even sales vs actual sales</div>
        <div class="chart-wrap tall"><canvas id="ch-cvp"></canvas></div>
      </div>
    </div>
  </div>

</div><!-- /main -->

<!-- UPLOAD OVERLAY -->
<div class="upload-overlay" id="uploadOverlay" onclick="hideUpload(event)">
  <div class="upload-box" id="uploadBox">
    <div class="upload-icon">📊</div>
    <h3>Update Dashboard Data</h3>
    <p>Upload a Screener.in Excel export for Maruti Suzuki or any other company. The dashboard will parse and refresh automatically.</p>
    <label for="fileInput2">
      <div class="upload-btn" style="display:inline-flex; margin: 0 auto; justify-content:center; width: 100%;">
        📁 &nbsp; Choose Excel File (.xlsx)
      </div>
    </label>
    <input type="file" id="fileInput2" accept=".xlsx,.xls" onchange="handleFileUpload(this)">
    <p class="data-note" style="margin-top: 16px;">Supports Screener.in standard export format (Data Sheet structure)</p>
  </div>
</div>

<!-- TOAST -->
<div class="toast" id="toast"></div>

<script>
// ========================
// DATA MODEL
// ========================
let D = {
  company: 'MARUTI SUZUKI INDIA LTD',
  currentPrice: 13314,
  marketCap: 418595.59,
  years: ['FY17','FY18','FY19','FY20','FY21','FY22','FY23','FY24','FY25','FY26'],
  sales:        [68085, 79809.4, 86068.5, 75660, 70372, 88329.8, 117571.3, 141858.2, 152913, 183316],
  rawMat:       [47121.5, 54945.3, 59346.6, 53402, 50550.5, 66137.1, 86654.7, 100119.5, 108718.3, 134848],
  invChange:    [379.3, -40.8, -211.6, 238.7, -273.6, 93.1, 403.9, 378.6, 1227.5, 2272.5],
  powerFuel:    [518.6, 673.4, 863.3, 699.5, 476.6, 630.9, 790, 1033.4, 1052.6, null],
  otherMfr:     [382.9, 462.1, 475.7, 378.1, 683.3, 852.5, 1120.7, 939.3, 1994.4, null],
  empCost:      [2360.3, 2863.4, 3285, 3416.2, 3431.6, 4051.4, 4634.6, 6301.6, 7026, 9049.7],
  sellAdmin:    [6111.3, 6755.3, 8390.3, 8456.3, 7924, 9119.1, 11695, 12082.6, 11523.3, null],
  otherExp:     [1548.4, 1951.2, 2439.7, 2191.6, 1621.8, 1879.9, 2051, 3134.1, 3602.3, 20234.6],
  otherIncome:  [2399.2, 2154.6, 2664.2, 3410.4, 3046.3, 1860.8, 2306.6, 4247.6, 5198.8, 4642.7],
  depreciation: [2603.9, 2759.8, 3020.8, 3528.4, 3034.1, 2789, 2825.7, 5255.8, 5608.2, 6741.7],
  interest:     [89.4, 345.8, 75.9, 134.2, 101.8, 126.6, 187, 193.6, 194.2, 238.7],
  pbt:          [10127.2, 11166.9, 10623.8, 7102.8, 5321, 4697.2, 10323.1, 17424.5, 19620, 19118.5],
  tax:          [2616.2, 3286.2, 2973.2, 1425.2, 931.9, 817.7, 2112.1, 3936.3, 5119.8, 4439],
  netProfit:    [7509.9, 7880, 7649.1, 5676, 4389.1, 3879.5, 8211, 13488.2, 14500.2, 14679.5],
  dividend:     [2265, 2416, 2416, 1812, 1359, 1812, 2718, 3930, 4244.4, 4401.6],
  // balance sheet
  reserves:     [36924.1, 42408.4, 46941.1, 49262, 52349.6, 55182.5, 61640.3, 85478.8, 96082.7, 106999.1],
  borrowings:   [483.6, 120.8, 159.6, 184.1, 540.9, 425.5, 1247.3, 118.6, 87, 102.5],
  totalAssets:  [51960.5, 60248.4, 63968.7, 63627.7, 71376.1, 74655.5, 84596.9, 115304.1, 131971.2, 148881],
  netBlock:     [13310.7, 13388.8, 15437.3, 15744.4, 14988.7, 13747.2, 17830.4, 27864.8, 32982.7, 34955.4],
  cwip:         [1252.3, 2132.1, 1606.9, 1415.2, 1496.8, 2936.5, 2904.1, 7734.8, 7929, 9406.2],
  investments:  [29150.6, 36123.1, 37503.6, 37488, 42944.8, 42034.7, 49184.3, 57296, 66265.4, 76838.5],
  otherAssets:  [8246.9, 8604.4, 9420.9, 8980.1, 11945.8, 15937.1, 14678.1, 22408.5, 24794.1, 27680.9],
  receivables:  [1202.6, 1465.4, 2312.8, 1977.7, 1279.9, 2034.5, 3301.4, 4596.8, 6539.7, 5342],
  inventory:    [3263.7, 3160.2, 3322.6, 3213.9, 3049, 3532.3, 4283.5, 5318.1, 6913.2, 11320.6],
  cash:         [23.5, 74, 187.8, 29, 3047.1, 3042.2, 41.6, 2827.4, 552.9, 1580],
  // cash flow
  ocf:          [10282, 11787.9, 6600.9, 3495.8, 8856.2, 1840.5, 9251.4, 16801.1, 16136.2, 19099.9],
  icf:          [-9173.2, -8301.7, -3539.9, -556.6, -7291.3, -239.2, -8036.1, -11864.8, -14456.1, -14733.5],
  fcf:          [-1129.3, -3436.1, -2947.9, -3104.3, -1544.9, -1607, -1213.1, -4062, -4155.1, -4484],
  // quarterly
  qLabels:     ['Dec23','Mar24','Jun24','Sep24','Dec24','Mar25','Jun25','Sep25','Dec25','Mar26'],
  qSales:      [33512.8, 38471.2, 35779.4, 37449.2, 38764.3, 40920.1, 38605.2, 42344.2, 49904.1, 52462.5],
  qExpenses:   [29072.8, 33250.1, 30672.8, 32450.4, 33687.8, 36076.1, 33982.6, 37258.2, 44331, 46304.2],
  qOpProfit:   [4440, 5221.1, 5106.6, 4998.8, 5076.5, 4844, 4622.6, 5086, 5573.1, 6158.3],
  qNetProfit:  [3206.8, 3952.3, 3759.7, 3102.5, 3726.9, 3911.1, 3792.4, 3349, 3879.1, 3659],
  price:       [6015.7, 8861.1, 6672.55, 4288.3, 6859.2, 7561.3, 8292.15, 12600.35, 11522.15, 12306, 13314],
};

// ========================
// CHART INSTANCES
// ========================
const charts = {};

const C = {
  gold: '#c8a84b', gold2: '#e8c96a',
  teal: '#2dd4bf', coral: '#f87171',
  lavender: '#a78bfa', green: '#4ade80',
  orange: '#fb923c', blue: '#60a5fa',
  text: '#8fa3c8', bg: '#111929', border: '#1e2d4a',
};

Chart.defaults.color = C.text;
Chart.defaults.borderColor = C.border;
Chart.defaults.font.family = "'DM Sans', sans-serif";
Chart.defaults.font.size = 11;
Chart.defaults.plugins.legend.labels.boxWidth = 10;
Chart.defaults.plugins.legend.labels.padding = 14;
Chart.defaults.plugins.tooltip.backgroundColor = '#0e1525';
Chart.defaults.plugins.tooltip.borderColor = '#1e2d4a';
Chart.defaults.plugins.tooltip.borderWidth = 1;
Chart.defaults.plugins.tooltip.padding = 10;

function pct(a, b) { return b ? ((a / b) * 100).toFixed(1) : '—'; }
function fmt(n) { if (!n && n !== 0) return '—'; return n >= 1000 ? (n/1000).toFixed(1)+'K' : n.toFixed(0); }
function fmtC(n) { if (!n && n !== 0) return '—'; return '₹' + n.toLocaleString('en-IN', {maximumFractionDigits: 0}); }
function arr(a) { return a.map(v => v === null || v === undefined || isNaN(v) ? null : v); }

function mkChart(id, cfg) {
  if (charts[id]) charts[id].destroy();
  const ctx = document.getElementById(id);
  if (!ctx) return;
  charts[id] = new Chart(ctx, cfg);
  return charts[id];
}

// ========================
// TAB SWITCHING
// ========================
function switchTab(name) {
  document.querySelectorAll('.tab-btn').forEach(b => b.classList.remove('active'));
  document.querySelectorAll('.tab-panel').forEach(p => p.classList.remove('active'));
  event.target.classList.add('active');
  document.getElementById('tab-' + name).classList.add('active');
}

// ========================
// INIT ALL CHARTS
// ========================
function initAll() {
  updateKPIs();
  buildOverview();
  buildCostAnalytics();
  buildPL();
  buildQuarterly();
  buildBalance();
  buildCashFlow();
  buildCostCVP(); // new tab
}

function updateKPIs() {
  const n = D.years.length - 1;
  const rev = D.sales[n], revPrev = D.sales[n-1];
  const np = D.netProfit[n], npPrev = D.netProfit[n-1];
  const rm = D.rawMat[n], rmPrev = D.rawMat[n-1];
  const borrow = D.borrowings[n];
  const reserves = D.reserves[n];
  const ocf = D.ocf[n], ocfPrev = D.ocf[n-1];

  document.getElementById('hdr-price').textContent = '₹' + D.currentPrice.toLocaleString('en-IN');
  document.getElementById('hdr-mcap').textContent = '₹' + Math.round(D.marketCap).toLocaleString('en-IN') + ' Cr';

  document.getElementById('kpi-rev').innerHTML = '₹' + Math.round(rev/100)/10 + ' <span>K Cr</span>';
  const revChg = ((rev - revPrev) / revPrev * 100).toFixed(1);
  document.getElementById('kpi-rev-chg').textContent = (revChg > 0 ? '▲ ' : '▼ ') + Math.abs(revChg) + '% YoY';
  document.getElementById('kpi-rev-chg').className = 'kpi-chg ' + (revChg > 0 ? 'up' : 'dn');

  document.getElementById('kpi-np').innerHTML = '₹' + Math.round(np/100)/10 + ' <span>K Cr</span>';
  const npChg = ((np - npPrev) / npPrev * 100).toFixed(1);
  document.getElementById('kpi-np-chg').textContent = (npChg > 0 ? '▲ ' : '▼ ') + Math.abs(npChg) + '% YoY';
  document.getElementById('kpi-np-chg').className = 'kpi-chg ' + (npChg > 0 ? 'up' : 'dn');

  const rmPct = (rm / rev * 100).toFixed(1);
  const rmPctPrev = (rmPrev / revPrev * 100).toFixed(1);
  document.getElementById('kpi-rm').innerHTML = rmPct + ' <span>%</span>';
  const rmBps = Math.round((rmPct - rmPctPrev) * 100);
  document.getElementById('kpi-rm-chg').textContent = (rmBps > 0 ? '▲ +' : '▼ ') + Math.abs(rmBps) + ' bps vs prev yr';
  document.getElementById('kpi-rm-chg').className = 'kpi-chg ' + (rmBps > 0 ? 'dn' : 'up');

  // OPM from quarterly
  const lastQOPM = (D.qOpProfit[D.qOpProfit.length-1] / D.qSales[D.qSales.length-1] * 100).toFixed(1);
  document.getElementById('kpi-opm').innerHTML = lastQOPM + ' <span>%</span>';
  document.getElementById('kpi-opm-chg').textContent = 'Latest Qtr Q4 FY26';

  const de = (borrow / (reserves + 151)).toFixed(3);
  document.getElementById('kpi-de').innerHTML = de + ' <span>x</span>';
  document.getElementById('kpi-de-chg').textContent = 'Reserves: ₹' + Math.round(reserves/100)/10 + 'K Cr';

  document.getElementById('kpi-ocf').innerHTML = '₹' + Math.round(ocf/100)/10 + ' <span>K Cr</span>';
  const ocfChg = ((ocf - ocfPrev) / Math.abs(ocfPrev) * 100).toFixed(1);
  document.getElementById('kpi-ocf-chg').textContent = (ocfChg > 0 ? '▲ ' : '▼ ') + Math.abs(ocfChg) + '% YoY';
  document.getElementById('kpi-ocf-chg').className = 'kpi-chg ' + (ocfChg > 0 ? 'up' : 'dn');
}

function buildOverview() {
  // Revenue vs Net Profit
  mkChart('ch-rev-np', {
    type: 'bar',
    data: {
      labels: D.years,
      datasets: [
        { label: 'Revenue', data: D.sales, backgroundColor: 'rgba(200,168,75,0.3)', borderColor: C.gold, borderWidth: 1.5, borderRadius: 3 },
        { label: 'Net Profit', data: D.netProfit, type: 'line', borderColor: C.teal, backgroundColor: 'rgba(45,212,191,0.1)', borderWidth: 2, pointRadius: 4, pointBackgroundColor: C.teal, tension: 0.3, yAxisID: 'y2' }
      ]
    },
    options: { responsive: true, maintainAspectRatio: true, interaction: { mode: 'index' },
      scales: {
        y: { grid: { color: C.border }, ticks: { callback: v => '₹'+Math.round(v/1000)+'K' } },
        y2: { position: 'right', grid: { display: false }, ticks: { callback: v => '₹'+Math.round(v/1000)+'K' } }
      },
      plugins: { legend: { position: 'bottom' } }
    }
  });

  // Cost donut FY26
  const n = D.years.length - 1;
  const rev = D.sales[n];
  const costsDonut = [
    { label: 'Raw Material', val: D.rawMat[n], color: C.coral },
    { label: 'Employee Cost', val: D.empCost[n], color: C.lavender },
    { label: 'Depreciation', val: D.depreciation[n], color: C.blue },
    { label: 'Other Costs', val: D.otherExp[n], color: C.orange },
    { label: 'Net Profit', val: D.netProfit[n], color: C.green },
  ];
  mkChart('ch-cost-donut', {
    type: 'doughnut',
    data: {
      labels: costsDonut.map(c => c.label),
      datasets: [{ data: costsDonut.map(c => c.val), backgroundColor: costsDonut.map(c => c.color), borderWidth: 2, borderColor: '#111929', hoverOffset: 6 }]
    },
    options: { responsive: true, cutout: '60%', plugins: { legend: { position: 'bottom', labels: { generateLabels: (chart) => {
      return chart.data.labels.map((lbl, i) => ({
        text: lbl + ' (' + pct(costsDonut[i].val, rev) + '%)',
        fillStyle: costsDonut[i].color, strokeStyle: costsDonut[i].color, lineWidth: 0,
        fontColor: '#e8edf5'  // <-- changed font color for legend text
      }));
    }}}} }
  });

  // Profitability margins
  const opmArr = D.years.map((_, i) => {
    const expenses = (D.rawMat[i]||0) + (D.invChange[i]||0) + (D.powerFuel[i]||0) + (D.otherMfr[i]||0) + (D.empCost[i]||0) + (D.sellAdmin[i]||0) + (D.otherExp[i]||0);
    return +((1 - expenses / D.sales[i]) * 100).toFixed(1);
  });
  const pbtArr = D.years.map((_, i) => +(D.pbt[i] / D.sales[i] * 100).toFixed(1));
  const patArr = D.years.map((_, i) => +(D.netProfit[i] / D.sales[i] * 100).toFixed(1));

  mkChart('ch-margins', {
    type: 'line',
    data: {
      labels: D.years,
      datasets: [
        { label: 'PAT Margin', data: patArr, borderColor: C.teal, backgroundColor: 'rgba(45,212,191,0.08)', borderWidth: 2, fill: true, tension: 0.3, pointRadius: 3 },
        { label: 'PBT Margin', data: pbtArr, borderColor: C.lavender, borderWidth: 1.5, tension: 0.3, pointRadius: 2 },
      ]
    },
    options: { responsive: true, plugins: { legend: { position: 'bottom' } },
      scales: { y: { grid: { color: C.border }, ticks: { callback: v => v + '%' } } }
    }
  });

  // OCF overview
  mkChart('ch-ocf-overview', {
    type: 'bar',
    data: {
      labels: D.years,
      datasets: [
        { label: 'Operating CF', data: D.ocf, backgroundColor: 'rgba(74,222,128,0.4)', borderColor: C.green, borderWidth: 1.5, borderRadius: 3 },
        { label: 'Investing CF', data: D.icf, backgroundColor: 'rgba(248,113,113,0.3)', borderColor: C.coral, borderWidth: 1.5, borderRadius: 3 },
      ]
    },
    options: { responsive: true, plugins: { legend: { position: 'bottom' } },
      scales: { y: { grid: { color: C.border }, stacked: false } }
    }
  });

  // Balance sheet overview
  mkChart('ch-bs-overview', {
    type: 'line',
    data: {
      labels: D.years,
      datasets: [
        { label: 'Total Assets', data: D.totalAssets, borderColor: C.gold, backgroundColor: 'rgba(200,168,75,0.08)', fill: true, borderWidth: 2, tension: 0.3, pointRadius: 3 },
        { label: 'Borrowings', data: D.borrowings, borderColor: C.coral, borderWidth: 1.5, tension: 0.3, pointRadius: 3 },
      ]
    },
    options: { responsive: true, plugins: { legend: { position: 'bottom' } },
      scales: { y: { grid: { color: C.border }, ticks: { callback: v => '₹'+Math.round(v/1000)+'K' } } }
    }
  });
}

function buildCostAnalytics() {
  // Cost bar FY25 (index 8, last complete with all line items)
  const idx = 8; // FY25
  const rev = D.sales[idx];
  const costItems = [
    { label: 'Raw Material', val: D.rawMat[idx], color: C.coral },
    { label: 'Inventory Δ', val: D.invChange[idx], color: C.orange },
    { label: 'Power & Fuel', val: D.powerFuel[idx] || 0, color: C.blue },
    { label: 'Other Mfr', val: D.otherMfr[idx] || 0, color: C.lavender },
    { label: 'Employee Cost', val: D.empCost[idx], color: '#34d399' },
    { label: 'Selling & Admin', val: D.sellAdmin[idx] || 0, color: C.gold },
    { label: 'Other Expenses', val: D.otherExp[idx], color: C.teal },
    { label: 'Depreciation', val: D.depreciation[idx], color: '#94a3b8' },
    { label: 'Interest', val: D.interest[idx], color: '#6366f1' },
  ].filter(c => c.val > 0);

  mkChart('ch-cost-bar', {
    type: 'bar',
    data: {
      labels: costItems.map(c => c.label),
      datasets: [{ data: costItems.map(c => c.val), backgroundColor: costItems.map(c => c.color + 'aa'), borderColor: costItems.map(c => c.color), borderWidth: 1.5, borderRadius: 4 }]
    },
    options: {
      indexAxis: 'y', responsive: true,
      plugins: { legend: { display: false },
        tooltip: { callbacks: { label: ctx => ' ₹' + ctx.parsed.x.toFixed(0) + ' Cr (' + (ctx.parsed.x / rev * 100).toFixed(1) + '% of rev)' } }
      },
      scales: { x: { grid: { color: C.border }, ticks: { callback: v => '₹'+Math.round(v/1000)+'K' } }, y: { grid: { display: false } } }
    }
  });

  // Revenue Bridge / Waterfall (FY25)
  buildWaterfall(idx);

  // Cost as % of revenue trend
  const rmPct = D.sales.map((s, i) => +(D.rawMat[i] / s * 100).toFixed(1));
  const empPct = D.sales.map((s, i) => +(D.empCost[i] / s * 100).toFixed(1));
  const saePct = D.sales.map((s, i) => +((D.sellAdmin[i] || D.otherExp[i]) / s * 100).toFixed(1));
  const depPct = D.sales.map((s, i) => +(D.depreciation[i] / s * 100).toFixed(1));

  mkChart('ch-cost-pct', {
    type: 'line',
    data: {
      labels: D.years,
      datasets: [
        { label: 'Raw Material %', data: rmPct, borderColor: C.coral, backgroundColor: 'rgba(248,113,113,0.08)', fill: false, borderWidth: 2, tension: 0.3, pointRadius: 3 },
        { label: 'Employee Cost %', data: empPct, borderColor: C.lavender, borderWidth: 2, tension: 0.3, pointRadius: 3 },
        { label: 'Sell+Admin / Other %', data: saePct, borderColor: C.gold, borderWidth: 2, tension: 0.3, pointRadius: 3 },
        { label: 'Depreciation %', data: depPct, borderColor: '#94a3b8', borderWidth: 2, tension: 0.3, pointRadius: 3 },
      ]
    },
    options: { responsive: true, plugins: { legend: { position: 'bottom' } },
      scales: { y: { grid: { color: C.border }, ticks: { callback: v => v + '%' } } }
    }
  });

  // Raw material deep dive
  mkChart('ch-rm-deep', {
    type: 'bar',
    data: {
      labels: D.years,
      datasets: [
        { label: 'Raw Material Cost', data: D.rawMat, backgroundColor: 'rgba(248,113,113,0.3)', borderColor: C.coral, borderWidth: 1.5, borderRadius: 3, order: 2 },
        { label: '% of Revenue', data: rmPct, type: 'line', borderColor: C.gold, borderWidth: 2, pointRadius: 4, pointBackgroundColor: C.gold, tension: 0.3, yAxisID: 'y2', order: 1 }
      ]
    },
    options: { responsive: true, plugins: { legend: { position: 'bottom' } }, interaction: { mode: 'index' },
      scales: {
        y: { grid: { color: C.border }, ticks: { callback: v => '₹'+Math.round(v/1000)+'K' } },
        y2: { position: 'right', grid: { display: false }, ticks: { callback: v => v + '%' }, min: 65, max: 80 }
      }
    }
  });

  // Stacked area cost composition
  mkChart('ch-cost-stack', {
    type: 'line',
    data: {
      labels: D.years,
      datasets: [
        { label: 'Interest', data: D.interest, borderColor: '#6366f1', backgroundColor: 'rgba(99,102,241,0.5)', fill: true, borderWidth: 1, tension: 0.3, pointRadius: 0 },
        { label: 'Depreciation', data: D.depreciation, borderColor: '#94a3b8', backgroundColor: 'rgba(148,163,184,0.5)', fill: true, borderWidth: 1, tension: 0.3, pointRadius: 0 },
        { label: 'Employee Cost', data: D.empCost, borderColor: '#34d399', backgroundColor: 'rgba(52,211,153,0.5)', fill: true, borderWidth: 1, tension: 0.3, pointRadius: 0 },
        { label: 'Raw Material', data: D.rawMat, borderColor: C.coral, backgroundColor: 'rgba(248,113,113,0.5)', fill: true, borderWidth: 1, tension: 0.3, pointRadius: 0 },
      ]
    },
    options: { responsive: true, maintainAspectRatio: false, plugins: { legend: { position: 'bottom' } },
      scales: { y: { stacked: true, grid: { color: C.border }, ticks: { callback: v => '₹'+Math.round(v/1000)+'K' } }, x: { grid: { display: false } } }
    }
  });

  // Cost efficiency table
  const rows = [
    { label: 'Raw Material %', fn: i => (D.rawMat[i] / D.sales[i] * 100).toFixed(1) + '%', thresh: (v, prev) => parseFloat(v) > parseFloat(prev) ? 'dn' : 'up' },
    { label: 'Employee Cost %', fn: i => (D.empCost[i] / D.sales[i] * 100).toFixed(1) + '%', thresh: (v, prev) => 'nt' },
    { label: 'PAT Margin %', fn: i => (D.netProfit[i] / D.sales[i] * 100).toFixed(1) + '%', thresh: (v, prev) => parseFloat(v) > parseFloat(prev) ? 'up' : 'dn' },
    { label: 'Depreciation %', fn: i => (D.depreciation[i] / D.sales[i] * 100).toFixed(1) + '%', thresh: () => 'nt' },
    { label: 'Tax Rate %', fn: i => (D.tax[i] / D.pbt[i] * 100).toFixed(1) + '%', thresh: () => 'nt' },
    { label: 'OCF/Rev %', fn: i => D.ocf[i] ? (D.ocf[i] / D.sales[i] * 100).toFixed(1) + '%' : '—', thresh: (v, prev) => parseFloat(v) > parseFloat(prev) ? 'up' : 'dn' },
  ];
  const tbody = document.getElementById('cost-eff-body');
  if (tbody) {
    tbody.innerHTML = rows.map(r => `<tr>
      <td>${r.label}</td>
      ${D.years.map((_, i) => {
        const v = r.fn(i);
        const cls = i > 0 ? r.thresh(v, r.fn(i - 1)) : 'nt';
        const arrow = cls === 'up' ? ' ▲' : cls === 'dn' ? ' ▼' : '';
        return `<td style="color: ${cls==='up'?'var(--green)':cls==='dn'?'var(--red)':'inherit'}">${v}${arrow}</td>`;
      }).join('')}
    </tr>`).join('');
  }
}

function buildWaterfall(idx) {
  const rev = D.sales[idx];
  const rmTotal = (D.rawMat[idx] || 0) + (D.invChange[idx] || 0);
  const mfgCost = (D.powerFuel[idx] || 0) + (D.otherMfr[idx] || 0);
  const emp = D.empCost[idx] || 0;
  const sa = D.sellAdmin[idx] || 0;
  const oe = D.otherExp[idx] || 0;
  const dep = D.depreciation[idx] || 0;
  const int_ = D.interest[idx] || 0;
  const oi = D.otherIncome[idx] || 0;
  const np = D.netProfit[idx] || 0;

  const items = [
    { label: 'Revenue', val: rev, color: C.teal, plus: true },
    { label: 'Raw Material (net)', val: -rmTotal, color: C.coral },
    { label: 'Mfg Expenses', val: -mfgCost, color: '#fb923c' },
    { label: 'Employee Cost', val: -emp, color: C.lavender },
    { label: 'Selling & Admin', val: -sa, color: C.gold },
    { label: 'Other Expenses', val: -oe, color: C.blue },
    { label: 'Other Income', val: oi, color: C.green },
    { label: 'Depreciation', val: -dep, color: '#94a3b8' },
    { label: 'Interest', val: -int_, color: '#6366f1' },
    { label: 'Tax', val: -(D.tax[idx]||0), color: '#f59e0b' },
    { label: 'Net Profit', val: np, color: C.teal, plus: true },
  ];

  const maxAbsVal = Math.max(...items.map(i => Math.abs(i.val)));

  const container = document.getElementById('wf-container');
  if (!container) return;
  container.innerHTML = `
    <div class="wf-header"><span class="l">Item</span><span class="r">Scale</span><span class="v">₹ Cr</span><span class="p">% Rev</span></div>
    ${items.map(item => {
      const barW = Math.round(Math.abs(item.val) / maxAbsVal * 130);
      const isPlus = item.val > 0;
      const pctRev = (Math.abs(item.val) / rev * 100).toFixed(1);
      return `<div class="wf-row">
        <div class="wf-label">${item.label}</div>
        <div class="wf-bar-area">
          <div class="wf-bar" style="width:${barW}px; background: ${item.color}; opacity: ${isPlus ? '0.9' : '0.6'};"></div>
        </div>
        <div class="wf-amount" style="color: ${item.color}">${isPlus ? '+' : ''}${item.val.toFixed(0)}</div>
        <div class="wf-pct">${pctRev}%</div>
      </div>`;
    }).join('')}
  `;
}

function buildPL() {
  mkChart('ch-pl-all', {
    type: 'line',
    data: {
      labels: D.years,
      datasets: [
        { label: 'Revenue', data: D.sales, borderColor: C.gold, borderWidth: 2.5, tension: 0.3, pointRadius: 4, pointBackgroundColor: C.gold },
        { label: 'Raw Material', data: D.rawMat, borderColor: C.coral, borderWidth: 1.5, tension: 0.3, pointRadius: 2, borderDash: [] },
        { label: 'Total Expenses', data: D.sales.map((s,i) => s - D.netProfit[i]), borderColor: '#fb923c', borderWidth: 1.5, tension: 0.3, pointRadius: 2 },
        { label: 'Net Profit', data: D.netProfit, borderColor: C.teal, borderWidth: 2, tension: 0.3, pointRadius: 3 },
      ]
    },
    options: { responsive: true, maintainAspectRatio: false, plugins: { legend: { position: 'bottom' } }, interaction: { mode: 'index' },
      scales: { y: { grid: { color: C.border }, ticks: { callback: v => '₹'+Math.round(v/1000)+'K' } } }
    }
  });

  mkChart('ch-emp-dep', {
    type: 'bar',
    data: {
      labels: D.years,
      datasets: [
        { label: 'Employee Cost', data: D.empCost, backgroundColor: 'rgba(167,139,250,0.4)', borderColor: C.lavender, borderWidth: 1.5, borderRadius: 3 },
        { label: 'Depreciation', data: D.depreciation, backgroundColor: 'rgba(148,163,184,0.4)', borderColor: '#94a3b8', borderWidth: 1.5, borderRadius: 3 },
      ]
    },
    options: { responsive: true, plugins: { legend: { position: 'bottom' } },
      scales: { y: { grid: { color: C.border }, ticks: { callback: v => '₹'+Math.round(v/1000)+'K' } } }
    }
  });

  mkChart('ch-oi-int', {
    type: 'bar',
    data: {
      labels: D.years,
      datasets: [
        { label: 'Other Income', data: D.otherIncome, backgroundColor: 'rgba(74,222,128,0.35)', borderColor: C.green, borderWidth: 1.5, borderRadius: 3 },
        { label: 'Interest', data: D.interest, backgroundColor: 'rgba(99,102,241,0.35)', borderColor: '#6366f1', borderWidth: 1.5, borderRadius: 3 },
      ]
    },
    options: { responsive: true, plugins: { legend: { position: 'bottom' } },
      scales: { y: { grid: { color: C.border } } }
    }
  });
}

function buildQuarterly() {
  mkChart('ch-q-rev-op', {
    type: 'bar',
    data: {
      labels: D.qLabels,
      datasets: [
        { label: 'Revenue', data: D.qSales, backgroundColor: 'rgba(200,168,75,0.3)', borderColor: C.gold, borderWidth: 1.5, borderRadius: 3, order: 2 },
        { label: 'Operating Profit', data: D.qOpProfit, type: 'line', borderColor: C.teal, borderWidth: 2.5, pointRadius: 5, pointBackgroundColor: C.teal, tension: 0.3, order: 1, yAxisID: 'y2' }
      ]
    },
    options: { responsive: true, plugins: { legend: { position: 'bottom' } }, interaction: { mode: 'index' },
      scales: {
        y: { grid: { color: C.border }, ticks: { callback: v => '₹'+Math.round(v/1000)+'K' } },
        y2: { position: 'right', grid: { display: false }, ticks: { callback: v => '₹'+Math.round(v/1000)+'K' } }
      }
    }
  });

  mkChart('ch-q-np', {
    type: 'line',
    data: {
      labels: D.qLabels,
      datasets: [{ label: 'Net Profit', data: D.qNetProfit, borderColor: C.teal, backgroundColor: 'rgba(45,212,191,0.1)', fill: true, borderWidth: 2.5, pointRadius: 5, pointBackgroundColor: C.teal, tension: 0.3 }]
    },
    options: { responsive: true, plugins: { legend: { position: 'bottom' } },
      scales: { y: { grid: { color: C.border } } }
    }
  });

  const qOPM = D.qSales.map((s, i) => +(D.qOpProfit[i] / s * 100).toFixed(1));
  const qCostRatio = D.qSales.map((s, i) => +(D.qExpenses[i] / s * 100).toFixed(1));
  mkChart('ch-q-margins', {
    type: 'line',
    data: {
      labels: D.qLabels,
      datasets: [
        { label: 'OPM %', data: qOPM, borderColor: C.gold, borderWidth: 2, pointRadius: 4, pointBackgroundColor: C.gold, tension: 0.3 },
        { label: 'Cost Ratio %', data: qCostRatio, borderColor: C.coral, borderWidth: 2, pointRadius: 4, pointBackgroundColor: C.coral, tension: 0.3 },
      ]
    },
    options: { responsive: true, plugins: { legend: { position: 'bottom' } },
      scales: { y: { grid: { color: C.border }, ticks: { callback: v => v + '%' } } }
    }
  });
}

function buildBalance() {
  mkChart('ch-assets', {
    type: 'bar',
    data: {
      labels: D.years,
      datasets: [
        { label: 'Net Block', data: D.netBlock, backgroundColor: 'rgba(96,165,250,0.5)', borderColor: C.blue, borderWidth: 1.5, borderRadius: 2 },
        { label: 'CWIP', data: D.cwip, backgroundColor: 'rgba(251,146,60,0.5)', borderColor: C.orange, borderWidth: 1.5, borderRadius: 2 },
        { label: 'Investments', data: D.investments, backgroundColor: 'rgba(200,168,75,0.4)', borderColor: C.gold, borderWidth: 1.5, borderRadius: 2 },
        { label: 'Other Assets', data: D.otherAssets, backgroundColor: 'rgba(167,139,250,0.4)', borderColor: C.lavender, borderWidth: 1.5, borderRadius: 2 },
      ]
    },
    options: { responsive: true, plugins: { legend: { position: 'bottom' } },
      scales: { x: { stacked: true }, y: { stacked: true, grid: { color: C.border }, ticks: { callback: v => '₹'+Math.round(v/1000)+'K' } } }
    }
  });

  mkChart('ch-res-borrow', {
    type: 'line',
    data: {
      labels: D.years,
      datasets: [
        { label: 'Reserves', data: D.reserves, borderColor: C.green, backgroundColor: 'rgba(74,222,128,0.08)', fill: true, borderWidth: 2.5, tension: 0.3, pointRadius: 3 },
        { label: 'Borrowings', data: D.borrowings, borderColor: C.coral, borderWidth: 2, tension: 0.3, pointRadius: 4, pointBackgroundColor: C.coral },
      ]
    },
    options: { responsive: true, plugins: { legend: { position: 'bottom' } },
      scales: { y: { grid: { color: C.border }, ticks: { callback: v => '₹'+Math.round(v/1000)+'K' } } }
    }
  });

  mkChart('ch-wc', {
    type: 'bar',
    data: {
      labels: D.years,
      datasets: [
        { label: 'Receivables', data: D.receivables, backgroundColor: 'rgba(45,212,191,0.4)', borderColor: C.teal, borderWidth: 1.5, borderRadius: 3 },
        { label: 'Inventory', data: D.inventory, backgroundColor: 'rgba(200,168,75,0.4)', borderColor: C.gold, borderWidth: 1.5, borderRadius: 3 },
        { label: 'Cash & Bank', data: D.cash, backgroundColor: 'rgba(74,222,128,0.4)', borderColor: C.green, borderWidth: 1.5, borderRadius: 3 },
      ]
    },
    options: { responsive: true, plugins: { legend: { position: 'bottom' } },
      scales: { y: { grid: { color: C.border } } }
    }
  });

  mkChart('ch-capex', {
    type: 'bar',
    data: {
      labels: D.years,
      datasets: [
        { label: 'Net Block', data: D.netBlock, backgroundColor: 'rgba(96,165,250,0.4)', borderColor: C.blue, borderWidth: 1.5, borderRadius: 3 },
        { label: 'CWIP', data: D.cwip, backgroundColor: 'rgba(251,146,60,0.4)', borderColor: C.orange, borderWidth: 1.5, borderRadius: 3 },
      ]
    },
    options: { responsive: true, plugins: { legend: { position: 'bottom' } },
      scales: { y: { grid: { color: C.border } } }
    }
  });

  mkChart('ch-invest', {
    type: 'line',
    data: {
      labels: D.years,
      datasets: [{ label: 'Investments', data: D.investments, borderColor: C.gold, backgroundColor: 'rgba(200,168,75,0.1)', fill: true, borderWidth: 2.5, tension: 0.4, pointRadius: 4, pointBackgroundColor: C.gold }]
    },
    options: { responsive: true, plugins: { legend: { position: 'bottom' } },
      scales: { y: { grid: { color: C.border }, ticks: { callback: v => '₹'+Math.round(v/1000)+'K' } } }
    }
  });
}

function buildCashFlow() {
  mkChart('ch-cf-3', {
    type: 'bar',
    data: {
      labels: D.years,
      datasets: [
        { label: 'Operating CF', data: D.ocf, backgroundColor: 'rgba(74,222,128,0.4)', borderColor: C.green, borderWidth: 1.5, borderRadius: 3 },
        { label: 'Investing CF', data: D.icf, backgroundColor: 'rgba(248,113,113,0.35)', borderColor: C.coral, borderWidth: 1.5, borderRadius: 3 },
        { label: 'Financing CF', data: D.fcf, backgroundColor: 'rgba(167,139,250,0.35)', borderColor: C.lavender, borderWidth: 1.5, borderRadius: 3 },
      ]
    },
    options: { responsive: true, maintainAspectRatio: false, plugins: { legend: { position: 'bottom' } }, interaction: { mode: 'index' },
      scales: { y: { grid: { color: C.border }, ticks: { callback: v => '₹'+Math.round(v/1000)+'K' } } }
    }
  });

  mkChart('ch-cf-np', {
    type: 'line',
    data: {
      labels: D.years,
      datasets: [
        { label: 'Operating CF', data: D.ocf, borderColor: C.green, borderWidth: 2.5, tension: 0.3, pointRadius: 4, pointBackgroundColor: C.green },
        { label: 'Net Profit', data: D.netProfit, borderColor: C.gold, borderWidth: 2, borderDash: [4, 3], tension: 0.3, pointRadius: 3 },
      ]
    },
    options: { responsive: true, plugins: { legend: { position: 'bottom' } },
      scales: { y: { grid: { color: C.border }, ticks: { callback: v => '₹'+Math.round(v/1000)+'K' } } }
    }
  });

  const fcf2 = D.ocf.map((o, i) => +(o + D.icf[i]).toFixed(0));
  mkChart('ch-fcf', {
    type: 'bar',
    data: {
      labels: D.years,
      datasets: [{ label: 'FCF Proxy (OCF + ICF)', data: fcf2, backgroundColor: fcf2.map(v => v >= 0 ? 'rgba(74,222,128,0.35)' : 'rgba(248,113,113,0.35)'), borderColor: fcf2.map(v => v >= 0 ? C.green : C.coral), borderWidth: 1.5, borderRadius: 3 }]
    },
    options: { responsive: true, plugins: { legend: { position: 'bottom' } },
      scales: { y: { grid: { color: C.border }, ticks: { callback: v => '₹'+Math.round(v/1000)+'K' } } }
    }
  });
}

// ========================
// COST SHEET & CVP TAB
// ========================
function buildCostCVP() {
  // Use FY25 (index 8) for the cost sheet because it's the last year with all components.
  const i = 8;
  const sales = D.sales[i];
  const rawMatNet = (D.rawMat[i] || 0) + (D.invChange[i] || 0);
  const mfgExp = (D.powerFuel[i] || 0) + (D.otherMfr[i] || 0);
  const emp = D.empCost[i] || 0;
  const sellAdmin = D.sellAdmin[i] || 0;
  const otherExp = D.otherExp[i] || 0;
  const dep = D.depreciation[i] || 0;
  const interest = D.interest[i] || 0;

  // Cost sheet rows (simplified)
  const primeCost = rawMatNet + mfgExp + emp; // assuming direct labour
  const factoryOverheads = dep; // as proxy
  const worksCost = primeCost + factoryOverheads;
  const officeAdmin = sellAdmin + otherExp;
  const costOfProduction = worksCost + officeAdmin; // COGS approx
  const totalCost = costOfProduction + interest;

  const rows = [
    { label: 'Raw Material (net)', val: rawMatNet, cls: '' },
    { label: 'Manufacturing Expenses', val: mfgExp, cls: '' },
    { label: 'Employee Cost', val: emp, cls: '' },
    { label: 'Prime Cost', val: primeCost, cls: 'strong' },
    { label: 'Depreciation (Factory Overheads)', val: dep, cls: '' },
    { label: 'Works Cost', val: worksCost, cls: 'strong' },
    { label: 'Selling & Admin + Other Exp.', val: officeAdmin, cls: '' },
    { label: 'Cost of Production', val: costOfProduction, cls: 'strong' },
    { label: 'Interest', val: interest, cls: '' },
    { label: 'Total Cost', val: totalCost, cls: 'strong' },
    { label: 'Net Profit', val: D.netProfit[i], cls: '' }
  ];

  const tbody = document.getElementById('cost-sheet-body');
  if (tbody) {
    tbody.innerHTML = rows.map(r => `<tr>
      <td style="font-weight:${r.cls==='strong'?'600':'400'}">${r.label}</td>
      <td>₹ ${Math.round(r.val).toLocaleString('en-IN')}</td>
      <td>${((r.val / sales) * 100).toFixed(1)}%</td>
    </tr>`).join('');
  }

  // CVP: classify costs as variable / fixed across all years
  const variableLines = D.sales.map((s, idx) => {
    const rm = (D.rawMat[idx] || 0) + (D.invChange[idx] || 0);
    const mfg = (D.powerFuel[idx] || 0) + (D.otherMfr[idx] || 0);
    return rm + mfg;
  });
  const fixedLines = D.sales.map((s, idx) => {
    const emp = D.empCost[idx] || 0;
    const sa = D.sellAdmin[idx] || 0;
    const oe = D.otherExp[idx] || 0;
    const dep = D.depreciation[idx] || 0;
    const int = D.interest[idx] || 0;
    return emp + sa + oe + dep + int;
  });

  const contribution = D.sales.map((s, idx) => s - variableLines[idx]);
  const cmRatio = contribution.map((c, idx) => c / D.sales[idx]);
  const breakEven = fixedLines.map((f, idx) => cmRatio[idx] ? f / cmRatio[idx] : 0);

  mkChart('ch-cvp', {
    type: 'line',
    data: {
      labels: D.years,
      datasets: [
        { label: 'Actual Sales', data: D.sales, borderColor: C.gold, borderWidth: 2, tension: 0.3, pointRadius: 3 },
        { label: 'Break‑Even Sales', data: breakEven, borderColor: C.coral, borderWidth: 2, borderDash: [5, 4], tension: 0.3, pointRadius: 3 },
        { label: 'Contribution (Sales – VC)', data: contribution, borderColor: C.teal, borderWidth: 2, tension: 0.3, pointRadius: 3, borderDash: [2, 2] },
      ]
    },
    options: { responsive: true, plugins: { legend: { position: 'bottom' } },
      scales: { y: { grid: { color: C.border }, ticks: { callback: v => '₹'+Math.round(v/1000)+'K' } } }
    }
  });
}

// ========================
// FILE UPLOAD
// ========================
function showUpload() { document.getElementById('uploadOverlay').classList.add('show'); }
function hideUpload(e) { if (e.target.id === 'uploadOverlay') document.getElementById('uploadOverlay').classList.remove('show'); }

function handleFileUpload(input) {
  const file = input.files[0];
  if (!file) return;
  const reader = new FileReader();
  reader.onload = function(e) {
    try {
      const wb = XLSX.read(e.target.result, { type: 'array', cellDates: true });
      parseScreenerWorkbook(wb);
      document.getElementById('uploadOverlay').classList.remove('show');
      showToast('✓ Dashboard updated from ' + file.name);
    } catch(err) {
      showToast('⚠ Error parsing file: ' + err.message);
      console.error(err);
    }
  };
  reader.readAsArrayBuffer(file);
}

function parseScreenerWorkbook(wb) {
  const dsName = wb.SheetNames.find(n => n.toLowerCase().includes('data'));
  if (!dsName) { showToast('Could not find Data Sheet'); return; }
  const ws = wb.Sheets[dsName];
  const raw = XLSX.utils.sheet_to_json(ws, { header: 1, defval: null });

  function findRow(label) {
    return raw.find(r => r[0] && r[0].toString().trim().toLowerCase() === label.toLowerCase().trim());
  }
  function numArr(row) {
    if (!row) return [];
    return row.slice(1).filter(v => v !== null && v !== undefined && v !== '' && !isNaN(parseFloat(v))).map(v => parseFloat(v));
  }
  function extractDates(row) {
    if (!row) return [];
    return row.slice(1).filter(v => v !== null && v !== undefined && v !== '').map(v => {
      const d = new Date(v);
      if (!isNaN(d.getTime())) return 'FY' + String(d.getFullYear()).slice(2);
      return v;
    });
  }

  const metaRow = raw.find(r => r[0] && r[0].toString().includes('COMPANY NAME'));
  if (metaRow && metaRow[1]) D.company = metaRow[1].toString().trim();

  const priceRow = raw.find(r => r[0] && r[0].toString().trim() === 'Current Price');
  if (priceRow && priceRow[1]) D.currentPrice = parseFloat(priceRow[1]);
  const mcapRow = raw.find(r => r[0] && r[0].toString().toLowerCase().includes('market cap'));
  if (mcapRow && mcapRow[1]) D.marketCap = parseFloat(mcapRow[1]);

  // P&L section
  const plDateRow = raw.find((r, i) => i > 10 && r[0] && r[0].toString().includes('PROFIT'));
  let plDateRowIdx = plDateRow ? raw.indexOf(plDateRow) : -1;
  let dateRowPL = null;
  for (let i = plDateRowIdx; i < raw.length && i < plDateRowIdx + 5; i++) {
    if (raw[i] && raw[i][0] && raw[i][0].toString().includes('Report Date')) { dateRowPL = raw[i]; break; }
  }
  if (dateRowPL) D.years = extractDates(dateRowPL);

  const pick = (lbl) => { const r = findRow(lbl); return r ? numArr(r) : []; };
  D.sales = pick('Sales') || D.sales;
  D.rawMat = pick('Raw Material Cost') || D.rawMat;
  D.invChange = pick('Change in Inventory') || D.invChange;
  D.powerFuel = pick('Power and Fuel') || D.powerFuel;
  D.otherMfr = pick('Other Mfr. Exp') || D.otherMfr;
  D.empCost = pick('Employee Cost') || D.empCost;
  D.sellAdmin = pick('Selling and admin') || D.sellAdmin;
  D.otherExp = pick('Other Expenses') || D.otherExp;
  D.otherIncome = pick('Other Income') || D.otherIncome;
  D.depreciation = pick('Depreciation') || D.depreciation;
  D.interest = pick('Interest') || D.interest;
  D.pbt = pick('Profit before tax') || D.pbt;
  D.tax = pick('Tax') || D.tax;
  D.netProfit = pick('Net profit') || D.netProfit;
  D.dividend = pick('Dividend Amount') || D.dividend;

  // Quarterly
  const qDateRow = raw.find((r, i) => {
    if (r[0] && r[0].toString().includes('Quarters')) {
      for (let j = i; j < Math.min(i+5, raw.length); j++) {
        if (raw[j] && raw[j][0] && raw[j][0].toString().includes('Report Date')) return true;
      }
    }
    return false;
  });
  const qSection = raw.findIndex(r => r[0] && r[0].toString().includes('Quarters'));
  if (qSection >= 0) {
    for (let i = qSection; i < Math.min(qSection+5, raw.length); i++) {
      if (raw[i] && raw[i][0] && raw[i][0].toString().includes('Report Date')) {
        D.qLabels = raw[i].slice(1).filter(v => v !== null && v !== undefined && v !== '').map(v => {
          if (!v) return '';
          const d = new Date(v);
          if (!isNaN(d.getTime())) return ['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'][d.getMonth()] + String(d.getFullYear()).slice(2);
          return v;
        });
        break;
      }
    }
    const afterQ = raw.slice(qSection+1);
    const qSalesRow = afterQ.find(r => r[0] && r[0].toString().trim() === 'Sales');
    const qExpRow = afterQ.find(r => r[0] && r[0].toString().trim() === 'Expenses');
    const qOPRow = afterQ.find(r => r[0] && r[0].toString().trim() === 'Operating Profit');
    const qNPRow = afterQ.find(r => r[0] && r[0].toString().trim() === 'Net profit');
    if (qSalesRow) D.qSales = numArr(qSalesRow);
    if (qExpRow) D.qExpenses = numArr(qExpRow);
    if (qOPRow) D.qOpProfit = numArr(qOPRow);
    if (qNPRow) D.qNetProfit = numArr(qNPRow);
  }

  // Balance Sheet
  D.reserves = pick('Reserves') || D.reserves;
  D.borrowings = pick('Borrowings') || D.borrowings;
  D.netBlock = pick('Net Block') || D.netBlock;
  D.cwip = pick('Capital Work in Progress') || D.cwip;
  D.investments = pick('Investments') || D.investments;
  D.otherAssets = pick('Other Assets') || D.otherAssets;
  D.receivables = pick('Receivables') || D.receivables;
  D.inventory = pick('Inventory') || D.inventory;
  D.cash = pick('Cash & Bank') || D.cash;

  const totalRow = raw.filter(r => r[0] && r[0].toString().trim() === 'Total');
  if (totalRow.length > 0) D.totalAssets = numArr(totalRow[0]);

  // Cash flow
  D.ocf = pick('Cash from Operating Activity') || D.ocf;
  D.icf = pick('Cash from Investing Activity') || D.icf;
  D.fcf = pick('Cash from Financing Activity') || D.fcf;

  document.querySelector('.brand-title').textContent = D.company;

  initAll();
}

function showToast(msg) {
  const t = document.getElementById('toast');
  t.textContent = msg;
  t.classList.add('show');
  setTimeout(() => t.classList.remove('show'), 3500);
}

// ========================
// INIT
// ========================
document.addEventListener('DOMContentLoaded', initAll);
</script>
</body>
</html>
