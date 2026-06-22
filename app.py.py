# ============================================================
#  MIS AutoReport Pro — AI-Powered MIS SaaS
#  Built with Streamlit + Claude AI + Plotly
# ============================================================
import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from plotly.subplots import make_subplots
import json, requests, io, datetime

# ── PAGE CONFIG ──────────────────────────────────────────────
st.set_page_config(
    page_title="MIS AutoReport Pro",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded"
)

# ── GLOBAL STYLES ────────────────────────────────────────────
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Syne:wght@700;800&family=DM+Sans:wght@400;500&family=DM+Mono:wght@400;500&display=swap');

html, body, [class*="css"] {
    font-family: 'DM Sans', sans-serif;
    background-color: #07070a;
    color: #eaeaf5;
}
.stApp { background: #07070a; }

/* SIDEBAR */
section[data-testid="stSidebar"] {
    background: #0e0e16 !important;
    border-right: 1px solid rgba(255,255,255,0.06);
}

/* METRIC CARDS */
.kpi-grid { display: grid; grid-template-columns: repeat(4,1fr); gap: 12px; margin: 16px 0; }
.kpi-card {
    background: #0f0f18;
    border: 1px solid rgba(245,200,66,0.15);
    border-radius: 14px;
    padding: 18px 16px;
    text-align: center;
    transition: border-color .2s;
}
.kpi-card:hover { border-color: rgba(245,200,66,0.4); }
.kpi-num  { font-family:'Syne',sans-serif; font-size:26px; font-weight:800; color:#f5c842; }
.kpi-lbl  { font-size:11px; color:rgba(234,234,245,0.45); margin-top:4px; font-family:'DM Mono',monospace; letter-spacing:1px; }
.kpi-delta{ font-size:11px; color:#00e5a0; margin-top:3px; }

/* AI REPORT BOX */
.ai-report {
    background: linear-gradient(135deg,rgba(245,200,66,0.06),rgba(0,229,160,0.03));
    border: 1px solid rgba(245,200,66,0.25);
    border-radius: 16px;
    padding: 24px 22px;
    margin: 16px 0;
    font-size: 14px;
    line-height: 1.8;
    color: #eaeaf5;
}
.ai-badge {
    display:inline-flex; align-items:center; gap:6px;
    background:rgba(245,200,66,0.1); border:1px solid rgba(245,200,66,0.25);
    border-radius:4px; padding:3px 10px; margin-bottom:12px;
    font-family:'DM Mono',monospace; font-size:10px; color:#f5c842; letter-spacing:2px;
}

/* SECTION HEADERS */
.sec-title {
    font-family:'Syne',sans-serif; font-size:18px; font-weight:800;
    color:#eaeaf5; margin:28px 0 12px; display:flex; align-items:center; gap:10px;
}
.sec-title::after {
    content:''; flex:1; height:1px; background:rgba(255,255,255,0.07);
}

/* CHART CONTAINER */
.chart-wrap {
    background:#0f0f18; border:1px solid rgba(255,255,255,0.07);
    border-radius:14px; padding:4px; margin-bottom:14px;
}

/* HERO */
.hero-bar {
    background: linear-gradient(135deg,#0f0f18,#12121f);
    border: 1px solid rgba(245,200,66,0.12);
    border-radius:16px; padding:22px 24px;
    display:flex; align-items:center; justify-content:space-between;
    margin-bottom:24px;
}
.hero-title { font-family:'Syne',sans-serif; font-size:24px; font-weight:800; color:#f5c842; }
.hero-sub   { font-size:13px; color:rgba(234,234,245,0.5); margin-top:4px; }
.hero-badge {
    background:rgba(0,229,160,0.1); border:1px solid rgba(0,229,160,0.25);
    border-radius:6px; padding:6px 14px;
    font-family:'DM Mono',monospace; font-size:10px; color:#00e5a0; letter-spacing:1px;
}
</style>
""", unsafe_allow_html=True)

# ── CONSTANTS ────────────────────────────────────────────────
CHART_THEME = dict(
    plot_bgcolor  ='#0f0f18',
    paper_bgcolor ='#0f0f18',
    font_color    ='#eaeaf5',
    font_family   ='DM Sans',
    margin        =dict(l=10,r=10,t=40,b=10),
    colorway      =['#f5c842','#00e5a0','#4a9eff','#ff4455','#b06bff','#ff8c42'],
)
GOLD   = '#f5c842'
GREEN  = '#00e5a0'
BLUE   = '#4a9eff'

# ── HELPERS ─────────────────────────────────────────────────
def valid(col, df):
    return col and isinstance(col, str) and col in df.columns

def fmt_inr(val):
    if val >= 1e7:  return f"₹{val/1e7:.2f} Cr"
    if val >= 1e5:  return f"₹{val/1e5:.1f} L"
    if val >= 1e3:  return f"₹{val/1e3:.1f} K"
    return f"₹{val:,.0f}"

def safe_numeric(df, col):
    if not valid(col, df): return df
    try:
        df[col] = pd.to_numeric(df[col], errors='coerce')
    except Exception:
        df[col] = df[col].astype(str).str.replace(r'[^\d.\-]','',regex=True)
        df[col] = pd.to_numeric(df[col], errors='coerce')
    return df

def safe_datetime(df, col):
    if not valid(col, df): return df
    df[col] = pd.to_datetime(df[col], errors='coerce')
    return df

# ── AI COLUMN DETECTION ──────────────────────────────────────
def detect_columns_ai(columns, sample_str, api_key):
    prompt = f"""You are a senior data analyst. Look at these Excel column names and sample data.
Identify which columns represent each role below.

Column names: {columns}
Sample (3 rows):
{sample_str}

Return ONLY valid JSON — no explanation, no markdown:
{{
  "date_col":     "exact column name or null",
  "value_col":    "exact column name or null",
  "qty_col":      "exact column name or null",
  "category_col": "exact column name or null",
  "name_col":     "exact column name or null",
  "product_col":  "exact column name or null",
  "manager_col":  "exact column name or null",
  "summary": "one sentence: what this dataset is about"
}}
Use EXACT names. null if not found."""
    try:
        r = requests.post(
            "https://api.anthropic.com/v1/messages",
            headers={"x-api-key": api_key,
                     "anthropic-version":"2023-06-01",
                     "content-type":"application/json"},
            json={"model":"claude-sonnet-4-6","max_tokens":400,
                  "messages":[{"role":"user","content":prompt}]},
            timeout=25
        )
        txt = r.json()["content"][0]["text"].strip()
        if "```" in txt:
            txt = txt.split("```")[1]
            if txt.startswith("json"): txt=txt[4:]
        return json.loads(txt.strip())
    except:
        return None

def detect_columns_fallback(df):
    cols_lower = {c.lower():c for c in df.columns}
    def find(*kws):
        for kw in kws:
            m=[v for k,v in cols_lower.items() if kw in k]
            if m: return m[0]
        return None
    d = find('date','time','day','dt','दिनांक')
    if not d:
        dc = df.select_dtypes(include='datetime64').columns
        d = dc[0] if len(dc) else None
    v = find('value','sales','amount','revenue','total','sale')
    if not v:
        nc = df.select_dtypes(include=['float64','int64']).columns
        v = nc[0] if len(nc) else None
    return {
        "date_col"    : d,
        "value_col"   : v,
        "qty_col"     : find('qty','quantity','units','pieces','volume'),
        "category_col": find('region','area','zone','city','state','branch','territory'),
        "name_col"    : find('party','customer','client','buyer','account','distributor'),
        "product_col" : find('item','product','sku','goods','material','article'),
        "manager_col" : find('manager','salesperson','sales person','executive','rep'),
        "summary"     : "Auto-detected via keyword matching"
    }

# ── AI REPORT GENERATION ─────────────────────────────────────
def generate_ai_report(stats_dict, api_key):
    prompt = f"""You are a Senior MIS Manager and Business Analyst with 15 years of experience in Indian SMEs.

Analyze this sales data and write a sharp, actionable MIS report for the CFO/Owner.

Data Summary:
{json.dumps(stats_dict, indent=2, default=str)}

Write the report in this EXACT structure (use these headers):
**📊 Executive Summary**
(2-3 sentences: key numbers, overall health)

**🔥 Top Insights**
(3-4 bullet points: most important patterns)

**⚠️ Red Flags**
(2-3 issues the management must act on immediately)

**✅ Recommendations**
(3 concrete actions with expected impact)

**💰 Revenue Opportunity**
(1 specific untapped opportunity with estimated value)

Keep it sharp, data-driven, and India-context aware. Use ₹ for currency. Max 300 words."""
    try:
        r = requests.post(
            "https://api.anthropic.com/v1/messages",
            headers={"x-api-key": api_key,
                     "anthropic-version":"2023-06-01",
                     "content-type":"application/json"},
            json={"model":"claude-sonnet-4-6","max_tokens":600,
                  "messages":[{"role":"user","content":prompt}]},
            timeout=30
        )
        return r.json()["content"][0]["text"]
    except Exception as e:
        return f"AI report generation failed: {e}"

# ═══════════════════════════════════════════════════════════════
#  SIDEBAR
# ═══════════════════════════════════════════════════════════════
with st.sidebar:
    st.markdown("""
    <div style='text-align:center;padding:16px 0 8px'>
      <div style='font-family:Syne,sans-serif;font-size:20px;font-weight:800;color:#f5c842'>
        MIS AutoReport
      </div>
      <div style='font-size:10px;color:rgba(234,234,245,0.4);
                  font-family:DM Mono,monospace;letter-spacing:2px;margin-top:4px'>
        POWERED BY AI
      </div>
    </div>
    """, unsafe_allow_html=True)
    st.divider()

    api_key = st.text_input(
        "🔑 Claude API Key (optional)",
        type="password",
        help="Add for AI report generation. Free at console.anthropic.com"
    )
    st.caption("Without API key — smart keyword detection still works!")
    st.divider()

    uploaded = st.file_uploader(
        "📂 Upload Excel / CSV",
        type=["xlsx","xls","csv"],
        help="Any Excel structure — AI will detect columns"
    )

    st.divider()
    st.markdown("""
    <div style='font-size:11px;color:rgba(234,234,245,0.35);
                font-family:DM Mono,monospace;line-height:1.8'>
    ✅ Any Excel structure<br>
    ✅ Auto column detection<br>
    ✅ AI MIS report<br>
    ✅ Multi-sheet support<br>
    ✅ Export to CSV
    </div>
    """, unsafe_allow_html=True)

# ═══════════════════════════════════════════════════════════════
#  MAIN AREA
# ═══════════════════════════════════════════════════════════════
if not uploaded:
    # LANDING PAGE
    st.markdown("""
    <div class='hero-bar'>
      <div>
        <div class='hero-title'>📊 MIS AutoReport Pro</div>
        <div class='hero-sub'>Upload any Excel → AI builds your MIS dashboard in 60 seconds</div>
      </div>
      <div class='hero-badge'>AI POWERED</div>
    </div>
    """, unsafe_allow_html=True)

    c1,c2,c3 = st.columns(3)
    with c1:
        st.markdown("""
        <div class='kpi-card'>
          <div class='kpi-num'>60s</div>
          <div class='kpi-lbl'>REPORT READY IN</div>
          <div class='kpi-delta'>vs 3-5 days manual</div>
        </div>""", unsafe_allow_html=True)
    with c2:
        st.markdown("""
        <div class='kpi-card'>
          <div class='kpi-num'>Any</div>
          <div class='kpi-lbl'>EXCEL FORMAT</div>
          <div class='kpi-delta'>AI auto-detects structure</div>
        </div>""", unsafe_allow_html=True)
    with c3:
        st.markdown("""
        <div class='kpi-card'>
          <div class='kpi-num'>₹0</div>
          <div class='kpi-lbl'>SETUP COST</div>
          <div class='kpi-delta'>Free to start today</div>
        </div>""", unsafe_allow_html=True)

    st.markdown("""
    <div style='margin-top:32px;padding:24px;background:#0f0f18;
                border:1px solid rgba(255,255,255,0.07);border-radius:16px'>
      <div style='font-family:Syne,sans-serif;font-size:16px;
                  font-weight:800;margin-bottom:16px'>
        🚀 How It Works
      </div>
      <div style='display:grid;grid-template-columns:repeat(3,1fr);gap:16px'>
        <div style='text-align:center;padding:16px;background:rgba(245,200,66,0.04);
                    border-radius:10px;border:1px solid rgba(245,200,66,0.1)'>
          <div style='font-size:28px'>📂</div>
          <div style='font-size:13px;margin-top:8px;font-weight:500'>Upload Excel</div>
          <div style='font-size:11px;color:rgba(234,234,245,0.45);margin-top:4px'>
            Any format, any columns
          </div>
        </div>
        <div style='text-align:center;padding:16px;background:rgba(0,229,160,0.04);
                    border-radius:10px;border:1px solid rgba(0,229,160,0.1)'>
          <div style='font-size:28px'>🤖</div>
          <div style='font-size:13px;margin-top:8px;font-weight:500'>AI Analyzes</div>
          <div style='font-size:11px;color:rgba(234,234,245,0.45);margin-top:4px'>
            Detects columns, patterns
          </div>
        </div>
        <div style='text-align:center;padding:16px;background:rgba(74,158,255,0.04);
                    border-radius:10px;border:1px solid rgba(74,158,255,0.1)'>
          <div style='font-size:28px'>📊</div>
          <div style='font-size:13px;margin-top:8px;font-weight:500'>Dashboard Ready</div>
          <div style='font-size:11px;color:rgba(234,234,245,0.45);margin-top:4px'>
            Charts + AI report
          </div>
        </div>
      </div>
    </div>
    """, unsafe_allow_html=True)
    st.stop()

# ── READ FILE ─────────────────────────────────────────────────
try:
    if uploaded.name.endswith('.csv'):
        df_raw = pd.read_csv(uploaded)
        sheet_used = "CSV"
    else:
        xl = pd.ExcelFile(uploaded)
        sheets = xl.sheet_names
        # Auto pick sheet with most rows
        best_sheet = sheets[0]
        best_rows  = 0
        for s in sheets:
            try:
                tmp = pd.read_excel(uploaded, sheet_name=s, nrows=5)
                if len(tmp.columns) > best_rows:
                    best_rows  = len(tmp.columns)
                    best_sheet = s
            except: pass

        if len(sheets) > 1:
            with st.sidebar:
                best_sheet = st.selectbox("📋 Sheet", sheets,
                    index=sheets.index(best_sheet))
        df_raw     = pd.read_excel(uploaded, sheet_name=best_sheet)
        sheet_used = best_sheet
except Exception as e:
    st.error(f"❌ File read error: {e}")
    st.stop()

# Clean unnamed columns & all-empty rows
df_raw = df_raw.dropna(how='all')
if df_raw.columns.str.startswith('Unnamed').sum() > len(df_raw.columns)*0.5:
    df_raw.columns = df_raw.iloc[0]
    df_raw = df_raw[1:].reset_index(drop=True)
df_raw.columns = [str(c).strip() for c in df_raw.columns]
col_list = [c for c in df_raw.columns if 'Unnamed' not in str(c)]
df_raw   = df_raw[col_list]

# ── AI COLUMN DETECTION ──────────────────────────────────────
with st.spinner("🤖 AI is reading your file structure..."):
    sample_str = df_raw.head(3).to_string()
    if api_key:
        mapping = detect_columns_ai(col_list, sample_str, api_key)
    else:
        mapping = None
    if not mapping:
        mapping = detect_columns_fallback(df_raw)

d = mapping.get('date_col')
v = mapping.get('value_col')
q = mapping.get('qty_col')
c = mapping.get('category_col')
n = mapping.get('name_col')
p = mapping.get('product_col')
m = mapping.get('manager_col')

# ── COLUMN OVERRIDE IN SIDEBAR ───────────────────────────────
with st.sidebar:
    st.markdown("**⚙️ Column Mapping**")
    with st.expander("Verify / Override", expanded=False):
        opts = ["(none)"] + col_list
        def pick(label, current):
            idx = opts.index(current) if current in opts else 0
            return st.selectbox(label, opts, index=idx)
        d = pick("📅 Date",     d)
        v = pick("💰 Value",    v)
        q = pick("📦 Qty",      q)
        c = pick("🗺️ Category", c)
        n = pick("🏢 Customer", n)
        p = pick("📦 Product",  p)
        m = pick("👤 Manager",  m)
        for k in [d,v,q,c,n,p,m]:
            if k == "(none)": k = None

d = None if d == "(none)" else d
v = None if v == "(none)" else v
q = None if q == "(none)" else q
c = None if c == "(none)" else c
n = None if n == "(none)" else n
p = None if p == "(none)" else p
m = None if m == "(none)" else m

# ── PREPARE DATA ─────────────────────────────────────────────
df = df_raw.copy()
df = safe_datetime(df, d)
df = safe_numeric(df, v)
df = safe_numeric(df, q)

# ── SIDEBAR FILTERS ──────────────────────────────────────────
with st.sidebar:
    st.divider()
    st.markdown("**🔍 Filters**")

    filtered = df.copy()

    if valid(d, df):
        dmin = df[d].min()
        dmax = df[d].max()
        if pd.notnull(dmin) and pd.notnull(dmax):
            dr = st.date_input("Date Range",
                [dmin.date(), dmax.date()])
            if len(dr)==2:
                filtered = filtered[
                    (filtered[d]>=pd.Timestamp(dr[0])) &
                    (filtered[d]<=pd.Timestamp(dr[1]))]

    if valid(c, df):
        cats = sorted(df[c].dropna().unique().tolist())
        sel  = st.multiselect(c, cats, default=cats)
        filtered = filtered[filtered[c].isin(sel)]

    if valid(m, df):
        mgrs = sorted(df[m].dropna().unique().tolist())
        selm = st.multiselect("Manager", mgrs, default=mgrs)
        filtered = filtered[filtered[m].isin(selm)]

# ═══════════════════════════════════════════════════════════════
#  HEADER
# ═══════════════════════════════════════════════════════════════
ai_label = "🤖 AI + Smart Detection" if api_key else "⚡ Smart Detection"
st.markdown(f"""
<div class='hero-bar'>
  <div>
    <div class='hero-title'>📊 MIS Dashboard</div>
    <div class='hero-sub'>{mapping.get('summary','Sales Analysis')} &nbsp;·&nbsp;
        Sheet: <b>{sheet_used}</b> &nbsp;·&nbsp;
        {len(filtered):,} records
    </div>
  </div>
  <div class='hero-badge'>{ai_label}</div>
</div>
""", unsafe_allow_html=True)

# ═══════════════════════════════════════════════════════════════
#  KPI CARDS
# ═══════════════════════════════════════════════════════════════
total_val  = filtered[v].sum()      if valid(v,filtered) else 0
total_qty  = filtered[q].sum()      if valid(q,filtered) else 0
total_cust = filtered[n].nunique()  if valid(n,filtered) else 0
total_txn  = len(filtered)
avg_order  = total_val/total_txn    if total_txn else 0

k1,k2,k3,k4 = st.columns(4)
for col_el, num, lbl, delta in [
    (k1, fmt_inr(total_val), "TOTAL SALES",     "Revenue"),
    (k2, f"{total_qty/1000:.1f}K" if total_qty>1000 else f"{total_qty:.0f}", "TOTAL QTY", "Units sold"),
    (k3, f"{total_cust:,}",        "CUSTOMERS",      "Unique buyers"),
    (k4, fmt_inr(avg_order),       "AVG ORDER VALUE", "Per transaction"),
]:
    col_el.markdown(f"""
    <div class='kpi-card'>
      <div class='kpi-num'>{num}</div>
      <div class='kpi-lbl'>{lbl}</div>
      <div class='kpi-delta'>{delta}</div>
    </div>""", unsafe_allow_html=True)

st.markdown("<div style='height:8px'></div>", unsafe_allow_html=True)

# ═══════════════════════════════════════════════════════════════
#  CHARTS — Row 1: Trend + Category
# ═══════════════════════════════════════════════════════════════
ch1, ch2 = st.columns([3,2])

# Monthly Trend
if valid(d,filtered) and valid(v,filtered):
    with ch1:
        st.markdown("<div class='sec-title'>📈 Sales Trend</div>", unsafe_allow_html=True)
        try:
            tr = filtered.copy()
            tr['_mo'] = tr[d].dt.to_period('M').astype(str)
            mo = tr.groupby('_mo')[v].sum().reset_index()
            mo.columns = ['Month','Sales']
            mo['MoM'] = mo['Sales'].pct_change()*100

            fig = make_subplots(specs=[[{"secondary_y":True}]])
            fig.add_trace(go.Bar(x=mo['Month'], y=mo['Sales'],
                name='Sales', marker_color=GOLD, opacity=0.85))
            fig.add_trace(go.Scatter(x=mo['Month'], y=mo['MoM'],
                name='MoM %', mode='lines+markers',
                line=dict(color=GREEN,width=2),
                marker=dict(size=5)), secondary_y=True)
            fig.update_layout(**CHART_THEME, title="Monthly Sales + MoM Growth %",
                              legend=dict(orientation='h',y=1.1))
            fig.update_yaxes(title_text="Sales (₹)", secondary_y=False,
                             gridcolor='rgba(255,255,255,0.04)')
            fig.update_yaxes(title_text="MoM %", secondary_y=True)
            st.plotly_chart(fig, use_container_width=True)
        except Exception as e:
            st.warning(f"Trend chart: {e}")

# Category Pie
if valid(c,filtered) and valid(v,filtered):
    with ch2:
        st.markdown(f"<div class='sec-title'>🗺️ By {c}</div>", unsafe_allow_html=True)
        try:
            cd = filtered.groupby(c)[v].sum().sort_values(ascending=False).head(10).reset_index()
            fig2 = px.pie(cd, names=c, values=v,
                          hole=0.45, color_discrete_sequence=[
                GOLD,'#00e5a0','#4a9eff','#ff4455','#b06bff',
                '#ff8c42','#00ccff','#ffcc00','#ff66aa','#88ff44'])
            fig2.update_layout(**CHART_THEME, title=f"Top {c} Share",
                showlegend=True,
                legend=dict(orientation='v', font=dict(size=10)))
            fig2.update_traces(textposition='inside',
                               textinfo='percent',
                               hovertemplate='%{label}<br>₹%{value:,.0f}')
            st.plotly_chart(fig2, use_container_width=True)
        except Exception as e:
            st.warning(f"Category chart: {e}")

# ═══════════════════════════════════════════════════════════════
#  CHARTS — Row 2: Top Customers + Manager Performance
# ═══════════════════════════════════════════════════════════════
ch3, ch4 = st.columns(2)

if valid(n,filtered) and valid(v,filtered):
    with ch3:
        st.markdown("<div class='sec-title'>🏆 Top 10 Customers</div>",
                    unsafe_allow_html=True)
        try:
            tn = filtered.groupby(n)[v].sum()\
                         .sort_values(ascending=True).tail(10).reset_index()
            fig3 = px.bar(tn, x=v, y=n, orientation='h',
                          color=v, color_continuous_scale='YlOrRd')
            fig3.update_layout(**CHART_THEME,
                title="Revenue by Customer",
                yaxis=dict(categoryorder='total ascending'),
                coloraxis_showscale=False)
            fig3.update_traces(
                hovertemplate='%{y}<br>₹%{x:,.0f}')
            st.plotly_chart(fig3, use_container_width=True)
        except Exception as e:
            st.warning(f"Customer chart: {e}")

if valid(m,filtered) and valid(v,filtered):
    with ch4:
        st.markdown("<div class='sec-title'>👤 Manager Performance</div>",
                    unsafe_allow_html=True)
        try:
            mgdf = filtered.groupby(m)[v].sum()\
                           .sort_values(ascending=False).reset_index()
            fig4 = px.bar(mgdf, x=m, y=v,
                          color=v, color_continuous_scale='Blues',
                          text_auto='.2s')
            fig4.update_layout(**CHART_THEME,
                title="Sales by Manager",
                coloraxis_showscale=False)
            fig4.update_traces(textfont_size=11,
                hovertemplate='%{x}<br>₹%{y:,.0f}')
            st.plotly_chart(fig4, use_container_width=True)
        except Exception as e:
            st.warning(f"Manager chart: {e}")

# ═══════════════════════════════════════════════════════════════
#  CHARTS — Row 3: Top Products
# ═══════════════════════════════════════════════════════════════
if valid(p,filtered) and valid(v,filtered):
    st.markdown("<div class='sec-title'>📦 Product Analysis</div>",
                unsafe_allow_html=True)
    try:
        pch1, pch2 = st.columns(2)
        prod = filtered.groupby(p).agg(
            Sales=(v,'sum'),
            Txns=(v,'count')
        ).sort_values('Sales',ascending=False).head(12).reset_index()

        with pch1:
            fig5 = px.bar(prod.head(10), x='Sales', y=p,
                orientation='h', color='Sales',
                color_continuous_scale='Oranges', text_auto='.2s')
            fig5.update_layout(**CHART_THEME,
                title="Top 10 Products by Revenue",
                yaxis=dict(categoryorder='total ascending'),
                coloraxis_showscale=False)
            st.plotly_chart(fig5, use_container_width=True)

        with pch2:
            fig6 = px.scatter(prod, x='Txns', y='Sales',
                size='Sales', color='Sales',
                color_continuous_scale='YlOrRd',
                hover_name=p, text=p)
            fig6.update_layout(**CHART_THEME,
                title="Volume vs Revenue (bubble = revenue)")
            fig6.update_traces(textposition='top center',
                               textfont=dict(size=9))
            st.plotly_chart(fig6, use_container_width=True)
    except Exception as e:
        st.warning(f"Product charts: {e}")

# ═══════════════════════════════════════════════════════════════
#  AI MIS REPORT
# ═══════════════════════════════════════════════════════════════
st.markdown("<div class='sec-title'>🤖 AI-Generated MIS Report</div>",
            unsafe_allow_html=True)

# Build stats for AI
stats = {
    "total_sales"     : fmt_inr(total_val),
    "total_transactions": total_txn,
    "total_customers" : total_cust,
    "avg_order_value" : fmt_inr(avg_order),
}
if valid(v,filtered) and valid(d,filtered):
    try:
        filtered['_mo'] = filtered[d].dt.to_period('M').astype(str)
        mo_sales = filtered.groupby('_mo')[v].sum()
        stats["best_month"]  = str(mo_sales.idxmax())
        stats["worst_month"] = str(mo_sales.idxmin())
        stats["last_month_sales"] = fmt_inr(float(mo_sales.iloc[-1]))
    except: pass

if valid(c,filtered) and valid(v,filtered):
    try:
        top_cat = filtered.groupby(c)[v].sum().idxmax()
        stats["top_region"] = str(top_cat)
    except: pass

if valid(n,filtered) and valid(v,filtered):
    try:
        top_cust = filtered.groupby(n)[v].sum().idxmax()
        stats["top_customer"] = str(top_cust)
    except: pass

if valid(m,filtered) and valid(v,filtered):
    try:
        top_mgr = filtered.groupby(m)[v].sum().idxmax()
        stats["top_manager"] = str(top_mgr)
    except: pass

if api_key:
    gen_btn = st.button("🚀 Generate AI MIS Report", type="primary",
                        use_container_width=True)
    if gen_btn:
        with st.spinner("🤖 Claude is writing your MIS report..."):
            report = generate_ai_report(stats, api_key)
        st.markdown(f"""
        <div class='ai-report'>
          <div class='ai-badge'>🤖 AI GENERATED · CLAUDE</div>
          {report.replace(chr(10),'<br>')}
        </div>""", unsafe_allow_html=True)
else:
    # Static smart report without API key
    report_lines = [
        f"**📊 Executive Summary**",
        f"Total sales: **{fmt_inr(total_val)}** across **{total_txn:,}** transactions "
        f"from **{total_cust:,}** unique customers. Average order value: **{fmt_inr(avg_order)}**.",
        "",
    ]
    if valid(c,filtered) and valid(v,filtered):
        try:
            top3 = filtered.groupby(c)[v].sum().sort_values(ascending=False).head(3)
            report_lines.append(f"**🔥 Top Regions:** " + ", ".join([f"{i} (₹{v2/1e5:.1f}L)" for i,v2 in top3.items()]))
        except: pass
    if valid(n,filtered) and valid(v,filtered):
        try:
            top_c = filtered.groupby(n)[v].sum().idxmax()
            report_lines.append(f"**🏆 Top Customer:** {top_c}")
        except: pass
    report_lines += [
        "",
        "**💡 Add Claude API Key** in the sidebar to get AI-powered insights, "
        "red flags, and revenue recommendations!"
    ]
    st.markdown(f"""
    <div class='ai-report'>
      <div class='ai-badge'>⚡ SMART REPORT</div>
      {'<br>'.join(report_lines)}
    </div>""", unsafe_allow_html=True)

# ═══════════════════════════════════════════════════════════════
#  DATA TABLE + DOWNLOAD
# ═══════════════════════════════════════════════════════════════
st.markdown("<div class='sec-title'>📋 Filtered Data</div>",
            unsafe_allow_html=True)
show_cols = [x for x in [d,v,q,c,n,p,m] if valid(x,filtered)]
display_df = filtered[show_cols] if show_cols else filtered
st.dataframe(display_df.head(500), use_container_width=True,
             height=280)

dl1, dl2 = st.columns(2)
dl1.download_button(
    "📥 Download Full Report (CSV)",
    filtered.to_csv(index=False),
    f"MIS_Report_{datetime.date.today()}.csv",
    "text/csv", use_container_width=True
)
if valid(v,filtered) and valid(d,filtered):
    try:
        filtered['_mo'] = filtered[d].dt.to_period('M').astype(str)
        mo_sum = filtered.groupby('_mo')[v].sum().reset_index()
        mo_sum.columns = ['Month','Sales']
        dl2.download_button(
            "📥 Monthly Summary (CSV)",
            mo_sum.to_csv(index=False),
            f"Monthly_Summary_{datetime.date.today()}.csv",
            "text/csv", use_container_width=True
        )
    except: pass

# ── FOOTER ───────────────────────────────────────────────────
st.markdown("""
<div style='margin-top:40px;text-align:center;
            font-family:DM Mono,monospace;font-size:10px;
            color:rgba(234,234,245,0.2);letter-spacing:2px'>
  MIS AUTOREPORT PRO · BUILT BY SATYENDRA · POWERED BY CLAUDE AI
</div>
""", unsafe_allow_html=True)
