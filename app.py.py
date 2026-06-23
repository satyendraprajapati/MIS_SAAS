# ============================================================
#  MIS AutoReport Pro — Universal AI-Powered MIS SaaS
#  Works with ANY company's Excel file — crash-proof
# ============================================================
import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from plotly.subplots import make_subplots
import json, requests, datetime, traceback

st.set_page_config(
    page_title="MIS AutoReport Pro",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded"
)

st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Syne:wght@700;800&family=DM+Sans:wght@400;500&family=DM+Mono:wght@400;500&display=swap');
html,body,[class*="css"]{font-family:'DM Sans',sans-serif;background:#07070a;color:#eaeaf5}
.stApp{background:#07070a}
section[data-testid="stSidebar"]{background:#0e0e16!important;border-right:1px solid rgba(255,255,255,.06)}
.kpi-card{background:#0f0f18;border:1px solid rgba(245,200,66,.15);border-radius:14px;padding:18px 16px;text-align:center}
.kpi-num{font-family:'Syne',sans-serif;font-size:26px;font-weight:800;color:#f5c842}
.kpi-lbl{font-size:11px;color:rgba(234,234,245,.45);margin-top:4px;font-family:'DM Mono',monospace;letter-spacing:1px}
.kpi-delta{font-size:11px;color:#00e5a0;margin-top:3px}
.ai-report{background:linear-gradient(135deg,rgba(245,200,66,.06),rgba(0,229,160,.03));border:1px solid rgba(245,200,66,.25);border-radius:16px;padding:24px 22px;margin:16px 0;font-size:14px;line-height:1.8;color:#eaeaf5}
.ai-badge{display:inline-flex;align-items:center;gap:6px;background:rgba(245,200,66,.1);border:1px solid rgba(245,200,66,.25);border-radius:4px;padding:3px 10px;margin-bottom:12px;font-family:'DM Mono',monospace;font-size:10px;color:#f5c842;letter-spacing:2px}
.sec-title{font-family:'Syne',sans-serif;font-size:18px;font-weight:800;color:#eaeaf5;margin:28px 0 12px}
.hero-bar{background:linear-gradient(135deg,#0f0f18,#12121f);border:1px solid rgba(245,200,66,.12);border-radius:16px;padding:22px 24px;display:flex;align-items:center;justify-content:space-between;margin-bottom:24px}
.hero-title{font-family:'Syne',sans-serif;font-size:24px;font-weight:800;color:#f5c842}
.hero-sub{font-size:13px;color:rgba(234,234,245,.5);margin-top:4px}
.hero-badge{background:rgba(0,229,160,.1);border:1px solid rgba(0,229,160,.25);border-radius:6px;padding:6px 14px;font-family:'DM Mono',monospace;font-size:10px;color:#00e5a0;letter-spacing:1px}
</style>
""", unsafe_allow_html=True)

CHART_THEME = dict(
    plot_bgcolor='#0f0f18', paper_bgcolor='#0f0f18',
    font_color='#eaeaf5', font_family='DM Sans',
    margin=dict(l=10,r=10,t=40,b=10),
    colorway=['#f5c842','#00e5a0','#4a9eff','#ff4455','#b06bff','#ff8c42'],
)
GOLD, GREEN, BLUE = '#f5c842', '#00e5a0', '#4a9eff'

# ════════════════════════════════════════════════════════
#  UNIVERSAL HELPERS — crash-proof for any Excel
# ════════════════════════════════════════════════════════

def valid(col, df):
    """Check column exists and is a single string name."""
    return bool(col and isinstance(col, str) and col in df.columns)

def get_col(df, col):
    """
    ALWAYS return a single pandas Series.
    Handles: duplicate columns, mixed types, empty columns.
    Returns None if column not usable.
    """
    if not valid(col, df):
        return None
    try:
        s = df[col]
        # Duplicate col names → DataFrame → take first column
        if isinstance(s, pd.DataFrame):
            s = s.iloc[:, 0]
        if not isinstance(s, pd.Series):
            return None
        return s
    except Exception:
        return None

def to_numeric_safe(series):
    """Convert any series to numeric — handles ₹, commas, mixed text."""
    if series is None:
        return None
    try:
        cleaned = (series.astype(str)
                   .str.replace(r'[₹,\s]', '', regex=True)
                   .str.replace(r'[^\d.\-]', '', regex=True)
                   .replace('', pd.NA))
        return pd.to_numeric(cleaned, errors='coerce')
    except Exception:
        return None

def to_datetime_safe(series):
    """Convert any series to datetime — handles Indian DD/MM/YYYY."""
    if series is None:
        return None
    try:
        return pd.to_datetime(series, errors='coerce', dayfirst=True)
    except Exception:
        return None

def safe_sum(df, col):
    s = get_col(df, col)
    if s is None: return 0
    try: return float(s.sum())
    except: return 0

def safe_nunique(df, col):
    s = get_col(df, col)
    if s is None: return 0
    try: return int(s.nunique())
    except: return 0

def safe_unique_list(series):
    """Get sorted unique values as strings — handles mixed int/str."""
    if series is None: return []
    try:
        vals = series.dropna().unique().tolist()
        return sorted([str(x) for x in vals])
    except: return []

def safe_groupby_sum(df, group_col, val_col):
    """groupby().sum() that never crashes."""
    if not valid(group_col, df) or not valid(val_col, df):
        return pd.DataFrame()
    try:
        gc = get_col(df, group_col)
        vc = get_col(df, val_col)
        if gc is None or vc is None: return pd.DataFrame()
        tmp = pd.DataFrame({'_g': gc.astype(str), '_v': vc})
        result = tmp.groupby('_g')['_v'].sum().reset_index()
        result.columns = [group_col, val_col]
        return result
    except: return pd.DataFrame()

def fmt_inr(val):
    try:
        val = float(val)
        if val >= 1e7:  return f"₹{val/1e7:.2f} Cr"
        if val >= 1e5:  return f"₹{val/1e5:.1f} L"
        if val >= 1e3:  return f"₹{val/1e3:.1f} K"
        return f"₹{val:,.0f}"
    except: return "₹0"

def clean_df(raw):
    """
    Universal Excel cleaner:
    - Remove all-empty rows/cols
    - Fix unnamed headers (try to detect real header row)
    - Rename duplicate columns
    - Strip whitespace from column names
    """
    df = raw.copy()
    df.dropna(how='all', inplace=True)
    df.dropna(axis=1, how='all', inplace=True)

    # If >50% columns are 'Unnamed' → first row is probably header
    unnamed_pct = df.columns.str.startswith('Unnamed').sum() / max(len(df.columns),1)
    if unnamed_pct > 0.5:
        df.columns = df.iloc[0].astype(str).str.strip()
        df = df[1:].reset_index(drop=True)

    # Strip whitespace + newlines from column names
    df.columns = [str(c).strip().replace('\n',' ') for c in df.columns]

    # Remove still-unnamed columns
    df = df[[c for c in df.columns if not c.startswith('Unnamed')]]

    # Rename duplicates: Value → Value, Value_1, Value_2
    seen = {}
    new_cols = []
    for c in df.columns:
        if c in seen:
            seen[c] += 1
            new_cols.append(f"{c}_{seen[c]}")
        else:
            seen[c] = 0
            new_cols.append(c)
    df.columns = new_cols

    # Remove rows where ALL numeric columns are NaN
    num_cols = df.select_dtypes(include=['float64','int64']).columns
    if len(num_cols) > 0:
        df = df.dropna(subset=num_cols, how='all')

    df = df.reset_index(drop=True)
    return df

# ════════════════════════════════════════════════════════
#  AI FUNCTIONS
# ════════════════════════════════════════════════════════

def detect_columns_ai(columns, sample_str, api_key):
    prompt = f"""You are a senior data analyst. Look at these Excel column names and sample data.
Identify which columns represent each role below.

Column names: {columns}
Sample (3 rows):
{sample_str}

Return ONLY valid JSON:
{{
  "date_col":     "exact column name or null",
  "value_col":    "exact column name or null",
  "qty_col":      "exact column name or null",
  "category_col": "exact column name or null",
  "name_col":     "exact column name or null",
  "product_col":  "exact column name or null",
  "manager_col":  "exact column name or null",
  "summary": "one sentence about this dataset"
}}
Use EXACT column names. null if not found."""
    try:
        r = requests.post(
            "https://api.anthropic.com/v1/messages",
            headers={"x-api-key": api_key,
                     "anthropic-version": "2023-06-01",
                     "content-type": "application/json"},
            json={"model": "claude-sonnet-4-6", "max_tokens": 400,
                  "messages": [{"role": "user", "content": prompt}]},
            timeout=25
        )
        txt = r.json()["content"][0]["text"].strip()
        if "```" in txt:
            txt = txt.split("```")[1]
            if txt.startswith("json"): txt = txt[4:]
        mapping = json.loads(txt.strip())
        # Validate: only keep keys that actually exist in columns
        for k, v in mapping.items():
            if k == "summary": continue
            if v and v not in columns:
                mapping[k] = None
        return mapping
    except:
        return None

def detect_columns_fallback(df):
    """Keyword-based detection — works without API key."""
    cols_lower = {c.lower(): c for c in df.columns}
    def find(*kws):
        for kw in kws:
            for lc, orig in cols_lower.items():
                if kw in lc: return orig
        return None

    d = find('date','time','day','dt','दिनांक')
    if not d:
        dc = df.select_dtypes(include='datetime64').columns
        d = dc[0] if len(dc) else None

    v = find('value','sales','amount','revenue','total','sale','turnover')
    if not v:
        # Pick numeric column with highest sum (likely to be revenue)
        num_cols = df.select_dtypes(include=['float64','int64']).columns.tolist()
        if num_cols:
            try:
                v = max(num_cols, key=lambda c: df[c].sum())
            except: v = num_cols[0]

    return {
        "date_col"    : d,
        "value_col"   : v,
        "qty_col"     : find('qty','quantity','units','pieces','vol'),
        "category_col": find('region','area','zone','city','state','branch','territory','location'),
        "name_col"    : find('party','customer','client','buyer','account','distributor','name'),
        "product_col" : find('item','product','sku','goods','material','article','desc'),
        "manager_col" : find('manager','salesperson','sales person','executive','rep','agent'),
        "summary"     : "Auto-detected via keyword matching"
    }

def generate_ai_report(stats_dict, api_key):
    prompt = f"""You are a Senior MIS Manager with 15 years experience in Indian SMEs.
Analyze this sales data and write a sharp MIS report for the CFO/Owner.

Data:
{json.dumps(stats_dict, indent=2, default=str)}

Structure:
**📊 Executive Summary** (2-3 sentences)
**🔥 Top Insights** (3-4 bullets)
**⚠️ Red Flags** (2-3 issues to act on)
**✅ Recommendations** (3 concrete actions)
**💰 Revenue Opportunity** (1 untapped opportunity)

Sharp, data-driven, India-context. Use ₹. Max 300 words."""
    try:
        r = requests.post(
            "https://api.anthropic.com/v1/messages",
            headers={"x-api-key": api_key,
                     "anthropic-version": "2023-06-01",
                     "content-type": "application/json"},
            json={"model": "claude-sonnet-4-6", "max_tokens": 600,
                  "messages": [{"role": "user", "content": prompt}]},
            timeout=30
        )
        return r.json()["content"][0]["text"]
    except Exception as e:
        return f"AI report generation failed: {e}"

# ════════════════════════════════════════════════════════
#  SIDEBAR
# ════════════════════════════════════════════════════════
with st.sidebar:
    st.markdown("""
    <div style='text-align:center;padding:16px 0 8px'>
      <div style='font-family:Syne,sans-serif;font-size:20px;font-weight:800;color:#f5c842'>
        MIS AutoReport</div>
      <div style='font-size:10px;color:rgba(234,234,245,.4);font-family:DM Mono,monospace;
                  letter-spacing:2px;margin-top:4px'>POWERED BY AI</div>
    </div>""", unsafe_allow_html=True)
    st.divider()

    api_key = st.text_input("🔑 Claude API Key (optional)",
        type="password",
        help="For AI report. Free at console.anthropic.com")
    st.caption("Without key — smart detection still works!")
    st.divider()

    uploaded = st.file_uploader("📂 Upload Excel / CSV",
        type=["xlsx","xls","csv"],
        help="Any company, any Excel structure — universal!")
    st.divider()
    st.markdown("""
    <div style='font-size:11px;color:rgba(234,234,245,.35);font-family:DM Mono,monospace;line-height:1.9'>
    ✅ Any company's Excel<br>✅ Any column structure<br>
    ✅ Multi-sheet support<br>✅ AI column detection<br>
    ✅ Crash-proof parsing
    </div>""", unsafe_allow_html=True)

# ════════════════════════════════════════════════════════
#  LANDING PAGE
# ════════════════════════════════════════════════════════
if not uploaded:
    st.markdown("""
    <div class='hero-bar'>
      <div>
        <div class='hero-title'>📊 MIS AutoReport Pro</div>
        <div class='hero-sub'>Upload any company's Excel → AI builds dashboard in 60 seconds</div>
      </div>
      <div class='hero-badge'>UNIVERSAL · AI</div>
    </div>""", unsafe_allow_html=True)

    c1,c2,c3 = st.columns(3)
    for col_, num, lbl, sub in [
        (c1,"Any","COMPANY","Any Excel structure"),
        (c2,"60s","REPORT IN","vs 3-5 days manual"),
        (c3,"₹0","SETUP COST","Free to start"),
    ]:
        col_.markdown(f"""
        <div class='kpi-card'>
          <div class='kpi-num'>{num}</div>
          <div class='kpi-lbl'>{lbl}</div>
          <div class='kpi-delta'>{sub}</div>
        </div>""", unsafe_allow_html=True)
    st.stop()

# ════════════════════════════════════════════════════════
#  READ FILE — crash-proof
# ════════════════════════════════════════════════════════
df_raw = None
sheet_used = "Data"

try:
    if uploaded.name.lower().endswith('.csv'):
        try:
            df_raw = pd.read_csv(uploaded, encoding='utf-8')
        except UnicodeDecodeError:
            uploaded.seek(0)
            df_raw = pd.read_csv(uploaded, encoding='latin-1')
        sheet_used = "CSV"
    else:
        xl = pd.ExcelFile(uploaded)
        sheets = xl.sheet_names

        # Score sheets by data density
        scores = {}
        for s in sheets:
            try:
                tmp = pd.read_excel(uploaded, sheet_name=s, nrows=20)
                clean = sum(1 for c in tmp.columns if 'Unnamed' not in str(c))
                full  = pd.read_excel(uploaded, sheet_name=s)
                scores[s] = len(full) * clean
            except:
                scores[s] = 0

        best = max(scores, key=scores.get) if scores else sheets[0]

        # Build labels with row counts
        labels = {}
        for s in sheets:
            try:
                n = pd.read_excel(uploaded, sheet_name=s).shape[0]
                labels[s] = f"{s}  ({n:,} rows)"
            except:
                labels[s] = s

        if len(sheets) > 1:
            st.info(f"📋 **{len(sheets)} sheets found** — select your data sheet:")
            chosen_label = st.selectbox("Select Sheet",
                list(labels.values()),
                index=list(labels.keys()).index(best))
            best = [k for k,v in labels.items() if v == chosen_label][0]
            st.success(f"✅ Using: **{best}**")

        df_raw = pd.read_excel(uploaded, sheet_name=best)
        sheet_used = best

except Exception as e:
    st.error(f"❌ Could not read file: {e}")
    st.info("💡 Try: Save your Excel as .xlsx format and re-upload.")
    st.stop()

if df_raw is None or len(df_raw) == 0:
    st.error("❌ File appears empty. Please check the file and re-upload.")
    st.stop()

# ════════════════════════════════════════════════════════
#  CLEAN DATA — universal cleaner
# ════════════════════════════════════════════════════════
with st.spinner("🔄 Cleaning data..."):
    try:
        df_clean = clean_df(df_raw)
    except Exception as e:
        st.warning(f"⚠️ Partial cleaning only: {e}")
        df_clean = df_raw.copy()
        df_clean.columns = [str(c).strip() for c in df_clean.columns]

col_list = list(df_clean.columns)

if len(col_list) == 0:
    st.error("❌ No readable columns found. Check if the file has data.")
    st.stop()

# ════════════════════════════════════════════════════════
#  AI COLUMN DETECTION
# ════════════════════════════════════════════════════════
with st.spinner("🤖 Detecting columns..."):
    sample_str = df_clean.head(3).to_string()
    mapping = None
    if api_key:
        mapping = detect_columns_ai(col_list, sample_str, api_key)
    if not mapping:
        mapping = detect_columns_fallback(df_clean)

d = mapping.get('date_col')
v = mapping.get('value_col')
q = mapping.get('qty_col')
c = mapping.get('category_col')
n = mapping.get('name_col')
p = mapping.get('product_col')
m = mapping.get('manager_col')

# ════════════════════════════════════════════════════════
#  COLUMN OVERRIDE in sidebar
# ════════════════════════════════════════════════════════
with st.sidebar:
    st.markdown("**⚙️ Column Mapping**")
    with st.expander("Verify / Change", expanded=False):
        opts = ["(none)"] + col_list
        def pick(label, cur):
            idx = opts.index(cur) if (cur and cur in opts) else 0
            return st.selectbox(label, opts, index=idx)
        d = pick("📅 Date",      d)
        v = pick("💰 Value",     v)
        q = pick("📦 Qty",       q)
        c = pick("🗺️ Category",  c)
        n = pick("🏢 Customer",  n)
        p = pick("📦 Product",   p)
        m = pick("👤 Manager",   m)

d = None if d == "(none)" else d
v = None if v == "(none)" else v
q = None if q == "(none)" else q
c = None if c == "(none)" else c
n = None if n == "(none)" else n
p = None if p == "(none)" else p
m = None if m == "(none)" else m

# ════════════════════════════════════════════════════════
#  PREPARE TYPED DATA
# ════════════════════════════════════════════════════════
df = df_clean.copy()

# Convert date
if valid(d, df):
    s = get_col(df, d)
    converted = to_datetime_safe(s)
    if converted is not None:
        df[d] = converted

# Convert numeric
for col_ in [v, q]:
    if valid(col_, df):
        s = get_col(df, col_)
        converted = to_numeric_safe(s)
        if converted is not None:
            df[col_] = converted

# ════════════════════════════════════════════════════════
#  SIDEBAR FILTERS — all safe
# ════════════════════════════════════════════════════════
with st.sidebar:
    st.divider()
    st.markdown("**🔍 Filters**")
    filtered = df.copy()

    # Date filter
    if valid(d, df):
        ds = get_col(df, d)
        if ds is not None:
            try:
                dmin, dmax = ds.min(), ds.max()
                if pd.notnull(dmin) and pd.notnull(dmax):
                    dr = st.date_input("Date Range",
                        [dmin.date(), dmax.date()])
                    if len(dr) == 2:
                        ds2 = get_col(filtered, d)
                        if ds2 is not None:
                            mask = (ds2 >= pd.Timestamp(dr[0])) & \
                                   (ds2 <= pd.Timestamp(dr[1]))
                            filtered = filtered[mask]
            except: pass

    # Category filter
    if valid(c, df):
        cs = get_col(df, c)
        if cs is not None:
            try:
                cats = safe_unique_list(cs)
                if cats:
                    sel = st.multiselect(c, cats, default=cats)
                    cf = get_col(filtered, c)
                    if cf is not None:
                        filtered = filtered[cf.astype(str).isin(sel)]
            except: pass

    # Manager filter
    if valid(m, df):
        ms = get_col(df, m)
        if ms is not None:
            try:
                mgrs = safe_unique_list(ms)
                if mgrs:
                    selm = st.multiselect("Manager", mgrs, default=mgrs)
                    mf = get_col(filtered, m)
                    if mf is not None:
                        filtered = filtered[mf.astype(str).isin(selm)]
            except: pass

# ════════════════════════════════════════════════════════
#  HEADER
# ════════════════════════════════════════════════════════
ai_label = "🤖 AI Detected" if api_key else "⚡ Smart Detection"
st.markdown(f"""
<div class='hero-bar'>
  <div>
    <div class='hero-title'>📊 MIS Dashboard</div>
    <div class='hero-sub'>
      {mapping.get('summary','Sales Analysis')} &nbsp;·&nbsp;
      Sheet: <b>{sheet_used}</b> &nbsp;·&nbsp;
      {len(filtered):,} records
    </div>
  </div>
  <div class='hero-badge'>{ai_label}</div>
</div>""", unsafe_allow_html=True)

# ════════════════════════════════════════════════════════
#  KPI CARDS
# ════════════════════════════════════════════════════════
total_val  = safe_sum(filtered, v)
total_qty  = safe_sum(filtered, q)
total_cust = safe_nunique(filtered, n)
total_txn  = len(filtered)
avg_order  = total_val / total_txn if total_txn else 0

k1,k2,k3,k4 = st.columns(4)
for col_el, num, lbl, delta in [
    (k1, fmt_inr(total_val),  "TOTAL SALES",      "Revenue"),
    (k2, f"{total_qty/1000:.1f}K" if total_qty > 999 else f"{int(total_qty):,}", "TOTAL QTY", "Units sold"),
    (k3, f"{total_cust:,}",   "CUSTOMERS",         "Unique buyers"),
    (k4, fmt_inr(avg_order),  "AVG ORDER VALUE",   "Per transaction"),
]:
    col_el.markdown(f"""
    <div class='kpi-card'>
      <div class='kpi-num'>{num}</div>
      <div class='kpi-lbl'>{lbl}</div>
      <div class='kpi-delta'>{delta}</div>
    </div>""", unsafe_allow_html=True)

st.markdown("<div style='height:8px'></div>", unsafe_allow_html=True)

# ════════════════════════════════════════════════════════
#  CHARTS — each in its own try/except
# ════════════════════════════════════════════════════════

def make_chart_safe(func, error_label):
    """Wrap any chart in crash protection."""
    try:
        func()
    except Exception as e:
        st.warning(f"⚠️ {error_label}: {str(e)[:120]}")

# ── Row 1: Monthly Trend + Category Pie ──
ch1, ch2 = st.columns([3,2])

with ch1:
    st.markdown("<div class='sec-title'>📈 Sales Trend</div>", unsafe_allow_html=True)
    def trend_chart():
        if not valid(d, filtered) or not valid(v, filtered): return
        tr = filtered.copy()
        ds = get_col(tr, d)
        vs = get_col(tr, v)
        if ds is None or vs is None: return
        tr['_mo'] = ds.dt.to_period('M').astype(str)
        mo = tr.groupby('_mo')[v].sum().reset_index()
        mo.columns = ['Month','Sales']
        mo['MoM'] = mo['Sales'].pct_change() * 100
        fig = make_subplots(specs=[[{"secondary_y": True}]])
        fig.add_trace(go.Bar(x=mo['Month'], y=mo['Sales'],
            name='Sales', marker_color=GOLD, opacity=0.85))
        fig.add_trace(go.Scatter(x=mo['Month'], y=mo['MoM'],
            name='MoM %', mode='lines+markers',
            line=dict(color=GREEN, width=2), marker=dict(size=5)),
            secondary_y=True)
        fig.update_layout(**CHART_THEME, title="Monthly Sales + MoM %",
            legend=dict(orientation='h', y=1.1))
        fig.update_yaxes(gridcolor='rgba(255,255,255,0.04)')
        st.plotly_chart(fig, use_container_width=True)
    make_chart_safe(trend_chart, "Trend chart")

with ch2:
    st.markdown(f"<div class='sec-title'>🗺️ By {c or 'Category'}</div>", unsafe_allow_html=True)
    def cat_chart():
        if not valid(c, filtered) or not valid(v, filtered): return
        cd = safe_groupby_sum(filtered, c, v)
        if cd.empty: return
        cd = cd.sort_values(v, ascending=False).head(10)
        fig2 = px.pie(cd, names=c, values=v, hole=0.45,
            color_discrete_sequence=[GOLD,GREEN,BLUE,'#ff4455','#b06bff',
                '#ff8c42','#00ccff','#ffcc00','#ff66aa','#88ff44'])
        fig2.update_layout(**CHART_THEME, title=f"Top {c} Share")
        fig2.update_traces(textposition='inside', textinfo='percent',
            hovertemplate='%{label}<br>₹%{value:,.0f}')
        st.plotly_chart(fig2, use_container_width=True)
    make_chart_safe(cat_chart, "Category chart")

# ── Row 2: Top Customers + Manager ──
ch3, ch4 = st.columns(2)

with ch3:
    st.markdown("<div class='sec-title'>🏆 Top 10 Customers</div>", unsafe_allow_html=True)
    def cust_chart():
        if not valid(n, filtered) or not valid(v, filtered): return
        tn = safe_groupby_sum(filtered, n, v)
        if tn.empty: return
        tn = tn.sort_values(v, ascending=True).tail(10)
        fig3 = px.bar(tn, x=v, y=n, orientation='h',
            color=v, color_continuous_scale='YlOrRd')
        fig3.update_layout(**CHART_THEME, title="Revenue by Customer",
            yaxis=dict(categoryorder='total ascending'),
            coloraxis_showscale=False)
        st.plotly_chart(fig3, use_container_width=True)
    make_chart_safe(cust_chart, "Customer chart")

with ch4:
    st.markdown("<div class='sec-title'>👤 Manager Performance</div>", unsafe_allow_html=True)
    def mgr_chart():
        if not valid(m, filtered) or not valid(v, filtered): return
        mgdf = safe_groupby_sum(filtered, m, v)
        if mgdf.empty: return
        mgdf = mgdf.sort_values(v, ascending=False)
        fig4 = px.bar(mgdf, x=m, y=v,
            color=v, color_continuous_scale='Blues', text_auto='.2s')
        fig4.update_layout(**CHART_THEME, title="Sales by Manager",
            coloraxis_showscale=False)
        st.plotly_chart(fig4, use_container_width=True)
    make_chart_safe(mgr_chart, "Manager chart")

# ── Row 3: Product ──
st.markdown("<div class='sec-title'>📦 Product Analysis</div>", unsafe_allow_html=True)
def prod_charts():
    if not valid(p, filtered) or not valid(v, filtered): return
    prod = safe_groupby_sum(filtered, p, v)
    if prod.empty: return
    prod = prod.sort_values(v, ascending=False).head(12)
    pc1, pc2 = st.columns(2)
    with pc1:
        fig5 = px.bar(prod.head(10), x=v, y=p, orientation='h',
            color=v, color_continuous_scale='Oranges', text_auto='.2s')
        fig5.update_layout(**CHART_THEME, title="Top 10 Products",
            yaxis=dict(categoryorder='total ascending'),
            coloraxis_showscale=False)
        st.plotly_chart(fig5, use_container_width=True)
    with pc2:
        txn_df = filtered.groupby(p)[v].count().reset_index()
        txn_df.columns = [p, 'Txns']
        merged = prod.merge(txn_df, on=p, how='left')
        fig6 = px.scatter(merged, x='Txns', y=v,
            size=v, color=v, color_continuous_scale='YlOrRd',
            hover_name=p)
        fig6.update_layout(**CHART_THEME, title="Volume vs Revenue")
        st.plotly_chart(fig6, use_container_width=True)
make_chart_safe(prod_charts, "Product charts")

# ════════════════════════════════════════════════════════
#  AI REPORT
# ════════════════════════════════════════════════════════
st.markdown("<div class='sec-title'>🤖 AI MIS Report</div>", unsafe_allow_html=True)

stats = {
    "total_sales"       : fmt_inr(total_val),
    "total_transactions": total_txn,
    "total_customers"   : total_cust,
    "avg_order_value"   : fmt_inr(avg_order),
    "file"              : uploaded.name,
    "sheet"             : sheet_used,
    "columns_detected"  : {
        "date": d, "value": v, "category": c,
        "customer": n, "product": p, "manager": m
    }
}

# Add monthly stats safely
if valid(v, filtered) and valid(d, filtered):
    try:
        tr2 = filtered.copy()
        ds2 = get_col(tr2, d)
        if ds2 is not None:
            tr2['_mo'] = ds2.dt.to_period('M').astype(str)
            mo2 = tr2.groupby('_mo')[v].sum()
            if len(mo2):
                stats['best_month']  = str(mo2.idxmax())
                stats['worst_month'] = str(mo2.idxmin())
    except: pass

if valid(c, filtered) and valid(v, filtered):
    try:
        top_cat = safe_groupby_sum(filtered, c, v)
        if not top_cat.empty:
            stats['top_region'] = str(top_cat.sort_values(v, ascending=False).iloc[0][c])
    except: pass

if valid(n, filtered) and valid(v, filtered):
    try:
        top_cust = safe_groupby_sum(filtered, n, v)
        if not top_cust.empty:
            stats['top_customer'] = str(top_cust.sort_values(v, ascending=False).iloc[0][n])
    except: pass

if api_key:
    if st.button("🚀 Generate AI MIS Report", type="primary", use_container_width=True):
        with st.spinner("🤖 Claude is writing your MIS report..."):
            report = generate_ai_report(stats, api_key)
        st.markdown(f"""
        <div class='ai-report'>
          <div class='ai-badge'>🤖 AI GENERATED · CLAUDE</div>
          {report.replace(chr(10), '<br>')}
        </div>""", unsafe_allow_html=True)
else:
    lines = [
        "**📊 Executive Summary**",
        f"Total sales: **{fmt_inr(total_val)}** across **{total_txn:,}** transactions "
        f"from **{total_cust:,}** unique customers. Avg order: **{fmt_inr(avg_order)}**.",
        ""
    ]
    if valid(c, filtered) and valid(v, filtered):
        try:
            top3 = safe_groupby_sum(filtered, c, v).sort_values(v, ascending=False).head(3)
            if not top3.empty:
                lines.append("**🔥 Top Regions:** " +
                    ", ".join([f"{row[c]} ({fmt_inr(row[v])})"
                               for _, row in top3.iterrows()]))
        except: pass
    lines += ["", "**💡 Add Claude API Key** for full AI-powered insights!"]
    st.markdown(f"""
    <div class='ai-report'>
      <div class='ai-badge'>⚡ SMART REPORT</div>
      {'<br>'.join(lines)}
    </div>""", unsafe_allow_html=True)

# ════════════════════════════════════════════════════════
#  DATA TABLE + DOWNLOADS
# ════════════════════════════════════════════════════════
st.markdown("<div class='sec-title'>📋 Filtered Data</div>", unsafe_allow_html=True)
show_cols = [x for x in [d,v,q,c,n,p,m] if valid(x, filtered)]
display_df = filtered[show_cols] if show_cols else filtered
st.dataframe(display_df.head(500), use_container_width=True, height=280)

dl1, dl2 = st.columns(2)
dl1.download_button(
    "📥 Download Full Report (CSV)",
    filtered.to_csv(index=False),
    f"MIS_{uploaded.name}_{datetime.date.today()}.csv",
    "text/csv", use_container_width=True
)
if valid(v, filtered) and valid(d, filtered):
    try:
        tr3 = filtered.copy()
        ds3 = get_col(tr3, d)
        if ds3 is not None:
            tr3['_mo'] = ds3.dt.to_period('M').astype(str)
            mo3 = tr3.groupby('_mo')[v].sum().reset_index()
            mo3.columns = ['Month', 'Sales']
            dl2.download_button(
                "📥 Monthly Summary (CSV)",
                mo3.to_csv(index=False),
                f"Monthly_{datetime.date.today()}.csv",
                "text/csv", use_container_width=True
            )
    except: pass

st.markdown("""
<div style='margin-top:40px;text-align:center;font-family:DM Mono,monospace;
            font-size:10px;color:rgba(234,234,245,.2);letter-spacing:2px'>
  MIS AUTOREPORT PRO · BUILT BY SATYENDRA · POWERED BY CLAUDE AI
</div>""", unsafe_allow_html=True)
