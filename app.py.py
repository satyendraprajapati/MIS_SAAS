import streamlit as st
import pandas as pd
import plotly.express as px
import json
import requests

st.set_page_config(
    page_title="MIS AutoReport — AI Powered",
    page_icon="📊",
    layout="wide"
)

# ── STYLING ──
st.markdown("""
<style>
    .main { background-color: #07070a; }
    .stApp { background-color: #07070a; color: #eaeaf5; }
    .metric-card {
        background: #0f0f16;
        border: 1px solid rgba(245,200,66,0.2);
        border-radius: 12px;
        padding: 16px;
        text-align: center;
    }
    .metric-num {
        font-size: 28px;
        font-weight: 800;
        color: #f5c842;
    }
    .metric-label {
        font-size: 12px;
        color: rgba(234,234,245,0.5);
        margin-top: 4px;
    }
</style>
""", unsafe_allow_html=True)

st.title("📊 MIS AutoReport")
st.caption("Upload any Excel — AI will auto-detect and build your dashboard")

ANTHROPIC_API_KEY = st.secrets.get("ANTHROPIC_API_KEY", "")

def detect_columns_with_ai(columns, sample_data):
    """Use Claude API to intelligently detect column types"""
    prompt = f"""You are a data analyst. Analyze these Excel column names and sample data, then identify which columns represent:
1. date_col: The main date/time column
2. value_col: The main sales/revenue/amount/value column  
3. category_col: A category like region/area/zone/department
4. name_col: Customer/party/client name column
5. product_col: Product/item name column
6. qty_col: Quantity column

Column names: {columns}

Sample data (first 3 rows):
{sample_data}

Respond ONLY with a valid JSON object like this:
{{
  "date_col": "exact column name or null",
  "value_col": "exact column name or null",
  "category_col": "exact column name or null", 
  "name_col": "exact column name or null",
  "product_col": "exact column name or null",
  "qty_col": "exact column name or null",
  "confidence": "high/medium/low",
  "summary": "one line about what this data contains"
}}

Use EXACT column names from the list. If not found, use null."""

    try:
        response = requests.post(
            "https://api.anthropic.com/v1/messages",
            headers={
                "x-api-key": ANTHROPIC_API_KEY,
                "anthropic-version": "2023-06-01",
                "content-type": "application/json"
            },
            json={
                "model": "claude-sonnet-4-6",
                "max_tokens": 500,
                "messages": [{"role": "user", "content": prompt}]
            },
            timeout=30
        )
        result = response.json()
        text = result["content"][0]["text"]
        # Clean JSON
        text = text.strip()
        if "```" in text:
            text = text.split("```")[1]
            if text.startswith("json"):
                text = text[4:]
        return json.loads(text.strip())
    except Exception as e:
        return None

def smart_fallback_detection(df):
    """Fallback: detect columns without AI using keywords"""
    cols = {c.lower(): c for c in df.columns}
    result = {
        "date_col": None, "value_col": None,
        "category_col": None, "name_col": None,
        "product_col": None, "qty_col": None,
        "confidence": "low",
        "summary": "Auto-detected using keyword matching"
    }
    date_kw   = ['date','time','day','month','year','dt','दिनांक']
    value_kw  = ['value','sales','amount','revenue','total','sale','turnover','बिक्री','मूल्य']
    cat_kw    = ['region','area','zone','city','state','branch','department','location','city']
    name_kw   = ['party','customer','client','name','buyer','account','distributor']
    prod_kw   = ['item','product','sku','goods','material','description','article']
    qty_kw    = ['qty','quantity','units','pieces','count','volume','मात्रा']

    for kw in date_kw:
        match = [c for lc,c in cols.items() if kw in lc]
        if match:
            result["date_col"] = match[0]; break

    for kw in value_kw:
        match = [c for lc,c in cols.items() if kw in lc]
        if match:
            result["value_col"] = match[0]; break

    for kw in cat_kw:
        match = [c for lc,c in cols.items() if kw in lc]
        if match:
            result["category_col"] = match[0]; break

    for kw in name_kw:
        match = [c for lc,c in cols.items() if kw in lc]
        if match:
            result["name_col"] = match[0]; break

    for kw in prod_kw:
        match = [c for lc,c in cols.items() if kw in lc]
        if match:
            result["product_col"] = match[0]; break

    for kw in qty_kw:
        match = [c for lc,c in cols.items() if kw in lc]
        if match:
            result["qty_col"] = match[0]; break

    # If nothing found — pick by dtype
    if not result["date_col"]:
        date_cols = df.select_dtypes(include=['datetime64']).columns
        if len(date_cols): result["date_col"] = date_cols[0]

    if not result["value_col"]:
        num_cols = df.select_dtypes(include=['float64','int64']).columns
        if len(num_cols): result["value_col"] = num_cols[0]

    return result

# ── FILE UPLOAD ──
file = st.file_uploader(
    "Upload your Excel or CSV file",
    type=["xlsx", "xls", "csv"]
)

if file:
    # Read file
    try:
        if file.name.endswith('.csv'):
            df_raw = pd.read_csv(file)
            sheet_name = None
        else:
            xl = pd.ExcelFile(file)
            sheets = xl.sheet_names
            # Pick sheet with most data
            if len(sheets) > 1:
                sheet_name = st.selectbox(
                    "📋 Which sheet to use?", sheets
                )
            else:
                sheet_name = sheets[0]
            df_raw = pd.read_excel(file, sheet_name=sheet_name)
    except Exception as e:
        st.error(f"File read error: {e}")
        st.stop()

    # Skip mostly-empty rows at top (handle merged headers)
    df_raw = df_raw.dropna(how='all')
    # If first row looks like header (many strings), reset
    if df_raw.columns.str.startswith('Unnamed').sum() > len(df_raw.columns) * 0.5:
        df_raw.columns = df_raw.iloc[0]
        df_raw = df_raw[1:].reset_index(drop=True)

    df_raw.columns = [str(c).strip() for c in df_raw.columns]
    col_list = [c for c in df_raw.columns if not c.startswith('Unnamed')]

    st.success(f"✅ File loaded — {len(df_raw):,} rows, {len(col_list)} columns")

    # ── AI COLUMN DETECTION ──
    with st.spinner("🤖 AI is analyzing your file structure..."):
        sample = df_raw[col_list].head(3).to_string()

        if ANTHROPIC_API_KEY:
            mapping = detect_columns_with_ai(col_list, sample)
        else:
            mapping = None

        if not mapping:
            mapping = smart_fallback_detection(df_raw)
            st.info("💡 AI fallback: keyword-based detection used")
        else:
            st.success(f"🤖 AI detected: {mapping.get('summary','')}")

    # Show detected mapping
    with st.expander("🔍 AI Column Mapping (click to verify)"):
        c1,c2,c3 = st.columns(3)
        c1.write(f"📅 Date: `{mapping['date_col']}`")
        c1.write(f"💰 Value: `{mapping['value_col']}`")
        c2.write(f"🗺️ Category: `{mapping['category_col']}`")
        c2.write(f"🏢 Customer: `{mapping['name_col']}`")
        c3.write(f"📦 Product: `{mapping['product_col']}`")
        c3.write(f"🔢 Quantity: `{mapping['qty_col']}`")

        st.markdown("**Override if wrong:**")
        all_cols_opt = ["(none)"] + col_list
        ov1,ov2 = st.columns(2)
        mapping['date_col']     = ov1.selectbox("Date col",     all_cols_opt, index=all_cols_opt.index(mapping['date_col'])     if mapping['date_col']     in all_cols_opt else 0)
        mapping['value_col']    = ov2.selectbox("Value col",    all_cols_opt, index=all_cols_opt.index(mapping['value_col'])    if mapping['value_col']    in all_cols_opt else 0)
        mapping['category_col'] = ov1.selectbox("Category col", all_cols_opt, index=all_cols_opt.index(mapping['category_col']) if mapping['category_col'] in all_cols_opt else 0)
        mapping['name_col']     = ov2.selectbox("Customer col", all_cols_opt, index=all_cols_opt.index(mapping['name_col'])     if mapping['name_col']     in all_cols_opt else 0)
        mapping['product_col']  = ov1.selectbox("Product col",  all_cols_opt, index=all_cols_opt.index(mapping['product_col'])  if mapping['product_col']  in all_cols_opt else 0)
        mapping['qty_col']      = ov2.selectbox("Qty col",      all_cols_opt, index=all_cols_opt.index(mapping['qty_col'])      if mapping['qty_col']      in all_cols_opt else 0)

        mapping = {k: (None if v=="(none)" else v) for k,v in mapping.items()}

    # ── PREPARE DATA ──
    df = df_raw.copy()
    d  = mapping['date_col']
    v  = mapping['value_col']
    c  = mapping['category_col']
    n  = mapping['name_col']
    p  = mapping['product_col']
    q  = mapping['qty_col']

    if d: df[d] = pd.to_datetime(df[d], errors='coerce')
    if v: df[v] = pd.to_numeric(df[v], errors='coerce')
    if q: df[q] = pd.to_numeric(df[q], errors='coerce')

    # ── SIDEBAR FILTERS ──
    st.sidebar.header("🔍 Filters")
    filtered = df.copy()

    if d and d in df.columns:
        df[d] = pd.to_datetime(df[d], errors='coerce')
        min_d = df[d].min()
        max_d = df[d].max()
        if pd.notnull(min_d) and pd.notnull(max_d):
            date_range = st.sidebar.date_input(
                "Date Range",
                [min_d.date(), max_d.date()]
            )
            if len(date_range) == 2:
                filtered = filtered[
                    (filtered[d] >= pd.Timestamp(date_range[0])) &
                    (filtered[d] <= pd.Timestamp(date_range[1]))
                ]

    if c and c in df.columns:
        cats = sorted(df[c].dropna().unique())
        sel = st.sidebar.multiselect(f"Filter by {c}", cats, default=cats)
        filtered = filtered[filtered[c].isin(sel)]

    # ── KPI CARDS ──
    st.markdown("---")
    st.subheader("📌 Key Metrics")
    m1,m2,m3,m4 = st.columns(4)

    total_val   = filtered[v].sum()   if v else 0
    total_qty   = filtered[q].sum()   if q else 0
    total_cust  = filtered[n].nunique() if n else 0
    total_rows  = len(filtered)

    with m1:
        st.markdown(f"""<div class="metric-card">
            <div class="metric-num">₹{total_val/1e7:.1f}Cr</div>
            <div class="metric-label">Total Sales Value</div>
        </div>""", unsafe_allow_html=True)
    with m2:
        st.markdown(f"""<div class="metric-card">
            <div class="metric-num">{total_qty/1000:.1f}K</div>
            <div class="metric-label">Total Quantity</div>
        </div>""", unsafe_allow_html=True)
    with m3:
        st.markdown(f"""<div class="metric-card">
            <div class="metric-num">{total_cust:,}</div>
            <div class="metric-label">Unique Customers</div>
        </div>""", unsafe_allow_html=True)
    with m4:
        st.markdown(f"""<div class="metric-card">
            <div class="metric-num">{total_rows:,}</div>
            <div class="metric-label">Transactions</div>
        </div>""", unsafe_allow_html=True)

    st.markdown("---")

    # ── CHARTS ──
    CHART_BG = dict(
        plot_bgcolor='#0f0f16',
        paper_bgcolor='#0f0f16',
        font_color='#eaeaf5'
    )

    # Monthly trend
    if d:
        st.subheader("📈 Sales Trend Over Time")
        trend = filtered.copy()
        trend['_month'] = trend[d].dt.to_period('M').astype(str)
        monthly = trend.groupby('_month')[v].sum().reset_index()
        fig = px.line(monthly, x='_month', y=v,
                      title="Monthly Sales",
                      color_discrete_sequence=['#f5c842'])
        fig.update_layout(**CHART_BG)
        st.plotly_chart(fig, use_container_width=True)

    # Category + Customer side by side
    col1, col2 = st.columns(2)

    if c:
        with col1:
            st.subheader(f"🗺️ By {c}")
            cat_df = filtered.groupby(c)[v].sum().sort_values(ascending=False).head(15).reset_index()
            fig2 = px.bar(cat_df, x=c, y=v,
                          color=v, color_continuous_scale='Oranges')
            fig2.update_layout(**CHART_BG)
            st.plotly_chart(fig2, use_container_width=True)

    if n:
        with col2:
            st.subheader("🏆 Top 10 Customers")
            top_n = filtered.groupby(n)[v].sum().sort_values(ascending=False).head(10).reset_index()
            fig3 = px.bar(top_n, x=v, y=n, orientation='h',
                          color_discrete_sequence=['#00e5a0'])
            fig3.update_layout(**CHART_BG,
                               yaxis={'categoryorder':'total ascending'})
            st.plotly_chart(fig3, use_container_width=True)

    # Product chart
    if p:
        st.subheader("📦 Top Products")
        prod_df = filtered.groupby(p)[v].sum().sort_values(ascending=False).head(10).reset_index()
        fig4 = px.pie(prod_df, names=p, values=v,
                      title="Product-wise Sales Share")
        fig4.update_layout(**CHART_BG)
        st.plotly_chart(fig4, use_container_width=True)

    # ── DATA TABLE + DOWNLOAD ──
    st.markdown("---")
    st.subheader("📋 Data Preview")
    show = [col for col in [d,v,c,n,p,q] if col and col in filtered.columns]
    st.dataframe(filtered[show].head(300) if show else filtered.head(300),
                 use_container_width=True)

    st.download_button(
        "📥 Download MIS Report (CSV)",
        filtered.to_csv(index=False),
        "MIS_AutoReport.csv",
        "text/csv"
    )

else:
    st.info("👆 Upload any Excel or CSV file to get started")
    st.markdown("""
    ### ✅ Works with ANY Excel structure:
    - Sales data with Date, Amount, Region columns
    - Tally exports
    - ERP downloads  
    - Custom MIS sheets
    - Multi-sheet workbooks
    
    **AI will auto-detect your columns!**
    """)
