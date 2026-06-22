import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go

st.set_page_config(
    page_title="Sales MIS Dashboard",
    page_icon="📊",
    layout="wide"
)

st.title("📊 Sales MIS Dashboard — 2023-2026")
st.markdown("---")

file = st.file_uploader(
    "Upload Sales Excel File",
    type=["xlsx", "csv"]
)

if file:
    # Read Main File sheet directly
    df = pd.read_excel(file, sheet_name='Main File')

    # Clean columns
    df.rename(columns={'Party Name\n': 'Party Name'}, inplace=True)
    df['Date'] = pd.to_datetime(df['Date'], errors='coerce')
    df['Value'] = pd.to_numeric(df['Value'], errors='coerce')
    df['QTY'] = pd.to_numeric(df['QTY'], errors='coerce')
    df = df.dropna(subset=['Date', 'Value'])

    # SIDEBAR FILTERS
    st.sidebar.header("🔍 Filters")

    # Year filter
    df['Year'] = df['Date'].dt.year
    years = sorted(df['Year'].unique())
    sel_year = st.sidebar.multiselect(
        "Year", years, default=years
    )

    # Region filter
    regions = sorted(df['Region'].dropna().unique())
    sel_region = st.sidebar.multiselect(
        "Region", regions, default=regions
    )

    # Manager filter
    managers = sorted(df['Sales Manager'].dropna().unique())
    sel_manager = st.sidebar.multiselect(
        "Sales Manager", managers, default=managers
    )

    # Season filter
    seasons = df['Season/Offseason'].dropna().unique()
    sel_season = st.sidebar.multiselect(
        "Season", seasons, default=seasons
    )

    # Apply filters
    filtered = df[
        (df['Year'].isin(sel_year)) &
        (df['Region'].isin(sel_region)) &
        (df['Sales Manager'].isin(sel_manager)) &
        (df['Season/Offseason'].isin(sel_season))
    ]

    st.success(f"✅ {len(filtered):,} records loaded")

    # ── KPI CARDS ──
    st.subheader("📌 Key Metrics")
    k1, k2, k3, k4 = st.columns(4)

    total_val = filtered['Value'].sum()
    total_qty = filtered['QTY'].sum()
    avg_rate  = filtered['Rate'].mean() if 'Rate' in filtered.columns else 0
    total_parties = filtered['Party Name'].nunique()

    k1.metric("💰 Total Sales Value",
              f"₹{total_val/1e7:.2f} Cr")
    k2.metric("📦 Total QTY",
              f"{total_qty/1000:.1f}K Units")
    k3.metric("🏢 Unique Parties",
              f"{total_parties:,}")
    k4.metric("📋 Total Transactions",
              f"{len(filtered):,}")

    st.markdown("---")

    # ── MONTHLY TREND ──
    st.subheader("📈 Monthly Sales Trend")
    monthly = (
        filtered.groupby(filtered['Date'].dt.to_period('M'))
        ['Value'].sum()
        .reset_index()
    )
    monthly['Date'] = monthly['Date'].astype(str)
    fig1 = px.line(
        monthly, x='Date', y='Value',
        title="Monthly Sales Value (₹)",
        color_discrete_sequence=['#f5c842']
    )
    fig1.update_layout(
        plot_bgcolor='#0f0f16',
        paper_bgcolor='#0f0f16',
        font_color='white'
    )
    st.plotly_chart(fig1, use_container_width=True)

    # ── REGION WISE ──
    st.subheader("🗺️ Region-wise Sales")
    col1, col2 = st.columns(2)

    region_df = (
        filtered.groupby('Region')['Value']
        .sum().sort_values(ascending=False)
        .reset_index()
    )
    fig2 = px.bar(
        region_df, x='Region', y='Value',
        title="Sales by Region",
        color='Value',
        color_continuous_scale='Oranges'
    )
    fig2.update_layout(
        plot_bgcolor='#0f0f16',
        paper_bgcolor='#0f0f16',
        font_color='white'
    )
    col1.plotly_chart(fig2, use_container_width=True)

    # ── MANAGER WISE ──
    mgr_df = (
        filtered.groupby('Sales Manager')['Value']
        .sum().sort_values(ascending=False)
        .reset_index()
    )
    fig3 = px.pie(
        mgr_df, names='Sales Manager', values='Value',
        title="Manager-wise Sales Share"
    )
    fig3.update_layout(
        paper_bgcolor='#0f0f16',
        font_color='white'
    )
    col2.plotly_chart(fig3, use_container_width=True)

    # ── SEASON ANALYSIS ──
    st.subheader("🌙 Season vs Off-Season")
    season_df = (
        filtered.groupby('Season/Offseason')['Value']
        .sum().reset_index()
    )
    fig4 = px.bar(
        season_df,
        x='Season/Offseason', y='Value',
        color='Season/Offseason',
        title="Season vs Off-Season Sales",
        color_discrete_sequence=['#f5c842','#00e5a0']
    )
    fig4.update_layout(
        plot_bgcolor='#0f0f16',
        paper_bgcolor='#0f0f16',
        font_color='white'
    )
    st.plotly_chart(fig4, use_container_width=True)

    # ── TOP 10 PARTIES ──
    st.subheader("🏆 Top 10 Customers")
    top_parties = (
        filtered.groupby('Party Name')['Value']
        .sum().sort_values(ascending=False)
        .head(10).reset_index()
    )
    fig5 = px.bar(
        top_parties, x='Value', y='Party Name',
        orientation='h',
        title="Top 10 Customers by Value",
        color_discrete_sequence=['#4a9eff']
    )
    fig5.update_layout(
        plot_bgcolor='#0f0f16',
        paper_bgcolor='#0f0f16',
        font_color='white',
        yaxis={'categoryorder': 'total ascending'}
    )
    st.plotly_chart(fig5, use_container_width=True)

    # ── RAW DATA + DOWNLOAD ──
    st.subheader("📋 Filtered Data")
    show_cols = ['Date','Party Name','Region',
                 'Sales Manager','Item Name',
                 'QTY','Value','Season/Offseason']
    st.dataframe(
        filtered[show_cols].head(500),
        use_container_width=True
    )

    st.download_button(
        "📥 Download Filtered MIS (CSV)",
        filtered[show_cols].to_csv(index=False),
        "MIS_Report_Filtered.csv",
        "text/csv"
    )
