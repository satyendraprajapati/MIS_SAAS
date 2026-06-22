import streamlit as st
import pandas as pd
import plotly.express as px

st.set_page_config(page_title="MIS AutoReport", layout="wide")
st.title("📊 MIS Auto Dashboard")
st.write("Upload your Excel → Get instant MIS")

file = st.file_uploader(
    "Upload Sales/Expense Excel",
    type=["xlsx", "csv"]
)

if file:
    # Read file
    if file.name.endswith('.csv'):
        df = pd.read_csv(file)
    else:
        df = pd.read_excel(file)

    st.success("✅ File loaded!")
    st.write("📋 Columns found:", list(df.columns))

    # Step 1: User selects columns manually
    st.subheader("⚙️ Select Your Columns")
    
    all_cols = list(df.columns)
    
    date_col = st.selectbox(
        "📅 Date column select karo:",
        options=all_cols
    )
    
    val_col = st.selectbox(
        "💰 Sales/Amount column select karo:",
        options=all_cols,
        index=min(1, len(all_cols)-1)
    )

    if st.button("🚀 Generate Dashboard"):
        try:
            # Convert date column
            df[date_col] = pd.to_datetime(
                df[date_col], errors='coerce'
            )
            df[val_col] = pd.to_numeric(
                df[val_col], errors='coerce'
            )
            df = df.dropna(subset=[date_col, val_col])

            # KPI Cards
            col1, col2, col3 = st.columns(3)
            col1.metric(
                "💰 Total Sales",
                f"₹{df[val_col].sum():,.0f}"
            )
            col2.metric(
                "📊 Avg Daily",
                f"₹{df[val_col].mean():,.0f}"
            )
            col3.metric(
                "🏆 Best Day",
                f"₹{df[val_col].max():,.0f}"
            )

            # Trend Chart
            fig = px.line(
                df, x=date_col, y=val_col,
                title="📈 Sales Trend"
            )
            st.plotly_chart(fig, use_container_width=True)

            # Bar Chart
            fig2 = px.bar(
                df, x=date_col, y=val_col,
                title="📊 Daily Sales Bar Chart"
            )
            st.plotly_chart(fig2, use_container_width=True)

            # Raw Data
            st.subheader("📋 Raw Data")
            st.dataframe(df, use_container_width=True)

            # Download
            st.download_button(
                "📥 Download MIS Report (CSV)",
                df.to_csv(index=False),
                "mis_report.csv",
                "text/csv"
            )

        except Exception as e:
            st.error(f"❌ Error: {e}")
            st.info("💡 Sahi columns select karo aur retry karo")
