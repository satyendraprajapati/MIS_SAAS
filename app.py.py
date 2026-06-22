# Install: pip install streamlit pandas plotly openpyxl

import streamlit as st
import pandas as pd
import plotly.express as px

st.set_page_config(
  page_title="MIS AutoReport",
  layout="wide"
)

st.title("📊 MIS Auto Dashboard")
st.write("Upload your Excel → Get instant MIS")

file = st.file_uploader(
  "Upload Sales/Expense Excel",
  type=["xlsx", "csv"]
)

if file:
  df = pd.read_excel(file)
  st.success("✅ File loaded!")

  # Auto detect columns
  date_col = df.columns[df.columns.str.contains(
    'date|Date|DATE')][0]
  val_col = df.columns[df.columns.str.contains(
    'sales|Sales|amount|Amount')][0]

  # KPI Cards
  col1, col2, col3 = st.columns(3)
  col1.metric("Total Sales",
    f"₹{df[val_col].sum():,.0f}")
  col2.metric("Avg Daily",
    f"₹{df[val_col].mean():,.0f}")
  col3.metric("Best Day",
    f"₹{df[val_col].max():,.0f}")

  # Trend Chart
  fig = px.line(df, x=date_col, y=val_col,
    title="Sales Trend")
  st.plotly_chart(fig, use_container_width=True)

  # Download Button
  st.download_button(
    "📥 Download MIS Report",
    df.to_csv(index=False),
    "mis_report.csv"
  )