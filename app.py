import streamlit as st
import pandas as pd
import openai
import os
from datetime import datetime

# Load all category sheets into a dictionary of DataFrames
@st.cache_data
def load_all_data(file_path):
    xl = pd.ExcelFile(file_path)
    data_dict = {}
    for sheet_name in xl.sheet_names:
        df = xl.parse(sheet_name)
        df['Category'] = sheet_name  # Add a column to identify the category
        data_dict[sheet_name] = df
    return data_dict

# Match outlet by normalized name or ID
def get_outlet_data(df, outlet_query):
    outlet_query = outlet_query.replace(" ", "").lower().strip()
    df['Normalized Name'] = df['Outlet Name'].astype(str).str.replace(" ", "").str.lower().str.strip()
    df['Normalized ID'] = df['Outlet ID'].astype(str).str.strip()
    match = df[(df['Normalized Name'] == outlet_query) | (df['Normalized ID'] == outlet_query)]
    return match

# Generate performance summary for an outlet
def generate_outlet_summary(df, outlet_name):
    if df.empty:
        return f"No outlet matching '{outlet_name}' found."

    outlet_id = df.iloc[0]['Outlet ID']
    head_office = df.iloc[0]['Head Office']
    category = df.iloc[0]['Category']

    current_month = df.iloc[0]['Month']
    cm_sales = df.iloc[0]['Current Month Sales']
    ly_sales = df.iloc[0]['Sales Last Year Same Month']
    ytd_sales = df.iloc[0]['YTD Sales']
    ytd_ly_sales = df.iloc[0]['YTD Sales Last Year']

    cm_growth = ((cm_sales - ly_sales) / ly_sales) * 100 if ly_sales != 0 else 0
    ytd_growth = ((ytd_sales - ytd_ly_sales) / ytd_ly_sales) * 100 if ytd_ly_sales != 0 else 0

    ho_count = df[df['Head Office'] == head_office]['Outlet ID'].nunique()

    summary = f"""
    📌 Outlet Performance for **{df.iloc[0]['Outlet Name']}**  
    🏷️ Head Office: **{head_office}** (Total Branches: {ho_count})  
    📦 Category: **{category}**  
    📅 Current Month: **{current_month}**  

    **Sales Performance:**  
    - Current Month Sales: {cm_sales:,.0f}  
    - Last Year Same Month Sales: {ly_sales:,.0f}  
    - % Growth vs LY (Month): **{cm_growth:.1f}%**  
    - YTD Sales: {ytd_sales:,.0f}  
    - YTD Sales Last Year: {ytd_ly_sales:,.0f}  
    - % Growth YTD: **{ytd_growth:.1f}%**
    """
    return summary

# Generate category summary

def generate_category_summary(data_dict):
    summary = "## 📦 Category Summary\n"
    for category, df in data_dict.items():
        current = df.iloc[0]['Current Month']
        prev_year = df.iloc[0]['Previous Year Month']

        cm_sales = df['Current Month Sales'].sum()
        ly_sales = df['Sales Last Year Same Month'].sum()
        ytd_sales = df['YTD Sales'].sum()
        ytd_ly_sales = df['YTD Sales Last Year'].sum()

        cm_growth = ((cm_sales - ly_sales) / ly_sales) * 100 if ly_sales != 0 else 0
        ytd_growth = ((ytd_sales - ytd_ly_sales) / ytd_ly_sales) * 100 if ytd_ly_sales != 0 else 0

        # Zero Sales Logic: no sales this month but with at least 2 of last 3 months
        zero_sales_outlets = df[
            (df['Current Month Sales'] == 0) &
            (df[['M-1 Sales', 'M-2 Sales', 'M-3 Sales']] > 0).sum(axis=1) >= 2
        ]

        zero_sales_count = zero_sales_outlets['Outlet ID'].nunique()

        summary += f"""
**{category}**  
- % Growth vs LY (Month): **{cm_growth:.1f}%**  
- % Growth YTD: **{ytd_growth:.1f}%**  
- Zero Sales Outlets: **{zero_sales_count}**\n\n"""
    return summary

# Load data
st.set_page_config(page_title="Sales AI Agent", page_icon="💬")
st.title("💬 Sales AI Agent Chat")
st.markdown("Hi Jerby! 👋 How can I help you today?")

uploaded_file = "MT Raw Data_Updated.xlsx"
data_dict = load_all_data(uploaded_file)
all_data = pd.concat(data_dict.values(), ignore_index=True)

# User input
query = st.text_input("Enter Outlet ID or Name (use voice input via your keyboard mic)")

if query:
    outlet_df = get_outlet_data(all_data.copy(), query)
    outlet_response = generate_outlet_summary(outlet_df, query)
    st.markdown(outlet_response)

# Category Summary
st.markdown(generate_category_summary(data_dict))
st.markdown("---")
st.markdown("Built with ❤️ by Jerby & ChatGPT")
