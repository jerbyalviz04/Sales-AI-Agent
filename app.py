import streamlit as st
import pandas as pd
import datetime
import os

# === PAGE CONFIG ===
st.set_page_config(page_title="Sales AI Agent", page_icon="💬", layout="centered")

# === TITLE ===
st.markdown("## 💬 Sales AI Agent Chat")
st.markdown("Hi Jerby! 👋 How can I help you today?")
st.markdown("Enter Outlet ID or Name (use voice input via your keyboard mic)")

# === FILE LOADING ===
@st.cache_data
def load_data():
    xls = pd.ExcelFile("MT Raw Data_Updated.xlsx")
    sheets = {sheet_name: xls.parse(sheet_name) for sheet_name in xls.sheet_names}
    return sheets

sheets = load_data()

# === HELPER FUNCTIONS ===
def calculate_growth(df, current_month, last_year_month):
    current = df.get(current_month, 0)
    last_year = df.get(last_year_month, 0)
    if last_year == 0:
        return None
    return round((current - last_year) / last_year * 100, 1)

def zero_sales_outlets(df):
    recent_months = df.columns[-4:-1]  # last 3 months excluding current
    current_month = df.columns[-1]
    zero_sales = df[(df[current_month] == 0) & (df[recent_months].gt(0).sum(axis=1) >= 2)]
    return len(zero_sales)

def get_head_office_count(df, ho_name):
    if 'Head Office Name' in df.columns:
        return df[df['Head Office Name'].str.lower() == ho_name.lower()]['Outlet ID'].nunique()
    return 0

def get_outlet_data(df, outlet_query):
    if outlet_query.isnumeric():
        return df[df['Outlet ID'].astype(str) == outlet_query]
    else:
        return df[df['Outlet Name'].str.lower().str.strip() == outlet_query.lower().strip()]

# === MAIN INPUT FIELD ===
user_query = st.text_input("", placeholder="e.g. Outlet001 or 1146231112942")

if user_query:
    found = False
    category_summary = []
    for category, df in sheets.items():
        outlet_df = get_outlet_data(df, user_query)
        if not outlet_df.empty:
            found = True
            st.subheader(f"📍 Outlet Performance: {user_query}")
            
            # Metrics
            current_month = df.columns[-1]
            last_year_same_month = df.columns[-13]
            ytd_this_year = df.iloc[:, -6:].sum(axis=1).values[0]
            ytd_last_year = df.iloc[:, -18:-12].sum(axis=1).values[0]

            growth_month = calculate_growth(outlet_df.iloc[0], current_month, last_year_same_month)
            growth_ytd = calculate_growth({'this': ytd_this_year, 'last': ytd_last_year}, 'this', 'last')

            st.markdown(f"- **Current Month Growth**: {growth_month}%")
            st.markdown(f"- **YTD Growth**: {growth_ytd}%")

            # Head Office
            ho_name = outlet_df['Head Office Name'].values[0]
            ho_count = get_head_office_count(df, ho_name)
            st.markdown(f"- **Head Office**: {ho_name} ({ho_count} total outlets)")

            break  # Found the outlet, skip checking other sheets

    if not found:
        st.warning(f"No outlet matching '{user_query}' found.")

    # === CATEGORY SUMMARY ===
    st.subheader("📦 Category Summary")
    for category, df in sheets.items():
        current_month = df.columns[-1]
        last_year_same_month = df.columns[-13]
        df_sum = df.sum(numeric_only=True)

        cmg = calculate_growth(df_sum, current_month, last_year_same_month)
        ytd_this = df.iloc[:, -6:].sum().sum()
        ytd_last = df.iloc[:, -18:-12].sum().sum()
        ytd_growth = calculate_growth({'this': ytd_this, 'last': ytd_last}, 'this', 'last')
        zero_outlets = zero_sales_outlets(df)

        st.markdown(f"**{category}**")
        st.markdown(f"- Current Month Growth vs LY: **{cmg}%**")
        st.markdown(f"- YTD Growth: **{ytd_growth}%**")
        st.markdown(f"- Zero Sales Outlets: **{zero_outlets}**")

# === FOOTER ===
st.markdown("---")
st.markdown("Built with ❤️ by Jerby & ChatGPT")
