import streamlit as st
import pandas as pd
import matplotlib.pyplot as plt
from openai import OpenAI

# Initialize OpenAI client
client = OpenAI(api_key=st.secrets["OPENAI_API_KEY"])

st.set_page_config(page_title="Sales AI Agent Chat", layout="wide")
st.title("💬 Sales AI Agent Chat")
st.markdown("Hi Jerby! 👋 How can I help you today?")

# Load Excel file
file_path = "MT Sales Raw Data.xlsx"
try:
    sheets = pd.read_excel(file_path, sheet_name=None, engine="openpyxl")
    for sheet in sheets.values():
        sheet.columns = sheet.columns.str.strip()
except Exception as e:
    st.error(f"Failed to load data: {e}")
    st.stop()

# Manual input (works with voice-to-text on phone)
user_input = st.text_input("Enter Outlet ID or Name (use voice input via your keyboard mic)")
if not user_input:
    st.stop()

# Search and display outlet info
found = False
for sheet_name, df in sheets.items():
    if 'Outlet ID' not in df.columns:
        continue

    matched_rows = df[
        (df['Outlet ID'].astype(str).str.strip() == user_input.strip()) |
        (df['Outlet Name'].astype(str).str.strip().str.lower() == user_input.strip().lower())
    ]
    if not matched_rows.empty:
        found = True
        outlet_row = matched_rows.iloc[0]

        st.subheader(f"Performance Dashboard for: {outlet_row['Outlet Name']} ({outlet_row['Outlet ID']})")

        st.markdown(f"**Head Office Name:** {outlet_row.get('Head Office Name', 'N/A')}")
        head_office = outlet_row.get('Head Office Name')
        if head_office:
            total_branches = df[df['Head Office Name'].astype(str).str.strip() == str(head_office).strip()].shape[0]
            st.markdown(f"**Total Branches under this Head Office:** {total_branches}")

        st.markdown(f"**Channel:** {outlet_row.get('Customer Channel', 'N/A')}")
        st.markdown(f"**Segment:** {outlet_row.get('Customer Segment', 'N/A')}")
        st.markdown(f"**Status:** {outlet_row.get('Customer Status', 'N/A')}")
        st.markdown(f"**Warehouse:** {outlet_row.get('Warehouse', 'N/A')}")

        # Monthly Sales Trend Table
        sales_cols = sorted([col for col in df.columns if col[:4].isdigit() and len(col) == 7])
        sales_data = outlet_row[sales_cols].fillna(0)
        sales_df = pd.DataFrame({"Month": sales_cols, "Sales": sales_data.values})
        sales_df['Sales'] = sales_df['Sales'].map('{:,.0f}'.format)

        st.subheader("📈 Monthly Sales Trend")
        st.dataframe(sales_df)

        # LRB Chart
        if "LRB Sales" in sheets:
            lrb_df = sheets["LRB Sales"]
            lrb_match = lrb_df[
                (lrb_df['Outlet ID'].astype(str).str.strip() == str(outlet_row['Outlet ID']).strip())
            ]
            if not lrb_match.empty:
                st.subheader("📊 LRB Sales Monthly Trend")
                lrb_sales_cols = sorted([col for col in lrb_df.columns if col[:4].isdigit() and len(col) == 7])
                lrb_sales_data = lrb_match.iloc[0][lrb_sales_cols].fillna(0)

                fig, ax = plt.subplots(figsize=(10, 4))
                ax.plot(lrb_sales_cols, lrb_sales_data.values, marker='o')
                ax.set_title("LRB Sales Trend")
                ax.set_xlabel("Month")
                ax.yaxis.set_visible(False)
                ax.tick_params(axis='x', rotation=45)
                for x, y in zip(lrb_sales_cols, lrb_sales_data.values):
                    ax.text(x, y, f"{y:,.0f}", ha='center', va='bottom', fontsize=9)
                st.pyplot(fig)

        # AI QA
        question = st.text_input("Ask a sales-related question about this outlet:")
        if st.button("Get AI Answer"):
            context = f"Outlet: {outlet_row.get('Outlet Name')} | Channel: {outlet_row.get('Customer Channel')}"
            prompt = (
                f"You are a sales data analyst.\nContext: {context}\n"
                f"Here are the monthly sales data:\n{sales_df.to_csv(index=False)}\n"
                f"Question: {question}\nProvide a detailed and precise answer."
            )
            try:
                response = client.chat.completions.create(
                    model="gpt-3.5-turbo",
                    messages=[{"role": "user", "content": prompt}],
                    temperature=0
                )
                st.markdown("### 🧠 AI Agent Answer")
                st.write(response.choices[0].message.content.strip())
            except Exception as e:
                st.error(f"OpenAI API error: {e}")
        break

if not found:
    st.warning(f"No outlet matching '{user_input}' found.")

# Category Summary Table
st.subheader("📦 Category Summary")
current_month = "2025-06"
previous_year = "2024"
current_year = "2025"
summary_data = []

for sheet_name, df in sheets.items():
    if 'Outlet ID' not in df.columns:
        continue
    sales_cols = sorted([col for col in df.columns if col[:4].isdigit() and len(col) == 7])
    current_month_col = current_month
    last_year_month_col = current_month.replace(current_year, previous_year)
    if current_month_col not in df.columns or last_year_month_col not in df.columns:
        continue

    df_outlet = df[
        (df['Outlet ID'].astype(str).str.strip() == user_input.strip()) |
        (df['Outlet Name'].astype(str).str.strip().str.lower() == user_input.strip().lower())
    ]
    if df_outlet.empty:
        continue

    current_month_sales = df_outlet[current_month_col].values[0]
    last_year_month_sales = df_outlet[last_year_month_col].values[0]
    month_growth = (current_month_sales / last_year_month_sales - 1) * 100 if last_year_month_sales else 0.0

    ytd_months_current = [f"{current_year}-{str(m).zfill(2)}" for m in range(1, int(current_month[-2:]) + 1)]
    ytd_months_previous = [f"{previous_year}-{str(m).zfill(2)}" for m in range(1, int(current_month[-2:]) + 1)]
    ytd_months_current = [m for m in ytd_months_current if m in df_outlet.columns]
    ytd_months_previous = [m for m in ytd_months_previous if m in df_outlet.columns]
    ytd_current = df_outlet[ytd_months_current].sum(axis=1).values[0]
    ytd_previous = df_outlet[ytd_months_previous].sum(axis=1).values[0]
    ytd_growth = (ytd_current / ytd_previous - 1) * 100 if ytd_previous else 0.0

    zero_sales_outlets = 0
    for _, row in df.iterrows():
        if row.get(current_month_col, 0) == 0:
            prev_sales = []
            for offset in range(1, 4):
                month_int = int(current_month[-2:]) - offset
                if month_int > 0:
                    prev_month = f"{current_year}-{str(month_int).zfill(2)}"
                    if prev_month in df.columns:
                        prev_sales.append(row.get(prev_month, 0) > 0)
            if sum(prev_sales) >= 2:
                zero_sales_outlets += 1

    summary_data.append({
        "Category": sheet_name,
        f"{current_month} Growth % vs LY": f"{month_growth:.1f}%",
        "YTD Growth %": f"{ytd_growth:.1f}%",
        "Zero Sales Outlets": zero_sales_outlets
    })

if summary_data:
    summary_df = pd.DataFrame(summary_data)
    st.dataframe(summary_df)

st.caption("Built with ❤️ by Jerby & ChatGPT")
