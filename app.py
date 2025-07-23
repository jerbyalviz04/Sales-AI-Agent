import streamlit as st
import pandas as pd
import speech_recognition as sr
from openai import OpenAI
import matplotlib.pyplot as plt

# Use OpenAI API key from Streamlit secrets
client = OpenAI(api_key=st.secrets["OPENAI_API_KEY"])

st.set_page_config(page_title="Sales AI Agent", layout="wide")
st.title("💬 Sales AI Agent Chat")

# Load the Excel file from the repo directory
file_path = "MT Sales Raw Data.xlsx"
try:
    sheets = pd.read_excel(file_path, sheet_name=None, engine="openpyxl")
    for sheet in sheets.values():
        sheet.columns = sheet.columns.str.strip()
except Exception as e:
    st.error(f"Failed to load data: {e}")
    st.stop()

if 'chat_history' not in st.session_state:
    st.session_state.chat_history = []

with st.chat_message("assistant"):
    st.markdown("Hi Jerby! 👋 How can I help you today?")

user_query = st.chat_input("Ask a question about a specific outlet (ID or name)...")

if user_query:
    st.chat_message("user").markdown(user_query)

    found = False
    for sheet_name, df in sheets.items():
        if 'Outlet ID' in df.columns:
            matched_rows = df[
                (df['Outlet ID'].astype(str).str.strip() == user_query.strip()) |
                (df['Outlet Name'].astype(str).str.strip().str.lower() == user_query.strip().lower())
            ]
            if not matched_rows.empty:
                found = True
                for _, outlet_row in matched_rows.iterrows():
                    context = f"Outlet: {outlet_row.get('Outlet Name', 'N/A')} | Channel: {outlet_row.get('Customer Channel', 'N/A')} | Segment: {outlet_row.get('Customer Segment', 'N/A')} | Status: {outlet_row.get('Customer Status', 'N/A')} | Warehouse: {outlet_row.get('Warehouse', 'N/A')}"

                    sales_cols = [col for col in df.columns if '-' in col]
                    sales_data = outlet_row[sales_cols]
                    sales_df = pd.DataFrame({
                        'Month': sales_cols,
                        'Sales': sales_data.values
                    })

                    # Convert to numeric for calculations
                    sales_df['Sales'] = pd.to_numeric(sales_df['Sales'], errors='coerce')

                    # Calculate Category Summaries
                    if sheet_name.lower() in ['lrb sales', 'csd sales', 'water sales']:
                        current_year = '2025'
                        previous_year = '2024'

                        june_2025 = outlet_row.get('2025-06', None)
                        june_2024 = outlet_row.get('2024-06', None)
                        ytd_2025 = outlet_row[[f"2025-{str(m).zfill(2)}" for m in range(1, 7)]].sum()
                        ytd_2024 = outlet_row[[f"2024-{str(m).zfill(2)}" for m in range(1, 7)]].sum()

                        june_growth = ((june_2025 - june_2024) / june_2024) * 100 if pd.notna(june_2025) and pd.notna(june_2024) and june_2024 != 0 else None
                        ytd_growth = ((ytd_2025 - ytd_2024) / ytd_2024) * 100 if pd.notna(ytd_2025) and pd.notna(ytd_2024) and ytd_2024 != 0 else None

                        st.subheader(f"📊 {sheet_name} Category Summary")
                        st.markdown(f"**June 2025 Sales:** {june_2025:,.0f}")
                        st.markdown(f"**June 2024 Sales:** {june_2024:,.0f}")
                        st.markdown(f"**June Growth % vs LY:** {june_growth:.1f}%")
                        st.markdown(f"**YTD Growth % (Jan–Jun):** {ytd_growth:.1f}%")

                        # Zero Sales: No sales this month but sales in at least 2 of the last 3 months
                        last_3_months = [f"2025-{str(m).zfill(2)}" for m in [4, 5, 6] if f"2025-{str(m).zfill(2)}" in outlet_row.index]
                        zero_sales = pd.isna(outlet_row[f"2025-06"]) or outlet_row[f"2025-06"] == 0
                        recent_sales = outlet_row[last_3_months[:-1]].fillna(0) > 0
                        if zero_sales and recent_sales.sum() >= 2:
                            st.markdown("**Zero Sales Flag:** Yes")
                        else:
                            st.markdown("**Zero Sales Flag:** No")

                        # Monthly Trend Table
                        st.markdown("---")
                        st.subheader(f"📈 Monthly {sheet_name} Sales Trend")

                        trend_df = pd.DataFrame({
                            'Month': sales_cols,
                            'Sales': sales_data.values
                        })
                        trend_df['Sales'] = pd.to_numeric(trend_df['Sales'], errors='coerce')
                        trend_df.dropna(inplace=True)

                        st.dataframe(trend_df)

                        # Plot trend
                        fig, ax = plt.subplots()
                        ax.plot(trend_df['Month'], trend_df['Sales'], marker='o')
                        for i, val in enumerate(trend_df['Sales']):
                            ax.text(i, val, f"{val:,.0f}", ha='center', va='bottom', fontsize=8)
                        ax.set_xticks(range(len(trend_df['Month'])))
                        ax.set_xticklabels(trend_df['Month'], rotation=45)
                        ax.set_ylabel("Sales")
                        ax.set_title(f"Monthly {sheet_name} Sales Trend")
                        ax.spines['left'].set_visible(False)
                        st.pyplot(fig)

                    # AI summary
                    formatted_df = sales_df.copy()
                    formatted_df['Sales'] = formatted_df['Sales'].apply(lambda x: f"{x:,.0f}" if pd.notna(x) else "0")
                    top_sales = formatted_df.to_csv(index=False)

                    prompt = (
                        f"You are a sales data analyst.\n"
                        f"Context: {context}\n"
                        f"Here are the monthly sales data:\n{top_sales}\n"
                        f"Provide a summary of sales performance for this outlet."
                    )

                    try:
                        response = client.chat.completions.create(
                            model="gpt-3.5-turbo",
                            messages=[{"role": "user", "content": prompt}],
                            temperature=0
                        )
                        ai_answer = response.choices[0].message.content.strip()
                        st.chat_message("assistant").markdown(ai_answer)

                    except Exception as e:
                        st.chat_message("assistant").markdown(f"⚠️ OpenAI API error: {e}")
                break
    if not found:
        st.chat_message("assistant").markdown(f"🔍 No outlet matching '{user_query}' found in any sheet.")
